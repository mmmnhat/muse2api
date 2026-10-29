#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Module đọc OTP Meta dùng chung cho auto-reg-mu.

Hỗ trợ ĐỒNG THỜI 2 đường đọc thư Outlook bằng OAuth2 refresh_token:

1. **IMAP XOAUTH2** — kết nối `outlook.office365.com:993`, quét Inbox + Junk.
   Cần token có scope IMAP (vd tài khoản bán dạng [IMAP/POP3/Graph]).
2. **Microsoft Graph API** — `GET graph.microsoft.com/v1.0/me/mailFolders/{inbox|junkemail}/messages`.
   Cần token có scope Graph Mail.Read (vd tài khoản bán dạng [GRAPH API]).

Chế độ `method="auto"` (mặc định): thử kết nối cả 2 đường, đường nào dùng được
thì quét song song trong cùng vòng lặp chờ OTP — tài khoản Graph-only hay
IMAP-only đều chạy, không cần chọn tay.

CHỈ HỖ TRỢ OUTLOOK: email phải có domain dạng `outlook.*` (outlook.com,
outlook.fr, ...). Hotmail/Live/MSN bị từ chối bởi `is_outlook_email`.
"""

from __future__ import annotations

import email
import email.utils
import imaplib
import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from typing import Dict, List, Optional, Tuple

TOKEN_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
IMAP_HOST = "outlook.office365.com"
IMAP_SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All offline_access"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = "https://graph.microsoft.com/Mail.Read offline_access"
GRAPH_FOLDERS = ("inbox", "junkemail")
IMAP_FOLDERS = ("Inbox", "Junk")

# Từ khóa nhận diện thư Meta (giữ nguyên như bản IMAP gốc)
META_KEYWORDS = ["Meta", "meta.com", "Xác nhận", "Confirm", "code", "Mã"]
# Các mã 6 số KHÔNG phải OTP (xuất hiện cố định trong template thư)
OTP_BLACKLIST = ("141823", "000000", "123456")


def is_outlook_email(email_addr: str) -> bool:
    """True nếu domain là outlook.* (outlook.com, outlook.fr, ...)."""
    if "@" not in email_addr:
        return False
    domain = email_addr.rsplit("@", 1)[1].strip().lower()
    return domain.startswith("outlook.")


def extract_meta_otp(frm: str, subj: str, raw_content: str) -> Optional[str]:
    """Bóc mã OTP 6 số từ nội dung thư Meta. Trả None nếu thư không khớp."""
    if not any(k in frm or k in subj for k in META_KEYWORDS):
        return None

    # Mẫu 1: "Mã xác nhận 123456"
    m = re.search(r"Mã xác nhận[:\s]+(\d{6})", raw_content)
    if m:
        return m.group(1)

    # Mẫu 2: HTML letter-spacing hoặc font-size của Meta email
    m2 = re.search(r"letter-spacing:[^>]*>(\d{6})<", raw_content)
    if m2:
        return m2.group(1)

    # Mẫu 3: "123456 is your Meta confirmation code"
    m3 = re.search(r"(\d{6})\s+is your Meta", raw_content)
    if m3:
        return m3.group(1)

    # Mẫu 4: Mọi chuỗi 6 chữ số hợp lệ
    for code in re.findall(r"\b\d{6}\b", raw_content):
        if code not in OTP_BLACKLIST:
            return code
    return None


def exchange_access_token(refresh_token: str, client_id: str, scope: Optional[str] = None) -> Optional[str]:
    """Đổi refresh_token lấy access_token. scope=None giữ nguyên scope gốc của token."""
    form = {
        "client_id": client_id,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
    }
    if scope:
        form["scope"] = scope
    data = urllib.parse.urlencode(form).encode("utf-8")
    try:
        req = urllib.request.Request(TOKEN_URL, data=data, method="POST")
        with urllib.request.urlopen(req, timeout=20) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
            return payload.get("access_token")
    except Exception:
        return None


def _token_candidates(refresh_token: str, client_id: str, scope: str) -> List[str]:
    """Thử token theo scope chuyên biệt trước, rồi tới token không scope (scope gốc)."""
    tokens: List[str] = []
    scoped = exchange_access_token(refresh_token, client_id, scope=scope)
    if scoped:
        tokens.append(scoped)
    plain = exchange_access_token(refresh_token, client_id)
    if plain and plain not in tokens:
        tokens.append(plain)
    return tokens


# ---------------------------------------------------------------------------
# IMAP
# ---------------------------------------------------------------------------

def _imap_connect(email_addr: str, refresh_token: str, client_id: str):
    """Mở kết nối IMAP XOAUTH2. Trả (mail, None) nếu OK, (None, lý-do) nếu fail."""
    tokens = _token_candidates(refresh_token, client_id, IMAP_SCOPE)
    if not tokens:
        return None, "token_exchange_failed"
    last_err: Optional[Exception] = None
    for token in tokens:
        mail = None
        try:
            auth_string = f"user={email_addr}\x01auth=Bearer {token}\x01\x01".encode("utf-8")
            mail = imaplib.IMAP4_SSL(IMAP_HOST, 993, timeout=25)
            mail.authenticate("XOAUTH2", lambda x: auth_string)
            return mail, None
        except Exception as e:  # thử token kế tiếp
            last_err = e
            if mail:
                try:
                    mail.logout()
                except Exception:
                    pass
    return None, f"imap_auth_failed: {last_err}"


def _imap_scan_once(mail, start_timestamp: float) -> Optional[Tuple[str, str, str]]:
    """Quét 1 vòng Inbox + Junk qua IMAP. Trả (otp, from, subject) hoặc None."""
    for folder in IMAP_FOLDERS:
        try:
            status, _ = mail.select(folder)
            if status != "OK":
                continue
            status, messages = mail.search(None, "ALL")
            if status != "OK" or not messages or not messages[0]:
                continue
            mids = messages[0].split()
            if not mids:
                continue

            # Quét các thư mới nhất
            for mid in reversed(mids[-5:]):
                res, msg_data = mail.fetch(mid, "(RFC822)")
                if res != "OK" or not msg_data or not msg_data[0]:
                    continue
                msg = email.message_from_bytes(msg_data[0][1])
                frm = str(msg.get("From", ""))
                subj = str(msg.get("Subject", ""))
                date_str = str(msg.get("Date", ""))

                # Bỏ qua thư cũ hơn thời điểm bắt đầu yêu cầu OTP (chênh lệch tối đa 30s)
                if date_str:
                    try:
                        msg_dt_obj = email.utils.parsedate_to_datetime(date_str)
                        if msg_dt_obj.tzinfo is None:
                            # Header dạng -0000 (không rõ múi giờ): hiểu là UTC,
                            # tránh lệch theo múi giờ local của máy chạy
                            msg_dt_obj = msg_dt_obj.replace(tzinfo=timezone.utc)
                        if msg_dt_obj.timestamp() < (start_timestamp - 30):
                            continue
                    except Exception:
                        pass

                raw_content = ""
                for part in msg.walk():
                    if part.get_content_type() in ["text/html", "text/plain"]:
                        payload = part.get_payload(decode=True)
                        if payload:
                            raw_content += payload.decode(errors="ignore")

                code = extract_meta_otp(frm, subj, raw_content)
                if code:
                    return code, frm, subj
        except Exception:
            continue
    return None


# ---------------------------------------------------------------------------
# Microsoft Graph
# ---------------------------------------------------------------------------

def _graph_request(access_token: str, path: str, params: Optional[Dict[str, str]] = None) -> Optional[dict]:
    url = GRAPH_BASE + path
    if params:
        url += "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {access_token}"}, method="GET")
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        return None


def _graph_connect(email_addr: str, refresh_token: str, client_id: str) -> Optional[str]:
    """Lấy access_token dùng được cho Graph (test bằng 1 request đọc thư thật)."""
    for token in _token_candidates(refresh_token, client_id, GRAPH_SCOPE):
        probe = _graph_request(token, "/me/mailFolders/inbox/messages", {"$top": "1", "$select": "id"})
        if probe is not None and "value" in probe:
            return token
    return None


def _parse_graph_datetime(value: str) -> Optional[float]:
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.timestamp()
    except Exception:
        return None


def _graph_scan_once(access_token: str, start_timestamp: float) -> Optional[Tuple[str, str, str]]:
    """Quét 1 vòng inbox + junkemail qua Graph. Trả (otp, from, subject) hoặc None."""
    params = {
        "$top": "10",
        "$orderby": "receivedDateTime desc",
        "$select": "subject,body,bodyPreview,receivedDateTime,from",
    }
    for folder in GRAPH_FOLDERS:
        data = _graph_request(access_token, f"/me/mailFolders/{folder}/messages", params)
        if not data:
            continue
        for msg in data.get("value", []):
            received = _parse_graph_datetime(str(msg.get("receivedDateTime", "")))
            if received is not None and received < (start_timestamp - 30):
                continue
            sender = (msg.get("from") or {}).get("emailAddress") or {}
            frm = f"{sender.get('name', '')} <{sender.get('address', '')}>"
            subj = str(msg.get("subject", ""))
            body = msg.get("body") or {}
            raw_content = str(body.get("content") or msg.get("bodyPreview") or "")
            code = extract_meta_otp(frm, subj, subj + "\n" + raw_content)
            if code:
                return code, frm, subj
    return None


# ---------------------------------------------------------------------------
# Polling dùng chung
# ---------------------------------------------------------------------------

def poll_meta_otp_detailed(
    email_addr: str,
    refresh_token: str,
    client_id: str,
    start_timestamp: float,
    timeout: int = 120,
    method: str = "auto",
) -> Optional[Tuple[str, str, str]]:
    """Chờ và lấy OTP Meta qua IMAP và/hoặc Graph.

    method: "auto" (mặc định, cả 2 đường), "imap" hoặc "graph".
    Trả (otp, from, subject) hoặc None khi hết timeout / cả 2 đường đều không dùng được.
    """
    method = (method or "auto").lower()
    if method not in ("auto", "imap", "graph"):
        raise ValueError(f"method không hợp lệ: {method!r} (chọn auto|imap|graph)")

    use_imap = method in ("auto", "imap")
    use_graph = method in ("auto", "graph")

    mail = None
    graph_token: Optional[str] = None
    errors: List[str] = []

    if use_imap:
        mail, err = _imap_connect(email_addr, refresh_token, client_id)
        if mail:
            print("    [✓] IMAP XOAUTH2: kết nối OK")
        else:
            errors.append(f"IMAP: {err}")
            if method == "imap":
                print(f"    [!] Lỗi kết nối IMAP XOAUTH2 {email_addr}: {err}")
    if use_graph:
        graph_token = _graph_connect(email_addr, refresh_token, client_id)
        if graph_token:
            print("    [✓] Microsoft Graph: kết nối OK")
        else:
            errors.append("Graph: token không dùng được cho Graph API")
            if method == "graph":
                print(f"    [!] Graph API không dùng được cho {email_addr} (token thiếu scope Mail.Read?)")

    if not mail and not graph_token:
        if method == "auto":
            print(f"    [!] Cả IMAP lẫn Graph đều không dùng được cho {email_addr}: {'; '.join(errors)}")
        return None

    deadline = time.time() + timeout
    try:
        while time.time() < deadline:
            if mail:
                hit = _imap_scan_once(mail, start_timestamp)
                if hit:
                    print("    [✓] Lấy OTP qua IMAP")
                    return hit
            if graph_token:
                hit = _graph_scan_once(graph_token, start_timestamp)
                if hit:
                    print("    [✓] Lấy OTP qua Microsoft Graph")
                    return hit
            time.sleep(1.2)
    finally:
        if mail:
            try:
                mail.logout()
            except Exception:
                pass
    return None


def poll_meta_otp(
    email_addr: str,
    refresh_token: str,
    client_id: str,
    start_timestamp: float,
    timeout: int = 120,
    method: str = "auto",
) -> Optional[str]:
    """Wrapper tương thích bản cũ: chỉ trả mã OTP (hoặc None)."""
    hit = poll_meta_otp_detailed(email_addr, refresh_token, client_id, start_timestamp, timeout, method)
    return hit[0] if hit else None
