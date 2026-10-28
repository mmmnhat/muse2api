#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Webmail OTP Viewer - Tương tự giao diện tuan89.site/mail/?email=...
Cho phép mở link xem hộp thư và lấy mã OTP tự động cho các tài khoản Hotmail đã đăng ký thành công.

Ví dụ sử dụng:
    http://localhost:8765/mail/?email=BusenbarkBlanch2303@hotmail.com
"""

import asyncio
import email
import imaplib
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
from aiohttp import web

PORT = 8765
HOST = "0.0.0.0"

SUCCESS_FILE = os.environ.get("MUSE_SUCCESS_FILE", os.path.expanduser("~/Downloads/muse_success.txt"))
ACCOUNTS_FILE = os.environ.get("MUSE_ACCOUNTS_FILE", "accounts.txt")

ACCOUNT_FILES = [
    SUCCESS_FILE,
    ACCOUNTS_FILE,
    os.path.expanduser("~/Downloads/muse_success.txt"),
    "muse_success.txt",
]

# Cache access tokens: { email: (access_token, expire_timestamp) }
TOKEN_CACHE = {}


def load_accounts():
    """Nạp danh sách tài khoản từ các file txt."""
    acc_map = {}
    for fpath in ACCOUNT_FILES:
        if os.path.isfile(fpath):
            with open(fpath, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or "|" not in line:
                        continue
                    parts = line.split("|")
                    if len(parts) >= 4:
                        em = parts[0].strip().lower()
                        if em not in acc_map:
                            acc_map[em] = {
                                "email": parts[0].strip(),
                                "pwd": parts[1].strip(),
                                "refresh_token": parts[2].strip(),
                                "client_id": parts[3].strip(),
                            }
    return acc_map


def get_access_token(email_addr: str, refresh_token: str, client_id: str) -> str:
    """Lấy Access Token qua Refresh Token (có cache 45 phút)."""
    now = time.time()
    if email_addr in TOKEN_CACHE:
        tok, exp = TOKEN_CACHE[email_addr]
        if now < exp:
            return tok

    token_url = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
    token_data = urllib.parse.urlencode({
        "client_id": client_id,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
    }).encode("utf-8")

    req = urllib.request.Request(token_url, data=token_data, method="POST")
    with urllib.request.urlopen(req, timeout=15) as resp:
        data = json.loads(resp.read().decode("utf-8"))
        tok = data["access_token"]
        expires_in = int(data.get("expires_in", 3600))
        TOKEN_CACHE[email_addr] = (tok, now + expires_in - 300)
        return tok


def fetch_emails_sync(email_addr: str, refresh_token: str, client_id: str, max_messages: int = 5):
    """Kết nối IMAP XOAUTH2 đọc thư trong Inbox và Junk."""
    access_token = get_access_token(email_addr, refresh_token, client_id)
    auth_string = f"user={email_addr}\x01auth=Bearer {access_token}\x01\x01".encode("utf-8")

    mail = imaplib.IMAP4_SSL("outlook.office365.com", 993, timeout=15)
    mail.authenticate("XOAUTH2", lambda x: auth_string)

    all_msgs = []
    latest_otp = None

    for folder in ["Inbox", "Junk"]:
        try:
            status, _ = mail.select(folder)
            if status != "OK":
                continue
            status, messages = mail.search(None, "ALL")
            mids = messages[0].split()
            for mid in reversed(mids[-max_messages:]):
                res, msg_data = mail.fetch(mid, "(RFC822)")
                if res != "OK" or not msg_data or not msg_data[0]:
                    continue
                msg = email.message_from_bytes(msg_data[0][1])
                frm = msg.get("From", "")
                subj = msg.get("Subject", "")
                date_str = msg.get("Date", "")

                # Decode header if needed
                subj_clean = ""
                for part, enc in email.header.decode_header(subj):
                    if isinstance(part, bytes):
                        subj_clean += part.decode(enc or "utf-8", errors="ignore")
                    else:
                        subj_clean += part

                raw_text = ""
                raw_html = ""
                for part in msg.walk():
                    ct = part.get_content_type()
                    if ct == "text/plain":
                        raw_text += part.get_payload(decode=True).decode(errors="ignore")
                    elif ct == "text/html":
                        raw_html += part.get_payload(decode=True).decode(errors="ignore")

                content_to_scan = raw_html or raw_text
                # Find OTP
                msg_otp = None
                m = re.search(r"Mã xác nhận\s+(\d{6})", content_to_scan)
                if m:
                    msg_otp = m.group(1)
                else:
                    m2 = re.search(r"letter-spacing:\s*2px[^>]*>(\d{6})<", content_to_scan)
                    if m2:
                        msg_otp = m2.group(1)
                    else:
                        for c in re.findall(r"\b\d{6}\b", content_to_scan):
                            if c not in ("141823", "000000"):
                                msg_otp = c
                                break

                if msg_otp and not latest_otp:
                    latest_otp = msg_otp

                all_msgs.append({
                    "id": mid.decode(),
                    "folder": folder,
                    "from": frm,
                    "subject": subj_clean or subj,
                    "date": date_str,
                    "otp": msg_otp,
                    "preview": (raw_text or re.sub(r"<[^>]+>", " ", raw_html))[:300].strip(),
                    "body_html": raw_html or f"<pre>{raw_text}</pre>",
                })
        except Exception:
            pass

    try:
        mail.logout()
    except Exception:
        pass

    return latest_otp, all_msgs


async def handle_api_inbox(request):
    """API trả về danh sách thư và mã OTP dạng JSON."""
    requested_email = request.query.get("email", "").strip().lower()
    if not requested_email:
        return web.json_response({"status": "error", "message": "Vui lòng truyền tham số ?email=..."}, status=400)

    acc_map = load_accounts()
    if requested_email not in acc_map:
        return web.json_response({
            "status": "error",
            "message": f"Không tìm thấy tài khoản {requested_email} trong danh sách hệ thống."
        }, status=404)

    acc = acc_map[requested_email]
    loop = asyncio.get_running_loop()
    try:
        otp, msgs = await loop.run_in_executor(
            None, fetch_emails_sync, acc["email"], acc["refresh_token"], acc["client_id"]
        )
        return web.json_response({
            "status": "success",
            "email": acc["email"],
            "latest_otp": otp,
            "messages": msgs,
            "timestamp": time.time()
        })
    except Exception as e:
        return web.json_response({
            "status": "error",
            "message": f"Lỗi lấy thư: {str(e)}"
        }, status=500)


async def handle_api_accounts(request):
    """API trả về danh sách các tài khoản có thể tra cứu."""
    acc_map = load_accounts()
    # Check success list
    success_emails = set()
    if os.path.isfile(SUCCESS_FILE):
        with open(SUCCESS_FILE, "r", encoding="utf-8") as f:
            for l in f:
                if "|" in l:
                    success_emails.add(l.split("|")[0].strip().lower())

    results = []
    for em, info in acc_map.items():
        results.append({
            "email": info["email"],
            "is_success": em in success_emails
        })
    results.sort(key=lambda x: (not x["is_success"], x["email"]))
    return web.json_response({"accounts": results})


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hộp Thư Tự Động Lấy OTP | Meta & Hotmail</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;700;800&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-base: #060913;
            --bg-card: rgba(14, 20, 34, 0.85);
            --bg-card-hover: rgba(22, 30, 48, 0.95);
            --border-color: rgba(255, 255, 255, 0.08);
            --border-glow: rgba(0, 242, 254, 0.4);
            --text-primary: #f8fafc;
            --text-secondary: #94a3b8;
            --text-muted: #64748b;
            --accent-cyan: #00f2fe;
            --accent-blue: #4facfe;
            --accent-purple: #8b5cf6;
            --accent-emerald: #10b981;
            --font-main: 'Plus Jakarta Sans', sans-serif;
            --font-mono: 'JetBrains Mono', monospace;
        }

        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            background-color: var(--bg-base);
            background-image: 
                radial-gradient(circle at 15% 15%, rgba(0, 242, 254, 0.12) 0%, transparent 45%),
                radial-gradient(circle at 85% 85%, rgba(139, 92, 246, 0.1) 0%, transparent 45%);
            color: var(--text-primary);
            font-family: var(--font-main);
            min-height: 100vh;
            padding: 30px 20px;
        }

        .container {
            max-width: 860px;
            margin: 0 auto;
        }

        header {
            text-align: center;
            margin-bottom: 28px;
        }
        .badge {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 6px 14px;
            background: rgba(0, 242, 254, 0.08);
            border: 1px solid rgba(0, 242, 254, 0.25);
            border-radius: 9999px;
            font-size: 13px;
            font-weight: 600;
            color: var(--accent-cyan);
            margin-bottom: 12px;
        }
        h1 {
            font-size: 28px;
            font-weight: 800;
            background: linear-gradient(135deg, #fff 30%, var(--accent-blue) 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 8px;
        }
        .subtitle {
            color: var(--text-secondary);
            font-size: 14px;
        }

        .search-box {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 16px;
            padding: 16px;
            display: flex;
            gap: 12px;
            margin-bottom: 24px;
            box-shadow: 0 8px 30px rgba(0,0,0,0.4);
        }
        .input-group {
            flex: 1;
            position: relative;
        }
        input[type="text"] {
            width: 100%;
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 12px 16px;
            color: #fff;
            font-size: 15px;
            font-family: var(--font-mono);
            outline: none;
            transition: all 0.2s;
        }
        input[type="text"]:focus {
            border-color: var(--accent-cyan);
            box-shadow: 0 0 12px rgba(0, 242, 254, 0.25);
        }
        .btn {
            background: linear-gradient(135deg, var(--accent-cyan), var(--accent-blue));
            color: #060913;
            border: none;
            border-radius: 12px;
            padding: 0 24px;
            font-size: 14px;
            font-weight: 700;
            cursor: pointer;
            transition: all 0.2s;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .btn:hover {
            opacity: 0.92;
            transform: translateY(-1px);
        }

        .quick-select {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-bottom: 24px;
            align-items: center;
        }
        .quick-label {
            font-size: 13px;
            color: var(--text-muted);
            margin-right: 4px;
        }
        .chip {
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 4px 10px;
            font-size: 12px;
            font-family: var(--font-mono);
            color: var(--text-secondary);
            cursor: pointer;
            transition: all 0.15s;
        }
        .chip:hover, .chip.active {
            border-color: var(--accent-cyan);
            color: var(--accent-cyan);
            background: rgba(0, 242, 254, 0.08);
        }

        /* OTP HERO CARD */
        .otp-card {
            background: linear-gradient(135deg, rgba(0, 242, 254, 0.1) 0%, rgba(139, 92, 246, 0.1) 100%);
            border: 1px solid rgba(0, 242, 254, 0.3);
            border-radius: 20px;
            padding: 24px;
            text-align: center;
            margin-bottom: 24px;
            display: none;
            position: relative;
            overflow: hidden;
        }
        .otp-card.show { display: block; }
        .otp-title {
            font-size: 13px;
            text-transform: uppercase;
            letter-spacing: 1.5px;
            color: var(--accent-cyan);
            font-weight: 700;
            margin-bottom: 8px;
        }
        .otp-code {
            font-family: var(--font-mono);
            font-size: 44px;
            font-weight: 800;
            color: #fff;
            letter-spacing: 6px;
            margin-bottom: 12px;
            text-shadow: 0 0 20px rgba(0, 242, 254, 0.5);
        }
        .btn-copy {
            background: rgba(255, 255, 255, 0.1);
            border: 1px solid rgba(255, 255, 255, 0.2);
            color: #fff;
            border-radius: 9999px;
            padding: 6px 18px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }
        .btn-copy:hover {
            background: #fff;
            color: #060913;
        }

        /* MESSAGES LIST */
        .card {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 16px;
            padding: 20px;
            margin-bottom: 14px;
            transition: all 0.2s;
        }
        .card:hover {
            border-color: rgba(255, 255, 255, 0.15);
            background: var(--bg-card-hover);
        }
        .msg-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 8px;
        }
        .msg-from {
            font-size: 13px;
            color: var(--accent-cyan);
            font-weight: 600;
        }
        .msg-date {
            font-size: 12px;
            color: var(--text-muted);
        }
        .msg-subject {
            font-size: 16px;
            font-weight: 700;
            color: #fff;
            margin-bottom: 8px;
        }
        .msg-preview {
            font-size: 13px;
            color: var(--text-secondary);
            line-height: 1.5;
        }

        .status-bar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 13px;
            color: var(--text-muted);
            margin-bottom: 16px;
        }
        .spinner {
            display: inline-block;
            width: 14px;
            height: 14px;
            border: 2px solid rgba(0, 242, 254, 0.3);
            border-top-color: var(--accent-cyan);
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            vertical-align: middle;
            margin-right: 6px;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div class="badge">⚡ LIVE OTP WEBMAIL VIEWER</div>
            <h1>Tra Cứu Mã OTP Tự Động</h1>
            <p class="subtitle">Xem hòm thư Hotmail & bắt mã xác thực Meta Muse AI trong thời gian thực</p>
        </header>

        <div class="search-box">
            <div class="input-group">
                <input type="text" id="emailInput" placeholder="Nhập địa chỉ email (VD: BusenbarkBlanch2303@hotmail.com)...">
            </div>
            <button class="btn" id="btnCheck" onclick="manualCheck()">
                <span>🔍 Lấy Mã</span>
            </button>
        </div>

        <div class="quick-select" id="quickSelect">
            <span class="quick-label">Gợi ý tài khoản:</span>
        </div>

        <div class="otp-card" id="otpCard">
            <div class="otp-title">Mã Xác Thực Gần Nhất (OTP)</div>
            <div class="otp-code" id="otpCode">------</div>
            <button class="btn-copy" id="btnCopy" onclick="copyOtp()">📋 Sao chép mã</button>
        </div>

        <div class="status-bar">
            <div id="statusText">Sẵn sàng</div>
            <div>
                <label style="cursor: pointer;">
                    <input type="checkbox" id="autoRefresh" checked> Tự động cập nhật (mỗi 5s)
                </label>
            </div>
        </div>

        <div id="messagesContainer">
            <!-- Messages render here -->
        </div>
    </div>

    <script>
        const urlParams = new URLSearchParams(window.location.search);
        const emailInput = document.getElementById('emailInput');
        const otpCard = document.getElementById('otpCard');
        const otpCode = document.getElementById('otpCode');
        const statusText = document.getElementById('statusText');
        const messagesContainer = document.getElementById('messagesContainer');
        const quickSelect = document.getElementById('quickSelect');

        let currentEmail = urlParams.get('email') || '';
        if (currentEmail) {
            emailInput.value = currentEmail;
        }

        // Load account chips
        fetch('/api/accounts')
            .then(r => r.json())
            .then(data => {
                (data.accounts || []).slice(0, 8).forEach(acc => {
                    const chip = document.createElement('div');
                    chip.className = 'chip' + (acc.email.toLowerCase() === currentEmail.toLowerCase() ? ' active' : '');
                    chip.textContent = acc.email.split('@')[0];
                    chip.title = acc.email;
                    chip.onclick = () => {
                        emailInput.value = acc.email;
                        selectEmail(acc.email);
                    };
                    quickSelect.appendChild(chip);
                });
            });

        function selectEmail(em) {
            currentEmail = em;
            const newUrl = new URL(window.location);
            newUrl.searchParams.set('email', em);
            window.history.pushState({}, '', newUrl);
            document.querySelectorAll('.chip').forEach(c => {
                c.classList.toggle('active', c.title.toLowerCase() === em.toLowerCase());
            });
            fetchInbox();
        }

        function manualCheck() {
            const em = emailInput.value.trim();
            if (em) {
                selectEmail(em);
            }
        }

        emailInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') manualCheck();
        });

        async function fetchInbox() {
            if (!currentEmail) return;
            statusText.innerHTML = '<span class="spinner"></span> Đang đồng bộ thư từ Microsoft Outlook...';
            try {
                const res = await fetch(`/api/inbox?email=${encodeURIComponent(currentEmail)}`);
                const data = await res.json();
                if (data.status === 'success') {
                    statusText.innerHTML = `Đã cập nhật lúc ${new Date().toLocaleTimeString()} • ${data.messages.length} thư`;
                    if (data.latest_otp) {
                        otpCode.textContent = data.latest_otp;
                        otpCard.classList.add('show');
                    } else {
                        otpCard.classList.remove('show');
                    }
                    renderMessages(data.messages);
                } else {
                    statusText.innerHTML = `<span style="color:#f43f5e;">Lỗi: ${data.message}</span>`;
                }
            } catch (err) {
                statusText.innerHTML = `<span style="color:#f43f5e;">Lỗi kết nối máy chủ</span>`;
            }
        }

        function renderMessages(msgs) {
            if (!msgs || msgs.length === 0) {
                messagesContainer.innerHTML = '<div class="card" style="text-align:center;color:#64748b;">Hòm thư trống hoặc chưa có thư mới</div>';
                return;
            }
            messagesContainer.innerHTML = msgs.map(m => `
                <div class="card">
                    <div class="msg-header">
                        <span class="msg-from">${m.from || 'Meta'}</span>
                        <span class="msg-date">${m.date}</span>
                    </div>
                    <div class="msg-subject">${m.subject}</div>
                    ${m.otp ? `<div style="display:inline-block;background:rgba(0,242,254,0.15);border:1px solid rgba(0,242,254,0.4);color:#00f2fe;font-family:var(--font-mono);font-size:16px;font-weight:700;padding:4px 12px;border-radius:6px;margin-bottom:8px;">OTP: ${m.otp}</div>` : ''}
                    <div class="msg-preview">${m.preview}</div>
                </div>
            `).join('');
        }

        function copyOtp() {
            const code = otpCode.textContent;
            navigator.clipboard.writeText(code).then(() => {
                const btn = document.getElementById('btnCopy');
                btn.textContent = '✓ Đã chép!';
                setTimeout(() => btn.textContent = '📋 Sao chép mã', 2000);
            });
        }

        if (currentEmail) {
            fetchInbox();
        }

        setInterval(() => {
            if (document.getElementById('autoRefresh').checked && currentEmail) {
                fetchInbox();
            }
        }, 5000);
    </script>
</body>
</html>
"""


async def handle_mail_page(request):
    """Trả về giao diện webmail HTML."""
    return web.Response(text=HTML_TEMPLATE, content_type="text/html")


def main():
    app = web.Application()
    app.router.add_get("/", handle_mail_page)
    app.router.add_get("/mail", handle_mail_page)
    app.router.add_get("/mail/", handle_mail_page)
    app.router.add_get("/api/inbox", handle_api_inbox)
    app.router.add_get("/api/accounts", handle_api_accounts)

    print(f"==================================================")
    print(f"[*] Khởi chạy Webmail OTP Viewer tại:")
    print(f"    👉 http://localhost:{PORT}/mail/?email=<dia_chi_email>")
    print(f"==================================================")
    web.run_app(app, host=HOST, port=PORT)


if __name__ == "__main__":
    main()
