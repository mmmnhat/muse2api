#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tool tự động đăng ký tài khoản Muse AI (Meta) hàng loạt từ file Hotmail OAuth2,
hỗ trợ xoay Proxy / 4G, phân loại triệt để 3 nhóm (Thành Công, Chờ Duyệt Waitlist, Thất Bại),
và tự động nạp thẳng tài khoản hợp lệ vào Pool muse2api.

Định dạng file đầu vào:
    email|password|refresh_token|client_id

Các file xuất ra (mặc định tại --out-dir .):
    - muse_success.txt & muse_success.json   (Tài khoản có hatch_vml, VM hoạt động, không bị waitlist)
    - muse_waitlist.txt & muse_waitlist.json (Tài khoản Meta tạo xong nhưng bị Meta đưa vào Waitlist do IP)
    - muse_failed.txt & muse_failed.json     (Tài khoản lỗi: sai OTP, timeout, lỗi mạng)
"""

from __future__ import annotations

import argparse
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
from typing import Dict, List, Optional, Tuple

try:
    from playwright.async_api import async_playwright
except ImportError:
    print("[!] Cần cài đặt Playwright: .venv/bin/pip install playwright")
    sys.exit(1)


def parse_name_from_email(email_addr: str) -> Tuple[str, str]:
    """Tách tên và họ từ email dạng AdrianceMadie870@hotmail.com"""
    username = email_addr.split("@")[0]
    clean = re.sub(r"\d+", "", username)
    parts = re.findall(r"[A-Z][a-z]*", clean)
    if len(parts) >= 2:
        return parts[0], parts[1]
    elif len(parts) == 1:
        return parts[0], "User"
    return username[:6].capitalize(), "AI"


def parse_proxy(proxy_str: Optional[str]) -> Optional[Dict[str, str]]:
    """Phân tích cấu hình proxy từ chuỗi đầu vào theo các định dạng phổ biến:
    - http://user:pass@host:port hoặc socks5://...
    - host:port:user:pass
    - host:port
    """
    if not proxy_str:
        return None
    p = proxy_str.strip()
    if not p:
        return None

    if "://" in p:
        parsed = urllib.parse.urlparse(p)
        cfg: Dict[str, str] = {"server": f"{parsed.scheme}://{parsed.hostname}:{parsed.port}"}
        if parsed.username:
            cfg["username"] = urllib.parse.unquote(parsed.username)
        if parsed.password:
            cfg["password"] = urllib.parse.unquote(parsed.password)
        return cfg

    parts = p.split(":")
    if len(parts) == 4:
        return {
            "server": f"http://{parts[0]}:{parts[1]}",
            "username": parts[2],
            "password": parts[3]
        }
    elif len(parts) == 2:
        return {"server": f"http://{parts[0]}:{parts[1]}"}

    return {"server": p}


def poll_meta_otp(email_addr: str, refresh_token: str, client_id: str, start_timestamp: float, timeout: int = 120) -> Optional[str]:
    """Kết nối Hotmail IMAP XOAUTH2 và quét mã OTP 6 chữ số từ Meta trong khoảng thời gian timeout."""
    token_url = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
    token_data = urllib.parse.urlencode({
        "client_id": client_id,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
    }).encode("utf-8")

    access_token = None
    try:
        req = urllib.request.Request(token_url, data=token_data, method="POST")
        with urllib.request.urlopen(req, timeout=20) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            access_token = data.get("access_token")
    except Exception as e:
        print(f"    [!] Lỗi đổi OAuth2 access token cho {email_addr}: {e}")
        return None

    if not access_token:
        return None

    auth_string = f"user={email_addr}\x01auth=Bearer {access_token}\x01\x01".encode("utf-8")
    mail = None
    try:
        mail = imaplib.IMAP4_SSL("outlook.office365.com", 993, timeout=25)
        mail.authenticate("XOAUTH2", lambda x: auth_string)
    except Exception as e:
        print(f"    [!] Lỗi kết nối IMAP XOAUTH2 {email_addr}: {e}")
        if mail:
            try:
                mail.logout()
            except Exception:
                pass
        return None

    deadline = time.time() + timeout
    while time.time() < deadline:
        for folder in ["Inbox", "Junk"]:
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

                    if any(k in frm or k in subj for k in ["Meta", "meta.com", "Xác nhận", "Confirm", "code", "Mã"]):
                        raw_content = ""
                        for part in msg.walk():
                            if part.get_content_type() in ["text/html", "text/plain"]:
                                payload = part.get_payload(decode=True)
                                if payload:
                                    raw_content += payload.decode(errors="ignore")

                        # Mẫu 1: "Mã xác nhận 123456"
                        m = re.search(r"Mã xác nhận[:\s]+(\d{6})", raw_content)
                        if m:
                            try: mail.logout()
                            except Exception: pass
                            return m.group(1)

                        # Mẫu 2: HTML letter-spacing hoặc font-size của Meta email
                        m2 = re.search(r"letter-spacing:[^>]*>(\d{6})<", raw_content)
                        if m2:
                            try: mail.logout()
                            except Exception: pass
                            return m2.group(1)

                        # Mẫu 3: "is your Meta confirmation code"
                        m3 = re.search(r"(\d{6})\s+is your Meta", raw_content)
                        if m3:
                            try: mail.logout()
                            except Exception: pass
                            return m3.group(1)

                        # Mẫu 4: Mọi chuỗi 6 chữ số hợp lệ
                        candidates = re.findall(r"\b\d{6}\b", raw_content)
                        for code in candidates:
                            if code not in ("141823", "000000", "123456"):
                                try: mail.logout()
                                except Exception: pass
                                return code
            except Exception:
                pass
        time.sleep(3)

    try:
        mail.logout()
    except Exception:
        pass
    return None


def upload_to_muse2api(base_url: str, api_key: str, label: str, cookies: Dict[str, str], cookies_exp: Dict[str, int]) -> bool:
    """Nạp cookie vào muse2api pool qua endpoint /admin/accounts"""
    payload = {
        "label": label,
        "cookies": cookies,
        "expires": {k: v for k, v in cookies_exp.items() if v > 0},
    }
    req = urllib.request.Request(
        base_url.rstrip("/") + "/admin/accounts",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        },
        method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            res = json.loads(resp.read().decode("utf-8"))
            return bool(res.get("added"))
    except Exception as e:
        print(f"    [!] Lỗi upload vào muse2api: {e}")
        return False


def save_record(out_dir: str, acc_data: Dict, res: Dict):
    """Lưu kết quả phân loại thành công / waitlist / thất bại vào cả JSON và TXT,
    đồng thời khử trùng lặp và chuyển trạng thái nếu tài khoản được xử lý lại."""
    os.makedirs(out_dir, exist_ok=True)
    email_addr = acc_data["email"].strip()
    email_lower = email_addr.lower()
    raw_line = f"{email_addr}|{acc_data['pwd']}|{acc_data['refresh_token']}|{acc_data['client_id']}"

    success_json = os.path.join(out_dir, "muse_success.json")
    success_txt = os.path.join(out_dir, "muse_success.txt")
    waitlist_json = os.path.join(out_dir, "muse_waitlist.json")
    waitlist_txt = os.path.join(out_dir, "muse_waitlist.txt")
    failed_json = os.path.join(out_dir, "muse_failed.json")
    failed_txt = os.path.join(out_dir, "muse_failed.txt")

    status = res.get("status")

    def read_json_dict(path: str) -> Dict[str, Dict]:
        if not os.path.isfile(path):
            return {}
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return {item["email"].lower(): item for item in data if isinstance(item, dict) and "email" in item}
                elif isinstance(data, dict):
                    return data
        except Exception:
            return {}
        return {}

    def write_json_dict(path: str, data_dict: Dict[str, Dict]):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(list(data_dict.values()), f, ensure_ascii=False, indent=2)

    def read_txt_lines(path: str) -> List[str]:
        if not os.path.isfile(path):
            return []
        with open(path, "r", encoding="utf-8") as f:
            return [line.strip() for line in f if line.strip()]

    def write_txt_lines(path: str, lines: List[str]):
        with open(path, "w", encoding="utf-8") as f:
            for line in lines:
                f.write(line + "\n")

    if status == "success":
        # 1. Update Success files
        s_dict = read_json_dict(success_json)
        s_dict[email_lower] = res
        write_json_dict(success_json, s_dict)

        s_lines = read_txt_lines(success_txt)
        s_emails = {line.split("|")[0].strip().lower() for line in s_lines}
        if email_lower not in s_emails:
            s_lines.append(raw_line)
            write_txt_lines(success_txt, s_lines)

        # 2. Xóa khỏi waitlist và failed nếu trước đó từng nằm ở đó
        w_dict = read_json_dict(waitlist_json)
        if email_lower in w_dict:
            del w_dict[email_lower]
            write_json_dict(waitlist_json, w_dict)
            w_lines = [l for l in read_txt_lines(waitlist_txt) if l.split("|")[0].strip().lower() != email_lower]
            write_txt_lines(waitlist_txt, w_lines)

        f_dict = read_json_dict(failed_json)
        if email_lower in f_dict:
            del f_dict[email_lower]
            write_json_dict(failed_json, f_dict)
            f_lines = [l for l in read_txt_lines(failed_txt) if l.split("|")[0].strip().lower() != email_lower]
            write_txt_lines(failed_txt, f_lines)

    elif status == "waitlist":
        # 1. Update Waitlist files
        w_dict = read_json_dict(waitlist_json)
        w_dict[email_lower] = res
        write_json_dict(waitlist_json, w_dict)

        w_lines = read_txt_lines(waitlist_txt)
        w_emails = {line.split("|")[0].strip().lower() for line in w_lines}
        if email_lower not in w_emails:
            w_lines.append(f"{raw_line}|{res.get('reason', 'waitlist')}")
            write_txt_lines(waitlist_txt, w_lines)

        # 2. Xóa khỏi failed nếu có
        f_dict = read_json_dict(failed_json)
        if email_lower in f_dict:
            del f_dict[email_lower]
            write_json_dict(failed_json, f_dict)
            f_lines = [l for l in read_txt_lines(failed_txt) if l.split("|")[0].strip().lower() != email_lower]
            write_txt_lines(failed_txt, f_lines)

    else:
        # Status failed
        f_dict = read_json_dict(failed_json)
        f_dict[email_lower] = res
        write_json_dict(failed_json, f_dict)

        f_lines = read_txt_lines(failed_txt)
        f_emails = {line.split("|")[0].strip().lower() for line in f_lines}
        if email_lower not in f_emails:
            f_lines.append(f"{raw_line}|{res.get('reason', 'failed')}")
            write_txt_lines(failed_txt, f_lines)


async def register_single_account(
    acc_data: Dict,
    proxy_server: Optional[str] = None,
    muse2api_base: str = "http://127.0.0.1:18610",
    muse2api_key: str = "",
    out_dir: str = ".",
    headless: bool = True
) -> Dict:
    """Quy trình đăng ký 1 tài khoản Muse AI tự động bằng Playwright Chrome."""
    email_addr = acc_data["email"].strip()
    pwd = acc_data["pwd"].strip()
    refresh_token = acc_data["refresh_token"].strip()
    client_id = acc_data["client_id"].strip()

    first_name, last_name = parse_name_from_email(email_addr)
    print(f"\n==================================================")
    print(f"[*] Bắt đầu xử lý: {email_addr} ({first_name} {last_name})")

    launch_args = {
        "headless": headless,
        "args": [
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-infobars",
            "--window-size=1280,800",
        ],
        "ignore_default_args": ["--enable-automation"]
    }
    chrome_path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    if os.path.isfile(chrome_path):
        launch_args["executable_path"] = chrome_path

    proxy_cfg = parse_proxy(proxy_server)
    if proxy_cfg:
        launch_args["proxy"] = proxy_cfg
        print(f"  -> Sử dụng Proxy: {proxy_cfg['server']}")

    async with async_playwright() as p:
        browser = await p.chromium.launch(**launch_args)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 800},
            locale="en-US",
            timezone_id="America/New_York",
        )
        # Ẩn cờ webdriver automation để qua mặt Meta Bot Detection
        await context.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            });
            window.chrome = {
                runtime: {}
            };
        """)

        page = await context.new_page()

        try:
            print("  1. Mở https://muse.ai ...")
            await page.goto("https://muse.ai", wait_until="domcontentloaded", timeout=35000)
            await page.wait_for_timeout(2500)

            # Kiểm tra ô nhập email trên form đăng nhập / đăng ký
            email_input = page.locator('input[autocomplete*="username"], input[type="email"], input[name="email"], input[placeholder*="email" i], input[placeholder*="di động" i]').first
            if await email_input.count() == 0:
                # Nếu trang chuyển đến landing page có nút đăng nhập riêng
                login_btn = page.locator('button:has-text("Đăng nhập"), button:has-text("Log in"), button:has-text("Sign in")').first
                if await login_btn.count() > 0:
                    await login_btn.click()
                    await page.wait_for_timeout(2000)
                    email_input = page.locator('input[autocomplete*="username"], input[type="email"], input[name="email"]').first

            if await email_input.count() == 0:
                print("  [!] Không tìm thấy ô nhập email trên màn hình!")
                await browser.close()
                res = {"email": email_addr, "status": "failed", "reason": "no_email_input"}
                save_record(out_dir, acc_data, res)
                return res

            await email_input.click()
            await email_input.press_sequentially(email_addr, delay=45)
            await page.wait_for_timeout(600)

            submit_btn = page.locator('button[type="submit"], button:has-text("Tiếp tục"), button:has-text("Continue")').first
            start_ts = time.time()
            await submit_btn.click()
            print("  2. Đã bấm gửi email. Bắt đầu quét mã OTP từ Hotmail qua IMAP XOAUTH2...")

            # Quét OTP từ Hotmail qua IMAP XOAUTH2
            loop = asyncio.get_running_loop()
            otp = await loop.run_in_executor(None, poll_meta_otp, email_addr, refresh_token, client_id, start_ts, 120)
            if not otp:
                print("  [!] Hết thời gian chờ OTP (120s) không nhận được mail từ Meta!")
                await browser.close()
                res = {"email": email_addr, "status": "failed", "reason": "otp_timeout"}
                save_record(out_dir, acc_data, res)
                return res

            print(f"  3. Nhận được mã OTP: {otp} -> Đang nhập vào form...")
            otp_input = page.locator('input[maxlength="6"], input[inputmode="numeric"], input[type="text"]').first
            try:
                await otp_input.wait_for(state="visible", timeout=20000)
            except Exception:
                cur_text = await page.locator("body").inner_text()
                print(f"  [!] Không thấy ô nhập OTP. Nội dung:\n{cur_text[:250]}")
                await browser.close()
                res = {"email": email_addr, "status": "failed", "reason": "otp_input_missing"}
                save_record(out_dir, acc_data, res)
                return res

            await otp_input.click()
            await otp_input.press_sequentially(otp, delay=120)
            await page.wait_for_timeout(6000)

            # Trường hợp Meta yêu cầu nhập mật khẩu (nếu tài khoản Meta đã tồn tại)
            pwd_input = page.locator('input[type="password"]').first
            if await pwd_input.count() > 0 and await pwd_input.is_visible():
                print("  -> Phát hiện yêu cầu mật khẩu Meta, đang điền mật khẩu...")
                await pwd_input.fill(pwd)
                await page.wait_for_timeout(400)
                confirm_pwd_btn = page.locator('button[type="submit"], button:has-text("Đăng nhập"), button:has-text("Log in")').first
                if await confirm_pwd_btn.count() > 0:
                    await confirm_pwd_btn.click()
                    await page.wait_for_timeout(6000)

            # Kiểm tra màn hình Profile (Hoàn tất quy trình tạo tài khoản Meta)
            body_text = await page.locator("body").inner_text()
            if any(k in body_text for k in ["Hoàn tất quy trình tạo tài khoản", "Tên bắt buộc", "Complete your profile", "First name"]):
                print("  4. Điền thông tin Profile (Họ tên, ngày sinh)...")
                inputs = await page.locator('input:not([type="checkbox"]):not([type="hidden"])').all()
                if len(inputs) >= 2:
                    await inputs[0].fill(first_name)
                    await inputs[1].fill(last_name)

                selects = await page.locator('select').all()
                if len(selects) >= 3:
                    await selects[0].select_option(value="15")
                    await selects[1].select_option(index=6)
                    await selects[2].select_option(value="1998")

                confirm_btn = page.locator('button[type="submit"], button:has-text("Xác nhận"), button:has-text("Confirm")').first
                await confirm_btn.click()
                print("  5. Đã bấm xác nhận tạo tài khoản Meta. Đang đợi xử lý...")
                await page.wait_for_timeout(8000)

            # Đợi tải xong và kiểm tra URL / Cookie
            await page.wait_for_timeout(5000)
            current_url = page.url
            print(f"  -> URL đích: {current_url}")

            cookies_list = await context.cookies()
            cookies_dict = {c["name"]: c["value"] for c in cookies_list}
            cookies_exp = {c["name"]: int(c.get("expires", 0)) for c in cookies_list}

            has_sess = "hatch_sess" in cookies_dict
            has_gw = "hatch_gw" in cookies_dict
            has_vml = "hatch_vml" in cookies_dict
            is_waitlist = "/access" in current_url

            # ĐÁNH GIÁ KẾT QUẢ
            if has_sess and has_vml and not is_waitlist:
                print("  [✓] ĐĂNG KÝ THÀNH CÔNG! Tài khoản có đầy đủ phiên (hatch_sess + hatch_vml).")
                if muse2api_base and muse2api_key:
                    uploaded = upload_to_muse2api(muse2api_base, muse2api_key, email_addr, cookies_dict, cookies_exp)
                    if uploaded:
                        print("  [✓] Đã tự động nạp thành công vào Pool muse2api!")
                res = {
                    "email": email_addr,
                    "status": "success",
                    "url": current_url,
                    "cookies": cookies_dict,
                    "expires": cookies_exp
                }
            elif is_waitlist or (has_sess and not has_vml):
                print("  [!] TÀI KHOẢN VÀO WAITLIST: Đã xác thực Meta nhưng Muse AI giới hạn vùng IP (thiếu hatch_vml)!")
                res = {
                    "email": email_addr,
                    "status": "waitlist",
                    "reason": "waitlist_geo_blocked (Cần IP US Residential / 4G sạch để cấp quyền hatch_vml)",
                    "url": current_url,
                    "cookies": cookies_dict
                }
            else:
                print("  [!] Thất bại: Không nhận được cookie phiên làm việc.")
                res = {
                    "email": email_addr,
                    "status": "failed",
                    "reason": "no_session_cookie",
                    "url": current_url
                }

            save_record(out_dir, acc_data, res)
            await browser.close()
            return res

        except Exception as e:
            print(f"  [!] Ngoại lệ trong quá trình đăng ký: {e}")
            await browser.close()
            res = {"email": email_addr, "status": "failed", "reason": str(e)}
            save_record(out_dir, acc_data, res)
            return res


def trigger_ip_rotate(rotate_url: str, wait_secs: int = 5):
    """Gửi request tới API xoay IP (thường dùng cho proxy 4G xoay)."""
    try:
        print(f"[*] Đang gửi yêu cầu đổi IP: {rotate_url} ...")
        req = urllib.request.Request(rotate_url, headers={"User-Agent": "curl/7.68.0"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            content = resp.read().decode("utf-8", errors="ignore")
            print(f"    -> Phản hồi đổi IP: {content[:150]}")
        print(f"[*] Đợi {wait_secs} giây để IP mới ổn định...")
        time.sleep(wait_secs)
    except Exception as e:
        print(f"[!] Lỗi khi gọi API đổi IP: {e}")


async def main():
    parser = argparse.ArgumentParser(
        description="Tool tự động đăng ký Muse AI từ Hotmail OAuth2, hỗ trợ Proxy/4G xoay & nạp pool."
    )
    default_file = os.environ.get("MUSE_ACCOUNTS_FILE", "accounts.txt")
    default_out_dir = os.environ.get("MUSE_OUT_DIR", os.path.expanduser("~/Downloads") if os.path.isdir(os.path.expanduser("~/Downloads")) else ".")

    parser.add_argument("--file", default=default_file, help="File txt danh sách acc (mặc định accounts.txt)")
    parser.add_argument("--proxy", default=None, help="Proxy đơn lẻ (VD: http://user:pass@host:port hoặc socks5://...)")
    parser.add_argument("--proxy-file", default=None, help="File chứa danh sách proxy, mỗi dòng 1 proxy để xoay vòng")
    parser.add_argument("--rotate-url", default=None, help="API URL để đổi IP proxy 4G trước mỗi tài khoản")
    parser.add_argument("--rotate-wait", type=int, default=5, help="Số giây chờ sau khi gọi API đổi IP (mặc định 5)")
    parser.add_argument("--prompt-rotate", action="store_true", help="Dừng lại hỏi người dùng bật/tắt 4G trước mỗi acc")
    parser.add_argument("--head", action="store_true", help="Mở giao diện trình duyệt trực quan (mặc định chạy headless)")
    parser.add_argument("--out-dir", default=default_out_dir, help="Thư mục xuất file kết quả (mặc định ~/Downloads hoặc .)")
    parser.add_argument("--base", default=os.environ.get("MUSE2API_BASE", "http://127.0.0.1:18610"), help="Địa chỉ muse2api pool")
    parser.add_argument("--key", default=os.environ.get("MUSE2API_KEY", ""), help="API Key muse2api")
    parser.add_argument("--limit", type=int, default=50, help="Số lượng tài khoản muốn chạy")
    parser.add_argument("--start", type=int, default=0, help="Vị trí bắt đầu trong file (0-indexed)")
    parser.add_argument("--force", action="store_true", help="Bỏ qua kiểm tra tài khoản đã thành công, chạy lại từ đầu")
    args = parser.parse_args()

    if not os.path.isfile(args.file):
        print(f"[!] Không tìm thấy file tài khoản: {args.file}")
        return

    accounts = []
    with open(args.file, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or "|" not in line:
                continue
            parts = line.split("|")
            if len(parts) >= 4:
                accounts.append({
                    "email": parts[0].strip(),
                    "pwd": parts[1].strip(),
                    "refresh_token": parts[2].strip(),
                    "client_id": parts[3].strip(),
                })

    print(f"[*] Tổng số tài khoản đọc được: {len(accounts)}")
    selected = accounts[args.start : args.start + args.limit]

    # Đọc danh sách proxy nếu có file
    proxy_list: List[str] = []
    if args.proxy_file and os.path.isfile(args.proxy_file):
        with open(args.proxy_file, "r", encoding="utf-8") as pf:
            proxy_list = [l.strip() for l in pf if l.strip()]
        print(f"[*] Đã nạp {len(proxy_list)} proxy từ {args.proxy_file}")

    # Lấy danh sách đã thành công để bỏ qua nếu không dùng cờ --force
    existing_success = set()
    if not args.force:
        success_txt = os.path.join(args.out_dir, "muse_success.txt")
        if os.path.isfile(success_txt):
            with open(success_txt, "r", encoding="utf-8") as f:
                for line in f:
                    existing_success.add(line.split("|")[0].strip().lower())

    print(f"[*] Đã có {len(existing_success)} tài khoản thành công từ trước.")
    print(f"[*] Bắt đầu xử lý {len(selected)} tài khoản...\n")

    for idx, acc in enumerate(selected, 1):
        email_clean = acc["email"].lower()
        if not args.force and email_clean in existing_success:
            print(f"[{idx}/{len(selected)}] Tài khoản {acc['email']} đã thành công từ trước -> Bỏ qua.")
            continue

        # Chọn proxy
        current_proxy = args.proxy
        if proxy_list:
            current_proxy = proxy_list[(idx - 1) % len(proxy_list)]

        # Xoay IP tự động qua API nếu có
        if args.rotate_url:
            trigger_ip_rotate(args.rotate_url, args.rotate_wait)

        # Dừng thủ công cho người dùng đổi IP 4G (bật tắt chế độ máy bay)
        if args.prompt_rotate and idx > 1:
            input(f"\n[?] Bật/tắt 4G trên điện thoại để lấy IP mới rồi nhấn [ENTER] để tiếp tục...")

        print(f"[{idx}/{len(selected)}] Tiến hành đăng ký: {acc['email']}")
        await register_single_account(
            acc,
            proxy_server=current_proxy,
            muse2api_base=args.base,
            muse2api_key=args.key,
            out_dir=args.out_dir,
            headless=not args.head
        )
        await asyncio.sleep(2)

    print("\n==================================================")
    print("[*] HOÀN TẤT TOÀN BỘ QUÁ TRÌNH!")
    print(f"    - Thành công: {os.path.join(args.out_dir, 'muse_success.txt')}")
    print(f"    - Chờ duyệt:  {os.path.join(args.out_dir, 'muse_waitlist.txt')}")
    print(f"    - Thất bại:   {os.path.join(args.out_dir, 'muse_failed.txt')}")


if __name__ == "__main__":
    asyncio.run(main())
