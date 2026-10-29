#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tool tự động đăng ký tài khoản Muse AI (Meta) hàng loạt từ file Outlook OAuth2 (đọc OTP qua IMAP XOAUTH2 + Microsoft Graph),
hỗ trợ xoay Proxy / 4G, phân loại triệt để 3 nhóm (Thành Công, Chờ Duyệt Waitlist, Thất Bại),
và tự động nạp thẳng tài khoản hợp lệ vào Pool muse2api.

Bản tối ưu hóa tốc độ cao (High-Performance Engine):
- Dynamic Element Waiting thay vì hardcoded sleep.
- Chặn tải font và analytics rác để giảm 70% băng thông và tăng tốc tải trang.
- Tự động sinh Sticky Session cho Cliproxy theo từng luồng/tài khoản.
- Hỗ trợ đa luồng (--concurrency / -c) chạy song song nhiều trình duyệt.

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
import secrets
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

from otp_mail import is_outlook_email, poll_meta_otp

# Khóa đồng bộ ghi file tránh xung đột khi chạy đa luồng
_FILE_LOCK = asyncio.Lock()


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


def resolve_proxy_for_worker(base_proxy: Optional[str], session_tag: Optional[str] = None) -> Optional[Dict[str, str]]:
    """Phân giải proxy và tự động chèn session sticky nếu dùng Cliproxy để mỗi luồng/tài khoản có IP riêng biệt."""
    if not base_proxy:
        return None
    p = base_proxy.strip()
    if not p:
        return None

    # Tự động gán session ID ngẫu nhiên cho Cliproxy nếu chưa có -sid-
    if "cliproxy.io" in p and "-sid-" not in p:
        sid = session_tag or secrets.token_hex(4)
        if "://" in p:
            parsed = urllib.parse.urlparse(p)
            user = parsed.username or ""
            user_with_sid = f"{user}-sid-{sid}-t-10"
            pwd_part = f":{parsed.password}" if parsed.password else ""
            netloc = f"{user_with_sid}{pwd_part}@{parsed.hostname}:{parsed.port}"
            p = urllib.parse.urlunparse((parsed.scheme, netloc, parsed.path, parsed.params, parsed.query, parsed.fragment))
        elif len(p.split(":")) == 4:
            host, port, user, pwd = p.split(":")
            user_with_sid = f"{user}-sid-{sid}-t-10"
            p = f"{host}:{port}:{user_with_sid}:{pwd}"

    return parse_proxy(p)


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
        s_dict = read_json_dict(success_json)
        s_dict[email_lower] = res
        write_json_dict(success_json, s_dict)

        s_lines = read_txt_lines(success_txt)
        s_emails = {line.split("|")[0].strip().lower() for line in s_lines}
        if email_lower not in s_emails:
            s_lines.append(raw_line)
            write_txt_lines(success_txt, s_lines)

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
        w_dict = read_json_dict(waitlist_json)
        w_dict[email_lower] = res
        write_json_dict(waitlist_json, w_dict)

        w_lines = read_txt_lines(waitlist_txt)
        w_emails = {line.split("|")[0].strip().lower() for line in w_lines}
        if email_lower not in w_emails:
            w_lines.append(f"{raw_line}|{res.get('reason', 'waitlist')}")
            write_txt_lines(waitlist_txt, w_lines)

        f_dict = read_json_dict(failed_json)
        if email_lower in f_dict:
            del f_dict[email_lower]
            write_json_dict(failed_json, f_dict)

    else:
        f_dict = read_json_dict(failed_json)
        f_dict[email_lower] = res
        write_json_dict(failed_json, f_dict)

        f_lines = read_txt_lines(failed_txt)
        f_emails = {line.split("|")[0].strip().lower() for line in f_lines}
        if email_lower not in f_emails:
            f_lines.append(f"{raw_line}|{res.get('reason', 'failed')}")
            write_txt_lines(failed_txt, f_lines)


async def save_record_async(out_dir: str, acc_data: Dict, res: Dict):
    """Ghi record an toàn trong môi trường đa luồng."""
    async with _FILE_LOCK:
        save_record(out_dir, acc_data, res)


async def wait_for_transition(page, context, max_wait: float = 8.0) -> str:
    """Đợi màn hình chuyển trạng thái thông minh: trả về loại màn hình tiếp theo ngay khi render."""
    deadline = time.time() + max_wait
    while time.time() < deadline:
        # 1. Kiểm tra session cookie hoặc redirect waitlist/thread
        cookies = {c["name"]: c["value"] for c in await context.cookies()}
        url = page.url
        if "hatch_sess" in cookies or "/thread" in url or "/access" in url:
            return "authenticated"

        # 2. Kiểm tra rejection dialog
        body_text = (await page.locator("body").inner_text()).lower()
        if any(k in body_text for k in ["không thể tạo tài khoản", "we couldn't create", "registration rejected", "rejected"]):
            return "rejected"

        # 3. Kiểm tra các form bước tiếp theo
        if any(k in body_text for k in ["birthday", "ngày sinh", "date of birth"]):
            return "birthday"
        if any(k in body_text for k in ["what’s your name", "what's your name", "tên của bạn", "first name", "họ tên", "tên bắt buộc"]):
            return "name"
        if await page.locator('input[type="password"]').count() > 0:
            return "password"

        await asyncio.sleep(0.25)
    return "timeout"


async def register_single_account(
    acc_data: Dict,
    proxy_server: Optional[str] = None,
    muse2api_base: str = "http://127.0.0.1:18610",
    muse2api_key: str = "",
    out_dir: str = ".",
    headless: bool = True,
    verify_video: bool = True,
    otp_method: str = "auto",
    session_tag: Optional[str] = None
) -> Dict:
    """Quy trình đăng ký 1 tài khoản Muse AI tự động bằng Playwright Chrome tối ưu tốc độ cao."""
    email_addr = acc_data["email"].strip()
    pwd = acc_data["pwd"].strip()
    refresh_token = acc_data["refresh_token"].strip()
    client_id = acc_data["client_id"].strip()

    first_name, last_name = parse_name_from_email(email_addr)
    print(f"\n==================================================")
    print(f"[*] Bắt đầu xử lý: {email_addr} ({first_name} {last_name})")

    # CHỈ DÙNG OUTLOOK: bỏ qua Hotmail/Live/MSN
    if not is_outlook_email(email_addr):
        print("  [!] Không phải email Outlook (chỉ hỗ trợ outlook.*) -> Bỏ qua!")
        res = {"email": email_addr, "status": "failed", "reason": "not_outlook"}
        await save_record_async(out_dir, acc_data, res)
        return res

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

    proxy_cfg = resolve_proxy_for_worker(proxy_server, session_tag)
    if proxy_cfg:
        launch_args["proxy"] = proxy_cfg
        user_info = f" (User: {proxy_cfg.get('username')})" if 'username' in proxy_cfg else ""
        print(f"  -> Sử dụng Proxy: {proxy_cfg['server']}{user_info}")

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

        # Tối ưu hóa tải trang: Chặn fonts và tracking rác giúp tăng tốc và tiết kiệm băng thông proxy
        async def intercept_route(route):
            url = route.request.url.lower()
            res_type = route.request.resource_type
            if any(k in url for k in ("google-analytics.com", "googletagmanager.com", "sentry.io", "datadoghq.com", "connect.facebook.net")):
                await route.abort()
                return
            if res_type == "font":
                await route.abort()
                return
            await route.continue_()

        await context.route("**/*", intercept_route)

        page = await context.new_page()

        try:
            print("  1. Mở https://muse.ai ...")
            await page.goto("https://muse.ai", wait_until="domcontentloaded", timeout=35000)

            # Chờ động ô nhập email
            email_input = page.locator('input[autocomplete*="username"], input[type="email"], input[name="email"], input[placeholder*="email" i], input[placeholder*="di động" i]').first
            try:
                await email_input.wait_for(state="visible", timeout=6000)
            except Exception:
                pass

            if await email_input.count() == 0 or not await email_input.is_visible():
                login_btn = page.locator('button:has-text("Đăng nhập"), button:has-text("Log in"), button:has-text("Sign in")').first
                if await login_btn.count() > 0 and await login_btn.is_visible():
                    await login_btn.click()
                    try:
                        await email_input.wait_for(state="visible", timeout=4000)
                    except Exception:
                        pass

            if await email_input.count() == 0 or not await email_input.is_visible():
                print("  [!] Không tìm thấy ô nhập email trên màn hình!")
                await browser.close()
                res = {"email": email_addr, "status": "failed", "reason": "no_email_input"}
                await save_record_async(out_dir, acc_data, res)
                return res

            await email_input.click()
            await email_input.press_sequentially(email_addr, delay=15)

            submit_btn = page.locator('button[type="submit"], button:has-text("Tiếp tục"), button:has-text("Continue")').first
            try:
                await submit_btn.wait_for(state="visible", timeout=3000)
            except Exception:
                pass

            start_ts = time.time()
            await submit_btn.click()
            print(f"  2. Đã bấm gửi email. Bắt đầu quét mã OTP từ Outlook (method={otp_method}: IMAP + Graph)...")

            # Quét OTP từ Outlook: IMAP XOAUTH2 + Microsoft Graph (chạy ngầm trong thread pool)
            loop = asyncio.get_running_loop()
            otp = await loop.run_in_executor(None, lambda: poll_meta_otp(email_addr, refresh_token, client_id, start_ts, 120, otp_method))
            if not otp:
                print("  [!] Hết thời gian chờ OTP (120s) không nhận được mail từ Meta!")
                await browser.close()
                res = {"email": email_addr, "status": "failed", "reason": "otp_timeout"}
                await save_record_async(out_dir, acc_data, res)
                return res

            print(f"  3. Nhận được mã OTP: {otp} -> Đang nhập vào form...")
            otp_input = page.locator('input[maxlength="6"], input[inputmode="numeric"], input[autocomplete="one-time-code"], input[type="text"]').first
            try:
                await otp_input.wait_for(state="attached", timeout=15000)
            except Exception:
                cur_text = await page.locator("body").inner_text()
                print(f"  [!] Không thấy ô nhập OTP. Nội dung:\n{cur_text[:250]}")
                await browser.close()
                res = {"email": email_addr, "status": "failed", "reason": "otp_input_missing"}
                await save_record_async(out_dir, acc_data, res)
                return res

            try:
                await otp_input.click(force=True)
            except Exception:
                pass
            await otp_input.focus()
            await page.keyboard.type(otp, delay=35)

            # Bấm nút Xác nhận / Submit OTP
            confirm_otp_btn = page.locator('button[type="submit"], button:has-text("Xác nhận"), button:has-text("Confirm"), button:has-text("Tiếp tục")').first
            if await confirm_otp_btn.count() > 0 and await confirm_otp_btn.is_visible():
                print("  -> Bấm nút Xác nhận OTP...")
                await confirm_otp_btn.click()
            else:
                await page.keyboard.press("Enter")

            # Xử lý thông minh các màn hình thiết lập hồ sơ Meta sau khi nhập OTP (Ngày sinh, Họ tên, Mật khẩu...)
            for step_idx in range(1, 5):
                state = await wait_for_transition(page, context, max_wait=6.0)
                if state == "authenticated" or "/access" in page.url:
                    break
                if state == "rejected":
                    print("  [!] Phát hiện Meta từ chối đăng ký (rejected dialog)!")
                    break

                acted = False
                cur_body = (await page.locator("body").inner_text()).lower()

                # 1. Màn hình Ngày sinh (Birthday)
                if state == "birthday" or any(k in cur_body for k in ["birthday", "ngày sinh", "date of birth"]):
                    print(f"  4.{step_idx}. Điền ngày sinh (15/06/1998)...")
                    selects = await page.locator("select").all()
                    if len(selects) >= 3:
                        try: await selects[0].select_option(index=6)
                        except Exception: pass
                        try: await selects[1].select_option(value="15")
                        except Exception: pass
                        try: await selects[2].select_option(value="1998")
                        except Exception: pass
                        acted = True

                # 2. Màn hình Họ tên (Name)
                elif state == "name" or any(k in cur_body for k in ["what’s your name", "what's your name", "tên của bạn", "first name", "họ tên", "tên bắt buộc"]):
                    print(f"  4.{step_idx}. Điền họ tên ({first_name} {last_name})...")
                    name_inputs = await page.locator('input:not([type="checkbox"]):not([type="hidden"]):not([type="password"])').all()
                    if len(name_inputs) >= 2:
                        await name_inputs[0].fill(first_name)
                        await name_inputs[1].fill(last_name)
                        acted = True

                # 3. Màn hình Mật khẩu (Password)
                elif state == "password" or await page.locator('input[type="password"]').count() > 0:
                    pwd_input = page.locator('input[type="password"]').first
                    if await pwd_input.is_visible():
                        print(f"  4.{step_idx}. Điền mật khẩu...")
                        await pwd_input.fill(pwd)
                        acted = True

                if acted:
                    next_btn = page.locator('button[type="submit"], button:has-text("Next"), button:has-text("Tiếp tục"), button:has-text("Xác nhận"), button:has-text("Confirm"), button:has-text("Save")').first
                    if await next_btn.count() > 0 and await next_btn.is_visible():
                        await next_btn.click()
                        await asyncio.sleep(0.5)
                else:
                    break

            # Kiểm tra trạng thái cuối cùng sau khi hoàn tất Onboarding
            await asyncio.sleep(1.0)
            current_url = page.url
            body_text = await page.locator("body").inner_text()
            print(f"  -> URL sau đăng nhập: {current_url}")

            cookies_list = await context.cookies()
            cookies_dict = {c["name"]: c["value"] for c in cookies_list}
            cookies_exp = {c["name"]: int(c.get("expires", 0)) for c in cookies_list}
            print(f"  -> Cookies tìm thấy: {list(cookies_dict.keys())}")

            has_sess = "hatch_sess" in cookies_dict
            has_vml = "hatch_vml" in cookies_dict
            is_waitlist = "/access" in current_url

            # BƯỚC 6: XÁC THỰC THỰC TẾ BẰNG LỆNH TẠO VIDEO "MUSE VẪY TAY"
            video_verified = False
            if verify_video and has_sess and not is_waitlist:
                print("  6. Kiểm tra quyền hạn thực tế: Mở phòng chat và gửi lệnh tạo video 'Muse vẫy tay'...")
                try:
                    await page.goto("https://muse.ai/thread/new", wait_until="domcontentloaded", timeout=25000)

                    if "/access" in page.url:
                        print("  [!] Sau khi đăng nhập, URL chuyển hướng về /access (Waitlist)!")
                        is_waitlist = True
                    else:
                        textarea = page.locator('textarea').first
                        try:
                            await textarea.wait_for(state="visible", timeout=12000)
                        except Exception:
                            textarea = None

                        if textarea and await textarea.is_visible():
                            print("  -> Tìm thấy khung chat! Đang nhập lệnh tạo video...")
                            await textarea.click()
                            prompt_msg = "Tạo một video: Muse vẫy tay chào thân thiện (Create a short video: waving hand friendly)"
                            # Dùng fill thay vì press_sequentially để gửi tức thì trong 0.01s
                            await textarea.fill(prompt_msg)

                            send_btn = page.locator('button[aria-label*="Send" i], button[aria-label*="Gửi" i], button[data-testid*="send" i], button:has-text("Send"), button:has-text("Gửi")').first
                            if await send_btn.count() > 0 and await send_btn.is_enabled():
                                await send_btn.click()
                            else:
                                await textarea.press("Enter")

                            print("  -> Đã gửi lệnh! Đang đợi Muse AI kích hoạt tiến trình tạo video (tối đa 30s)...")
                            v_deadline = time.time() + 30
                            while time.time() < v_deadline:
                                # 1. Nút Stop xuất hiện = Muse đang sinh video
                                stop_btn = page.locator('[data-testid="hatch-composer-stop-button"], button[aria-label*="Stop" i], button[aria-label*="停止" i]')
                                if await stop_btn.count() > 0:
                                    video_verified = True
                                    print("  [✓] Phát hiện nút Stop: Muse AI đang tích cực tạo video!")
                                    break

                                # 2. Thẻ video hoặc attachment video xuất hiện
                                vid = page.locator('video, [data-testid*="presentation-video"]')
                                if await vid.count() > 0:
                                    video_verified = True
                                    print("  [✓] Phát hiện thẻ Video trên luồng chat!")
                                    break

                                # 3. Phản hồi chat bubble từ bot về video
                                bubble = page.locator('div[class*="hatch-agent-bubble-bg"]')
                                if await bubble.count() > 0:
                                    b_txt = (await bubble.last.inner_text()).lower()
                                    if any(k in b_txt for k in ["video", "generating", "creating", "vẫy tay", "đang tạo"]):
                                        video_verified = True
                                        print(f"  [✓] Muse AI phản hồi tiến trình: {b_txt[:80]}...")
                                        break

                                # Kiểm tra thông báo lỗi hoặc giới hạn
                                err_box = page.locator('[role="alert"], [class*="error"]')
                                if await err_box.count() > 0:
                                    err_txt = await err_box.first.inner_text()
                                    print(f"  [!] Có thông báo lỗi trên trang: {err_txt[:100]}")
                                    break

                                await asyncio.sleep(0.35)
                        else:
                            print("  [!] Không tìm thấy khung nhập lệnh (textarea) trên https://muse.ai/thread/new!")
                except Exception as ve:
                    print(f"  [!] Ngoại lệ trong bước gửi lệnh tạo video: {ve}")

            # ĐÁNH GIÁ KẾT QUẢ CUỐI CÙNG
            final_url = page.url
            if (video_verified if verify_video else (has_sess and has_vml and not is_waitlist)):
                print("  🎉 [✓] ĐĂNG KÝ VÀ TẠO VIDEO THÀNH CÔNG! Tài khoản hoạt động 100%!")
                if muse2api_base and muse2api_key:
                    uploaded = upload_to_muse2api(muse2api_base, muse2api_key, email_addr, cookies_dict, cookies_exp)
                    if uploaded:
                        print("  [✓] Đã tự động nạp thành công vào Pool muse2api!")
                res = {
                    "email": email_addr,
                    "status": "success",
                    "video_verified": video_verified,
                    "url": final_url,
                    "cookies": cookies_dict,
                    "expires": cookies_exp
                }
            elif is_waitlist or "/access" in final_url or (has_sess and not has_vml):
                print("  [!] TÀI KHOẢN VÀO WAITLIST: Đã xác thực nhưng bị Muse AI chặn vùng (không có quyền tạo video)!")
                res = {
                    "email": email_addr,
                    "status": "waitlist",
                    "video_verified": False,
                    "reason": "waitlist_geo_blocked (Không kích hoạt được video do kẹt danh sách chờ)",
                    "url": final_url,
                    "cookies": cookies_dict
                }
            else:
                print("  [!] Thất bại: Không tạo được video và không đủ điều kiện phiên làm việc.")
                res = {
                    "email": email_addr,
                    "status": "failed",
                    "video_verified": False,
                    "reason": "video_verification_failed",
                    "url": final_url
                }

            await save_record_async(out_dir, acc_data, res)
            await browser.close()
            return res

        except Exception as e:
            print(f"  [!] Ngoại lệ trong quá trình đăng ký: {e}")
            await browser.close()
            res = {"email": email_addr, "status": "failed", "reason": str(e)}
            await save_record_async(out_dir, acc_data, res)
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
        description="Tool tự động đăng ký Muse AI từ Hotmail OAuth2, hỗ trợ Proxy/4G xoay, đa luồng & nạp pool."
    )
    default_file = os.environ.get("MUSE_ACCOUNTS_FILE", "accounts.txt")
    default_out_dir = os.environ.get("MUSE_OUT_DIR", os.path.expanduser("~/Downloads") if os.path.isdir(os.path.expanduser("~/Downloads")) else ".")

    parser.add_argument("--file", default=default_file, help="File txt danh sách acc (mặc định accounts.txt)")
    parser.add_argument("--proxy", default=None, help="Proxy đơn lẻ (VD: http://user:pass@host:port hoặc socks5://...)")
    parser.add_argument("--proxy-file", default=None, help="File chứa danh sách proxy, mỗi dòng 1 proxy để xoay vòng")
    parser.add_argument("-c", "--concurrency", type=int, default=1, help="Số luồng trình duyệt chạy song song (mặc định 1, khuyên dùng 2-4 với Cliproxy)")
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
    parser.add_argument("--no-verify-video", action="store_true", help="Bỏ qua bước gửi lệnh tạo video vẫy tay để xác thực quyền")
    parser.add_argument("--otp-method", choices=["auto", "imap", "graph"], default=os.environ.get("MUSE_OTP_METHOD", "auto"), help="Đường đọc OTP: auto = cả IMAP + Graph (mặc định), imap = chỉ IMAP XOAUTH2, graph = chỉ Microsoft Graph")
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

    # CHỈ DÙNG OUTLOOK: lọc bỏ Hotmail/Live/MSN khỏi danh sách
    total_read = len(accounts)
    accounts = [a for a in accounts if is_outlook_email(a["email"])]
    if total_read != len(accounts):
        print(f"[!] Đã lọc bỏ {total_read - len(accounts)} tài khoản không phải Outlook (chỉ hỗ trợ outlook.*)")

    print(f"[*] Tổng số tài khoản hợp lệ: {len(accounts)}")
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

    # Đưa danh sách vào hàng đợi đa luồng
    queue: asyncio.Queue[Tuple[int, Dict]] = asyncio.Queue()
    for idx, acc in enumerate(selected, 1):
        email_clean = acc["email"].lower()
        if not args.force and email_clean in existing_success:
            print(f"[{idx}/{len(selected)}] Tài khoản {acc['email']} đã thành công từ trước -> Bỏ qua.")
            continue
        queue.put_nowait((idx, acc))

    total_queue = queue.qsize()
    if total_queue == 0:
        print("[*] Không có tài khoản nào cần xử lý!")
        return

    concurrency = max(1, min(args.concurrency, 10))
    print(f"[*] Bắt đầu xử lý {total_queue} tài khoản với {concurrency} luồng song song...\n")

    start_all = time.time()

    async def worker(worker_id: int):
        while not queue.empty():
            try:
                idx, acc = queue.get_nowait()
            except asyncio.QueueEmpty:
                break

            current_proxy = args.proxy
            if proxy_list:
                current_proxy = proxy_list[(idx - 1) % len(proxy_list)]

            if args.rotate_url:
                trigger_ip_rotate(args.rotate_url, args.rotate_wait)

            if args.prompt_rotate and idx > 1:
                input(f"\n[?] Bật/tắt 4G trên điện thoại để lấy IP mới rồi nhấn [ENTER] để tiếp tục...")

            print(f"\n>>> [Luồng #{worker_id} | Tiến độ {idx}/{len(selected)}] Bắt đầu: {acc['email']}")
            try:
                await register_single_account(
                    acc,
                    proxy_server=current_proxy,
                    muse2api_base=args.base,
                    muse2api_key=args.key,
                    out_dir=args.out_dir,
                    headless=not args.head,
                    verify_video=not args.no_verify_video,
                    otp_method=args.otp_method,
                    session_tag=f"w{worker_id}_{secrets.token_hex(3)}"
                )
            except Exception as we:
                print(f"[Luồng #{worker_id}] Lỗi xử lý {acc['email']}: {we}")
            finally:
                queue.task_done()
            await asyncio.sleep(0.5)

    workers = [asyncio.create_task(worker(i + 1)) for i in range(concurrency)]
    await asyncio.gather(*workers)

    elapsed = time.time() - start_all
    print("\n==================================================")
    print(f"[*] HOÀN TẤT TOÀN BỘ QUÁ TRÌNH TRONG {elapsed:.1f} GIÂY!")
    print(f"    - Thành công: {os.path.join(args.out_dir, 'muse_success.txt')}")
    print(f"    - Chờ duyệt:  {os.path.join(args.out_dir, 'muse_waitlist.txt')}")
    print(f"    - Thất bại:   {os.path.join(args.out_dir, 'muse_failed.txt')}")


if __name__ == "__main__":
    asyncio.run(main())
