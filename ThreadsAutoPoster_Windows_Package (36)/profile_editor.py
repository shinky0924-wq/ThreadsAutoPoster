#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Threads Profile Editor (スレッズ プロフィール編集専用ツール)
アカウントを選択して専用ブラウザ（Chrome）を起動し、
Threadsのプロフィール画面（自己紹介文・アイコン画像・名前・リンク等）を手動でじっくり編集するための独立アプリです。

【主な特徴】
1. 完全独立アプリ: 自動投稿エンジンとは完全に切り離され、投稿タイマーや自動入力の干渉を一切受けません。
2. 5大ログアウト対策完全維持: アカウントごとの端末指紋 (fingerprint.json) とセッションCookieを読み込み、安全にChromeを起動。
3. プロフィール直行: 起動時にあなたのアカウントのプロフィールページ (https://www.threads.net/@アカウント名) を自動表示。
4. セッション自動最新保存: 編集完了後にブラウザを閉じると、Metaから発行された最新Cookieを自動抽出してプロファイルおよびExcelに永久保存。
"""

import os
import sys
import json
import re
import hashlib
import platform
import time
import urllib.parse
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime

# パス解決（EXE化対応）
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

os.chdir(BASE_DIR)

try:
    import openpyxl
except ImportError:
    openpyxl = None


def find_account_file():
    candidates = [
        "アカウント情報.xlsx",
        "アカウント情報.xls",
        "アカウント.xlsx",
        "account.xlsx",
        "accounts.xlsx"
    ]
    for c in candidates:
        p = os.path.join(BASE_DIR, c)
        if os.path.exists(p):
            return p
    return os.path.join(BASE_DIR, "アカウント情報.xlsx")


def read_excel_records(filepath):
    if not openpyxl or not filepath or not os.path.exists(filepath):
        return []
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        target_ws = None
        target_header_row_idx = 1

        for ws in wb.worksheets:
            for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
                if not row:
                    continue
                row_cells = [str(c).strip() for c in row if c is not None]
                if any(k in row_cells for k in ["アカウントID", "アカウント名", "auth_token", "sessionid", "ID"]):
                    target_ws = ws
                    target_header_row_idx = r_idx
                    break
            if target_ws:
                break

        if not target_ws:
            target_ws = wb.active
            target_header_row_idx = 1

        rows = list(target_ws.iter_rows(values_only=True))
        wb.close()
        if len(rows) < target_header_row_idx:
            return []

        header_cells = rows[target_header_row_idx - 1]
        headers = [str(c).strip() if c is not None else "" for c in header_cells]

        records = []
        for r_idx, r in enumerate(rows[target_header_row_idx:], start=target_header_row_idx + 1):
            if not r or all(c is None for c in r):
                continue
            rec = {"_row_idx": r_idx}
            has_val = False
            for h_idx, h in enumerate(headers):
                if not h:
                    continue
                val = r[h_idx] if h_idx < len(r) else None
                if val is not None:
                    has_val = True
                rec[h] = val
            if has_val:
                records.append(rec)
        return records
    except Exception as e:
        print(f"[WARN] read_excel_records error: {e}")
        return []


def clean_session_id_token(token):
    t = str(token or "").strip()
    if t.startswith("sessionid="):
        t = t[len("sessionid="):].strip()
    while t and t[0] in ('"', "'", ';'):
        t = t[1:].strip()
    while t and t[-1] in ('"', "'", ';'):
        t = t[:-1].strip()
    import urllib.parse
    return urllib.parse.unquote(t)


def _get_current_machine_id():
    raw = "|".join([
        platform.node() or os.environ.get("COMPUTERNAME", ""),
        os.environ.get("USERNAME", "") or os.environ.get("USER", ""),
        platform.machine() or "",
        sys.platform or "",
    ])
    return hashlib.sha256(raw.encode("utf-8", errors="ignore")).hexdigest()[:16]


def parse_playwright_proxy(raw_proxy):
    s = str(raw_proxy or "").strip()
    if not s or s.lower() in ["nan", "none", "-", "なし"] or s.startswith(("(", "（", "※")):
        return None
    if not re.match(r'^(?:https?|socks5?)://', s, re.IGNORECASE):
        m_colon = re.fullmatch(r'([^:@/\s]+):(\d{2,5}):([^:@\s]+):([^@\s]+)', s)
        if m_colon:
            host, port, user, pw = m_colon.groups()
            return {"server": f"http://{host}:{port}", "username": user, "password": pw}
        m_at = re.fullmatch(r'([^:@\s]+):([^@\s]+)@([^:@/\s]+):(\d{2,5})', s)
        if m_at:
            user, pw, host, port = m_at.groups()
            return {"server": f"http://{host}:{port}", "username": user, "password": pw}
        s = "http://" + s
    try:
        u = urllib.parse.urlparse(s)
        if not u.hostname:
            return None
        scheme = (u.scheme or "http").lower()
        port_part = f":{u.port}" if u.port else ""
        cfg = {"server": f"{scheme}://{u.hostname}{port_part}"}
        if u.username is not None:
            cfg["username"] = urllib.parse.unquote(u.username)
        if u.password is not None:
            cfg["password"] = urllib.parse.unquote(u.password)
        return cfg
    except Exception:
        return None


def clear_auth_cookies_preserve_device(context, domains=None, log_func=None):
    target_names = {"sessionid", "ds_user_id", "rur", "csrftoken"}
    target_domains = domains or ["threads.net", "threads.com", "instagram.com"]
    try:
        cur = context.cookies()
        kept = []
        removed = 0
        now_ts = time.time()
        for c in cur:
            dom = str(c.get("domain") or "").lower().lstrip(".")
            name = str(c.get("name") or "")
            is_meta = any(td in dom for td in target_domains)
            if is_meta and name in target_names:
                removed += 1
                continue
            c_exp = c.get("expires")
            if isinstance(c_exp, (int, float)) and c_exp != -1 and c_exp < now_ts:
                continue
            kept.append(c)
        context.clear_cookies()
        if kept:
            try:
                context.add_cookies(kept)
            except Exception:
                for kc in kept:
                    try:
                        context.add_cookies([kc])
                    except Exception:
                        pass
        if log_func and removed:
            log_func(f"[INFO] 🧹 端末保護Cookie(datr/mid/ig_did)を維持しつつ認証Cookieのみ {removed} 件リセットしました。")
    except Exception:
        pass


def detect_phone_checkpoint_screen(page):
    try:
        cur_u = (page.url or "").lower()
    except Exception:
        cur_u = ""
    if any(k in cur_u for k in [
        "/challenge/", "/checkpoint/", "auth_platform", "two_step_verification",
        "/accounts/suspended", "add_phone", "phone_number"
    ]):
        return True, f"URL検知: {cur_u[:120]}"
    try:
        for inp_sel in ['input[name="phone_number"]', 'input[type="tel"]', 'input[name="security_code"]', 'input[name="verificationCode"]']:
            try:
                if page.locator(inp_sel).first.is_visible(timeout=200):
                    return True, f"入力欄検知: {inp_sel}"
            except Exception:
                pass
        body_txt = ""
        try:
            body_txt = (page.inner_text("body") or "")[:3000]
        except Exception:
            try:
                body_txt = (page.content() or "")[:4000]
            except Exception:
                body_txt = ""
        low = body_txt.lower()
        for kw in [
            "電話番号を入力", "携帯電話番号を入力", "電話番号を確認", "本人確認が必要",
            "不審なログイン", "コードを入力", "認証コードを入力", "アカウントが一時的にロック",
            "enter your mobile number", "confirm your phone", "confirm your info",
            "check your notifications", "suspicious login", "we detected unusual"
        ]:
            if kw in low or kw in body_txt:
                return True, f"画面文言検知: {kw}"
    except Exception:
        pass
    return False, ""


def detect_and_mark_host_migration(profile_dir, account_name="", log_func=None):
    if not profile_dir or not os.path.exists(profile_dir):
        return False
    cur_host = _get_current_machine_id()
    fp_path = os.path.join(profile_dir, "fingerprint.json")
    fp = {}
    if os.path.exists(fp_path):
        try:
            with open(fp_path, "r", encoding="utf-8") as f:
                fp = json.load(f) or {}
        except Exception:
            fp = {}
    prev_host = str(fp.get("last_host_id") or "").strip()
    has_portable = os.path.exists(os.path.join(profile_dir, "portable_session.json")) or os.path.exists(os.path.join(profile_dir, "cookies.json"))
    has_default = os.path.exists(os.path.join(profile_dir, "Default"))
    host_changed = bool(prev_host and prev_host != cur_host)
    legacy_migrated = bool(not prev_host and has_default and has_portable)
    if host_changed or legacy_migrated:
        if has_portable:
            for rel in ["Default/Cookies", "Default/Cookies-journal", "Default/Network/Cookies", "Default/Network/Cookies-journal"]:
                sp = os.path.join(profile_dir, rel)
                if os.path.exists(sp):
                    try:
                        os.remove(sp)
                    except Exception:
                        pass
            try:
                with open(os.path.join(profile_dir, ".portable_needs_rebind"), "w", encoding="utf-8") as mf:
                    mf.write(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            except Exception:
                pass
        if host_changed:
            fp["needs_migration_warmup"] = True
            if log_func:
                log_func(f"[WARN] 🌐 [{account_name}] 別PC/VPS起動を検知したためDPAPI再バインドを実施します。")
    fp["last_host_id"] = cur_host
    try:
        with open(fp_path, "w", encoding="utf-8") as f:
            json.dump(fp, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return host_changed


def get_or_create_fingerprint(profile_dir):
    os.makedirs(profile_dir, exist_ok=True)
    fp_file = os.path.join(profile_dir, "fingerprint.json")
    cur_host = _get_current_machine_id()
    if os.path.exists(fp_file):
        try:
            with open(fp_file, "r", encoding="utf-8") as f:
                fp_existing = json.load(f) or {}
            if not fp_existing.get("last_host_id"):
                fp_existing["last_host_id"] = cur_host
                with open(fp_file, "w", encoding="utf-8") as fw:
                    json.dump(fp_existing, fw, ensure_ascii=False, indent=2)
            return fp_existing
        except Exception:
            pass

    fingerprint = {
        "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        "locale": "ja-JP",
        "timezone_id": "Asia/Tokyo",
        "viewport": {"width": 1280, "height": 850},
        "color_scheme": "light",
        "platform": "Win32",
        "last_host_id": cur_host,
    }
    try:
        with open(fp_file, "w", encoding="utf-8") as f:
            json.dump(fingerprint, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return fingerprint


def safe_launch_persistent_context(p, profile_dir, headless=False, log_func=print, proxy_str=""):
    os.makedirs(profile_dir, exist_ok=True)
    for lk in ["SingletonLock", "SingletonCookie", "SingletonSocket"]:
        try:
            lk_p = os.path.join(profile_dir, lk)
            if os.path.exists(lk_p):
                os.remove(lk_p)
        except Exception:
            pass

    detect_and_mark_host_migration(profile_dir, account_name=os.path.basename(profile_dir), log_func=log_func)
    fp = get_or_create_fingerprint(profile_dir)
    args = [
        "--lang=ja-JP",
        "--disable-dev-shm-usage"
    ]
    if sys.platform != "win32":
        args.append("--no-sandbox")
    proxy_cfg = parse_playwright_proxy(proxy_str)
    if proxy_cfg and log_func:
        log_func(f"[INFO] 🌐 [{os.path.basename(profile_dir)}] 専用プロキシ経由で起動: {proxy_cfg['server']}")

    for attempt in range(1, 3):
        try:
            lkwargs = dict(
                user_data_dir=profile_dir,
                headless=headless,
                channel="chrome",
                locale=fp.get("locale", "ja-JP"),
                timezone_id=fp.get("timezone_id", "Asia/Tokyo"),
                viewport=fp.get("viewport", {"width": 1280, "height": 850}),
                color_scheme=fp.get("color_scheme", "light"),
                args=args,
                ignore_default_args=["--enable-automation"]
            )
            if proxy_cfg:
                lkwargs["proxy"] = proxy_cfg
            context = p.chromium.launch_persistent_context(**lkwargs)
            try:
                context.add_init_script("Object.defineProperty(navigator, 'webdriver', { get: () => undefined });")
            except Exception:
                pass
            return context
        except Exception as e:
            if attempt == 1:
                log_func("[WARN] ロックファイルを再クリーンアップして再試行します...")
                time.sleep(1.5)
                for lk in ["SingletonLock", "SingletonCookie", "SingletonSocket"]:
                    try:
                        lk_p = os.path.join(profile_dir, lk)
                        if os.path.exists(lk_p):
                            os.remove(lk_p)
                    except Exception:
                        pass
                continue
            raise e


def inject_session_cookies(context, token, account_name, log_func=print):
    if not token or str(token).lower() in ["nan", "none"] or str(token).startswith("ここに"):
        return False
    t = clean_session_id_token(token)
    if not t:
        return False
    try:
        cur_cks = context.cookies(["https://www.threads.net", "https://www.threads.com", "https://www.instagram.com"])
        if any(c.get("name") == "sessionid" and str(c.get("value") or "").strip() == t for c in cur_cks):
            return True
    except Exception:
        pass
    uid = ""
    if ":" in t:
        potential = t.split(":")[0].strip()
        if potential.isdigit():
            uid = potential

    cookie_expires = int(time.time()) + 31536000
    domains = [
        ".threads.net", "www.threads.net", "threads.net",
        ".threads.com", "www.threads.com", "threads.com",
        ".instagram.com", "www.instagram.com", "instagram.com"
    ]

    cookies_list = []
    for d in domains:
        cookies_list.append({"name": "sessionid", "value": t, "domain": d, "path": "/", "httpOnly": True, "secure": True, "sameSite": "Lax", "expires": cookie_expires})
        if uid:
            cookies_list.append({"name": "ds_user_id", "value": str(uid), "domain": d, "path": "/", "secure": True, "sameSite": "Lax", "expires": cookie_expires})

    try:
        context.add_cookies(cookies_list)
        return True
    except Exception as e:
        log_func(f"[DEBUG] add_cookies: {e}")
        return False


def execute_id_pw_login(page, context, account_name, password, log_func=print):
    """
    【ID/PW自動ログイン】
    セッションIDに一切依存せず、ExcelのIDとパスワードを使ってChromeプロファイル上で即時自動ログインします。
    """
    clean_id = str(account_name or "").lstrip("@").strip()
    clean_pw = str(password or "").strip()
    if not clean_pw:
        return False

    log_func(f"[INFO] 🔐 [{account_name}] 【ID/PW自動ログイン】を実行中（ID: @{clean_id}）...")
    try:
        is_cp_pre, cp_why_pre = detect_phone_checkpoint_screen(page)
        if is_cp_pre:
            log_func(f"[WARN] 🚨 [{account_name}] 電話番号確認画面（{cp_why_pre}）を検知したため自動ログインを停止します。")
            return False
        # 端末保護Cookie(datr/mid/ig_did)は維持したまま、壊れた認証Cookieのみをリセット
        clear_auth_cookies_preserve_device(context, log_func=log_func)
        time.sleep(0.5)

        cur_url = (page.url or "").lower()
        if "about:blank" in cur_url or "/login" not in cur_url:
            for login_url in ["https://www.threads.net/login", "https://www.instagram.com/accounts/login/"]:
                try:
                    page.goto(login_url, timeout=30000, wait_until="domcontentloaded")
                    time.sleep(2)
                    if "about:blank" not in (page.url or "").lower():
                        break
                except Exception:
                    pass

        # 1. 連携ボタン
        for cont_sel in ['button:has-text("としてログイン")', 'div[role="button"]:has-text("としてログイン")', 'span:has-text("としてログイン")', f'button:has-text("{clean_id}")', 'button:has-text("Continue as")']:
            try:
                c_btn = page.locator(cont_sel).first
                if c_btn.is_visible(timeout=800):
                    log_func(f"[INFO] [{account_name}] ワンクリック連携ボタン ({cont_sel}) を自動クリック...")
                    c_btn.click(force=True)
                    time.sleep(3)
                    if "/login" not in (page.url or ""):
                        return True
            except Exception:
                pass

        # 2. Instagramでログインボタン
        for ig_sel in ['div[role="button"]:has-text("Instagramでログイン")', 'button:has-text("Instagramでログイン")', 'span:has-text("Instagramでログイン")', 'button:has-text("Log in with Instagram")']:
            try:
                btn = page.locator(ig_sel).first
                if btn.is_visible(timeout=1000):
                    btn.click(force=True)
                    time.sleep(2.5)
                    break
            except Exception:
                pass

        # 3. フォーム入力
        user_loc = None
        for u_sel in ['input[name="username"]', 'input[autocomplete="username"]', 'input[aria-label*="ユーザーネーム"]', 'input[aria-label*="ユーザー名"]', 'input[placeholder*="ユーザーネーム"]', 'input[placeholder*="ユーザー名"]', 'input[type="text"]']:
            try:
                loc = page.locator(u_sel).first
                if loc.is_visible(timeout=1500):
                    user_loc = loc
                    break
            except Exception:
                pass

        pass_loc = None
        for p_sel in ['input[name="password"]', 'input[autocomplete="current-password"]', 'input[type="password"]']:
            try:
                loc = page.locator(p_sel).first
                if loc.is_visible(timeout=1500):
                    pass_loc = loc
                    break
            except Exception:
                pass

        # Threads側でフォームが見つからない場合、Instagram公式ログインへフォールバック
        if not user_loc or not pass_loc:
            log_func(f"[INFO] [{account_name}] Threads側フォーム未検出のため、Instagramログイン画面へアクセスします...")
            try:
                page.goto("https://www.instagram.com/accounts/login/", timeout=30000, wait_until="domcontentloaded")
                time.sleep(2.5)
                for u_sel in ['input[name="username"]', 'input[type="text"]']:
                    try:
                        loc = page.locator(u_sel).first
                        if loc.is_visible(timeout=1500):
                            user_loc = loc
                            break
                    except Exception:
                        pass
                for p_sel in ['input[name="password"]', 'input[type="password"]']:
                    try:
                        loc = page.locator(p_sel).first
                        if loc.is_visible(timeout=1500):
                            pass_loc = loc
                            break
                    except Exception:
                        pass
            except Exception as e:
                log_func(f"[WARN] Instagramフォールバック遷移エラー: {e}")

        if not user_loc or not pass_loc:
            return False

        user_loc.click()
        time.sleep(0.3)
        user_loc.fill("")
        time.sleep(0.2)
        import random
        user_loc.type(clean_id, delay=random.randint(50, 110))
        time.sleep(0.5)

        pass_loc.click()
        time.sleep(0.3)
        pass_loc.fill("")
        time.sleep(0.2)
        pass_loc.type(clean_pw, delay=random.randint(50, 110))
        time.sleep(0.8)

        # 送信
        for s_sel in ['button[type="submit"]', 'button:has-text("ログイン")', 'div[role="button"]:has-text("ログイン")']:
            try:
                b = page.locator(s_sel).first
                if b.is_visible(timeout=1000):
                    b.click()
                    break
            except Exception:
                pass

        time.sleep(4)
        for pop_sel in ['button:has-text("ログイン情報を保存")', 'button:has-text("後で")', 'div[role="button"]:has-text("後で")', 'button:has-text("Save info")']:
            try:
                pop = page.locator(pop_sel).first
                if pop.is_visible(timeout=1000):
                    pop.click(force=True)
                    time.sleep(1.5)
            except Exception:
                pass

        for _ in range(12):
            cur = (page.url or "").lower()
            if "/login" not in cur:
                log_func(f"[SUCCESS] 🎉 [{account_name}] 【ID/PWログイン成功！】プロファイルが正規復帰しました！")
                return True
            time.sleep(1)

        return False
    except Exception as e:
        log_func(f"[ERROR] [{account_name}] execute_id_pw_login: {e}")
        return False


def auto_refresh_and_save_session(context, account_name, profile_dir, excel_path=None, log_func=print):
    try:
        for pg in getattr(context, "pages", []):
            try:
                if not pg.is_closed():
                    pg_u = (pg.url or "").lower()
                    if any(k in pg_u for k in ["/challenge/", "/checkpoint/", "auth_platform", "two_step_verification", "/accounts/suspended"]):
                        return None
            except Exception:
                pass
        cookies = context.cookies(["https://www.threads.net", "https://www.instagram.com", "https://www.threads.com"])
        if not cookies:
            return None

        session_val = None
        for c in cookies:
            if c.get("name") == "sessionid" and c.get("value"):
                v = str(c.get("value")).strip()
                if len(v) > 15:
                    session_val = v
                    break

        # 🛡️ ログアウト状態・未認証状態での上書き破壊防止ガード！
        if not session_val:
            return None

        os.makedirs(profile_dir, exist_ok=True)
        pfile = os.path.join(profile_dir, "portable_session.json")
        cfile = os.path.join(profile_dir, "cookies.json")
        context.storage_state(path=pfile)
        with open(cfile, "w", encoding="utf-8") as f:
            json.dump(cookies, f, ensure_ascii=False, indent=2)

        if session_val and excel_path and os.path.exists(excel_path) and openpyxl:
            wb = openpyxl.load_workbook(excel_path)
            for ws in wb.worksheets:
                header_row = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]
                id_col = None
                tok_col = None
                for idx, h in enumerate(header_row, start=1):
                    if h in ["アカウントID", "アカウント名", "アカウント", "ID"]:
                        id_col = idx
                    elif h in ["auth_token", "Authトークン_sessionid", "sessionid", "Authトークン"]:
                        tok_col = idx
                if id_col and tok_col:
                    for row_idx in range(2, ws.max_row + 1):
                        row_id = str(ws.cell(row=row_idx, column=id_col).value or "").strip()
                        if row_id.lstrip("@") == account_name.lstrip("@"):
                            ws.cell(row=row_idx, column=tok_col, value=session_val)
                            wb.save(excel_path)
                            break
            wb.close()

        log_func(f"[SUCCESS] 💾 [@{account_name}] 最新セッションCookie（sessionid）を自動保存しました！")
        return session_val
    except Exception as e:
        log_func(f"[WARN] auto_refresh_and_save_session error: {e}")
        return None


class ProfileEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Threads Profile Editor (スレッズ プロフィール編集ツール)")
        self.root.geometry("860x700")
        self.root.minsize(780, 600)
        self.root.configure(bg="#121212")

        self.account_excel_path = find_account_file()
        self.accounts_list = []
        self.selected_account = None

        self.active_context = None
        self.active_page = None
        self.is_browser_running = False

        self._setup_style()
        self._build_ui()
        self.load_accounts()

    def _setup_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#18181b", foreground="#e4e4e7", fieldbackground="#18181b", rowheight=28, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", background="#27272a", foreground="#ffffff", font=("Segoe UI", 9, "bold"))
        style.map("Treeview", background=[("selected", "#4f46e5")], foreground=[("selected", "#ffffff")])

    def _build_ui(self):
        # 1. Header
        header = tk.Frame(self.root, bg="#1a1a1a", pady=12, padx=16)
        header.pack(fill="x")

        title_lbl = tk.Label(header, text="👤 Threads Profile Editor", font=("Segoe UI", 14, "bold"), fg="#ffffff", bg="#1a1a1a")
        title_lbl.pack(side="left")

        desc_lbl = tk.Label(header, text="アカウント選択 ➔ 専用Chrome起動 ➔ プロフィール編集（自己紹介・画像・名前・リンク等）", font=("Segoe UI", 9), fg="#a1a1aa", bg="#1a1a1a")
        desc_lbl.pack(side="left", padx=15)

        self.status_lbl = tk.Label(header, text="● 待機中", font=("Segoe UI", 10, "bold"), fg="#10b981", bg="#1a1a1a")
        self.status_lbl.pack(side="right")

        # 2. Main Action Bar
        action_bar = tk.Frame(self.root, bg="#18181b", pady=10, padx=16, highlightbackground="#27272a", highlightthickness=1)
        action_bar.pack(fill="x", padx=16, pady=(10, 5))

        self.btn_open_browser = tk.Button(
            action_bar,
            text="🚀 選択アカウントでChrome起動（プロフィール編集）",
            font=("Segoe UI", 11, "bold"),
            bg="#6366f1",
            fg="#ffffff",
            activebackground="#4f46e5",
            activeforeground="#ffffff",
            padx=16,
            pady=8,
            relief="flat",
            cursor="hand2",
            command=self.launch_profile_editor
        )
        self.btn_open_browser.pack(side="left", padx=(0, 8))

        self.btn_close_and_save = tk.Button(
            action_bar,
            text="💾 編集完了（保存してChrome終了）",
            font=("Segoe UI", 10, "bold"),
            bg="#059669",
            fg="#ffffff",
            activebackground="#047857",
            activeforeground="#ffffff",
            padx=14,
            pady=8,
            relief="flat",
            cursor="hand2",
            state="disabled",
            command=self.close_and_save_browser
        )
        self.btn_close_and_save.pack(side="left", padx=5)

        btn_open_excel = tk.Button(
            action_bar,
            text="⚙️ アカウントExcelを開く",
            font=("Segoe UI", 9),
            bg="#27272a",
            fg="#ffffff",
            padx=10,
            pady=8,
            relief="flat",
            cursor="hand2",
            command=self.open_excel_file
        )
        btn_open_excel.pack(side="right", padx=3)

        btn_reload = tk.Button(
            action_bar,
            text="🔄 一覧更新",
            font=("Segoe UI", 9, "bold"),
            bg="#27272a",
            fg="#38bdf8",
            padx=10,
            pady=8,
            relief="flat",
            cursor="hand2",
            command=self.load_accounts
        )
        btn_reload.pack(side="right", padx=3)

        # 3. Account Listview Frame
        list_frame = tk.Frame(self.root, bg="#121212", padx=16, pady=5)
        list_frame.pack(fill="both", expand=True)

        list_header = tk.Label(list_frame, text="📋 プロフィールを編集したいアカウントを以下から選択してください:", font=("Segoe UI", 10, "bold"), fg="#e4e4e7", bg="#121212")
        list_header.pack(anchor="w", pady=(0, 4))

        columns = ("id", "prof", "session", "fingerprint", "status")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=8, selectmode="browse")
        self.tree.heading("id", text="アカウントID")
        self.tree.heading("prof", text="Chromeプロファイル")
        self.tree.heading("session", text="セッションCookie状態")
        self.tree.heading("fingerprint", text="端末指紋 (ボット対策)")
        self.tree.heading("status", text="ステータス")

        self.tree.column("id", width=180, anchor="w")
        self.tree.column("prof", width=180, anchor="w")
        self.tree.column("session", width=180, anchor="center")
        self.tree.column("fingerprint", width=160, anchor="center")
        self.tree.column("status", width=90, anchor="center")

        tree_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)

        self.tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", lambda e: self.launch_profile_editor())

        # 4. Console Log Frame
        log_frame = tk.Frame(self.root, bg="#121212", padx=16, pady=(0, 10))
        log_frame.pack(fill="x")

        lbl_log = tk.Label(log_frame, text="操作ログ / ステータス:", font=("Segoe UI", 9, "bold"), fg="#a1a1aa", bg="#121212")
        lbl_log.pack(anchor="w", pady=(2, 2))

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap="word", bg="#0a0a0a", fg="#e5e5e5", font=("Consolas", 9), height=7, relief="solid", bd=1)
        self.log_text.pack(fill="x")

        self.log_text.tag_config("INFO", foreground="#60a5fa")
        self.log_text.tag_config("SUCCESS", foreground="#34d399")
        self.log_text.tag_config("WARN", foreground="#fbbf24")
        self.log_text.tag_config("ERROR", foreground="#f87171")

    def log(self, msg):
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {msg}\n"
        tag = "INFO"
        if "[SUCCESS]" in msg: tag = "SUCCESS"
        elif "[WARN]" in msg: tag = "WARN"
        elif "[ERROR]" in msg: tag = "ERROR"

        def _append():
            self.log_text.insert(tk.END, line, tag)
            self.log_text.see(tk.END)
        self.root.after(0, _append)

    def load_accounts(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        accounts_data = read_excel_records(self.account_excel_path)
        self.accounts_list = []

        seen_ids = set()
        for acc in accounts_data:
            acc_id = str(acc.get("アカウントID") or acc.get("アカウント名") or acc.get("アカウント") or "").strip()
            if not acc_id or acc_id.lower() == "nan":
                continue
            clean_id = acc_id.lstrip("@").strip()
            seen_ids.add(clean_id.lower())
            prof_name = clean_id
            raw_token = str(acc.get("auth_token") or acc.get("Authトークン_sessionid") or acc.get("sessionid") or "").strip()
            password_val = str(acc.get("パスワード") or acc.get("password") or acc.get("PW") or acc.get("Pass") or acc.get("暗証番号") or "").strip()
            proxy_val = str(acc.get("プロキシ") or acc.get("proxy") or acc.get("Proxy") or acc.get("PROXY") or "").strip()
            status_val = str(acc.get("有効/無効") or acc.get("有効無効") or acc.get("ステータス") or "有効").strip()

            prof_dir = os.path.join(BASE_DIR, f"chrome_profiles/{prof_name}")
            pfile = os.path.join(prof_dir, "portable_session.json")
            has_pfile = os.path.exists(pfile)
            has_fp = os.path.exists(os.path.join(prof_dir, "fingerprint.json"))

            sess_label = "✅ 登録済" if (raw_token and not raw_token.startswith("ここに")) or has_pfile else "⚠️ 未設定"
            fp_label = "🛡️ 固定端末指紋済" if has_fp else "⚡ 自動付与"

            item_id = self.tree.insert("", "end", values=(
                f"@{clean_id}",
                f"chrome_profiles/{prof_name}",
                sess_label,
                fp_label,
                status_val
            ))

            acc_obj = {
                "id": clean_id,
                "prof_name": prof_name,
                "prof_dir": prof_dir,
                "token": raw_token,
                "password": password_val,
                "proxy": proxy_val,
                "status": status_val
            }
            self.accounts_list.append((item_id, acc_obj))

        # chrome_profiles フォルダ内にあるがエクセルにないフォルダもスキャン
        profiles_root = os.path.join(BASE_DIR, "chrome_profiles")
        if os.path.exists(profiles_root):
            for d in os.listdir(profiles_root):
                full_d = os.path.join(profiles_root, d)
                if os.path.isdir(full_d) and d.lower() not in seen_ids:
                    has_pfile = os.path.exists(os.path.join(full_d, "portable_session.json"))
                    has_fp = os.path.exists(os.path.join(full_d, "fingerprint.json"))
                    item_id = self.tree.insert("", "end", values=(
                        f"@{d}",
                        f"chrome_profiles/{d}",
                        "✅ プロファイル保存済" if has_pfile else "📁 フォルダのみ",
                        "🛡️ 固定端末指紋済" if has_fp else "⚡ 自動付与",
                        "手動"
                    ))
                    self.accounts_list.append((item_id, {
                        "id": d,
                        "prof_name": d,
                        "prof_dir": full_d,
                        "token": "",
                        "status": "有効"
                    }))

        if self.accounts_list:
            first_item = self.accounts_list[0][0]
            self.tree.selection_set(first_item)
            self.selected_account = self.accounts_list[0][1]
            self.log(f"[INFO] {len(self.accounts_list)} 件のアカウントを検出しました。編集したいアカウントを選択してください。")
        else:
            self.log("[WARN] アカウント情報.xlsx または chrome_profiles にアカウントが見つかりません。")

    def _on_tree_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        selected_item = sel[0]
        for it_id, acc_obj in self.accounts_list:
            if it_id == selected_item:
                self.selected_account = acc_obj
                self.log(f"[INFO] アカウント選択: @{acc_obj['id']} (プロファイル: {acc_obj['prof_name']})")
                break

    def open_excel_file(self):
        if os.path.exists(self.account_excel_path):
            try:
                os.startfile(self.account_excel_path)
            except Exception as e:
                self.log(f"[ERROR] Excelを開けませんでした: {e}")
        else:
            messagebox.showwarning("警告", f"'{self.account_excel_path}' が見つかりません。")

    def launch_profile_editor(self):
        if not self.selected_account:
            messagebox.showwarning("警告", "プロフィールを編集するアカウントを一覧から選択してください。")
            return

        if self.is_browser_running:
            messagebox.showinfo("お知らせ", "すでにブラウザが起動しています。現在のウィンドウで編集を行ってください。")
            return

        acc = self.selected_account
        account_name = acc["id"]
        prof_dir = acc["prof_dir"]
        token = acc["token"]
        password = acc.get("password", "")
        proxy_str = acc.get("proxy", "")

        self.is_browser_running = True
        self.btn_open_browser.config(state="disabled")
        self.btn_close_and_save.config(state="normal")
        self.status_lbl.config(text=f"● 編集作業中 (@{account_name})", fg="#38bdf8")

        self.log("=" * 60)
        self.log(f"[INFO] 🚀 @{account_name} のプロフィール編集ブラウザを起動しています...")
        self.log(f"[INFO] 📁 プロファイルフォルダ: {prof_dir}")

        def _worker():
            try:
                from playwright.sync_api import sync_playwright
                with sync_playwright() as p:
                    context = safe_launch_persistent_context(p, prof_dir, headless=False, log_func=self.log, proxy_str=proxy_str)
                    self.active_context = context

                    # sessionid または portable_session.json からCookieを注入
                    pfile = os.path.join(prof_dir, "portable_session.json")
                    rebind_marker = os.path.join(prof_dir, ".portable_needs_rebind")
                    if os.path.exists(pfile):
                        try:
                            with open(pfile, "r", encoding="utf-8") as pf:
                                pdata = json.load(pf)
                            cks = pdata.get("cookies", []) if isinstance(pdata, dict) else pdata
                            if cks:
                                context.add_cookies(cks)
                                self.log(f"[INFO] 🍪 portable_session.json から {len(cks)} 件のCookieを完全復元しました。")
                        except Exception as ex:
                            self.log(f"[DEBUG] pfile load: {ex}")
                    if os.path.exists(rebind_marker):
                        try:
                            os.remove(rebind_marker)
                        except Exception:
                            pass

                    if token:
                        inject_session_cookies(context, token, account_name, log_func=self.log)

                    page = context.pages[0] if context.pages else context.new_page()
                    self.active_page = page

                    # あなたのThreadsプロフィール画面に直行
                    clean_id = account_name.lstrip("@")
                    profile_url = f"https://www.threads.net/@{clean_id}"
                    self.log(f"[INFO] 🌐 プロフィール画面へアクセス中: {profile_url}")

                    try:
                        page.goto(profile_url, wait_until="domcontentloaded", timeout=45000)
                    except Exception:
                        page.goto("https://www.threads.net/", wait_until="domcontentloaded", timeout=45000)

                    time.sleep(2)

                    is_cp_ed, cp_why_ed = detect_phone_checkpoint_screen(page)
                    if is_cp_ed:
                        self.log(f"[WARN] 🚨 @{account_name} に電話番号確認/本人確認画面（{cp_why_ed}）が表示されています。自動入力は行いませんので画面の案内に従って手動で確認してください。")

                    # ログイン画面(/login)にリダイレクトされた場合の自動復帰
                    cur_u = (page.url or "").lower()
                    if "/login" in cur_u and not is_cp_ed:
                        self.log(f"[WARN] 🚪 @{account_name} のセッションが切れているためログイン画面を検知しました。")
                        if password:
                            self.log(f"[INFO] 🔐 パスワード情報を検出。ID/PWによるプロファイル自動ログインを実行します...")
                            ok = execute_id_pw_login(page, context, account_name, password, log_func=self.log)
                            if ok:
                                time.sleep(2)
                                page.goto(profile_url, wait_until="domcontentloaded", timeout=30000)
                        else:
                            self.log(f"[INFO] 💡 画面上の「Instagramでログイン」ボタンからログインしてください。ログイン後は最新セッションが自動保存されます。")

                    self.log("=" * 60)
                    self.log(f"[SUCCESS] 🎉 @{account_name} のプロフィール画面を開きました！")
                    self.log("[INFO] 💡 画面上の「プロフィールを編集」ボタンを押して、自己紹介・画像・名前・リンク等を自由に変更してください。")
                    self.log("[INFO] ⏳ 自動で閉じるタイマーや自動投稿は一切動きませんので、じっくり編集していただけます。")
                    self.log("[INFO] 🚪 編集完了後、ブラウザを「×」で閉じるか、画面の「💾 編集完了（保存してChrome終了）」を押してください。")
                    self.log("=" * 60)

                    # リアルタイム監視＆ユーザーがブラウザを閉じるまで待機
                    last_detected_session = token
                    while not page.is_closed() and self.is_browser_running:
                        try:
                            # ログイン画面(/login)に飛ばされていないかチェック＆案内
                            cur_u = page.url or ""
                            if "/login" in cur_u:
                                pass

                            # セッションのリアルタイム監視（ブラウザ内でログインやセッション更新があったら即座に永続化）
                            cks = context.cookies(["https://www.threads.net", "https://www.instagram.com"])
                            cur_sess = next((c.get("value") for c in cks if c.get("name") == "sessionid" and len(str(c.get("value") or "")) > 15), None)
                            if cur_sess and cur_sess != last_detected_session:
                                last_detected_session = cur_sess
                                auto_refresh_and_save_session(context, account_name, prof_dir, self.account_excel_path, log_func=self.log)
                                self.log(f"[SUCCESS] 🎉 [@{account_name}] ブラウザログイン・最新セッションを即座に検知！プロファイルとExcelへ永続保存しました！")
                        except Exception:
                            pass
                        time.sleep(1.0)

                    try:
                        auto_refresh_and_save_session(context, account_name, prof_dir, self.account_excel_path, log_func=self.log)
                    except Exception:
                        pass

                    try:
                        context.close()
                    except Exception:
                        pass
            except Exception as e:
                self.log(f"[ERROR] ブラウザ実行エラー: {e}")
            finally:
                self.active_context = None
                self.active_page = None
                self.is_browser_running = False
                self.root.after(0, lambda: self.btn_open_browser.config(state="normal"))
                self.root.after(0, lambda: self.btn_close_and_save.config(state="disabled"))
                self.root.after(0, lambda: self.status_lbl.config(text="● 待機中", fg="#10b981"))
                self.root.after(0, self.load_accounts)

        threading.Thread(target=_worker, daemon=True).start()

    def close_and_save_browser(self):
        if not self.is_browser_running or not self.active_context:
            return
        self.log("[INFO] 💾 保存リクエストを受信しました。最新セッションを保存してブラウザを閉じます...")
        self.is_browser_running = False


def main():
    root = tk.Tk()
    app = ProfileEditorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
