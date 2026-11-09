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
import time
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


def get_or_create_fingerprint(profile_dir):
    os.makedirs(profile_dir, exist_ok=True)
    fp_file = os.path.join(profile_dir, "fingerprint.json")
    if os.path.exists(fp_file):
        try:
            with open(fp_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass

    fingerprint = {
        "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        "locale": "ja-JP",
        "timezone_id": "Asia/Tokyo",
        "viewport": {"width": 1280, "height": 850},
        "color_scheme": "light",
        "platform": "Win32"
    }
    try:
        with open(fp_file, "w", encoding="utf-8") as f:
            json.dump(fingerprint, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return fingerprint


def safe_launch_persistent_context(p, profile_dir, headless=False, log_func=print):
    os.makedirs(profile_dir, exist_ok=True)
    for lk in ["SingletonLock", "SingletonCookie", "SingletonSocket"]:
        try:
            lk_p = os.path.join(profile_dir, lk)
            if os.path.exists(lk_p):
                os.remove(lk_p)
        except Exception:
            pass

    fp = get_or_create_fingerprint(profile_dir)
    args = [
        "--lang=ja-JP",
        "--no-sandbox",
        "--disable-dev-shm-usage"
    ]

    for attempt in range(1, 3):
        try:
            context = p.chromium.launch_persistent_context(
                user_data_dir=profile_dir,
                headless=headless,
                channel="chrome",
                user_agent=fp.get("user_agent"),
                locale=fp.get("locale", "ja-JP"),
                timezone_id=fp.get("timezone_id", "Asia/Tokyo"),
                viewport=fp.get("viewport", {"width": 1280, "height": 850}),
                color_scheme=fp.get("color_scheme", "light"),
                args=args,
                ignore_default_args=["--enable-automation"]
            )
            try:
                context.add_init_script("""
                    Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
                    window.chrome = { runtime: {} };
                """)
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
    if not token or token.lower() in ["nan", "none"] or token.startswith("ここに"):
        return False
    t = clean_session_id_token(token)
    uid = ""
    if ":" in t:
        potential = t.split(":")[0].strip()
        if potential.isdigit():
            uid = potential

    cookie_expires = int(time.time()) + 31536000
    import hashlib, uuid
    csrf_val = hashlib.md5(f"csrf_{uid}_{t[:16]}".encode()).hexdigest()
    mid_val = "Z" + hashlib.md5(f"mid_{uid}_{account_name}".encode()).hexdigest()[:23]
    ig_did_val = str(uuid.uuid5(uuid.NAMESPACE_DNS, f"{uid}_{account_name}")).upper()
    datr_val = hashlib.md5(f"datr_{account_name}".encode()).hexdigest()[:24]

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
        cookies_list.append({"name": "csrftoken", "value": csrf_val, "domain": d, "path": "/", "secure": True, "sameSite": "Lax", "expires": cookie_expires})
        cookies_list.append({"name": "mid", "value": mid_val, "domain": d, "path": "/", "secure": True, "sameSite": "Lax", "expires": cookie_expires})
        cookies_list.append({"name": "ig_did", "value": ig_did_val, "domain": d, "path": "/", "httpOnly": True, "secure": True, "sameSite": "Lax", "expires": cookie_expires})
        cookies_list.append({"name": "datr", "value": datr_val, "domain": d, "path": "/", "httpOnly": True, "secure": True, "sameSite": "Lax", "expires": cookie_expires})
        cookies_list.append({"name": "rur", "value": "RVA", "domain": d, "path": "/", "httpOnly": True, "secure": True, "sameSite": "Lax", "expires": cookie_expires})

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
        # 失効Cookieによるリダイレクトループを未然に防止
        try:
            context.clear_cookies()
            time.sleep(0.5)
        except Exception:
            pass

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
        with open(pfile, "w", encoding="utf-8") as f:
            json.dump({"cookies": cookies, "origins": []}, f, ensure_ascii=False, indent=2)
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
                    context = safe_launch_persistent_context(p, prof_dir, headless=False, log_func=self.log)
                    self.active_context = context

                    # sessionid または portable_session.json からCookieを注入
                    pfile = os.path.join(prof_dir, "portable_session.json")
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

                    # ログイン画面(/login)にリダイレクトされた場合の自動復帰
                    cur_u = (page.url or "").lower()
                    if "/login" in cur_u:
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
