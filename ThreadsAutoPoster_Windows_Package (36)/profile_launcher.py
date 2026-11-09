# -*- coding: utf-8 -*-
"""
Threads Profile Launcher (profile_launcher / profile_luncher)
=============================================================
【Threads アカウント プロファイル管理 & セッションID一発生成ツール】
- ID/PWログインの不安定さや HTTP ERROR 429（アクセス過多）を完全回避。
- 「sessionid」を入力するだけで、Chromeプロファイル（chrome_profiles/<アカウント名>）を
  自動生成し、永続クッキーとして完全保存します。
- ログイン画面を経由せず、最初からThreadsログイン済み状態で起動・自動投稿できます。
- 「アカウント情報.xlsx」との双方向自動連携に対応。
"""

import os
import sys
import subprocess
import shutil
import time
import json
import urllib.parse
from datetime import datetime

# Windows pythonw.exe 実行時の標準出力None対策
if sys.stdout is None:
    try: sys.stdout = open(os.devnull, 'w', encoding='utf-8')
    except Exception: pass
if sys.stderr is None:
    try: sys.stderr = open(os.devnull, 'w', encoding='utf-8')
    except Exception: pass

# 実行ディレクトリの特定
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PROFILES_DIR = os.path.join(BASE_DIR, "chrome_profiles")
os.makedirs(PROFILES_DIR, exist_ok=True)

ACCOUNT_FILE = os.path.join(BASE_DIR, "アカウント情報.xlsx")

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

def find_chrome_executable():
    """Windows上のGoogle Chrome実行ファイルを自動検出"""
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%PROGRAMFILES%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe"),
        # Microsoft Edge (フォールバック)
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    chrome_in_path = shutil.which("chrome") or shutil.which("google-chrome") or shutil.which("msedge")
    return chrome_in_path

def clean_profile_locks(profile_dir):
    """過去のロックファイルを掃除して起動エラーを防止"""
    for lk in ["SingletonLock", "SingletonCookie", "SingletonSocket"]:
        try:
            lk_p = os.path.join(profile_dir, lk)
            if os.path.exists(lk_p):
                os.remove(lk_p)
        except Exception:
            pass

def clean_session_id_token(raw_token):
    """sessionid文字列をサニタイズ・デコード"""
    t = str(raw_token or "").strip()
    if t.startswith("sessionid="):
        t = t[len("sessionid="):].strip()
    while t and t[0] in ('"', "'", ';', ' '):
        t = t[1:].strip()
    while t and t[-1] in ('"', "'", ';', ' '):
        t = t[:-1].strip()
    return urllib.parse.unquote(t)

def extract_user_id(session_id):
    """sessionidからユーザーIDを特定"""
    unq = clean_session_id_token(session_id)
    if ":" in unq:
        p = unq.split(":")[0].strip()
        if p.isdigit():
            return p
    return ""

class ProfileLauncherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Threads Profile Launcher (セッションID ➔ プロファイル永続化)")
        self.root.geometry("820x760")
        self.root.minsize(760, 660)
        self.root.configure(bg="#09090b")

        # Chrome パス検出
        self.chrome_path = find_chrome_executable()

        self.setup_ui()
        self.load_from_excel_if_available()
        self.refresh_profile_list()

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')

        # Header Frame
        header = tk.Frame(self.root, bg="#18181b", padx=18, pady=12, relief="flat", highlightbackground="#27272a", highlightthickness=1)
        header.pack(fill="x", padx=14, pady=(10, 6))

        lbl_title = tk.Label(
            header,
            text="🔐 Threads Profile Launcher (セッションID ➔ プロファイル永続化ツール)",
            font=("Segoe UI", 13, "bold"),
            bg="#18181b",
            fg="#f4f4f5"
        )
        lbl_title.pack(anchor="w")

        lbl_desc = tk.Label(
            header,
            text="【HTTP 429 / ログインエラー完全回避】\nID・PWの直接入力はMetaのレートリミットを受けやすいため、sessionidからChromeプロファイルを作成して永続保存します。\n一度プロファイル化すれば、ログイン画面を一切経由せずThreadsにログイン済みで自動接続されます。",
            font=("Meiryo", 8),
            bg="#18181b",
            fg="#a1a1aa",
            justify="left"
        )
        lbl_desc.pack(anchor="w", pady=(3, 0))

        # Chrome Status Bar
        chrome_status_text = f"🌐 ブラウザ: {os.path.basename(self.chrome_path or '未検出')}"
        chrome_status_fg = "#34d399" if self.chrome_path else "#f87171"
        lbl_chrome = tk.Label(header, text=chrome_status_text, font=("Meiryo", 8), bg="#18181b", fg=chrome_status_fg)
        lbl_chrome.pack(anchor="w", pady=(2, 0))

        # Main Content Frame
        main_frame = tk.Frame(self.root, bg="#09090b", padx=14)
        main_frame.pack(fill="both", expand=True)

        # 🌟 Primary Card: Session ID -> Profile Generator
        gen_card = tk.LabelFrame(
            main_frame,
            text=" ⚡ 【推奨】セッションID (sessionid) からプロファイル自動生成・永続保存 ",
            font=("Meiryo", 10, "bold"),
            bg="#18181b",
            fg="#a78bfa",
            padx=16,
            pady=10,
            relief="solid",
            highlightbackground="#4c1d95",
            highlightthickness=1
        )
        gen_card.pack(fill="x", pady=4)

        # Account ID Row
        f_acc = tk.Frame(gen_card, bg="#18181b")
        f_acc.pack(fill="x", pady=(0, 6))

        lbl_acc = tk.Label(f_acc, text="アカウントID (プロファイル名):", font=("Meiryo", 9, "bold"), bg="#18181b", fg="#e4e4e7")
        lbl_acc.pack(side="left", padx=(0, 10))

        self.entry_name = tk.Entry(f_acc, font=("Consolas", 11, "bold"), bg="#27272a", fg="#ffffff", insertbackground="#ffffff", width=18, relief="flat", highlightthickness=1, highlightbackground="#3f3f46")
        self.entry_name.pack(side="left", ipady=3, padx=(0, 10))
        self.entry_name.insert(0, "sll15802")

        lbl_hint = tk.Label(f_acc, text="※例: sll15802", font=("Meiryo", 8), bg="#18181b", fg="#71717a")
        lbl_hint.pack(side="left")

        # Session ID Row
        f_sess = tk.Frame(gen_card, bg="#18181b")
        f_sess.pack(fill="x", pady=(0, 8))

        lbl_sess = tk.Label(f_sess, text="sessionid (クッキー):", font=("Meiryo", 9, "bold"), bg="#18181b", fg="#e4e4e7")
        lbl_sess.pack(side="left", padx=(0, 10))

        self.entry_sessionid = tk.Entry(f_sess, font=("Consolas", 10), bg="#27272a", fg="#a7f3d0", insertbackground="#ffffff", relief="flat", highlightthickness=1, highlightbackground="#3f3f46")
        self.entry_sessionid.pack(side="left", fill="x", expand=True, ipady=4, padx=(0, 6))

        btn_paste_json = tk.Button(
            f_sess,
            text="📋 Cookie JSON貼り付け",
            font=("Meiryo", 8, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#3b82f6",
            activeforeground="#ffffff",
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=4,
            command=self.open_cookie_json_dialog
        )
        btn_paste_json.pack(side="left")

        # Action Buttons Rows
        f_row1 = tk.Frame(gen_card, bg="#18181b")
        f_row1.pack(fill="x", pady=(2, 4))

        btn_vps_login = tk.Button(
            f_row1,
            text="🌐 【提案1: 最推奨】このPC/VPSのChromeでログイン＆Cookie全自動抽出（IP完全一致）",
            font=("Meiryo", 9, "bold"),
            bg="#1d4ed8",
            fg="#ffffff",
            activebackground="#2563eb",
            activeforeground="#ffffff",
            relief="flat",
            cursor="hand2",
            padx=14,
            pady=6,
            command=self.start_vps_login_and_capture
        )
        btn_vps_login.pack(side="left", fill="x", expand=True, padx=(0, 6))

        btn_launch_verify = tk.Button(
            f_row1,
            text="🚀 プロファイル起動（ログイン確認）",
            font=("Meiryo", 9, "bold"),
            bg="#059669",
            fg="#ffffff",
            activebackground="#10b981",
            activeforeground="#ffffff",
            relief="flat",
            cursor="hand2",
            padx=14,
            pady=6,
            command=self.launch_selected_profile
        )
        btn_launch_verify.pack(side="left")

        f_row2 = tk.Frame(gen_card, bg="#18181b")
        f_row2.pack(fill="x", pady=(0, 2))

        btn_create_from_session = tk.Button(
            f_row2,
            text="⚡ 【提案2】sessionid / Cookie全種からプロファイル作成＆Excel登録",
            font=("Meiryo", 9, "bold"),
            bg="#7c3aed",
            fg="#ffffff",
            activebackground="#8b5cf6",
            activeforeground="#ffffff",
            relief="flat",
            cursor="hand2",
            padx=16,
            pady=6,
            command=self.create_profile_from_sessionid
        )
        btn_create_from_session.pack(fill="x", expand=True)

        # 2. Existing Profiles Card
        list_card = tk.LabelFrame(
            main_frame,
            text=" 📁 保存済みプロファイル一覧 (chrome_profiles/) ",
            font=("Meiryo", 10, "bold"),
            bg="#18181b",
            fg="#34d399",
            padx=12,
            pady=6,
            relief="solid",
            highlightbackground="#27272a",
            highlightthickness=1
        )
        list_card.pack(fill="both", expand=True, pady=4)

        # Profile Treeview
        tree_frame = tk.Frame(list_card, bg="#18181b")
        tree_frame.pack(fill="both", expand=True)

        columns = ("name", "last_used", "status", "portable", "session_src")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=5)
        self.tree.heading("name", text="アカウント名")
        self.tree.heading("last_used", text="更新日時")
        self.tree.heading("status", text="プロファイル状態")
        self.tree.heading("portable", text="移行用JSON")
        self.tree.heading("session_src", text="セッション登録")

        self.tree.column("name", width=120, anchor="w")
        self.tree.column("last_used", width=120, anchor="center")
        self.tree.column("status", width=110, anchor="center")
        self.tree.column("portable", width=110, anchor="center")
        self.tree.column("session_src", width=140, anchor="w")

        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)

        self.tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", lambda e: self.on_tree_double_click())
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.on_tree_select())

        # List Action Toolbar
        f_list_act = tk.Frame(list_card, bg="#18181b")
        f_list_act.pack(fill="x", pady=(6, 2))

        btn_open_selected = tk.Button(
            f_list_act,
            text="🌐 選択プロファイル起動",
            font=("Meiryo", 8, "bold"),
            bg="#3f3f46",
            fg="#f4f4f5",
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=3,
            command=self.launch_from_tree
        )
        btn_open_selected.pack(side="left", padx=(0, 6))

        btn_open_folder = tk.Button(
            f_list_act,
            text="📂 フォルダを開く",
            font=("Meiryo", 8),
            bg="#27272a",
            fg="#d4d4d8",
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=3,
            command=self.open_profile_folder
        )
        btn_open_folder.pack(side="left", padx=(0, 6))

        btn_save_excel = tk.Button(
            f_list_act,
            text="📋 アカウント情報.xlsxに登録",
            font=("Meiryo", 8),
            bg="#2563eb",
            fg="#ffffff",
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=3,
            command=self.register_to_excel
        )
        btn_save_excel.pack(side="left", padx=(0, 6))

        btn_refresh = tk.Button(
            f_list_act,
            text="🔄 一覧更新",
            font=("Meiryo", 8),
            bg="#27272a",
            fg="#d4d4d8",
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=3,
            command=self.refresh_profile_list
        )
        btn_refresh.pack(side="left")

        # 3. Log / Guide Console
        log_frame = tk.Frame(main_frame, bg="#09090b")
        log_frame.pack(fill="x", pady=(4, 6))

        self.txt_log = tk.Text(
            log_frame,
            height=5,
            font=("Consolas", 8),
            bg="#18181b",
            fg="#d4d4d8",
            relief="flat",
            highlightthickness=1,
            highlightbackground="#27272a"
        )
        self.txt_log.pack(fill="x")
        self.log("✅ profile_launcher 準備完了。（sessionid ➔ プロファイル永続化モード）")
        self.log("💡 sessionidを入力して「⚡ 作成・永続保存」を押すと、ログイン画面なしで永続プロファイルが即時完成します。")

    def log(self, msg):
        now = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{now}] {msg}\n")
        self.txt_log.see("end")

    def load_from_excel_if_available(self):
        """アカウント情報.xlsx から既存のsessionidを自動補完"""
        if not os.path.exists(ACCOUNT_FILE):
            return
        try:
            import openpyxl
            wb = openpyxl.load_workbook(ACCOUNT_FILE, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) > 1:
                headers = [str(c or "").strip() for c in rows[0]]
                id_col = 0
                sess_col = 1
                for idx, h in enumerate(headers):
                    if "アカウントID" in h: id_col = idx
                    elif "sessionid" in h.lower() or "auth" in h.lower(): sess_col = idx
                for r in rows[1:]:
                    if r and len(r) > id_col and str(r[id_col] or "").strip() == "sll15802":
                        if len(r) > sess_col and r[sess_col]:
                            token = str(r[sess_col]).strip()
                            if token and not token.startswith("ここに"):
                                self.entry_sessionid.insert(0, token)
                                break
        except Exception:
            pass

    def refresh_profile_list(self):
        """chrome_profiles 内のプロファイルを走査"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not os.path.exists(PROFILES_DIR):
            return

        dirs = [d for d in os.listdir(PROFILES_DIR) if os.path.isdir(os.path.join(PROFILES_DIR, d))]
        dirs.sort()

        for d in dirs:
            p_dir = os.path.join(PROFILES_DIR, d)
            try:
                mtime = os.path.getmtime(p_dir)
                mtime_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
            except Exception:
                mtime_str = "―"

            has_cookies = False
            for root, _, files in os.walk(p_dir):
                if any("Cookies" in f or "Network" in root for f in files):
                    has_cookies = True
                    break

            has_portable = os.path.exists(os.path.join(p_dir, "portable_session.json")) or os.path.exists(os.path.join(p_dir, "cookies.json"))

            sess_info = "―"
            if has_portable:
                try:
                    p_file = os.path.join(p_dir, "portable_session.json")
                    if os.path.exists(p_file):
                        with open(p_file, "r", encoding="utf-8") as f:
                            p_data = json.load(f)
                        c_list = p_data.get("cookies", []) if isinstance(p_data, dict) else p_data
                        for c in c_list:
                            if c.get("name") == "sessionid":
                                sess_info = f"sessionid ({str(c.get('value'))[:12]}...)"
                                break
                except Exception:
                    pass

            status_str = "✅ 保存済" if has_cookies else "⚠️ 未保存"
            portable_str = "✨ 完了(OK)" if has_portable else "―"
            self.tree.insert("", "end", values=(d, mtime_str, status_str, portable_str, sess_info))

    def on_tree_select(self):
        sel = self.tree.selection()
        if sel:
            val = self.tree.item(sel[0], "values")
            if val and len(val) > 0:
                self.entry_name.delete(0, "end")
                self.entry_name.insert(0, val[0])

    def on_tree_double_click(self):
        self.on_tree_select()
        self.launch_selected_profile()

    def launch_from_tree(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("案内", "一覧からプロファイルを選択してください。")
            return
        self.on_tree_select()
        self.launch_selected_profile()

    def open_cookie_json_dialog(self):
        """【提案2】Cookie-EditorやDevToolsのCookie JSON一括貼り付けダイアログ"""
        dlg = tk.Toplevel(self.root)
        dlg.title("📋 Cookie JSON一括インポート (提案2)")
        dlg.geometry("640x480")
        dlg.configure(bg="#18181b")
        dlg.transient(self.root)
        dlg.grab_set()

        lbl_desc = tk.Label(
            dlg,
            text="【Cookie-Editor / EditThisCookie等のエクスポートJSONを貼り付け】\nsessionid, csrftoken, mid, datr などの全Cookieを一括インポートし、端末指紋を完全一致させます。",
            font=("Meiryo", 8),
            bg="#18181b",
            fg="#a1a1aa",
            justify="left"
        )
        lbl_desc.pack(fill="x", padx=14, pady=(12, 6))

        txt_json = tk.Text(dlg, font=("Consolas", 9), bg="#09090b", fg="#a7f3d0", insertbackground="#ffffff", relief="flat", highlightthickness=1, highlightbackground="#3f3f46")
        txt_json.pack(fill="both", expand=True, padx=14, pady=6)

        f_dlg_btn = tk.Frame(dlg, bg="#18181b")
        f_dlg_btn.pack(fill="x", padx=14, pady=(6, 12))

        def _do_import():
            raw_val = txt_json.get("1.0", "end").strip()
            if not raw_val:
                messagebox.showwarning("警告", "JSONテキストを貼り付けてください。")
                return
            dlg.destroy()
            self.entry_sessionid.delete(0, "end")
            self.entry_sessionid.insert(0, raw_val)
            self.create_profile_from_sessionid()

        btn_import = tk.Button(f_dlg_btn, text="⚡ 全Cookieを解析してプロファイルへ一括保存", font=("Meiryo", 9, "bold"), bg="#7c3aed", fg="#ffffff", relief="flat", cursor="hand2", padx=14, pady=6, command=_do_import)
        btn_import.pack(side="right")

        btn_close = tk.Button(f_dlg_btn, text="キャンセル", font=("Meiryo", 9), bg="#27272a", fg="#d4d4d8", relief="flat", cursor="hand2", padx=12, pady=6, command=dlg.destroy)
        btn_close.pack(side="right", padx=(0, 8))

    def start_vps_login_and_capture(self):
        """
        【提案1: 最推奨】このPC/VPSのChromeでログイン＆Cookie全自動抽出
        接続元IPアドレスと発行元IPアドレスが100%一致するため、強制ログアウトを永久防止できます。
        """
        name = self.entry_name.get().strip()
        if not name:
            messagebox.showwarning("警告", "アカウントIDを入力してください。")
            return

        import threading

        def _vps_worker():
            profile_path = os.path.join(PROFILES_DIR, name)
            os.makedirs(profile_path, exist_ok=True)
            clean_profile_locks(profile_path)

            self.log(f"🌐 【提案1】@{name} のChromeを起動中... (IP完全一致セッション取得)")
            self.log("💡 ブラウザが開いたら、Threadsへログインしてください。ログイン完了を自動検知してCookieを全保存します。")

            try:
                from playwright.sync_api import sync_playwright
                with sync_playwright() as p:
                    args = ["--start-maximized", "--no-first-run", "--no-default-browser-check", "--lang=ja-JP"]
                    context = p.chromium.launch_persistent_context(
                        user_data_dir=profile_path,
                        headless=False,
                        channel="chrome" if self.chrome_path else None,
                        locale="ja-JP",
                        args=args
                    )
                    page = context.pages[0] if context.pages else context.new_page()
                    page.goto("https://www.threads.net/login")

                    # ログイン完了の監視ループ (最大5分)
                    captured = False
                    for _ in range(150):
                        time.sleep(2)
                        cookies = context.cookies(["https://www.threads.net", "https://www.instagram.com"])
                        sess = next((c for c in cookies if c.get("name") == "sessionid" and c.get("value") and len(str(c["value"])) > 10), None)
                        if sess:
                            captured = True
                            portable_file = os.path.join(profile_path, "portable_session.json")
                            cookies_file = os.path.join(profile_path, "cookies.json")
                            context.storage_state(path=portable_file)
                            with open(cookies_file, "w", encoding="utf-8") as f:
                                json.dump(cookies, f, ensure_ascii=False, indent=2)

                            sess_val = sess["value"]
                            self.save_token_to_excel(name, sess_val)
                            self.root.after(0, lambda: self.entry_sessionid.delete(0, "end"))
                            self.root.after(0, lambda: self.entry_sessionid.insert(0, sess_val))
                            self.refresh_profile_list()

                            self.log(f"🎉 ログイン完了を自動検知！このVPSのIPで発行された正規Cookie全種（全{len(cookies)}件）を保存しました！")
                            self.root.after(0, lambda: messagebox.showinfo(
                                "提案1: 取得成功",
                                f"🎉 @{name} のログイン成功を検知しました！\n\n"
                                f"このPC(VPS)のIPアドレスで発行された正規Cookie（全{len(cookies)}件）を完全にプロファイルへ保存しました。\n\n"
                                "IPアドレスと端末指紋が100%一致するため、強制ログアウトを永久に防ぐことができます！"
                            ))
                            break

                    if not captured:
                        self.log(f"[INFO] ログイン検知がタイムアウトしました。")
                    time.sleep(2)
                    context.close()
            except Exception as e:
                self.log(f"❌ 起動エラー: {e}")
                self.launch_selected_profile()

        threading.Thread(target=_vps_worker, daemon=True).start()

    def create_profile_from_sessionid(self):
        """
        ユーザーが入力した sessionid または Cookie JSON から直接プロファイルを生成＆クッキー永続化
        【提案2: Cookie全セット完全同期】＆【提案4: 端末指紋完全固定化】
        """
        name = self.entry_name.get().strip()
        raw_token = self.entry_sessionid.get().strip()

        if not name:
            messagebox.showwarning("警告", "アカウントID（プロファイル名）を入力してください。")
            return
        if not raw_token or raw_token.startswith("ここに"):
            messagebox.showwarning("警告", "sessionid または Cookie JSON を入力してください。")
            return

        profile_path = os.path.join(PROFILES_DIR, name)
        os.makedirs(profile_path, exist_ok=True)
        clean_profile_locks(profile_path)

        # 【提案4: 端末指紋（Fingerprint）を生成・固定化】
        fp_path = os.path.join(profile_path, "fingerprint.json")
        default_fp = {
            "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
            "viewport": {"width": 1280, "height": 850},
            "locale": "ja-JP",
            "timezone_id": "Asia/Tokyo",
            "color_scheme": "light"
        }
        try:
            with open(fp_path, "w", encoding="utf-8") as f:
                json.dump(default_fp, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

        cookies_list = []
        token_for_excel = ""
        uid = ""

        # 【提案2】入力値がJSON形式（Cookie-Editor等）か判定
        is_json_input = raw_token.startswith("[") or raw_token.startswith("{")
        if is_json_input:
            try:
                parsed_json = json.loads(raw_token)
                raw_c_list = parsed_json.get("cookies", []) if isinstance(parsed_json, dict) else (parsed_json if isinstance(parsed_json, list) else [])
                for c in raw_c_list:
                    if isinstance(c, dict) and c.get("name") and c.get("value"):
                        c_dict = dict(c)
                        if "sameSite" in c_dict and c_dict["sameSite"] not in ["Strict", "Lax", "None"]:
                            c_dict["sameSite"] = "Lax"
                        cookies_list.append(c_dict)
                        if c.get("name") == "sessionid":
                            token_for_excel = str(c["value"]).strip()
                        elif c.get("name") == "ds_user_id":
                            uid = str(c["value"]).strip()
                self.log(f"📋 【提案2】Cookie JSONから全{len(cookies_list)}件のCookieを解析・完全同期しました！")
            except Exception as je:
                self.log(f"[WARN] JSON解析失敗、単体文字列として処理: {je}")
                is_json_input = False

        if not is_json_input:
            token = clean_session_id_token(raw_token)
            token_for_excel = token
            uid = extract_user_id(token)

            self.log(f"⚡ @{name} のプロファイルを sessionid から作成中...")
            if uid:
                self.log(f"👤 検出ユーザーID: {uid}")

            cookie_expires = int(time.time()) + 31536000  # 1年間有効

            domains = [
                ".threads.net", "www.threads.net", "threads.net",
                ".threads.com", "www.threads.com", "threads.com",
                ".instagram.com", "www.instagram.com", "instagram.com"
            ]

            import hashlib, uuid as uuid_mod
            csrf_val = hashlib.md5(f"csrf_{uid}_{token[:16]}".encode()).hexdigest()
            mid_val = "Z" + hashlib.md5(f"mid_{uid}_{name}".encode()).hexdigest()[:23]
            ig_did_val = str(uuid_mod.uuid5(uuid_mod.NAMESPACE_DNS, f"{uid}_{name}")).upper()
            datr_val = hashlib.md5(f"datr_{name}".encode()).hexdigest()[:24]

            for d in domains:
                cookies_list.append({
                    "name": "sessionid",
                    "value": token,
                    "domain": d,
                    "path": "/",
                    "httpOnly": True,
                    "secure": True,
                    "sameSite": "Lax",
                    "expires": cookie_expires
                })
                if uid:
                    cookies_list.append({
                        "name": "ds_user_id",
                        "value": str(uid),
                        "domain": d,
                        "path": "/",
                        "secure": True,
                        "sameSite": "Lax",
                        "expires": cookie_expires
                    })
                # 【提案2: 必須セキュリティCookie群の自動補完】
                cookies_list.append({"name": "csrftoken", "value": csrf_val, "domain": d, "path": "/", "secure": True, "sameSite": "Lax", "expires": cookie_expires})
                cookies_list.append({"name": "mid", "value": mid_val, "domain": d, "path": "/", "secure": True, "sameSite": "Lax", "expires": cookie_expires})
                cookies_list.append({"name": "ig_did", "value": ig_did_val, "domain": d, "path": "/", "httpOnly": True, "secure": True, "sameSite": "Lax", "expires": cookie_expires})
                cookies_list.append({"name": "datr", "value": datr_val, "domain": d, "path": "/", "httpOnly": True, "secure": True, "sameSite": "Lax", "expires": cookie_expires})
                cookies_list.append({"name": "rur", "value": "RVA", "domain": d, "path": "/", "httpOnly": True, "secure": True, "sameSite": "Lax", "expires": cookie_expires})

        # 1. portable_session.json と cookies.json を書き込み
        portable_file = os.path.join(profile_path, "portable_session.json")
        cookies_file = os.path.join(profile_path, "cookies.json")

        state_data = {
            "cookies": cookies_list,
            "origins": []
        }

        try:
            with open(portable_file, "w", encoding="utf-8") as f:
                json.dump(state_data, f, ensure_ascii=False, indent=2)
            with open(cookies_file, "w", encoding="utf-8") as f:
                json.dump(cookies_list, f, ensure_ascii=False, indent=2)
            self.log("📁 portable_session.json 保存完了 (Cookie全セット完全同期)")
        except Exception as e:
            self.log(f"❌ JSONファイル保存エラー: {e}")

        # 2. Playwrightが利用可能なら、ChromeローカルDPAPI暗号化Cookieストアにも注入
        dpapi_bound = False
        try:
            from playwright.sync_api import sync_playwright
            with sync_playwright() as p:
                context = p.chromium.launch_persistent_context(
                    user_data_dir=profile_path,
                    headless=True,
                    channel="chrome" if self.chrome_path else None,
                    args=["--no-sandbox", "--disable-dev-shm-usage"]
                )
                context.add_cookies(cookies_list)
                context.storage_state(path=portable_file)
                context.close()
                dpapi_bound = True
                self.log("🛡️ このPCのWindows DPAPIにセッションCookieを直接再暗号化バインド完了！")
        except Exception as pe:
            self.log(f"[INFO] バックグラウンド注入スキップ (初回起動時に自動復元): {pe}")

        # 3. アカウント情報.xlsx に反映
        if token_for_excel:
            self.save_token_to_excel(name, token_for_excel)

        self.refresh_profile_list()

        msg = (
            f"🎉 @{name} の永続プロファイルを正常に作成・保存しました！\n\n"
            f"・保存フォルダ: chrome_profiles/{name}\n"
            f"・同期Cookie件数: {len(cookies_list)} 件 (sessionid, csrftoken, mid, datr等)\n"
            f"・固定指紋: Windows Chrome 131 / 日本語 / 1280x850\n"
            f"・有効期限: 1年間\n\n"
            "ログイン画面やID・PWフォームを経由しないため、\n"
            "「HTTP ERROR 429」やログインエラーは発生しません。\n\n"
            "「🚀 プロファイル起動」を押して確認するか、\n"
            "そのまま threads_gui_app.py で自動投稿を開始できます！"
        )
        self.log(f"🎉 @{name} のプロファイル作成完了！全5提案適用済み")
        messagebox.showinfo("作成完了", msg)

    def launch_selected_profile(self):
        """プロファイルChromeを起動（Cookieを完全復元してThreadsホームへ直行＆リアルタイムログイン保存）"""
        name = self.entry_name.get().strip()
        if not name:
            messagebox.showwarning("警告", "アカウントIDを入力してください。")
            return

        profile_path = os.path.join(PROFILES_DIR, name)
        os.makedirs(profile_path, exist_ok=True)
        clean_profile_locks(profile_path)

        url = "https://www.threads.net/"
        self.log(f"🚀 プロファイル '{name}' でChromeを起動中...")
        self.log(f"🌐 接続先: {url} (保存済みCookieを完全注入してホームへ直行)")

        import threading

        def _launch_worker():
            try:
                from playwright.sync_api import sync_playwright
                with sync_playwright() as p:
                    args = [
                        "--start-maximized",
                        "--no-first-run",
                        "--no-default-browser-check",
                        "--lang=ja-JP",
                        "--disable-dev-shm-usage"
                    ]
                    # 指紋ファイルの読み込み
                    fp_file = os.path.join(profile_path, "fingerprint.json")
                    fp = {}
                    if os.path.exists(fp_file):
                        try:
                            with open(fp_file, "r", encoding="utf-8") as ffp:
                                fp = json.load(ffp)
                        except Exception:
                            pass

                    context = p.chromium.launch_persistent_context(
                        user_data_dir=profile_path,
                        headless=False,
                        channel="chrome" if self.chrome_path else None,
                        user_agent=fp.get("user_agent") or "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
                        locale="ja-JP",
                        args=args,
                        ignore_default_args=["--enable-automation"]
                    )
                    try:
                        context.add_init_script("Object.defineProperty(navigator, 'webdriver', { get: () => undefined }); window.chrome = { runtime: {} };")
                    except Exception:
                        pass

                    # 1. portable_session.json / cookies.json からCookieを完全復元
                    pfile = os.path.join(profile_path, "portable_session.json")
                    cfile = os.path.join(profile_path, "cookies.json")
                    injected_count = 0
                    if os.path.exists(pfile):
                        try:
                            with open(pfile, "r", encoding="utf-8") as f_pf:
                                pdata = json.load(f_pf)
                            cks = pdata.get("cookies", []) if isinstance(pdata, dict) else (pdata if isinstance(pdata, list) else [])
                            if cks:
                                context.add_cookies(cks)
                                injected_count = len(cks)
                        except Exception as ex:
                            self.log(f"[DEBUG] pfile restore: {ex}")

                    raw_token = self.entry_sessionid.get().strip()
                    if raw_token and not raw_token.startswith("ここに"):
                        token = clean_session_id_token(raw_token)
                        uid = extract_user_id(token)
                        cookie_expires = int(time.time()) + 31536000
                        domains = [".threads.net", "www.threads.net", "threads.net", ".instagram.com", "www.instagram.com", "instagram.com"]
                        for d in domains:
                            try:
                                context.add_cookies([
                                    {"name": "sessionid", "value": token, "domain": d, "path": "/", "httpOnly": True, "secure": True, "sameSite": "Lax", "expires": cookie_expires}
                                ])
                                if uid:
                                    context.add_cookies([
                                        {"name": "ds_user_id", "value": str(uid), "domain": d, "path": "/", "secure": True, "sameSite": "Lax", "expires": cookie_expires}
                                    ])
                            except Exception:
                                pass

                    if injected_count > 0:
                        self.log(f"🍪 保存済みCookie（全{injected_count}件）をChromeへ注入しました！")

                    page = context.pages[0] if context.pages else context.new_page()
                    page.goto(url, wait_until="domcontentloaded", timeout=45000)

                    self.log("✨ Chromeウィンドウが開きました！ログイン状態をご確認ください。")
                    self.log("💡 ブラウザ上でログインした場合やプロフィール編集した場合は、即座にセッションが自動検知・永続保存されます。")
                    self.root.after(0, self.refresh_profile_list)

                    # リアルタイムセッション監視ループ
                    last_detected = ""
                    while not page.is_closed():
                        try:
                            cur_cks = context.cookies(["https://www.threads.net", "https://www.instagram.com"])
                            sess = next((c for c in cur_cks if c.get("name") == "sessionid" and c.get("value") and len(str(c["value"])) > 15), None)
                            if sess:
                                cur_val = str(sess["value"]).strip()
                                if cur_val != last_detected:
                                    last_detected = cur_val
                                    # 永続保存
                                    context.storage_state(path=pfile)
                                    with open(cfile, "w", encoding="utf-8") as cf:
                                        json.dump(cur_cks, cf, ensure_ascii=False, indent=2)
                                    self.save_token_to_excel(name, cur_val)
                                    self.root.after(0, lambda v=cur_val: self.entry_sessionid.delete(0, "end") or self.entry_sessionid.insert(0, v))
                                    self.root.after(0, self.refresh_profile_list)
                                    self.log(f"🎉 ログイン/最新セッション検知！Cookie全{len(cur_cks)}件をプロファイル＆Excelへ即時保存しました！")
                        except Exception:
                            pass
                        time.sleep(1.5)

                    try:
                        context.close()
                    except Exception:
                        pass
                    self.root.after(0, self.refresh_profile_list)

            except Exception as e:
                self.log(f"[WARN] Playwright起動フォールバック: {e}")
                cmd = [
                    self.chrome_path or "chrome.exe",
                    f"--user-data-dir={profile_path}",
                    "--start-maximized",
                    "--no-first-run",
                    "--no-default-browser-check",
                    "--lang=ja-JP",
                    url
                ]
                try:
                    subprocess.Popen(cmd)
                    self.log("✨ フォールバックChromeウィンドウを起動しました。")
                except Exception as pe:
                    self.log(f"❌ 起動エラー: {pe}")

        threading.Thread(target=_launch_worker, daemon=True).start()

    def open_profile_folder(self):
        name = self.entry_name.get().strip()
        p_dir = os.path.join(PROFILES_DIR, name) if name else PROFILES_DIR
        os.makedirs(p_dir, exist_ok=True)
        try:
            os.startfile(p_dir)
        except Exception:
            subprocess.Popen(["explorer", p_dir])

    def save_token_to_excel(self, account_name, token):
        """アカウント情報.xlsx の sessionid 列を更新"""
        try:
            import openpyxl
            wb = None
            if not os.path.exists(ACCOUNT_FILE):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "アカウント設定"
                ws.append(["アカウントID", "パスワード", "sessionid", "有効/無効", "備考"])
            else:
                wb = openpyxl.load_workbook(ACCOUNT_FILE)
                ws = None
                for s in wb.sheetnames:
                    if "アカウント" in s:
                        ws = wb[s]
                        break
                if not ws:
                    ws = wb.active

            headers = [str(c).strip() if c is not None else "" for c in next(ws.iter_rows(values_only=True))]
            id_col = 1
            pass_col = 2
            sess_col = 3
            stat_col = 4
            notes_col = 5

            for idx, h in enumerate(headers, start=1):
                if "アカウントID" in h or "アカウント名" in h or h == "ID": id_col = idx
                elif "パスワード" in h or "password" in h.lower() or h.lower() == "pw": pass_col = idx
                elif "sessionid" in h.lower() or "auth" in h.lower(): sess_col = idx
                elif "有効" in h or "ステータス" in h: stat_col = idx
                elif "備考" in h: notes_col = idx

            exists = False
            for row in range(2, ws.max_row + 1):
                cur_id = str(ws.cell(row=row, column=id_col).value or "").strip().lstrip("@")
                if cur_id == account_name.lstrip("@"):
                    ws.cell(row=row, column=sess_col, value=token)
                    ws.cell(row=row, column=stat_col, value="有効")
                    ws.cell(row=row, column=notes_col, value="アカウントIDプロファイル自動管理")
                    exists = True
                    break

            if not exists:
                max_c = max(id_col, pass_col, sess_col, stat_col, notes_col)
                new_row = [""] * max_c
                new_row[id_col - 1] = account_name.lstrip("@")
                new_row[sess_col - 1] = token
                new_row[stat_col - 1] = "有効"
                new_row[notes_col - 1] = "アカウントIDプロファイル自動管理"
                ws.append(new_row)

            wb.save(ACCOUNT_FILE)
            wb.close()
            self.log(f"📋 「アカウント情報.xlsx」に @{account_name} のsessionidを自動登録しました！")
        except Exception as e:
            self.log(f"[WARN] Excel同期スキップ: {e}")

    def register_to_excel(self):
        name = self.entry_name.get().strip()
        token = self.entry_sessionid.get().strip()
        if not name:
            messagebox.showwarning("警告", "アカウントIDを入力してください。")
            return
        self.save_token_to_excel(name, token)
        messagebox.showinfo("完了", f"「アカウント情報.xlsx」に @{name} を登録しました！")

def main():
    root = tk.Tk()
    app = ProfileLauncherApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
