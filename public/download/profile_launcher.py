# -*- coding: utf-8 -*-
"""
Threads Profile Launcher (profile_launcher / profile_luncher)
=============================================================
【Threads アカウント プロファイル管理・ログイン永続化ツール】
- アカウントごとに独立したChromeプロファイルを作成・起動します。
- 一度IDとパスワードでログインするだけで、Cookie・LocalStorage・IndexedDBが
  プロファイルフォルダ（chrome_profiles/<アカウント名>）に永久保存されます。
- sessionidをコピー＆ペーストする必要がなく、Metaの定期強制ログアウトも回避できます。
- 作成したプロファイルは自動的に「アカウント情報.xlsx」と連携可能です。
"""

import os
import sys
import subprocess
import shutil
import time
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
    # PATHから探索
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

class ProfileLauncherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Threads Profile Launcher (プロファイル管理ツール)")
        self.root.geometry("740x660")
        self.root.minsize(680, 560)
        self.root.configure(bg="#09090b")

        # Chrome パス検出
        self.chrome_path = find_chrome_executable()

        self.setup_ui()
        self.refresh_profile_list()

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')

        # Header Frame
        header = tk.Frame(self.root, bg="#18181b", padx=18, pady=14, relief="flat", highlightbackground="#27272a", highlightthickness=1)
        header.pack(fill="x", padx=14, pady=(12, 8))

        lbl_title = tk.Label(
            header,
            text="🔐 Threads Profile Launcher (profile_launcher)",
            font=("Segoe UI", 14, "bold"),
            bg="#18181b",
            fg="#f4f4f5"
        )
        lbl_title.pack(anchor="w")

        lbl_desc = tk.Label(
            header,
            text="アカウントごとに独立したChromeプロファイルを作成・起動します。\nID・パスワードで1度ログインするだけでCookieとセッションが永続保存され、強制ログアウトを防げます。（※sessionid入力不要）",
            font=("Meiryo", 9),
            bg="#18181b",
            fg="#a1a1aa",
            justify="left"
        )
        lbl_desc.pack(anchor="w", pady=(4, 0))

        # Chrome Status Bar
        chrome_status_text = f"🌐 ブラウザ: {os.path.basename(self.chrome_path or '未検出')}"
        chrome_status_fg = "#34d399" if self.chrome_path else "#f87171"
        lbl_chrome = tk.Label(header, text=chrome_status_text, font=("Meiryo", 8), bg="#18181b", fg=chrome_status_fg)
        lbl_chrome.pack(anchor="w", pady=(4, 0))

        # Main Content Frame
        main_frame = tk.Frame(self.root, bg="#09090b", padx=14)
        main_frame.pack(fill="both", expand=True)

        # 1. New/Launch Profile Card
        launch_card = tk.LabelFrame(
            main_frame,
            text=" 🚀 プロファイル作成・Chrome起動 ",
            font=("Meiryo", 10, "bold"),
            bg="#18181b",
            fg="#60a5fa",
            padx=16,
            pady=12,
            relief="solid",
            highlightbackground="#27272a",
            highlightthickness=1
        )
        launch_card.pack(fill="x", pady=6)

        f_input = tk.Frame(launch_card, bg="#18181b")
        f_input.pack(fill="x", pady=(2, 8))

        lbl_name = tk.Label(f_input, text="プロファイル名 (アカウントID):", font=("Meiryo", 9, "bold"), bg="#18181b", fg="#e4e4e7")
        lbl_name.pack(side="left", padx=(0, 10))

        self.entry_name = tk.Entry(f_input, font=("Consolas", 11, "bold"), bg="#27272a", fg="#ffffff", insertbackground="#ffffff", width=22, relief="flat", highlightthickness=1, highlightbackground="#3f3f46")
        self.entry_name.pack(side="left", ipady=4, padx=(0, 10))
        self.entry_name.insert(0, "sll15802")

        lbl_hint = tk.Label(f_input, text="※例: sll15802, jhtwj2 など", font=("Meiryo", 8), bg="#18181b", fg="#71717a")
        lbl_hint.pack(side="left")

        # URL Options
        f_url = tk.Frame(launch_card, bg="#18181b")
        f_url.pack(fill="x", pady=(0, 10))

        lbl_url = tk.Label(f_url, text="起動時のページ:", font=("Meiryo", 9), bg="#18181b", fg="#d4d4d8")
        lbl_url.pack(side="left", padx=(0, 10))

        self.var_url_choice = tk.StringVar(value="threads_login")
        rb1 = tk.Radiobutton(f_url, text="Threads ログイン画面 (推奨)", variable=self.var_url_choice, value="threads_login", bg="#18181b", fg="#e4e4e7", selectcolor="#27272a", activebackground="#18181b", font=("Meiryo", 9))
        rb1.pack(side="left", padx=(0, 10))

        rb2 = tk.Radiobutton(f_url, text="Threads ホーム画面", variable=self.var_url_choice, value="threads_home", bg="#18181b", fg="#e4e4e7", selectcolor="#27272a", activebackground="#18181b", font=("Meiryo", 9))
        rb2.pack(side="left", padx=(0, 10))

        rb3 = tk.Radiobutton(f_url, text="Instagram ログイン", variable=self.var_url_choice, value="instagram", bg="#18181b", fg="#e4e4e7", selectcolor="#27272a", activebackground="#18181b", font=("Meiryo", 9))
        rb3.pack(side="left")

        # Big Launch Button
        f_btn = tk.Frame(launch_card, bg="#18181b")
        f_btn.pack(fill="x", pady=(4, 2))

        btn_launch = tk.Button(
            f_btn,
            text="🚀 このプロファイルでChromeを起動（ログイン用）",
            font=("Meiryo", 11, "bold"),
            bg="#059669",
            fg="#ffffff",
            activebackground="#10b981",
            activeforeground="#ffffff",
            relief="flat",
            cursor="hand2",
            padx=20,
            pady=8,
            command=self.launch_selected_profile
        )
        btn_launch.pack(side="left", fill="x", expand=True, padx=(0, 8))

        btn_save_excel = tk.Button(
            f_btn,
            text="📋 アカウント情報.xlsxに登録",
            font=("Meiryo", 9, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#3b82f6",
            activeforeground="#ffffff",
            relief="flat",
            cursor="hand2",
            padx=14,
            pady=8,
            command=self.register_to_excel
        )
        btn_save_excel.pack(side="left")

        # 2. Existing Profiles Card
        list_card = tk.LabelFrame(
            main_frame,
            text=" 📁 保存済みプロファイル一覧 (chrome_profiles/) ",
            font=("Meiryo", 10, "bold"),
            bg="#18181b",
            fg="#34d399",
            padx=12,
            pady=8,
            relief="solid",
            highlightbackground="#27272a",
            highlightthickness=1
        )
        list_card.pack(fill="both", expand=True, pady=6)

        # Profile Treeview
        tree_frame = tk.Frame(list_card, bg="#18181b")
        tree_frame.pack(fill="both", expand=True)

        columns = ("name", "last_used", "status", "path")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=5)
        self.tree.heading("name", text="プロファイル名 (アカウントID)")
        self.tree.heading("last_used", text="最終更新日時")
        self.tree.heading("status", text="保存状態")
        self.tree.heading("path", text="保存フォルダパス")

        self.tree.column("name", width=160, anchor="w")
        self.tree.column("last_used", width=140, anchor="center")
        self.tree.column("status", width=120, anchor="center")
        self.tree.column("path", width=220, anchor="w")

        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)

        self.tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", lambda e: self.on_tree_double_click())
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.on_tree_select())

        # List Action Toolbar
        f_list_act = tk.Frame(list_card, bg="#18181b")
        f_list_act.pack(fill="x", pady=(8, 2))

        btn_open_selected = tk.Button(
            f_list_act,
            text="🌐 選択プロファイルを起動",
            font=("Meiryo", 8, "bold"),
            bg="#3f3f46",
            fg="#f4f4f5",
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=4,
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
            padx=10,
            pady=4,
            command=self.open_profile_folder
        )
        btn_open_folder.pack(side="left", padx=(0, 6))

        btn_refresh = tk.Button(
            f_list_act,
            text="🔄 一覧更新",
            font=("Meiryo", 8),
            bg="#27272a",
            fg="#d4d4d8",
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=4,
            command=self.refresh_profile_list
        )
        btn_refresh.pack(side="left")

        # 3. Log / Guide Console
        log_frame = tk.Frame(main_frame, bg="#09090b")
        log_frame.pack(fill="x", pady=(4, 8))

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
        self.log("✅ profile_launcher 準備完了。プロファイル名を入力して「🚀 Chrome起動」を押してください。")
        self.log("💡 【手順】開いたChromeでThreadsにID/PWでログイン ➔ 閉じるだけでセッションが永久保存されます。")

    def log(self, msg):
        now = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{now}] {msg}\n")
        self.txt_log.see("end")

    def refresh_profile_list(self):
        """chrome_profiles 内のプロファイルを再走査"""
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

            # Cookieファイルの有無でログインデータ保存済みか推定
            has_cookies = False
            for root, _, files in os.walk(p_dir):
                if any("Cookies" in f or "Network" in root for f in files):
                    has_cookies = True
                    break

            status_str = "✅ データ保存済" if has_cookies else "⚠️ 新規/未保存"
            rel_path = f"chrome_profiles/{d}"
            self.tree.insert("", "end", values=(d, mtime_str, status_str, rel_path))

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

    def get_target_url(self):
        choice = self.var_url_choice.get()
        if choice == "threads_login":
            return "https://www.threads.net/login"
        elif choice == "instagram":
            return "https://www.instagram.com/accounts/login/"
        else:
            return "https://www.threads.net/"

    def launch_selected_profile(self):
        name = self.entry_name.get().strip()
        if not name:
            messagebox.showwarning("警告", "プロファイル名（アカウントID）を入力してください。")
            return

        if not self.chrome_path or not os.path.exists(self.chrome_path):
            # Chromeの場所を手動選択
            chrome_file = filedialog.askopenfilename(
                title="Chromeの実行ファイル (chrome.exe) を選択してください",
                filetypes=[("Chrome Executable", "chrome.exe"), ("All Files", "*.*")]
            )
            if chrome_file and os.path.exists(chrome_file):
                self.chrome_path = chrome_file
            else:
                messagebox.showerror("エラー", "Google Chromeが見つかりませんでした。Chromeをインストールしてください。")
                return

        profile_path = os.path.join(PROFILES_DIR, name)
        os.makedirs(profile_path, exist_ok=True)
        clean_profile_locks(profile_path)

        url = self.get_target_url()
        self.log(f"🚀 プロファイル '{name}' でChromeを起動中...")
        self.log(f"📂 プロファイル保存先: {profile_path}")
        self.log(f"🌐 起動URL: {url}")

        cmd = [
            self.chrome_path,
            f"--user-data-dir={profile_path}",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-blink-features=AutomationControlled",
            "--lang=ja-JP",
            url
        ]

        try:
            # 非同期でChromeを起動（ウィンドウは独立して開き、Launcherを操作可能に保つ）
            subprocess.Popen(cmd)
            self.log("✨ Chromeウィンドウが開きました！ThreadsにIDとパスワードでログインしてください。")
            self.log("💡 ログイン後、Chromeを閉じるだけでセッション情報が永続保存されます。")
            self.refresh_profile_list()
        except Exception as e:
            self.log(f"❌ 起動エラー: {e}")
            messagebox.showerror("起動エラー", f"Chromeの起動に失敗しました:\n{e}")

    def open_profile_folder(self):
        name = self.entry_name.get().strip()
        p_dir = os.path.join(PROFILES_DIR, name) if name else PROFILES_DIR
        os.makedirs(p_dir, exist_ok=True)
        try:
            os.startfile(p_dir)
        except Exception:
            subprocess.Popen(["explorer", p_dir])

    def register_to_excel(self):
        name = self.entry_name.get().strip()
        if not name:
            messagebox.showwarning("警告", "プロファイル名を入力してください。")
            return

        try:
            import openpyxl
        except ImportError:
            try:
                subprocess.run([sys.executable, "-m", "pip", "install", "--user", "openpyxl"], check=False)
                import openpyxl
            except Exception:
                messagebox.showerror("エラー", "Excel操作ライブラリ (openpyxl) が読み込めませんでした。")
                return

        try:
            if not os.path.exists(ACCOUNT_FILE):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "アカウント設定"
                ws.append(["アカウントID", "プロファイル名", "有効/無効", "備考"])
            else:
                wb = openpyxl.load_workbook(ACCOUNT_FILE)
                ws = None
                for s in wb.sheetnames:
                    if "アカウント" in s:
                        ws = wb[s]
                        break
                if not ws:
                    ws = wb.active

            # ヘッダー検査
            headers = [str(c).strip() if c is not None else "" for c in next(ws.iter_rows(values_only=True))]
            id_col = 1
            profile_col = 2
            status_col = 3
            notes_col = 4

            for idx, h in enumerate(headers, start=1):
                if "アカウントID" in h or "アカウント名" in h:
                    id_col = idx
                elif "プロファイル" in h:
                    profile_col = idx
                elif "有効" in h or "ステータス" in h:
                    status_col = idx
                elif "備考" in h:
                    notes_col = idx

            # 既存レコードチェック
            exists = False
            for row in range(2, ws.max_row + 1):
                cur_id = str(ws.cell(row=row, column=id_col).value or "").strip()
                if cur_id == name:
                    ws.cell(row=row, column=profile_col, value=name)
                    ws.cell(row=row, column=status_col, value="有効")
                    ws.cell(row=row, column=notes_col, value="profile_launcherで作成・保存")
                    exists = True
                    break

            if not exists:
                new_row = [""] * max(id_col, profile_col, status_col, notes_col)
                new_row[id_col - 1] = name
                new_row[profile_col - 1] = name
                new_row[status_col - 1] = "有効"
                new_row[notes_col - 1] = "profile_launcherで新規登録"
                ws.append(new_row)

            wb.save(ACCOUNT_FILE)
            wb.close()
            self.log(f"✅ 「アカウント情報.xlsx」にアカウント '{name}' を登録・保存しました！")
            messagebox.showinfo("完了", f"「アカウント情報.xlsx」に @{name} を登録しました！\n（プロファイル名: {name}）")
        except Exception as e:
            self.log(f"❌ Excel登録エラー: {e}")
            messagebox.showerror("エラー", f"Excelの更新に失敗しました:\n{e}")

def main():
    root = tk.Tk()
    app = ProfileLauncherApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
