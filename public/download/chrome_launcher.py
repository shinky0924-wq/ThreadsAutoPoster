# -*- coding: utf-8 -*-
"""
Chrome カスタムプロファイル起動ツール (chrome_launcher.py)
=========================================================
初回アカウント登録用の完全手動Chrome起動ツール。
指定したプロファイル名 (例: wqwq1320) 用の独立したChromeプロファイルを
chrome_profiles/<プロファイル名> に作成・保持し、通常のGoogle Chromeを起動します。
- 自動化フラグや勝手なURL遷移は一切行いません。
- 初回起動時は「Chrome にログイン」画面が開き、任意のサイト（X / Threads等）へ
  手動でアクセスしてログインできます。
- ログイン情報はプロファイルフォルダに永続保存され、2回目以降は再ログイン不要です。
"""

import os
import sys
import re
import shutil
import subprocess

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PROFILES_ROOT = os.path.join(BASE_DIR, "chrome_profiles")


def find_chrome_executable():
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%PROGRAMFILES%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe"),
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c

    if sys.platform == "win32":
        try:
            import winreg
            for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                for sub in (
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
                    r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
                ):
                    try:
                        with winreg.OpenKey(root, sub) as k:
                            val, _ = winreg.QueryValueEx(k, None)
                            if val and os.path.exists(str(val).strip('"')):
                                return str(val).strip('"')
                    except Exception:
                        pass
        except Exception:
            pass

    return shutil.which("chrome") or shutil.which("google-chrome")


def sanitize_profile_name(raw_name):
    cleaned = str(raw_name or "").strip().lstrip("@").strip()
    cleaned = re.sub(r'[\\/:*?"<>|]+', "", cleaned)
    return cleaned.strip()


def clean_lock_and_crash_files(profile_dir):
    for lk in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
        try:
            p = os.path.join(profile_dir, lk)
            if os.path.exists(p):
                os.remove(p)
        except Exception:
            pass

    for pref_sub in (os.path.join("Default", "Preferences"), "Preferences"):
        pref_path = os.path.join(profile_dir, pref_sub)
        if os.path.exists(pref_path):
            try:
                with open(pref_path, "r", encoding="utf-8") as f:
                    text = f.read()
                if '"exit_type":"Crashed"' in text or '"exited_cleanly":false' in text:
                    text = text.replace('"exit_type":"Crashed"', '"exit_type":"Normal"')
                    text = text.replace('"exited_cleanly":false', '"exited_cleanly":true')
                    with open(pref_path, "w", encoding="utf-8") as f:
                        f.write(text)
            except Exception:
                pass


def main():
    if sys.platform == "win32":
        try:
            os.system("chcp 65001 >nul")
        except Exception:
            pass
    for stream_name in ("stdout", "stderr"):
        try:
            stream = getattr(sys, stream_name, None)
            if stream and hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    os.makedirs(PROFILES_ROOT, exist_ok=True)
    chrome_exe = find_chrome_executable()
    if not chrome_exe:
        print("=========================================")
        print("[エラー] Google Chrome (chrome.exe) が見つかりませんでした。")
        print("Google Chrome をインストールしてから再度お試しください。")
        print("=========================================")
        input("Enterキーを押すと終了します...")
        return

    cli_arg = sys.argv[1] if len(sys.argv) > 1 else None

    while True:
        print("=========================================")
        print("Chrome カスタムプロファイル起動ツール")
        if cli_arg:
            raw_input = cli_arg
            print(f"作成・使用するプロファイル名を入力してください: {raw_input}")
            cli_arg = None
        else:
            raw_input = input("作成・使用するプロファイル名を入力してください: ")

        profile_name = sanitize_profile_name(raw_input)
        if not profile_name:
            print("※ 有効なプロファイル名を入力してください（例: wqwq1320）\n")
            continue

        profile_dir = os.path.join(PROFILES_ROOT, profile_name)
        is_new = not os.path.exists(os.path.join(profile_dir, "Default"))
        os.makedirs(profile_dir, exist_ok=True)
        clean_lock_and_crash_files(profile_dir)

        print()
        if is_new:
            print("[初回作成] 新規プロファイルを作成してGoogle Chromeを起動します...")
        else:
            print("[既存利用] 保存済みプロファイルを読み込んでGoogle Chromeを起動します...")
        print(f"プロファイル保存先: {profile_dir}")
        print("※ 勝手に外部サイトへ遷移しません。ご自身でログインしたいサイトを開いて手動ログインしてください。")
        print("※ ログイン後、Chromeウィンドウを閉じるとログイン情報がプロファイルに完全保存されます。")
        print("=========================================")

        try:
            # URLや自動化フラグは一切渡さず、純粋なGoogle Chromeとして起動
            proc = subprocess.Popen([chrome_exe, f"--user-data-dir={profile_dir}"])
            print("Google Chrome を起動しました。Chromeを閉じると保存が完了します...")
            proc.wait()
            print()
            print(f"✅ プロファイル [{profile_name}] の情報を保持・保存しました！")
            print("2回目以降は同じプロファイル名、または自動投稿アプリからログイン済みのまま利用できます。")
        except Exception as e:
            print(f"[エラー] Chromeの起動に失敗しました: {e}")

        print()
        again = input("続けて別のプロファイルを起動しますか？ (y=続行 / Enter=終了): ").strip().lower()
        if again != "y":
            break
        print()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
    except Exception:
        import traceback
        print("=========================================")
        print("[エラー] 予期しないエラーが発生しました:")
        traceback.print_exc()
        print("=========================================")
        try:
            input("Enterキーを押すと終了します...")
        except Exception:
            pass
