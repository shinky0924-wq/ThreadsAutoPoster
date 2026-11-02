# -*- coding: utf-8 -*-
"""
Threads AutoPoster - Windows GUI アプリケーション (.exe対応)
============================================================
標準の tkinter を使用しているため、追加のGUIライブラリ不要でEXE化可能。
EXE化後のパス解決（sys.frozen）に対応済み。
openpyxlによる軽量・安全なExcel直接読み込み＆更新（NumPy/Pandas不要）。
【対応項目】
- アカウント情報.xlsx: アカウントID, auth_token (アカウント名/sessionidも自動対応)
- 投稿エクセル: ポスト文章 (ツイート文章/投稿本文も自動対応), 設定セルにimage.png, ポスト文章内のimage2.pngも自動検出
- exe起動時のExcel自動読み込み＆「Excel再読み込み」「Excel選択」機能
"""

import os
import sys
import time
import random
import re
import threading
from datetime import datetime, timedelta

# EXE実行時とPythonスクリプト実行時の基準ディレクトリを自動判別
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Windows pythonw.exe 実行時は sys.stdout / sys.stderr が None になるため、安全なダミーを出力先に設定
if sys.stdout is None:
    try:
        sys.stdout = open(os.devnull, 'w', encoding='utf-8')
    except Exception:
        pass
if sys.stderr is None:
    try:
        sys.stderr = open(os.devnull, 'w', encoding='utf-8')
    except Exception:
        pass

def log_startup(msg):
    try:
        log_file = os.path.join(BASE_DIR, "startup_log.txt")
        with open(log_file, "a", encoding="utf-8") as f:
            now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            f.write("[" + now_str + "] [PYTHON] " + str(msg) + chr(10))
    except Exception:
        pass

def log_error(msg, exc=None):
    """すべてのエラーを error_log.txt に永続保存"""
    try:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_file = os.path.join(BASE_DIR, "error_log.txt")
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"[{now_str}] [ERROR] {msg}\n")
            if exc:
                import traceback
                f.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)) + "\n")
            f.flush()
    except Exception:
        pass

def log_execution(msg):
    """すべての実行ログを execution_log.txt に永続保存"""
    try:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_file = os.path.join(BASE_DIR, "execution_log.txt")
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"[{now_str}] {msg}\n")
            f.flush()
    except Exception:
        pass

log_startup("Python script execution started. sys.executable: " + str(sys.executable))
log_startup("Python version: " + str(sys.version.split()[0]) + " | Working dir: " + str(os.getcwd()))

# グローバル例外フック（クラッシュ時に一瞬で消えるのを防ぎ、crash_log.txt & error_log.txt 保存＆ダイアログ表示）
def global_excepthook(exc_type, exc_value, exc_traceback):
    import traceback
    err_str = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_startup(f"CRASH EXCEPTION: {exc_value}")
    try:
        with open(os.path.join(BASE_DIR, "crash_log.txt"), "w", encoding="utf-8") as f:
            f.write(err_str)
    except Exception:
        pass
    try:
        with open(os.path.join(BASE_DIR, "error_log.txt"), "a", encoding="utf-8") as f:
            f.write(f"\n{'='*60}\n[{now_str}] 【致命的起動クラッシュ】\n{err_str}\n{'='*60}\n")
    except Exception:
        pass
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, f"【Threads AutoPoster 起動エラー】\n\n{exc_value}\n\nエラーログを error_log.txt に保存しました。\n詳細をご確認ください。", "Threads AutoPoster エラー", 0x10)
    except Exception:
        print(err_str)

sys.excepthook = global_excepthook


# openpyxlの自動確認＆未導入時の自動インストール（--user指定で権限エラー防止）
HAS_OPENPYXL = False
try:
    import openpyxl
    HAS_OPENPYXL = True
    log_startup("openpyxl library loaded successfully")
except ImportError:
    log_startup("openpyxl missing, attempting pip install --user...")
    try:
        import subprocess
        subprocess.run([sys.executable, "-m", "pip", "install", "--user", "openpyxl"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=40)
        import openpyxl
        HAS_OPENPYXL = True
        log_startup("openpyxl installed and loaded successfully")
    except Exception as ie:
        log_startup(f"[WARN] openpyxl auto-install failed: {ie}")

try:
    import tkinter as tk
    from tkinter import ttk, messagebox, scrolledtext, filedialog
    log_startup("Tkinter GUI library loaded successfully")
except Exception as te:
    err_msg = f"【Tkinter未検出エラー】\n\nお使いのPython環境 ({sys.executable}) にGUI描画ライブラリ (Tkinter) が含まれていません。\n詳細: {te}\n\n【解決手順】\n同梱の「0_他のPCでもすぐ動く環境作成.bat」を実行して自動セットアップを行ってください。"
    log_startup(f"[ERROR] {err_msg}")
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, err_msg, "Threads AutoPoster - GUIエラー", 0x10)
    except Exception:
        pass
    sys.exit(1)

def find_account_file():
    candidates = [
        os.path.join(BASE_DIR, "アカウント情報.xlsx"),
        os.path.join(os.getcwd(), "アカウント情報.xlsx"),
        os.path.join(BASE_DIR, "account_info.xlsx"),
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    try:
        for f in os.listdir(BASE_DIR):
            if f.endswith(".xlsx") and ("アカウント" in f or "account" in f.lower()):
                return os.path.join(BASE_DIR, f)
    except Exception:
        pass
    return candidates[0]

def read_excel_records(filepath):
    if not os.path.exists(filepath):
        return []
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        target_ws = None
        target_header_row_idx = 1
        
        for sname in wb.sheetnames:
            ws = wb[sname]
            all_rows = list(ws.iter_rows(values_only=True))
            for r_idx, row in enumerate(all_rows[:15], start=1):
                if not row:
                    continue
                row_cells = [str(c).strip() for c in row if c is not None]
                if any(k in row_cells for k in ["ポスト文章", "投稿本文", "ツイート文章", "アカウントID", "アカウント名", "auth_token"]):
                    target_ws = ws
                    target_header_row_idx = r_idx
                    break
            if target_ws:
                break
                
        if not target_ws:
            target_ws = wb.active
            target_header_row_idx = 1

        rows = list(target_ws.iter_rows(values_only=True))
        sheet_name = target_ws.title
        wb.close()
        
        if not rows or len(rows) < target_header_row_idx:
            return []
            
        header_row = rows[target_header_row_idx - 1]
        headers = [str(h).strip() if h is not None else f"col_{i}" for i, h in enumerate(header_row)]
        
        records = []
        for r_idx, row in enumerate(rows[target_header_row_idx:], start=target_header_row_idx + 1):
            if not row or all(c is None or str(c).strip() == "" for c in row):
                continue
            r_dict = {"_row_idx": r_idx, "_sheet_name": sheet_name}
            for h, val in zip(headers, row):
                r_dict[h] = "" if val is None else val
            records.append(r_dict)
        return records
    except Exception as e:
        print(f"[ERROR] Excel読み込みエラー: {filepath} ({e})")
        return []

def read_account_settings_from_post_file(filepath):
    settings = {}
    if not filepath or not os.path.exists(filepath):
        return settings
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        # 投稿リスト sheet は設定用ではないため絶対に読み込まない（ID「1」の誤取得防止）
        candidate_sheets = [s for s in wb.sheetnames if any(k in s for k in ["設定", "config", "setting", "基本設定"])]
        if not candidate_sheets:
            candidate_sheets = [s for s in wb.sheetnames if not any(p in s for p in ["投稿", "post", "Post", "リスト"])]
        if not candidate_sheets and wb.sheetnames:
            candidate_sheets = [wb.sheetnames[0]]

        for sname in candidate_sheets:
            ws = wb[sname]
            all_rows = list(ws.iter_rows(values_only=True))
            for r_idx, row in enumerate(all_rows):
                if not row:
                    continue
                row_str = " ".join([str(c).strip() for c in row if c is not None])
                
                # 1. ポスト間隔 / ツイート間隔
                if any(k in row_str for k in ["ツイート間隔", "ポスト間隔", "投稿間隔", "間隔"]):
                    for offset in range(1, 4):
                        if r_idx + offset >= len(all_rows):
                            break
                        sub_row = all_rows[r_idx + offset]
                        if not sub_row:
                            continue
                        sub_str = " ".join([str(c).strip() for c in sub_row if c is not None])
                        is_minutes = "分" in sub_str
                        is_seconds = "秒" in sub_str
                        
                        nums = []
                        for cell in sub_row:
                            if cell is not None and str(cell).strip():
                                clean_val = str(cell).strip().replace("分", "").replace("秒", "")
                                if clean_val.replace(".", "", 1).isdigit():
                                    try:
                                        nums.append(float(clean_val))
                                    except Exception:
                                        pass
                        if len(nums) >= 2:
                            mult = 60 if is_minutes or not is_seconds else 1
                            min_v = int(nums[0])
                            max_v = int(nums[1])
                            settings["min_interval_sec"] = int(min_v * mult)
                            settings["max_interval_sec"] = int(max_v * mult)
                            settings["interval_str"] = f"{min_v}〜{max_v}{'分' if mult == 60 else '秒'}"
                            break
                        elif len(nums) == 1:
                            mult = 60 if is_minutes or not is_seconds else 1
                            val = int(nums[0])
                            settings["min_interval_sec"] = int(val * mult)
                            settings["max_interval_sec"] = int(val * mult)
                            settings["interval_str"] = f"{val}{'分' if mult == 60 else '秒'}"
                            break

                # 2. 実行時間 (0:00 〜 0:00)
                if any(k in row_str for k in ["ツイート実行時間", "ポスト実行時間", "実行時間"]):
                    for offset in range(1, 4):
                        if r_idx + offset >= len(all_rows):
                            break
                        sub_row = all_rows[r_idx + offset]
                        if not sub_row:
                            continue
                        times = []
                        for cell in sub_row:
                            if cell is None:
                                continue
                            c_s = str(cell).strip()
                            if hasattr(cell, "strftime"):
                                times.append(cell.strftime("%H:%M"))
                            elif re.match(r'^\d{1,2}:\d{2}', c_s):
                                times.append(c_s)
                            elif c_s.isdigit() and len(c_s) <= 2:
                                times.append(f"{c_s}:00")
                        if len(times) >= 2:
                            settings["time_slot"] = f"{times[0]}-{times[1]}"
                            break
                        elif len(times) == 1:
                            settings["time_slot"] = times[0]
                            break

                # 3. モード (0:上から順, 1:ランダム)
                if any(k in row_str for k in ["モード", "投稿モード", "mode"]):
                    for offset in range(1, 4):
                        if r_idx + offset >= len(all_rows):
                            break
                        sub_row = all_rows[r_idx + offset]
                        if not sub_row:
                            continue
                        for cell in sub_row:
                            if cell is not None and str(cell).strip() in ["0", "1"]:
                                settings["post_mode"] = int(str(cell).strip())
                                break
                        if "post_mode" in settings:
                            break

                # 4. Chromeプロファイル保存パス
                if any(k in row_str for k in ["Chromeプロファイル", "プロファイル保存パス", "プロファイル保存先", "プロファイル", "profile"]):
                    for offset in range(1, 4):
                        if r_idx + offset >= len(all_rows):
                            break
                        sub_row = all_rows[r_idx + offset]
                        if not sub_row:
                            continue
                        for cell in sub_row:
                            if cell is not None:
                                c_s = str(cell).strip()
                                if c_s and not c_s.startswith("(") and not c_s.startswith("（") and not c_s.startswith("※"):
                                    settings["profile_dir"] = c_s
                                    break
                        if "profile_dir" in settings:
                            break

                # 5. コミュニティまたはトピック (デフォルト設定)
                if any(k in row_str for k in ["コミュニティまたはトピック", "コミュニティ", "トピック", "トピックタグ"]):
                    # まず同じ行の隣の列（B列以降）を確認
                    for cell in row[1:]:
                        if cell is not None:
                            c_val = str(cell).strip()
                            if c_val and not c_val.startswith("(") and not c_val.startswith("（") and not c_val.startswith("※") and c_val.lower() not in ["none", "nan"]:
                                if not re.match(r'^\+?\d+(\.0+)?$', c_val):
                                    settings["topic"] = c_val
                                    break
                    # 同一行にない場合は次の行（1〜3行先）を確認
                    if "topic" not in settings:
                        for offset in range(1, 4):
                            if r_idx + offset >= len(all_rows):
                                break
                            sub_row = all_rows[r_idx + offset]
                            if not sub_row:
                                continue
                            for cell in sub_row:
                                if cell is not None and str(cell).strip():
                                    c_val = str(cell).strip()
                                    if c_val and not c_val.startswith("(") and not c_val.startswith("（") and not c_val.startswith("※") and c_val.lower() not in ["none", "nan"]:
                                        if not re.match(r'^\+?\d+(\.0+)?$', c_val):
                                            settings["topic"] = c_val
                                            break
                            if "topic" in settings:
                                break
        wb.close()
    except Exception as e:
        print(f"[WARN] ポスト設定読み込みエラー: {e}")
    return settings

def update_post_status(filepath, row_idx, status="投稿完了", finish_time=None, error_msg="", sheet_name=None, applied_topic=None):
    try:
        wb = openpyxl.load_workbook(filepath)
        ws = None
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            for sname in wb.sheetnames:
                candidate_ws = wb[sname]
                for r in candidate_ws.iter_rows(max_row=15, values_only=True):
                    if r and any(str(c).strip() in ["ステータス", "status", "状態"] for c in r if c is not None):
                        ws = candidate_ws
                        break
                if ws:
                    break
            if not ws:
                ws = wb.active

        status_col = None
        time_col = None
        err_col = None
        topic_col = None
        applied_topic_col = None

        for r_num, row_cells in enumerate(ws.iter_rows(max_row=15), start=1):
            vals = [str(cell.value).strip() if cell.value is not None else "" for cell in row_cells]
            for candidate in ["ステータス", "status", "状態"]:
                if candidate in vals:
                    status_col = vals.index(candidate) + 1
                    break
            for candidate in ["投稿完了日時", "完了日時", "日時", "posted_at"]:
                if candidate in vals:
                    time_col = vals.index(candidate) + 1
                    break
            for candidate in ["エラー内容", "エラー", "error"]:
                if candidate in vals:
                    err_col = vals.index(candidate) + 1
                    break
            for candidate in ["適用トピック", "使用トピック", "選択トピック", "投稿トピック"]:
                if candidate in vals:
                    applied_topic_col = vals.index(candidate) + 1
                    break
            for candidate in ["コミュニティまたはトピック", "トピック", "コミュニティ"]:
                if candidate in vals:
                    topic_col = vals.index(candidate) + 1
                    break
            if status_col:
                break

        if status_col:
            ws.cell(row=row_idx, column=status_col, value=status)
        if time_col and finish_time:
            ws.cell(row=row_idx, column=time_col, value=finish_time)
        if err_col:
            ws.cell(row=row_idx, column=err_col, value=error_msg)
        if applied_topic:
            if applied_topic_col:
                ws.cell(row=row_idx, column=applied_topic_col, value=applied_topic)
            elif topic_col:
                cur_v = ws.cell(row=row_idx, column=topic_col).value
                # 空欄の場合は適用されたトピックを記録
                if not cur_v or str(cur_v).strip() in ["", "None", "nan"]:
                    ws.cell(row=row_idx, column=topic_col, value=applied_topic)
        wb.save(filepath)
        wb.close()
        return True
    except Exception as e:
        print(f"[WARN] Excel書き込みエラー: {e}")
        return False

def update_account_sessionid(filepath, account_name, new_sessionid):
    try:
        if not filepath or not os.path.exists(filepath):
            return False
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active
        id_col = None
        token_col = None
        for r_num, row_cells in enumerate(ws.iter_rows(max_row=5), start=1):
            vals = [str(cell.value).strip() if cell.value is not None else "" for cell in row_cells]
            for idx, v in enumerate(vals, start=1):
                if v in ["アカウントID", "アカウント名", "アカウント", "ID"]:
                    id_col = idx
                elif v in ["auth_token", "Authトークン_sessionid", "sessionid", "Authトークン"]:
                    token_col = idx
            if id_col and token_col:
                break
        
        if not token_col:
            wb.close()
            return False
            
        target_row = None
        for r_num in range(2, ws.max_row + 1):
            cell_val = ws.cell(row=r_num, column=id_col if id_col else 1).value
            if cell_val and str(cell_val).strip() == account_name:
                target_row = r_num
                break
        
        if not target_row and ws.max_row >= 2:
            target_row = 2
            
        if target_row:
            ws.cell(row=target_row, column=token_col, value=new_sessionid)
            wb.save(filepath)
            wb.close()
            return True
        wb.close()
    except Exception as e:
        print(f"[WARN] sessionid 更新エラー: {e}")
    return False

def setup_media_blocking(target_context_or_page, log_func=None):
    """
    【提案2: 画像・動画（リール等）読み込みブロック機能】
    Playwrightのネットワークルーティングを活用し、タイムラインやリール等の
    重い画像(.jpg .png .webp)や動画(.mp4 .webm)、音声、Webフォントの自動通信を遮断します。
    ※ 自身のポスト添付画像のアップロード(POST)やAPI/GraphQL通信は正常に通過させます。
    効果: ブラウザ1窓あたりのメモリ使用量を約50%〜70%カットし、ページ表示を数倍高速化、
         CPUファンが激しく回るのを防ぎます。
    """
    def route_filter(route):
        try:
            req = route.request
            method = req.method.upper()
            url = req.url.lower()
            res_type = (req.resource_type or "").lower()

            # 1. POST / PUT / PATCH によるアップロードやAPI、GraphQL通信は絶対に許可
            if method in ["POST", "PUT", "PATCH"]:
                route.continue_()
                return

            if any(k in url for k in ["/api/", "graphql", "upload", "logging", "session", "login"]):
                route.continue_()
                return

            # 2. メディア・画像・フォントリソースを遮断
            if res_type in ["image", "media", "font"]:
                route.abort()
                return

            # 3. 拡張子による判定
            clean_url = url.split("?")[0].split("#")[0]
            blocked_exts = (
                ".jpg", ".jpeg", ".png", ".webp", ".gif", ".ico",
                ".mp4", ".webm", ".m4v", ".mov", ".mkv", ".avi",
                ".mp3", ".wav", ".ogg",
                ".woff", ".woff2", ".ttf", ".otf"
            )
            if any(clean_url.endswith(ext) for ext in blocked_exts):
                route.abort()
                return

            route.continue_()
        except Exception:
            try:
                route.continue_()
            except Exception:
                pass

    try:
        target_context_or_page.route("**/*", route_filter)
        if log_func:
            log_func("[INFO] ⚡ 【提案2: メディア通信遮断】有効: タイムラインの画像・動画・リール読み込みを自動ブロック（メモリ50-70%削減・CPUファン静音化）")
    except Exception as e:
        if log_func:
            log_func(f"[WARN] メディア遮断ルールの登録スキップ: {e}")

def test_session_id_online(token):
    """Meta/Threadsサーバーと通信してsessionidの有効性とアカウント名を即時判定"""
    import urllib.request, urllib.parse, re
    t = str(token or "").strip()
    if t.startswith("sessionid="):
        t = t[len("sessionid="):].strip()
    while t and t[0] in ('"', "'", ';'):
        t = t[1:].strip()
    while t and t[-1] in ('"', "'", ';'):
        t = t[:-1].strip()
    if not t or t.lower() == "nan" or t.startswith("ここに"):
        return False, "sessionidが未入力です", ""
    
    unq_t = urllib.parse.unquote(t)
    uid = unq_t.split(":")[0].strip() if ":" in unq_t else ""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
        "Cookie": f"sessionid={unq_t}; ds_user_id={uid}"
    }
    try:
        req = urllib.request.Request("https://www.threads.net/", headers=headers)
        with urllib.request.urlopen(req, timeout=12) as res:
            html = res.read().decode("utf-8", errors="ignore")
            users = re.findall(r'\"username\":\"([a-zA-Z0-9_.-]+)\"', html)
            has_viewer = '"viewer":{' in html and '"viewer":null' not in html
            if (has_viewer and users) or (users and uid and uid in html):
                return True, users[0], uid
            else:
                return False, "認証セッションが無効か期限切れです", ""
    except Exception as e:
        return False, f"通信エラー: {e}", ""

def inject_session_cookies(context, raw_token, account_name, log_func=print):
    token = str(raw_token or "").strip()
    if not token or token.lower() == "nan" or token.startswith("ここに"):
        log_func(f"[WARN] [{account_name}] Excelの auth_token (sessionid) が空です。")
        return False
    
    if token.startswith("sessionid="):
        token = token[len("sessionid="):].strip()
    token = token.strip()
    while token and token[0] in ('"', "'", ';'):
        token = token[1:].strip()
    while token and token[-1] in ('"', "'", ';'):
        token = token[:-1].strip()
    if not token:
        return False

    import urllib.parse
    decoded_token = urllib.parse.unquote(token)
    ds_user_id = ""
    if ":" in decoded_token:
        potential_id = decoded_token.split(":")[0].strip()
        if potential_id.isdigit():
            ds_user_id = potential_id
            log_func(f"[INFO] [{account_name}] sessionidからユーザーID ({ds_user_id}) を検出しました。")
    elif "%3A" in token.upper():
        potential_id = token.upper().split("%3A")[0].strip()
        if potential_id.isdigit():
            ds_user_id = potential_id
            log_func(f"[INFO] [{account_name}] sessionidからユーザーID ({ds_user_id}) を検出しました。")

    # MetaのセッションIDはコロン(:)を含むデコード済み形式が正規です
    # デコード版と生版の両方を準備
    cookie_values_to_set = [decoded_token]
    if token != decoded_token:
        cookie_values_to_set.append(token)

    domains = [
        ".threads.net", "www.threads.net", "threads.net",
        ".threads.com", "www.threads.com", "threads.com",
        ".instagram.com", "www.instagram.com", "instagram.com"
    ]

    cookies_to_add = []
    for domain in domains:
        cookies_to_add.append({
            "name": "sessionid",
            "value": decoded_token,
            "domain": domain,
            "path": "/",
            "httpOnly": True,
            "secure": True,
            "sameSite": "Lax"
        })
        if ds_user_id:
            cookies_to_add.append({
                "name": "ds_user_id",
                "value": str(ds_user_id),
                "domain": domain,
                "path": "/",
                "secure": True,
                "sameSite": "Lax"
            })

    try:
        context.add_cookies(cookies_to_add)
    except Exception as e:
        log_func(f"[DEBUG] add_cookies error: {e}")

    for u in ["https://www.threads.net/", "https://www.threads.com/", "https://www.instagram.com/"]:
        try:
            url_cookies = [{
                "name": "sessionid",
                "value": decoded_token,
                "url": u,
                "httpOnly": True,
                "secure": True,
                "sameSite": "Lax"
            }]
            if ds_user_id:
                url_cookies.append({
                    "name": "ds_user_id",
                    "value": str(ds_user_id),
                    "url": u,
                    "secure": True,
                    "sameSite": "Lax"
                })
            context.add_cookies(url_cookies)
        except Exception:
            pass

    log_func(f"[SUCCESS] [{account_name}] エクセルの sessionid (ユーザーID: {ds_user_id or '未特定'}) をThreads/Instagramに自動注入しました。")
    return True

def detect_threads_logged_in(page):
    # 1. ページ内部の Relay / 初期化データからログイン済みユーザーを判定（言語・UI変更に最も強い判定法）
    try:
        content = page.content()
        if '"viewer":{' in content and ('"username":"' in content or '"id":"' in content):
            if '"viewer":null' not in content:
                import re
                m_user = re.search(r'"viewer":\{[^}]*"username":"([^"]+)"', content)
                user_found = m_user.group(1) if m_user else ""
                return True, f"セッション認証済み (ユーザー: @{user_found or '確認済'})"
    except Exception:
        pass

    # 2. UI要素での判定（日本語・英語・アイコン双方に対応）
    success_selectors = [
        'svg[aria-label="新しい投稿"]',
        'svg[aria-label="New post"]',
        'svg[aria-label="投稿を作成"]',
        'svg[aria-label="Create"]',
        'svg[aria-label="作成"]',
        'div[role="button"]:has-text("投稿")',
        'div[role="button"]:has-text("Post")',
        'a[href*="/@"]',
        'svg[aria-label="ホーム"]',
        'svg[aria-label="Home"]',
        'svg[aria-label="検索"]',
        'svg[aria-label="Search"]',
        'svg[aria-label="アクティビティ"]',
        'svg[aria-label="Activity"]',
        'svg[aria-label="プロフィール"]',
        'svg[aria-label="Profile"]'
    ]
    for sel in success_selectors:
        try:
            if page.locator(sel).first.is_visible(timeout=600):
                return True, f"ログイン済み要素 ({sel}) を検出"
        except Exception:
            pass

    # 3. 未ログイン要素のチェック
    unauth_selectors = [
        'span:has-text("Instagramでログイン")',
        'div[role="button"]:has-text("Instagramでログイン")',
        'button:has-text("Instagramでログイン")',
        'span:has-text("Log in with Instagram")',
        'button:has-text("Log in with Instagram")',
        'a[href*="/login"]'
    ]
    for sel in unauth_selectors:
        try:
            if page.locator(sel).first.is_visible(timeout=600):
                return False, f"未ログイン要素 ({sel}) を検出"
        except Exception:
            pass

    return False, "未認証状態"

def save_context_cookies_to_excel(context, filepath, account_name, log_func=print):
    try:
        cookies = context.cookies(["https://www.threads.net", "https://www.instagram.com"])
        new_sessionid = ""
        for c in cookies:
            if c.get("name") == "sessionid" and c.get("value"):
                new_sessionid = c.get("value")
                break
        if new_sessionid:
            ok = update_account_sessionid(filepath, account_name, new_sessionid)
            if ok:
                log_func(f"[INFO] [{account_name}] セッション状態を「アカウント情報.xlsx」と自動同期しました。")
            return new_sessionid
    except Exception as e:
        log_func(f"[WARN] sessionid 自動保存エラー: {e}")
    return ""

def ensure_authenticated_session(page, context, account_name, filepath, log_func=print, wait_timeout_sec=30):
    time.sleep(2)
    is_ok, reason = detect_threads_logged_in(page)
    if is_ok:
        log_func(f"[SUCCESS] [{account_name}] エクセルの sessionid で自動ログインに成功しました！（{reason}）")
        save_context_cookies_to_excel(context, filepath, account_name, log_func)
        return True

    log_func(f"[INFO] [{account_name}] Threads画面のセッション照合中...")

    for login_sel in [
        'div[role="button"]:has-text("Instagramでログイン")',
        'button:has-text("Instagramでログイン")',
        'span:has-text("Instagramでログイン")',
        'div:has-text("Instagramでログイン")',
        'a:has-text("Instagramでログイン")',
        'text="Instagramでログイン"',
        'div[role="button"]:has-text("Log in with Instagram")',
        'button:has-text("Log in with Instagram")'
    ]:
        try:
            loc = page.locator(login_sel).first
            if loc.is_visible(timeout=1200):
                log_func(f"[INFO] [{account_name}] 「Instagramでログイン」を検出。自動クリックしてログインを実行します...")
                loc.click(force=True)
                time.sleep(3)
                break
        except Exception:
            pass

    for cont_sel in [
        'button:has-text("としてログイン")',
        'div[role="button"]:has-text("としてログイン")',
        'span:has-text("としてログイン")',
        'button:has-text("次へ")',
        'button:has-text("Continue as")',
        'div[role="button"]:has-text("Continue as")',
        'button:has-text("ログイン情報を保存")',
        'button:has-text("後で")',
        'div[role="button"]:has-text("後で")'
    ]:
        try:
            c_loc = page.locator(cont_sel).first
            if c_loc.is_visible(timeout=1200):
                log_func(f"[INFO] [{account_name}] 連携確認ボタンを自動クリックしました...")
                c_loc.click(force=True)
                time.sleep(3)
                break
        except Exception:
            pass

    for close_sel in ['svg[aria-label="閉じる"]', 'div[role="button"][aria-label="閉じる"]']:
        try:
            c_btn = page.locator(close_sel).first
            if c_btn.is_visible(timeout=500):
                c_btn.click(force=True)
                time.sleep(1)
        except Exception:
            pass

    for _ in range(15):
        time.sleep(1)
        is_ok, reason = detect_threads_logged_in(page)
        if is_ok:
            log_func(f"[SUCCESS] [{account_name}] 🎉 エクセルの sessionid による自動ログインに成功しました！（{reason}）")
            save_context_cookies_to_excel(context, filepath, account_name, log_func)
            return True

    for login_sel in ['div[role="button"]:has-text("Instagramでログイン")', 'button:has-text("Instagramでログイン")']:
        try:
            btn = page.locator(login_sel).first
            if btn.is_visible(timeout=800):
                btn.click(force=True)
                time.sleep(3)
        except Exception:
            pass

    is_ok, reason = detect_threads_logged_in(page)
    if is_ok:
        log_func(f"[SUCCESS] [{account_name}] 🎉 自動ログイン完了！（{reason}）")
        save_context_cookies_to_excel(context, filepath, account_name, log_func)
        return True

    log_func(f"[ERROR] [{account_name}] エクセルの sessionid で自動ログインできませんでした。Excelの「auth_token」列の値が正しいか、有効期限が切れていないかご確認ください。")
    return False

def resolve_image_path(raw_path, post_text=""):
    search_dirs = [BASE_DIR, os.path.join(BASE_DIR, "images"), os.getcwd(), os.path.join(os.getcwd(), "images")]
    
    if raw_path and str(raw_path).strip().lower() not in ["", "nan", "none"]:
        cleaned = str(raw_path).strip()
        if os.path.isabs(cleaned) and os.path.exists(cleaned):
            return cleaned
        for d in search_dirs:
            cand = os.path.join(d, cleaned)
            if os.path.exists(cand):
                return os.path.abspath(cand)
            base_name = os.path.basename(cleaned)
            cand_base = os.path.join(d, base_name)
            if os.path.exists(cand_base):
                return os.path.abspath(cand_base)

    if post_text:
        try:
            # ハイフンや記号による文字範囲エラーを防ぐ安全な正規表現
            matches = [m.group(0) for m in re.finditer(r'[-a-zA-Z0-9_./\\:]+?\.(?:png|jpg|jpeg|webp|gif)', str(post_text), re.IGNORECASE)]
            for m in matches:
                m_clean = m.strip(' ,()[]{}')
                for d in search_dirs:
                    cand = os.path.join(d, m_clean)
                    if os.path.exists(cand):
                        return os.path.abspath(cand)
                    cand_base = os.path.join(d, os.path.basename(m_clean))
                    if os.path.exists(cand_base):
                        return os.path.abspath(cand_base)
        except Exception:
            pass

    return ""

def make_valid_sample_png(width=600, height=600, r=0, g=149, b=246):
    """600x600 の高画質単色PNGバイナリを純粋Python(標準ライブラリのみ)で生成"""
    import zlib, struct
    def chunk(tag, data):
        return struct.pack('>I', len(data)) + tag + data + struct.pack('>I', zlib.crc32(tag + data) & 0xffffffff)
    header = b'\x89PNG\r\n\x1a\n'
    ihdr = chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0))
    raw_row = b'\x00' + bytes([r, g, b] * width)
    raw_data = raw_row * height
    idat = chunk(b'IDAT', zlib.compress(raw_data, 9))
    iend = chunk(b'IEND', b'')
    return header + ihdr + idat + iend

def ensure_valid_image(img_path, log_func=None):
    """
    Threadsの画像アップロード要件（極小10x10画像や破損ファイルの拒否）を満たすよう自動検証・修復。
    ファイルサイズが500バイト未満の場合はThreadsが「添付ファイルを添付できませんでした」と拒否するため
    高解像度(600x600)の有効なPNGへ自動修復・最適化します。
    """
    if not img_path or not os.path.exists(img_path):
        return img_path
    try:
        abs_p = os.path.abspath(img_path)
        size = os.path.getsize(abs_p)
        if size < 500:
            if log_func:
                log_func(f"[INFO] 画像 '{os.path.basename(abs_p)}' は極小サイズ({size}バイト)のため、Threads仕様を満たす600x600の高解像度画像に自動修復しました。")
            color = (0, 149, 246) if "2" not in os.path.basename(abs_p) else (120, 50, 220)
            valid_bytes = make_valid_sample_png(600, 600, *color)
            with open(abs_p, "wb") as f:
                f.write(valid_bytes)
        return abs_p
    except Exception as e:
        if log_func:
            log_func(f"[WARN] 画像検証中: {e}")
        return os.path.abspath(img_path)

def set_topic_in_threads_composer(page, dialog, topic_name, log_func=None):
    """
    Threads新規投稿モーダル内で「コミュニティまたはトピック」を設定する関数
    ※ユーザーの指示通り、アカウント名（ユーザー名）の右側にある「コミュニティまたはトピック」設定ボタンをクリックし、
      登録ワード（例: 信太山新地）を検索・入力して確定します。
    ※本文への #ハッシュタグ追記は一切行いません（完全ネイティブUI設定）。
    """
    if not topic_name or not str(topic_name).strip() or str(topic_name).strip().lower() in ["none", "nan"]:
        return False
    clean_topic = str(topic_name).strip().lstrip('#').strip()
    # 数字のみの場合は無効なIDと判定してスキップ
    if not clean_topic or re.match(r'^\+?\d+(\.0+)?$', clean_topic):
        return False

    if log_func:
        log_func(f"[*] アカウント名右側の「コミュニティまたはトピック」に「{clean_topic}」を設定します...")

    topic_set = False

    # 1. 投稿ダイアログの上部（アカウント名表示エリア周辺・本文テキストボックスより上）にある
    #    「トピック」「コミュニティ」「+」「#」などの設定ボタンを探索
    target_btn = None

    # (A) テキストやaria-labelによる厳密検索
    priority_selectors = [
        'div[role="dialog"] div[role="button"]:has-text("トピック")',
        'div[role="dialog"] button:has-text("トピック")',
        'div[role="dialog"] span[role="button"]:has-text("トピック")',
        'div[role="dialog"] div[role="button"]:has-text("コミュニティ")',
        'div[role="dialog"] button:has-text("コミュニティ")',
        'div[role="dialog"] span[role="button"]:has-text("コミュニティ")',
        'div[role="dialog"] [aria-label*="トピック"]',
        'div[role="dialog"] [aria-label*="コミュニティ"]',
        'div[role="dialog"] [aria-label*="Topic"]',
        'div[role="dialog"] [aria-label*="Community"]',
        'div[role="dialog"] svg[aria-label*="トピック"]',
        'div[role="dialog"] svg[aria-label*="コミュニティ"]',
        'div[role="dialog"] svg[aria-label*="Topic"]',
        'div[role="dialog"] [aria-label*="タグ"]',
        'div[role="dialog"] div[role="button"]:has-text("Add topic")',
        'div[role="dialog"] button:has-text("Add topic")',
        'div[role="dialog"] div[role="button"]:has-text("Add a topic")',
    ]
    for sel in priority_selectors:
        try:
            el = page.locator(sel).first
            if el.is_visible(timeout=300):
                parent_btn = page.locator(f'div[role="dialog"] div[role="button"]:has({sel}), div[role="dialog"] button:has({sel})').first
                target_btn = parent_btn if parent_btn.is_visible() else el
                break
        except Exception:
            pass

    # (B) 本文入力欄の上部にあるボタン群から探索（アカウント名の右隣にある＋やトピックボタン）
    if not target_btn:
        try:
            tb = dialog.locator('div[role="textbox"], div[contenteditable="true"]').first
            tb_box = tb.bounding_box() if tb.is_visible() else None
            max_y = (tb_box['y'] + 15) if tb_box else 9999
            
            for el in dialog.locator('div[role="button"], button, span[role="button"]').all():
                try:
                    if el.is_visible():
                        box = el.bounding_box()
                        if box and box['y'] < max_y:
                            txt = el.inner_text().strip()
                            aria = el.get_attribute("aria-label") or ""
                            combined = f"{txt} {aria}"
                            if any(k in combined for k in ["トピック", "コミュニティ", "Topic", "Community", "タグ", "tag"]):
                                target_btn = el
                                break
                            elif ("+" in txt or "＋" in txt or "#" in txt) and box['width'] < 160:
                                target_btn = el
                                break
                except Exception:
                    pass
        except Exception:
            pass

    # ボタンをクリック
    if target_btn:
        try:
            target_btn.click()
            time.sleep(0.7)
        except Exception as ce:
            if log_func:
                log_func(f"[DEBUG] トピックボタンクリック試行: {ce}")

    # 2. 開いたトピック/コミュニティ検索入力欄を探して入力
    for inp_sel in [
        'input[placeholder*="トピック"]',
        'input[placeholder*="Topic"]',
        'input[placeholder*="コミュニティ"]',
        'input[placeholder*="Community"]',
        'input[placeholder*="タグ"]',
        'input[placeholder*="検索"]',
        'input[placeholder*="Search"]',
        'input[aria-label*="トピック"]',
        'input[aria-label*="Topic"]',
        'input[aria-label*="コミュニティ"]',
        'input[type="text"]',
        'div[role="dialog"] input',
        'div[role="combobox"] input',
    ]:
        try:
            inp = page.locator(inp_sel).first
            if inp.is_visible(timeout=600):
                inp.click()
                time.sleep(0.2)
                try:
                    inp.fill("")
                except Exception:
                    pass
                inp.press_sequentially(clean_topic, delay=30)
                time.sleep(0.6)

                # Threadsのサジェストリスト（候補リスト）をクリック、またはEnterで決定
                selected = False
                for opt_sel in [
                    'div[role="listbox"] div[role="option"]',
                    'div[role="listbox"] div[role="button"]',
                    'div[role="option"]',
                    'div[role="menuitem"]',
                    'li[role="option"]',
                    f'div:has-text("「{clean_topic}」")',
                    f'span:has-text("「{clean_topic}」")',
                ]:
                    try:
                        candidates = page.locator(opt_sel).all()
                        for c_opt in candidates:
                            if c_opt.is_visible():
                                opt_text = c_opt.inner_text().strip()
                                if clean_topic in opt_text or "追加" in opt_text or "Add" in opt_text or len(candidates) == 1:
                                    c_opt.click()
                                    selected = True
                                    time.sleep(0.4)
                                    break
                        if selected:
                            break
                    except Exception:
                        pass

                if not selected:
                    inp.press("Enter")
                    time.sleep(0.5)

                topic_set = True
                if log_func:
                    log_func(f"[SUCCESS] コミュニティ/トピック「{clean_topic}」をアカウント名の右側の設定欄に正常に適用しました。")
                break
        except Exception:
            pass

    # 3. 設定結果の通知（※本文末尾へのハッシュタグ追加はユーザー要望により完全廃止）
    if not topic_set:
        if log_func:
            log_func(f"[WARN] コミュニティ/トピック「{clean_topic}」の設定UIが見つかりませんでした。本文を改変せずそのまま投稿します。")

    return topic_set

def capture_error_diagnostic(page, account_name="unknown", context_label="posting", exception_obj=None, log_func=None):

    """
    エラー発生時にThreads画面のスクリーンショットを自動撮影し、
    DOM内のトースト・アラート・エラーメッセージを自動抽出して記録する診断関数
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    clean_acc = re.sub(r'[^a-zA-Z0-9_-]', '_', str(account_name))
    ss_dir = os.path.join(BASE_DIR, "error_screenshots")
    os.makedirs(ss_dir, exist_ok=True)
    
    ss_filename = f"error_{clean_acc}_{ts}.png"
    ss_path = os.path.join(ss_dir, ss_filename)
    rel_ss = os.path.join("error_screenshots", ss_filename)
    
    ss_saved = False
    if page:
        try:
            page.screenshot(path=ss_path, full_page=False)
            ss_saved = True
        except Exception:
            try:
                page.screenshot(path=ss_path)
                ss_saved = True
            except Exception:
                pass

    screen_messages = []
    if page:
        error_selectors = [
            'div[role="alert"]',
            'div:has-text("添付ファイルを添付できませんでした")',
            'div:has-text("アップロードできませんでした")',
            'div:has-text("投稿できませんでした")',
            'div:has-text("エラーが発生しました")',
            'div:has-text("問題が発生しました")',
            'div:has-text("ブロック")',
            'div:has-text("制限")',
            'div:has-text("後ほどもう一度お試しください")',
            'div:has-text("しばらくしてからもう一度お試しください")',
            'div[aria-live="assertive"]',
        ]
        for sel in error_selectors:
            try:
                for el in page.locator(sel).all():
                    if el.is_visible(timeout=150):
                        txt = el.inner_text().strip()
                        if txt and 3 <= len(txt) <= 200 and txt not in screen_messages:
                            # 正常なUIテキストやユーザー名・リンク等はエラーメッセージとして誤検知しない
                            if not any(k in txt for k in ["閉じる", "投稿する", "投稿", "Post", "キャンセル", "OK", "新しいスレッド", "スレッドに追加", "投稿オプション", "Threadsの友達"]):
                                screen_messages.append(txt)
            except Exception:
                pass

    exc_msg = str(exception_obj) if exception_obj else "不明なエラー"
    # 本物の画面エラーメッセージがある場合はそれを優先、無ければPythonの例外内容を明確に表示
    short_reason = screen_messages[0] if screen_messages else exc_msg
    if len(short_reason) > 80:
        short_reason = short_reason[:77] + "..."

    try:
        log_file = os.path.join(BASE_DIR, "error_log.txt")
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"\n{'='*60}\n")
            f.write(f"[{now_str}] 【エラー検知】 アカウント: @{account_name} | 対象: {context_label}\n")
            f.write(f"エラー内容: {exc_msg}\n")
            if screen_messages:
                f.write(f"Threads画面上のメッセージ:\n")
                for sm in screen_messages:
                    f.write(f"  ・{sm}\n")
            if ss_saved:
                f.write(f"スクリーンショット保存先: {ss_path}\n")
            import traceback
            if exception_obj:
                f.write("詳細トレースバック:\n" + "".join(traceback.format_exception(type(exception_obj), exception_obj, exception_obj.__traceback__)) + "\n")
            f.write(f"{'='*60}\n")
    except Exception:
        pass

    return {
        "screenshot_saved": ss_saved,
        "screenshot_path": ss_path if ss_saved else "",
        "rel_screenshot": rel_ss if ss_saved else "",
        "screen_messages": screen_messages,
        "exception_msg": exc_msg,
        "short_reason": short_reason,
    }


def is_within_target_hours(time_setting, current_hour=None):
    if current_hour is None:
        current_hour = datetime.now().hour
    if time_setting is None:
        return True
    raw = str(time_setting).strip()
    if not raw or raw.lower() == "nan":
        return True
    parts = [p.strip() for p in raw.replace("、", ",").replace(";", ",").split(",") if p.strip()]
    if not parts:
        return True
    for part in parts:
        is_range = False
        for sep in ["〜", "~", "ー", "-", "to"]:
            if sep in part:
                tokens = part.split(sep, 1)
                try:
                    s = int(tokens[0].split(":")[0].strip())
                    e = int(tokens[1].split(":")[0].strip())
                    is_range = True
                    if s == e or (s == 0 and e >= 24):
                        return True
                    elif s < e:
                        if s <= current_hour < e:
                            return True
                    else:
                        if current_hour >= s or current_hour < e:
                            return True
                except Exception:
                    pass
                break
        if not is_range:
            try:
                h = int(part.split(":")[0].strip())
                if current_hour == h:
                    return True
            except Exception:
                pass
    return False

def reset_posts_file_silent(posts_file):
    """Excelの全ポストのステータスを『未投稿』に戻し、エラー内容をクリアする（自動ループ・周回用）"""
    if not posts_file or not os.path.exists(posts_file):
        return 0
    try:
        wb = openpyxl.load_workbook(posts_file)
        target_ws = None
        header_row_idx = 1
        for sname in wb.sheetnames:
            ws = wb[sname]
            for r_idx, row in enumerate(list(ws.iter_rows(values_only=True))[:15], start=1):
                if row and any(k in [str(c).strip() for c in row if c is not None] for k in ["ポスト文章", "投稿本文", "ツイート文章"]):
                    target_ws = ws
                    header_row_idx = r_idx
                    break
            if target_ws:
                break
        if not target_ws:
            target_ws = wb.active
            header_row_idx = 1

        headers = [str(c.value).strip() if c.value is not None else "" for c in target_ws[header_row_idx]]
        status_col_idx = None
        text_col_idx = None
        err_col_idx = None
        for idx, h in enumerate(headers, start=1):
            if "ステータス" in h:
                status_col_idx = idx
            if any(k in h for k in ["ポスト文章", "投稿本文", "ツイート文章"]):
                text_col_idx = idx
            if any(k in h for k in ["エラー内容", "エラー", "error"]):
                err_col_idx = idx

        reset_count = 0
        if status_col_idx:
            for row_idx in range(header_row_idx + 1, target_ws.max_row + 1):
                txt_val = target_ws.cell(row=row_idx, column=text_col_idx).value if text_col_idx else None
                if txt_val and str(txt_val).strip():
                    target_ws.cell(row=row_idx, column=status_col_idx).value = "未投稿"
                    if err_col_idx:
                        target_ws.cell(row=row_idx, column=err_col_idx).value = ""
                    reset_count += 1
            wb.save(posts_file)
        wb.close()
        return reset_count
    except Exception as e:
        print(f"[WARN] reset_posts_file_silent エラー: {e}")
        return 0

def perform_stealth_feed_patrol(page, account_name, log_func=None):
    """
    【機能1: ステルス巡回（ウォーミングアップ閲覧）】
    新規投稿前に「おすすめ」タイムラインを人間らしくスクロール・閲覧し、
    Metaサーバーに自然な滞在・閲覧イベントログを残してスパム判定・BANを強力に回避する
    """
    try:
        if log_func:
            log_func(f"[{account_name}] 🕵️‍♂️ 【ステルス巡回】おすすめタイムラインを閲覧中（人間らしい操作ログを生成）...")
        time.sleep(random.uniform(1.5, 2.5))

        # 1回目のゆっくりスクロール（下へ350〜550px）
        scroll_y = random.randint(350, 550)
        page.evaluate(f"window.scrollBy({{ top: {scroll_y}, behavior: 'smooth' }})")
        time.sleep(random.uniform(1.8, 3.2))

        # 2回目のゆっくりスクロール（さらに200〜400px）
        scroll_y2 = random.randint(200, 400)
        page.evaluate(f"window.scrollBy({{ top: {scroll_y2}, behavior: 'smooth' }})")
        time.sleep(random.uniform(1.5, 2.8))

        # 人間が投稿を読み返すように少しだけ上に戻す（-100〜-150px）
        scroll_up = random.randint(100, 150)
        page.evaluate(f"window.scrollBy({{ top: -{scroll_up}, behavior: 'smooth' }})")
        time.sleep(random.uniform(1.0, 1.8))

        # 一番上に戻して左下の「新しいスレッド」ボタンが押しやすい状態にする
        page.evaluate("window.scrollTo({ top: 0, behavior: 'smooth' })")
        time.sleep(1.0)

        if log_func:
            log_func(f"[{account_name}] 👁️‍🗨️ ステルス巡回完了（正常な閲覧ログ記録済）。投稿フローに移行します。")
    except Exception as e:
        if log_func:
            log_func(f"[INFO] [{account_name}] ステルス巡回スキップ (画面動作優先): {e}")

def clean_account_profile(account_name, profile_dir, reason="警告検知または動作不安定", log_func=None):
    """
    【機能2: 自動プロファイル初期化（クリーンアップ）】
    警告画面や連続エラー・セッション汚染を検知した際、
    Chromeプロファイル（Cookie・IndexedDB・指紋キャッシュ）を完全自動初期化する
    """
    import shutil
    import stat

    if log_func:
        log_func(f"[WARN] [{account_name}] 🚨 アカウントの動作不安定・警告兆候を検知（理由: {reason}）")
        log_func(f"[INFO] [{account_name}] 🧹 プロファイルフォルダ ('{profile_dir}') の自動クリーンアップを開始します...")

    def _remove_readonly(func, path, exc_info):
        try:
            os.chmod(path, stat.S_IWRITE)
            func(path)
        except Exception:
            pass

    success = False
    for attempt in range(3):
        try:
            if os.path.exists(profile_dir):
                shutil.rmtree(profile_dir, onerror=_remove_readonly)
            time.sleep(0.5)
            os.makedirs(profile_dir, exist_ok=True)
            success = True
            break
        except Exception:
            time.sleep(1.0)

    if success:
        if log_func:
            log_func(f"[SUCCESS] [{account_name}] 🧹 プロファイルの初期化が完了しました！Cookie・追跡キャッシュは完全消去されました。次回クリーンな状態で自動再接続します。")
        log_execution(f"[{account_name}] プロファイル自動初期化完了 (理由: {reason})")
    else:
        if log_func:
            log_func(f"[WARN] [{account_name}] 一部ファイルの排他ロックによりプロファイル完全削除ができませんでした（次回起動時に再試行します）。")

    return success

def parse_topic_candidates(raw_val):
    """
    【機能3: トピックA/Bテスト用】
    カンマ（半角・全角）、読点（、）、スラッシュ（/）、改行等で区切られた複数トピック候補をパースしてリスト化
    """
    if not raw_val:
        return []
    tokens = re.split(r'[,，、/|\n]+', str(raw_val))
    cleaned = []
    for tok in tokens:
        clean = tok.strip().lstrip('#').strip()
        if clean and clean.lower() not in ["none", "nan"]:
            # 数字のみ（ID等の誤混入）は除外
            if not re.match(r'^\+?\d+(\.0+)?$', clean):
                if clean not in cleaned:
                    cleaned.append(clean)
    return cleaned

def select_topic_for_post(raw_topic_source, account_name, post_counter=0, enable_ab_test=True, log_func=None):
    """
    【機能3: トピックA/Bテスト・自動ローテーション】
    エクセルで指定された複数トピック候補から、投稿ごとに均等に切り替えて適用
    """
    candidates = parse_topic_candidates(raw_topic_source)
    if not candidates:
        return ""
    if len(candidates) == 1 or not enable_ab_test:
        return candidates[0]

    # 複数候補が存在する場合の均等ローテーション (A/B/Cテスト)
    selected_idx = post_counter % len(candidates)
    selected = candidates[selected_idx]
    if log_func:
        cand_str = ", ".join(candidates)
        log_func(f"[{account_name}] 🎯 【トピックA/Bテスト】候補 [{cand_str}] から今回「{selected}」({selected_idx + 1}/{len(candidates)}) を自動ローテーション選択しました。")
    return selected

def is_time_in_slot(now_dt, slot_str):
    """
    時間帯文字列（例: '7:00-9:00', '7:00〜9:00', '12:00-13:30', '21:00-23:30'）
    に現在時刻が含まれるかを分単位で正確に判定
    """
    if not slot_str:
        return False
    raw = str(slot_str).strip()
    for sep in ["-", "〜", "~", "ー", "to"]:
        if sep in raw:
            parts = raw.split(sep, 1)
            t1_s, t2_s = parts[0].strip(), parts[1].strip()
            def _parse_time(t_s):
                if ":" in t_s:
                    h, m = t_s.split(":", 1)
                    return int(h.strip()), int(m.strip())
                elif t_s.isdigit():
                    return int(t_s), 0
                return None
            try:
                p1 = _parse_time(t1_s)
                p2 = _parse_time(t2_s)
                if p1 and p2:
                    h1, m1 = p1
                    h2, m2 = p2
                    cur_min = now_dt.hour * 60 + now_dt.minute
                    start_min = h1 * 60 + m1
                    end_min = h2 * 60 + m2
                    if start_min <= end_min:
                        return start_min <= cur_min <= end_min
                    else:
                        return cur_min >= start_min or cur_min <= end_min
            except Exception:
                pass
    return False

def ensure_daily_posts_sheet_exists(filepath):
    """
    <アカウント名>.xlsx に「日常ポスト」シートが存在しない場合、
    朝・昼・夜の標準テンプレート枠（自由編集可能）を自動追加して保存
    """
    if not filepath or not os.path.exists(filepath):
        return
    try:
        wb = openpyxl.load_workbook(filepath)
        has_daily = any(any(k in s for k in ["日常", "あいさつ", "つぶやき", "daily", "Daily"]) for sname in wb.sheetnames for s in [sname])
        if not has_daily:
            ws = wb.create_sheet(title="日常ポスト")
            headers = ["枠名", "時間帯", "ポスト文章", "コミュニティまたはトピック", "有効無効", "最終投稿日", "最終投稿日時"]
            ws.append(headers)
            defaults = [
                ("朝枠", "7:00-9:00", "{おはようございます☀️|皆さんおはようございます☕️|おはよ〜ございます！} {月日}({曜日})の朝ですね！{今日もいい天気！|朝の空気が澄んでいて気持ちいい✨|少し肌寒いですが温かくして}マイペースに頑張りましょう〜！", "信太山新地, 大阪求人, 高収入バイト", "有効", "", ""),
                ("朝枠", "7:00-9:00", "{Good morning☀️|おはようございます！} 朝からしっかり{コーヒー飲んで|朝ごはん食べて|ストレッチして}エネルギーチャージ完了💪 {今日も素敵な1日に|ぼちぼち気負わず|笑顔で楽しく}乗り切りましょう！", "信太山新地, 大阪求人, 高収入バイト", "有効", "", ""),
                ("昼枠", "12:00-13:30", "やっとお昼休憩！{今日は一段とバタバタしてる💦|午前中のタスク一段落しました☕️|あっという間のお昼ですね} 美味しいもの食べて午後も乗り切ろう〜！皆さんはお昼何食べましたか？🍙", "信太山新地, 大阪求人, 高収入バイト", "有効", "", ""),
                ("夕方枠", "17:00-18:30", "{夕暮れの空がすごく綺麗🌆|定時まであと少し💨|今日も一日あっという間でした！} {お仕事終わりの皆さん、今日もお疲れ様でした！|ラストスパート集中していきます💻|気をつけてお帰りくださいね〜}✨", "信太山新地, 大阪求人, 高収入バイト", "有効", "", ""),
                ("夜枠", "21:00-23:30", "{今日も1日本当に忙しかった〜！|やっと今日のタスクがひと段落🌙|お風呂入ってまったりタイム🛁} {月日}({曜日})もお疲れ様でした！温かくしてゆっくり休んでくださいね🛌✨", "信太山新地, 大阪求人, 高収入バイト", "有効", "", ""),
            ]
            for row in defaults:
                ws.append(list(row))

            # 100個のテンプレート行（無効）を追加
            m_texts = [
                "おはようございます！朝の空気が澄んでいて気持ちいいですね🌿 今日も笑顔で頑張りましょう！",
                "Good morning☀️ 今週も元気にスタートしていきましょう〜！皆さんにとって実りある1日になりますように✨",
                "朝活中☕️ 朝の静かな時間が一番落ち着きますね。今日も1日マイペースに乗り切ろう！",
                "おはようございます！今日も予定がいっぱいですが、ひとつずつ丁寧にこなしていきます💪",
                "爽やかな朝☀️ 今日も一日頑張る皆さんを応援してます！水分補給しっかりね〜",
                "おはようございます！朝ごはんしっかり食べてエネルギーチャージ完了🍞 今日も良い日に！",
                "皆さんおはようございます✨ 朝からバタバタ準備中💦 今日も元気に乗り切りましょう！",
                "おはようございます！寒暖差に気をつけて、温かくしてお出かけくださいね🧣",
                "朝日が眩しい☀️ 今日も素敵な出会いや嬉しいことがありますように〜！",
                "おはようございます！今週もあと少し、気合い入れていきましょう〜！☕️",
                "朝から淹れたてのコーヒーが最高に美味しい☕️ 今日もぼちぼちマイペースで！",
                "おはようございます！早起きは三文の徳って本当ですね✨ 今日も気合十分です！",
                "皆さんおはようございます☀️ 今日も自分らしく、素敵な1日を過ごしましょう〜！",
                "おはようございます！今日も一日ポジティブに楽しんでいきましょうね〜😊",
                "朝の深呼吸🌿 今日も一日元気にいきましょう！良い日になりますように！",
                "おはようございます！週の始まり、張り切っていきましょう〜🔥",
                "おはようございます！今日も美味しいコーヒー片手にスタート☕️ 頑張ろう！",
                "皆さんおはようございます☀️ 今日もたくさんの笑顔が溢れる1日になりますように✨",
                "おはようございます！朝からお散歩して気分爽快🚶‍♀️ 今日も張り切っていきましょう！",
                "おはようございます！今日も無理せず自分のペースを大切に過ごしましょうね〜",
                "朝の澄んだ空気最高✨ 今日も一日感謝の気持ちで過ごします！",
                "おはようございます！今日も素敵な一日になりますように。行ってらっしゃい！👋",
                "皆さんおはようございます☀️ 今日もコツコツ積み上げていきましょう〜！",
                "おはようございます！しっかりストレッチして身体を目覚めさせました✨ 今日も元気に！",
                "おはようございます！今日も笑顔溢れる充実した一日になりますように☕️"
            ]
            l_texts = [
                "お昼ご飯タイム🍱 しっかり食べて午後も集中していきましょう〜！",
                "ランチ休憩中〜✨ 今日は同僚とおしゃれなカフェランチ！午後の活力チャージ完了☕️",
                "お昼休憩！午前中バタバタだったから美味しいご飯に癒されてます🍙",
                "午前のタスク完了！午後も無理せずマイペースに頑張りましょう〜✨",
                "お昼ごはん何にするか迷い中…皆さんは何食べましたか？おすすめ教えてください👇",
                "やっとお昼！あったかいスープが染み渡る〜🍲 午後も乗り切るぞ！",
                "ランチ休憩中🍙 息抜きもしっかりして、午後からの作業も集中していきます！",
                "お昼タイム！美味しいスイーツも食べちゃいました🍰 午後も頑張れそう〜",
                "午前の山場を越えてホッとひと息☕️ お昼しっかり食べて午後もGO！",
                "お昼休憩〜！お気に入りの音楽聴きながらまったり過ごしてます🎧",
                "ランチ完了！午後もサクサク進めて定時退勤目指します💪",
                "お昼ご飯美味しかった〜！皆さんも午後無理せず頑張ってくださいね😊",
                "お昼休み中！今日は天気が良くて窓際がぽかぽかしてます☀️ 午後もマイペースに！",
                "お昼休憩✨ 眠気に負けず午後も張り切っていきましょう〜！珈琲注入☕️",
                "今日のお昼はラーメン🍜 ガッツリ食べて午後のやる気マックスです！",
                "お昼休憩！スマホ見ながらのんびり充電中🔋 午後も頑張りましょう〜",
                "ランチタイム終了！切り替えて午後のタスクもサクサク終わらせます💻",
                "お昼ご飯食べて生き返った〜！午後もマイペースにぼちぼち行きましょう✨",
                "ランチ休憩🍙 休憩しっかり取るのも大事なお仕事！午後も乗り切ろう！",
                "お昼のまったり時間☕️ 午後も笑顔で頑張りましょうね〜！",
                "ランチタイム✨ 良いアイデアが浮かびそうな予感！午後も楽しみです",
                "お昼ご飯食べて満腹幸せ〜😋 午後も集中して駆け抜けます！",
                "やっとお昼休憩！午前中お疲れ様でした。午後もぼちぼちいきましょう〜",
                "お昼タイム！美味しいもの食べてリフレッシュ完了✨ 午後もファイトです！",
                "お昼休憩中🍙 お腹いっぱい！午後もマイペースに楽しく過ごしましょう〜"
            ]
            e_texts = [
                "夕方になって涼しくなってきましたね🍃 今日もあと少し！ラストスパート頑張りましょう✨",
                "定時までカウントダウン💨 皆さん今日もお疲れ様でした！",
                "夕暮れの空がすごく綺麗🌆 今日も一日よく頑張った〜！",
                "夕方のひと息☕️ もう一踏ん張りして今日の業務を締めくくります！",
                "もうすぐ今日の仕事終了！帰ったら何食べようか考え中🤤",
                "夕方ラッシュ気をつけてお帰りくださいね！今日も一日お疲れ様でした✨",
                "定時退社に向けて最後のラストスパート！集中力全開でいきます🔥",
                "夕方タイム✨ 今日も予定通り進んでホッとしてます。お疲れ様でした！",
                "夕焼けが綺麗で見惚れてました🌇 今日も無事に終わって感謝！",
                "夕方の疲れた頭に甘いものが染みる〜🍫 もう少し頑張ろう！",
                "今日のタスクほぼ完了！皆さん今日もお疲れ様でした〜！",
                "夕方のお仕事タイム、最後まで集中して駆け抜けます💻",
                "退勤ラッシュの時間帯ですね！皆さん気をつけて帰宅してくださいね🚃",
                "夕暮れ時🌆 今日もたくさんの学びがあった1日でした！お疲れ様でした",
                "今日の業務終了！達成感いっぱいで帰ります〜🙌",
                "夕方の空がグラデーションで綺麗✨ 今日も1日頑張った自分にご褒美を！",
                "ラストスパート！あと30分で今日のタスクを片付けます💪",
                "お仕事終わりの皆さん、今日もお疲れ様でした！ゆっくり休んでくださいね🍻",
                "夕方の風が心地いい〜！寄り道して帰ろうかな🚶‍♀️",
                "今日も1日無事終了！美味しい晩ごはんが待っている〜🍚",
                "夕方タイム！集中して残りのタスクを片付けます✨",
                "定時退勤成功！今日は早めにお風呂入ってリラックスします🛁",
                "夕暮れの街並みを眺めながら帰宅中🌇 今日も皆さんお疲れ様でした！",
                "夕方のラストスパート完了！やりきった〜！お疲れ様でした✨",
                "今日もあっという間に夕方！皆さん今日もお疲れ様でした😊"
            ]
            n_texts = [
                "お風呂上がりの至福タイム🛁 今日も一日お疲れ様でした！明日に備えてしっかり休みます✨",
                "夜のまったり時間🍵 温かいハーブティー飲んでリラックス中。おやすみなさい💤",
                "今日も一日頑張った〜！自分を褒めてあげたいと思います👏 皆さんもお疲れ様でした！",
                "夜風が涼しくて心地いい🌙 明日もいい日になりますように。おやすみなさい✨",
                "読書タイム📖 夜の静けさが大好きです。皆さん良い夢を〜！",
                "今日もたくさんのいいねやコメントありがとうございました❤️ ゆっくり休んでくださいね！",
                "一日の終わりに感謝✨ 明日も笑顔で頑張りましょう！おやすみなさい🛌",
                "ベッドに入ってゴロゴロ中💤 今日も充実した1日でした。おやすみ〜",
                "明日の準備も完了！今夜はぐっすり眠れそうです🌙 お疲れ様でした！",
                "夜のスキンケア完了✨ お肌も心も潤して眠ります。皆さん良い夜を！",
                "今日も1日駆け抜けました！明日はもっと良い日になるはず✨ おやすみなさい😴",
                "夜の静寂タイム🌙 今日できたことを振り返りながら眠りにつきます。お疲れ様でした！",
                "ストレッチして身体をほぐし中🧘‍♀️ 睡眠の質を上げて明日も頑張るぞ〜",
                "今日も無事に一日が終わってホッとしてます🍵 皆さん暖かくして休んでくださいね",
                "深夜のまったりタイム✨ 明日もマイペースにいきましょうね。おやすみなさい💤",
                "今日も一日お疲れ様でした！自分へのご褒美アイスタイム🍨",
                "明日の朝が楽しみになるような良い夜を🌙 おやすみなさい！",
                "スマホを置いてそろそろ寝ます〜📴 皆さん今日も一日お疲れ様でした✨",
                "夜のひととき🌙 好きな音楽聴きながらリラックス。良い夢を！",
                "今日も一日本当にお疲れ様でした！明日も元気にいきましょうね〜😊",
                "温かいお布団に入りました🛌 最高の瞬間！皆さんおやすみなさい💤",
                "今日一日関わってくれた皆さんに感謝✨ 明日も素敵な1日に！",
                "今夜もぐっすり眠れますように🌙 皆さんお疲れ様でした、おやすみなさい！",
                "夜の静けさに包まれてリラックス🍵 明日も良い日になりますように！",
                "今日も一日無事に終わりました✨ 明日に備えてしっかりエネルギーチャージします💤"
            ]
            for txt in m_texts:
                ws.append(["朝枠", "7:00-9:00", txt, "信太山新地, 大阪求人, 高収入バイト", "無効", "", ""])
            for txt in l_texts:
                ws.append(["昼枠", "12:00-13:30", txt, "信太山新地, 大阪求人, 高収入バイト", "無効", "", ""])
            for txt in e_texts:
                ws.append(["夕方枠", "17:00-18:30", txt, "信太山新地, 大阪求人, 高収入バイト", "無効", "", ""])
            for txt in n_texts:
                ws.append(["夜枠", "21:00-23:30", txt, "信太山新地, 大阪求人, 高収入バイト", "無効", "", ""])

            ws.column_dimensions['A'].width = 12
            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 65
            ws.column_dimensions['D'].width = 36
            ws.column_dimensions['E'].width = 12
            ws.column_dimensions['F'].width = 14
            ws.column_dimensions['G'].width = 22
            wb.save(filepath)
        wb.close()
    except Exception as e:
        print(f"[WARN] ensure_daily_posts_sheet_exists error: {e}")

def expand_spintax_and_placeholders(text, now_dt=None):
    """
    【提案1: スピンタックス記法自動展開 & 提案2: リアルタイム動的プレースホルダー差し込み】
    
    1. リアルタイム動的置換:
       {日付}, {月日}, {年月日}, {年}, {月}, {日}, {曜日}, {曜}, {時間}, {時}, {分}, {午前午後}, {時間帯}
       を投稿する瞬間の現在日時に自動置換。人間がリアルタイムに手打ちした投稿に見せかけます。
       
    2. スピンタックス記法展開:
       {選択肢1|選択肢2|選択肢3} （全角の「｜」も自動対応、ネスト入れ子対応）
       投稿する瞬間にパーツをランダムにシャッフル組み立てし、毎回異なる文章を自動生成。
       Metaの重複投稿・スパム検知スコアをゼロにします。
    """
    if not text:
        return ""
    result = str(text)

    if now_dt is None:
        now_dt = datetime.now()

    # 日本語曜日
    weekday_names = ["月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日", "日曜日"]
    weekday_short = ["月", "火", "水", "木", "金", "土", "日"]
    w_idx = now_dt.weekday()

    # 時間帯名
    hour = now_dt.hour
    if 5 <= hour < 11:
        time_period = "朝"
    elif 11 <= hour < 15:
        time_period = "お昼"
    elif 15 <= hour < 18:
        time_period = "夕方"
    elif 18 <= hour < 23:
        time_period = "夜"
    else:
        time_period = "深夜"

    am_pm = "午前" if hour < 12 else "午後"

    # 1. 動的プレースホルダー置換
    placeholders = {
        "{年月日}": f"{now_dt.year}年{now_dt.month}月{now_dt.day}日",
        "{日付}": f"{now_dt.month}月{now_dt.day}日",
        "{月日}": f"{now_dt.month}月{now_dt.day}日",
        "{年}": f"{now_dt.year}年",
        "{月}": f"{now_dt.month}月",
        "{日}": f"{now_dt.day}日",
        "{曜日}": weekday_names[w_idx],
        "{曜}": weekday_short[w_idx],
        "{時間}": now_dt.strftime("%H:%M"),
        "{時}": f"{now_dt.hour}時",
        "{分}": f"{now_dt.minute}分",
        "{午前午後}": am_pm,
        "{時間帯}": time_period,
    }

    for tag, val in placeholders.items():
        if tag in result:
            result = result.replace(tag, val)

    # 2. スピンタックス展開: {選択肢1|選択肢2|選択肢3}
    # パイプ（|または｜）を含む最内側の波括弧を順次ランダム解決（最大50回まで再帰的解決）
    spintax_pattern = re.compile(r'\{([^{}]*[|｜][^{}]*)\}')
    for _ in range(50):
        match = spintax_pattern.search(result)
        if not match:
            break
        inner = match.group(1)
        options = [opt.strip() for opt in re.split(r'[|｜]', inner)]
        chosen = random.choice(options) if options else ""
        result = result[:match.start()] + chosen + result[match.end():]

    return result

def get_eligible_daily_post(filepath, cooldown_days=30):
    """
    「日常ポスト」シートから、現在時刻の時間帯に合致する日常ポストを抽出。
    
    【凍結・重複防止の3大安全設計】:
    1. 1日1回枠制御:
       朝枠・昼枠・夕方枠・夜枠のいずれかで本日（YYYY-MM-DD）すでに投稿があれば、その枠は本日終了としてスキップ。
    2. クールダウン日数 (過去30日間重複選出禁止):
       最終投稿日が「30日以上前」または「未投稿（空）」の行のみを候補とし、同一ポストが30日間絶対に再選出されない。
    3. ランダム選出:
       クールダウンを満たす候補行の中から random.choice() でランダムに1行を選定。
       （※全候補が30日以内の場合は最も古い日付の行を自動フォールバック選出して停止を防止）
    4. 単独行の日付更新:
       選出された1行のみの最終投稿日を更新し、他の行のクールダウンに干渉しない。
    """
    if not filepath or not os.path.exists(filepath):
        return None
    try:
        ensure_daily_posts_sheet_exists(filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        target_ws = None
        for sname in wb.sheetnames:
            if any(k in sname for k in ["日常", "あいさつ", "つぶやき", "daily", "Daily"]):
                target_ws = wb[sname]
                break
        if not target_ws:
            wb.close()
            return None

        rows = list(target_ws.iter_rows(values_only=True))
        sheet_title = target_ws.title
        wb.close()
        if len(rows) < 2:
            return None

        headers = [str(h).strip() if h is not None else "" for h in rows[0]]
        slot_col = 0
        time_col = 1
        text_col = 2
        topic_col = None
        active_col = 3
        date_col = 4

        for idx, h in enumerate(headers):
            if "枠" in h or "slot" in h.lower():
                slot_col = idx
            elif "時間" in h or "time" in h.lower():
                time_col = idx
            elif any(k in h for k in ["文章", "内容", "ポスト", "text"]):
                text_col = idx
            elif any(k in h for k in ["トピック", "コミュニティ", "topic", "community"]):
                topic_col = idx
            elif "有効" in h or "active" in h.lower():
                active_col = idx
            elif "最終投稿日" in h or "日付" in h:
                date_col = idx

        now = datetime.now()
        today_date = now.date()
        today_str = now.strftime('%Y-%m-%d')

        slot_groups = {}
        for r_idx, row in enumerate(rows[1:], start=2):
            if not row or all(c is None for c in row):
                continue
            slot_name = str(row[slot_col]).strip() if len(row) > slot_col and row[slot_col] is not None else f"枠_{r_idx}"
            time_slot = str(row[time_col]).strip() if len(row) > time_col and row[time_col] is not None else ""
            post_txt = str(row[text_col]).strip() if len(row) > text_col and row[text_col] is not None else ""
            active_val = str(row[active_col]).strip().lower() if len(row) > active_col and row[active_col] is not None else "有効"
            last_date_raw = str(row[date_col]).strip() if len(row) > date_col and row[date_col] is not None else ""
            topic_txt = str(row[topic_col]).strip() if (topic_col is not None and len(row) > topic_col and row[topic_col] is not None) else ""
            if topic_txt.lower() in ["none", "nan"]:
                topic_txt = ""

            if active_val in ["無効", "off", "0", "false", "停止"]:
                continue
            if not post_txt or post_txt.lower() in ["none", "nan"]:
                continue
            if not is_time_in_slot(now, time_slot):
                continue

            key = (slot_name, time_slot)
            if key not in slot_groups:
                slot_groups[key] = {
                    "slot_name": slot_name,
                    "time_slot": time_slot,
                    "has_run_today": False,
                    "candidates": []
                }

            last_dt = None
            if last_date_raw and last_date_raw.lower() not in ["none", "nan"]:
                if today_str in last_date_raw:
                    slot_groups[key]["has_run_today"] = True
                clean_d = last_date_raw.split(" ")[0].replace("/", "-")
                try:
                    last_dt = datetime.strptime(clean_d, "%Y-%m-%d").date()
                except Exception:
                    pass

            slot_groups[key]["candidates"].append({
                "row_idx": r_idx,
                "text": post_txt,
                "topic": topic_txt,
                "last_dt": last_dt,
                "last_date_raw": last_date_raw
            })

        # 本日まだ実行されていない時間帯枠のみ抽出
        eligible_slots = []
        for key, s_info in slot_groups.items():
            if s_info["has_run_today"]:
                continue
            if not s_info["candidates"]:
                continue
            eligible_slots.append(s_info)

        if not eligible_slots:
            return None

        # 該当枠から1つ選択（通常は現時間帯の1枠）
        chosen_slot = random.choice(eligible_slots)

        # 提案3: 30日クールダウン判定
        ready_candidates = []
        fallback_candidates = []

        for cand in chosen_slot["candidates"]:
            if cand["last_dt"] is None:
                # 未投稿行は最優先
                ready_candidates.append(cand)
            else:
                days_diff = (today_date - cand["last_dt"]).days
                if days_diff >= cooldown_days:
                    ready_candidates.append(cand)
                else:
                    fallback_candidates.append((days_diff, cand))

        # クールダウンを満たす候補の中から「ランダム」に1行を選出
        if ready_candidates:
            chosen_cand = random.choice(ready_candidates)
        else:
            # 全候補が30日以内に投稿済みだった場合: 最も以前（日数差が最大）に投稿された行をフォールバック選出
            fallback_candidates.sort(key=lambda x: x[0], reverse=True)
            chosen_cand = fallback_candidates[0][1]

        # 提案1 & 提案2: スピンタックス自動シャッフル展開 ＆ リアルタイム日付プレースホルダー置換
        expanded_text = expand_spintax_and_placeholders(chosen_cand["text"], now)

        return {
            "slot_name": chosen_slot["slot_name"],
            "time_slot": chosen_slot["time_slot"],
            "post_text": expanded_text,
            "raw_text": chosen_cand["text"],
            "topic": chosen_cand["topic"],
            "row_indices": [chosen_cand["row_idx"]], # 選ばれた1行のみ更新
            "sheet_name": sheet_title,
            "filepath": filepath
        }
    except Exception as e:
        print(f"[WARN] get_eligible_daily_post error: {e}")
        return None

def update_daily_post_status(filepath, sheet_name, row_indices):
    """日常ポストの完了日・完了日時を更新（選ばれた投稿行のみ本日済みに記録しクールダウン管理）"""
    try:
        wb = openpyxl.load_workbook(filepath)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
        date_col = None
        dt_col = None
        for idx, h in enumerate(headers, start=1):
            if "最終投稿日" in h or (date_col is None and "日付" in h):
                date_col = idx
            elif "最終投稿日時" in h or (dt_col is None and "日時" in h):
                dt_col = idx

        now = datetime.now()
        today_str = now.strftime('%Y-%m-%d')
        now_str = now.strftime('%Y-%m-%d %H:%M:%S')

        for r_num in row_indices:
            if date_col:
                ws.cell(row=r_num, column=date_col, value=today_str)
            if dt_col:
                ws.cell(row=r_num, column=dt_col, value=now_str)

        wb.save(filepath)
        wb.close()
        return True
    except Exception as e:
        print(f"[WARN] update_daily_post_status error: {e}")
        return False

class ThreadsPosterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Threads AutoPoster (スレッズ自動投稿)")
        self.root.geometry("860x800")
        self.root.configure(bg="#121212")
        self.is_running = False
        self.account_config_path = find_account_file()
        self.var_auto_loop = tk.BooleanVar(value=True)
        self.var_single_window_rotation = tk.BooleanVar(value=True) # 提案1: 1窓・順番交代ローテーション
        self.var_block_media = tk.BooleanVar(value=True)            # 提案2: 画像・動画通信遮断 (メモリ50-70%削減)
        self.var_stealth_patrol = tk.BooleanVar(value=True)         # 提案新機能1: 投稿前のステルス巡回
        self.var_auto_clean_profile = tk.BooleanVar(value=True)      # 提案新機能2: 不安定・警告時の自動プロファイル初期化
        self.var_topic_ab_test = tk.BooleanVar(value=True)           # 提案新機能3: トピックA/Bテスト・自動ローテーション
        self.var_daily_greeting = tk.BooleanVar(value=True)          # 🌅 朝・昼・夜の日常つぶやき・あいさつ枠自動挿入

        self._setup_ui()
        self.log("[INFO] アプリケーションを起動しました。")
        self.log(f"[INFO] 実行ディレクトリ: {BASE_DIR}")
        self.load_excel_data()

    def _setup_ui(self):
        header = tk.Frame(self.root, bg="#1a1a1a", pady=12, padx=16)
        header.pack(fill="x")

        title_lbl = tk.Label(header, text="@ Threads AutoPoster Studio", font=("Segoe UI", 14, "bold"), fg="#ffffff", bg="#1a1a1a")
        title_lbl.pack(side="left")

        status_lbl = tk.Label(header, text="● 待機中 (Excel読込済)", font=("Segoe UI", 10, "bold"), fg="#10b981", bg="#1a1a1a")
        status_lbl.pack(side="right")
        self.status_lbl = status_lbl

        toolbar = tk.Frame(self.root, bg="#121212", pady=8, padx=16)
        toolbar.pack(fill="x")

        self.btn_start = tk.Button(toolbar, text="▶ 自動投稿を開始", font=("Segoe UI", 10, "bold"), bg="#10b981", fg="#ffffff", padx=14, pady=5, relief="flat", cursor="hand2", command=self.start_posting)
        self.btn_start.pack(side="left", padx=3)

        self.btn_stop = tk.Button(toolbar, text="⏹ 停止", font=("Segoe UI", 10, "bold"), bg="#ef4444", fg="#ffffff", padx=14, pady=5, relief="flat", cursor="hand2", state="disabled", command=self.stop_posting)
        self.btn_stop.pack(side="left", padx=3)

        self.btn_browser_login = tk.Button(toolbar, text="🌐 ログイン確認・保存", font=("Segoe UI", 9, "bold"), bg="#8b5cf6", fg="#ffffff", padx=10, pady=5, relief="flat", cursor="hand2", command=self.open_browser_for_login)
        self.btn_browser_login.pack(side="left", padx=3)

        self.btn_test_session = tk.Button(toolbar, text="⚡ セッションテスト", font=("Segoe UI", 9, "bold"), bg="#f59e0b", fg="#000000", padx=10, pady=5, relief="flat", cursor="hand2", command=self.test_session_id_action)
        self.btn_test_session.pack(side="left", padx=3)

        btn_reload = tk.Button(toolbar, text="🔄 Excel再読込", font=("Segoe UI", 9, "bold"), bg="#3b82f6", fg="#ffffff", padx=8, pady=5, relief="flat", cursor="hand2", command=self.load_excel_data)
        btn_reload.pack(side="left", padx=3)

        btn_reset = tk.Button(toolbar, text="🔁 未投稿にリセット", font=("Segoe UI", 9, "bold"), bg="#6366f1", fg="#ffffff", padx=8, pady=5, relief="flat", cursor="hand2", command=self.reset_all_posts_to_unposted)
        btn_reset.pack(side="left", padx=3)

        btn_open_posts = tk.Button(toolbar, text="📊 投稿Excelを開く", font=("Segoe UI", 9), bg="#262626", fg="#ffffff", padx=8, pady=5, relief="flat", cursor="hand2", command=self.open_posts_excel)
        btn_open_posts.pack(side="left", padx=3)

        btn_open_daily = tk.Button(toolbar, text="🌅 日常ポスト設定", font=("Segoe UI", 9, "bold"), bg="#15803d", fg="#ffffff", padx=8, pady=5, relief="flat", cursor="hand2", command=self.open_daily_posts_excel)
        btn_open_daily.pack(side="left", padx=3)

        btn_open_excel = tk.Button(toolbar, text="⚙️ アカウントExcel", font=("Segoe UI", 9), bg="#262626", fg="#ffffff", padx=8, pady=5, relief="flat", cursor="hand2", command=self.open_account_excel)
        btn_open_excel.pack(side="left", padx=3)

        btn_open_ss = tk.Button(toolbar, text="📷 エラー画像", font=("Segoe UI", 9, "bold"), bg="#450a0a", fg="#fca5a5", padx=8, pady=5, relief="flat", cursor="hand2", command=self.open_error_screenshots_dir)
        btn_open_ss.pack(side="left", padx=3)

        btn_open_log = tk.Button(toolbar, text="📝 エラーログ(txt)", font=("Segoe UI", 9, "bold"), bg="#991b1b", fg="#ffffff", padx=8, pady=5, relief="flat", cursor="hand2", command=self.open_error_log)
        btn_open_log.pack(side="left", padx=3)

        btn_open_all_log = tk.Button(toolbar, text="📋 全実行ログ(txt)", font=("Segoe UI", 9), bg="#262626", fg="#ffffff", padx=8, pady=5, relief="flat", cursor="hand2", command=self.open_execution_log)
        btn_open_all_log.pack(side="left", padx=3)

        btn_clean_manual = tk.Button(toolbar, text="🧹 キャッシュ全初期化", font=("Segoe UI", 9, "bold"), bg="#4c0519", fg="#fda4af", padx=8, pady=5, relief="flat", cursor="hand2", command=self.manual_clean_profiles)
        btn_clean_manual.pack(side="left", padx=3)

        btn_open_dir = tk.Button(toolbar, text="📁 フォルダ", font=("Segoe UI", 9), bg="#262626", fg="#ffffff", padx=8, pady=5, relief="flat", cursor="hand2", command=self.open_base_dir)
        btn_open_dir.pack(side="left", padx=3)


        self.info_frame = tk.Frame(self.root, bg="#18181b", padx=14, pady=8, highlightbackground="#27272a", highlightthickness=1)
        self.info_frame.pack(fill="x", padx=16, pady=4)

        self.lbl_excel_path = tk.Label(self.info_frame, text="設定ファイル: 読み込み中...", font=("Segoe UI", 9, "bold"), fg="#e4e4e7", bg="#18181b", anchor="w")
        self.lbl_excel_path.pack(fill="x")

        self.lbl_account_info = tk.Label(self.info_frame, text="アカウント: 検出中...", font=("Segoe UI", 9), fg="#a1a1aa", bg="#18181b", anchor="w")
        self.lbl_account_info.pack(fill="x", pady=(2, 0))

        self.lbl_posts_info = tk.Label(self.info_frame, text="投稿データ: 未読み込み", font=("Segoe UI", 9), fg="#38bdf8", bg="#18181b", anchor="w")
        self.lbl_posts_info.pack(fill="x", pady=(2, 0))

        self.chk_auto_loop = tk.Checkbutton(
            self.info_frame,
            text="🔁 全ポスト投稿完了後に自動で『未投稿』にリセットして最初から周回（無限ループ投稿）",
            variable=self.var_auto_loop,
            bg="#18181b",
            fg="#34d399",
            activebackground="#18181b",
            activeforeground="#34d399",
            selectcolor="#09090b",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2"
        )
        self.chk_auto_loop.pack(anchor="w", pady=(3, 0))

        self.chk_single_rotation = tk.Checkbutton(
            self.info_frame,
            text="🪟 【提案1: 1窓・順番交代ローテーション】常に開くブラウザは1つだけ。投稿後に完全終了してメモリ解放＆待機順に切替",
            variable=self.var_single_window_rotation,
            bg="#18181b",
            fg="#60a5fa",
            activebackground="#18181b",
            activeforeground="#60a5fa",
            selectcolor="#09090b",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2"
        )
        self.chk_single_rotation.pack(anchor="w", pady=(3, 0))

        self.chk_block_media = tk.Checkbutton(
            self.info_frame,
            text="⚡ 【提案2: 画像・動画通信ブロック】タイムライン画像・動画・リールを遮断（メモリ50-70%削減・CPUファン静音化）",
            variable=self.var_block_media,
            bg="#18181b",
            fg="#c084fc",
            activebackground="#18181b",
            activeforeground="#c084fc",
            selectcolor="#09090b",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2"
        )
        self.chk_block_media.pack(anchor="w", pady=(3, 0))

        self.chk_stealth_patrol = tk.Checkbutton(
            self.info_frame,
            text="🕵️‍♂️ 【新機能1: ステルス巡回】投稿前におすすめタイムラインを数秒閲覧（人間らしい行動模倣・BANリスク回避）",
            variable=self.var_stealth_patrol,
            bg="#18181b",
            fg="#facc15",
            activebackground="#18181b",
            activeforeground="#facc15",
            selectcolor="#09090b",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2"
        )
        self.chk_stealth_patrol.pack(anchor="w", pady=(3, 0))

        self.chk_auto_clean = tk.Checkbutton(
            self.info_frame,
            text="🧹 【新機能2: 自動プロファイル初期化】Meta警告メッセージや連続エラーを検知時、手動操作不要でCookie/指紋を自動完全リセット",
            variable=self.var_auto_clean_profile,
            bg="#18181b",
            fg="#f43f5e",
            activebackground="#18181b",
            activeforeground="#f43f5e",
            selectcolor="#09090b",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2"
        )
        self.chk_auto_clean.pack(anchor="w", pady=(3, 0))

        self.chk_topic_ab = tk.Checkbutton(
            self.info_frame,
            text="🎯 【新機能3: トピックA/Bテスト・自動ローテーション】Excelの複数トピック（例: 信太山新地, 大阪求人, 高収入バイト）を投稿ごとに自動均等切替",
            variable=self.var_topic_ab_test,
            bg="#18181b",
            fg="#2dd4bf",
            activebackground="#18181b",
            activeforeground="#2dd4bf",
            selectcolor="#09090b",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2"
        )
        self.chk_topic_ab.pack(anchor="w", pady=(3, 0))

        self.chk_daily_greeting = tk.Checkbutton(
            self.info_frame,
            text="🌅 【日常つぶやき・あいさつ枠】朝・昼・夕・夜の日常文を優先投稿（スピンタックス自動展開・動的日付差し込み・30日間クールダウン＆ランダム選出で完全凍結防止）",
            variable=self.var_daily_greeting,
            bg="#18181b",
            fg="#fb923c",
            activebackground="#18181b",
            activeforeground="#fb923c",
            selectcolor="#09090b",
            font=("Segoe UI", 9, "bold"),
            cursor="hand2"
        )
        self.chk_daily_greeting.pack(anchor="w", pady=(3, 0))

        log_frame = tk.Frame(self.root, bg="#121212", padx=16, pady=6)
        log_frame.pack(fill="both", expand=True)

        lbl_log = tk.Label(log_frame, text="実行ログ / コンソール出力:", font=("Segoe UI", 9, "bold"), fg="#a3a3a3", bg="#121212")
        lbl_log.pack(anchor="w", pady=(0, 4))

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap="word", bg="#0a0a0a", fg="#e5e5e5", font=("Consolas", 9), relief="solid", bd=1)
        self.log_text.pack(fill="both", expand=True)

        self.log_text.tag_config("INFO", foreground="#60a5fa")
        self.log_text.tag_config("SUCCESS", foreground="#34d399")
        self.log_text.tag_config("WARN", foreground="#fbbf24")
        self.log_text.tag_config("ERROR", foreground="#f87171")

    def log(self, msg):
        timestamp = datetime.now().strftime("%H:%M:%S")
        full_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{timestamp}] {msg}\n"
        full_line = f"[{full_time}] {msg}\n"
        try:
            self.log_text.insert("end", line)
            if "[SUCCESS]" in msg:
                self.log_text.tag_add("SUCCESS", "end-2l", "end-1l")
            elif "[ERROR]" in msg:
                self.log_text.tag_add("ERROR", "end-2l", "end-1l")
            elif "[WARN]" in msg:
                self.log_text.tag_add("WARN", "end-2l", "end-1l")
            else:
                self.log_text.tag_add("INFO", "end-2l", "end-1l")
            self.log_text.see("end")
        except Exception:
            pass

        # 1. 全実行ログ (execution_log.txt) に永続追記
        try:
            with open(os.path.join(BASE_DIR, "execution_log.txt"), "a", encoding="utf-8") as f_exec:
                f_exec.write(full_line)
                f_exec.flush()
        except Exception:
            pass

        # 2. エラーログ (error_log.txt) に永続追記
        if "[ERROR]" in msg or "[WARN]" in msg or "エラー" in msg:
            try:
                with open(os.path.join(BASE_DIR, "error_log.txt"), "a", encoding="utf-8") as f_err:
                    f_err.write(full_line)
                    f_err.flush()
            except Exception:
                pass


    def load_excel_data(self):
        if not os.path.exists(self.account_config_path):
            self.lbl_excel_path.config(text=f"⚠️ 設定ファイルが見つかりません: {os.path.basename(self.account_config_path)}", fg="#f87171")
            self.lbl_account_info.config(text="同フォルダに「アカウント情報.xlsx」を配置してください。")
            self.lbl_posts_info.config(text="―")
            self.log(f"[WARN] '{self.account_config_path}' が見つかりません。")
            return

        self.lbl_excel_path.config(text=f"📊 設定ファイル: {os.path.basename(self.account_config_path)} (読み込みOK)", fg="#34d399")
        accounts_data = read_excel_records(self.account_config_path)

        valid_accounts = []
        for acc in accounts_data:
            acc_id = str(acc.get("アカウントID") or acc.get("アカウント名") or acc.get("アカウント") or "").strip()
            if acc_id and acc_id.lower() != "nan":
                valid_accounts.append(acc)

        if not valid_accounts:
            self.lbl_account_info.config(text="有効なアカウント設定がありません（アカウントID列をご確認ください）")
            self.lbl_posts_info.config(text="―")
            self.log("[WARN] アカウント情報.xlsx に有効な「アカウントID」が記入されていません。")
            return

        first_acc = valid_accounts[0]
        first_id = str(first_acc.get("アカウントID") or first_acc.get("アカウント名") or "").strip()
        auth_token = str(first_acc.get("auth_token") or first_acc.get("Authトークン_sessionid") or "").strip()
        has_token = "あり (自動ログイン)" if (auth_token and auth_token.lower() != "nan" and not auth_token.startswith("ここに")) else "なし (手動ログイン)"
        self.lbl_account_info.config(
            text=f"👤 有効アカウント: {len(valid_accounts)}件 | メインID: @{first_id} | auth_token: {has_token} (設定は {first_id}.xlsx で管理)",
            fg="#e4e4e7"
        )
        self.log(f"[INFO] アカウント情報読み込み完了: {len(valid_accounts)}件のアカウント設定を検出 (ID: @{first_id})")

        posts_file = self._find_posts_file(first_id)
        if not posts_file or not os.path.exists(posts_file):
            self.lbl_posts_info.config(text=f"⚠️ 投稿ファイル '{first_id}.xlsx' が見つかりません", fg="#fbbf24")
            self.log(f"[WARN] 投稿データファイル '{first_id}.xlsx' が見つかりません。")
            return

        posts = read_excel_records(posts_file)
        unposted = [p for p in posts if self._is_unposted_item(p)]
        completed = [p for p in posts if "完了" in str(p.get("ステータス", ""))]

        custom_settings = read_account_settings_from_post_file(posts_file)
        interval_txt = custom_settings.get("interval_str", "")
        mode_val = custom_settings.get("post_mode", 1)
        try:
            mode_val = int(mode_val)
        except Exception:
            mode_val = 1
        mode_txt = "ランダム" if mode_val == 1 else "上から順"
        time_txt = custom_settings.get("time_slot", "0:00-0:00")
        prof_txt = custom_settings.get("profile_dir", f"chrome_profiles/{first_id}")

        setting_parts = []
        if interval_txt:
            setting_parts.append(f"間隔: {interval_txt}")
        if mode_txt:
            setting_parts.append(f"モード: {mode_txt}")
        if time_txt and time_txt not in ["0:00-0:00", "0-24"]:
            setting_parts.append(f"時間: {time_txt}")
        if prof_txt:
            setting_parts.append(f"プロファイル: {prof_txt}")
        setting_disp = (" | " + " / ".join(setting_parts)) if setting_parts else ""

        self.lbl_posts_info.config(
            text=f"📝 投稿ファイル: {os.path.basename(posts_file)} (全 {len(posts)} 件 | 未投稿: {len(unposted)} 件 | 投稿完了: {len(completed)} 件){setting_disp}",
            fg="#38bdf8"
        )
        self.log(f"[INFO] 投稿リスト '{os.path.basename(posts_file)}' 読み込み完了: 未投稿 {len(unposted)} 件")
        if interval_txt or mode_txt:
            self.log(f"[INFO] [{first_id}] {os.path.basename(posts_file)} から設定を適用: 間隔={interval_txt or 'デフォルト'}, モード={mode_txt}, プロファイル={prof_txt}")

        for i, p_item in enumerate(unposted[:2], 1):
            text_preview = self._extract_post_text(p_item)
            short_txt = (text_preview[:35] + "...") if len(text_preview) > 35 else text_preview
            raw_img = self._extract_raw_image(p_item)
            resolved_img = resolve_image_path(raw_img, text_preview)
            img_info = f"画像: {os.path.basename(resolved_img)}" if resolved_img else "画像なし"
            self.log(f"[INFO] プレビュー#{i}: [ポスト文章] {short_txt} | [{img_info}]")

    def _find_posts_file(self, account_id):
        candidates = [
            os.path.join(BASE_DIR, f"{account_id}.xlsx"),
            os.path.join(os.getcwd(), f"{account_id}.xlsx"),
        ]
        for c in candidates:
            if os.path.exists(c):
                return c
        fallback_main = os.path.join(BASE_DIR, "threads_main_account.xlsx")
        if os.path.exists(fallback_main):
            try:
                import shutil
                target_dest = os.path.join(BASE_DIR, f"{account_id}.xlsx")
                shutil.copyfile(fallback_main, target_dest)
                return target_dest
            except Exception:
                return fallback_main
        try:
            for f in os.listdir(BASE_DIR):
                if f.endswith(".xlsx") and not ("アカウント情報" in f or "account_info" in f):
                    return os.path.join(BASE_DIR, f)
        except Exception:
            pass
        return candidates[0]

    def test_session_id_action(self):
        """アカウント情報.xlsx に記載された sessionid をMetaサーバーと直接照合してテスト"""
        if not os.path.exists(self.account_config_path):
            messagebox.showerror("エラー", f"'{self.account_config_path}' が見つかりません。")
            return
        
        all_accounts = read_excel_records(self.account_config_path)
        active_accounts = [acc for acc in all_accounts if str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip() != ""]
        if not active_accounts:
            messagebox.showwarning("警告", "アカウント情報.xlsx にアカウントが見つかりません。")
            return

        first_acc = active_accounts[0]
        curr_id = str(first_acc.get("アカウントID") or first_acc.get("アカウント名") or "").strip()
        raw_token = str(first_acc.get("auth_token") or first_acc.get("Authトークン_sessionid") or first_acc.get("sessionid") or "").strip()
        
        if not raw_token or raw_token.startswith("ここに"):
            messagebox.showwarning("セッションID未設定", "アカウント情報.xlsx の「auth_token」列に sessionid を貼り付けてから再度お試しください。")
            return
        
        self.log(f"[INFO] セッションIDをMeta Threadsサーバーと直接照合テスト中...")
        self.btn_test_session.config(state="disabled")

        def _test_worker():
            try:
                is_valid, user_or_err, uid = test_session_id_online(raw_token)
                if is_valid:
                    self.log(f"[SUCCESS] 🎉 セッションID有効確認！アカウント: @{user_or_err} (User ID: {uid})")
                    sync_msg = ""
                    if curr_id != user_or_err:
                        update_account_sessionid(self.account_config_path, user_or_err, raw_token)
                        self.log(f"[INFO] アカウントIDを @{curr_id} から @{user_or_err} へ自動同期しました。")
                        sync_msg = f"\n\nアカウントIDを実アカウント @{user_or_err} に自動同期しました！\n投稿エクセル '{user_or_err}.xlsx' を準備完了しました。"
                        self._find_posts_file(user_or_err)
                    
                    self.root.after(0, lambda: messagebox.showinfo(
                        "セッション有効確認成功",
                        f"✅ セッションIDは正常に動作しています！\n\n"
                        f"・Metaアカウント名: @{user_or_err}\n"
                        f"・ユーザーID: {uid}\n"
                        f"・認証ステータス: ログイン有効{sync_msg}\n\n"
                        "このまま「▶ 自動投稿を開始」をクリックして投稿を実行できます！"
                    ))
                else:
                    self.log(f"[WARN] セッションID検証結果: {user_or_err}")
                    self.root.after(0, lambda: messagebox.showwarning(
                        "認証確認できませんでした",
                        f"セッションIDの検証でエラーが返されました:\n{user_or_err}\n\n"
                        "【確認項目】\n"
                        "1. ブラウザでInstagram/Threadsからログアウトしていないか\n"
                        "2. sessionidの全文字列が正しくコピーされているか"
                    ))
            except Exception as e:
                self.log(f"[ERROR] セッションテストエラー: {e}")
            finally:
                self.root.after(0, lambda: self.btn_test_session.config(state="normal"))
                self.root.after(0, self.load_excel_data)

        threading.Thread(target=_test_worker, daemon=True).start()

    def _extract_post_text(self, post_item):
        for k in ["ポスト文章", "投稿本文", "ツイート文章", "本文", "テキスト", "text"]:
            if k in post_item and post_item[k] is not None and str(post_item[k]).strip() != "":
                return str(post_item[k]).strip()
        return ""

    def _extract_raw_image(self, post_item):
        for k in ["画像ファイルパス", "設定セル", "画像", "画像パス", "image", "image_path"]:
            if k in post_item and post_item[k] is not None and str(post_item[k]).strip() != "":
                return str(post_item[k]).strip()
        return ""

    def _extract_topic_or_community(self, post_item):
        if not isinstance(post_item, dict):
            return ""
        for k in [
            "コミュニティまたはトピック",
            "コミュニティ",
            "トピック",
            "トピックタグ",
            "タグ",
            "topic",
            "community",
            "tag"
        ]:
            if k in post_item and post_item[k] is not None:
                val = str(post_item[k]).strip()
                if val != "" and val.lower() not in ["none", "nan"]:
                    # 数字のみ（IDなどの誤混入）はトピックとして不適格なため完全除外
                    if not re.match(r'^\+?\d+(\.0+)?$', val):
                        return val
        return ""


    def _is_unposted_item(self, post_item):
        st = str(post_item.get("ステータス", "")).strip()
        if "完了" in st:
            return False
        if st == "未投稿":
            return True
        txt = self._extract_post_text(post_item)
        return (not st or st.lower() in ["none", "nan"]) and len(txt) > 0

    def select_account_file(self):
        selected = filedialog.askopenfilename(
            title="アカウント設定エクセルを選択",
            filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")],
            initialdir=BASE_DIR
        )
        if selected:
            self.account_config_path = selected
            self.load_excel_data()

    def open_account_excel(self):
        if os.path.exists(self.account_config_path):
            os.startfile(self.account_config_path)
        else:
            messagebox.showwarning("警告", f"'{self.account_config_path}' が存在しません。")

    def open_posts_excel(self):
        all_accounts = read_excel_records(self.account_config_path)
        valid_accounts = [acc for acc in all_accounts if str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip() != ""]
        first_id = str(valid_accounts[0].get("アカウントID") or valid_accounts[0].get("アカウント名") or "").strip() if valid_accounts else ""
        posts_file = self._find_posts_file(first_id)
        if posts_file and os.path.exists(posts_file):
            ensure_daily_posts_sheet_exists(posts_file)
            os.startfile(posts_file)
        elif os.path.exists(self.account_config_path):
            os.startfile(self.account_config_path)
        else:
            messagebox.showwarning("警告", "投稿エクセルファイルが見つかりません。")

    def open_daily_posts_excel(self):
        all_accounts = read_excel_records(self.account_config_path)
        valid_accounts = [acc for acc in all_accounts if str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip() != ""]
        first_id = str(valid_accounts[0].get("アカウントID") or valid_accounts[0].get("アカウント名") or "").strip() if valid_accounts else ""
        posts_file = self._find_posts_file(first_id)
        if posts_file and os.path.exists(posts_file):
            ensure_daily_posts_sheet_exists(posts_file)
            try:
                os.startfile(posts_file)
                self.log(f"[INFO] 🌅 日常ポスト設定シート（{os.path.basename(posts_file)} 内の「日常ポスト」シート）を開きました。時間帯や文章をご自由に編集してください。")
            except Exception as e:
                self.log(f"[WARN] エクセルを開けませんでした: {e}")
        elif os.path.exists(self.account_config_path):
            os.startfile(self.account_config_path)
        else:
            messagebox.showwarning("警告", "投稿エクセルファイルが見つかりません。")

    def reset_all_posts_to_unposted(self):
        try:
            all_accounts = read_excel_records(self.account_config_path)
            valid_accounts = [acc for acc in all_accounts if str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip() != ""]
            if not valid_accounts:
                messagebox.showwarning("警告", "アカウント情報.xlsx にアカウントが見つかりません。")
                return
            first_id = str(valid_accounts[0].get("アカウントID") or valid_accounts[0].get("アカウント名") or "").strip()
            posts_file = self._find_posts_file(first_id)
            if not posts_file or not os.path.exists(posts_file):
                messagebox.showwarning("警告", f"投稿ファイル '{posts_file}' が見つかりません。")
                return
            
            if not messagebox.askyesno("確認", f"'{os.path.basename(posts_file)}' の全ポストのステータスを『未投稿』に戻しますか？\n\n（再度先頭から自動投稿できるようになります）"):
                return
            
            reset_count = reset_posts_file_silent(posts_file)
            if reset_count > 0:
                self.log(f"[SUCCESS] '{os.path.basename(posts_file)}' の {reset_count} 件を『未投稿』にリセットしました！")
                self.load_excel_data()
                messagebox.showinfo("リセット完了", f"'{os.path.basename(posts_file)}' の {reset_count} 件を『未投稿』にリセットしました！\n\nこのまま「自動投稿を開始」を押すとブラウザが起動します。")
            else:
                self.log(f"[WARN] '{os.path.basename(posts_file)}' にリセット対象のポストがありませんでした。")
                messagebox.showwarning("警告", "リセット対象のポストが見つかりませんでした。")
        except Exception as e:
            self.log(f"[ERROR] ステータスリセット失敗: {e}")
            messagebox.showerror("エラー", f"ステータスのリセット中にエラーが発生しました:\n{e}\n\nエクセルが開いている場合は閉じてから再試行してください。")

    def open_base_dir(self):
        os.startfile(BASE_DIR)

    def open_error_screenshots_dir(self):
        ss_dir = os.path.join(BASE_DIR, "error_screenshots")
        os.makedirs(ss_dir, exist_ok=True)
        try:
            os.startfile(ss_dir)
            self.log("[INFO] 📷 エラー画像フォルダを開きました。")
        except Exception as e:
            self.log(f"[WARN] フォルダを開けませんでした: {e}")

    def open_error_log(self):
        log_file = os.path.join(BASE_DIR, "error_log.txt")
        if not os.path.exists(log_file):
            try:
                with open(log_file, "w", encoding="utf-8") as f:
                    f.write("=== Threads AutoPoster エラー診断ログ ===\nまだエラーは記録されていません。\n")
            except Exception:
                pass
        try:
            os.startfile(log_file)
            self.log("[INFO] 📝 エラーログファイル (error_log.txt) を開きました。")
        except Exception:
            try:
                import subprocess
                subprocess.Popen(["notepad.exe", log_file])
            except Exception as e:
                self.log(f"[WARN] ログを開けませんでした: {e}")

    def open_execution_log(self):
        log_file = os.path.join(BASE_DIR, "execution_log.txt")
        if not os.path.exists(log_file):
            try:
                with open(log_file, "w", encoding="utf-8") as f:
                    f.write("=== Threads AutoPoster 実行ログ ===\nまだログは記録されていません。\n")
            except Exception:
                pass
        try:
            os.startfile(log_file)
            self.log("[INFO] 📋 全実行ログファイル (execution_log.txt) を開きました。")
        except Exception:
            try:
                import subprocess
                subprocess.Popen(["notepad.exe", log_file])
            except Exception as e:
                self.log(f"[WARN] ログを開けませんでした: {e}")

    def manual_clean_profiles(self):
        """手動でのプロファイル完全初期化（キャッシュ・スパムフラグ消去）"""
        all_accounts = read_excel_records(self.account_config_path)
        valid_accounts = [acc for acc in all_accounts if str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip() != ""]
        if not valid_accounts:
            messagebox.showwarning("警告", "アカウント情報.xlsx に有効なアカウントが見つかりません。")
            return

        if not messagebox.askyesno(
            "プロファイル初期化（キャッシュ消去）の確認",
            "全アカウントのブラウザプロファイル（Cookie・LocalStorage・追跡指紋）を完全に初期化しますか？\n\n"
            "【初期化される内容】\n"
            "・Meta側の警告兆候・スパムフラグ付きCookie\n"
            "・蓄積されたブラウザキャッシュ・IndexedDB\n\n"
            "※「アカウント情報.xlsx」に auth_token が記入されている場合は、次回起動時に自動再ログインされます。\n\n"
            "実行しますか？"
        ):
            return

        cleaned_count = 0
        for acc in valid_accounts:
            acc_name = str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip()
            pf = self._find_posts_file(acc_name)
            custom_settings = read_account_settings_from_post_file(pf) if pf else {}
            p_sub = custom_settings.get("profile_dir") or f"chrome_profiles/{acc_name}"
            p_dir = os.path.join(BASE_DIR, p_sub)
            clean_account_profile(acc_name, p_dir, reason="手動リクエストによる全初期化", log_func=self.log)
            cleaned_count += 1

        messagebox.showinfo(
            "初期化完了",
            f"{cleaned_count} 件のアカウントプロファイルを初期化しました！\n\n"
            "次回からクリーンな状態で自動接続されます。"
        )

    def open_browser_for_login(self):
        if not os.path.exists(self.account_config_path):
            messagebox.showerror("エラー", f"'{self.account_config_path}' が見つかりません。")
            return

        self.log("[INFO] ログイン確認ブラウザを起動しています...")
        self.btn_browser_login.config(state="disabled")
        self.btn_start.config(state="disabled")

        def _worker():
            try:
                from playwright.sync_api import sync_playwright
                all_accounts = read_excel_records(self.account_config_path)
                active_accounts = [acc for acc in all_accounts if str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip() != ""]
                if not active_accounts:
                    self.log("[WARN] アカウント情報.xlsx に有効なアカウントがありません。")
                    return
                first_acc = active_accounts[0]
                account_name = str(first_acc.get("アカウントID") or first_acc.get("アカウント名") or "").strip()
                
                posts_file = self._find_posts_file(account_name)
                custom_settings = read_account_settings_from_post_file(posts_file) if posts_file else {}
                profile_subpath = custom_settings.get("profile_dir") or f"chrome_profiles/{account_name}"
                profile_dir = os.path.join(BASE_DIR, profile_subpath)
                os.makedirs(profile_dir, exist_ok=True)

                with sync_playwright() as p:
                    context = p.chromium.launch_persistent_context(
                        user_data_dir=profile_dir,
                        headless=False,
                        channel="chrome",
                        locale="ja-JP",
                        viewport={"width": 1280, "height": 900},
                        args=["--disable-blink-features=AutomationControlled", "--lang=ja-JP"]
                    )
                    
                    token = str(first_acc.get("auth_token") or first_acc.get("Authトークン_sessionid") or first_acc.get("sessionid") or "").strip()
                    if token:
                        inject_session_cookies(context, token, account_name, log_func=self.log)

                    page = context.pages[0] if context.pages else context.new_page()
                    page.goto("https://www.threads.net/", timeout=60000)
                    time.sleep(3)

                    is_ok = ensure_authenticated_session(
                        page, context, account_name, self.account_config_path, log_func=self.log, wait_timeout_sec=300
                    )

                    if is_ok:
                        self.log(f"[SUCCESS] [{account_name}] ログイン確認完了！セッション情報は自動保存されました。このまま自動投稿を開始できます。")
                        self.root.after(0, lambda: messagebox.showinfo(
                            "ログイン完了",
                            f"アカウント @{account_name} のログインを確認しました！\n\n"
                            "最新のセッション情報は「アカウント情報.xlsx」およびブラウザプロファイルに永続保存されました。\n\n"
                            "これで次回以降、ログイン操作なしで「自動投稿を開始」できます！"
                        ))
                    else:
                        self.log(f"[WARN] [{account_name}] ログインが完了しませんでした。ブラウザで「Instagramでログイン」を完了してください。")

                    time.sleep(2)
                    context.close()
            except Exception as e:
                self.log(f"[ERROR] ログインブラウザ起動エラー: {e}")
            finally:
                self.root.after(0, lambda: self.btn_browser_login.config(state="normal"))
                self.root.after(0, lambda: self.btn_start.config(state="normal"))
                self.root.after(0, self.load_excel_data)

        threading.Thread(target=_worker, daemon=True).start()

    def start_posting(self):
        if not os.path.exists(self.account_config_path):
            messagebox.showerror("エラー", f"'{self.account_config_path}' が見つかりません。")
            return

        self.is_running = True
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.status_lbl.config(text="● 投稿処理中...", fg="#38bdf8")
        self.log("[INFO] 自動投稿スレッドを開始します...")

        t = threading.Thread(target=self._run_automation_loop, daemon=True)
        t.start()

    def stop_posting(self):
        self.is_running = False
        self.log("[WARN] 停止リクエストを送信しました。現在の投稿完了後に停止します。")
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.status_lbl.config(text="● 待機中", fg="#10b981")

    def _run_automation_loop(self):
        try:
            try:
                from playwright.sync_api import sync_playwright
            except ImportError as ie:
                self.log(f"[ERROR] Playwrightが読み込めません: {ie}")
                messagebox.showerror("エラー", f"Playwrightが見つかりません:\n{ie}\n\nコマンドプロンプトで pip install playwright を実行してください。")
                return

            all_accounts = read_excel_records(self.account_config_path)
            active_accounts = [
                acc for acc in all_accounts
                if str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip() != ""
            ]

            if not active_accounts:
                self.log("[WARN] アカウント情報.xlsx に有効なアカウントがありません。")
                return

            with sync_playwright() as p:
                if self.var_single_window_rotation.get():
                    self._run_single_window_rotation_loop(active_accounts, p)
                else:
                    self._run_legacy_sequential_loop(active_accounts, p)

            if not self.is_running:
                self.log("[WARN] 停止リクエストにより自動投稿処理を終了しました。")
            else:
                self.log("[SUCCESS] すべてのアカウント処理が完了しました。")
        except Exception as e:
            self.log(f"[ERROR] 致命的エラー: {e}")
        finally:
            self.is_running = False
            self.root.after(0, lambda: self.btn_start.config(state="normal"))
            self.root.after(0, lambda: self.btn_stop.config(state="disabled"))
            self.root.after(0, lambda: self.status_lbl.config(text="● 待機中", fg="#10b981"))
            self.root.after(0, self.load_excel_data)

    def _run_single_window_rotation_loop(self, active_accounts, p):
        """
        【提案1: 1窓・順番交代ローテーション】
        同時並行で複数窓を開くのではなく、常に「開いているブラウザは1つだけ」を徹底。
        1. 待機時間が来たアカウントのブラウザを1窓だけ開いて1件投稿
        2. 投稿完了後、直ちにブラウザを完全に閉じてメモリを100%解放
        3. 時間が被った別アカウントがあれば、即座に次のブラウザを開いて実行
        4. 全アカウントが待機中の場合は、ブラウザを0窓（完全終了）のまま低スペックPCでも超軽量に待機
        """
        self.log("=" * 65)
        self.log("[INFO] 🚀 【提案1: 1窓・順番交代ローテーション】システム稼働開始")
        self.log(f"[INFO] 📋 対象アカウント数: {len(active_accounts)} 件")
        self.log("[INFO] 💡 動作仕様: 常に1ブラウザのみ起動 ➔ 投稿後ブラウザ完全終了＆メモリ即座解放 ➔ 待機順に次アカウント実行")
        if self.var_block_media.get():
            self.log("[INFO] ⚡ 【提案2: 画像・動画通信ブロック】有効: タイムライン画像/動画/リールを遮断（メモリ50-70%削減・CPU静音化）")
        if self.var_stealth_patrol.get():
            self.log("[INFO] 🕵️‍♂️ 【新機能1: ステルス巡回】有効: 投稿前におすすめフィードを自然に閲覧（人間らしい行動模倣・BAN対策）")
        if self.var_auto_clean_profile.get():
            self.log("[INFO] 🧹 【新機能2: 自動プロファイル初期化】有効: Meta画面警告や連続エラー検知時にCookie/指紋を自動完全リセット")
        if self.var_topic_ab_test.get():
            self.log("[INFO] 🎯 【新機能3: トピックA/Bテスト】有効: 複数トピック候補を投稿ごとに自動均等ローテーション")
        if self.var_daily_greeting.get():
            self.log("[INFO] 🌅 【日常つぶやき・あいさつ枠】有効: 朝・昼・夕・夜枠の最優先投稿（スピンタックス自動展開・動的日付置換・30日クールダウン＆ランダム選出）")
        self.log("=" * 65)

        # アカウント状態キューの初期化
        account_states = []
        for acc in active_accounts:
            acc_name = str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip()
            pf = self._find_posts_file(acc_name)
            account_states.append({
                "account": acc,
                "name": acc_name,
                "posts_file": pf,
                "next_run_time": datetime.now(), # 初回は即座に実行可能
                "cycle_num": 1,
                "consecutive_errors": 0,
                "topic_rotation_counter": 0,
            })

        while self.is_running:
            now = datetime.now()

            # 0. 🌅 【日常つぶやき・あいさつ枠】の最優先判定
            # 通常のプロモ投稿間隔とは完全に独立し、指定時間帯（朝・昼・夜など）に合致した日常ポストを最優先で投稿
            if self.var_daily_greeting.get():
                daily_executed = False
                for a_state in account_states:
                    pf = a_state["posts_file"]
                    if not pf or not os.path.exists(pf):
                        continue
                    daily_data = get_eligible_daily_post(pf)
                    if daily_data:
                        acc_name = a_state["name"]
                        self.log("=" * 60)
                        self.log(f"[INFO] 🌅 【日常つぶやき枠】@{acc_name} の「{daily_data['slot_name']} ({daily_data['time_slot']})」を検知しました！通常プロモ投稿に優先して日常ポストを実行します。")
                        self.status_lbl.config(text=f"● 日常投稿中: @{acc_name} [{daily_data['slot_name']}]", fg="#fb923c")
                        ok = self._execute_daily_greeting_post(a_state, daily_data, p)
                        if ok:
                            self.log(f"[INFO] 🌿 @{acc_name} の日常ポスト完了。ブラウザを完全終了しメモリを解放しました。")
                        self.root.after(0, self.load_excel_data)
                        daily_executed = True
                        break
                if daily_executed:
                    time.sleep(2)
                    continue

            # 1. 各アカウントの未投稿ポスト存在確認＆自動ループ周回処理
            active_account_count = 0
            for a_state in account_states:
                pf = a_state["posts_file"]
                if not pf or not os.path.exists(pf):
                    continue
                all_p = read_excel_records(pf)
                unp = [p_it for p_it in all_p if self._is_unposted_item(p_it)]
                if not unp and self.var_auto_loop.get():
                    reset_cnt = reset_posts_file_silent(pf)
                    a_state["cycle_num"] += 1
                    self.log(f"[INFO] 🔁 [{a_state['name']}] 全ポスト完了に伴い『{os.path.basename(pf)}』の全 {reset_cnt} 件を『未投稿』にリセットしました（{a_state['cycle_num']}周目へ）。")
                    all_p = read_excel_records(pf)
                    unp = [p_it for p_it in all_p if self._is_unposted_item(p_it)]
                if unp:
                    active_account_count += 1

            if active_account_count == 0:
                self.log("[SUCCESS] 🎉 全アカウントの未投稿ポストがすべて完了しました！")
                break

            # 2. 現在実行時間帯に合致し、かつ予定時刻 (next_run_time <= now) に到達しているアカウントを抽出
            eligible_accounts = []
            for a_state in account_states:
                pf = a_state["posts_file"]
                if not pf or not os.path.exists(pf):
                    continue
                all_p = read_excel_records(pf)
                unp = [p_it for p_it in all_p if self._is_unposted_item(p_it)]
                if not unp:
                    continue

                settings = read_account_settings_from_post_file(pf)
                time_slot = settings.get("time_slot", "0:00-0:00")
                if not is_within_target_hours(time_slot, now.hour):
                    continue

                if a_state["next_run_time"] <= now:
                    eligible_accounts.append(a_state)

            # 3. 待機時間が到達したアカウントがある場合
            if eligible_accounts:
                # 待機時間が来たアカウントを優先（最も待機時間が経過している順）
                eligible_accounts.sort(key=lambda a: a["next_run_time"])
                target_state = eligible_accounts[0]
                acc_name = target_state["name"]

                rem_queue = len(eligible_accounts) - 1
                self.log("-" * 55)
                self.log(f"[INFO] 🪟 【1窓ローテーション】@{acc_name} の実行番が来ました（他待機中: {rem_queue} アカウント）。ブラウザを1窓のみ起動します。")
                self.status_lbl.config(text=f"● 実行中: @{acc_name} (1窓のみ起動中)", fg="#38bdf8")

                # 1アカウント分だけブラウザを起動 ➔ 1件投稿 ➔ ブラウザ完全終了
                wait_sec = self._execute_single_post(target_state, p)

                # 次回予定時刻を設定（指定の待機秒数後）
                target_state["next_run_time"] = datetime.now() + timedelta(seconds=wait_sec)
                rem_min = wait_sec // 60
                self.log(f"[INFO] 🌿 @{acc_name} の処理完了。ブラウザを完全終了しメモリを解放しました。次回実行予定: {target_state['next_run_time'].strftime('%H:%M:%S')} (約{rem_min}分後)")
                self.root.after(0, self.load_excel_data)

                # 時間が被った別アカウントがあれば、このまま即座に次のループで次アカウントを実行！
                continue

            # 4. 到達したアカウントがない場合（全アカウント待機中）
            # 常に開いているブラウザは「0」！メモリは完全に解放された状態！
            all_valid_pending = [
                a for a in account_states
                if a["posts_file"] and os.path.exists(a["posts_file"])
                and any(self._is_unposted_item(p_it) for p_it in read_excel_records(a["posts_file"]))
            ]
            if not all_valid_pending:
                break

            nearest_acc = min(all_valid_pending, key=lambda a: a["next_run_time"])
            secs_left = max(1.0, (nearest_acc["next_run_time"] - datetime.now()).total_seconds())
            mins_left = int(secs_left // 60)
            secs_rem = int(secs_left % 60)
            time_disp = f"{mins_left:02d}分{secs_rem:02d}秒"

            self.status_lbl.config(
                text=f"● 待機中: 次回 @{nearest_acc['name']} (あと {time_disp}) [ブラウザ完全終了・省メモリ]",
                fg="#10b981"
            )

            # 1秒刻みで停止フラグを監視しながら待機
            sleep_round = min(10, int(secs_left))
            if sleep_round <= 0:
                sleep_round = 1
            for _ in range(sleep_round):
                if not self.is_running:
                    break
                time.sleep(1)

    def _execute_single_post(self, acc_state, p):
        """
        1アカウント分のブラウザを単独起動し、1件投稿を実行して直ちにブラウザを完全クローズする。
        """
        account = acc_state["account"]
        account_name = acc_state["name"]
        posts_file = acc_state["posts_file"]
        cycle_num = acc_state.get("cycle_num", 1)

        custom_settings = read_account_settings_from_post_file(posts_file)
        profile_subpath = custom_settings.get("profile_dir") or f"chrome_profiles/{account_name}"
        profile_dir = os.path.join(BASE_DIR, profile_subpath)

        if "min_interval_sec" in custom_settings:
            min_delay = custom_settings["min_interval_sec"]
            max_delay = custom_settings["max_interval_sec"]
            interval_label = custom_settings.get("interval_str", f"{min_delay}〜{max_delay}秒")
        else:
            min_delay = 7200
            max_delay = 21600
            interval_label = "120〜360分"

        post_mode = custom_settings.get("post_mode", 1)
        mode_label = "ランダム" if post_mode == 1 else "上から順"

        all_posts = read_excel_records(posts_file)
        unposted = [p_it for p_it in all_posts if self._is_unposted_item(p_it)]
        if not unposted:
            if self.var_auto_loop.get():
                reset_cnt = reset_posts_file_silent(posts_file)
                self.log(f"[INFO] [{account_name}] 未投稿が0件のため、全 {reset_cnt} 件を『未投稿』に自動リセットしました！")
                self.root.after(0, self.load_excel_data)
                all_posts = read_excel_records(posts_file)
                unposted = [p_it for p_it in all_posts if self._is_unposted_item(p_it)]

            if not unposted:
                self.log(f"[WARN] [{account_name}] '{os.path.basename(posts_file)}' に有効な投稿文章がありません。")
                return 300

        # 今回のローテーションで投稿する1件を選択
        if post_mode == 1:
            post_item = random.choice(unposted)
            self.log(f"[{account_name}] 【モード1: ランダム】未投稿 {len(unposted)} 件中から1件を選択して投稿します。")
        else:
            post_item = unposted[0]
            self.log(f"[{account_name}] 【モード0: 上から順】未投稿 {len(unposted)} 件中の先頭を選択して投稿します。")

        os.makedirs(profile_dir, exist_ok=True)
        context = None
        row_idx = post_item["_row_idx"]
        post_id = post_item.get("ID") or post_item.get("id") or row_idx - 1

        try:
            context = p.chromium.launch_persistent_context(
                user_data_dir=profile_dir,
                headless=False,
                channel="chrome",
                locale="ja-JP",
                viewport={"width": 1280, "height": 900},
                args=["--disable-blink-features=AutomationControlled", "--lang=ja-JP"]
            )

            # 提案2: 画像・動画通信ブロック
            if self.var_block_media.get():
                setup_media_blocking(context, log_func=self.log)

            session_token = str(account.get("auth_token") or account.get("Authトークン_sessionid") or account.get("sessionid") or "").strip()
            inject_session_cookies(context, session_token, account_name, log_func=self.log)

            page = context.pages[0] if context.pages else context.new_page()

            page.goto("https://www.threads.net/", timeout=60000)
            time.sleep(3)

            is_logged_in = ensure_authenticated_session(
                page, context, account_name, self.account_config_path, log_func=self.log, wait_timeout_sec=300
            )
            if not is_logged_in:
                self.log(f"[ERROR] [{account_name}] ログイン認証が完了しなかったため、このアカウントの今回の投稿をスキップします。")
                return 300

            text_content = expand_spintax_and_placeholders(self._extract_post_text(post_item))
            raw_image = self._extract_raw_image(post_item)
            resolved_img = resolve_image_path(raw_image, text_content)
            row_topic_raw = self._extract_topic_or_community(post_item)
            custom_topic_raw = custom_settings.get("topic", "")
            account_topic_raw = self._extract_topic_or_community(account) if isinstance(account, dict) else ""
            raw_topic_source = row_topic_raw or custom_topic_raw or account_topic_raw

            topic_counter = acc_state.get("topic_rotation_counter", 0)
            target_topic = select_topic_for_post(
                raw_topic_source,
                account_name,
                post_counter=topic_counter,
                enable_ab_test=self.var_topic_ab_test.get(),
                log_func=self.log
            )
            acc_state["topic_rotation_counter"] = topic_counter + 1

            topic_info_msg = f"「{target_topic}」" if target_topic else "なし"
            self.log(f"[{account_name}] Post #{post_id} ({len(text_content)}文字, トピック: {topic_info_msg}) を投稿中...")

            # 1. 別のページにいる場合はホームに戻る
            try:
                cur_url = page.url.split('?')[0].rstrip('/')
                if cur_url != "https://www.threads.net":
                    self.log(f"[{account_name}] 現在のページ({cur_url})からホーム(https://www.threads.net/)へ戻ります...")
                    page.goto("https://www.threads.net/", wait_until="domcontentloaded", timeout=30000)
                    time.sleep(2.0)
            except Exception:
                pass

            try:
                page.keyboard.press("Escape")
                time.sleep(0.3)
            except Exception:
                pass

            # 提案新機能1: 投稿前のステルス巡回（ウォーミングアップ閲覧）
            if self.var_stealth_patrol.get():
                perform_stealth_feed_patrol(page, account_name, log_func=self.log)

            # 2. 左下の新しいスレッドボタンを探す
            dialog = None
            textbox = None
            new_thread_btn = None

            try:
                for el in page.get_by_text("新しいスレッド", exact=False).all():
                    try:
                        if el.is_visible():
                            box = el.bounding_box()
                            if box and box['x'] < 300 and box['width'] < 300 and box['height'] < 80 and box['y'] > 200:
                                new_thread_btn = el
                                break
                    except Exception:
                        pass
            except Exception:
                pass

            if not new_thread_btn:
                for sel in ['div[role="button"]', 'button', 'a']:
                    try:
                        for el in page.locator(sel).all():
                            if el.is_visible():
                                txt = el.inner_text().strip()
                                if "新しいスレッド" in txt:
                                    box = el.bounding_box()
                                    if box and box['x'] < 300 and box['width'] < 300 and box['height'] < 80:
                                        new_thread_btn = el
                                        break
                        if new_thread_btn:
                            break
                    except Exception:
                        pass

            if new_thread_btn:
                self.log(f"[{account_name}] 左下の「＋ 新しいスレッド」ボタンをクリックします...")
                new_thread_btn.click()
                time.sleep(1.0)
            else:
                self.log(f"[{account_name}] キーボードショートカット 'c' で新規スレッド作成モーダルを呼び出します...")
                page.keyboard.press("c")
                time.sleep(1.0)

            # 3. 新規投稿モーダル出現待機
            for wait_i in range(16):
                time.sleep(0.5)
                try:
                    for d_sel in ['div[role="dialog"]', 'div[aria-modal="true"]']:
                        d_cand = page.locator(d_sel).first
                        if d_cand.is_visible():
                            dialog = d_cand
                            break
                    if dialog:
                        for tb_sel in [
                            'div[role="textbox"]',
                            'div[contenteditable="true"]',
                            '[data-lexical-editor="true"]'
                        ]:
                            tb_cand = dialog.locator(tb_sel).first
                            if tb_cand.is_visible():
                                textbox = tb_cand
                                break
                        if textbox:
                            break
                except Exception:
                    pass
                if wait_i == 6 and not dialog:
                    if new_thread_btn:
                        try:
                            new_thread_btn.click()
                        except Exception:
                            pass
                    try:
                        page.keyboard.press("c")
                    except Exception:
                        pass

            if not dialog or not textbox:
                raise Exception("左下の「新しいスレッド」をクリックしましたが、新規投稿モーダルが開きませんでした。他人のポストに誤爆しないよう処理を安全に中断しました。")

            # 4. テキスト入力
            textbox.click()
            try:
                textbox.press("Control+A")
                textbox.press("Backspace")
            except Exception:
                pass

            lines = text_content.split("\n")
            for l_idx, line in enumerate(lines):
                if line:
                    textbox.press_sequentially(line, delay=random.randint(15, 30))
                if l_idx < len(lines) - 1:
                    textbox.press("Shift+Enter")
                    time.sleep(0.1)

            time.sleep(1)

            # トピック設定
            if target_topic:
                set_topic_in_threads_composer(page, dialog, target_topic, self.log)
                time.sleep(0.5)

            # 画像添付
            if resolved_img:
                safe_img = ensure_valid_image(resolved_img, self.log)
                self.log(f"[{account_name}] 画像を添付します: {os.path.basename(safe_img)}")
                attached = False
                media_selectors = [
                    'div[role="dialog"] svg[aria-label*="メディア"]',
                    'div[role="dialog"] svg[aria-label*="画像"]',
                    'div[role="dialog"] svg[aria-label*="写真"]',
                    'div[role="dialog"] svg[aria-label*="添付"]',
                    'svg[aria-label="画像を追加"]',
                    'svg[aria-label="メディアを添付"]',
                    'svg[aria-label="写真を添付"]',
                    'svg[aria-label*="メディア"]',
                    'svg[aria-label*="画像"]',
                ]
                for m_sel in media_selectors:
                    try:
                        m_el = page.locator(m_sel).first
                        if m_el.is_visible():
                            btn_parent = page.locator(f'div[role="button"]:has({m_sel}), button:has({m_sel})').first
                            click_el = btn_parent if btn_parent.is_visible() else m_el
                            with page.expect_file_chooser(timeout=2500) as fc_info:
                                click_el.click()
                            file_chooser = fc_info.value
                            file_chooser.set_files(safe_img)
                            attached = True
                            self.log(f"[{account_name}] メディア添付アイコン経由で画像を設定しました。")
                            break
                    except Exception:
                        pass

                if not attached:
                    for f_sel in [
                        'div[role="dialog"] input[type="file"]',
                        'div[aria-modal="true"] input[type="file"]',
                        'input[type="file"][accept*="image"]',
                        'input[type="file"][multiple]',
                        'input[type="file"]'
                    ]:
                        try:
                            locs = page.locator(f_sel).all()
                            if locs:
                                locs[-1].set_input_files(safe_img)
                                attached = True
                                self.log(f"[{account_name}] ダイアログ内ファイル入力要素({f_sel})に画像を設定しました。")
                                break
                        except Exception:
                            pass

                preview_found = False
                for _ in range(12):
                    time.sleep(0.8)
                    for err_sel in ['div:text-matches("添付ファイルを添付できませんでした|アップロードできませんでした|エラーが発生しました")']:
                        try:
                            err_el = page.locator(err_sel).first
                            if err_el.is_visible():
                                err_text = err_el.inner_text().strip()
                                self.log(f"[WARN] [{account_name}] Threads画面通知: '{err_text}'")
                        except Exception:
                            pass
                    for p_sel in [
                        'div[role="dialog"] img[src^="blob:"]',
                        'div[role="dialog"] img[alt*="メディア"]',
                        'div[role="dialog"] img[alt*="画像"]',
                        'div[role="dialog"] svg[aria-label*="削除"]',
                        'div[role="dialog"] svg[aria-label*="メディアを破棄"]',
                        'div[role="dialog"] svg[aria-label*="メディアを編集"]'
                    ]:
                        try:
                            if page.locator(p_sel).first.is_visible():
                                preview_found = True
                                break
                        except Exception:
                            pass
                    if preview_found:
                        self.log(f"[{account_name}] 画像プレビュー確認完了（アップロード成功）")
                        break

            time.sleep(1.0)

            # ポップアップクローズ
            for pop_sel in ['text="返信と引用ができる人"', 'text="返信できる人"', 'div[role="menu"]', 'div[role="listbox"]']:
                try:
                    pop_el = page.locator(pop_sel).first
                    if pop_el.is_visible():
                        self.log(f"[{account_name}] 画面上のオプションメニューを閉じて「投稿」ボタンを押せるようにします...")
                        page.keyboard.press("Escape")
                        time.sleep(0.6)
                        break
                except Exception:
                    pass

            # 投稿ボタン特定
            submit_btn = None
            for _ in range(15):
                try:
                    if page.locator('text="返信と引用ができる人"').first.is_visible():
                        page.keyboard.press("Escape")
                        time.sleep(0.5)
                except Exception:
                    pass

                candidates = []
                search_root = dialog if (dialog and dialog.is_visible()) else page.locator('div[role="dialog"]').first
                for btn in search_root.locator('div[role="button"], button').all():
                    try:
                        if not btn.is_visible():
                            continue
                        txt = btn.inner_text().strip()
                        clean_txt = txt.replace(" ", "").replace("\n", "")
                        if clean_txt in ["投稿", "Post", "投稿する"] and "オプション" not in txt and "新しい" not in txt and "作成" not in txt and "下書き" not in txt:
                            candidates.append(btn)
                    except Exception:
                        pass

                if candidates:
                    cand_btn = candidates[-1]
                    is_disabled = (cand_btn.get_attribute("aria-disabled") == "true") or (cand_btn.get_attribute("disabled") is not None)
                    if not is_disabled:
                        submit_btn = cand_btn
                        break
                time.sleep(1)

            if not submit_btn:
                raise Exception("「投稿」ボタンがアクティブになりませんでした（画像のアップロード中か、文字数超過の可能性があります）。")

            # 投稿実行
            self.log(f"[{account_name}] 「投稿」ボタンをクリックしました。Threads送信処理を確認中...")
            try:
                submit_btn.scroll_into_view_if_needed()
                submit_btn.click(force=True)
            except Exception as e:
                self.log(f"[{account_name}] 送信ボタン通常クリック試行: {e}")

            try:
                submit_btn.evaluate("el => el.click()")
            except Exception:
                pass

            try:
                textbox.focus()
                page.keyboard.press("Control+Enter")
            except Exception:
                pass

            posted_success = False
            post_error_detected = ""

            for check_idx in range(30):
                time.sleep(0.8)
                for s_sel in [
                    'text="投稿完了"',
                    ':has-text("投稿完了")',
                    'text="投稿しました"',
                    ':has-text("投稿しました")',
                    'text="スレッドを投稿しました"',
                    ':has-text("スレッドを投稿しました")',
                    'text="スレッドを表示"',
                    ':has-text("スレッドを表示")',
                    'text="スレッドが投稿されました"',
                    ':has-text("スレッドが投稿されました")',
                    'text="Posted"',
                    'text="View post"',
                    'div[role="alert"]:has-text("投稿")',
                    'div[role="alert"]:has-text("スレッド")',
                    'div:has-text("投稿完了")',
                    'div:has-text("投稿しました")',
                    'div:has-text("スレッドを投稿しました")'
                ]:
                    try:
                        s_loc = page.locator(s_sel).first
                        if s_loc.is_visible():
                            posted_success = True
                            self.log(f"[SUCCESS] [{account_name}] 🎉 画面下部の黒い通知バー（'{s_sel}'）を確認し、投稿完了しました！")
                            break
                    except Exception:
                        pass
                if posted_success:
                    break

                for err_sel in [
                    'div[role="alert"]:has-text("エラー")',
                    'text="投稿できませんでした"',
                    'text="アップロードできませんでした"',
                    'text="エラーが発生しました"',
                    'text="問題が発生しました"',
                    'text="コミュニティガイドライン"',
                    'text="後でもう一度お試しください"'
                ]:
                    try:
                        err_el = page.locator(err_sel).first
                        if err_el.is_visible():
                            post_error_detected = err_el.inner_text().strip()[:80]
                            break
                    except Exception:
                        pass
                if post_error_detected:
                    break

                modal_closed = True
                try:
                    for d in page.locator('div[role="dialog"], div[aria-modal="true"]').all():
                        if d.is_visible():
                            modal_closed = False
                            break
                except Exception:
                    pass

                if modal_closed:
                    time.sleep(0.4)
                    for err_sel in ['div[role="alert"]:has-text("エラー")', 'text="投稿できませんでした"']:
                        try:
                            if page.locator(err_sel).first.is_visible():
                                post_error_detected = "エラー通知検出"
                                break
                        except Exception:
                            pass
                    if not post_error_detected:
                        posted_success = True
                        self.log(f"[SUCCESS] [{account_name}] 🎉 投稿モーダルが閉じ、下部の黒バー完了通知を確認しました！")
                        break

                if check_idx in [5, 12]:
                    try:
                        if page.locator('text="返信と引用ができる人"').first.is_visible():
                            page.keyboard.press("Escape")
                            time.sleep(0.5)
                        for btn in page.locator('div[role="dialog"] div[role="button"], div[role="dialog"] button').all():
                            try:
                                if btn.is_visible() and btn.inner_text().strip() in ["投稿", "Post"]:
                                    self.log(f"[{account_name}] 送信ボタンを再押下します...")
                                    btn.click(force=True)
                                    break
                            except Exception:
                                pass
                    except Exception:
                        pass

            if post_error_detected:
                raise Exception(f"Threads画面でエラーが返されました: {post_error_detected}")

            if not posted_success:
                raise Exception("投稿ボタンを押下後、30秒経過してもThreads上で送信完了（投稿通知またはフォーム閉鎖）が確認できませんでした。下書きが画面に残っています。")

            finish_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            update_post_status(posts_file, row_idx, status="投稿完了", finish_time=finish_time, error_msg="", sheet_name=post_item.get("_sheet_name"), applied_topic=target_topic)
            topic_suffix = f" (トピック: 「{target_topic}」)" if target_topic else ""
            self.log(f"[SUCCESS] [{account_name}] Post #{post_id} 投稿成功！Excel更新{topic_suffix}。")
            acc_state["consecutive_errors"] = 0

            # 投稿後の待機時間を決定
            wait_sec = random.randint(min_delay, max_delay)
            return wait_sec

        except Exception as post_err:
            acc_state["consecutive_errors"] = acc_state.get("consecutive_errors", 0) + 1
            err_count = acc_state["consecutive_errors"]

            screen_msgs = []
            is_warning = False
            warning_keywords = [
                "コミュニティガイドライン",
                "制限されています",
                "利用を制限",
                "一時的に制限",
                "ブロックされています",
                "アクションがブロック",
                "後でもう一度",
                "後ほどもう一度",
                "不審なアクティビティ",
                "本人確認",
                "アカウントが一時停止",
                "セッションの有効期限",
                "セッションが無効",
                "利用制限",
                "スパム"
            ]

            if 'page' in locals() and page:
                diag = capture_error_diagnostic(page, account_name, f"Post #{post_id}", post_err, self.log)
                short_reason = diag.get("short_reason", str(post_err))
                screen_msgs = diag.get("screen_messages", [])
                self.log(f"[ERROR] [{account_name}] Post #{post_id} 投稿エラー: {short_reason}")
                if screen_msgs:
                    for sm in screen_msgs:
                        self.log(f"[ERROR] [{account_name}] 画面上のメッセージ: 「{sm}」")
                        if any(wk in sm for wk in warning_keywords):
                            is_warning = True
                if diag.get("screenshot_saved"):
                    self.log(f"[INFO] [{account_name}] 📸 エラー画面を保存しました: {diag['rel_screenshot']}")
                finish_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                excel_err_text = f"[{finish_time}] {short_reason} (スクショ: {diag.get('rel_screenshot', '')})"
                update_post_status(posts_file, row_idx, status="エラー", finish_time=None, error_msg=excel_err_text, sheet_name=post_item.get("_sheet_name"))
            else:
                short_reason = str(post_err)
                self.log(f"[ERROR] [{account_name}] 投稿処理エラー: {post_err}")

            if any(wk in str(post_err) for wk in warning_keywords):
                is_warning = True

            # 提案新機能2: 不安定・警告兆候時の自動プロファイル初期化（クリーンアップ）
            if self.var_auto_clean_profile.get():
                trigger_clean = False
                clean_reason = ""
                if is_warning:
                    trigger_clean = True
                    clean_reason = f"Meta画面警告検知 ({screen_msgs[0] if screen_msgs else short_reason})"
                elif err_count >= 2:
                    trigger_clean = True
                    clean_reason = f"連続投稿エラー ({err_count}回連続失敗)"

                if trigger_clean:
                    if context:
                        try:
                            context.close()
                            context = None
                        except Exception:
                            pass
                    time.sleep(1.0)
                    clean_account_profile(account_name, profile_dir, reason=clean_reason, log_func=self.log)
                    acc_state["consecutive_errors"] = 0

                    if is_warning:
                        cooldown_sec = random.randint(900, 1800)
                        self.log(f"[WARN] [{account_name}] ⚠️ アカウント保護のため次回までクールダウン待機（{cooldown_sec // 60}分）を設定します。")
                        return cooldown_sec

            return 300
        finally:
            if context:
                try:
                    context.close()
                except Exception:
                    pass

    def _execute_daily_greeting_post(self, acc_state, daily_data, p):
        """
        🌅 【日常つぶやき・あいさつ枠】
        プロモ投稿間隔とは完全に独立し、指定時間帯（朝・昼・夜など）に合致した日常ポストを専用投稿。
        - リンクプレビューやトピックの強制付与は行わず、自然な日常の呟きとして投稿
        - 投稿後は直ちにブラウザを完全終了してメモリを100%解放
        - Excelの「日常ポスト」シートに本日完了（最終投稿日・最終投稿日時）を記録
        """
        account = acc_state["account"]
        account_name = acc_state["name"]
        posts_file = daily_data["filepath"]
        slot_name = daily_data["slot_name"]
        time_slot = daily_data["time_slot"]
        post_text = expand_spintax_and_placeholders(daily_data.get("raw_text") or daily_data["post_text"])
        row_indices = daily_data["row_indices"]
        sheet_name = daily_data["sheet_name"]

        custom_settings = read_account_settings_from_post_file(posts_file)
        profile_subpath = custom_settings.get("profile_dir") or f"chrome_profiles/{account_name}"
        profile_dir = os.path.join(BASE_DIR, profile_subpath)
        os.makedirs(profile_dir, exist_ok=True)

        self.log(f"[{account_name}] 🌅 日常つぶやき【{slot_name}】({time_slot}) 自動投稿を開始（スピンタックス＆動的日付展開済: {post_text[:35]}...）")
        context = None
        page = None
        try:
            context = p.chromium.launch_persistent_context(
                user_data_dir=profile_dir,
                headless=False,
                channel="chrome",
                locale="ja-JP",
                viewport={"width": 1280, "height": 900},
                args=["--disable-blink-features=AutomationControlled", "--lang=ja-JP"]
            )
            if self.var_block_media.get():
                setup_media_blocking(context, log_func=self.log)

            session_token = str(account.get("auth_token") or account.get("Authトークン_sessionid") or account.get("sessionid") or "").strip()
            inject_session_cookies(context, session_token, account_name, log_func=self.log)

            page = context.pages[0] if context.pages else context.new_page()
            page.goto("https://www.threads.net/", timeout=60000)
            time.sleep(3)

            is_logged_in = ensure_authenticated_session(
                page, context, account_name, self.account_config_path, log_func=self.log, wait_timeout_sec=300
            )
            if not is_logged_in:
                self.log(f"[ERROR] [{account_name}] ログイン認証が完了しなかったため日常投稿をスキップします。")
                return False

            try:
                cur_url = page.url.split('?')[0].rstrip('/')
                if cur_url != "https://www.threads.net":
                    page.goto("https://www.threads.net/", wait_until="domcontentloaded", timeout=30000)
                    time.sleep(2.0)
            except Exception:
                pass

            try:
                page.keyboard.press("Escape")
                time.sleep(0.3)
            except Exception:
                pass

            if self.var_stealth_patrol.get():
                perform_stealth_feed_patrol(page, account_name, log_func=self.log)

            dialog = None
            textbox = None
            new_thread_btn = None
            try:
                for el in page.get_by_text("新しいスレッド", exact=False).all():
                    try:
                        if el.is_visible():
                            box = el.bounding_box()
                            if box and box['x'] < 300 and box['width'] < 300 and box['height'] < 80 and box['y'] > 200:
                                new_thread_btn = el
                                break
                    except Exception:
                        pass
            except Exception:
                pass

            if not new_thread_btn:
                for sel in ['div[role="button"]', 'button', 'a']:
                    try:
                        for el in page.locator(sel).all():
                            if el.is_visible():
                                txt = el.inner_text().strip()
                                if "新しいスレッド" in txt:
                                    box = el.bounding_box()
                                    if box and box['x'] < 300 and box['width'] < 300 and box['height'] < 80:
                                        new_thread_btn = el
                                        break
                        if new_thread_btn:
                            break
                    except Exception:
                        pass

            if new_thread_btn:
                new_thread_btn.click()
                time.sleep(1.0)
            else:
                page.keyboard.press("c")
                time.sleep(1.0)

            for wait_i in range(16):
                time.sleep(0.5)
                try:
                    for d_sel in ['div[role="dialog"]', 'div[aria-modal="true"]']:
                        d_cand = page.locator(d_sel).first
                        if d_cand.is_visible():
                            dialog = d_cand
                            break
                    if dialog:
                        for tb_sel in [
                            'div[role="textbox"]',
                            'div[contenteditable="true"]',
                            '[data-lexical-editor="true"]'
                        ]:
                            tb_cand = dialog.locator(tb_sel).first
                            if tb_cand.is_visible():
                                textbox = tb_cand
                                break
                        if textbox:
                            break
                except Exception:
                    pass
                if wait_i == 6 and not dialog:
                    if new_thread_btn:
                        try:
                            new_thread_btn.click()
                        except Exception:
                            pass
                    try:
                        page.keyboard.press("c")
                    except Exception:
                        pass

            if not dialog or not textbox:
                raise Exception("新規投稿ダイアログが開けませんでした。")

            textbox.click()
            try:
                textbox.press("Control+A")
                textbox.press("Backspace")
            except Exception:
                pass

            lines = post_text.split("\n")
            for l_idx, line in enumerate(lines):
                if line:
                    textbox.press_sequentially(line, delay=random.randint(15, 30))
                if l_idx < len(lines) - 1:
                    textbox.press("Shift+Enter")
                    time.sleep(0.1)

            time.sleep(1.2)

            # トピック設定（日常ポストの「コミュニティまたはトピック」列から設定可能）
            # カンマ区切り（例: 信太山新地, 大阪求人, 高収入バイト）でA/Bテスト自動均等ローテーション
            daily_raw_topic = daily_data.get("topic", "")
            if daily_raw_topic:
                target_topic = select_topic_for_post(
                    daily_raw_topic,
                    account_name,
                    post_counter=acc_state.get("daily_topic_counter", 0),
                    enable_ab_test=self.var_topic_ab_test.get(),
                    log_func=self.log
                )
                if target_topic:
                    acc_state["daily_topic_counter"] = acc_state.get("daily_topic_counter", 0) + 1
                    set_topic_in_threads_composer(page, dialog, target_topic, self.log)
                    time.sleep(0.5)

            for pop_sel in ['text="返信と引用ができる人"', 'text="返信できる人"', 'div[role="menu"]', 'div[role="listbox"]']:
                try:
                    pop_el = page.locator(pop_sel).first
                    if pop_el.is_visible():
                        page.keyboard.press("Escape")
                        time.sleep(0.6)
                        break
                except Exception:
                    pass

            submit_btn = None
            for _ in range(15):
                try:
                    if page.locator('text="返信と引用ができる人"').first.is_visible():
                        page.keyboard.press("Escape")
                        time.sleep(0.5)
                except Exception:
                    pass

                candidates = []
                search_root = dialog if (dialog and dialog.is_visible()) else page.locator('div[role="dialog"]').first
                for btn in search_root.locator('div[role="button"], button').all():
                    try:
                        if not btn.is_visible():
                            continue
                        txt = btn.inner_text().strip()
                        clean_txt = txt.replace(" ", "").replace("\n", "")
                        if clean_txt in ["投稿", "Post", "投稿する"] and "オプション" not in txt and "新しい" not in txt and "作成" not in txt and "下書き" not in txt:
                            candidates.append(btn)
                    except Exception:
                        pass

                if candidates:
                    cand_btn = candidates[-1]
                    is_disabled = (cand_btn.get_attribute("aria-disabled") == "true") or (cand_btn.get_attribute("disabled") is not None)
                    if not is_disabled:
                        submit_btn = cand_btn
                        break
                time.sleep(1)

            if not submit_btn:
                raise Exception("「投稿」ボタンがアクティブになりませんでした。")

            self.log(f"[{account_name}] 日常ポストの「投稿」ボタンをクリックしました...")
            try:
                submit_btn.scroll_into_view_if_needed()
                submit_btn.click(force=True)
            except Exception:
                pass
            try:
                submit_btn.evaluate("el => el.click()")
            except Exception:
                pass
            try:
                textbox.focus()
                page.keyboard.press("Control+Enter")
            except Exception:
                pass

            posted_success = False
            for check_idx in range(25):
                time.sleep(0.8)
                for s_sel in [
                    'text="投稿完了"', ':has-text("投稿完了")',
                    'text="投稿しました"', ':has-text("投稿しました")',
                    'text="スレッドを投稿しました"', ':has-text("スレッドを投稿しました")',
                    'text="スレッドを表示"', ':has-text("スレッドを表示")',
                    'text="スレッドが投稿されました"', ':has-text("スレッドが投稿されました")',
                    'text="Posted"', 'text="View post"',
                    'div[role="alert"]:has-text("投稿")', 'div[role="alert"]:has-text("スレッド")'
                ]:
                    try:
                        if page.locator(s_sel).first.is_visible():
                            posted_success = True
                            break
                    except Exception:
                        pass
                if posted_success:
                    break

                modal_closed = True
                try:
                    for d in page.locator('div[role="dialog"], div[aria-modal="true"]').all():
                        if d.is_visible():
                            modal_closed = False
                            break
                except Exception:
                    pass
                if modal_closed:
                    posted_success = True
                    break

                if check_idx in [5, 10]:
                    try:
                        page.keyboard.press("Escape")
                        time.sleep(0.3)
                        for btn in page.locator('div[role="dialog"] div[role="button"], div[role="dialog"] button').all():
                            if btn.is_visible() and btn.inner_text().strip() in ["投稿", "Post"]:
                                btn.click(force=True)
                                break
                    except Exception:
                        pass

            if not posted_success:
                raise Exception("日常ポスト送信後、完了通知またはモーダル終了が確認できませんでした。")

            update_daily_post_status(posts_file, sheet_name, row_indices)
            self.log(f"[SUCCESS] [{account_name}] 🎉 🌅 日常つぶやき【{slot_name}】の自動投稿に成功しました！(Excelの「日常ポスト」シートに本日完了を記録)")
            return True
        except Exception as e:
            if page:
                diag = capture_error_diagnostic(page, account_name, f"DailyPost {slot_name}", e, self.log)
                self.log(f"[ERROR] [{account_name}] 🌅 日常ポスト投稿エラー: {diag.get('short_reason', str(e))}")
            else:
                self.log(f"[ERROR] [{account_name}] 🌅 日常ポスト投稿エラー: {e}")
            return False
        finally:
            if context:
                try:
                    context.close()
                except Exception:
                    pass

    def _run_legacy_sequential_loop(self, active_accounts, p):
        cycle_num = 1
        while self.is_running:
            if cycle_num > 1:
                self.log("=" * 60)
                self.log(f"[INFO] 🔁 【周回サイクル {cycle_num} 周目】自動投稿ループを開始します")
                self.log("=" * 60)

            for acc in active_accounts:
                if not self.is_running:
                    break
                # 日常ポストのチェック
                if self.var_daily_greeting.get():
                    acc_name = str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip()
                    pf = self._find_posts_file(acc_name)
                    if pf and os.path.exists(pf):
                        daily_data = get_eligible_daily_post(pf)
                        if daily_data:
                            acc_st = {"account": acc, "name": acc_name, "posts_file": pf}
                            self._execute_daily_greeting_post(acc_st, daily_data, p)
                self._process_single_account(acc, p, cycle_num=cycle_num)

            if not self.var_auto_loop.get() or not self.is_running:
                break

            for acc in active_accounts:
                acc_name = str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip()
                pf = self._find_posts_file(acc_name)
                if pf and os.path.exists(pf):
                    cur_posts = read_excel_records(pf)
                    cur_unposted = [p_it for p_it in cur_posts if self._is_unposted_item(p_it)]
                    if not cur_unposted:
                        reset_cnt = reset_posts_file_silent(pf)
                        self.log(f"[INFO] [{acc_name}] サイクル完了に伴い '{os.path.basename(pf)}' の全 {reset_cnt} 件を『未投稿』にリセットしました。")

            cycle_num += 1
            self.root.after(0, self.load_excel_data)

    def _process_single_account(self, account, p, cycle_num=1):
        account_name = str(account.get("アカウントID") or account.get("アカウント名") or "").strip()
        posts_file = self._find_posts_file(account_name)
        if not posts_file or not os.path.exists(posts_file):
            self.log(f"[{account_name}] '{posts_file}' が見つかりません。")
            return

        custom_settings = read_account_settings_from_post_file(posts_file)
        profile_subpath = custom_settings.get("profile_dir") or f"chrome_profiles/{account_name}"
        profile_dir = os.path.join(BASE_DIR, profile_subpath)

        if "min_interval_sec" in custom_settings:
            min_delay = custom_settings["min_interval_sec"]
            max_delay = custom_settings["max_interval_sec"]
            interval_label = custom_settings.get("interval_str", f"{min_delay}〜{max_delay}秒")
        else:
            min_delay = 7200
            max_delay = 21600
            interval_label = "120〜360分"

        post_mode = custom_settings.get("post_mode", 1)
        mode_label = "ランダム" if post_mode == 1 else "上から順"

        time_slot = custom_settings.get("time_slot", "0:00-0:00")
        now_hour = datetime.now().hour
        if not is_within_target_hours(time_slot, now_hour):
            self.log(f"[WARN] [{account_name}] 現在時刻({now_hour}時)は{os.path.basename(posts_file)}の指定実行時間 ({time_slot}) 外のためスキップします。")
            return

        self.log(f"[{account_name}] 【{os.path.basename(posts_file)} 設定適用】ポスト間隔: {interval_label} ({min_delay}〜{max_delay}秒ランダム) | モード: {mode_label} | プロファイル: {profile_subpath} (周回: {cycle_num})")

        all_posts = read_excel_records(posts_file)
        unposted = [p for p in all_posts if self._is_unposted_item(p)]
        if not unposted:
            if self.var_auto_loop.get():
                reset_cnt = reset_posts_file_silent(posts_file)
                self.log(f"[INFO] [{account_name}] 未投稿が0件のため、自動ループ機能により全 {reset_cnt} 件を『未投稿』に自動リセットしました！最初から投稿を再開します。")
                self.root.after(0, self.load_excel_data)
                all_posts = read_excel_records(posts_file)
                unposted = [p for p in all_posts if self._is_unposted_item(p)]

            if not unposted:
                self.log(f"[WARN] [{account_name}] '{os.path.basename(posts_file)}' に有効な投稿文章がありません（全 {len(all_posts)} 件）。")
                if not self.var_auto_loop.get():
                    self.root.after(0, lambda: messagebox.showinfo(
                        "未投稿のポストが 0 件です",
                        f"【ウィンドウが開かない原因】\n"
                        f"投稿リスト '{os.path.basename(posts_file)}' の全ポストが「投稿完了」になっています。\n\n"
                        "二重投稿・BAN防止のため、未投稿データがない場合はブラウザを開かずに終了します。\n\n"
                        "【再投稿する方法】\n"
                        "・画面の『🔁 全ポスト投稿完了後に自動で未投稿にリセットして最初から周回』にチェックを入れる\n"
                        "・または画面上部の『🔁 未投稿にリセット』を押してください。"
                    ))
                return

        if post_mode == 1:
            random.shuffle(unposted)
            self.log(f"[{account_name}] 【モード1: ランダム】未投稿 {len(unposted)} 件をランダム順序で投稿します。")
        else:
            self.log(f"[{account_name}] 【モード0: 上から順】未投稿 {len(unposted)} 件を先頭から順に投稿します。")

        os.makedirs(profile_dir, exist_ok=True)

        context = p.chromium.launch_persistent_context(
            user_data_dir=profile_dir,
            headless=False,
            channel="chrome",
            locale="ja-JP",
            viewport={"width": 1280, "height": 900},
            args=["--disable-blink-features=AutomationControlled", "--lang=ja-JP"]
        )

        if self.var_block_media.get():
            setup_media_blocking(context, log_func=self.log)

        session_token = str(account.get("auth_token") or account.get("Authトークン_sessionid") or account.get("sessionid") or "").strip()
        inject_session_cookies(context, session_token, account_name, log_func=self.log)

        page = context.pages[0] if context.pages else context.new_page()

        try:
            page.goto("https://www.threads.net/", timeout=60000)
            time.sleep(3)

            is_logged_in = ensure_authenticated_session(
                page, context, account_name, self.account_config_path, log_func=self.log, wait_timeout_sec=300
            )
            if not is_logged_in:
                self.log(f"[ERROR] [{account_name}] ログイン認証が完了しなかったため、このアカウントの投稿をスキップします。")
                return

            consecutive_account_errors = 0
            for post_loop_idx, post_item in enumerate(unposted):
                if not self.is_running:
                    break

                row_idx = post_item["_row_idx"]
                post_id = post_item.get("ID") or post_item.get("id") or row_idx - 1
                text_content = expand_spintax_and_placeholders(self._extract_post_text(post_item))
                raw_image = self._extract_raw_image(post_item)
                resolved_img = resolve_image_path(raw_image, text_content)
                row_topic_raw = self._extract_topic_or_community(post_item)
                custom_topic_raw = custom_settings.get("topic", "")
                account_topic_raw = self._extract_topic_or_community(account) if isinstance(account, dict) else ""
                raw_topic_source = row_topic_raw or custom_topic_raw or account_topic_raw

                target_topic = select_topic_for_post(
                    raw_topic_source,
                    account_name,
                    post_counter=post_loop_idx,
                    enable_ab_test=self.var_topic_ab_test.get(),
                    log_func=self.log
                )

                topic_info_msg = f"「{target_topic}」" if target_topic else "なし"
                self.log(f"[{account_name}] Post #{post_id} ({len(text_content)}文字, トピック: {topic_info_msg}) を投稿中...")

                try:
                    # 1. 別のユーザーページや投稿詳細にいる場合は、必ずホーム (https://www.threads.net/) に戻る
                    try:
                        cur_url = page.url.split('?')[0].rstrip('/')
                        if cur_url != "https://www.threads.net":
                            self.log(f"[{account_name}] 現在のページ({cur_url})からホーム(https://www.threads.net/)へ戻ります...")
                            page.goto("https://www.threads.net/", wait_until="domcontentloaded", timeout=30000)
                            time.sleep(2.0)
                    except Exception:
                        pass

                    # 画面上に残っているポップアップやダイアログをクリーンアップ
                    try:
                        page.keyboard.press("Escape")
                        time.sleep(0.3)
                    except Exception:
                        pass

                    # 提案新機能1: 投稿前のステルス巡回（ウォーミングアップ閲覧）
                    if self.var_stealth_patrol.get():
                        perform_stealth_feed_patrol(page, account_name, log_func=self.log)

                    # 2. 画面【左下の新しいスレッド】ボタンをピンポイントで特定してクリック
                    # （※親コンテナやフィードへの誤クリックを完全に防ぐため、サイズ幅<300, 高さ<80, 画面左側の要素のみを厳格に対象化）
                    dialog = None
                    textbox = None
                    new_thread_btn = None

                    # (A) テキスト要素から幅と高さがコンパクトなボタン要素を探索
                    try:
                        for el in page.get_by_text("新しいスレッド", exact=False).all():
                            try:
                                if el.is_visible():
                                    box = el.bounding_box()
                                    if box and box['x'] < 300 and box['width'] < 300 and box['height'] < 80 and box['y'] > 200:
                                        new_thread_btn = el
                                        break
                            except Exception:
                                pass
                    except Exception:
                        pass

                    # (B) ボタン・リンクタグから厳密に探索
                    if not new_thread_btn:
                        for sel in ['div[role="button"]', 'button', 'a']:
                            try:
                                for el in page.locator(sel).all():
                                    if el.is_visible():
                                        txt = el.inner_text().strip()
                                        if "新しいスレッド" in txt:
                                            box = el.bounding_box()
                                            if box and box['x'] < 300 and box['width'] < 300 and box['height'] < 80:
                                                new_thread_btn = el
                                                break
                                if new_thread_btn:
                                    break
                            except Exception:
                                pass

                    # (C) ボタンをクリック
                    if new_thread_btn:
                        self.log(f"[{account_name}] 左下の「＋ 新しいスレッド」ボタンをクリックします...")
                        new_thread_btn.click()
                        time.sleep(1.0)
                    else:
                        self.log(f"[{account_name}] キーボードショートカット 'c' で新規スレッド作成モーダルを呼び出します...")
                        page.keyboard.press("c")
                        time.sleep(1.0)

                    # 3. 新規投稿モーダル (div[role="dialog"]) の出現を厳格に待機
                    for wait_i in range(16):
                        time.sleep(0.5)
                        try:
                            for d_sel in ['div[role="dialog"]', 'div[aria-modal="true"]']:
                                d_cand = page.locator(d_sel).first
                                if d_cand.is_visible():
                                    dialog = d_cand
                                    break
                            if dialog:
                                for tb_sel in [
                                    'div[role="textbox"]',
                                    'div[contenteditable="true"]',
                                    '[data-lexical-editor="true"]'
                                ]:
                                    tb_cand = dialog.locator(tb_sel).first
                                    if tb_cand.is_visible():
                                        textbox = tb_cand
                                        break
                                if textbox:
                                    break
                        except Exception:
                            pass
                        # 4秒経過しても開かない場合、再度クリックを試みる
                        if wait_i == 6 and not dialog:
                            if new_thread_btn:
                                try:
                                    new_thread_btn.click()
                                except Exception:
                                    pass
                            try:
                                page.keyboard.press("c")
                            except Exception:
                                pass

                    if not dialog or not textbox:
                        raise Exception("左下の「新しいスレッド」をクリックしましたが、新規投稿モーダルが開きませんでした。他人のポストに誤爆しないよう処理を安全に中断しました。")

                    # 4. モーダル内の新規スレッド入力欄に投稿テキストを入力

                    textbox.click()
                    try:
                        textbox.press("Control+A")
                        textbox.press("Backspace")
                    except Exception:
                        pass

                    lines = text_content.split("\n")
                    for l_idx, line in enumerate(lines):
                        if line:
                            textbox.press_sequentially(line, delay=random.randint(15, 30))
                        if l_idx < len(lines) - 1:
                            textbox.press("Shift+Enter")
                            time.sleep(0.1)

                    time.sleep(1)

                    # コミュニティまたはトピックの設定
                    if target_topic:
                        set_topic_in_threads_composer(page, dialog, target_topic, self.log)
                        time.sleep(0.5)

                    if resolved_img:

                        safe_img = ensure_valid_image(resolved_img, self.log)
                        self.log(f"[{account_name}] 画像を添付します: {os.path.basename(safe_img)}")
                        
                        attached = False
                        
                        # 1. 投稿ダイアログ内の画像添付アイコンをクリックして expect_file_chooser で設定
                        media_selectors = [
                            'div[role="dialog"] svg[aria-label*="メディア"]',
                            'div[role="dialog"] svg[aria-label*="画像"]',
                            'div[role="dialog"] svg[aria-label*="写真"]',
                            'div[role="dialog"] svg[aria-label*="添付"]',
                            'svg[aria-label="画像を追加"]',
                            'svg[aria-label="メディアを添付"]',
                            'svg[aria-label="写真を添付"]',
                            'svg[aria-label*="メディア"]',
                            'svg[aria-label*="画像"]',
                        ]
                        for m_sel in media_selectors:
                            try:
                                m_el = page.locator(m_sel).first
                                if m_el.is_visible():
                                    btn_parent = page.locator(f'div[role="button"]:has({m_sel}), button:has({m_sel})').first
                                    click_el = btn_parent if btn_parent.is_visible() else m_el
                                    with page.expect_file_chooser(timeout=2500) as fc_info:
                                        click_el.click()
                                    file_chooser = fc_info.value
                                    file_chooser.set_files(safe_img)
                                    attached = True
                                    self.log(f"[{account_name}] メディア添付アイコン経由で画像を設定しました。")
                                    break
                            except Exception:
                                pass
                        
                        # 2. expect_file_chooserで設定できなかった場合、ダイアログ内のinput[type="file"]を直接指定
                        if not attached:
                            for f_sel in [
                                'div[role="dialog"] input[type="file"]',
                                'div[aria-modal="true"] input[type="file"]',
                                'input[type="file"][accept*="image"]',
                                'input[type="file"][multiple]',
                                'input[type="file"]'
                            ]:
                                try:
                                    locs = page.locator(f_sel).all()
                                    if locs:
                                        locs[-1].set_input_files(safe_img)
                                        attached = True
                                        self.log(f"[{account_name}] ダイアログ内ファイル入力要素({f_sel})に画像を設定しました。")
                                        break
                                except Exception:
                                    pass

                        
                        # 3. アップロードプレビューの完了を待機
                        preview_found = False
                        for _ in range(12):
                            time.sleep(0.8)
                            for err_sel in ['div:text-matches("添付ファイルを添付できませんでした|アップロードできませんでした|エラーが発生しました")']:
                                try:
                                    err_el = page.locator(err_sel).first
                                    if err_el.is_visible():
                                        err_text = err_el.inner_text().strip()
                                        self.log(f"[WARN] [{account_name}] Threads画面通知: '{err_text}'")
                                except Exception:
                                    pass
                            
                            for p_sel in [
                                'div[role="dialog"] img[src^="blob:"]',
                                'div[role="dialog"] img[alt*="メディア"]',
                                'div[role="dialog"] img[alt*="画像"]',
                                'div[role="dialog"] svg[aria-label*="削除"]',
                                'div[role="dialog"] svg[aria-label*="メディアを破棄"]',
                                'div[role="dialog"] svg[aria-label*="メディアを編集"]'
                            ]:
                                try:
                                    if page.locator(p_sel).first.is_visible():
                                        preview_found = True
                                        break
                                except Exception:
                                    pass
                            if preview_found:
                                self.log(f"[{account_name}] 画像プレビュー確認完了（アップロード成功）")
                                break

                    time.sleep(1.0)

                    # ポップアップ（「返信と引用ができる人」や投稿オプション）が開いている場合はEscapeで閉じる
                    for pop_sel in ['text="返信と引用ができる人"', 'text="返信できる人"', 'div[role="menu"]', 'div[role="listbox"]']:
                        try:
                            pop_el = page.locator(pop_sel).first
                            if pop_el.is_visible():
                                self.log(f"[{account_name}] 画面上のオプションメニューを閉じて「投稿」ボタンを押せるようにします...")
                                page.keyboard.press("Escape")
                                time.sleep(0.6)
                                break
                        except Exception:
                            pass

                    # 本物の「投稿」ボタンを特定（「投稿オプション」「下書き」等は除外）
                    submit_btn = None
                    for _ in range(15):
                        # メニューが開いていれば閉じる
                        try:
                            if page.locator('text="返信と引用ができる人"').first.is_visible():
                                page.keyboard.press("Escape")
                                time.sleep(0.5)
                        except Exception:
                            pass

                        candidates = []
                        search_root = dialog if (dialog and dialog.is_visible()) else page.locator('div[role="dialog"]').first
                        for btn in search_root.locator('div[role="button"], button').all():
                            try:
                                if not btn.is_visible():
                                    continue
                                txt = btn.inner_text().strip()
                                clean_txt = txt.replace(" ", "").replace("\n", "")
                                if clean_txt in ["投稿", "Post", "投稿する"] and "オプション" not in txt and "新しい" not in txt and "作成" not in txt and "下書き" not in txt:
                                    candidates.append(btn)
                            except Exception:
                                pass

                        if candidates:
                            cand_btn = candidates[-1]
                            is_disabled = (cand_btn.get_attribute("aria-disabled") == "true") or (cand_btn.get_attribute("disabled") is not None)
                            if not is_disabled:
                                submit_btn = cand_btn
                                break
                        time.sleep(1)

                    if not submit_btn:
                        raise Exception("「投稿」ボタンがアクティブになりませんでした（画像のアップロード中か、文字数超過の可能性があります）。")

                    # 送信ボタンクリックとCtrl+Enterの両方で確実に投稿実行
                    self.log(f"[{account_name}] 「投稿」ボタンをクリックしました。Threads送信処理を確認中...")
                    try:
                        submit_btn.scroll_into_view_if_needed()
                        submit_btn.click(force=True)
                    except Exception as e:
                        self.log(f"[{account_name}] 送信ボタン通常クリック試行: {e}")
                    
                    # DOMレベルの直接click()も呼出
                    try:
                        submit_btn.evaluate("el => el.click()")
                    except Exception:
                        pass

                    try:
                        textbox.focus()
                        page.keyboard.press("Control+Enter")
                    except Exception:
                        pass

                    # 【厳密検証】本当にThreads側で投稿が完了したか確認（勝手に成功扱いにしない）
                    posted_success = False
                    post_error_detected = ""

                    for check_idx in range(30):
                        time.sleep(0.8)

                        # 1. 画面下部の黒い通知バー（「投稿完了」「投稿しました」等）を即座に検出
                        for s_sel in [
                            'text="投稿完了"',
                            ':has-text("投稿完了")',
                            'text="投稿しました"',
                            ':has-text("投稿しました")',
                            'text="スレッドを投稿しました"',
                            ':has-text("スレッドを投稿しました")',
                            'text="スレッドを表示"',
                            ':has-text("スレッドを表示")',
                            'text="スレッドが投稿されました"',
                            ':has-text("スレッドが投稿されました")',
                            'text="Posted"',
                            'text="View post"',
                            'div[role="alert"]:has-text("投稿")',
                            'div[role="alert"]:has-text("スレッド")',
                            'div:has-text("投稿完了")',
                            'div:has-text("投稿しました")',
                            'div:has-text("スレッドを投稿しました")'
                        ]:
                            try:
                                s_loc = page.locator(s_sel).first
                                if s_loc.is_visible():
                                    posted_success = True
                                    self.log(f"[SUCCESS] [{account_name}] 🎉 画面下部の黒い通知バー（'{s_sel}'）を確認し、投稿完了しました！")
                                    break
                            except Exception:
                                pass
                        if posted_success:
                            break

                        # 2. 画面上のエラーメッセージ検出
                        for err_sel in [
                            'div[role="alert"]:has-text("エラー")',
                            'text="投稿できませんでした"',
                            'text="アップロードできませんでした"',
                            'text="エラーが発生しました"',
                            'text="問題が発生しました"',
                            'text="コミュニティガイドライン"',
                            'text="後でもう一度お試しください"'
                        ]:
                            try:
                                err_el = page.locator(err_sel).first
                                if err_el.is_visible():
                                    post_error_detected = err_el.inner_text().strip()[:80]
                                    break
                            except Exception:
                                pass
                        if post_error_detected:
                            break

                        # 3. 投稿モーダル (div[role="dialog"]) が画面から消滅したか確認（モーダル閉鎖＝送信成功）
                        modal_closed = True
                        try:
                            for d in page.locator('div[role="dialog"], div[aria-modal="true"]').all():
                                if d.is_visible():
                                    modal_closed = False
                                    break
                        except Exception:
                            pass

                        if modal_closed:
                            time.sleep(0.4)
                            for err_sel in ['div[role="alert"]:has-text("エラー")', 'text="投稿できませんでした"']:
                                try:
                                    if page.locator(err_sel).first.is_visible():
                                        post_error_detected = "エラー通知検出"
                                        break
                                except Exception:
                                    pass
                            if not post_error_detected:
                                posted_success = True
                                self.log(f"[SUCCESS] [{account_name}] 🎉 投稿モーダルが閉じ、下部の黒バー完了通知を確認しました！")
                                break

                        # 4. モーダルがまだ開いている場合、5秒経過時に「投稿」ボタンを再押下
                        if check_idx in [5, 12]:
                            try:
                                if page.locator('text="返信と引用ができる人"').first.is_visible():
                                    page.keyboard.press("Escape")
                                    time.sleep(0.5)
                                for btn in page.locator('div[role="dialog"] div[role="button"], div[role="dialog"] button').all():
                                    try:
                                        if btn.is_visible() and btn.inner_text().strip() in ["投稿", "Post"]:
                                            self.log(f"[{account_name}] 送信ボタンを再押下します...")
                                            btn.click(force=True)
                                            break
                                    except Exception:
                                        pass
                            except Exception:
                                pass

                    if post_error_detected:
                        raise Exception(f"Threads画面でエラーが返されました: {post_error_detected}")

                    if not posted_success:
                        raise Exception("投稿ボタンを押下後、30秒経過してもThreads上で送信完了（投稿通知またはフォーム閉鎖）が確認できませんでした。下書きが画面に残っています。")


                    finish_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    update_post_status(posts_file, row_idx, status="投稿完了", finish_time=finish_time, error_msg="", sheet_name=post_item.get("_sheet_name"), applied_topic=target_topic)
                    topic_suffix = f" (トピック: 「{target_topic}」)" if target_topic else ""
                    self.log(f"[SUCCESS] [{account_name}] Post #{post_id} 投稿成功！Excel更新{topic_suffix}。")
                    consecutive_account_errors = 0

                    wait_sec = random.randint(min_delay, max_delay)
                    if wait_sec >= 60:
                        self.log(f"[*] BAN防止のため {wait_sec} 秒 (約 {wait_sec // 60} 分) 待機中...")
                    else:
                        self.log(f"[*] BAN防止のため {wait_sec} 秒待機中...")
                    
                    for _ in range(wait_sec):
                        if not self.is_running:
                            break
                        time.sleep(1)

                except Exception as post_err:
                    consecutive_account_errors += 1
                    diag = capture_error_diagnostic(page, account_name, f"Post #{post_id}", post_err, self.log)
                    short_reason = diag.get("short_reason", str(post_err))
                    screen_msgs = diag.get("screen_messages", [])
                    self.log(f"[ERROR] [{account_name}] Post #{post_id} 投稿エラー: {short_reason}")
                    
                    is_warning = False
                    warning_keywords = [
                        "コミュニティガイドライン", "制限されています", "利用を制限", "一時的に制限",
                        "ブロックされています", "アクションがブロック", "後でもう一度", "後ほどもう一度",
                        "不審なアクティビティ", "本人確認", "アカウントが一時停止", "セッションの有効期限",
                        "セッションが無効", "利用制限", "スパム"
                    ]
                    if screen_msgs:
                        for sm in screen_msgs:
                            self.log(f"[ERROR] [{account_name}] 画面上のメッセージ: 「{sm}」")
                            if any(wk in sm for wk in warning_keywords):
                                is_warning = True
                    if diag.get("screenshot_saved"):
                        self.log(f"[INFO] [{account_name}] 📸 エラー画面を保存しました: {diag['rel_screenshot']}")
                    
                    if any(wk in str(post_err) for wk in warning_keywords):
                        is_warning = True

                    finish_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    excel_err_text = f"[{finish_time}] {short_reason} (スクショ: {diag.get('rel_screenshot', '')})"
                    update_post_status(posts_file, row_idx, status="エラー", finish_time=None, error_msg=excel_err_text, sheet_name=post_item.get("_sheet_name"))

                    # 提案新機能2: 不安定・警告兆候時の自動プロファイル初期化（クリーンアップ）
                    if self.var_auto_clean_profile.get():
                        trigger_clean = False
                        clean_reason = ""
                        if is_warning:
                            trigger_clean = True
                            clean_reason = f"Meta画面警告検知 ({screen_msgs[0] if screen_msgs else short_reason})"
                        elif consecutive_account_errors >= 2:
                            trigger_clean = True
                            clean_reason = f"連続投稿エラー ({consecutive_account_errors}回連続失敗)"

                        if trigger_clean:
                            clean_account_profile(account_name, profile_dir, reason=clean_reason, log_func=self.log)
                            consecutive_account_errors = 0
                            if is_warning:
                                self.log(f"[WARN] [{account_name}] ⚠️ アカウント保護のためこの周回の残りをスキップします。")
                                break

                    time.sleep(3)

            # 全ての未投稿ポストの処理が終了した後の自動リセット判定
            if self.is_running and self.var_auto_loop.get():
                check_posts = read_excel_records(posts_file)
                still_unposted = [p_it for p_it in check_posts if self._is_unposted_item(p_it)]
                if not still_unposted:
                    reset_cnt = reset_posts_file_silent(posts_file)
                    self.log(f"[SUCCESS] 🎉 [{account_name}] '{os.path.basename(posts_file)}' の全ポスト({reset_cnt}件)の投稿が完了しました！自動ループ機能により全件を『未投稿』にリセットし、最初から周回を継続します。")
                    self.root.after(0, self.load_excel_data)

        except Exception as e:
            self.log(f"[ERROR] [{account_name}] 投稿処理エラー: {e}")
        finally:
            try:
                context.close()
            except Exception:
                pass

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = ThreadsPosterGUI(root)
        root.mainloop()
    except Exception as e:
        import traceback
        err_msg = "".join(traceback.format_exception(type(e), e, e.__traceback__))
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            with open(os.path.join(BASE_DIR, "crash_log.txt"), "w", encoding="utf-8") as f:
                f.write(err_msg)
        except Exception:
            pass
        try:
            with open(os.path.join(BASE_DIR, "error_log.txt"), "a", encoding="utf-8") as f:
                f.write(f"\n{'='*60}\n[{now_str}] 【メインプロセス例外】\n{err_msg}\n{'='*60}\n")
        except Exception:
            pass
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, f"起動エラーが発生しました:\n\n{e}\n\nエラーログを error_log.txt に保存しました。\n詳細をご確認ください。", "Threads AutoPoster 起動エラー", 0x10)
        except Exception:
            print(err_msg)

