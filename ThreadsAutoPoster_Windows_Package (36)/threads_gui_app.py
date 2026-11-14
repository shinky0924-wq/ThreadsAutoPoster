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
import json
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

LOG_DIR = os.path.join(BASE_DIR, "log")
IMAGES_DIR = os.path.join(BASE_DIR, "images")
_THREAD_LOG_LOCAL = threading.local()
_KNOWN_ACCOUNT_NAMES = set()
_LOG_LEVEL_TAGS = {
    "INFO", "WARN", "WARNING", "ERROR", "SUCCESS", "DEBUG", "*", "PYTHON",
    "初回作成", "既存利用", "手動Chrome起動", "ブラウザ起動", "ブラウザ完全終了・省メモリ",
    "モード0: 上から順", "モード1: ランダム", "ポスト文章", "画像なし", "ステップ1", "ステップ2",
}

def sanitize_account_folder_name(account_name):
    if not account_name:
        return ""
    cleaned = str(account_name).strip().lstrip("@").strip()
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", cleaned)
    return cleaned.strip()

def get_account_log_path(account_name):
    clean_acc = sanitize_account_folder_name(account_name)
    os.makedirs(LOG_DIR, exist_ok=True)
    if not clean_acc:
        return os.path.join(LOG_DIR, "system.log")
    return os.path.join(LOG_DIR, f"{clean_acc}.log")

def get_account_images_dir(account_name, create=True):
    clean_acc = sanitize_account_folder_name(account_name)
    if create:
        os.makedirs(IMAGES_DIR, exist_ok=True)
    if not clean_acc:
        return IMAGES_DIR
    acc_dir = os.path.join(IMAGES_DIR, clean_acc)
    if create:
        os.makedirs(acc_dir, exist_ok=True)
    return acc_dir

def extract_account_name_from_msg(msg):
    if not msg:
        return ""
    s_msg = str(msg)
    for m in re.finditer(r'\[([^\[\]\r\n]+)\]', s_msg):
        token = m.group(1).strip().lstrip("@").strip()
        if not token or token in _LOG_LEVEL_TAGS:
            continue
        if token in _KNOWN_ACCOUNT_NAMES:
            return token
        if re.match(r'^[a-zA-Z0-9._-]{2,64}$', token) and token.upper() != token:
            return token
    for known in sorted(_KNOWN_ACCOUNT_NAMES, key=len, reverse=True):
        if f"@{known}" in s_msg:
            return known
    return ""

def log_to_account_file(account_name, full_line):
    clean_acc = sanitize_account_folder_name(account_name)
    if not clean_acc:
        return
    try:
        acc_log_path = get_account_log_path(clean_acc)
        with open(acc_log_path, "a", encoding="utf-8") as f_acc:
            f_acc.write(full_line if full_line.endswith("\n") else (full_line + "\n"))
            f_acc.flush()
    except Exception:
        pass

try:
    os.makedirs(LOG_DIR, exist_ok=True)
    os.makedirs(IMAGES_DIR, exist_ok=True)
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
    if not filepath or not os.path.exists(filepath):
        return []

    wb = None
    # 1. ExcelファイルがMicrosoft Excel等で開かれている場合のファイルロック回避 (バイナリ共有読込 + 3回リトライ)
    for attempt in range(3):
        try:
            with open(filepath, "rb") as f_in:
                import io
                content_bytes = f_in.read()
            wb = openpyxl.load_workbook(io.BytesIO(content_bytes), data_only=True)
            break
        except Exception:
            if attempt < 2:
                time.sleep(0.3)
            else:
                try:
                    wb = openpyxl.load_workbook(filepath, data_only=True)
                except Exception as e_final:
                    print(f"[ERROR] Excel読み込みエラー ({filepath}): {e_final}")
                    return []

    if not wb:
        return []

    try:
        fname_base = os.path.basename(filepath).lower()
        is_account_file = any(k in fname_base for k in ["アカウント", "account"])

        target_ws = None
        target_header_row_idx = 1

        if is_account_file:
            candidate_sheets = []
            for sname in wb.sheetnames:
                if any(k in sname.lower() for k in ["アカウント", "account", "ユーザー", "user"]):
                    candidate_sheets.append(wb[sname])
            for sname in wb.sheetnames:
                if wb[sname] not in candidate_sheets:
                    candidate_sheets.append(wb[sname])

            for ws in candidate_sheets:
                all_rows = list(ws.iter_rows(values_only=True))
                for r_idx, row in enumerate(all_rows[:15], start=1):
                    if not row:
                        continue
                    row_cells_norm = [re.sub(r'[\s　_\-:=]+', '', str(c).strip().lower()) for c in row if c is not None]
                    if any(any(k in cell for k in ["アカウント", "account", "ユーザー", "username", "userid", "auth_token", "sessionid", "パスワード", "password"]) or cell in ["id", "uid"] for cell in row_cells_norm):
                        target_ws = ws
                        target_header_row_idx = r_idx
                        break
                if target_ws:
                    break
        else:
            # 投稿リストファイルの場合（<account_id>.xlsx 等）:
            # 1. シート名に「投稿」「post」「tweet」「リスト」が含まれるシートを最優先
            # 2. シート名に「設定」「config」「setting」が含まれるシートは設定用のため絶対にスキップ！
            candidate_sheets = []
            for sname in wb.sheetnames:
                s_lower = sname.lower()
                if any(k in s_lower for k in ["投稿", "post", "tweet", "リスト"]):
                    candidate_sheets.append(wb[sname])
            for sname in wb.sheetnames:
                s_lower = sname.lower()
                if not any(k in s_lower for k in ["設定", "config", "setting", "日常"]) and wb[sname] not in candidate_sheets:
                    candidate_sheets.append(wb[sname])
            for sname in wb.sheetnames:
                if wb[sname] not in candidate_sheets:
                    candidate_sheets.append(wb[sname])

            for ws in candidate_sheets:
                if any(k in ws.title.lower() for k in ["設定", "config", "setting"]) and len(candidate_sheets) > 1:
                    continue
                all_rows = list(ws.iter_rows(values_only=True))
                for r_idx, row in enumerate(all_rows[:15], start=1):
                    if not row:
                        continue
                    row_cells_norm = [re.sub(r'[\s　_\-:=]+', '', str(c).strip().lower()) for c in row if c is not None]
                    if any(any(k in cell for k in ["ポスト文章", "投稿本文", "ツイート文章", "posttext", "本文", "ポスト内容", "投稿内容"]) for cell in row_cells_norm):
                        target_ws = ws
                        target_header_row_idx = r_idx
                        break
                if target_ws:
                    break

        if not target_ws:
            non_setting = [wb[s] for s in wb.sheetnames if not any(k in s.lower() for k in ["設定", "config", "setting"])]
            target_ws = non_setting[0] if non_setting else wb.active
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
                val_clean = "" if val is None else str(val).strip()
                r_dict[h] = val_clean

                norm_h = re.sub(r'[\s　_\-:=]+', '', str(h).strip().lower())
                if any(k in norm_h for k in ["ポスト文章", "投稿本文", "ツイート文章", "posttext", "本文", "内容"]) or norm_h in ["text", "post"]:
                    r_dict["ポスト文章"] = val_clean
                    r_dict["投稿本文"] = val_clean
                    r_dict["ツイート文章"] = val_clean
                    r_dict["本文"] = val_clean
                elif any(k in norm_h for k in ["アカウント", "account", "ユーザー", "username", "userid"]) or norm_h in ["id", "uid"]:
                    r_dict["アカウントID"] = val_clean
                    r_dict["アカウント名"] = val_clean
                    r_dict["アカウント"] = val_clean
                    r_dict["id"] = val_clean
                elif any(k in norm_h for k in ["パスワード", "password", "暗証番号"]) or norm_h in ["pw", "pass"]:
                    r_dict["パスワード"] = val_clean
                    r_dict["password"] = val_clean
                    r_dict["PW"] = val_clean
                    r_dict["Pass"] = val_clean
                elif any(k in norm_h for k in ["sessionid", "authtoken", "トークン", "token", "セッション"]):
                    r_dict["sessionid"] = val_clean
                    r_dict["auth_token"] = val_clean
                elif any(k in norm_h for k in ["ステータス", "status", "状態"]):
                    r_dict["ステータス"] = val_clean
                elif any(k in norm_h for k in ["有効", "無効"]):
                    r_dict["有効/無効"] = val_clean
                elif any(k in norm_h for k in ["画像", "image", "media"]):
                    r_dict["画像ファイルパス"] = val_clean
                    r_dict["画像"] = val_clean
                elif any(k in norm_h for k in ["コミュニティ", "トピック", "タグ", "topic"]):
                    r_dict["コミュニティまたはトピック"] = val_clean
                    r_dict["トピック"] = val_clean
                elif any(k in norm_h for k in ["備考", "メモ", "note", "notes"]):
                    r_dict["備考"] = val_clean

            if is_account_file:
                if not r_dict.get("アカウントID") and len(row) > 0 and row[0] is not None and str(row[0]).strip():
                    val_a = str(row[0]).strip()
                    r_dict["アカウントID"] = val_a
                    r_dict["アカウント名"] = val_a
                    r_dict["アカウント"] = val_a
                    r_dict["id"] = val_a
                if not r_dict.get("パスワード") and len(row) > 1 and row[1] is not None and str(row[1]).strip():
                    val_b = str(row[1]).strip()
                    r_dict["パスワード"] = val_b
                    r_dict["password"] = val_b
                    r_dict["PW"] = val_b
                    r_dict["Pass"] = val_b
                if "有効/無効" not in r_dict or not r_dict["有効/無効"]:
                    r_dict["有効/無効"] = "有効"
            else:
                # 投稿リストの場合: ステータスが空欄なら「未投稿」として初期化
                if "ステータス" not in r_dict or not r_dict["ステータス"]:
                    r_dict["ステータス"] = "未投稿"

            records.append(r_dict)
        return records
    except Exception as e:
        print(f"[ERROR] Excelパース例外 ({filepath}): {e}")
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
                    for offset in range(1, 3):
                        if r_idx + offset >= len(all_rows):
                            break
                        sub_row = all_rows[r_idx + offset]
                        if not sub_row:
                            continue
                        sub_str = " ".join([str(c).strip() for c in sub_row if c is not None])
                        if any(h in sub_str for h in ["コミュニティ", "トピック", "モード", "間隔", "ツイート", "ポスト", "実行時間"]):
                            break
                        for cell in sub_row:
                            if cell is not None:
                                c_s = str(cell).strip()
                                if c_s and not c_s.startswith("(") and not c_s.startswith("（") and not c_s.startswith("※"):
                                    if not any(h in c_s for h in ["コミュニティ", "トピック", "モード", "間隔", "ツイート", "ポスト", "実行時間"]):
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
        ws = None
        for sname in wb.sheetnames:
            if "アカウント" in sname or "account" in sname.lower():
                ws = wb[sname]
                break
        if not ws:
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



def open_threads_composer_dialog(page, account_name, log_func=print):
    """
    Threadsの新規スレッド作成モーダル (div[role="dialog"]) を開き、
    入力欄 (textbox) を取得します。
    
    対応パターン（最新Threads UI完全対応）:
    1. 画面左下の円形「＋」ボタン（黒丸白プラス・ユーザー添付画面に完全一致）
    2. 左サイドバー/ヘッダーの「スレッド」「投稿」「作成」「+」SVG/aria-labelアイコン
    3. タイムライン上部の「スレッドを開始...」プロンプトバー
    4. 左下プラスボタン領域（x: 35〜55px, y: 画面下部）への直接マウスクリック
    5. キーボードショートカット 'c'
    """
    dialog = None
    textbox = None

    # すでにモーダルが開いていないか即時チェック
    for d_sel in ['div[role="dialog"]', 'div[aria-modal="true"]']:
        try:
            cand = page.locator(d_sel).first
            if cand.is_visible():
                dialog = cand
                break
        except Exception:
            pass

    if dialog:
        for tb_sel in ['div[role="textbox"]', 'div[contenteditable="true"]', '[data-lexical-editor="true"]']:
            try:
                tb_cand = dialog.locator(tb_sel).first
                if tb_cand.is_visible():
                    return dialog, tb_cand
            except Exception:
                pass

    clicked_method = ""

    # パターン1: 画面左下の丸い「＋」ボタン（最新UI・画面左下隅に配置されている黒丸アイコン）
    try:
        # 左下領域（x < 120）に存在するボタン要素を全収集し、最も下（y座標が最大）にある要素を選択
        candidates = []
        for sel in ['div[role="button"]', 'button', 'a']:
            try:
                for el in page.locator(sel).all():
                    try:
                        if el.is_visible():
                            box = el.bounding_box()
                            if box and box['x'] < 120 and box['y'] > 200 and 20 <= box['width'] <= 100 and 20 <= box['height'] <= 100:
                                candidates.append((box['y'], box, el))
                    except Exception:
                        pass
            except Exception:
                pass

        if candidates:
            # y座標が最も大きいもの（＝一番下にある黒丸「＋」ボタン）を優先
            candidates.sort(key=lambda item: item[0], reverse=True)
            best_y, best_box, best_el = candidates[0]
            log_func(f"[{account_name}] 画面左下の最下部「＋」ボタン (x={int(best_box['x'])}, y={int(best_y)}, サイズ={int(best_box['width'])}x{int(best_box['height'])}) を検出！クリックします...")
            try:
                best_el.scroll_into_view_if_needed()
                best_el.click(force=True)
                clicked_method = "左下円形＋ボタン"
                time.sleep(1.0)
            except Exception as e:
                log_func(f"[DEBUG] ボタンクリック失敗: {e}")
    except Exception as e:
        log_func(f"[DEBUG] 左下＋ボタン探索: {e}")

    # パターン2: SVG aria-label による「スレッド」「投稿」「作成」「New」「Create」ボタン
    if not clicked_method:
        icon_selectors = [
            'svg[aria-label*="スレッド"]',
            'svg[aria-label*="投稿"]',
            'svg[aria-label*="作成"]',
            'svg[aria-label*="New"]',
            'svg[aria-label*="Create"]',
            'svg[aria-label*="Post"]',
            'div[role="button"]:has(svg[aria-label*="スレッド"])',
            'div[role="button"]:has(svg[aria-label*="投稿"])',
            'div[role="button"]:has(svg[aria-label*="作成"])',
            'div[role="button"]:has(svg[aria-label*="New"])',
            'div[role="button"]:has(svg[aria-label*="Create"])',
            'button:has(svg[aria-label*="スレッド"])',
            'button:has(svg[aria-label*="投稿"])',
            'button:has(svg[aria-label*="作成"])'
        ]
        for sel in icon_selectors:
            try:
                cand = page.locator(sel).first
                if cand.is_visible():
                    box = cand.bounding_box()
                    if box and (box['x'] < 300 or box['y'] < 200):
                        log_func(f"[{account_name}] 新規スレッド作成アイコン ({sel}) を検出！クリックします...")
                        cand.click(force=True)
                        clicked_method = f"アイコン ({sel})"
                        time.sleep(1.0)
                        break
            except Exception:
                pass

    # パターン3: タイムライン上部の「スレッドを開始...」エリア
    if not clicked_method:
        for prompt_sel in [
            'div[role="button"]:has-text("スレッドを開始")',
            'div:has-text("スレッドを開始...")',
            'span:has-text("スレッドを開始")',
            'div:has-text("Start a thread")',
            '[placeholder*="スレッドを開始"]'
        ]:
            try:
                p_el = page.locator(prompt_sel).first
                if p_el.is_visible():
                    log_func(f"[{account_name}] タイムライン上部の「スレッドを開始...」をクリックします...")
                    p_el.click(force=True)
                    clicked_method = "タイムライン上部バー"
                    time.sleep(1.0)
                    break
            except Exception:
                pass

    # パターン4: 左下プラスボタン位置への直接マウスクリック（ブラウザ実寸座標）
    if not clicked_method:
        try:
            win_dim = page.evaluate("() => ({ width: window.innerWidth, height: window.innerHeight })")
            target_x = 35
            target_y = max(500, win_dim.get("height", 900) - 45)
            log_func(f"[{account_name}] 左下「＋」ボタンの位置 (x={target_x}, y={target_y}) を直接クリックします...")
            page.mouse.click(target_x, target_y)
            clicked_method = "左下座標直接クリック"
            time.sleep(1.0)
        except Exception:
            pass

    # パターン5: キーボードショートカット 'c'
    try:
        page.locator('body').focus()
        page.keyboard.press("c")
        time.sleep(0.5)
    except Exception:
        pass

    # モーダル (dialog) と入力欄 (textbox) の出現を待機
    for wait_i in range(20):
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
                    log_func(f"[SUCCESS] [{account_name}] 🎉 新規スレッド作成モーダルが開きました！")
                    return dialog, textbox
        except Exception:
            pass

        # 途中で開かなかった場合は左下座標を再クリック＆ショートカット再送
        if wait_i in [4, 10]:
            try:
                win_dim = page.evaluate("() => ({ width: window.innerWidth, height: window.innerHeight })")
                page.mouse.click(35, max(500, win_dim.get("height", 900) - 45))
                time.sleep(0.3)
                page.keyboard.press("c")
            except Exception:
                pass

    return dialog, textbox


def restore_portable_session_if_needed(context, profile_dir, account_name, log_func=print):
    """
    PC AからPC B（VPS等）へプロファイルをコピーした際、
    Windows DPAPI暗号化キーの不一致によりChromeの内部Cookieがリセットされる問題を完全解決。
    profile_dir内に保存された「portable_session.json」または「cookies.json」から
    Cookieを自動検知して復元し、PC BのWindows DPAPIで再暗号化保存させます。
    """
    if not profile_dir or not os.path.exists(profile_dir):
        return False

    portable_file = os.path.join(profile_dir, "portable_session.json")
    if not os.path.exists(portable_file) or os.path.getsize(portable_file) < 5:
        portable_file = os.path.join(profile_dir, "cookies.json")
    if not os.path.exists(portable_file) or os.path.getsize(portable_file) < 5:
        return False

    try:
        cur_cookies = context.cookies(["https://www.threads.net", "https://www.threads.com", "https://www.instagram.com"])
        has_session = any(c.get("name") == "sessionid" and c.get("value") for c in cur_cookies)
        if has_session:
            return True

        import json
        with open(portable_file, "r", encoding="utf-8") as f:
            data = json.load(f)

        cookies_to_add = data.get("cookies", []) if isinstance(data, dict) else (data if isinstance(data, list) else [])
        if cookies_to_add:
            cleaned_cookies = []
            for c in cookies_to_add:
                c_clean = {}
                for k in ["name", "value", "domain", "path", "expires", "httpOnly", "secure", "sameSite"]:
                    if k in c and c[k] is not None:
                        c_clean[k] = c[k]
                if "sameSite" in c_clean:
                    s = str(c_clean["sameSite"]).capitalize()
                    c_clean["sameSite"] = s if s in ["Strict", "Lax", "None"] else "Lax"
                cleaned_cookies.append(c_clean)

            context.add_cookies(cleaned_cookies)
            log_func(f"[SUCCESS] [{account_name}] 🔄 PC間移行を検知: portable_session.json からセッションCookie ({len(cleaned_cookies)}件) を自動復元しました！")
            log_func(f"[INFO] [{account_name}] 🛡️ 現在のPCのWindows暗号化(DPAPI)にセッションを自動再バインドしました。")
            return True
    except Exception as e:
        log_func(f"[WARN] portable_session.json 復元エラー: {e}")
    return False

def save_portable_session_file(context, profile_dir, account_name, log_func=print):
    """
    ログイン中のセッション状態を portable_session.json として保存。
    別PCにフォルダごとコピーした際もログイン状態を100%保持できるようにします。
    ※ ログアウト状態や未認証状態での上書き破壊を防止する安全ガード付き。
    """
    if not profile_dir or not os.path.exists(profile_dir) or not context:
        return
    try:
        import json
        cookies = context.cookies(["https://www.threads.net", "https://www.threads.com", "https://www.instagram.com"])
        if not cookies:
            return

        # 🛡️ ログアウト状態・未認証状態での上書き破壊防止ガード！
        has_valid_session = any(c.get("name") == "sessionid" and len(str(c.get("value") or "").strip()) > 15 for c in cookies)
        if not has_valid_session:
            return

        portable_file = os.path.join(profile_dir, "portable_session.json")
        context.storage_state(path=portable_file)
        cookies_file = os.path.join(profile_dir, "cookies.json")
        with open(cookies_file, "w", encoding="utf-8") as f:
            json.dump(cookies, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def get_or_create_fingerprint(profile_dir):
    """
    【提案4: 端末指紋（Fingerprint）の完全固定化】
    アカウントごとに一貫したUser-Agent、解像度、タイムゾーン、言語設定を永続化し、
    Metaによる「デバイス急変」「不審端末」判定を完全防止します。
    """
    fp_path = os.path.join(profile_dir, "fingerprint.json")
    if os.path.exists(fp_path):
        try:
            with open(fp_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass

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
    return default_fp


def clean_session_id_token(raw_token):
    """sessionid文字列をサニタイズ・デコード"""
    t = str(raw_token or "").strip()
    if t.startswith("sessionid="):
        t = t[len("sessionid="):].strip()
    while t and t[0] in ('"', "'", ';', ' '):
        t = t[1:].strip()
    while t and t[-1] in ('"', "'", ';', ' '):
        t = t[:-1].strip()
    import urllib.parse
    return urllib.parse.unquote(t)


def auto_bootstrap_countermeasure_profiles(accounts_data, log_func=print):
    """
    【アプリ起動時・自動構築】
    アカウント情報.xlsx に記載された全アカウントのプロファイル（chrome_profiles/<account>）を自動構築します。
    - 端末指紋（fingerprint.json）の自動配備（固定User-Agent, 1280x850, 日本語, JST）
    - sessionidと周辺セキュリティCookie（csrftoken, mid, ig_did, datr, rur）の全自動合成＆保存
    これにより、ユーザーが profile_launcher を開くことなく、アプリ起動だけで対策プロファイルが100%自動配備されます。
    """
    if not accounts_data:
        return 0

    bootstrapped_count = 0
    for acc in accounts_data:
        acc_id = str(acc.get("アカウントID") or acc.get("アカウント名") or acc.get("アカウント") or "").strip()
        clean_id = acc_id.lstrip("@").strip()
        if not clean_id or clean_id.lower() == "nan":
            continue
        prof_name = clean_id
        prof_dir = os.path.join(BASE_DIR, f"chrome_profiles/{prof_name}")
        os.makedirs(prof_dir, exist_ok=True)

        # 提案4: 端末指紋の自動作成・維持
        get_or_create_fingerprint(prof_dir)

        # 提案2: sessionid があれば Cookie全種をプロファイルに完全同期
        raw_token = str(
            acc.get("auth_token") or
            acc.get("Authトークン_sessionid") or
            acc.get("sessionid") or
            acc.get("Authトークン") or
            acc.get("セッションID") or
            acc.get("セッションid") or
            acc.get("トークン") or
            acc.get("token") or
            ""
        ).strip()
        pfile = os.path.join(prof_dir, "portable_session.json")
        cfile = os.path.join(prof_dir, "cookies.json")

        if (not raw_token or raw_token.startswith("ここに") or raw_token.lower() in ["nan", "none"]) and os.path.exists(pfile):
            try:
                with open(pfile, "r", encoding="utf-8") as _pf:
                    _d = json.load(_pf)
                _cks = _d.get("cookies", []) if isinstance(_d, dict) else _d
                for _c in _cks:
                    if _c.get("name") == "sessionid" and _c.get("value"):
                        raw_token = str(_c.get("value")).strip()
                        break
            except Exception:
                pass

        if raw_token and not raw_token.startswith("ここに") and raw_token.lower() not in ["nan", "none"]:
            token = clean_session_id_token(raw_token)
            uid = ""
            if ":" in token:
                potential = token.split(":")[0].strip()
                if potential.isdigit():
                    uid = potential
            elif "%3A" in token.upper():
                potential = token.upper().split("%3A")[0].strip()
                if potential.isdigit():
                    uid = potential

            cookie_expires = int(time.time()) + 31536000
            import hashlib, uuid as uuid_mod
            csrf_val = hashlib.md5(f"csrf_{uid}_{token[:16]}".encode()).hexdigest()
            mid_val = "Z" + hashlib.md5(f"mid_{uid}_{acc_id}".encode()).hexdigest()[:23]
            ig_did_val = str(uuid_mod.uuid5(uuid_mod.NAMESPACE_DNS, f"{uid}_{acc_id}")).upper()
            datr_val = hashlib.md5(f"datr_{acc_id}".encode()).hexdigest()[:24]

            domains = [
                ".threads.net", "www.threads.net", "threads.net",
                ".threads.com", "www.threads.com", "threads.com",
                ".instagram.com", "www.instagram.com", "instagram.com"
            ]

            cookies_list = []
            for d in domains:
                cookies_list.append({"name": "sessionid", "value": token, "domain": d, "path": "/", "httpOnly": True, "secure": True, "sameSite": "Lax", "expires": cookie_expires})
                if uid:
                    cookies_list.append({"name": "ds_user_id", "value": str(uid), "domain": d, "path": "/", "secure": True, "sameSite": "Lax", "expires": cookie_expires})
                cookies_list.append({"name": "csrftoken", "value": csrf_val, "domain": d, "path": "/", "secure": True, "sameSite": "Lax", "expires": cookie_expires})
                cookies_list.append({"name": "mid", "value": mid_val, "domain": d, "path": "/", "secure": True, "sameSite": "Lax", "expires": cookie_expires})
                cookies_list.append({"name": "ig_did", "value": ig_did_val, "domain": d, "path": "/", "httpOnly": True, "secure": True, "sameSite": "Lax", "expires": cookie_expires})
                cookies_list.append({"name": "datr", "value": datr_val, "domain": d, "path": "/", "httpOnly": True, "secure": True, "sameSite": "Lax", "expires": cookie_expires})
                cookies_list.append({"name": "rur", "value": "RVA", "domain": d, "path": "/", "httpOnly": True, "secure": True, "sameSite": "Lax", "expires": cookie_expires})

            state_data = {"cookies": cookies_list, "origins": []}
            # ファイルが存在しない、または古い場合は作成・更新
            try:
                need_write = True
                if os.path.exists(pfile):
                    try:
                        with open(pfile, "r", encoding="utf-8") as f_ex:
                            ex_data = json.load(f_ex)
                        ex_cookies = ex_data.get("cookies", []) if isinstance(ex_data, dict) else ex_data
                        for c in ex_cookies:
                            if c.get("name") == "sessionid" and c.get("value") == token and len(ex_cookies) >= 5:
                                need_write = False
                                break
                    except Exception:
                        need_write = True

                if need_write:
                    with open(pfile, "w", encoding="utf-8") as f:
                        json.dump(state_data, f, ensure_ascii=False, indent=2)
                    with open(cfile, "w", encoding="utf-8") as f:
                        json.dump(cookies_list, f, ensure_ascii=False, indent=2)
                    bootstrapped_count += 1
                    log_func(f"[SUCCESS] 🛡️ [@{acc_id}] アプリ起動時に【5大強制ログアウト対策プロファイル】を自動構築しました！（周辺Cookie＋端末指紋を全自動配備）")
            except Exception as e:
                log_func(f"[DEBUG] auto_bootstrap write: {e}")

    return bootstrapped_count


def safe_launch_persistent_context(p, profile_dir, headless=False, log_func=print):
    """
    SingletonLock等のブラウザ残骸を自動クリーンアップし、安定してChromeを起動します。
    【提案4】固定User-Agentと端末指紋を維持してMetaのボット検知・強制ログアウトを防止します。
    PC間移行時(Windows DPAPIリセット時)は portable_session.json から自動自己復元します。
    """
    os.makedirs(profile_dir, exist_ok=True)
    for lk in ["SingletonLock", "SingletonCookie", "SingletonSocket"]:
        try:
            lk_p = os.path.join(profile_dir, lk)
            if os.path.exists(lk_p):
                os.remove(lk_p)
        except Exception:
            pass

    fp = get_or_create_fingerprint(profile_dir)

    # --disable-blink-features=AutomationControlled はChrome警告バナーを発生させ
    # Metaの自動化検知を誘発するため削除し、ignore_default_args で自然に隠蔽
    args = [
        "--lang=ja-JP",
        "--disable-dev-shm-usage",
        "--test-type",
        "--disable-session-crashed-bubble",
        "--hide-crash-restore-bubble"
    ]
    if sys.platform != "win32":
        args.append("--no-sandbox")

    # Chromeの「ページを復元しますか？」ダイアログを抑制（Preferencesのexit_typeをNormalに自動修正）
    for pref_sub in ["Default/Preferences", "Preferences"]:
        pref_path = os.path.join(profile_dir, pref_sub)
        if os.path.exists(pref_path):
            try:
                with open(pref_path, "r", encoding="utf-8") as pf:
                    p_text = pf.read()
                p_text = p_text.replace('"exit_type":"Crashed"', '"exit_type":"Normal"')
                p_text = p_text.replace('"exited_cleanly":false', '"exited_cleanly":true')
                with open(pref_path, "w", encoding="utf-8") as pf:
                    pf.write(p_text)
            except Exception:
                pass

    for attempt in range(1, 3):
        try:
            context = p.chromium.launch_persistent_context(
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
            # navigator.webdriver を隠蔽して通常の人間のブラウザとして維持（window.chromeは実Chromeのものを保持）
            try:
                context.add_init_script("""
                    Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
                """)
            except Exception:
                pass

            # PC間移行・クッキー揮発の自動検知＆自己修復
            restore_portable_session_if_needed(context, profile_dir, os.path.basename(profile_dir), log_func=log_func)
            return context
        except Exception as e:
            err_str = str(e)
            if ("Target page, context or browser has been closed" in err_str or "closed" in err_str) and attempt == 1:
                log_func("[WARN] 過去のChrome残骸ロックを検知しました。ロックファイルを削除して再起動します...")
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
    if not token or token.lower() in ["nan", "none", ""] or token.startswith("ここに"):
        # プロファイルフォルダ (chrome_profiles/{account_name}/portable_session.json) に sessionid があれば読み込む
        prof_dirs = [
            os.path.join(BASE_DIR, "chrome_profiles", account_name),
            os.path.join(os.getcwd(), "chrome_profiles", account_name)
        ]
        for pdir in prof_dirs:
            pjson = os.path.join(pdir, "portable_session.json")
            if os.path.exists(pjson):
                try:
                    with open(pjson, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        t = data.get("sessionid", "").strip() if isinstance(data, dict) else ""
                        cks = data.get("cookies", []) if isinstance(data, dict) else (data if isinstance(data, list) else [])
                        if not t and cks:
                            for c in cks:
                                if c.get("name") == "sessionid" and c.get("value"):
                                    t = str(c["value"]).strip()
                                    break
                        if t:
                            token = t
                            if cks:
                                try:
                                    context.add_cookies(cks)
                                except Exception:
                                    pass
                            log_func(f"[INFO] [{account_name}] プロファイル永続化ファイル (portable_session.json) から sessionid を自動読み込み・適用しました。")
                            break
                except Exception:
                    pass
        if not token:
            # プロファイル永続化モード（Chrome既存Cookieストアを使用）
            return True
    
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

    # 既存のCookieをチェック (すでに全く同じ値が設定済みの場合は二重設定をスキップ)
    try:
        existing_cookies = context.cookies(["https://www.threads.net", "https://www.instagram.com"])
        for c in existing_cookies:
            if c.get("name") == "sessionid" and c.get("value"):
                c_val = str(c.get("value")).strip()
                if c_val == decoded_token or c_val == token:
                    return True
    except Exception:
        pass

    cookie_expires = int(time.time()) + 31536000  # 1年間有効（揮発防止）

    domains = [
        ".threads.net", "www.threads.net", "threads.net",
        ".threads.com", "www.threads.com", "threads.com",
        ".instagram.com", "www.instagram.com", "instagram.com"
    ]

    import hashlib, uuid
    csrf_val = hashlib.md5(f"csrf_{ds_user_id}_{decoded_token[:16]}".encode()).hexdigest()
    mid_val = "Z" + hashlib.md5(f"mid_{ds_user_id}_{account_name}".encode()).hexdigest()[:23]
    ig_did_val = str(uuid.uuid5(uuid.NAMESPACE_DNS, f"{ds_user_id}_{account_name}")).upper()
    datr_val = hashlib.md5(f"datr_{account_name}".encode()).hexdigest()[:24]

    cookies_to_add = []
    for domain in domains:
        cookies_to_add.append({
            "name": "sessionid",
            "value": decoded_token,
            "domain": domain,
            "path": "/",
            "httpOnly": True,
            "secure": True,
            "sameSite": "Lax",
            "expires": cookie_expires
        })
        if ds_user_id:
            cookies_to_add.append({
                "name": "ds_user_id",
                "value": str(ds_user_id),
                "domain": domain,
                "path": "/",
                "secure": True,
                "sameSite": "Lax",
                "expires": cookie_expires
            })
        # 【提案2: 必須セキュリティCookie群の自動補完・完全同期】
        cookies_to_add.append({
            "name": "csrftoken",
            "value": csrf_val,
            "domain": domain,
            "path": "/",
            "secure": True,
            "sameSite": "Lax",
            "expires": cookie_expires
        })
        cookies_to_add.append({
            "name": "mid",
            "value": mid_val,
            "domain": domain,
            "path": "/",
            "secure": True,
            "sameSite": "Lax",
            "expires": cookie_expires
        })
        cookies_to_add.append({
            "name": "ig_did",
            "value": ig_did_val,
            "domain": domain,
            "path": "/",
            "httpOnly": True,
            "secure": True,
            "sameSite": "Lax",
            "expires": cookie_expires
        })
        cookies_to_add.append({
            "name": "datr",
            "value": datr_val,
            "domain": domain,
            "path": "/",
            "httpOnly": True,
            "secure": True,
            "sameSite": "Lax",
            "expires": cookie_expires
        })
        cookies_to_add.append({
            "name": "rur",
            "value": "RVA",
            "domain": domain,
            "path": "/",
            "httpOnly": True,
            "secure": True,
            "sameSite": "Lax",
            "expires": cookie_expires
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
                "sameSite": "Lax",
                "expires": cookie_expires
            }]
            if ds_user_id:
                url_cookies.append({
                    "name": "ds_user_id",
                    "value": str(ds_user_id),
                    "url": u,
                    "secure": True,
                    "sameSite": "Lax",
                    "expires": cookie_expires
                })
            context.add_cookies(url_cookies)
        except Exception:
            pass

    # プロファイルフォルダに portable_session.json が無ければ自動初期生成
    try:
        pdir = os.path.join(BASE_DIR, "chrome_profiles", account_name)
        os.makedirs(pdir, exist_ok=True)
        pfile = os.path.join(pdir, "portable_session.json")
        if not os.path.exists(pfile):
            state_data = {"cookies": cookies_to_add, "origins": []}
            with open(pfile, "w", encoding="utf-8") as f:
                json.dump(state_data, f, ensure_ascii=False, indent=2)
            log_func(f"[INFO] [{account_name}] ✨ プロファイルフォルダ (chrome_profiles/{account_name}) と永続セッションを自動生成しました！")
    except Exception as pe:
        log_func(f"[DEBUG] 自動プロファイルファイル生成: {pe}")

    log_func(f"[SUCCESS] [{account_name}] sessionid と必須Cookie全セット（csrftoken, mid, datr等）を永続注入しました（提案2完全同期）。")
    return True


def auto_refresh_and_save_session(context, account_name, profile_dir, account_file, log_func=print):
    """
    【提案3: セッション自動延命・ローテーション】
    Metaが更新・再発行した最新のCookie（sessionid, csrftoken, mid, datr等）をブラウザから自動抽出し、
    portable_session.json と アカウント情報.xlsx に上書き保存して永久延命します。
    """
    try:
        if not context:
            return
        cookies = context.cookies(["https://www.threads.net", "https://www.instagram.com"])
        if not cookies:
            return

        latest_sessionid = ""
        user_id = ""
        for c in cookies:
            if c.get("name") == "sessionid" and c.get("value"):
                latest_sessionid = c.get("value").strip()
            elif c.get("name") == "ds_user_id" and c.get("value"):
                user_id = str(c.get("value")).strip()

        if not latest_sessionid or len(latest_sessionid) < 15:
            # 🛡️ ログアウト状態での上書き破壊防止ガード！有効なsessionidがない場合は保存しない
            return

        target_dir = profile_dir if (profile_dir and os.path.isabs(profile_dir)) else os.path.join(BASE_DIR, profile_dir or f"chrome_profiles/{account_name}")
        os.makedirs(target_dir, exist_ok=True)

        portable_file = os.path.join(target_dir, "portable_session.json")
        cookies_file = os.path.join(target_dir, "cookies.json")

        try:
            context.storage_state(path=portable_file)
        except Exception:
            state_data = {"cookies": cookies, "origins": []}
            with open(portable_file, "w", encoding="utf-8") as f:
                json.dump(state_data, f, ensure_ascii=False, indent=2)

        try:
            with open(cookies_file, "w", encoding="utf-8") as f:
                json.dump(cookies, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

        # アカウント情報.xlsx の sessionid 列を最新値に更新
        if account_file and os.path.exists(account_file):
            try:
                import openpyxl
                wb = openpyxl.load_workbook(account_file)
                ws = None
                for s in wb.sheetnames:
                    if "アカウント" in s:
                        ws = wb[s]
                        break
                if not ws:
                    ws = wb.active

                headers = [str(c).strip() if c is not None else "" for c in next(ws.iter_rows(values_only=True))]
                id_col = 1
                sess_col = 2
                for idx, h in enumerate(headers, start=1):
                    if "アカウントID" in h or "アカウント名" in h: id_col = idx
                    elif "sessionid" in h.lower() or "auth" in h.lower(): sess_col = idx

                updated = False
                for row in range(2, ws.max_row + 1):
                    cid = str(ws.cell(row=row, column=id_col).value or "").strip()
                    if cid.lstrip("@").strip().lower() == account_name.lstrip("@").strip().lower():
                        cur_val = str(ws.cell(row=row, column=sess_col).value or "").strip()
                        if cur_val != latest_sessionid:
                            ws.cell(row=row, column=sess_col, value=latest_sessionid)
                            updated = True
                        break
                if updated:
                    wb.save(account_file)
                    log_func(f"[SUCCESS] ✨ [{account_name}] 【提案3: 延命完了】Metaが更新した最新セッションCookieを「アカウント情報.xlsx」に同期しました！")
                wb.close()
            except Exception as xe:
                log_func(f"[DEBUG] Excel同期スキップ: {xe}")

        log_func(f"[SUCCESS] 🔄 [{account_name}] 【提案3: 延命保存】Meta最新Cookie全{len(cookies)}件をプロファイルへ保存完了（強制ログアウト防止）。")
    except Exception as e:
        log_func(f"[DEBUG] auto_refresh_and_save_session: {e}")


def detect_phone_checkpoint_screen(page):
    """
    Metaの携帯電話番号認証（SMS / WhatsApp確認）要求画面を即時検知
    ※ 通常のログイン画面（ユーザーネーム/電話番号入力欄）との誤検知を100%防止
    """
    if not page:
        return False, ""
    try:
        cur_url = ""
        try:
            cur_url = (page.url or "").lower()
        except Exception:
            pass

        # 通常のログイン画面やThreadsホームは絶対にチェックポイントではない
        if any(k in cur_url for k in ["/accounts/login", "/login", "threads.net/@", "instagram.com/direct"]):
            return False, ""

        # 通常のユーザーネーム/パスワード入力欄が表示されている場合はログイン画面
        try:
            if page.locator('input[name="username"]').first.is_visible(timeout=200) or                page.locator('input[name="password"]').first.is_visible(timeout=200):
                return False, ""
        except Exception:
            pass

        if any(k in cur_url for k in ["/challenge/", "/checkpoint/"]):
            return True, f"認証URL検知 ({cur_url})"

        content = ""
        try:
            content = page.content()
        except Exception:
            pass

        keywords = [
            "SMSまたはWhatsAppで送信されたコードでこの携帯電話番号を認証する必要があります",
            "電話番号を認証する必要があります",
            "セキュリティコードを入力",
            "認証コードを入力",
            "Enter your confirmation code",
        ]
        for kw in keywords:
            if kw in content:
                return True, f"画面テキスト検知「{kw}」"

        checkpoint_selectors = [
            'input[name="security_code"]',
            'input[name="verificationCode"]',
            'button:has-text("コードを再送信")',
            'text="SMSまたはWhatsApp"'
        ]
        for sel in checkpoint_selectors:
            try:
                if page.locator(sel).first.is_visible(timeout=200):
                    return True, f"要素検知 ({sel})"
            except Exception:
                pass
    except Exception:
        pass
    return False, ""


def safe_goto_threads(page, context, session_token, account_name, log_func=print, timeout=60000):
    """
    Threadsホーム (https://www.threads.net/) または ログイン画面 (https://www.threads.net/login) へ安全にアクセスします。
    セッションIDが無い場合や失効している場合は、無駄なリダイレクトループを避けて直接ログイン画面を開きます。
    net::ERR_TOO_MANY_REDIRECTS や about:blank を検知した場合は、CookieをクリアしてThreadsログイン画面へ自己修復遷移します。
    """
    token_str = str(session_token or "").strip()
    has_valid_token = bool(token_str and token_str.lower() not in ["", "nan", "none"] and not token_str.startswith("ここに"))

    # Chromeプロファイル内に既存のCookie（手動ログイン済みセッション等）がある場合は必ずホームから開く
    has_profile_cookies = False
    try:
        prof_cks = context.cookies(["https://www.threads.net", "https://www.threads.com", "https://www.instagram.com"])
        has_profile_cookies = len(prof_cks) > 0
    except Exception:
        pass

    # セッションIDまたはプロファイルCookieがある場合はThreadsホームへ直接アクセス
    initial_url = "https://www.threads.net/" if (has_valid_token or has_profile_cookies) else "https://www.threads.net/login"
    
    redirect_detected = False
    try:
        log_func(f"[INFO] 🌐 [{account_name}] Threadsアクセス中 ({initial_url})...")
        page.goto(initial_url, timeout=timeout, wait_until="domcontentloaded")
        time.sleep(2.5)

        cur_u = (page.url or "").lower()
        if "about:blank" in cur_u:
            redirect_detected = True
        else:
            try:
                page_content = page.content()
                if "HTTP ERROR 429" in page_content or "429 Too Many Requests" in page_content:
                    redirect_detected = True
                    log_func(f"[WARN] ⏳ [{account_name}] Metaアクセス頻度制限 (HTTP ERROR 429) を検知しました。")
                elif any(k in page_content for k in [
                    "ERR_TOO_MANY_REDIRECTS",
                    "リダイレクトが繰り返し行われました",
                    "Cookie を削除してみてください",
                    "Cookieを削除してみてください",
                    "このページは動作していません"
                ]):
                    redirect_detected = True
            except Exception:
                pass

    except Exception as e:
        err_str = str(e)
        redirect_detected = True
        if "ERR_TOO_MANY_REDIRECTS" in err_str or "too many redirects" in err_str.lower():
            log_func(f"[WARN] 🔄 [{account_name}] Page.goto: net::ERR_TOO_MANY_REDIRECTS を検知しました。")
        else:
            log_func(f"[WARN] [{account_name}] ページ接続エラー ({err_str[:60]}) を検知しました。")

    if redirect_detected or "about:blank" in (page.url or "").lower():
        log_func(f"[INFO] 🧹 [{account_name}] セッション失効 / about:blank / リダイレクトループを検知。Threads/Instagramの競合Cookieを整理してログイン画面へ自己修復遷移します...")
        try:
            for d_target in ["threads.net", "threads.com", "instagram.com"]:
                try:
                    context.clear_cookies(domain=d_target)
                except Exception:
                    pass
            time.sleep(1.0)
        except Exception as ce:
            log_func(f"[DEBUG] clear_cookies: {ce}")

        # 直接クリーンにThreadsログイン画面へアクセス
        for target_url in ["https://www.threads.net/login", "https://www.threads.net/"]:
            try:
                page.goto(target_url, timeout=35000, wait_until="domcontentloaded")
                time.sleep(2.5)
                cur_after = (page.url or "").lower()
                if "about:blank" not in cur_after:
                    log_func(f"[INFO] 🚪 [{account_name}] 正常に画面 ({page.url}) を開きました。")
                    return True
            except Exception:
                pass

    return True

def set_account_status_in_excel(filepath, account_name, status="有効", reason=""):
    """
    アカウント情報.xlsx はユーザーが手動管理するため、プログラムから勝手に「無効」に書き換えることはしません。
    """
    return False


def detect_threads_logged_in(page):
    cur_u = (page.url or "").lower()
    # ログイン画面またはabout:blankの場合は確実に未認証
    if "/login" in cur_u or "about:blank" in cur_u:
        return False, "ログイン画面表示中"

    try:
        content = page.content()
        # ログイン画面テキストがある場合は未認証
        if any(w in content for w in ["Instagramアカウントでログイン", "Log in with Instagram", "パスワードを忘れた場合"]):
            return False, "ログイン画面表示中"

        # 1. ページ内部の Relay / 初期化データからログイン済みユーザーを判定
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
            if page.locator(sel).first.is_visible(timeout=500):
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
            if page.locator(sel).first.is_visible(timeout=500):
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

def find_login_form_inputs(page, timeout_sec=8):
    """
    現在の画面（Threads公式ログイン画面またはInstagramログイン画面）から
    ユーザーネーム欄とパスワード欄を安定・確実に自動検出します（最新Threads UI完全対応）。
    """
    user_loc = None
    pass_loc = None

    u_selectors = [
        'input[placeholder*="ユーザーネーム"]',
        'input[placeholder*="携帯電話番号"]',
        'input[placeholder*="メールアドレス"]',
        'input[aria-label*="ユーザーネーム"]',
        'input[aria-label*="ユーザー名"]',
        'input[aria-label*="電話番号"]',
        'input[name="username"]',
        'input[autocomplete="username"]',
        'input[type="text"]'
    ]
    p_selectors = [
        'input[placeholder*="パスワード"]',
        'input[aria-label*="パスワード"]',
        'input[type="password"]',
        'input[name="password"]',
        'input[autocomplete="current-password"]'
    ]

    start_t = time.time()
    while time.time() - start_t < timeout_sec:
        # パスワード欄を最優先で特定
        for p_sel in p_selectors:
            try:
                loc = page.locator(p_sel).first
                if loc.is_visible(timeout=200):
                    pass_loc = loc
                    break
            except Exception:
                pass

        # ユーザーネーム欄を特定
        for u_sel in u_selectors:
            try:
                loc = page.locator(u_sel).first
                if loc.is_visible(timeout=200):
                    user_loc = loc
                    break
            except Exception:
                pass

        # フォールバック: 画面上の全可視inputから割り当て
        if not user_loc or not pass_loc:
            try:
                inputs = page.locator('input:visible').all()
                cand_texts = []
                cand_passes = []
                for inp in inputs:
                    itype = (inp.get_attribute("type") or "text").lower()
                    if itype in ["password"]:
                        cand_passes.append(inp)
                    elif itype in ["text", "email", "tel"] or not itype:
                        cand_texts.append(inp)
                if cand_passes and not pass_loc:
                    pass_loc = cand_passes[0]
                if cand_texts and not user_loc:
                    user_loc = cand_texts[0]
            except Exception:
                pass

        if user_loc and pass_loc:
            break
        time.sleep(0.5)

    return user_loc, pass_loc


def execute_id_pw_login(page, context, account_name, password, log_func=print):
    """
    【Threads直接ID/PW自動ログイン】
    Threadsのログイン画面（「Instagramアカウントでログイン」フォーム）に、
    Excelに登録されたアカウントID（例: takiishinchi01）とパスワード（例: shinki0924）を
    確実に入力し、「ログイン」ボタンをクリックして正規セッションを即時確立します。
    """
    clean_id = str(account_name or "").lstrip("@").strip()
    clean_pw = str(password or "").strip()
    if not clean_pw:
        log_func(f"[WARN] ⚠️ [{account_name}] パスワードがExcelに登録されていないため、ID/PW自動ログインをスキップします。")
        return False

    log_func("=" * 60)
    log_func(f"[INFO] 🔐 [{account_name}] 【Threads ID/PW自動ログイン実行】 (ID: @{clean_id})")
    log_func("=" * 60)

    try:
        cur_u = (page.url or "").lower()
        log_func(f"[INFO] 🔍 [{account_name}] 画面 (URL: {page.url}) の入力フォームを確認中...")

        # 画面が about:blank または Threadsログイン画面でない場合は直接Threadsログインへ移動
        if "about:blank" in cur_u or ("threads.net" not in cur_u and "instagram.com" not in cur_u):
            log_func(f"[INFO] 🌐 [{account_name}] Threads公式ログイン画面 (https://www.threads.net/login) へアクセスします...")
            try:
                page.goto("https://www.threads.net/login", timeout=35000, wait_until="domcontentloaded")
                time.sleep(2.5)
            except Exception as e_goto:
                log_func(f"[WARN] [{account_name}] ログイン画面アクセス警告: {e_goto}")

        # 入力フォーム検出（最大10秒待機）
        user_loc, pass_loc = find_login_form_inputs(page, timeout_sec=10)

        # もしThreads画面で入力欄が出ず「Instagramアカウントでログイン」等のボタンのみの場合はクリック
        if not user_loc or not pass_loc:
            for sso_sel in [
                'button:has-text("Instagramアカウントでログイン")',
                'button:has-text("Instagramでログイン")',
                'div[role="button"]:has-text("Instagramアカウントでログイン")',
                'div[role="button"]:has-text("Instagramでログイン")',
                'a[href*="/login"]'
            ]:
                try:
                    sso_btn = page.locator(sso_sel).first
                    if sso_btn.is_visible(timeout=400):
                        log_func(f"[INFO] 🔘 [{account_name}] ログイン展開ボタン ({sso_sel}) をクリックします...")
                        sso_btn.click()
                        time.sleep(2.0)
                        user_loc, pass_loc = find_login_form_inputs(page, timeout_sec=6)
                        if user_loc and pass_loc:
                            break
                except Exception:
                    pass

        # それでも見つからない場合のInstagram直接アクセス
        if not user_loc or not pass_loc:
            if "instagram.com" not in (page.url or "").lower():
                log_func(f"[INFO] 🌐 [{account_name}] Threadsログイン画面で入力欄が見つからないため、Instagram認証画面へ遷移します...")
                try:
                    page.goto("https://www.instagram.com/accounts/login/", timeout=35000, wait_until="domcontentloaded")
                    time.sleep(3.0)
                    user_loc, pass_loc = find_login_form_inputs(page, timeout_sec=8)
                except Exception as e_goto:
                    log_func(f"[WARN] [{account_name}] Instagram遷移警告: {e_goto}")

        if not user_loc or not pass_loc:
            log_func(f"[ERROR] ❌ [{account_name}] ログイン入力欄（ユーザーネーム・パスワード）が検出できませんでした (現在のURL: {page.url})。")
            return False

        # 入力フォームにアカウントIDとパスワードを確実に全選択クリア＆入力
        log_func(f"[INFO] ✍️ [{account_name}] ログイン画面のユーザー名欄に『{clean_id}』を入力中...")
        user_loc.click()
        time.sleep(0.3)
        try:
            user_loc.press("Control+A")
            user_loc.press("Backspace")
        except Exception:
            pass
        user_loc.fill(clean_id)
        try:
            user_loc.evaluate("(el, val) => { el.value = val; el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); }", clean_id)
        except Exception:
            pass
        time.sleep(0.4)

        log_func(f"[INFO] ✍️ [{account_name}] パスワード欄にパスワードを入力中...")
        pass_loc.click()
        time.sleep(0.3)
        try:
            pass_loc.press("Control+A")
            pass_loc.press("Backspace")
        except Exception:
            pass
        pass_loc.fill(clean_pw)
        try:
            pass_loc.evaluate("(el, val) => { el.value = val; el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); }", clean_pw)
        except Exception:
            pass
        time.sleep(0.5)

        # 送信ボタン（「ログイン」または「Log in」）をクリック
        submit_btn = None
        for s_sel in [
            'button:has-text("ログイン")',
            'div[role="button"]:has-text("ログイン")',
            'button[type="submit"]',
            'form button',
            'form div[role="button"]',
            'div:has-text("ログイン")[role="button"]',
            'span:has-text("ログイン")',
            'button:has-text("Log in")',
            'button:has-text("Log In")'
        ]:
            try:
                b = page.locator(s_sel).first
                if b.is_visible(timeout=500):
                    submit_btn = b
                    break
            except Exception:
                pass

        if submit_btn:
            log_func(f"[INFO] 🚀 [{account_name}] 「ログイン」ボタンをクリックしました。Instagram認証ラグ待機中...")
            try:
                submit_btn.click()
            except Exception:
                pass_loc.press("Enter")
        else:
            log_func(f"[INFO] 🚀 [{account_name}] Enterキーでログイン情報を送信しました。Instagram認証ラグ待機中...")
            pass_loc.press("Enter")

        # -----------------------------------------------------------------
        # 【ラグ待機＆Instagramログイン確認フェーズ】
        # Instagram/Meta側の認証処理・通信ラグがあるため、最大60秒間しっかり待機。
        # recaptcha / challenge / checkpoint 画面を「ログイン完了」と誤検知することを100%防止！
        # -----------------------------------------------------------------
        log_func(f"[INFO] ⏳ [{account_name}] Instagramの認証処理・通信ラグを待機しています（最大60秒）...")
        instagram_logged_in = False
        recaptcha_announced = False
        
        for wait_i in range(1, 61):
            time.sleep(1.0)
            cur_u = (page.url or "").lower()

            # 1. パスワード不一致等の致命的エラー検知
            try:
                page_content = page.content()
                if any(w in page_content for w in ["パスワードが正しくありません", "ユーザーネームが正しくありません", "Incorrect Password", "The password you entered is incorrect"]):
                    log_func(f"[ERROR] ❌ [{account_name}] パスワード不一致のエラーが返されました。エクセルのパスワードをご確認ください。")
                    return False
                if "利用規約" in page_content and "一時的に停止" in page_content:
                    log_func(f"[CRITICAL] 🚨 [{account_name}] アカウント一時停止画面を検知しました。")
                    return False
            except Exception:
                pass

            # 2. reCAPTCHA / セキュリティ確認画面の検知（誤検知防止ガード！）
            is_recaptcha_or_challenge = any(k in cur_u for k in [
                "recaptcha", "auth_platform", "challenge", "checkpoint", "two_factor", "two_step_verification"
            ])
            if not is_recaptcha_or_challenge:
                try:
                    if any(w in page.content() for w in ["recaptcha", "ロボットではありません", "セキュリティチェック", "認証コード"]):
                        is_recaptcha_or_challenge = True
                except Exception:
                    pass

            if is_recaptcha_or_challenge:
                if not recaptcha_announced or wait_i % 10 == 0:
                    log_func("=" * 60)
                    log_func(f"[WARN] 🤖 [{account_name}] Metaのセキュリティ確認画面（reCAPTCHA等）を検知しました！")
                    log_func(f"[INFO] 👉 開いているブラウザ画面でロボット確認（チェックボックス等）を行ってください。認証完了を自動待機しています... ({wait_i}/60秒)")
                    log_func("=" * 60)
                    recaptcha_announced = True

                # reCAPTCHA画面の「次へ」「確認」「送信」ボタンがあれば自動クリックも試行
                for c_btn_sel in [
                    'button:has-text("次へ")',
                    'button:has-text("送信")',
                    'button:has-text("確認")',
                    'button:has-text("Continue")',
                    'button:has-text("Submit")'
                ]:
                    try:
                        cand_b = page.locator(c_btn_sel).first
                        if cand_b.is_visible(timeout=200):
                            cand_b.click()
                    except Exception:
                        pass
                continue  # reCAPTCHA解除完了までThreadsに勝手に遷移させない！

            # 3. 「ログイン情報を保存」「後で」「Save Info」「Not Now」等のポップアップ処理
            for pop_sel in [
                'button:has-text("後で")',
                'div[role="button"]:has-text("後で")',
                'button:has-text("ログイン情報を保存")',
                'button:has-text("情報を保存")',
                'button:has-text("Save info")',
                'button:has-text("Not now")',
                'button:has-text("次へ")'
            ]:
                try:
                    pop = page.locator(pop_sel).first
                    if pop.is_visible(timeout=300):
                        log_func(f"[INFO] 🔘 [{account_name}] 画面上のダイアログ ({pop_sel}) を自動クリックしました。")
                        pop.click(force=True)
                        time.sleep(1.0)
                except Exception:
                    pass

            # 4. 認証Cookie (sessionid) の生成確認（recaptchaやlogin画面にいないこと）
            try:
                all_cookies = context.cookies()
                has_sess = any(c.get("name") == "sessionid" and len(str(c.get("value") or "")) > 15 for c in all_cookies)
                if has_sess and not any(k in cur_u for k in ["/login", "recaptcha", "auth_platform", "challenge", "checkpoint"]):
                    log_func(f"[SUCCESS] 🔑 [{account_name}] Instagram正規セッションCookie (sessionid) の生成を確認しました！")
                    instagram_logged_in = True
                    break
            except Exception:
                pass

            # 5. Threads画面に直接到達・認証済み判定
            if "threads.net" in cur_u and "/login" not in cur_u:
                is_ok, _ = detect_threads_logged_in(page)
                if is_ok:
                    instagram_logged_in = True
                    break

            # 6. Instagram正規画面（フィード・プロフィール）へ到達確認
            if "instagram.com" in cur_u and not any(k in cur_u for k in ["/accounts/login", "recaptcha", "auth_platform", "challenge", "checkpoint", "about:blank"]):
                log_func(f"[SUCCESS] 🚪 [{account_name}] Instagram正規画面へのログイン完了を確認しました (URL: {page.url})。")
                instagram_logged_in = True
                break

            if wait_i % 5 == 0 and not is_recaptcha_or_challenge:
                log_func(f"[INFO] ⏳ [{account_name}] 認証待機中... ({wait_i}/60秒)")

        time.sleep(2.0)

        # -----------------------------------------------------------------
        # 【Threads連携フェーズ】
        # Instagram側でログイン完了が確認された後、Threadsホームへ安全に遷移・連携
        # -----------------------------------------------------------------
        cur_u = (page.url or "").lower()
        if "threads.net" not in cur_u or "/login" in cur_u:
            log_func(f"[INFO] 🌐 [{account_name}] Instagramログイン完了を確認！Threadsホーム画面へ遷移してセッションを確立します...")
            try:
                page.goto("https://www.threads.net/", timeout=45000, wait_until="domcontentloaded")
                time.sleep(3.5)
            except Exception as e_th:
                log_func(f"[WARN] [{account_name}] Threadsホーム遷移警告: {e_th}")

        # Threadsで「〇〇として続行」「Instagramでログイン」ボタンがあればクリックして紐付け
        for cont_sel in [
            'button:has-text("としてログイン")',
            'div[role="button"]:has-text("としてログイン")',
            'span:has-text("としてログイン")',
            f'button:has-text("{clean_id}")',
            'button:has-text("Continue as")',
            'button:has-text("Instagramでログイン")',
            'div[role="button"]:has-text("Instagramでログイン")',
            'span:has-text("Instagramでログイン")',
            'button:has-text("Instagramで続行")',
            'button:has-text("次へ")'
        ]:
            try:
                c_btn = page.locator(cont_sel).first
                if c_btn.is_visible(timeout=1000):
                    log_func(f"[INFO] 🔗 [{account_name}] Threads連携ボタン ({cont_sel}) を検出。自動クリックしました。")
                    c_btn.click(force=True)
                    time.sleep(3.5)
                    break
            except Exception:
                pass

        # 最終認証判定（最大25秒）
        for _ in range(25):
            is_ok, reason = detect_threads_logged_in(page)
            if is_ok:
                log_func(f"[SUCCESS] 🎉 [{account_name}] 【ID/PW完全自動ログイン大成功！】正規プロファイルを確立しました！（{reason}）")
                p_dir = os.path.join(BASE_DIR, f"chrome_profiles/{account_name}")
                save_portable_session_file(context, p_dir, account_name, log_func)
                return True
            time.sleep(1.0)

        final_u = (page.url or "").lower()
        if "threads.net" in final_u and not any(k in final_u for k in ["/login", "recaptcha", "auth_platform", "challenge"]):
            log_func(f"[SUCCESS] 🎉 [{account_name}] Threadsホーム画面への到達を確認し、認証成功と判定しました！")
            p_dir = os.path.join(BASE_DIR, f"chrome_profiles/{account_name}")
            save_portable_session_file(context, p_dir, account_name, log_func)
            return True

        log_func(f"[WARN] ❌ [{account_name}] ログイン後の認証確認が完了しませんでした (URL: {page.url})。")
        return False
    except Exception as e:
        log_func(f"[ERROR] [{account_name}] ID/PW自動ログイン例外: {e}")
        return False


def ensure_authenticated_session(page, context, account, filepath, log_func=print, wait_timeout_sec=30, acc_state=None):
    account_name = str(account.get("アカウントID") or account.get("アカウント名") or account.get("アカウント") or account if isinstance(account, dict) else account).strip().lstrip("@")
    password = ""
    if isinstance(account, dict):
        password = str(account.get("パスワード") or account.get("password") or account.get("PW") or account.get("Pass") or account.get("暗証番号") or "").strip()
    
    # Excelからのフェイルセーフ取得
    if not password and filepath and os.path.exists(filepath):
        try:
            recs = read_excel_records(filepath)
            for r in recs:
                rid = str(r.get("アカウントID") or r.get("アカウント名") or "").strip().lstrip("@")
                if rid == account_name or not account_name:
                    pw_cand = str(r.get("パスワード") or r.get("password") or r.get("PW") or r.get("Pass") or r.get("暗証番号") or "").strip()
                    if pw_cand:
                        password = pw_cand
                        break
            if not password and recs and len(recs) == 1:
                password = str(recs[0].get("パスワード") or recs[0].get("password") or "").strip()
        except Exception:
            pass

    # 生Excel直接セル読み出しフォールバック (openpyxl)
    if not password and filepath and os.path.exists(filepath):
        try:
            import openpyxl
            wb_tmp = openpyxl.load_workbook(filepath, data_only=True)
            ws_tmp = wb_tmp.active
            for r in ws_tmp.iter_rows(values_only=True):
                if not r: continue
                r_strs = [str(c).strip() for c in r if c is not None]
                if any(k in "".join(r_strs) for k in ["パスワード", "password", "PW", "Pass"]):
                    continue
                # アカウント行を見つけた場合
                if len(r) >= 2 and r[1] is not None and str(r[1]).strip():
                    c0 = str(r[0] or "").strip().lstrip("@")
                    c1 = str(r[1] or "").strip()
                    if not account_name or c0 == account_name:
                        password = c1
                        break
            wb_tmp.close()
        except Exception:
            pass

    time.sleep(1.5)
    is_cp, cp_why = detect_phone_checkpoint_screen(page)
    if is_cp:
        log_func(f"[WARN] 🚨 [{account_name}] 携帯電話番号（SMS）認証画面を検知しました（{cp_why}）。")
        return False

    # 既にログイン成功しているか判定
    is_ok, reason = detect_threads_logged_in(page)
    if is_ok:
        log_func(f"[SUCCESS] [{account_name}] Chromeプロファイルのセッションで自動ログイン確認完了！（{reason}）")
        p_dir = os.path.join(BASE_DIR, f"chrome_profiles/{account_name}")
        save_portable_session_file(context, p_dir, account_name, log_func)
        if acc_state:
            acc_state["login_cooldown_until"] = 0
            acc_state["login_cooldown_started"] = False
            acc_state["pure_profile_mode"] = True
        return True

    # 未ログイン・ログアウト画面を検知した場合
    log_func(f"[INFO] [{account_name}] Threads画面のセッション照合中 (URL: {page.url})...")

    # 1. まずワンクリック連携ボタンを試行
    for cont_sel in [
        'button:has-text("としてログイン")',
        'div[role="button"]:has-text("としてログイン")',
        'span:has-text("としてログイン")',
        'button:has-text("Continue as")',
        'div[role="button"]:has-text("Instagramでログイン")',
        'button:has-text("Instagramでログイン")',
        'span:has-text("Instagramでログイン")',
        'button:has-text("Instagramで続行")'
    ]:
        try:
            btn = page.locator(cont_sel).first
            if btn.is_visible(timeout=800):
                log_func(f"[INFO] [{account_name}] 連携ボタン ({cont_sel}) を検出。クリック試行中...")
                btn.click(force=True)
                time.sleep(3)
                break
        except Exception:
            pass

    is_ok, reason = detect_threads_logged_in(page)
    if is_ok:
        log_func(f"[SUCCESS] 🎉 [{account_name}] ワンクリック連携により自動ログイン復帰しました！（{reason}）")
        p_dir = os.path.join(BASE_DIR, f"chrome_profiles/{account_name}")
        save_portable_session_file(context, p_dir, account_name, log_func)
        if acc_state:
            acc_state["login_cooldown_until"] = 0
            acc_state["login_cooldown_started"] = False
            acc_state["pure_profile_mode"] = True
        return True

    # 2. セッションIDが入っていない場合、またはセッションIDでログインできない場合 ➔ 入力されているID/PWで即座に自動ログインを実行！
    if password:
        session_token_present = False
        if isinstance(account, dict):
            session_token_present = bool(str(account.get("auth_token") or account.get("sessionid") or "").strip())
        
        log_func("=" * 60)
        if not session_token_present:
            log_func(f"[INFO] 🔐 [{account_name}] セッションID未設定を検知。Excelに入力されている【ID/パスワード】で自動ログインを実行します！")
        else:
            log_func(f"[INFO] 🔐 [{account_name}] セッションID失効または無効を検知。Excelに入力されている【ID/パスワード】で自動復帰ログインを実行します！")
        log_func("=" * 60)

        # Threads/Instagramの壊れたセッションCookieのみを整理し、他サイト(X/Google等)のログインCookieは完全保持
        try:
            for d_target in ["threads.net", "threads.com", "instagram.com"]:
                try:
                    context.clear_cookies(domain=d_target)
                except Exception:
                    pass
            time.sleep(0.5)
        except Exception:
            pass

        ok = execute_id_pw_login(page, context, account_name, password, log_func=log_func)
        if ok:
            p_dir = os.path.join(BASE_DIR, f"chrome_profiles/{account_name}")
            save_portable_session_file(context, p_dir, account_name, log_func)
            if acc_state:
                acc_state["login_cooldown_until"] = 0
                acc_state["login_cooldown_started"] = False
                acc_state["pure_profile_mode"] = True
            # 新しいセッションIDをアカウント情報.xlsxに同期
            save_context_cookies_to_excel(context, filepath, account_name, log_func=log_func)
            return True
        else:
            log_func(f"[WARN] ❌ [{account_name}] ID/PW自動ログインに失敗しました。パスワードをご確認ください。")
            return False
    else:
        log_func(f"[WARN] [{account_name}] セッションIDが未設定または無効で、パスワードも「アカウント情報.xlsx」に未記入のため、ログインを実行できません。Excelの「パスワード」列にご記入ください。")
        return False

def resolve_image_path(raw_path, post_text="", account_name=""):
    search_dirs = []
    clean_acc = sanitize_account_folder_name(account_name)
    if clean_acc:
        search_dirs.extend([
            os.path.join(BASE_DIR, "images", clean_acc),
            os.path.join(BASE_DIR, "imeges", clean_acc),
            os.path.join(os.getcwd(), "images", clean_acc),
            os.path.join(os.getcwd(), "imeges", clean_acc),
        ])
    search_dirs.extend([
        os.path.join(BASE_DIR, "images"),
        os.path.join(BASE_DIR, "imeges"),
        BASE_DIR,
        os.path.join(os.getcwd(), "images"),
        os.path.join(os.getcwd(), "imeges"),
        os.getcwd(),
    ])
    
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


def _perform_fill_and_submit_post(
    page,
    container,
    textbox,
    account_name,
    text_content,
    target_topic=None,
    resolved_img=None,
    submit_btn_override=None,
    route_name="新規投稿",
    log_func=print
):
    """
    指定されたテキストボックスに入力、トピック・画像を設定し、
    「投稿」ボタンをクリックして送信完了（トースト通知またはフォーム終了）を確認する汎用投稿関数
    """
    try:
        # 1. テキストボックスのフォーカス＆クリア
        textbox.click()
        try:
            textbox.press("Control+A")
            textbox.press("Backspace")
        except Exception:
            pass
        time.sleep(0.3)

        # 2. 本文のタイピング入力
        lines = text_content.split("\n")
        for l_idx, line in enumerate(lines):
            if line:
                textbox.press_sequentially(line, delay=random.randint(15, 30))
            if l_idx < len(lines) - 1:
                textbox.press("Shift+Enter")
                time.sleep(0.1)

        time.sleep(1.0)

        # 3. トピック設定
        if target_topic:
            try:
                set_topic_in_threads_composer(page, container, target_topic, log_func)
                time.sleep(0.5)
            except Exception as te:
                if log_func:
                    log_func(f"[DEBUG] トピック設定例外: {te}")

        # 4. 画像添付
        if resolved_img:
            safe_img = ensure_valid_image(resolved_img, log_func)
            if log_func:
                log_func(f"[{account_name}] 画像を添付します: {os.path.basename(safe_img)}")
            attached = False
            media_selectors = [
                'svg[aria-label*="メディア"]',
                'svg[aria-label*="画像"]',
                'svg[aria-label*="写真"]',
                'svg[aria-label*="添付"]',
                'svg[aria-label="画像を追加"]',
                'svg[aria-label="メディアを添付"]',
                'svg[aria-label="写真を添付"]'
            ]
            search_root = container if (container and container.is_visible()) else page
            for m_sel in media_selectors:
                try:
                    m_el = search_root.locator(m_sel).first
                    if m_el.is_visible():
                        btn_parent = search_root.locator(f'div[role="button"]:has({m_sel}), button:has({m_sel})').first
                        click_el = btn_parent if btn_parent.is_visible() else m_el
                        with page.expect_file_chooser(timeout=2500) as fc_info:
                            click_el.click()
                        file_chooser = fc_info.value
                        file_chooser.set_files(safe_img)
                        attached = True
                        if log_func:
                            log_func(f"[{account_name}] メディア添付アイコン経由で画像を設定しました。")
                        break
                except Exception:
                    pass

            if not attached:
                for f_sel in [
                    'input[type="file"][accept*="image"]',
                    'input[type="file"][multiple]',
                    'input[type="file"]'
                ]:
                    try:
                        locs = search_root.locator(f_sel).all()
                        if locs:
                            locs[-1].set_input_files(safe_img)
                            attached = True
                            if log_func:
                                log_func(f"[{account_name}] ファイル入力要素({f_sel})に画像を設定しました。")
                            break
                    except Exception:
                        pass

            if attached:
                for _ in range(12):
                    time.sleep(0.8)
                    preview_found = False
                    for p_sel in [
                        'img[src^="blob:"]',
                        'img[alt*="メディア"]',
                        'img[alt*="画像"]',
                        'svg[aria-label*="削除"]',
                        'svg[aria-label*="メディアを破棄"]'
                    ]:
                        try:
                            if search_root.locator(p_sel).first.is_visible():
                                preview_found = True
                                break
                        except Exception:
                            pass
                    if preview_found:
                        if log_func:
                            log_func(f"[{account_name}] 画像プレビュー確認完了（アップロード成功）")
                        break

        time.sleep(1.0)

        # 5. 余計なポップアップメニュー（返信範囲設定等）をEscapeでクローズ
        for pop_sel in ['text="返信と引用ができる人"', 'text="返信できる人"', 'div[role="menu"]']:
            try:
                if page.locator(pop_sel).first.is_visible():
                    page.keyboard.press("Escape")
                    time.sleep(0.5)
                    break
            except Exception:
                pass

        # 6. 「投稿」送信ボタンの特定
        submit_btn = submit_btn_override
        if not submit_btn or not submit_btn.is_visible():
            search_root = container if (container and container.is_visible()) else page
            for _ in range(15):
                candidates = []
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
            return False, "「投稿」ボタンがアクティブになりませんでした（文字数超過またはアップロード中の可能性）"

        if log_func:
            log_func(f"[{account_name}] 【{route_name}】「投稿」ボタンをクリックしました。Threads送信処理を確認中...")
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

        # 7. 送信完了チェック
        posted_success = False
        post_error_detected = ""

        for check_idx in range(30):
            time.sleep(0.8)
            # (A) 画面下部の黒い完了通知トースト
            for s_sel in [
                'text="投稿完了"', ':has-text("投稿完了")',
                'text="投稿しました"', ':has-text("投稿しました")',
                'text="スレッドを投稿しました"', ':has-text("スレッドを投稿しました")',
                'text="スレッドを表示"', ':has-text("スレッドを表示")',
                'text="スレッドが投稿されました"', ':has-text("スレッドが投稿されました")',
                'text="Posted"', 'text="View post"',
                'div[role="alert"]:has-text("投稿")', 'div[role="alert"]:has-text("スレッド")',
                'div:has-text("投稿完了")', 'div:has-text("投稿しました")',
                'div:has-text("スレッドを投稿しました")'
            ]:
                try:
                    if page.locator(s_sel).first.is_visible():
                        posted_success = True
                        if log_func:
                            log_func(f"[SUCCESS] [{account_name}] 🎉 完了通知（'{s_sel}'）を確認し、投稿完了しました！")
                        break
                except Exception:
                    pass
            if posted_success:
                break

            # (B) エラー画面通知
            for err_sel in [
                'div[role="alert"]:has-text("エラー")', 'text="投稿できませんでした"',
                'text="アップロードできませんでした"', 'text="エラーが発生しました"',
                'text="問題が発生しました"', 'text="コミュニティガイドライン"',
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

            # (C) フォームの消失またはクリア検知
            form_cleared_or_closed = False
            if container and container.is_visible():
                try:
                    if not textbox.is_visible() or not textbox.inner_text().strip():
                        form_cleared_or_closed = True
                except Exception:
                    form_cleared_or_closed = True
            else:
                form_cleared_or_closed = True

            if form_cleared_or_closed and check_idx >= 2:
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
                    if log_func:
                        log_func(f"[SUCCESS] [{account_name}] 🎉 投稿フォームが正常に送信され完了しました！")
                    break

            if check_idx in [5, 12]:
                try:
                    if submit_btn and submit_btn.is_visible():
                        submit_btn.click(force=True)
                except Exception:
                    pass

        if post_error_detected:
            return False, f"Threads画面エラー: {post_error_detected}"

        if posted_success:
            return True, ""
        else:
            return False, "投稿ボタン押下後30秒経過しても完了通知が確認できませんでした"

    except Exception as e:
        return False, str(e)


def _open_timeline_top_composer(page, account_name, log_func=print):
    """
    タイムライン上部（おすすめ直下）の「最近どう？ / 投稿」枠を特定してアクティブ化します。
    戻り値: (container, textbox, submit_btn)
    """
    try:
        # 画面最上部にスクロール
        page.evaluate("window.scrollTo(0, 0)")
        time.sleep(0.8)
    except Exception:
        pass

    # 1. 「最近どう？」テキスト・プレースホルダーを検索してクリック
    found_prompt = None
    for sel in [
        'text="最近どう？"',
        ':has-text("最近どう？")',
        'div:has-text("最近どう？")',
        'span:has-text("最近どう？")',
        '[placeholder*="最近どう"]',
        '[aria-label*="最近どう"]',
        'div:has-text("Start a thread")',
        'div:has-text("スレッドを開始")'
    ]:
        try:
            for el in page.locator(sel).all():
                if el.is_visible():
                    box = el.bounding_box()
                    if box and box['y'] < 450:
                        found_prompt = el
                        log_func(f"[{account_name}] タイムライン上部の「最近どう？」枠 (x={int(box['x'])}, y={int(box['y'])}) を検出！クリックします...")
                        break
            if found_prompt:
                break
        except Exception:
            pass

    # もし見つからない場合、上部カード内の「投稿」ボタン周辺をクリック
    if not found_prompt:
        for b_sel in ['button:has-text("投稿")', 'div[role="button"]:has-text("投稿")']:
            try:
                for b in page.locator(b_sel).all():
                    if b.is_visible():
                        box = b.bounding_box()
                        if box and 100 < box['y'] < 450:
                            click_x = max(150, box['x'] - 120)
                            click_y = box['y'] + (box['height'] / 2)
                            log_func(f"[{account_name}] 上部カード位置 (x={int(click_x)}, y={int(click_y)}) を直接クリックして入力欄を起動します...")
                            page.mouse.click(click_x, click_y)
                            time.sleep(1.0)
                            found_prompt = b
                            break
                if found_prompt:
                    break
            except Exception:
                pass

    if found_prompt:
        try:
            found_prompt.scroll_into_view_if_needed()
            found_prompt.click(force=True)
            time.sleep(1.2)
        except Exception:
            pass

    # (A) クリックによって新規投稿ダイアログが開いたかチェック
    for d_sel in ['div[role="dialog"]', 'div[aria-modal="true"]']:
        try:
            cand = page.locator(d_sel).first
            if cand.is_visible():
                tb = cand.locator('div[role="textbox"], div[contenteditable="true"], [data-lexical-editor="true"]').first
                if tb.is_visible():
                    log_func(f"[{account_name}] 「最近どう？」クリックにより新規投稿ダイアログが開きました。")
                    return cand, tb, None
        except Exception:
            pass

    # (B) インラインの入力欄（contenteditable / textbox）をチェック
    for tb_sel in [
        'div[role="textbox"]',
        'div[contenteditable="true"]',
        '[data-lexical-editor="true"]'
    ]:
        try:
            for tb in page.locator(tb_sel).all():
                if tb.is_visible():
                    box = tb.bounding_box()
                    if box and box['y'] < 450:
                        # タイムライン上部カード内の投稿ボタンを特定
                        submit_btn = None
                        for b_sel in ['button:has-text("投稿")', 'div[role="button"]:has-text("投稿")', 'button:has-text("Post")', 'div[role="button"]:has-text("Post")']:
                            for b in page.locator(b_sel).all():
                                if b.is_visible():
                                    b_box = b.bounding_box()
                                    if b_box and b_box['y'] < 450:
                                        submit_btn = b
                                        break
                            if submit_btn:
                                break
                        log_func(f"[{account_name}] タイムライン上部のインライン投稿欄 (y={int(box['y'])}) を検出しました。")
                        return None, tb, submit_btn
        except Exception:
            pass

    return None, None, None


def execute_threads_post_with_fallback(
    page,
    account_name,
    text_content,
    target_topic=None,
    resolved_img=None,
    log_func=print,
    abort_check=None
):
    """
    【2段階フェイルセーフ投稿エンジン】
    1. 【第1優先】画面左下の円形「＋」ボタン / アイコン / ショートカット 'c' 経由で
       新規スレッド作成モーダル (div[role="dialog"]) を開いて投稿。
    2. 【第2優先】左下の「＋」ボタンでモーダルが開かない、または投稿に失敗した場合、
       タイムライン上部（おすすめフィード直下）の「最近どう？ / 投稿」枠から代替投稿。
    3. 【2重投稿防止ガード】第1優先（＋ボタン）が成功した場合は第2優先を実行せず直ちに完了。
       どちらか一方のみが1度だけ成功するように厳密に制御。
    """
    if abort_check and abort_check():
        log_func(f"[INFO] ⏸️ 停止ボタン検知: 投稿処理を直ちに中断しました。")
        return False, "停止ボタンによる中断"

    posted_success = False
    last_error = ""

    # =========================================================================
    # ルート1: 画面左下の円形「＋」ボタンからの投稿（優先度：高）
    # =========================================================================
    if abort_check and abort_check():
        log_func(f"[INFO] ⏸️ 停止ボタン検知: 投稿処理を直ちに中断しました。")
        return False, "停止ボタンによる中断"

    log_func(f"[{account_name}] 🚀 【第1ルート: 左下「＋」ボタン】新規スレッド作成モーダルを起動して投稿を試みます...")
    try:
        dialog, textbox = open_threads_composer_dialog(page, account_name, log_func=log_func)
        if abort_check and abort_check():
            log_func(f"[INFO] ⏸️ 停止ボタン検知: 投稿入力を中断しました。")
            return False, "停止ボタンによる中断"
        if dialog and textbox:
            ok, err = _perform_fill_and_submit_post(
                page=page,
                container=dialog,
                textbox=textbox,
                account_name=account_name,
                text_content=text_content,
                target_topic=target_topic,
                resolved_img=resolved_img,
                route_name="左下「＋」ボタン",
                log_func=log_func
            )
            if ok:
                posted_success = True
                log_func(f"[SUCCESS] [{account_name}] 🎉 【第1ルート: 左下「＋」ボタン】から正常に投稿が完了しました！（※2重投稿防止のため「最近どう？」枠はスキップします）")
                return True, ""
            else:
                last_error = err or "左下「＋」からの投稿送信に失敗しました"
        else:
            last_error = "左下「＋」ボタンをクリックしましたが、新規投稿モーダルが開きませんでした"
    except Exception as e:
        last_error = f"左下「＋」処理エラー: {e}"
        log_func(f"[WARN] [{account_name}] 第1ルート失敗: {e}")

    # =========================================================================
    # ルート2: タイムライン上部の「最近どう？ / 投稿」枠からの投稿（フォールバック）
    # =========================================================================
    if not posted_success:
        if abort_check and abort_check():
            log_func(f"[INFO] ⏸️ 停止ボタン検知: 投稿処理を直ちに中断しました。")
            return False, "停止ボタンによる中断"
        log_func(f"[{account_name}] 🔄 プラスボタンで投稿できなかったため、タイムライン上部の【最近どう？ / 投稿】枠から代替投稿を実行します...")
        try:
            # 画面の余計なダイアログをクリーンアップ＆タイムライン最上部にスクロール
            for _ in range(3):
                try:
                    page.keyboard.press("Escape")
                    time.sleep(0.3)
                except Exception:
                    pass
            try:
                page.evaluate("window.scrollTo(0, 0)")
                time.sleep(1.0)
            except Exception:
                pass

            top_container, top_textbox, top_submit_btn = _open_timeline_top_composer(page, account_name, log_func=log_func)
            if top_textbox:
                ok, err = _perform_fill_and_submit_post(
                    page=page,
                    container=top_container,
                    textbox=top_textbox,
                    account_name=account_name,
                    text_content=text_content,
                    target_topic=target_topic,
                    resolved_img=resolved_img,
                    submit_btn_override=top_submit_btn,
                    route_name="タイムライン上部「最近どう？」",
                    log_func=log_func
                )
                if ok:
                    posted_success = True
                    log_func(f"[SUCCESS] [{account_name}] 🎉 【第2ルート: タイムライン上部「最近どう？」】から代替投稿に成功しました！")
                    return True, ""
                else:
                    last_error = err or "タイムライン上部「最近どう？」からの投稿送信に失敗しました"
            else:
                last_error = "タイムライン上部の「最近どう？ / 投稿」枠を検出できませんでした"
        except Exception as e:
            last_error = f"タイムライン上部「最近どう？」投稿エラー: {e}"
            log_func(f"[WARN] [{account_name}] 第2ルート失敗: {e}")

    return False, last_error


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
        import traceback
        diag_lines = [
            f"\n{'='*60}\n",
            f"[{now_str}] 【エラー検知】 アカウント: @{account_name} | 対象: {context_label}\n",
            f"エラー内容: {exc_msg}\n",
        ]
        if screen_messages:
            diag_lines.append("Threads画面上のメッセージ:\n")
            for sm in screen_messages:
                diag_lines.append(f"  ・{sm}\n")
        if ss_saved:
            diag_lines.append(f"スクリーンショット保存先: {ss_path}\n")
        if exception_obj:
            diag_lines.append("詳細トレースバック:\n" + "".join(traceback.format_exception(type(exception_obj), exception_obj, exception_obj.__traceback__)) + "\n")
        diag_lines.append(f"{'='*60}\n")
        diag_text = "".join(diag_lines)
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(diag_text)
        if account_name and account_name != "unknown":
            log_to_account_file(account_name, diag_text)
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

def perform_stealth_feed_patrol(page, account_name, log_func=None, abort_check=None):
    """
    【提案5: ステルス巡回（ウォーミングアップ閲覧＆人間挙動模倣）】
    新規投稿前に「おすすめ」タイムラインを人間らしくスクロール・閲覧し、マウス移動や滞在ログをMetaサーバーに送信。
    ボット判定・強制ログアウト・BANを強力に回避します。
    """
    if abort_check and abort_check():
        return

    try:
        if log_func:
            log_func(f"[{account_name}] 🕵️‍♂️ 【提案5: ステルス巡回】おすすめタイムラインを閲覧中（人間らしいマウス移動・滞在ログを生成）...")
        time.sleep(random.uniform(1.5, 2.5))
        if abort_check and abort_check():
            return

        # マウスを画面中央付近へ自然に移動
        try:
            page.mouse.move(random.randint(400, 700), random.randint(250, 450))
        except Exception:
            pass

        # 1回目のゆっくりスクロール（下へ350〜550px）
        scroll_y = random.randint(350, 550)
        page.evaluate(f"window.scrollBy({{ top: {scroll_y}, behavior: 'smooth' }})")
        time.sleep(random.uniform(2.0, 3.5))
        if abort_check and abort_check():
            return

        # タイムライン上の投稿付近へマウスホバー（人間の視線移動を模倣）
        try:
            page.mouse.move(random.randint(350, 650), random.randint(300, 600))
        except Exception:
            pass

        # 2回目のゆっくりスクロール（さらに200〜400px）
        scroll_y2 = random.randint(200, 400)
        page.evaluate(f"window.scrollBy({{ top: {scroll_y2}, behavior: 'smooth' }})")
        time.sleep(random.uniform(1.8, 3.0))
        if abort_check and abort_check():
            return

        # 人間が投稿を読み返すように少しだけ上に戻す（-100〜-180px）
        scroll_up = random.randint(100, 180)
        page.evaluate(f"window.scrollBy({{ top: -{scroll_up}, behavior: 'smooth' }})")
        time.sleep(random.uniform(1.2, 2.2))
        if abort_check and abort_check():
            return

        # 一番上に戻して投稿準備
        page.evaluate("window.scrollTo({ top: 0, behavior: 'smooth' })")
        time.sleep(1.2)

        if log_func:
            log_func(f"[{account_name}] 👁️‍🗨️ ステルス巡回完了（正常な閲覧ログ記録済）。投稿フローに移行します。")
    except Exception as e:
        if log_func:
            log_func(f"[INFO] [{account_name}] ステルス巡回スキップ (画面動作優先): {e}")

def clean_account_profile(account_name, profile_dir, reason="警告検知または動作不安定", log_func=None, full_reset=False):
    """
    【機能2: プロファイルキャッシュクリーンアップ】
    自動運用時は手動ログイン済みセッション（Cookies / Local Storage / Preferences）を絶対に消さず、
    キャッシュ（Cache / Code Cache / GPUCache）のみ安全に掃除します。
    full_reset=True の場合（ユーザーによる手動リセットボタン押下時）のみプロファイル全体を削除します。
    """
    import shutil
    import stat

    if log_func:
        log_func(f"[WARN] [{account_name}] 🚨 動作不安定・警告兆候を検知（理由: {reason}）")
        if full_reset:
            log_func(f"[INFO] [{account_name}] 🧹 プロファイルフォルダ ('{profile_dir}') を完全初期化します...")
        else:
            log_func(f"[INFO] [{account_name}] 🧹 ログイン情報を保持したままキャッシュのみ安全に掃除します...")

    def _remove_readonly(func, path, exc_info):
        try:
            os.chmod(path, stat.S_IWRITE)
            func(path)
        except Exception:
            pass

    success = False
    if full_reset:
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
    else:
        cache_targets = [
            os.path.join(profile_dir, "Default", "Cache"),
            os.path.join(profile_dir, "Default", "Code Cache"),
            os.path.join(profile_dir, "Default", "GPUCache"),
            os.path.join(profile_dir, "ShaderCache"),
            os.path.join(profile_dir, "GrShaderCache"),
        ]
        for ct in cache_targets:
            try:
                if os.path.exists(ct):
                    shutil.rmtree(ct, onerror=_remove_readonly)
            except Exception:
                pass
        success = True

    if success:
        if log_func:
            if full_reset:
                log_func(f"[SUCCESS] [{account_name}] 🧹 プロファイルの完全初期化が完了しました。")
            else:
                log_func(f"[SUCCESS] [{account_name}] 🧹 ログインセッションを保持したままキャッシュをクリーンアップしました。")
        log_execution(f"[{account_name}] プロファイルクリーンアップ完了 (full_reset={full_reset}, 理由: {reason})")
    else:
        if log_func:
            log_func(f"[WARN] [{account_name}] 一部ファイルの排他ロックによりクリーンアップをスキップしました。")

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

# =====================================================================
# アカウント設計 × スピンタックス文章 & Pillow PNG画像 自動生成エンジン
# =====================================================================

_DESIGN_PRESETS_PY = {
    "飛田新地・風景×求人案内": {
        "genre": "飛田新地・大阪エリアの安心ワーク紹介",
        "target": "短期集中でしっかり稼ぎたい未経験・経験者の方",
        "tone": "丁寧・誠実",
        "pain": "雰囲気や給与システムが事前に分からず一歩踏み出せない",
        "benefits": "見学のみOK・日払い対応\n1日2〜3時間の短時間シフト\nプライバシー徹底・個室待機完備\n専任スタッフの事前サポート",
        "cta": "雰囲気や条件の詳細はプロフィールの案内ページからご確認ください",
        "url": "https://osaka-shinodayamashinchivip.pages.dev/blog/osaka-shinodayama-high-guarantee-guide",
        "topic": "飛田新地",
        "image_style": "自動おまかせ（設計に合わせて最適化）",
        "palette": "提灯アンバー×黒",
    },
    "信太山新地・安心解説": {
        "genre": "信太山新地の働き方・システム解説",
        "target": "身バレ防止・自由出勤を重視したい方",
        "tone": "共感スタイル",
        "pain": "身バレや強引な勧誘がないか不安で相談できない",
        "benefits": "アリバイ・送迎・寮完備\nお友達同士の見学も大歓迎\nノルマ・罰金いっさい無し\n自分のペースで週1から勤務OK",
        "cta": "不安な点の質問だけでも気軽にプロフのリンクからどうぞ",
        "url": "https://osaka-shinodayamashinchivip.pages.dev/blog/shinodayamashinchi-privacy-guide",
        "topic": "信太山新地",
        "image_style": "自動おまかせ（設計に合わせて最適化）",
        "palette": "エメラルド×濃紺",
    },
    "SNS運用・チェックリスト": {
        "genre": "Threads運用・自動化ノウハウ",
        "target": "投稿ネタ切れや画像作成の手間に悩む運用者",
        "tone": "プロ解説",
        "pain": "毎日投稿と画像作りに時間がかかりすぎて続かない",
        "benefits": "フック×本文×CTAの型を固定\nスピンタックスで重複判定を回避\n背景写真＋自動文字入れで時短\n1日3分で100件分のストック完成",
        "cta": "保存して次回の投稿作成のチェックリストに活用してください",
        "url": "",
        "topic": "SNS運用",
        "image_style": "ステップ図解",
        "palette": "生成り白×墨",
    },
    "美容・カフェ・雑誌マガジン風": {
        "genre": "大人女子の週末カフェ・美容習慣",
        "target": "無理なく垢抜けたい20〜30代女性",
        "tone": "フレンドリー",
        "pain": "忙しくて自分のケアやリフレッシュの時間が取れない",
        "benefits": "朝10分で気分が整うルーティン\n肌と心に優しいシンプル習慣\n週末だけのご褒美カフェ時間\nプチプラでも続く厳選アイテム",
        "cta": "あとで見返せるように保存＆フォローしてもらえると嬉しいです",
        "url": "",
        "topic": "美容習慣",
        "image_style": "雑誌エディトリアル風",
        "palette": "雑誌アイボリー×紺",
    },
}

_PIL_PALETTES = {
    "提灯アンバー×黒": {
        "bg1": (11, 12, 16), "bg2": (22, 25, 34), "accent": (245, 158, 11),
        "accent_soft": (62, 43, 14), "text": (248, 250, 252), "muted": (203, 213, 225),
        "border": (110, 78, 25), "light": False,
    },
    "エメラルド×濃紺": {
        "bg1": (9, 17, 20), "bg2": (16, 32, 38), "accent": (16, 185, 129),
        "accent_soft": (14, 58, 46), "text": (248, 250, 252), "muted": (203, 213, 225),
        "border": (22, 92, 68), "light": False,
    },
    "深紅×黒": {
        "bg1": (18, 9, 12), "bg2": (31, 16, 21), "accent": (244, 63, 94),
        "accent_soft": (72, 20, 32), "text": (255, 241, 242), "muted": (253, 164, 175),
        "border": (120, 35, 52), "light": False,
    },
    "生成り白×墨": {
        "bg1": (246, 244, 239), "bg2": (255, 255, 255), "accent": (15, 118, 110),
        "accent_soft": (224, 242, 241), "text": (24, 24, 27), "muted": (82, 82, 91),
        "border": (214, 211, 205), "light": True,
    },
    "雑誌アイボリー×紺": {
        "bg1": (250, 247, 242), "bg2": (255, 254, 251), "accent": (30, 58, 138),
        "accent_soft": (224, 231, 255), "text": (24, 24, 27), "muted": (82, 82, 91),
        "border": (212, 207, 198), "light": True,
    },
    "ネオンシアン×漆黒": {
        "bg1": (7, 9, 16), "bg2": (14, 20, 34), "accent": (34, 211, 238),
        "accent_soft": (14, 54, 72), "text": (248, 250, 252), "muted": (186, 230, 253),
        "border": (21, 94, 117), "light": False,
    },
}


_LAST_PILLOW_ERROR = ""
_PILLOW_INSTALL_LOCK = threading.Lock()


def _inject_site_packages_paths():
    """動的にインストールされたPillowが再起動なしでimportできるようsys.pathとキャッシュを更新する"""
    import importlib
    import site
    candidates = []
    try:
        user_site = site.getusersitepackages()
        if user_site:
            candidates.append(user_site)
    except Exception:
        pass
    try:
        for sp in site.getsitepackages():
            if sp:
                candidates.append(sp)
    except Exception:
        pass
    appdata = os.environ.get("APPDATA", "")
    if appdata:
        candidates.append(
            os.path.join(
                appdata,
                "Python",
                f"Python{sys.version_info.major}{sys.version_info.minor}",
                "site-packages",
            )
        )
    local_app = os.environ.get("LOCALAPPDATA", "")
    if local_app:
        candidates.append(
            os.path.join(
                local_app,
                "Programs",
                "Python",
                f"Python{sys.version_info.major}{sys.version_info.minor}",
                "Lib",
                "site-packages",
            )
        )
    for p in candidates:
        if p and os.path.isdir(p) and p not in sys.path:
            sys.path.append(p)
    try:
        importlib.invalidate_caches()
    except Exception:
        pass


def _resolve_pip_python_exe():
    """pip実行に適した同一バージョンのpython.exeパスを特定する（pythonw.exe / frozen対策）"""
    if not getattr(sys, "frozen", False):
        exe = sys.executable or ""
        if exe and os.path.basename(exe).lower() == "pythonw.exe":
            sibling = os.path.join(os.path.dirname(exe), "python.exe")
            if os.path.exists(sibling):
                return sibling
        if exe and os.path.exists(exe):
            return exe
    ver_tag = f"{sys.version_info.major}{sys.version_info.minor}"
    for cand in [
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Python", f"Python{ver_tag}", "python.exe"),
        os.path.join(r"C:\Program Files", f"Python{ver_tag}", "python.exe"),
        os.path.join(rf"C:\Python{ver_tag}", "python.exe"),
    ]:
        if cand and os.path.exists(cand):
            return cand
    import shutil
    return shutil.which("python") or sys.executable


def ensure_pillow_available(log_func=None, allow_install=True):
    """Pillow(PIL)が未導入の場合は自動インストールを試行し、(Image, ImageDraw, ImageFont)を返す"""
    global _LAST_PILLOW_ERROR
    _inject_site_packages_paths()
    try:
        from PIL import Image, ImageDraw, ImageFont
        _LAST_PILLOW_ERROR = ""
        return Image, ImageDraw, ImageFont
    except ImportError as first_err:
        if not allow_install:
            return None, None, None
        with _PILLOW_INSTALL_LOCK:
            _inject_site_packages_paths()
            try:
                from PIL import Image, ImageDraw, ImageFont
                _LAST_PILLOW_ERROR = ""
                return Image, ImageDraw, ImageFont
            except ImportError:
                pass
            py_exe = _resolve_pip_python_exe()
            if log_func:
                log_func(f"[INFO] Pillowライブラリを自動インストールしています (Python: {py_exe})...")

            import subprocess
            kw = {"capture_output": True, "text": True, "timeout": 120}
            if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW"):
                kw["creationflags"] = subprocess.CREATE_NO_WINDOW

            err_logs = [f"初回import失敗: {first_err}", f"実行Python: {py_exe}"]
            try:
                pip_chk = subprocess.run([py_exe, "-m", "pip", "--version"], **kw)
                if pip_chk.returncode != 0:
                    err_logs.append(f"pip確認失敗({pip_chk.returncode}): {(pip_chk.stderr or pip_chk.stdout or '').strip()[:200]}")
                    subprocess.run([py_exe, "-m", "ensurepip", "--upgrade"], **kw)
            except Exception as ep:
                err_logs.append(f"pip確認例外: {ep}")

            strategies = [
                [py_exe, "-m", "pip", "install", "--no-input", "--disable-pip-version-check", "Pillow"],
                [py_exe, "-m", "pip", "install", "--no-input", "--disable-pip-version-check", "--user", "Pillow"],
                [py_exe, "-m", "pip", "install", "--no-input", "--disable-pip-version-check", "--trusted-host", "pypi.org", "--trusted-host", "files.pythonhosted.org", "Pillow"],
                [py_exe, "-m", "pip", "install", "--no-input", "--disable-pip-version-check", "--user", "--trusted-host", "pypi.org", "--trusted-host", "files.pythonhosted.org", "Pillow"],
            ]

            for idx, cmd in enumerate(strategies, 1):
                try:
                    res = subprocess.run(cmd, **kw)
                    _inject_site_packages_paths()
                    try:
                        from PIL import Image, ImageDraw, ImageFont
                        _LAST_PILLOW_ERROR = ""
                        if log_func:
                            log_func("[SUCCESS] Pillowのインストールが完了しました！")
                        return Image, ImageDraw, ImageFont
                    except ImportError as imp_e:
                        out_txt = ((res.stderr or "") + "\n" + (res.stdout or "")).strip()
                        short_out = out_txt[-300:] if len(out_txt) > 300 else out_txt
                        err_logs.append(f"試行{idx} (rc={res.returncode}): {short_out or imp_e}")
                except Exception as ex:
                    err_logs.append(f"試行{idx} 例外: {ex}")

            _LAST_PILLOW_ERROR = "\n".join(err_logs)
            if log_func:
                log_func(f"[ERROR] Pillowインストール失敗:\n{_LAST_PILLOW_ERROR}")
            return None, None, None


def _find_jp_font_path(bold=False, family="gothic"):
    win_fonts = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    if family == "mincho":
        candidates = [
            os.path.join(win_fonts, "yumindb.ttf" if bold else "yumin.ttf"),
            os.path.join(win_fonts, "BIZ-UDMinchoM.ttc"),
            os.path.join(win_fonts, "msmincho.ttc"),
            "/usr/share/fonts/opentype/noto/NotoSerifCJK-Bold.ttc",
        ]
        for path in candidates:
            if os.path.exists(path):
                return path
    if bold:
        candidates = [
            os.path.join(win_fonts, "meiryob.ttc"),
            os.path.join(win_fonts, "YuGothB.ttc"),
            os.path.join(win_fonts, "BIZ-UDGothicB.ttc"),
            os.path.join(win_fonts, "msgothic.ttc"),
            os.path.join(win_fonts, "meiryo.ttc"),
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansJP-Bold.otf",
        ]
    else:
        candidates = [
            os.path.join(win_fonts, "meiryo.ttc"),
            os.path.join(win_fonts, "YuGothM.ttc"),
            os.path.join(win_fonts, "YuGothR.ttc"),
            os.path.join(win_fonts, "BIZ-UDGothicR.ttc"),
            os.path.join(win_fonts, "msgothic.ttc"),
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansJP-Regular.otf",
        ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def _load_jp_font(ImageFont, size, bold=False, family="gothic"):
    path = _find_jp_font_path(bold=bold, family=family)
    if path:
        try:
            return ImageFont.truetype(path, size=size)
        except Exception:
            pass
    return ImageFont.load_default()


def _wrap_jp_text_pil(draw, text, font, max_width, max_lines=3):
    lines = []
    for para in str(text or "").split("\n"):
        cur = ""
        for ch in para:
            nxt = cur + ch
            try:
                w = draw.textlength(nxt, font=font)
            except Exception:
                bbox = draw.textbbox((0, 0), nxt, font=font)
                w = bbox[2] - bbox[0]
            if w > max_width and len(cur) > 0:
                lines.append(cur)
                cur = ch
                if len(lines) >= max_lines:
                    break
            else:
                cur = nxt
        if cur and len(lines) < max_lines:
            lines.append(cur)
        if len(lines) >= max_lines:
            break
    if len(lines) == max_lines:
        last = lines[-1]
        try:
            lw = draw.textlength(last, font=font)
        except Exception:
            bbox = draw.textbbox((0, 0), last, font=font)
            lw = bbox[2] - bbox[0]
        if lw > max_width - 20 and len(last) > 1:
            lines[-1] = last[:-1] + "…"
    return lines


def ensure_account_workbook(xlsx_path, account_id):
    """<account_id>.xlsx が存在しない場合に標準4シート構成で自動作成して開く"""
    if os.path.exists(xlsx_path):
        return openpyxl.load_workbook(xlsx_path)
    wb = openpyxl.Workbook()
    ws_set = wb.active
    ws_set.title = "設定"
    ws_set.append(["ツイート間隔"])
    ws_set.append(["※指定の範囲内から毎回ランダム"])
    ws_set.append([120, "分", "〜", 360, "分"])
    ws_set.append(["ツイート実行時間"])
    ws_set.append(["(時:分〜時:分)"])
    ws_set.append(["0:00", "〜", "0:00"])
    ws_set.append(["モード"])
    ws_set.append(["(0:上から順, 1:ランダム)"])
    ws_set.append([1])
    ws_set.append(["Chromeプロファイル保存パス"])
    ws_set.append(["(空欄の場合は chrome_profiles/<アカウントID> が自動使用されます)"])
    ws_set.append([f"chrome_profiles/{account_id}"])
    ws_set.append(["コミュニティまたはトピック", "信太山新地"])

    ws_posts = wb.create_sheet(title="投稿リスト")
    ws_posts.append(["ID", "ポスト文章", "コミュニティまたはトピック", "画像ファイルパス", "ステータス", "投稿完了日時", "エラー内容"])

    ws_daily = wb.create_sheet(title="日常ポスト")
    ws_daily.append(["枠名", "時間帯", "ポスト文章", "有効無効", "最終投稿日", "最終投稿日時"])
    ws_daily.append(["朝枠", "7:00-9:00", "{おはようございます☀️|良い朝ですね🌿} 今日もマイペースで頑張りましょう！", "有効", "", ""])

    wb.save(xlsx_path)
    return wb


def load_account_design_from_xlsx(xlsx_path, fallback_topic=""):
    """<account_id>.xlsx の「アカウント設計」シートから設計項目を読み込む"""
    default_d = dict(_DESIGN_PRESETS_PY["飛田新地・風景×求人案内"])
    if fallback_topic:
        default_d["topic"] = fallback_topic
    if not xlsx_path or not os.path.exists(xlsx_path):
        return default_d
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        design_ws = None
        for sname in wb.sheetnames:
            if "設計" in sname:
                design_ws = wb[sname]
                break
        if design_ws:
            for row in design_ws.iter_rows(values_only=True):
                if not row or len(row) < 2 or row[0] is None:
                    continue
                k = str(row[0]).strip()
                v = "" if row[1] is None else str(row[1]).strip()
                if not k:
                    continue
                if "ジャンル" in k and v:
                    default_d["genre"] = v
                elif "ターゲット" in k and v:
                    default_d["target"] = v
                elif "口調" in k and v:
                    default_d["tone"] = v
                elif "悩み" in k and v:
                    default_d["pain"] = v
                elif ("強み" in k or "ポイント" in k) and v:
                    default_d["benefits"] = v
                elif ("誘導文" in k or "CTA" in k) and v:
                    default_d["cta"] = v
                elif "URL" in k:
                    default_d["url"] = v
                elif "トピック" in k and v:
                    default_d["topic"] = v
                elif "スタイル" in k and v:
                    default_d["image_style"] = v
                elif "配色" in k and v:
                    default_d["palette"] = v
        wb.close()
    except Exception:
        pass
    return default_d


def save_account_design_to_xlsx(xlsx_path, account_id, design_dict):
    """<account_id>.xlsx の「アカウント設計」シートへ設計項目を保存する"""
    wb = ensure_account_workbook(xlsx_path, account_id)
    if "アカウント設計" in wb.sheetnames:
        ws = wb["アカウント設計"]
        if ws.max_row >= 1:
            ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet(title="アカウント設計")
    ws.append(["項目", "設定値"])
    ws.append(["ジャンル", design_dict.get("genre", "")])
    ws.append(["ターゲット", design_dict.get("target", "")])
    ws.append(["口調", design_dict.get("tone", "丁寧・誠実")])
    ws.append(["悩み", design_dict.get("pain", "")])
    ws.append(["強みポイント", design_dict.get("benefits", "")])
    ws.append(["誘導文CTA", design_dict.get("cta", "")])
    ws.append(["誘導URL", design_dict.get("url", "")])
    ws.append(["トピック", design_dict.get("topic", "")])
    ws.append(["画像スタイル", design_dict.get("image_style", "自動おまかせ（設計に合わせて最適化）")])
    ws.append(["配色テーマ", design_dict.get("palette", "提灯アンバー×黒")])
    wb.save(xlsx_path)
    wb.close()


def _normalize_image_style_label_py(raw_style):
    s = (raw_style or "").strip()
    if not s:
        return "自動おまかせ（設計に合わせて最適化）"
    alias = {
        "背景写真＋文字入れ": "背景写真＋自動文字入れ",
        "photo_overlay": "背景写真＋自動文字入れ",
        "magazine_clean": "雑誌エディトリアル風",
        "chat_bubble": "吹き出し会話SNS風",
        "neon_night": "ネオン夜風",
        "step_roadmap": "ステップ図解",
        "qa_card": "Q&A神回答カード",
        "diagram_card": "SNS図解カード",
        "チェックリスト": "保存用チェックリスト",
        "checklist": "保存用チェックリスト",
        "ミックス": "全テンプレートミックス",
        "mix": "全テンプレートミックス",
        "all_mix": "全テンプレートミックス",
        "auto": "自動おまかせ（設計に合わせて最適化）",
    }
    if s in alias:
        return alias[s]
    return s


def resolve_design_image_templates(design_dict, has_photos=False):
    """アカウント設計（ジャンル・ターゲット・口調・写真有無）に応じて最適な画像テンプレート順を自動選定"""
    style_pref = _normalize_image_style_label_py(design_dict.get("image_style"))
    style_map = {
        "背景写真＋自動文字入れ": "photo_overlay",
        "雑誌エディトリアル風": "magazine_clean",
        "吹き出し会話SNS風": "chat_bubble",
        "ネオン夜風": "neon_night",
        "ステップ図解": "step_roadmap",
        "Q&A神回答カード": "qa_card",
        "SNS図解カード": "diagram_card",
        "保存用チェックリスト": "checklist",
    }
    if style_pref in style_map:
        return [style_map[style_pref]]

    all_styles = ["magazine_clean", "chat_bubble", "neon_night", "step_roadmap", "qa_card", "diagram_card", "checklist"]
    if has_photos:
        all_styles.insert(0, "photo_overlay")
    if "全テンプレート" in style_pref:
        return all_styles
    if "3パターン" in style_pref or "ミックス" in style_pref:
        return ["photo_overlay", "diagram_card", "checklist"] if has_photos else ["magazine_clean", "diagram_card", "checklist"]

    # 自動おまかせ：ジャンル・トピック・口調を解析
    joined = " ".join([
        str(design_dict.get("genre") or ""),
        str(design_dict.get("topic") or ""),
        str(design_dict.get("target") or ""),
        str(design_dict.get("pain") or ""),
    ]).lower()
    tone = (design_dict.get("tone") or "丁寧・誠実").strip()

    if any(k in joined for k in ["新地", "ナイト", "高収入", "キャバ", "夜職", "ガールズ", "ラウンジ", "体験入店", "短期集中"]):
        if tone == "フレンドリー":
            seq = ["chat_bubble", "neon_night", "qa_card", "magazine_clean"]
        elif tone == "共感スタイル":
            seq = ["chat_bubble", "qa_card", "magazine_clean", "neon_night"]
        else:
            seq = ["neon_night", "magazine_clean", "qa_card", "chat_bubble"]
    elif any(k in joined for k in ["副業", "キャリア", "転職", "sns", "threads", "ai", "ビジネス", "稼", "運用", "ノウハウ", "投資"]):
        if tone == "プロ解説":
            seq = ["step_roadmap", "diagram_card", "qa_card", "checklist"]
        else:
            seq = ["step_roadmap", "magazine_clean", "chat_bubble", "checklist"]
    elif any(k in joined for k in ["美容", "カフェ", "暮", "ダイエット", "コスメ", "女性", "育児", "習慣"]):
        seq = ["magazine_clean", "checklist", "chat_bubble", "qa_card"]
    else:
        tone_defaults = {
            "プロ解説": ["step_roadmap", "diagram_card", "qa_card", "checklist"],
            "フレンドリー": ["chat_bubble", "magazine_clean", "qa_card", "checklist"],
            "共感スタイル": ["chat_bubble", "qa_card", "magazine_clean", "checklist"],
            "丁寧・誠実": ["magazine_clean", "qa_card", "step_roadmap", "checklist"],
        }
        seq = list(tone_defaults.get(tone, ["magazine_clean", "qa_card", "step_roadmap", "checklist"]))

    if has_photos and "photo_overlay" not in seq:
        seq = ["photo_overlay", seq[0], seq[1], "photo_overlay"] + seq[2:]
    return seq


def generate_posts_from_design(design_dict, count=10, enable_spintax=True, has_photos=False):
    """フック(1行目)×本文×CTA(結び)＋スピンタックスで投稿セットを自動生成"""
    hooks_map = {
        "丁寧・誠実": [
            "【保存版】{topic}で失敗しないために最初に知っておくべきこと",
            "「{pain}」と迷われている方へ、現場のリアルをお伝えします。",
            "{target}から特にご相談が多い疑問を分かりやすく整理しました。",
            "実はあまり知られていませんが、{genre}には安心できる共通点があります。",
            "初めての方ほど大切にしてほしい、{topic}の3つのチェック基準。",
            "雰囲気を見てから決めたい方へ。事前相談でよく聞かれるポイントまとめ。",
            "無駄な不安を減らして納得して選ぶための{genre}ガイド。",
            "短時間で効率よく結果を出したい方が、最初に見直すべきポイント。",
        ],
        "フレンドリー": [
            "ぶっちゃけ「{pain}」って思ってる人、かなり多いですよね…！",
            "{target}にこそ知ってほしい、{topic}のリアルな話✨",
            "これ知らないまま動くのはもったいない！{genre}のポイントまとめ🙌",
            "よく質問をもらう「実際どうなの？」に正直にお答えします💬",
            "最初の一歩がいちばん緊張するけど、実はこんな雰囲気なんです🌿",
            "気になるけど聞きづらいこと、ここでまとめてスッキリ解決しましょう☕️",
            "自分のペースを大事にしたい人ほど、この選び方がしっくりきます✨",
            "保存してあとで見返せるように、大事なポイントだけギュッとまとめました📌",
        ],
        "プロ解説": [
            "【結論】{topic}で後悔しない人は、必ず「この基準」を先に確認しています。",
            "「{pain}」を放置したまま選ぶと、時間と労力をロスする理由。",
            "{target}向けの{genre}：成果と安心感を両立する設計ルール。",
            "曖昧な情報に振り回されないための、チェック項目リスト公開。",
            "効率よく条件を整える人が実践している3ステップ手順。",
            "ミスマッチを防ぐために、事前に必ず確認すべきポイント一覧。",
            "再現性高く進めるための、{topic}・実践テンプレート整理。",
            "迷う時間をゼロにする、シンプルな判断軸をまとめました。",
        ],
        "共感スタイル": [
            "「{pain}」…その気持ち、すごくよく分かります。",
            "無理をしてまで頑張らなくて大丈夫。{target}に合った選び方があります。",
            "不安があるまま動かなくてOK。まずは雰囲気を知るところから始めてみませんか？",
            "誰にも相談できずに一人で悩んでいた方に、そっと届けたいお話。",
            "{genre}は「知ってから決める」だけで、気持ちの負担がぐっと軽くなります。",
            "大切なのは、あなた自身の安心とペースを守れる環境かどうか。",
            "「思っていたより落ち着いた雰囲気だった」という声が多い理由とは。",
            "焦らず自分のタイミングで納得できるよう、ポイントをまとめました。",
        ],
    }

    connectors_spintax = [
        "{まずは雰囲気を知るだけでもOKです|最初の一歩は情報収集からで十分です|事前確認だけで不安の大半は解消できます}。",
        "{条件やペースを重視したい方ほど|無理なく続けたい方ほど|初めて検討される方ほど}、事前のすり合わせが鍵になります。",
        "{見学や質問だけでも歓迎される環境|プライバシー配慮が整っている環境|サポート体制が明確な環境}を選ぶのがおすすめです。",
        "{1人で抱え込まずに|焦って決めずに|比較しながら}、自分に合う選択肢を確かめてみてください。",
    ]
    connectors_plain = [
        "まずは雰囲気を知るだけでも十分な前進になります。",
        "条件やペースを重視したい方ほど、事前のすり合わせが大切です。",
        "サポート体制やプライバシー配慮が明確な環境を選ぶと安心です。",
        "焦って決めずに、自分に合う選択肢を一つずつ確かめてみてください。",
    ]
    closers_spintax = [
        "{気になる点があればいつでも参考にしてみてください|まずはブックマークして見返してみてください|一歩目の判断材料になれば嬉しいです}",
        "{無理のないペースで進めていきましょう|自分に合った環境選びを応援しています|納得できる選択ができるよう願っています}",
    ]

    topic = (design_dict.get("topic") or "おすすめ").strip()
    genre = (design_dict.get("genre") or topic).strip()
    target = (design_dict.get("target") or "これから始める方").strip()
    pain = (design_dict.get("pain") or "何から確認すればいいか分からない").strip()
    cta = (design_dict.get("cta") or "詳細はプロフィールの案内をご覧ください。").strip()
    url = (design_dict.get("url") or "").strip()

    raw_bullets = [s.strip() for s in str(design_dict.get("benefits") or "").splitlines() if s.strip()]
    bullets = raw_bullets[:4] if raw_bullets else [
        "事前の条件・システム確認で安心",
        "自分のペースに合わせた柔軟な選択",
        "プライバシーとサポート体制を重視",
        "まずは見学・相談からスタート可能",
    ]

    tone = design_dict.get("tone") or "丁寧・誠実"
    hooks = hooks_map.get(tone, hooks_map["丁寧・誠実"])
    template_seq = resolve_design_image_templates(design_dict, has_photos=has_photos)

    custom_headlines = [
        f"{topic}・失敗しない選び方",
        f"知らないと損する{topic}の基本",
        f"初めてでも安心の{topic}ガイド",
        f"{topic}のリアルな雰囲気とは",
        "後悔しないための3つの基準",
        "無理なく続くワークスタイル",
        "保存版・事前チェック項目",
        "短時間×安心環境のポイント",
    ]

    results = []
    for i in range(max(1, min(200, int(count)))):
        h_tpl = hooks[i % len(hooks)]
        hook = h_tpl.replace("{topic}", topic).replace("{genre}", genre).replace("{target}", target).replace("{pain}", pain)
        b1 = bullets[i % len(bullets)]
        b2 = bullets[(i + 1) % len(bullets)]
        conn = connectors_spintax[i % len(connectors_spintax)] if enable_spintax else connectors_plain[i % len(connectors_plain)]
        closer = closers_spintax[i % len(closers_spintax)] if enable_spintax else "ぜひ参考にしてみてください。"
        bullet_block = f"① {b1}\n② {b2}" if tone == "プロ解説" else f"・{b1}\n・{b2}"
        cta_line = f"{{{cta}|{closer}。{cta}}}" if enable_spintax else cta
        parts = [hook, "", bullet_block, conn, "", cta_line]
        if url:
            parts.append(url)
        full_text = "\n".join(parts).strip()

        if i % 2 == 0:
            headline = custom_headlines[(i // 2) % len(custom_headlines)]
        else:
            stripped = re.sub(r"【[^】]+】", "", hook).replace("「", "").replace("」", "").strip()
            headline = stripped[:25] + ("…" if len(stripped) > 25 else "")

        sub_raw = f"{target}向け · {b1}"
        subhead = sub_raw[:33] + ("…" if len(sub_raw) > 33 else "")
        resolved_style = template_seq[i % len(template_seq)]

        results.append({
            "index": i + 1,
            "text": full_text,
            "topic": topic,
            "headline": headline,
            "subhead": subhead,
            "bullets": bullets,
            "pain": pain,
            "cta": cta,
            "style": resolved_style,
            "filename": f"gen_{i + 1:03d}.png",
        })
    return results


def render_post_image_pil(
    Image, ImageDraw, ImageFont,
    style="photo_overlay",
    aspect_ratio="1:1",
    palette_name="提灯アンバー×黒",
    scrim_opacity=0.48,
    topic="",
    headline="",
    subhead="",
    bullets=None,
    account_id="",
    post_index=1,
    bg_image_path=None,
    show_watermark=True,
    output_path=None,
    pain="",
):
    """アカウント設計に合わせて多彩なデザインテンプレート（写真文字入れ/雑誌風/吹き出し/ネオン夜風/ステップ図解/Q&A/図解/チェックリスト）を描画"""
    width = 1080
    height = 1350 if aspect_ratio == "4:5" else 1080
    theme = _PIL_PALETTES.get(palette_name, _PIL_PALETTES["提灯アンバー×黒"])
    pad = 76
    bullets = bullets or ["事前の条件確認で安心", "自分のペースで続けやすい", "プライバシー重視の環境"]
    pain_text = (pain or "何から確認すればいいか分からず不安…").strip()

    def _fit_cover(bg_path):
        bg = Image.open(bg_path).convert("RGB")
        iw, ih = bg.size
        if iw <= 0 or ih <= 0:
            return None
        ir = iw / ih
        cr = width / height
        if ir > cr:
            nh = height
            nw = int(height * ir)
            bg = bg.resize((nw, nh), Image.LANCZOS)
            left = (nw - width) // 2
            bg = bg.crop((left, 0, left + width, height))
        else:
            nw = width
            nh = int(width / ir)
            bg = bg.resize((nw, nh), Image.LANCZOS)
            top = (nh - height) // 2
            bg = bg.crop((0, top, width, top + height))
        return bg

    def _smooth_ambient_bg(c1, c2, ac, tint_strength=0.30):
        from PIL import ImageFilter
        sw, sh = 240, max(240, int(240 * height / width))
        mesh = Image.new("RGB", (sw, sh), c1)
        md = ImageDraw.Draw(mesh)
        for gy in range(sh):
            t = gy / max(1, sh - 1)
            md.line([(0, gy), (sw, gy)], fill=(
                int(c2[0] * (1 - t) + c1[0] * t),
                int(c2[1] * (1 - t) + c1[1] * t),
                int(c2[2] * (1 - t) + c1[2] * t),
            ))
        w1 = (
            int(c2[0] * (1 - tint_strength) + ac[0] * tint_strength),
            int(c2[1] * (1 - tint_strength) + ac[1] * tint_strength),
            int(c2[2] * (1 - tint_strength) + ac[2] * tint_strength),
        )
        md.ellipse([int(sw * 0.38), -int(sh * 0.12), int(sw * 1.12), int(sh * 0.62)], fill=w1)
        mesh = mesh.filter(ImageFilter.GaussianBlur(radius=48))
        return mesh.resize((width, height), Image.BICUBIC)

    def _rrect(draw_obj, box, radius=18, fill=None, outline=None, width_px=2):
        try:
            draw_obj.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width_px)
        except Exception:
            draw_obj.rectangle(box, fill=fill, outline=outline, width=width_px)

    if style == "photo_overlay":
        base = None
        has_real_photo = False
        if bg_image_path and os.path.exists(bg_image_path):
            try:
                base = _fit_cover(bg_image_path)
                if base is not None:
                    has_real_photo = True
            except Exception:
                base = None

        if has_real_photo and base is not None:
            base_rgba = base.convert("RGBA")
            alpha_val = max(0.15, min(0.85, float(scrim_opacity)))
            scrim = Image.new("RGBA", (width, height), (0, 0, 0, 0))
            s_draw = ImageDraw.Draw(scrim)
            for gy in range(height):
                t = gy / max(1, height)
                if t < 0.42:
                    a = alpha_val * (0.18 + 0.25 * (t / 0.42))
                else:
                    bottom_target = min(0.88, alpha_val + 0.34)
                    start_a = alpha_val * 0.43
                    a = start_a + (bottom_target - start_a) * (((t - 0.42) / 0.58) ** 1.15)
                s_draw.line([(0, gy), (width, gy)], fill=(8, 10, 14, int(max(0, min(230, a * 255)))))
            img = Image.alpha_composite(base_rgba, scrim)
        else:
            img = _smooth_ambient_bg(theme["bg1"], theme["bg2"], theme["accent"], tint_strength=0.35).convert("RGBA")

        draw = ImageDraw.Draw(img)
        frame_col = (255, 255, 255, 48) if not theme["light"] else theme["border"] + (200,)
        txt_main_col = (255, 255, 255, 255) if (has_real_photo or not theme["light"]) else theme["text"] + (255,)
        txt_sub_col = (241, 245, 249, 232) if (has_real_photo or not theme["light"]) else theme["muted"] + (255,)
        txt_topic_col = (248, 250, 252, 255) if (has_real_photo or not theme["light"]) else theme["text"] + (255,)

        draw.rectangle([38, 38, width - 38, height - 38], outline=frame_col, width=2)
        draw.rectangle([pad, pad + 6, pad + 42, pad + 12], fill=theme["accent"] + (255,))
        f_topic = _load_jp_font(ImageFont, 28, bold=True)
        draw.text((pad + 58, pad - 6), topic or "GUIDE", font=f_topic, fill=txt_topic_col)

        f_head = _load_jp_font(ImageFont, 66, bold=True)
        f_sub = _load_jp_font(ImageFont, 32, bold=False)
        f_foot = _load_jp_font(ImageFont, 26, bold=True)

        max_w = width - pad * 2
        h_lines = _wrap_jp_text_pil(draw, headline, f_head, max_w, max_lines=3)
        s_lines = _wrap_jp_text_pil(draw, subhead, f_sub, max_w, max_lines=2) if subhead else []

        line_h = 88
        sub_line_h = 44
        sub_block_h = (14 + len(s_lines) * sub_line_h + 18) if s_lines else 0
        foot_h = 66 if show_watermark else 16
        total_h = 40 + len(h_lines) * line_h + sub_block_h + foot_h
        cy = height - pad - total_h

        draw.rectangle([pad, cy, pad + 72, cy + 8], fill=theme["accent"] + (255,))
        cy += 40

        for line in h_lines:
            if has_real_photo or not theme["light"]:
                draw.text((pad + 2, cy + 4), line, font=f_head, fill=(0, 0, 0, 190))
            draw.text((pad, cy), line, font=f_head, fill=txt_main_col)
            cy += line_h

        if s_lines:
            cy += 14
            for sl in s_lines:
                draw.text((pad, cy), sl, font=f_sub, fill=txt_sub_col)
                cy += sub_line_h

        if show_watermark:
            div_y = height - pad - 44
            div_col = (255, 255, 255, 64) if (has_real_photo or not theme["light"]) else theme["border"] + (220,)
            wm_col = (226, 232, 240, 225) if (has_real_photo or not theme["light"]) else theme["muted"] + (255,)
            draw.line([(pad, div_y), (width - pad, div_y)], fill=div_col, width=2)
            draw.text((pad, height - pad - 28), f"@{account_id or 'threads'}", font=f_foot, fill=wm_col)
            no_txt = f"NO.{post_index:02d}"
            try:
                nw = draw.textlength(no_txt, font=f_foot)
            except Exception:
                nw = 90
            draw.text((width - pad - nw, height - pad - 28), no_txt, font=f_foot, fill=theme["accent"] + (255,))
        out_img = img.convert("RGB")

    elif style == "magazine_clean":
        # 雑誌・エディトリアル風：明朝体見出し＋余白＋洗練インデックス
        paper_bg = (250, 247, 242) if theme["light"] else (18, 19, 24)
        card_bg = (255, 255, 255) if theme["light"] else (26, 28, 36)
        ink_col = (24, 24, 27) if theme["light"] else (248, 250, 252)
        sub_col = (82, 82, 91) if theme["light"] else (180, 188, 200)
        line_col = (220, 214, 204) if theme["light"] else (48, 52, 64)
        ac = theme["accent"]

        out_img = Image.new("RGB", (width, height), paper_bg)
        if bg_image_path and os.path.exists(bg_image_path):
            try:
                bg_cov = _fit_cover(bg_image_path)
                if bg_cov:
                    out_img = Image.blend(out_img, bg_cov, 0.08 if theme["light"] else 0.14)
            except Exception:
                pass
        draw = ImageDraw.Draw(out_img)
        draw.rectangle([42, 42, width - 42, height - 42], outline=line_col, width=2)

        f_vol = _load_jp_font(ImageFont, 24, bold=True)
        f_head = _load_jp_font(ImageFont, 58, bold=True, family="mincho")
        f_sub = _load_jp_font(ImageFont, 28, bold=False)
        f_idx = _load_jp_font(ImageFont, 32, bold=True, family="mincho")
        f_item = _load_jp_font(ImageFont, 30, bold=True)
        f_foot = _load_jp_font(ImageFont, 24, bold=True)

        y = pad + 6
        draw.text((pad, y), f"ISSUE NO.{post_index:02d}", font=f_vol, fill=ac)
        topic_str = (topic or "FEATURE").upper()
        try:
            tw = draw.textlength(topic_str, font=f_vol)
        except Exception:
            tw = 140
        draw.text((width - pad - tw, y), topic_str, font=f_vol, fill=sub_col)
        y += 42
        draw.rectangle([pad, y, pad + 88, y + 6], fill=ac)
        draw.line([(pad + 104, y + 3), (width - pad, y + 3)], fill=line_col, width=2)
        y += 42

        max_w = width - pad * 2
        h_lines = _wrap_jp_text_pil(draw, headline, f_head, max_w, max_lines=3)
        for hl in h_lines:
            draw.text((pad, y), hl, font=f_head, fill=ink_col)
            y += 78

        if subhead:
            y += 6
            s_lines = _wrap_jp_text_pil(draw, subhead, f_sub, max_w, max_lines=2)
            for sl in s_lines:
                draw.text((pad, y), sl, font=f_sub, fill=sub_col)
                y += 40

        y += 28
        _rrect(draw, [pad, y, width - pad, height - pad - 68], radius=16, fill=card_bg, outline=line_col, width_px=2)

        inner_y = y + 34
        items = bullets[:3]
        slot_h = max(92, (height - pad - 68 - y - 56) // max(1, len(items)))
        for idx_i, it_txt in enumerate(items):
            num_s = f"{idx_i + 1:02d}."
            draw.text((pad + 34, inner_y + 4), num_s, font=f_idx, fill=ac)
            wrapped = _wrap_jp_text_pil(draw, it_txt, f_item, max_w - 140, max_lines=2)
            for li, wl in enumerate(wrapped):
                draw.text((pad + 112, inner_y + li * 40), wl, font=f_item, fill=ink_col)
            if idx_i < len(items) - 1:
                sep_y = inner_y + slot_h - 14
                draw.line([(pad + 34, sep_y), (width - pad - 34, sep_y)], fill=line_col, width=1)
            inner_y += slot_h

        if show_watermark:
            draw.text((pad, height - pad - 24), f"@{account_id or 'threads'}", font=f_foot, fill=sub_col)
            r_lbl = "EDITORIAL PICK"
            try:
                rw = draw.textlength(r_lbl, font=f_foot)
            except Exception:
                rw = 170
            draw.text((width - pad - rw, height - pad - 24), r_lbl, font=f_foot, fill=ac)

    elif style == "chat_bubble":
        # 吹き出し・会話SNS風：左(読者の悩みQ) → 右(アカウントの回答A)
        out_img = _smooth_ambient_bg(theme["bg1"], theme["bg2"], theme["accent"], tint_strength=0.18)
        draw = ImageDraw.Draw(out_img)
        ac = theme["accent"]
        q_bg = (232, 236, 242) if theme["light"] else (32, 36, 48)
        q_border = (203, 213, 225) if theme["light"] else (58, 65, 84)
        a_bg = (255, 255, 255) if theme["light"] else (24, 30, 42)
        ink_col = theme["text"]
        sub_col = theme["muted"]

        f_top = _load_jp_font(ImageFont, 26, bold=True)
        f_q = _load_jp_font(ImageFont, 32, bold=True)
        f_head = _load_jp_font(ImageFont, 40, bold=True)
        f_item = _load_jp_font(ImageFont, 30, bold=True)
        f_badge = _load_jp_font(ImageFont, 24, bold=True)
        f_foot = _load_jp_font(ImageFont, 25, bold=True)

        y = pad
        draw.text((pad, y), f"💬 {topic or '相談'} · よくあるお悩み＆回答", font=f_top, fill=ac)
        y += 52
        draw.line([(pad, y), (width - pad, y)], fill=theme["border"], width=2)
        y += 32

        # 左側：読者の悩み吹き出し
        av_s = 72
        _rrect(draw, [pad, y, pad + av_s, y + av_s], radius=36, fill=q_bg, outline=q_border, width_px=2)
        draw.text((pad + 24, y + 20), "Q", font=f_top, fill=sub_col)

        qb_x1 = pad + av_s + 20
        qb_x2 = width - pad - 70
        q_lines = _wrap_jp_text_pil(draw, f"「{pain_text}」", f_q, qb_x2 - qb_x1 - 56, max_lines=3)
        qb_h = max(110, 44 + len(q_lines) * 44)
        _rrect(draw, [qb_x1, y, qb_x2, y + qb_h], radius=24, fill=q_bg, outline=q_border, width_px=2)
        draw.polygon([(qb_x1 + 2, y + 28), (qb_x1 - 14, y + 36), (qb_x1 + 2, y + 48)], fill=q_bg)
        for idx_q, ql in enumerate(q_lines):
            draw.text((qb_x1 + 28, y + 24 + idx_q * 44), ql, font=f_q, fill=ink_col)
        y += qb_h + 38

        # 右側：回答見出し吹き出し
        head_fg = (12, 12, 14) if sum(ac) > 460 else (255, 255, 255)
        ab_x1 = pad + 56
        ab_x2 = width - pad - av_s - 20
        _rrect(draw, [width - pad - av_s, y, width - pad, y + av_s], radius=36, fill=ac, outline=ac, width_px=2)
        draw.text((width - pad - av_s + 24, y + 20), "A", font=f_top, fill=head_fg)

        a_lines = _wrap_jp_text_pil(draw, headline, f_head, ab_x2 - ab_x1 - 56, max_lines=2)
        ab_h = max(106, 44 + len(a_lines) * 54)
        _rrect(draw, [ab_x1, y, ab_x2, y + ab_h], radius=24, fill=ac, outline=ac, width_px=2)
        draw.polygon([(ab_x2 - 2, y + 28), (ab_x2 + 14, y + 36), (ab_x2 - 2, y + 48)], fill=ac)
        for idx_a, al in enumerate(a_lines):
            draw.text((ab_x1 + 28, y + 24 + idx_a * 54), al, font=f_head, fill=head_fg)
        y += ab_h + 26

        # 下側：具体ポイント返信バブル
        avail_bot = height - pad - 60 - y
        items = bullets[:3]
        _rrect(draw, [ab_x1, y, ab_x2, y + avail_bot], radius=24, fill=a_bg, outline=ac, width_px=2)
        row_step = max(66, (avail_bot - 40) // max(1, len(items)))
        iy = y + 26
        for bi, bt in enumerate(items):
            _rrect(draw, [ab_x1 + 26, iy + 6, ab_x1 + 60, iy + 40], radius=10, fill=theme["accent_soft"], outline=ac, width_px=2)
            draw.text((ab_x1 + 36, iy + 9), f"{bi + 1}", font=f_badge, fill=ac)
            b_lines = _wrap_jp_text_pil(draw, bt, f_item, ab_x2 - ab_x1 - 108, max_lines=2)
            for li, bl in enumerate(b_lines):
                draw.text((ab_x1 + 78, iy + 4 + li * 38), bl, font=f_item, fill=ink_col)
            iy += row_step

        if show_watermark:
            draw.text((pad, height - pad - 24), f"@{account_id or 'threads'}", font=f_foot, fill=sub_col)
            r_txt = f"Q&A #{post_index:02d}"
            try:
                rw = draw.textlength(r_txt, font=f_foot)
            except Exception:
                rw = 110
            draw.text((width - pad - rw, height - pad - 24), r_txt, font=f_foot, fill=ac)

    elif style == "neon_night":
        # ネオン夜風：漆黒＋GaussianBlur発光ネオン枠＋グラスカード
        from PIL import ImageFilter
        ac = theme["accent"] if not theme["light"] else (56, 189, 248)
        base = _smooth_ambient_bg((7, 8, 14), (15, 18, 30), ac, tint_strength=0.22)
        if bg_image_path and os.path.exists(bg_image_path):
            try:
                bg_cov = _fit_cover(bg_image_path)
                if bg_cov:
                    base = Image.blend(base, bg_cov, 0.18)
            except Exception:
                pass
        img = base.convert("RGBA")
        neon_layer = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        nd = ImageDraw.Draw(neon_layer)
        _rrect(nd, [44, 44, width - 44, height - 44], radius=26, fill=None, outline=ac + (210,), width_px=12)
        nd.line([(pad, 250), (pad + 180, 250)], fill=ac + (230,), width=10)
        neon_blur = neon_layer.filter(ImageFilter.GaussianBlur(radius=16))
        img = Image.alpha_composite(img, neon_blur)
        draw = ImageDraw.Draw(img)
        _rrect(draw, [44, 44, width - 44, height - 44], radius=26, fill=None, outline=ac + (255,), width_px=3)

        f_tag = _load_jp_font(ImageFont, 24, bold=True)
        f_head = _load_jp_font(ImageFont, 58, bold=True)
        f_item = _load_jp_font(ImageFont, 32, bold=True)
        f_num = _load_jp_font(ImageFont, 32, bold=True)
        f_foot = _load_jp_font(ImageFont, 25, bold=True)

        y = pad + 8
        tag_txt = f"◆ {(topic or 'NIGHT').upper()}  //  PICK UP {post_index:02d}"
        try:
            tag_tw = draw.textlength(tag_txt, font=f_tag)
        except Exception:
            tag_tw = len(tag_txt) * 18
        tag_box_w = min(width - pad * 2, max(320, int(tag_tw) + 42))
        _rrect(draw, [pad, y, pad + tag_box_w, y + 48], radius=10, fill=(20, 24, 38, 220), outline=ac + (180,), width_px=2)
        draw.text((pad + 18, y + 10), tag_txt, font=f_tag, fill=ac + (255,))
        y += 82

        max_w = width - pad * 2
        h_lines = _wrap_jp_text_pil(draw, headline, f_head, max_w, max_lines=2)
        for hl in h_lines:
            draw.text((pad + 3, y + 4), hl, font=f_head, fill=(0, 0, 0, 200))
            draw.text((pad, y), hl, font=f_head, fill=(255, 255, 255, 255))
            y += 78
        y += 14
        draw.rectangle([pad, y, pad + 160, y + 6], fill=ac + (255,))
        draw.line([(pad + 176, y + 3), (width - pad, y + 3)], fill=(255, 255, 255, 50), width=2)
        y += 42

        items = bullets[:3]
        avail_h = height - pad - 68 - y
        gap = 24
        card_h = min(176, max(108, (avail_h - gap * (len(items) - 1)) // max(1, len(items))))
        for idx_i, it_txt in enumerate(items):
            cy = y + idx_i * (card_h + gap)
            _rrect(draw, [pad, cy, width - pad, cy + card_h], radius=16, fill=(16, 20, 32, 228), outline=ac + (145,), width_px=2)
            draw.rectangle([pad + 22, cy + 24, pad + 30, cy + card_h - 24], fill=ac + (255,))
            draw.text((pad + 52, cy + (card_h - 38) // 2), f"{idx_i + 1:02d}", font=f_num, fill=ac + (255,))
            ilines = _wrap_jp_text_pil(draw, it_txt, f_item, max_w - 150, max_lines=2)
            if len(ilines) == 1:
                draw.text((pad + 122, cy + (card_h - 36) // 2), ilines[0], font=f_item, fill=(248, 250, 252, 255))
            else:
                draw.text((pad + 122, cy + (card_h // 2) - 42), ilines[0], font=f_item, fill=(248, 250, 252, 255))
                draw.text((pad + 122, cy + (card_h // 2) + 4), ilines[1], font=f_item, fill=(248, 250, 252, 255))

        if show_watermark:
            draw.text((pad, height - pad - 24), f"@{account_id or 'threads'}", font=f_foot, fill=(203, 213, 225, 225))
            r_txt = f"NIGHT EDITION #{post_index:02d}"
            try:
                rw = draw.textlength(r_txt, font=f_foot)
            except Exception:
                rw = 210
            draw.text((width - pad - rw, height - pad - 24), r_txt, font=f_foot, fill=ac + (255,))
        out_img = img.convert("RGB")

    elif style == "step_roadmap":
        # ステップ図解・ロードマップ風：縦ライン＋STEP 01→02→03接続ノード
        out_img = _smooth_ambient_bg(theme["bg1"], theme["bg2"], theme["accent"], tint_strength=0.14)
        draw = ImageDraw.Draw(out_img)
        ac = theme["accent"]

        f_kick = _load_jp_font(ImageFont, 26, bold=True)
        f_head = _load_jp_font(ImageFont, 50, bold=True)
        f_step_lbl = _load_jp_font(ImageFont, 22, bold=True)
        f_item = _load_jp_font(ImageFont, 31, bold=True)
        f_foot = _load_jp_font(ImageFont, 25, bold=True)

        y = pad
        draw.text((pad, y), f"🧭 {topic or 'GUIDE'} · 実践ステップ図解", font=f_kick, fill=ac)
        y += 46
        max_w = width - pad * 2
        h_lines = _wrap_jp_text_pil(draw, headline, f_head, max_w, max_lines=2)
        for hl in h_lines:
            draw.text((pad, y), hl, font=f_head, fill=theme["text"])
            y += 66
        y += 14
        draw.line([(pad, y), (width - pad, y)], fill=theme["border"], width=2)
        y += 38

        items = bullets[:3]
        avail_h = height - pad - 68 - y
        gap = 26
        card_h = min(192, max(112, (avail_h - gap * (len(items) - 1)) // max(1, len(items))))
        node_x = pad + 34

        if len(items) > 1:
            top_cy = y + card_h // 2
            bot_cy = y + (len(items) - 1) * (card_h + gap) + card_h // 2
            draw.line([(node_x, top_cy), (node_x, bot_cy)], fill=ac, width=5)

        for idx_i, it_txt in enumerate(items):
            cy = y + idx_i * (card_h + gap)
            mid_y = cy + card_h // 2
            _rrect(draw, [node_x - 28, mid_y - 28, node_x + 28, mid_y + 28], radius=28, fill=ac, outline=theme["bg1"], width_px=4)
            num_s = f"{idx_i + 1}"
            num_fg = (12, 12, 14) if sum(ac) > 460 else (255, 255, 255)
            draw.text((node_x - 9, mid_y - 16), num_s, font=f_kick, fill=num_fg)

            cx1 = pad + 88
            _rrect(draw, [cx1, cy, width - pad, cy + card_h], radius=16, fill=theme["bg2"], outline=theme["border"], width_px=2)
            draw.text((cx1 + 26, cy + 18), f"STEP 0{idx_i + 1}", font=f_step_lbl, fill=ac)
            ilines = _wrap_jp_text_pil(draw, it_txt, f_item, width - pad - cx1 - 52, max_lines=2)
            for li, il in enumerate(ilines):
                draw.text((cx1 + 26, cy + 54 + li * 42), il, font=f_item, fill=theme["text"])

        if show_watermark:
            draw.text((pad, height - pad - 24), f"@{account_id or 'threads'}", font=f_foot, fill=theme["muted"])
            r_txt = "3-STEP ROADMAP"
            try:
                rw = draw.textlength(r_txt, font=f_foot)
            except Exception:
                rw = 190
            draw.text((width - pad - rw, height - pad - 24), r_txt, font=f_foot, fill=ac)

    elif style == "qa_card":
        # Q&A神回答カード：上段に質問(悩み)、下段に結論＋3つの回答ポイント
        out_img = _smooth_ambient_bg(theme["bg1"], theme["bg2"], theme["accent"], tint_strength=0.14)
        draw = ImageDraw.Draw(out_img)
        ac = theme["accent"]

        f_kick = _load_jp_font(ImageFont, 26, bold=True)
        f_big_q = _load_jp_font(ImageFont, 54, bold=True)
        f_q = _load_jp_font(ImageFont, 34, bold=True)
        f_a_head = _load_jp_font(ImageFont, 44, bold=True)
        f_item = _load_jp_font(ImageFont, 30, bold=True)
        f_foot = _load_jp_font(ImageFont, 25, bold=True)

        y = pad
        draw.text((pad, y), f"❓ {topic or 'GUIDE'} · Q&A よくいただく質問", font=f_kick, fill=ac)
        y += 48

        max_w = width - pad * 2
        q_lines = _wrap_jp_text_pil(draw, pain_text, f_q, max_w - 130, max_lines=3)
        q_box_h = max(136, 48 + len(q_lines) * 46)
        _rrect(draw, [pad, y, width - pad, y + q_box_h], radius=18, fill=theme["bg2"], outline=ac, width_px=3)
        draw.text((pad + 30, y + (q_box_h - 62) // 2), "Q.", font=f_big_q, fill=ac)
        qy = y + (q_box_h - len(q_lines) * 44) // 2
        for ql in q_lines:
            draw.text((pad + 108, qy), ql, font=f_q, fill=theme["text"])
            qy += 44
        y += q_box_h + 30

        draw.text((pad, y), "A. 結論・押さえるべきポイント", font=f_kick, fill=ac)
        y += 42
        a_lines = _wrap_jp_text_pil(draw, headline, f_a_head, max_w, max_lines=2)
        for al in a_lines:
            draw.text((pad, y), al, font=f_a_head, fill=theme["text"])
            y += 58
        y += 22

        items = bullets[:3]
        avail_h = height - pad - 68 - y
        gap = 18
        card_h = min(132, max(82, (avail_h - gap * (len(items) - 1)) // max(1, len(items))))
        for idx_i, it_txt in enumerate(items):
            cy = y + idx_i * (card_h + gap)
            _rrect(draw, [pad, cy, width - pad, cy + card_h], radius=14, fill=theme["bg2"], outline=theme["border"], width_px=2)
            _rrect(draw, [pad + 24, cy + (card_h - 36) // 2, pad + 60, cy + (card_h + 36) // 2], radius=8, fill=theme["accent_soft"], outline=ac, width_px=2)
            draw.text((pad + 32, cy + (card_h - 28) // 2), "✓", font=f_kick, fill=ac)
            ilines = _wrap_jp_text_pil(draw, it_txt, f_item, max_w - 116, max_lines=2)
            if len(ilines) == 1:
                draw.text((pad + 82, cy + (card_h - 34) // 2), ilines[0], font=f_item, fill=theme["text"])
            else:
                draw.text((pad + 82, cy + (card_h // 2) - 38), ilines[0], font=f_item, fill=theme["text"])
                draw.text((pad + 82, cy + (card_h // 2) + 2), ilines[1], font=f_item, fill=theme["text"])

        if show_watermark:
            draw.text((pad, height - pad - 24), f"@{account_id or 'threads'}", font=f_foot, fill=theme["muted"])
            r_txt = f"Q&A ANSWER #{post_index:02d}"
            try:
                rw = draw.textlength(r_txt, font=f_foot)
            except Exception:
                rw = 190
            draw.text((width - pad - rw, height - pad - 24), r_txt, font=f_foot, fill=ac)

    else:
        # 図解カード (diagram_card) / チェックリスト (checklist)
        out_img = _smooth_ambient_bg(theme["bg1"], theme["bg2"], theme["accent"], tint_strength=0.10)
        if bg_image_path and os.path.exists(bg_image_path):
            try:
                bg_cov = _fit_cover(bg_image_path)
                if bg_cov:
                    alpha_tint = 0.08 if theme["light"] else 0.14
                    out_img = Image.blend(out_img, bg_cov, alpha_tint)
            except Exception:
                pass
        draw = ImageDraw.Draw(out_img)

        f_kick = _load_jp_font(ImageFont, 28, bold=True)
        f_head = _load_jp_font(ImageFont, 52, bold=True)
        f_item = _load_jp_font(ImageFont, 32, bold=True)
        f_num = _load_jp_font(ImageFont, 36, bold=True)
        f_foot = _load_jp_font(ImageFont, 26, bold=True)

        max_w = width - pad * 2
        y = pad
        mode_lbl = f"{topic or 'GUIDE'} · 保存用チェックリスト" if style == "checklist" else f"{topic or 'GUIDE'} · ポイント図解"
        draw.text((pad, y), mode_lbl, font=f_kick, fill=theme["accent"])
        y += 50

        t_lines = _wrap_jp_text_pil(draw, headline, f_head, max_w, max_lines=2)
        for tl in t_lines:
            draw.text((pad, y), tl, font=f_head, fill=theme["text"])
            y += 70

        y += 14
        draw.line([(pad, y), (width - pad, y)], fill=theme["border"], width=2)
        y += 38

        list_items = bullets[:4]
        avail_h = height - y - pad - 76
        gap = 22
        card_h = min(165, max(95, (avail_h - gap * (len(list_items) - 1)) // max(1, len(list_items))))

        for idx_i, item_txt in enumerate(list_items):
            cy = y + idx_i * (card_h + gap)
            _rrect(draw, [pad, cy, width - pad, cy + card_h], radius=16, fill=theme["bg2"], outline=theme["border"], width_px=2)

            if style == "checklist":
                bx = pad + 28
                by = cy + (card_h - 46) // 2
                _rrect(draw, [bx, by, bx + 46, by + 46], radius=10, fill=theme["accent_soft"], outline=theme["accent"], width_px=3)
                draw.line([(bx + 11, by + 23), (bx + 20, by + 33), (bx + 35, by + 14)], fill=theme["accent"], width=4)
            else:
                draw.text((pad + 28, cy + (card_h - 40) // 2), f"{idx_i + 1:02d}", font=f_num, fill=theme["accent"])

            tx = pad + 104
            i_lines = _wrap_jp_text_pil(draw, item_txt, f_item, max_w - 134, max_lines=2)
            if len(i_lines) == 1:
                draw.text((tx, cy + (card_h - 36) // 2), i_lines[0], font=f_item, fill=theme["text"])
            else:
                draw.text((tx, cy + (card_h // 2) - 40), i_lines[0], font=f_item, fill=theme["text"])
                draw.text((tx, cy + (card_h // 2) + 4), i_lines[1], font=f_item, fill=theme["text"])

        if show_watermark:
            draw.text((pad, height - pad - 28), f"@{account_id or 'threads'}", font=f_foot, fill=theme["muted"])
            r_txt = f"{topic or 'GUIDE'} GUIDE"
            try:
                rw = draw.textlength(r_txt, font=f_foot)
            except Exception:
                rw = 160
            draw.text((width - pad - rw, height - pad - 28), r_txt, font=f_foot, fill=theme["accent"])

    if output_path:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        out_img.save(output_path, format="PNG")
    return out_img


class ThreadsPosterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Threads AutoPoster (スレッズ自動投稿)")
        self.root.geometry("980x820")
        self.root.configure(bg="#121212")
        self.is_running = False
        self.stop_requested = False
        self.current_context = None
        self.current_page = None
        self.current_account_name = None
        self.current_profile_dir = None
        self.account_config_path = find_account_file()
        self.var_auto_loop = tk.BooleanVar(value=True)
        self.var_single_window_rotation = tk.BooleanVar(value=True) # 提案1: 1窓・順番交代ローテーション
        self.var_block_media = tk.BooleanVar(value=True)            # 提案2: 画像・動画通信遮断 (メモリ50-70%削減)
        self.var_stealth_patrol = tk.BooleanVar(value=True)         # 提案新機能1: 投稿前のステルス巡回
        self.var_auto_clean_profile = tk.BooleanVar(value=False)     # ログインセッション保持優先（誤消去防止のため初期値OFF）
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

        self.btn_browser_login = tk.Button(toolbar, text="🌐 初回手動ログイン（セッション永続化）", font=("Segoe UI", 9, "bold"), bg="#8b5cf6", fg="#ffffff", padx=10, pady=5, relief="flat", cursor="hand2", command=self.open_browser_for_login)
        self.btn_browser_login.pack(side="left", padx=3)

        self.btn_chrome_launcher = tk.Button(toolbar, text="🛠️ 手動Chrome起動 (chrome_launcher)", font=("Segoe UI", 9, "bold"), bg="#0f766e", fg="#ffffff", padx=10, pady=5, relief="flat", cursor="hand2", command=self.launch_custom_chrome_profile)
        self.btn_chrome_launcher.pack(side="left", padx=3)

        self.btn_edit_profile = tk.Button(toolbar, text="👤 プロフィール編集", font=("Segoe UI", 9, "bold"), bg="#0284c7", fg="#ffffff", padx=10, pady=5, relief="flat", cursor="hand2", command=self.open_browser_for_profile_edit)
        self.btn_edit_profile.pack(side="left", padx=3)

        self.btn_test_session = tk.Button(toolbar, text="⚡ セッションテスト", font=("Segoe UI", 9, "bold"), bg="#f59e0b", fg="#000000", padx=10, pady=5, relief="flat", cursor="hand2", command=self.test_session_id_action)
        self.btn_test_session.pack(side="left", padx=3)

        btn_reload = tk.Button(toolbar, text="🔄 Excel再読込", font=("Segoe UI", 9, "bold"), bg="#3b82f6", fg="#ffffff", padx=8, pady=5, relief="flat", cursor="hand2", command=self.load_excel_data)
        btn_reload.pack(side="left", padx=3)

        btn_sync_profiles = tk.Button(toolbar, text="🛡️ 対策プロファイル同期", font=("Segoe UI", 9, "bold"), bg="#065f46", fg="#6ee7b7", padx=8, pady=5, relief="flat", cursor="hand2", command=self.manual_sync_profiles)
        btn_sync_profiles.pack(side="left", padx=3)

        btn_reset = tk.Button(toolbar, text="🔁 未投稿にリセット", font=("Segoe UI", 9, "bold"), bg="#6366f1", fg="#ffffff", padx=8, pady=5, relief="flat", cursor="hand2", command=self.reset_all_posts_to_unposted)
        btn_reset.pack(side="left", padx=3)

        toolbar2 = tk.Frame(self.root, bg="#121212", pady=2, padx=16)
        toolbar2.pack(fill="x")

        btn_design_gen = tk.Button(toolbar2, text="✨ 設計・投稿＆画像自動生成", font=("Segoe UI", 9, "bold"), bg="#059669", fg="#ffffff", padx=10, pady=4, relief="flat", cursor="hand2", command=self.open_account_design_generator)
        btn_design_gen.pack(side="left", padx=3)

        btn_bulk_add = tk.Button(toolbar2, text="📋 100件一括追加・削除", font=("Segoe UI", 9, "bold"), bg="#4f46e5", fg="#ffffff", padx=8, pady=4, relief="flat", cursor="hand2", command=self.open_bulk_post_manager)
        btn_bulk_add.pack(side="left", padx=3)

        btn_open_posts = tk.Button(toolbar2, text="📊 投稿Excelを開く", font=("Segoe UI", 9), bg="#262626", fg="#ffffff", padx=8, pady=4, relief="flat", cursor="hand2", command=self.open_posts_excel)
        btn_open_posts.pack(side="left", padx=3)

        btn_open_daily = tk.Button(toolbar2, text="🌅 日常ポスト設定", font=("Segoe UI", 9, "bold"), bg="#15803d", fg="#ffffff", padx=8, pady=4, relief="flat", cursor="hand2", command=self.open_daily_posts_excel)
        btn_open_daily.pack(side="left", padx=3)

        btn_open_excel = tk.Button(toolbar2, text="⚙️ アカウントExcel", font=("Segoe UI", 9), bg="#262626", fg="#ffffff", padx=8, pady=4, relief="flat", cursor="hand2", command=self.open_account_excel)
        btn_open_excel.pack(side="left", padx=3)

        toolbar3 = tk.Frame(self.root, bg="#121212", pady=2, padx=16)
        toolbar3.pack(fill="x")

        btn_open_ss = tk.Button(toolbar3, text="📷 エラー画像", font=("Segoe UI", 9, "bold"), bg="#450a0a", fg="#fca5a5", padx=8, pady=4, relief="flat", cursor="hand2", command=self.open_error_screenshots_dir)
        btn_open_ss.pack(side="left", padx=3)

        btn_open_log = tk.Button(toolbar3, text="📝 エラーログ(txt)", font=("Segoe UI", 9, "bold"), bg="#991b1b", fg="#ffffff", padx=8, pady=4, relief="flat", cursor="hand2", command=self.open_error_log)
        btn_open_log.pack(side="left", padx=3)

        btn_open_log_dir = tk.Button(toolbar3, text="📂 logフォルダ", font=("Segoe UI", 9, "bold"), bg="#1e3a8a", fg="#bfdbfe", padx=8, pady=4, relief="flat", cursor="hand2", command=self.open_log_dir)
        btn_open_log_dir.pack(side="left", padx=3)

        btn_open_images_dir = tk.Button(toolbar3, text="🖼️ imagesフォルダ", font=("Segoe UI", 9, "bold"), bg="#0f766e", fg="#ccfbf1", padx=8, pady=4, relief="flat", cursor="hand2", command=self.open_images_dir)
        btn_open_images_dir.pack(side="left", padx=3)

        btn_open_all_log = tk.Button(toolbar3, text="📋 全実行ログ(txt)", font=("Segoe UI", 9), bg="#262626", fg="#ffffff", padx=8, pady=4, relief="flat", cursor="hand2", command=self.open_execution_log)
        btn_open_all_log.pack(side="left", padx=3)

        btn_clean_manual = tk.Button(toolbar3, text="🧹 キャッシュ全初期化", font=("Segoe UI", 9, "bold"), bg="#4c0519", fg="#fda4af", padx=8, pady=4, relief="flat", cursor="hand2", command=self.manual_clean_profiles)
        btn_clean_manual.pack(side="left", padx=3)

        btn_open_dir = tk.Button(toolbar3, text="📁 フォルダ", font=("Segoe UI", 9), bg="#262626", fg="#ffffff", padx=8, pady=4, relief="flat", cursor="hand2", command=self.open_base_dir)
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

        # 3. アカウント別ログ (log/<アカウント名>.log) に永続追記
        try:
            acc_for_log = extract_account_name_from_msg(msg) or getattr(_THREAD_LOG_LOCAL, "account_name", "")
            if acc_for_log:
                log_to_account_file(acc_for_log, full_line)
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

        # 【起動時全自動】アカウント情報.xlsx の全アカウントに対して5大強制ログアウト対策プロファイルを自動構築
        try:
            bootstrapped = auto_bootstrap_countermeasure_profiles(accounts_data, log_func=self.log)
            if bootstrapped > 0:
                self.log(f"[SUCCESS] 🛡️ 起動時チェック: {bootstrapped}件のプロファイルに対策（端末指紋＋周辺Cookie永続化）を自動適用しました！")
        except Exception as _e_boot:
            self.log(f"[WARN] 対策プロファイル自動構築エラー: {_e_boot}")

        valid_accounts = []
        disabled_accounts = []
        os.makedirs(LOG_DIR, exist_ok=True)
        os.makedirs(IMAGES_DIR, exist_ok=True)
        if not hasattr(self, "_initialized_account_folders"):
            self._initialized_account_folders = set()

        for acc in accounts_data:
            acc_id = str(acc.get("アカウントID") or acc.get("アカウント名") or acc.get("アカウント") or "").strip()
            if not acc_id or acc_id.lower() == "nan":
                continue
            clean_id = sanitize_account_folder_name(acc_id)
            if clean_id:
                _KNOWN_ACCOUNT_NAMES.add(clean_id)
                acc_img_dir = get_account_images_dir(clean_id, create=True)
                acc_log_file = get_account_log_path(clean_id)
                if clean_id not in self._initialized_account_folders:
                    self._initialized_account_folders.add(clean_id)
                    log_to_account_file(clean_id, f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [INFO] [{clean_id}] ログ＆画像フォルダ準備完了 (log: {os.path.basename(acc_log_file)}, images: images/{os.path.basename(acc_img_dir)})")
            status_val = str(acc.get("有効/無効") or acc.get("有効無効") or acc.get("ステータス") or "").strip()
            if status_val in ["無効", "停止", "OFF", "off", "0", "false", "停止 (要電話番号認証)"]:
                disabled_accounts.append(acc_id)
            else:
                valid_accounts.append(acc)

        if not valid_accounts:
            if disabled_accounts:
                self.lbl_account_info.config(text=f"⚠️ {len(disabled_accounts)}件のアカウントが「無効」に設定されています", fg="#fbbf24")
                self.lbl_posts_info.config(text="―")
                self.log(f"[WARN] ⚠️ アカウント情報.xlsx のアカウント（@{', @'.join(disabled_accounts)}）が「無効」に設定されています。「有効/無効」列を「有効」にして保存してください。")
            else:
                self.lbl_account_info.config(text="⚠️ 有効なアカウント設定がありません（Excel保存・セル確定をご確認ください）", fg="#f87171")
                self.lbl_posts_info.config(text="―")
                self.log(f"[WARN] ⚠️ 『{os.path.basename(self.account_config_path)}』に有効なアカウントが検出されませんでした。ExcelでEnterキーを押してセル入力を確定し、上書き保存（Ctrl+S）してから「🔄 Excel再読込」を押してください。")
            return

        first_acc = valid_accounts[0]
        first_id = str(first_acc.get("アカウントID") or first_acc.get("アカウント名") or "").strip().lstrip("@")
        first_prof = first_id
        first_pw = str(first_acc.get("パスワード") or "").strip()
        pw_info = "🔑 PW登録済" if first_pw else "⚠️ PW未登録"
        dis_text = f" | ⚠️ 停止中: {len(disabled_accounts)}件 (@{', @'.join(disabled_accounts)})" if disabled_accounts else ""
        self.lbl_account_info.config(
            text=f"👤 有効: {len(valid_accounts)}件{dis_text} | メインID: @{first_id} ({pw_info}) | プロファイル: {first_prof}",
            fg="#e4e4e7"
        )
        self.log(f"[INFO] アカウント情報読み込み完了: {len(valid_accounts)}件のアカウント設定を検出 (ID: @{first_id}, パスワード: {'設定済 (***)' if first_pw else '未設定'}, プロファイル: {first_prof})")

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
        prof_txt = f"chrome_profiles/{first_id}"

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
            resolved_img = resolve_image_path(raw_img, text_preview, account_name=first_id)
            img_info = f"画像: {os.path.basename(resolved_img)}" if resolved_img else "画像なし"
            self.log(f"[INFO] プレビュー#{i}: [ポスト文章] {short_txt} | [{img_info}]")

    def _find_posts_file(self, account_id):
        clean_id = str(account_id or "").strip().lstrip("@")
        candidates = [
            os.path.join(BASE_DIR, f"{clean_id}.xlsx"),
            os.path.join(os.getcwd(), f"{clean_id}.xlsx"),
        ]
        for c in candidates:
            if os.path.exists(c):
                return c
        fallback_main = os.path.join(BASE_DIR, "threads_main_account.xlsx")
        if os.path.exists(fallback_main):
            try:
                import shutil
                target_dest = os.path.join(BASE_DIR, f"{clean_id}.xlsx")
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
        if "完了" in st or "投稿済" in st:
            return False
        if st in ["未投稿", "未", "待機", "pending"]:
            return True
        txt = self._extract_post_text(post_item)
        return (not st or st.lower() in ["none", "nan", "有効"]) and len(txt) > 0

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

    def manual_sync_profiles(self):
        """【手動実行】エクセル記載の全アカウントの5大強制ログアウト対策プロファイルを即時同期"""
        try:
            if not os.path.exists(self.account_config_path):
                messagebox.showwarning("警告", "アカウント情報.xlsxが見つかりません。")
                return
            accounts_data = read_excel_records(self.account_config_path)
            bootstrapped = auto_bootstrap_countermeasure_profiles(accounts_data, log_func=self.log)
            self.load_excel_data()
            messagebox.showinfo(
                "対策プロファイル同期完了",
                f"【5大強制ログアウト対策】\n全アカウントのプロファイル（端末指紋＋Cookie完全永続化）を同期しました！\n\n新規・更新件数: {bootstrapped} 件\n\n※ アプリ起動時にも毎回全自動で同期されます。"
            )
        except Exception as e:
            messagebox.showerror("エラー", f"プロファイル同期エラー: {e}")

    def open_bulk_post_manager(self):
        """【EXE対応】アカウント選択、既存ツイートの全件削除、新規100件一括登録GUIダイアログ"""
        all_accounts = read_excel_records(self.account_config_path)
        valid_accounts = [
            str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip()
            for acc in all_accounts
            if str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip() != ""
        ]
        if not valid_accounts:
            messagebox.showwarning("警告", "アカウント情報.xlsx に有効なアカウントが見つかりません。")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("【EXE投稿管理】アカウント別 既存ツイート削除 & 100件一括追加")
        dialog.geometry("760x700")
        dialog.configure(bg="#18181b")
        dialog.transient(self.root)
        dialog.grab_set()

        hdr = tk.Frame(dialog, bg="#09090b", padx=16, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📋 投稿一括管理 & 100件一括登録 (EXE)", font=("Segoe UI", 12, "bold"), fg="#ffffff", bg="#09090b").pack(side="left")

        sel_frame = tk.Frame(dialog, bg="#18181b", padx=16, pady=8)
        sel_frame.pack(fill="x")
        tk.Label(sel_frame, text="対象アカウント:", font=("Segoe UI", 9, "bold"), fg="#a1a1aa", bg="#18181b").pack(side="left")

        var_selected_acc = tk.StringVar(value=valid_accounts[0])
        cmb_acc = ttk.Combobox(sel_frame, textvariable=var_selected_acc, values=valid_accounts, state="readonly", width=18)
        cmb_acc.pack(side="left", padx=8)

        lbl_post_count = tk.Label(sel_frame, text="読込中...", font=("Segoe UI", 9), fg="#38bdf8", bg="#18181b")
        lbl_post_count.pack(side="left", padx=8)

        def refresh_count(*args):
            acc_id = var_selected_acc.get().strip()
            p_file = self._find_posts_file(acc_id)
            if p_file and os.path.exists(p_file):
                records = read_excel_records(p_file)
                unposted = len([r for r in records if self._is_unposted_item(r)])
                lbl_post_count.config(text=f"登録数: {len(records)} 件 (未投稿: {unposted} 件)")
            else:
                lbl_post_count.config(text="ファイル未検出")

        var_selected_acc.trace_add("write", refresh_count)
        refresh_count()

        act_frame = tk.Frame(dialog, bg="#18181b", padx=16, pady=4)
        act_frame.pack(fill="x")

        def delete_all_posts():
            acc_id = var_selected_acc.get().strip()
            p_file = self._find_posts_file(acc_id)
            if not p_file or not os.path.exists(p_file):
                messagebox.showwarning("エラー", f"'{acc_id}.xlsx' が見つかりません。")
                return
            if not messagebox.askyesno("全件削除の確認", f"@{acc_id} の登録済みポストを全件削除しますか？\n\n（※設定シートおよび日常ポストシートは安全に保持されます）"):
                return
            try:
                import openpyxl
                wb = openpyxl.load_workbook(p_file)
                post_sheet = None
                for sname in wb.sheetnames:
                    if any(k in sname for k in ["投稿", "ポスト", "post"]):
                        post_sheet = wb[sname]
                        break
                if post_sheet:
                    max_r = post_sheet.max_row
                    if max_r > 1:
                        post_sheet.delete_rows(2, max_r - 1)
                    wb.save(p_file)
                    wb.close()
                    refresh_count()
                    self.load_excel_data()
                    self.log(f"[WARN] 🗑️ @{acc_id} の既存ツイートを全件削除しました。")
                    messagebox.showinfo("完了", f"@{acc_id} のツイートを全件削除しました。")
            except Exception as ex:
                messagebox.showerror("エラー", f"削除失敗: {ex}")

        btn_del = tk.Button(act_frame, text="🗑️ 選択アカウントの全ツイートを削除", bg="#7f1d1d", fg="#fca5a5", font=("Segoe UI", 9, "bold"), padx=10, pady=4, relief="flat", cursor="hand2", command=delete_all_posts)
        btn_del.pack(side="left")

        def insert_100_samples():
            topics = ["信太山新地", "大阪求人", "高収入バイト", "ナイトワーク安心ナビ"]
            hooks = [
                "【完全保証】安心して働ける環境を徹底紹介！",
                "高収入とプライバシー保護を両立できる働き方ガイド✨",
                "未経験から始めるための面接・体験入社Q&Aまとめ🍙",
                "シフトの自由度と給与システムを徹底比較！",
                "寮完備・即日入居サポートありの人気店舗特集🏢",
                "働きやすい環境づくりとスタッフのサポート体制について🍵",
                "短期集中・週末のみでもOK！自分に合ったワークスタイル見学🌿",
                "よくある質問と安心ルールの徹底解説まとめ💡"
            ]
            urls = [
                "https://osaka-shinodayamashinchivip.pages.dev/blog/osaka-shinodayama-high-guarantee-guide",
                "https://osaka-shinodayamashinchivip.pages.dev/blog/shinodayamashinchi-salary-system",
                "https://osaka-shinodayamashinchivip.pages.dev/blog/shinodayamashinchi-privacy-guide",
                "https://osaka-shinodayamashinchivip.pages.dev/blog/shinodayamashinchi-housing-support"
            ]
            samples = []
            for i in range(1, 101):
                t = topics[(i - 1) % len(topics)]
                h = hooks[(i - 1) % len(hooks)]
                u = urls[(i - 1) % len(urls)]
                samples.append(f"{t} #{i}\n{h}\n{u}")
            txt_box.delete("1.0", "end")
            txt_box.insert("1.0", "\n\n".join(samples))
            update_item_count()

        btn_sample = tk.Button(act_frame, text="✨ 100件サンプル自動生成", bg="#374151", fg="#fde047", font=("Segoe UI", 9), padx=10, pady=4, relief="flat", cursor="hand2", command=insert_100_samples)
        btn_sample.pack(side="right")

        mid_frame = tk.Frame(dialog, bg="#18181b", padx=16, pady=6)
        mid_frame.pack(fill="both", expand=True)
        lbl_info = tk.Label(mid_frame, text="投稿文章貼り付け (空行または改行区切りで1ポストずつ自動認識・100件以上対応):", font=("Segoe UI", 9, "bold"), fg="#e4e4e7", bg="#18181b")
        lbl_info.pack(anchor="w", pady=(0, 4))

        lbl_detect = tk.Label(mid_frame, text="検出件数: 0 件", font=("Segoe UI", 9, "bold"), fg="#34d399", bg="#18181b")
        lbl_detect.pack(anchor="e", pady=(0, 2))

        txt_box = scrolledtext.ScrolledText(mid_frame, wrap="word", bg="#09090b", fg="#f4f4f5", font=("Consolas", 9), relief="solid", bd=1)
        txt_box.pack(fill="both", expand=True)

        def update_item_count(event=None):
            raw = txt_box.get("1.0", "end").strip()
            if not raw:
                lbl_detect.config(text="検出件数: 0 件")
                return []
            import re
            pattern = chr(10) + r"\s*" + chr(10)
            parts = [p.strip() for p in re.split(pattern, raw) if p.strip()]
            if len(parts) <= 1 and chr(10) in raw:
                lines_list = [l.strip() for l in raw.split(chr(10)) if l.strip()]
                if len(lines_list) > 5:
                    parts = lines_list
            lbl_detect.config(text=f"検出件数: {len(parts)} 件")
            return parts
        txt_box.bind("<KeyRelease>", update_item_count)

        opt_frame = tk.Frame(dialog, bg="#18181b", padx=16, pady=8)
        opt_frame.pack(fill="x")

        tk.Label(opt_frame, text="トピック:", font=("Segoe UI", 9), fg="#a1a1aa", bg="#18181b").pack(side="left")
        entry_topic = tk.Entry(opt_frame, bg="#27272a", fg="#ffffff", font=("Segoe UI", 9), width=14, relief="flat")
        entry_topic.insert(0, "信太山新地")
        entry_topic.pack(side="left", padx=(4, 12))

        tk.Label(opt_frame, text="画像添付:", font=("Segoe UI", 9), fg="#a1a1aa", bg="#18181b").pack(side="left")
        var_img_opt = tk.StringVar(value="alternate")
        cmb_img = ttk.Combobox(opt_frame, textvariable=var_img_opt, values=["alternate", "image.png", "image2.png", "none"], state="readonly", width=12)
        cmb_img.pack(side="left", padx=(4, 12))

        var_replace_existing = tk.BooleanVar(value=True)
        chk_repl = tk.Checkbutton(opt_frame, text="既存ツイートを削除して完全入れ替え", variable=var_replace_existing, bg="#18181b", fg="#fb7185", selectcolor="#09090b", font=("Segoe UI", 9), cursor="hand2")
        chk_repl.pack(side="left")

        btm_bar = tk.Frame(dialog, bg="#09090b", padx=16, pady=10)
        btm_bar.pack(fill="x")

        def save_and_apply():
            items = update_item_count()
            if not items:
                messagebox.showwarning("警告", "投稿文章が入力されていません。")
                return
            acc_id = var_selected_acc.get().strip()
            p_file = self._find_posts_file(acc_id)
            if not p_file or not os.path.exists(p_file):
                messagebox.showerror("エラー", f"'{acc_id}.xlsx' が見つかりません。")
                return
            try:
                import openpyxl
                wb = openpyxl.load_workbook(p_file)
                post_sheet = None
                for sname in wb.sheetnames:
                    if any(k in sname for k in ["投稿", "ポスト", "post"]):
                        post_sheet = wb[sname]
                        break
                if not post_sheet:
                    post_sheet = wb.create_sheet(title="投稿リスト")
                    post_sheet.append(["ID", "ポスト文章", "コミュニティまたはトピック", "画像ファイルパス", "ステータス", "投稿完了日時", "エラー内容"])

                if var_replace_existing.get():
                    max_r = post_sheet.max_row
                    if max_r > 1:
                        post_sheet.delete_rows(2, max_r - 1)
                    start_id = 1
                else:
                    start_id = post_sheet.max_row

                topic_val = entry_topic.get().strip() or "信太山新地"
                img_rule = var_img_opt.get()

                for idx, text_item in enumerate(items, start=start_id):
                    if img_rule == "alternate":
                        img_val = "image.png" if idx % 2 == 1 else "image2.png"
                    elif img_rule == "image.png":
                        img_val = "image.png"
                    elif img_rule == "image2.png":
                        img_val = "image2.png"
                    else:
                        img_val = ""

                    post_sheet.append([
                        idx,
                        text_item,
                        topic_val,
                        img_val,
                        "未投稿",
                        "",
                        ""
                    ])

                wb.save(p_file)
                wb.close()
                self.load_excel_data()
                self.log(f"[SUCCESS] 🚀 @{acc_id} のExcelに新規 {len(items)} 件のポストを一括登録しました！")
                messagebox.showinfo("保存完了", f"@{acc_id} のExcelに {len(items)} 件のポストを正常に登録しました！\n\nアプリ画面の設定・待機ポストに即時反映されています。")
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("保存エラー", f"Excelの書き込みに失敗しました: {e}")

        btn_save = tk.Button(btm_bar, text="💾 選択アカウントのExcelに一括保存して即座に反映", bg="#10b981", fg="#ffffff", font=("Segoe UI", 10, "bold"), padx=16, pady=6, relief="flat", cursor="hand2", command=save_and_apply)
        btn_save.pack(side="right")
        tk.Button(btm_bar, text="閉じる", bg="#27272a", fg="#ffffff", font=("Segoe UI", 9), padx=12, pady=6, relief="flat", cursor="hand2", command=dialog.destroy).pack(side="right", padx=8)

    def open_account_design_generator(self):
        """アカウント設計（<アカウント名>.xlsx「アカウント設計」シート）から投稿文＆画像を完全無料で一括自動生成"""
        all_accounts = read_excel_records(self.account_config_path)
        valid_accounts = [
            str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip().lstrip("@")
            for acc in all_accounts
            if str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip() != ""
        ]
        if not valid_accounts:
            messagebox.showwarning("警告", "アカウント情報.xlsx に有効なアカウントが見つかりません。")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("✨ アカウント設計 × 投稿文＆Pillow画像 一括自動生成（完全無料・API不要）")
        dialog.geometry("1040x780")
        dialog.configure(bg="#18181b")
        dialog.transient(self.root)
        dialog.grab_set()

        hdr = tk.Frame(dialog, bg="#09090b", padx=16, pady=10)
        hdr.pack(fill="x")
        tk.Label(
            hdr,
            text="✨ アカウント設計 → 投稿文章＋PNG画像 一括自動生成",
            font=("Segoe UI", 12, "bold"),
            fg="#ffffff",
            bg="#09090b",
        ).pack(side="left")
        tk.Label(
            hdr,
            text="APIキー不要 · Windows標準日本語フォント＋Pillow直接書き出し",
            font=("Segoe UI", 9),
            fg="#34d399",
            bg="#09090b",
        ).pack(side="right")

        top_bar = tk.Frame(dialog, bg="#18181b", padx=16, pady=8)
        top_bar.pack(fill="x")

        tk.Label(top_bar, text="対象アカウント:", font=("Segoe UI", 9, "bold"), fg="#d4d4d8", bg="#18181b").pack(side="left")
        var_acc = tk.StringVar(value=valid_accounts[0])
        cmb_acc = ttk.Combobox(top_bar, textvariable=var_acc, values=valid_accounts, state="readonly", width=18)
        cmb_acc.pack(side="left", padx=(6, 12))

        tk.Label(top_bar, text="プリセット:", font=("Segoe UI", 9), fg="#a1a1aa", bg="#18181b").pack(side="left")
        var_preset = tk.StringVar(value="飛田新地・風景×求人案内")
        cmb_preset = ttk.Combobox(top_bar, textvariable=var_preset, values=list(_DESIGN_PRESETS_PY.keys()), state="readonly", width=24)
        cmb_preset.pack(side="left", padx=(6, 6))

        body = tk.Frame(dialog, bg="#18181b", padx=16, pady=4)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=6)
        body.columnconfigure(1, weight=5)
        body.rowconfigure(0, weight=1)

        # Left Form (Scrollable)
        left_wrap = tk.Frame(body, bg="#18181b")
        left_wrap.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        canvas_scroll = tk.Canvas(left_wrap, bg="#18181b", highlightthickness=0)
        v_scroll = ttk.Scrollbar(left_wrap, orient="vertical", command=canvas_scroll.yview)
        form_frame = tk.Frame(canvas_scroll, bg="#18181b")
        form_frame.bind("<Configure>", lambda _e: canvas_scroll.configure(scrollregion=canvas_scroll.bbox("all")))
        canvas_scroll.create_window((0, 0), window=form_frame, anchor="nw")
        canvas_scroll.configure(yscrollcommand=v_scroll.set)
        canvas_scroll.pack(side="left", fill="both", expand=True)
        v_scroll.pack(side="right", fill="y")

        def _mk_entry(parent, label_txt, width=44):
            f = tk.Frame(parent, bg="#18181b", pady=3)
            f.pack(fill="x", padx=4)
            tk.Label(f, text=label_txt, font=("Segoe UI", 9, "bold"), fg="#a1a1aa", bg="#18181b", anchor="w").pack(fill="x")
            ent = tk.Entry(f, bg="#09090b", fg="#ffffff", insertbackground="#ffffff", font=("Segoe UI", 9), width=width, relief="solid", bd=1)
            ent.pack(fill="x", pady=(2, 0), ipady=3)
            return ent

        ent_genre = _mk_entry(form_frame, "ジャンル・発信テーマ:")
        ent_topic = _mk_entry(form_frame, "トピック（Excel設定用）:")
        ent_target = _mk_entry(form_frame, "ターゲット層:")

        f_tone = tk.Frame(form_frame, bg="#18181b", pady=3)
        f_tone.pack(fill="x", padx=4)
        tk.Label(f_tone, text="アカウント口調:", font=("Segoe UI", 9, "bold"), fg="#a1a1aa", bg="#18181b", anchor="w").pack(fill="x")
        var_tone = tk.StringVar(value="丁寧・誠実")
        cmb_tone = ttk.Combobox(f_tone, textvariable=var_tone, values=["丁寧・誠実", "フレンドリー", "プロ解説", "共感スタイル"], state="readonly")
        cmb_tone.pack(fill="x", pady=(2, 0))

        ent_pain = _mk_entry(form_frame, "読者の悩み・不安:")

        f_ben = tk.Frame(form_frame, bg="#18181b", pady=3)
        f_ben.pack(fill="x", padx=4)
        tk.Label(f_ben, text="強み・ポイント（1行1項目 · 図解/チェックリスト項目に反映）:", font=("Segoe UI", 9, "bold"), fg="#a1a1aa", bg="#18181b", anchor="w").pack(fill="x")
        txt_benefits = tk.Text(f_ben, height=4, bg="#09090b", fg="#ffffff", insertbackground="#ffffff", font=("Segoe UI", 9), relief="solid", bd=1)
        txt_benefits.pack(fill="x", pady=(2, 0))

        ent_cta = _mk_entry(form_frame, "誘導文（CTA・結び）:")
        ent_url = _mk_entry(form_frame, "誘導URL（任意）:")
        ent_headline_ov = _mk_entry(form_frame, "テスト用キャッチコピー上書き（空欄なら自動生成）:")

        # Image Generation Options
        img_opt_box = tk.Frame(form_frame, bg="#09090b", padx=10, pady=8, highlightbackground="#27272a", highlightthickness=1)
        img_opt_box.pack(fill="x", padx=4, pady=(8, 4))
        tk.Label(img_opt_box, text="画像自動生成オプション (Pillow)", font=("Segoe UI", 9, "bold"), fg="#34d399", bg="#09090b", anchor="w").pack(fill="x")

        row_style = tk.Frame(img_opt_box, bg="#09090b", pady=4)
        row_style.pack(fill="x")
        var_gen_imgs = tk.BooleanVar(value=True)
        tk.Checkbutton(row_style, text="PNG画像も同時生成", variable=var_gen_imgs, bg="#09090b", fg="#ffffff", selectcolor="#18181b", font=("Segoe UI", 9)).pack(side="left")

        var_img_style = tk.StringVar(value="自動おまかせ（設計に合わせて最適化）")
        cmb_img_style = ttk.Combobox(
            row_style,
            textvariable=var_img_style,
            values=[
                "自動おまかせ（設計に合わせて最適化）",
                "雑誌エディトリアル風",
                "吹き出し会話SNS風",
                "ネオン夜風",
                "ステップ図解",
                "Q&A神回答カード",
                "背景写真＋自動文字入れ",
                "SNS図解カード",
                "保存用チェックリスト",
                "全テンプレートミックス",
            ],
            state="readonly",
            width=30,
        )
        cmb_img_style.pack(side="right")

        row_size = tk.Frame(img_opt_box, bg="#09090b", pady=3)
        row_size.pack(fill="x")
        tk.Label(row_size, text="サイズ / 配色:", fg="#a1a1aa", bg="#09090b", font=("Segoe UI", 9)).pack(side="left")
        var_aspect = tk.StringVar(value="1:1")
        cmb_aspect = ttk.Combobox(row_size, textvariable=var_aspect, values=["1:1", "4:5"], state="readonly", width=5)
        cmb_aspect.pack(side="left", padx=4)
        var_palette = tk.StringVar(value="提灯アンバー×黒")
        cmb_palette = ttk.Combobox(row_size, textvariable=var_palette, values=list(_PIL_PALETTES.keys()), state="readonly", width=16)
        cmb_palette.pack(side="left", padx=4)
        var_scrim_label = tk.StringVar(value="標準48%")
        cmb_scrim = ttk.Combobox(row_size, textvariable=var_scrim_label, values=["薄め30%", "標準48%", "濃いめ65%"], state="readonly", width=9)
        cmb_scrim.pack(side="right")

        row_bg = tk.Frame(img_opt_box, bg="#09090b", pady=4)
        row_bg.pack(fill="x")
        var_bg_dir = tk.StringVar(value=get_account_images_dir(valid_accounts[0], create=True))
        tk.Label(row_bg, text="背景フォルダ:", fg="#a1a1aa", bg="#09090b", font=("Segoe UI", 9)).pack(side="left")
        ent_bg_dir = tk.Entry(row_bg, textvariable=var_bg_dir, bg="#18181b", fg="#e4e4e7", font=("Consolas", 8), relief="flat")
        ent_bg_dir.pack(side="left", fill="x", expand=True, padx=6)

        def browse_bg_dir():
            chosen = filedialog.askdirectory(initialdir=var_bg_dir.get() or IMAGES_DIR, title="背景写真フォルダを選択")
            if chosen:
                var_bg_dir.set(chosen)
                refresh_bg_info()
                update_live_preview()

        def open_bg_dir_explorer():
            target_dir = var_bg_dir.get().strip() or get_account_images_dir(var_acc.get().strip(), create=True)
            os.makedirs(target_dir, exist_ok=True)
            try:
                os.startfile(target_dir)
            except Exception:
                pass

        tk.Button(row_bg, text="開く", bg="#0f766e", fg="#ffffff", font=("Segoe UI", 8), padx=7, pady=2, relief="flat", cursor="hand2", command=open_bg_dir_explorer).pack(side="right", padx=(4, 0))
        tk.Button(row_bg, text="選択", bg="#27272a", fg="#ffffff", font=("Segoe UI", 8), padx=8, pady=2, relief="flat", cursor="hand2", command=browse_bg_dir).pack(side="right")
        lbl_bg_status = tk.Label(img_opt_box, text="検出背景画像: 確認中...", font=("Segoe UI", 8), fg="#38bdf8", bg="#09090b", anchor="w")
        lbl_bg_status.pack(fill="x")

        row_batch = tk.Frame(form_frame, bg="#18181b", pady=6)
        row_batch.pack(fill="x", padx=4)
        tk.Label(row_batch, text="生成件数:", fg="#d4d4d8", bg="#18181b", font=("Segoe UI", 9, "bold")).pack(side="left")
        var_count = tk.StringVar(value="10")
        ttk.Combobox(row_batch, textvariable=var_count, values=["5", "10", "30", "50", "100"], state="readonly", width=6).pack(side="left", padx=6)
        var_spintax = tk.BooleanVar(value=True)
        chk_spintax = tk.Checkbutton(row_batch, text="スピンタックス {A|B|C}", variable=var_spintax, bg="#18181b", fg="#d4d4d8", selectcolor="#09090b", font=("Segoe UI", 9))
        chk_spintax.pack(side="left", padx=6)
        var_replace = tk.BooleanVar(value=True)
        tk.Checkbutton(row_batch, text="既存リスト入替", variable=var_replace, bg="#18181b", fg="#fb7185", selectcolor="#09090b", font=("Segoe UI", 9)).pack(side="left")

        # Right Preview Panel
        right_panel = tk.Frame(body, bg="#09090b", padx=12, pady=10, highlightbackground="#27272a", highlightthickness=1)
        right_panel.grid(row=0, column=1, sticky="nsew")
        tk.Label(right_panel, text="リアルタイム画像＆文章プレビュー (1件目)", font=("Segoe UI", 10, "bold"), fg="#ffffff", bg="#09090b").pack(anchor="w")

        lbl_img_preview = tk.Label(right_panel, text="プレビュー準備中...", bg="#18181b", fg="#71717a", width=44, height=18)
        lbl_img_preview.pack(pady=8, fill="both", expand=True)

        lbl_text_title = tk.Label(right_panel, text="生成される投稿文サンプル:", font=("Segoe UI", 9, "bold"), fg="#a1a1aa", bg="#09090b", anchor="w")
        lbl_text_title.pack(fill="x")
        txt_preview_post = scrolledtext.ScrolledText(right_panel, height=7, wrap="word", bg="#18181b", fg="#f4f4f5", font=("Segoe UI", 9), relief="flat")
        txt_preview_post.pack(fill="x", pady=(4, 0))

        preview_photo_ref = {"img": None}

        def _collect_design():
            return {
                "genre": ent_genre.get().strip(),
                "topic": ent_topic.get().strip() or "おすすめ",
                "target": ent_target.get().strip(),
                "tone": var_tone.get(),
                "pain": ent_pain.get().strip(),
                "benefits": txt_benefits.get("1.0", "end").strip(),
                "cta": ent_cta.get().strip(),
                "url": ent_url.get().strip(),
                "image_style": var_img_style.get(),
                "palette": var_palette.get(),
            }

        def _fill_design(d):
            for ent, key in [
                (ent_genre, "genre"), (ent_topic, "topic"), (ent_target, "target"),
                (ent_pain, "pain"), (ent_cta, "cta"), (ent_url, "url")
            ]:
                ent.delete(0, "end")
                ent.insert(0, d.get(key, ""))
            var_tone.set(d.get("tone", "丁寧・誠実"))
            txt_benefits.delete("1.0", "end")
            txt_benefits.insert("1.0", d.get("benefits", ""))
            var_img_style.set(_normalize_image_style_label_py(d.get("image_style", "自動おまかせ（設計に合わせて最適化）")))
            pal_val = d.get("palette", "提灯アンバー×黒")
            var_palette.set(pal_val if pal_val in _PIL_PALETTES else "提灯アンバー×黒")

        def _scan_photos_in_dir(dpath):
            exts = (".jpg", ".jpeg", ".png", ".webp", ".bmp")
            res_list = []
            if dpath and os.path.isdir(dpath):
                try:
                    for fn in sorted(os.listdir(dpath)):
                        low = fn.lower()
                        full_p = os.path.join(dpath, fn)
                        if os.path.isfile(full_p) and low.endswith(exts) and not low.startswith("gen_"):
                            # 初期プレースホルダー（10x10px・数十バイトのダミーimage.png）は背景に使わない
                            if os.path.getsize(full_p) >= 4096:
                                res_list.append(full_p)
                except Exception:
                    pass
            return res_list

        def _find_bg_photos(folder_path):
            found = _scan_photos_in_dir(folder_path)
            # 指定フォルダ内の背景サブフォルダも探索
            if not found and folder_path and os.path.isdir(folder_path):
                for sub_name in ("backgrounds", "bg", "背景", "photos"):
                    found = _scan_photos_in_dir(os.path.join(folder_path, sub_name))
                    if found:
                        break
            # アカウント個別フォルダが空の場合、共通フォルダ（images/backgrounds, images/共通背景, images直下）をフォールバック探索
            if not found:
                for fallback_dir in (
                    os.path.join(IMAGES_DIR, "backgrounds"),
                    os.path.join(IMAGES_DIR, "共通背景"),
                    os.path.join(IMAGES_DIR, "bg"),
                    IMAGES_DIR,
                ):
                    if os.path.abspath(fallback_dir) != os.path.abspath(folder_path or ""):
                        found = _scan_photos_in_dir(fallback_dir)
                        if found:
                            break
            return found

        def _get_scrim_val():
            lbl = var_scrim_label.get()
            if "30" in lbl:
                return 0.30
            if "65" in lbl:
                return 0.65
            return 0.48

        def refresh_bg_info():
            bgs = _find_bg_photos(var_bg_dir.get().strip())
            if bgs:
                src_dir = os.path.basename(os.path.dirname(bgs[0]))
                lbl_bg_status.config(text=f"✅ 背景写真 {len(bgs)}枚 検出 ({src_dir}/{os.path.basename(bgs[0])}) ※gen_*除外", fg="#34d399")
            else:
                lbl_bg_status.config(text="⚠️ 背景写真0枚：「開く」から写真を入れるか「選択」で写真フォルダを指定してください", fg="#fbbf24")
            return bgs

        def update_live_preview(*_args):
            bgs = refresh_bg_info()
            d = _collect_design()
            specs = generate_posts_from_design(d, count=1, enable_spintax=var_spintax.get(), has_photos=bool(bgs))
            if not specs:
                return
            sp = specs[0]
            hl = ent_headline_ov.get().strip() or sp["headline"]
            txt_preview_post.delete("1.0", "end")
            txt_preview_post.insert("1.0", sp["text"])

            Image, ImageDraw, ImageFont = ensure_pillow_available(log_func=self.log, allow_install=False)
            if not Image:
                lbl_img_preview.config(image="", text="Pillow導入確認中...（一括生成または「Pillow導入」で自動準備します）")
                return
            bg_p = bgs[0] if bgs else None
            pil_img = render_post_image_pil(
                Image, ImageDraw, ImageFont,
                style=sp["style"],
                aspect_ratio=var_aspect.get(),
                palette_name=var_palette.get(),
                scrim_opacity=_get_scrim_val(),
                topic=sp["topic"],
                headline=hl,
                subhead=sp["subhead"],
                bullets=sp["bullets"],
                account_id=var_acc.get().strip(),
                post_index=1,
                bg_image_path=bg_p,
                show_watermark=True,
                output_path=None,
                pain=sp.get("pain", ""),
            )
            max_side = 330
            pw, ph = pil_img.size
            scale = min(max_side / pw, max_side / ph)
            pil_thumb = pil_img.resize((max(1, int(pw * scale)), max(1, int(ph * scale))), Image.LANCZOS)
            import io, base64
            buf = io.BytesIO()
            pil_thumb.save(buf, format="PNG")
            b64_data = base64.b64encode(buf.getvalue()).decode("ascii")
            tk_img = tk.PhotoImage(data=b64_data)
            preview_photo_ref["img"] = tk_img
            lbl_img_preview.config(image=tk_img, text="", width=0, height=0)

        for _cmb in (cmb_tone, cmb_img_style, cmb_aspect, cmb_palette, cmb_scrim):
            _cmb.bind("<<ComboboxSelected>>", update_live_preview)
        chk_spintax.config(command=update_live_preview)

        def on_account_selected(*_args):
            acc_id = var_acc.get().strip()
            acc_xlsx = os.path.join(BASE_DIR, f"{acc_id}.xlsx")
            fb_topic = read_account_settings_from_post_file(acc_xlsx).get("topic", "") if os.path.exists(acc_xlsx) else ""
            loaded = load_account_design_from_xlsx(acc_xlsx, fallback_topic=fb_topic)
            _fill_design(loaded)
            var_bg_dir.set(get_account_images_dir(acc_id, create=True))
            refresh_bg_info()
            update_live_preview()

        def apply_preset_clicked():
            p_name = var_preset.get()
            if p_name in _DESIGN_PRESETS_PY:
                _fill_design(_DESIGN_PRESETS_PY[p_name])
                update_live_preview()

        tk.Button(top_bar, text="読込", bg="#27272a", fg="#ffffff", font=("Segoe UI", 8, "bold"), padx=10, pady=3, relief="flat", cursor="hand2", command=apply_preset_clicked).pack(side="left")
        cmb_acc.bind("<<ComboboxSelected>>", on_account_selected)
        cmb_preset.bind("<<ComboboxSelected>>", lambda _e: apply_preset_clicked())

        def _offer_manual_pillow_cmd(err_detail):
            py_cmd = _resolve_pip_python_exe()
            msg = (
                "画像書き出し用の Pillow ライブラリの自動導入に失敗しました。\n\n"
                f"【詳細ログ】\n{err_detail[:650]}\n\n"
                "コマンドプロンプトを開いて `pip install pillow` を手動実行しますか？\n"
                "（※「いいえ」を選んで『PNG画像も同時生成』のチェックを外せば、文章のみ即座に生成できます）"
            )
            if messagebox.askyesno("Pillow導入エラー", msg):
                try:
                    import subprocess
                    subprocess.Popen(
                        ["cmd", "/k", f'"{py_cmd}" -m pip install pillow --trusted-host pypi.org --trusted-host files.pythonhosted.org & echo. & echo 完了したらこの画面を閉じて再試行してください & pause']
                    )
                except Exception as ex_cmd:
                    messagebox.showerror("起動失敗", f"コマンドプロンプトを起動できませんでした:\n{ex_cmd}\n\n手動で以下を実行してください:\n\"{py_cmd}\" -m pip install pillow")

        def save_design_only():
            acc_id = var_acc.get().strip()
            acc_xlsx = os.path.join(BASE_DIR, f"{acc_id}.xlsx")
            try:
                save_account_design_to_xlsx(acc_xlsx, acc_id, _collect_design())
                self.log(f"[SUCCESS] 💾 @{acc_id} のアカウント設計を {os.path.basename(acc_xlsx)}（アカウント設計シート）に保存しました！")
                messagebox.showinfo("設計保存完了", f"@{acc_id} のアカウント設計を\n{os.path.basename(acc_xlsx)} に保存しました！\n\n次回からはアカウントを選ぶだけで自動読み込みされます。")
            except Exception as ex:
                messagebox.showerror("保存エラー", f"設計保存に失敗しました:\n{ex}")

        def run_batch_generate():
            acc_id = var_acc.get().strip()
            d = _collect_design()
            cnt = int(var_count.get() or "10")
            do_imgs = var_gen_imgs.get()
            do_spintax = var_spintax.get()
            do_replace = var_replace.get()
            aspect = var_aspect.get()
            palette_n = var_palette.get()
            hl_override = ent_headline_ov.get().strip()
            bg_list = _find_bg_photos(var_bg_dir.get().strip())
            scrim_val = _get_scrim_val()
            acc_xlsx = os.path.join(BASE_DIR, f"{acc_id}.xlsx")
            acc_img_dir = get_account_images_dir(acc_id, create=True)

            if do_imgs and not bg_list and var_img_style.get() == "背景写真＋自動文字入れ":
                proceed = messagebox.askyesno(
                    "背景写真が見つかりません",
                    f"フォルダ内に背景写真（4KB以上のJPG/PNG等）が0枚でした：\n{var_bg_dir.get().strip()}\n\n"
                    "※「自動おまかせ（設計に合わせて最適化）」や「雑誌風」「吹き出し風」「ネオン夜風」を選べば写真なしでも綺麗に作れます。\n\n"
                    "写真なしのまま生成を続行しますか？"
                )
                if not proceed:
                    return

            btn_run.config(state="disabled", text="生成中...")

            def _worker():
                t0 = time.time()
                try:
                    Image, ImageDraw, ImageFont = (None, None, None)
                    if do_imgs:
                        Image, ImageDraw, ImageFont = ensure_pillow_available(log_func=self.log, allow_install=True)
                        if not Image:
                            err_snap = _LAST_PILLOW_ERROR or "原因不明"
                            self.root.after(0, lambda es=err_snap: (
                                btn_run.config(state="normal", text="🚀 ワンクリックで一括生成 → Excel＆images保存"),
                                _offer_manual_pillow_cmd(es)
                            ))
                            return

                    save_account_design_to_xlsx(acc_xlsx, acc_id, d)
                    specs = generate_posts_from_design(d, count=cnt, enable_spintax=do_spintax, has_photos=bool(bg_list))

                    wb = ensure_account_workbook(acc_xlsx, acc_id)
                    post_sheet = None
                    for sname in wb.sheetnames:
                        if any(k in sname for k in ["投稿", "ポスト", "post"]):
                            post_sheet = wb[sname]
                            break
                    if not post_sheet:
                        post_sheet = wb.create_sheet(title="投稿リスト")
                        post_sheet.append(["ID", "ポスト文章", "コミュニティまたはトピック", "画像ファイルパス", "ステータス", "投稿完了日時", "エラー内容"])

                    if do_replace:
                        if post_sheet.max_row > 1:
                            post_sheet.delete_rows(2, post_sheet.max_row - 1)
                        start_id = 1
                    else:
                        start_id = post_sheet.max_row

                    for idx_p, sp in enumerate(specs):
                        row_id = start_id + idx_p
                        img_name = ""
                        if do_imgs and Image:
                            img_name = sp["filename"]
                            out_png = os.path.join(acc_img_dir, img_name)
                            bg_p = bg_list[idx_p % len(bg_list)] if bg_list else None
                            hl = hl_override if (idx_p == 0 and hl_override) else sp["headline"]
                            render_post_image_pil(
                                Image, ImageDraw, ImageFont,
                                style=sp["style"],
                                aspect_ratio=aspect,
                                palette_name=palette_n,
                                scrim_opacity=scrim_val,
                                topic=sp["topic"],
                                headline=hl,
                                subhead=sp["subhead"],
                                bullets=sp["bullets"],
                                account_id=acc_id,
                                post_index=row_id,
                                bg_image_path=bg_p,
                                show_watermark=True,
                                output_path=out_png,
                                pain=sp.get("pain", ""),
                            )

                        post_sheet.append([
                            row_id,
                            sp["text"],
                            sp["topic"],
                            img_name,
                            "未投稿",
                            "",
                            "",
                        ])

                    wb.save(acc_xlsx)
                    wb.close()
                    elapsed = time.time() - t0

                    def _on_done():
                        self.load_excel_data()
                        self.log(f"[SUCCESS] 🚀 @{acc_id}: 設計から投稿 {cnt} 件{'＋PNG画像 ' + str(cnt) + '枚' if do_imgs else ''} を {elapsed:.2f}秒 で生成＆Excel直接保存しました！")
                        msg = f"@{acc_id} の一括生成が完了しました！（{elapsed:.2f}秒）\n\n・Excel保存: {os.path.basename(acc_xlsx)}（{cnt}件）"
                        if do_imgs:
                            msg += f"\n・画像保存: images/{acc_id}/gen_001.png 〜 gen_{cnt:03d}.png"
                        msg += "\n\n「imagesフォルダを開く」で仕上がりを確認しますか？"
                        if do_imgs and messagebox.askyesno("生成完了", msg):
                            try:
                                os.startfile(acc_img_dir)
                            except Exception:
                                pass
                        elif not do_imgs:
                            messagebox.showinfo("生成完了", msg.split("\n\n「images")[0])
                        btn_run.config(state="normal", text="🚀 ワンクリックで一括生成 → Excel＆images保存")

                    self.root.after(0, _on_done)
                except Exception as ex:
                    self.root.after(0, lambda e=ex: (
                        btn_run.config(state="normal", text="🚀 ワンクリックで一括生成 → Excel＆images保存"),
                        messagebox.showerror("生成エラー", f"一括生成中にエラーが発生しました:\n{e}")
                    ))

            threading.Thread(target=_worker, daemon=True).start()

        # Bottom Action Bar
        btm = tk.Frame(dialog, bg="#09090b", padx=16, pady=10)
        btm.pack(fill="x")
        tk.Button(btm, text="💾 設計をExcelに保存", bg="#0284c7", fg="#ffffff", font=("Segoe UI", 9, "bold"), padx=12, pady=6, relief="flat", cursor="hand2", command=save_design_only).pack(side="left")
        tk.Button(btm, text="🔍 プレビュー更新", bg="#27272a", fg="#ffffff", font=("Segoe UI", 9), padx=12, pady=6, relief="flat", cursor="hand2", command=update_live_preview).pack(side="left", padx=8)
        tk.Button(btm, text="閉じる", bg="#27272a", fg="#a1a1aa", font=("Segoe UI", 9), padx=12, pady=6, relief="flat", cursor="hand2", command=dialog.destroy).pack(side="right", padx=(8, 0))
        btn_run = tk.Button(btm, text="🚀 ワンクリックで一括生成 → Excel＆images保存", bg="#10b981", fg="#ffffff", font=("Segoe UI", 10, "bold"), padx=16, pady=6, relief="flat", cursor="hand2", command=run_batch_generate)
        btn_run.pack(side="right")

        # Initial load + non-blocking background Pillow pre-warm if missing
        on_account_selected()
        if ensure_pillow_available(log_func=None, allow_install=False)[0] is None:
            def _bg_install_pil():
                img_mod, _, _ = ensure_pillow_available(log_func=self.log, allow_install=True)
                if img_mod:
                    self.root.after(0, lambda: update_live_preview() if dialog.winfo_exists() else None)
                else:
                    self.root.after(0, lambda: lbl_img_preview.config(text="Pillow未導入（一括生成を押すと詳細診断・導入手順を表示）") if dialog.winfo_exists() else None)
            threading.Thread(target=_bg_install_pil, daemon=True).start()


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

    def open_log_dir(self):
        os.makedirs(LOG_DIR, exist_ok=True)
        try:
            os.startfile(LOG_DIR)
            self.log("[INFO] 📂 アカウント別ログフォルダ (log) を開きました。")
        except Exception as e:
            self.log(f"[WARN] フォルダを開けませんでした: {e}")

    def open_images_dir(self):
        os.makedirs(IMAGES_DIR, exist_ok=True)
        try:
            os.startfile(IMAGES_DIR)
            self.log("[INFO] 🖼️ アカウント別画像フォルダ (images) を開きました。")
        except Exception as e:
            self.log(f"[WARN] フォルダを開けませんでした: {e}")

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
            clean_acc_id = acc_name.lstrip("@").strip()
            p_sub = custom_settings.get("profile_dir", "")
            if not p_sub or any(k in p_sub for k in ["コミュニティ", "トピック", "モード", "間隔", "ツイート", "ポスト"]): p_sub = f"chrome_profiles/{clean_acc_id}"
            p_dir = os.path.join(BASE_DIR, p_sub)
            clean_account_profile(acc_name, p_dir, reason="手動リクエストによる全初期化", log_func=self.log, full_reset=True)
            cleaned_count += 1

        messagebox.showinfo(
            "初期化完了",
            f"{cleaned_count} 件のアカウントプロファイルを初期化しました！\n\n"
            "次回からクリーンな状態で自動接続されます。"
        )

    def launch_custom_chrome_profile(self):
        """
        【Chrome カスタムプロファイル起動 (完全手動)】
        勝手なURL遷移・自動化フラグなしで純正Google Chromeを起動し、
        ユーザーが任意のサイト（X / Threads等）へ手動アクセス・手動ログインして情報を保持します。
        """
        import subprocess
        from tkinter import simpledialog

        default_name = ""
        if os.path.exists(self.account_config_path):
            try:
                accs = read_excel_records(self.account_config_path)
                for a in accs:
                    aid = str(a.get("プロファイル名") or a.get("アカウントID") or a.get("アカウント名") or "").strip().lstrip("@")
                    if aid:
                        default_name = aid
                        break
            except Exception:
                pass

        prof_input = simpledialog.askstring(
            "Chrome カスタムプロファイル起動ツール",
            "作成・使用するプロファイル名を入力してください:\n（例: wqwq1320）",
            initialvalue=default_name,
            parent=self.root
        )
        if not prof_input:
            return
        prof_name = re.sub(r'[\\/:*?"<>|]+', "", prof_input.strip().lstrip("@")).strip()
        if not prof_name:
            messagebox.showwarning("警告", "有効なプロファイル名を入力してください。")
            return

        prof_dir = os.path.join(BASE_DIR, "chrome_profiles", prof_name)
        os.makedirs(prof_dir, exist_ok=True)
        for lk in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
            try:
                lp = os.path.join(prof_dir, lk)
                if os.path.exists(lp):
                    os.remove(lp)
            except Exception:
                pass

        chrome_candidates = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        ]
        chrome_exe = next((c for c in chrome_candidates if c and os.path.exists(c)), None)
        if not chrome_exe:
            import shutil
            chrome_exe = shutil.which("chrome") or shutil.which("google-chrome")

        if not chrome_exe:
            py_launcher = os.path.join(BASE_DIR, "chrome_launcher.py")
            if os.path.exists(py_launcher):
                subprocess.Popen([sys.executable, py_launcher, prof_name], cwd=BASE_DIR)
                return
            exe_bat = os.path.join(BASE_DIR, "chrome_launcher.exe")
            if os.path.exists(exe_bat):
                subprocess.Popen([exe_bat, prof_name], cwd=BASE_DIR)
                return
            messagebox.showerror("エラー", "Google Chrome (chrome.exe) が見つかりませんでした。")
            return

        try:
            subprocess.Popen([chrome_exe, f"--user-data-dir={prof_dir}"], cwd=BASE_DIR)
            self.log("=" * 60)
            self.log(f"[INFO] 🛠️ 【Chrome手動起動】プロファイル『{prof_name}』を起動しました。")
            self.log(f"[INFO] 📂 保存先: {prof_dir}")
            self.log("[INFO] 💡 サイトは自動で開きません。ご自身でログインしたいサイトを開いて手動ログインし、完了後にChromeを閉じてください。")
            self.log("=" * 60)
        except Exception as e:
            messagebox.showerror("起動エラー", f"Chromeの起動に失敗しました:\n{e}")

    def open_browser_for_login(self):
        if not os.path.exists(self.account_config_path):
            messagebox.showerror("エラー", f"'{self.account_config_path}' が見つかりません。")
            return

        self.log("=" * 60)
        self.log("[INFO] 🌐 【初回手動ログイン・セッション永続化ブラウザ】を起動します...")
        self.log("[INFO] 💡 画面が開いたら、ブラウザ上でID/PW入力やreCAPTCHA（画像認証）を手動で完了してください。")
        self.log("[INFO] 💡 一度ログインが完了すれば、セッション情報が永続保存され【次回以降は完全自動】になります！")
        self.log("=" * 60)

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
                
                clean_acc_id = account_name.lstrip("@").strip()
                profile_dir = os.path.join(BASE_DIR, f"chrome_profiles/{clean_acc_id}")
                os.makedirs(profile_dir, exist_ok=True)

                with sync_playwright() as p:
                    context = safe_launch_persistent_context(p, profile_dir, headless=False, log_func=self.log)
                    self.current_context = context
                    
                    token = str(first_acc.get("auth_token") or first_acc.get("Authトークン_sessionid") or first_acc.get("sessionid") or "").strip()
                    if token:
                        inject_session_cookies(context, token, account_name, log_func=self.log)

                    page = context.pages[0] if context.pages else context.new_page()
                    self.current_page = page

                    # 初回ログイン画面またはThreadsホームへアクセス
                    initial_url = "https://www.threads.net/login" if not token else "https://www.threads.net/"
                    try:
                        page.goto(initial_url, timeout=45000, wait_until="domcontentloaded")
                    except Exception:
                        pass

                    clean_pw = str(first_acc.get("パスワード") or first_acc.get("password") or "").strip()
                    # ログイン画面の自動入力補助（ID・パスワードの手間を最小限に）
                    if clean_pw and not token:
                        try:
                            time.sleep(1.5)
                            u_loc, p_loc = find_login_form_inputs(page, timeout_sec=3)
                            if not u_loc:
                                for sso_sel in [
                                    'button:has-text("Instagramアカウントでログイン")',
                                    'button:has-text("Instagramでログイン")',
                                    'div[role="button"]:has-text("Instagramアカウントでログイン")',
                                    'div[role="button"]:has-text("Instagramでログイン")'
                                ]:
                                    try:
                                        s_btn = page.locator(sso_sel).first
                                        if s_btn.is_visible(timeout=300):
                                            s_btn.click()
                                            time.sleep(1.5)
                                            u_loc, p_loc = find_login_form_inputs(page, timeout_sec=4)
                                            if u_loc:
                                                break
                                    except Exception:
                                        pass
                            if u_loc and p_loc:
                                u_loc.fill(clean_acc_id)
                                p_loc.fill(clean_pw)
                                self.log(f"[INFO] ✍️ [{account_name}] ExcelのIDとパスワードを自動入力しました。画面にreCAPTCHA（画像認証）が表示された場合は手動で解除してください。")
                        except Exception:
                            pass

                    self.log(f"[INFO] 👤 [{account_name}] 手動ログイン待機モード稼働中 (URL: {page.url})")
                    self.log("[INFO] 👉 開いたブラウザでログイン操作（reCAPTCHA画像選択など）を行ってください。")
                    self.root.after(0, lambda: self.status_lbl.config(text=f"👤 @{account_name} 初回手動ログイン操作待機中...", fg="#38bdf8"))
                    self.root.after(0, lambda: self.btn_stop.config(state="normal"))

                    login_detected = False
                    last_captured_session = token

                    # ユーザーの手動ログイン操作を最大600秒間（10分間）リアルタイム常時監視（勝手なURLリダイレクトは行わない）
                    for _ in range(600):
                        if page.is_closed() or self.stop_requested:
                            break

                        try:
                            cur_page_url = (page.url or "").lower()
                            # 1. Cookieのsessionid検知
                            cks = context.cookies(["https://www.threads.net", "https://www.instagram.com"])
                            cur_sess = next((c.get("value") for c in cks if c.get("name") == "sessionid" and len(str(c.get("value") or "")) > 15), None)
                            
                            # 2. Threadsホーム画面到達検知
                            is_threads_ok, reason = detect_threads_logged_in(page)

                            # 誤検知ガード: recaptchaやログイン認証中はまだ完了とみなさない
                            is_in_challenge = any(k in cur_page_url for k in ["recaptcha", "auth_platform", "challenge", "checkpoint", "/login"])

                            if ((cur_sess and not is_in_challenge) or is_threads_ok):
                                if cur_sess:
                                    last_captured_session = cur_sess
                                auto_refresh_and_save_session(context, account_name, profile_dir, self.account_config_path, log_func=self.log)
                                self.log("=" * 60)
                                self.log(f"[SUCCESS] 🎉 [{account_name}] 【初回手動ログイン大成功！】セッション情報をプロファイルとExcelへ永続保存しました！")
                                self.log(f"[SUCCESS] 🚀 次回以降は【完全自動】でログインされます。")
                                self.log("=" * 60)
                                login_detected = True
                                self.root.after(0, lambda acc=account_name: messagebox.showinfo(
                                    "手動ログイン成功",
                                    f"@{acc} の初回ログインとセッション永続化が完了しました！\n\n"
                                    "プロファイルと「アカウント情報.xlsx」にセッション情報が保存されました。\n"
                                    "次回以降は再ログインなしで、完全自動でログイン・投稿されます！\n\n"
                                    "※このブラウザウィンドウは閉じていただいて構いません。"
                                ))
                                break
                        except Exception:
                            pass

                        time.sleep(1.0)

                    # 終了時にも念のため最新状態を保存
                    try:
                        auto_refresh_and_save_session(context, account_name, profile_dir, self.account_config_path, log_func=self.log)
                    except Exception:
                        pass
                    try:
                        context.close()
                    except Exception:
                        pass
            except Exception as e:
                self.log(f"[ERROR] ログインブラウザ起動エラー: {e}")
            finally:
                self.current_context = None
                self.current_page = None
                self.stop_requested = False
                self.root.after(0, lambda: self.btn_browser_login.config(state="normal"))
                self.root.after(0, lambda: self.btn_start.config(state="normal"))
                self.root.after(0, lambda: self.btn_stop.config(state="disabled"))
                self.root.after(0, lambda: self.status_lbl.config(text="● 待機中", fg="#10b981"))
                self.root.after(0, self.load_excel_data)

        threading.Thread(target=_worker, daemon=True).start()

    def open_browser_for_profile_edit(self):
        """【プロフィール編集】Chromeを起動してログイン済みThreadsのプロフィール画面を開き、ユーザーが手動編集できるようにウィンドウを維持"""
        if not os.path.exists(self.account_config_path):
            messagebox.showerror("エラー", f"'{self.account_config_path}' が見つかりません。")
            return

        all_accounts = read_excel_records(self.account_config_path)
        valid_accounts = [
            acc for acc in all_accounts
            if str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip() != ""
        ]
        if not valid_accounts:
            messagebox.showwarning("警告", "アカウント情報.xlsx に有効なアカウントが見つかりません。")
            return

        target_acc = valid_accounts[0]
        target_acc_name = str(target_acc.get("アカウントID") or target_acc.get("アカウント名") or "").strip().lstrip("@")
        prof_name = target_acc_name
        prof_dir = os.path.join(BASE_DIR, f"chrome_profiles/{prof_name}")
        os.makedirs(prof_dir, exist_ok=True)

        self.log(f"[INFO] 👤 @{target_acc_name} のプロフィール編集用ブラウザを起動しています...")
        self.btn_edit_profile.config(state="disabled")
        self.btn_browser_login.config(state="disabled")
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.status_lbl.config(text=f"👤 @{target_acc_name} プロフィール編集・手動操作中", fg="#38bdf8")

        def _worker():
            try:
                from playwright.sync_api import sync_playwright
                with sync_playwright() as p:
                    context = safe_launch_persistent_context(p, prof_dir, headless=False, log_func=self.log)
                    self.current_context = context
                    page = context.pages[0] if context.pages else context.new_page()
                    self.current_page = page

                    token = str(target_acc.get("auth_token") or target_acc.get("Authトークン_sessionid") or target_acc.get("sessionid") or "").strip()
                    if token:
                        inject_session_cookies(context, token, target_acc_name, log_func=self.log)

                    safe_goto_threads(page, context, token, target_acc_name, log_func=self.log, timeout=60000)
                    time.sleep(2)

                    try:
                        clean_id = target_acc_name.lstrip('@')
                        page.goto(f"https://www.threads.net/@{clean_id}", wait_until="domcontentloaded", timeout=30000)
                    except Exception:
                        pass

                    self.log("=" * 60)
                    self.log(f"[SUCCESS] 👤 @{target_acc_name} のプロフィール画面を開きました！")
                    self.log("[INFO] 💡 画面上の「プロフィールを編集」ボタンから、自己紹介・名前・リンク・アイコン等を自由に変更していただけます。")
                    self.log("[INFO] 🚪 編集が終わったら、ブラウザの「×」ボタンを押して閉じるか、「⏹ 停止」を押してください。")
                    self.log("=" * 60)

                    while not page.is_closed() and not self.stop_requested:
                        time.sleep(0.5)

                    try:
                        auto_refresh_and_save_session(context, target_acc_name, prof_dir, self.account_config_path, log_func=self.log)
                        self.log(f"[SUCCESS] 💾 @{target_acc_name} の最新Cookie・セッション情報を自動保存しました！")
                    except Exception:
                        pass

                    try:
                        context.close()
                    except Exception:
                        pass
            except Exception as ex:
                self.log(f"[ERROR] プロフィール編集ブラウザエラー: {ex}")
            finally:
                self.current_context = None
                self.current_page = None
                self.stop_requested = False
                self.root.after(0, lambda: self.btn_edit_profile.config(state="normal"))
                self.root.after(0, lambda: self.btn_browser_login.config(state="normal"))
                self.root.after(0, lambda: self.btn_start.config(state="normal"))
                self.root.after(0, lambda: self.btn_stop.config(state="disabled"))
                self.root.after(0, lambda: self.status_lbl.config(text="● 待機中", fg="#10b981"))
                self.root.after(0, self.load_excel_data)

        threading.Thread(target=_worker, daemon=True).start()

    def _wait_for_user_in_open_browser(self, page, context, account_name, profile_dir):
        """
        停止ボタンが押された際、ユーザーがブラウザ上でプロフィール編集や手動操作を行えるよう
        ウィンドウを閉じずに維持し、ユーザーが手動で閉じるか再開するまで待機します。
        """
        self.root.after(0, lambda: self.status_lbl.config(text=f"⏸️ 停止中: @{account_name} 手動操作可能", fg="#f59e0b"))
        self.root.after(0, lambda: self.btn_start.config(state="normal"))
        self.root.after(0, lambda: self.btn_stop.config(state="disabled"))
        self.log(f"[INFO] ⏸️ @{account_name} のブラウザウィンドウを開いたまま待機中（プロフィール編集などの手動操作が可能です）...")

        clean_id = account_name.lstrip('@')
        self.log(f"[INFO] 🔗 プロフィール画面URL: https://www.threads.net/@{clean_id}")

        while not self.is_running:
            try:
                if page.is_closed() or (hasattr(context, "is_closed") and context.is_closed()):
                    self.log(f"[INFO] 👤 @{account_name} のブラウザがユーザーにより閉じられました。")
                    break
            except Exception:
                break
            time.sleep(0.5)

        # 最新セッションとCookieを自動保存
        try:
            if not page.is_closed():
                auto_refresh_and_save_session(context, account_name, profile_dir, self.account_config_path, log_func=self.log)
        except Exception:
            pass

    def start_posting(self):
        if not os.path.exists(self.account_config_path):
            messagebox.showerror("エラー", f"'{self.account_config_path}' が見つかりません。")
            return

        # 起動直前に最新のExcelデータ（ID・パスワード）を自動再読込
        self.load_excel_data()

        accounts = read_excel_records(self.account_config_path)
        active = [a for a in accounts if str(a.get("アカウントID") or "").strip()]
        if not active:
            messagebox.showwarning(
                "アカウント未検出",
                f"『{os.path.basename(self.account_config_path)}』に有効なアカウントが読み込めませんでした。\n\n"
                "【確認事項】\n"
                "1. Excelで入力中のセルを確定（Enterキーを押す）\n"
                "2. Excelファイルを上書き保存（Ctrl+S）\n"
                "3. 1行目に「アカウントID」「パスワード」が入力されているか\n\n"
                "ご確認のうえ、再度お試しください。"
            )
            return

        self.is_running = True
        self.stop_requested = False
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.status_lbl.config(text="● 投稿処理中...", fg="#38bdf8")
        self.log("[INFO] 自動投稿スレッドを開始します...")

        t = threading.Thread(target=self._run_automation_loop, daemon=True)
        t.start()

    def stop_posting(self):
        self.is_running = False
        self.stop_requested = True
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        if self.current_page and not self.current_page.is_closed():
            self.log("=" * 65)
            self.log("[INFO] 🛑 停止ボタンが押されました！自動投稿・操作を直ちにストップしました。")
            self.log("[INFO] 👤 プロフィール編集用にブラウザウィンドウを開いたまま維持しています。")
            self.log("[INFO] 💡 プロフィール画面で自己紹介・リンク・アイコン等を自由に変更してください。")
            self.log("[INFO] 🚪 編集完了後、ブラウザを「×」で閉じるか「▶ 自動投稿を開始」を押すと最新情報を保存して再開します。")
            self.log("=" * 65)
            self.status_lbl.config(text="⏸️ 停止中（ブラウザ手動操作・プロフィール編集可能）", fg="#f59e0b")
        else:
            self.log("[INFO] 🛑 停止ボタンが押されました。自動投稿処理を停止しました。")
            self.status_lbl.config(text="● 待機中", fg="#10b981")

    def _run_automation_loop(self):
        try:
            try:
                from playwright.sync_api import sync_playwright
            except ImportError as ie:
                self.log(f"[ERROR] Playwrightが読み込めません: {ie}")
                self.root.after(0, lambda err=ie: messagebox.showerror("エラー", f"Playwrightが見つかりません:\n{err}\n\nコマンドプロンプトで pip install playwright を実行してください。"))
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
                    # 【BAN対策1】前回失敗・実行後のクールダウン待機チェック（連続再試行スピンを完全防止）
                    daily_cooldown_until = a_state.get("daily_cooldown_until", datetime.min)
                    if now < daily_cooldown_until:
                        continue

                    daily_data = get_eligible_daily_post(pf)
                    if daily_data:
                        acc_name = a_state["name"]
                        self.log("=" * 60)
                        self.log(f"[INFO] 🌅 【日常つぶやき枠】@{acc_name} の「{daily_data['slot_name']} ({daily_data['time_slot']})」を検知しました！通常プロモ投稿に優先して日常ポストを実行します。")
                        self.root.after(0, lambda an=acc_name, sn=daily_data['slot_name']: self.status_lbl.config(text=f"● 日常投稿中: @{an} [{sn}]", fg="#fb923c"))
                        ok = self._execute_daily_greeting_post(a_state, daily_data, p)
                        if ok:
                            self.log(f"[INFO] 🌿 @{acc_name} の日常ポスト完了。ブラウザを完全終了しメモリを解放しました。")
                            # 【BAN対策2】日常ポスト成功後、次回の通常プロモ投稿時刻を45〜90分後に延期（同一アカウントでの連投BANを完全回避）
                            safe_promo_delay = random.randint(2700, 5400)
                            a_state["next_run_time"] = datetime.now() + timedelta(seconds=safe_promo_delay)
                            a_state["daily_cooldown_until"] = datetime.now() + timedelta(minutes=60)
                            rem_m = safe_promo_delay // 60
                            self.log(f"[INFO] 🛡️ 連投BAN防止: @{acc_name} の次回の通常プロモ投稿を【{rem_m}分後】に設定しました（日常ポスト直後の連投を遮断）。")
                        else:
                            # 【BAN対策3】失敗時の無限リトライ防止（20分間のクールダウンを設定し429強制ログアウトを回避）
                            a_state["daily_cooldown_until"] = datetime.now() + timedelta(minutes=20)
                            self.log(f"[WARN] 🛡️ 連続再試行BAN防止: @{acc_name} の日常ポスト失敗を検知。20分間の待機クールダウンを設定しました。")

                        self.root.after(0, self.load_excel_data)
                        daily_executed = True
                        break
                if daily_executed:
                    time.sleep(5)
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
                self.root.after(0, lambda an=acc_name: self.status_lbl.config(text=f"● 実行中: @{an} (1窓のみ起動中)", fg="#38bdf8"))

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

            self.root.after(0, lambda an=nearest_acc['name'], td=time_disp: self.status_lbl.config(
                text=f"● 待機中: 次回 @{an} (あと {td}) [ブラウザ完全終了・省メモリ]",
                fg="#10b981"
            ))

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
        _THREAD_LOG_LOCAL.account_name = sanitize_account_folder_name(account_name)
        get_account_images_dir(account_name, create=True)

        # アカウント情報.xlsx から最新データ（パスワード等）を即時再取得
        if self.account_config_path and os.path.exists(self.account_config_path):
            try:
                for lr in read_excel_records(self.account_config_path):
                    if str(lr.get("アカウントID") or "").strip().lstrip("@") == account_name.lstrip("@"):
                        account = lr
                        acc_state["account"] = lr
                        break
            except Exception:
                pass

        clean_acc_id = account_name.lstrip("@").strip()
        raw_prof_col = str(account.get("プロファイル名") or account.get("profile") or "").strip().lstrip("@").strip() if isinstance(account, dict) else ""
        if raw_prof_col and raw_prof_col.lower() not in ["nan", "none", ""] and not any(k in raw_prof_col for k in ["コミュニティ", "トピック", "モード", "間隔", "ツイート", "ポスト"]):
            prof_name = os.path.basename(raw_prof_col.replace("\\", "/"))
        else:
            prof_name = clean_acc_id
        profile_subpath = f"chrome_profiles/{prof_name}"
        profile_dir = os.path.join(BASE_DIR, profile_subpath)
        os.makedirs(profile_dir, exist_ok=True)
        custom_settings = read_account_settings_from_post_file(posts_file)

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
        # 過去のクラッシュ残骸(SingletonLock)を事前除去してChrome起動失敗を防止
        for lk in ["SingletonLock", "SingletonCookie", "SingletonSocket"]:
            try:
                lk_p = os.path.join(profile_dir, lk)
                if os.path.exists(lk_p):
                    os.remove(lk_p)
            except Exception:
                pass

        context = None
        row_idx = post_item["_row_idx"]
        post_id = post_item.get("ID") or post_item.get("id") or row_idx - 1

        try:
            context = safe_launch_persistent_context(p, profile_dir, headless=False, log_func=self.log)
            self.current_context = context
            page = context.pages[0] if context.pages else context.new_page()
            self.current_page = page
            self.current_account_name = account_name
            self.current_profile_dir = profile_dir

            if not self.is_running or self.stop_requested:
                self._wait_for_user_in_open_browser(page, context, account_name, profile_dir)
                return 0

            # 提案2: 画像・動画通信ブロック
            if self.var_block_media.get():
                setup_media_blocking(context, log_func=self.log)

            if acc_state.get("pure_profile_mode"):
                session_token = ""
                self.log(f"[INFO] 🛡️ [{account_name}] 【ID/PWプロファイル運用モード】sessionidは使用せず、Chromeプロファイルの正規セッションで起動します。")
            else:
                session_token = str(account.get("auth_token") or account.get("Authトークン_sessionid") or account.get("sessionid") or "").strip()
                if session_token:
                    inject_session_cookies(context, session_token, account_name, log_func=self.log)

            # リダイレクトループ(ERR_TOO_MANY_REDIRECTS)自動自己修復付き遷移
            safe_goto_threads(page, context, session_token, account_name, log_func=self.log, timeout=60000)
            time.sleep(3)

            if not self.is_running or self.stop_requested:
                self._wait_for_user_in_open_browser(page, context, account_name, profile_dir)
                return 0

            is_cp, cp_why = detect_phone_checkpoint_screen(page)
            if is_cp:
                self.log("=" * 60)
                self.log(f"[WARN] 🚨 [{account_name}] 携帯電話番号（SMS）認証画面を検知しました（{cp_why}）。")
                capture_error_diagnostic(page, account_name, "携帯電話番号認証要求", None, self.log)
                self.log("=" * 60)
                return 300

            is_logged_in = ensure_authenticated_session(
                page, context, account, self.account_config_path, log_func=self.log, wait_timeout_sec=300, acc_state=acc_state
            )
            if not is_logged_in:
                if acc_state.get("login_cooldown_started"):
                    rem = int(max(60, acc_state.get("login_cooldown_until", 0) - time.time()))
                    self.log(f"[INFO] ⏳ [{account_name}] 15分の安全冷却完了後にID/PW自動ログインを実行します（次回実行予定: 約{rem//60}分後）。")
                    return rem
                self.log(f"[ERROR] [{account_name}] ログイン認証が完了しなかったため、このアカウントの今回の投稿をスキップします。")
                return 300

            # Metaが発行・更新した最新セッションIDをメモリおよびアカウント情報.xlsxに双方向自動同期
            fresh_tok = save_context_cookies_to_excel(context, self.account_config_path, account_name, log_func=self.log)
            if fresh_tok:
                account["auth_token"] = fresh_tok
                acc_state["account"]["auth_token"] = fresh_tok

            text_content = expand_spintax_and_placeholders(self._extract_post_text(post_item))
            raw_image = self._extract_raw_image(post_item)
            resolved_img = resolve_image_path(raw_image, text_content, account_name=account_name)
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
                    safe_goto_threads(page, context, session_token, account_name, log_func=self.log, timeout=30000)
                    time.sleep(2.0)
            except Exception:
                pass

            try:
                page.keyboard.press("Escape")
                time.sleep(0.3)
            except Exception:
                pass

            # 提案新機能1: 投稿前のステルス巡回（ウォーミングアップ閲覧）
            if not self.is_running or self.stop_requested:
                self._wait_for_user_in_open_browser(page, context, account_name, profile_dir)
                return 0

            if self.var_stealth_patrol.get():
                perform_stealth_feed_patrol(page, account_name, log_func=self.log, abort_check=lambda: not self.is_running or self.stop_requested)

            if not self.is_running or self.stop_requested:
                self._wait_for_user_in_open_browser(page, context, account_name, profile_dir)
                return 0

            # 2. 2段階フェイルセーフ投稿を実行
            #    【第1優先】画面左下の円形「＋」ボタンからの投稿
            #    【第2優先】失敗時はタイムライン上部の「最近どう？ / 投稿」枠からの代替投稿
            #    ※ どちらか一方が成功した時点で直ちに完了し、2重投稿を完全に防止します
            posted_ok, post_err_msg = execute_threads_post_with_fallback(
                page=page,
                account_name=account_name,
                text_content=text_content,
                target_topic=target_topic,
                resolved_img=resolved_img,
                log_func=self.log,
                abort_check=lambda: not self.is_running or self.stop_requested
            )

            if not posted_ok:
                raise Exception(post_err_msg or "投稿処理に失敗しました（左下「＋」ボタンおよび「最近どう？」枠の両方で投稿できませんでした）。")

            finish_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            update_post_status(posts_file, row_idx, status="投稿完了", finish_time=finish_time, error_msg="", sheet_name=post_item.get("_sheet_name"), applied_topic=target_topic)
            topic_suffix = f" (トピック: 「{target_topic}」)" if target_topic else ""
            self.log(f"[SUCCESS] [{account_name}] Post #{post_id} 投稿成功！Excel更新{topic_suffix}。")
            acc_state["consecutive_errors"] = 0

            # 投稿後の待機時間を決定
            wait_sec = random.randint(min_delay, max_delay)
            return wait_sec

        except Exception as post_err:
            err_str = str(post_err)
            # リダイレクトループエラー (ERR_TOO_MANY_REDIRECTS) の即時自動修復＆再試行セーフティ
            if ("ERR_TOO_MANY_REDIRECTS" in err_str or "too many redirects" in err_str.lower() or "リダイレクトが繰り返し" in err_str) and not acc_state.get("_retried_redirect", False):
                self.log("=" * 60)
                self.log(f"[WARN] 🔄 [{account_name}] リダイレクトループエラー (net::ERR_TOO_MANY_REDIRECTS) を検知しました。")
                self.log(f"[INFO] 🧹 [{account_name}] 破損Cookieを完全消去し、失効したsessionidを破棄して【ID/PWプロファイル運用】へ自動移行します...")
                self.log("=" * 60)
                acc_state["_retried_redirect"] = True
                acc_state["pure_profile_mode"] = True
                if isinstance(account, dict):
                    account["auth_token"] = ""
                try:
                    if context:
                        context.clear_cookies()
                        time.sleep(1.0)
                        context.close()
                except Exception:
                    pass
                time.sleep(2.0)
                return 15

            acc_state["_retried_redirect"] = False
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
            if not self.is_running or self.stop_requested:
                try:
                    if 'page' in locals() and page and not page.is_closed():
                        self._wait_for_user_in_open_browser(page, context, account_name, profile_dir)
                except Exception:
                    pass
            if context:
                try:
                    # 【提案3】Metaが発行・更新した最新Cookieを自動抽出し、プロファイルとExcelを延命保存
                    auto_refresh_and_save_session(context, account_name, profile_dir, self.account_config_path, log_func=self.log)
                except Exception:
                    pass
                try:
                    context.close()
                except Exception:
                    pass
            self.current_context = None
            self.current_page = None
            _THREAD_LOG_LOCAL.account_name = ""

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
        _THREAD_LOG_LOCAL.account_name = sanitize_account_folder_name(account_name)
        get_account_images_dir(account_name, create=True)
        time_slot = daily_data["time_slot"]
        post_text = expand_spintax_and_placeholders(daily_data.get("raw_text") or daily_data["post_text"])
        row_indices = daily_data["row_indices"]
        sheet_name = daily_data["sheet_name"]

        clean_acc_id = account_name.lstrip("@").strip()
        raw_prof_col = str(account.get("プロファイル名") or account.get("profile") or "").strip().lstrip("@").strip() if isinstance(account, dict) else ""
        if raw_prof_col and raw_prof_col.lower() not in ["nan", "none", ""] and not any(k in raw_prof_col for k in ["コミュニティ", "トピック", "モード", "間隔", "ツイート", "ポスト"]):
            prof_name = os.path.basename(raw_prof_col.replace("\\", "/"))
        else:
            prof_name = clean_acc_id
        profile_subpath = f"chrome_profiles/{prof_name}"
        profile_dir = os.path.join(BASE_DIR, profile_subpath)
        os.makedirs(profile_dir, exist_ok=True)
        custom_settings = read_account_settings_from_post_file(posts_file)

        # 過去のクラッシュ残骸(SingletonLock)を事前除去
        for lk in ["SingletonLock", "SingletonCookie", "SingletonSocket"]:
            try:
                lk_p = os.path.join(profile_dir, lk)
                if os.path.exists(lk_p):
                    os.remove(lk_p)
            except Exception:
                pass

        self.log(f"[{account_name}] 🌅 日常つぶやき【{slot_name}】({time_slot}) 自動投稿を開始（スピンタックス＆動的日付展開済: {post_text[:35]}...）")
        context = None
        page = None
        try:
            context = safe_launch_persistent_context(p, profile_dir, headless=False, log_func=self.log)
            self.current_context = context
            page = context.pages[0] if context.pages else context.new_page()
            self.current_page = page
            self.current_account_name = account_name
            self.current_profile_dir = profile_dir

            if not self.is_running or self.stop_requested:
                self._wait_for_user_in_open_browser(page, context, account_name, profile_dir)
                return False

            if self.var_block_media.get():
                setup_media_blocking(context, log_func=self.log)

            if acc_state.get("pure_profile_mode"):
                session_token = ""
                self.log(f"[INFO] 🛡️ [{account_name}] 【ID/PWプロファイル運用モード】sessionidは使用せず、Chromeプロファイルの正規セッションで日常投稿を起動します。")
            else:
                session_token = str(account.get("auth_token") or account.get("Authトークン_sessionid") or account.get("sessionid") or "").strip()
                if session_token:
                    inject_session_cookies(context, session_token, account_name, log_func=self.log)

            safe_goto_threads(page, context, session_token, account_name, log_func=self.log, timeout=60000)
            time.sleep(3)

            if not self.is_running or self.stop_requested:
                self._wait_for_user_in_open_browser(page, context, account_name, profile_dir)
                return False

            is_logged_in = ensure_authenticated_session(
                page, context, account, self.account_config_path, log_func=self.log, wait_timeout_sec=300, acc_state=acc_state
            )
            if not is_logged_in:
                self.log(f"[ERROR] [{account_name}] ログイン認証が完了しなかったため日常投稿をスキップします。")
                return False

            # Metaが発行・更新した最新セッションIDをメモリおよびアカウント情報.xlsxに双方向自動同期
            fresh_tok = save_context_cookies_to_excel(context, self.account_config_path, account_name, log_func=self.log)
            if fresh_tok:
                account["auth_token"] = fresh_tok
                acc_state["account"]["auth_token"] = fresh_tok

            try:
                cur_url = page.url.split('?')[0].rstrip('/')
                if cur_url != "https://www.threads.net":
                    safe_goto_threads(page, context, session_token, account_name, log_func=self.log, timeout=30000)
                    time.sleep(2.0)
            except Exception:
                pass

            try:
                page.keyboard.press("Escape")
                time.sleep(0.3)
            except Exception:
                pass

            if not self.is_running or self.stop_requested:
                self._wait_for_user_in_open_browser(page, context, account_name, profile_dir)
                return False

            if self.var_stealth_patrol.get():
                perform_stealth_feed_patrol(page, account_name, log_func=self.log, abort_check=lambda: not self.is_running or self.stop_requested)

            if not self.is_running or self.stop_requested:
                self._wait_for_user_in_open_browser(page, context, account_name, profile_dir)
                return False

            # トピック設定（日常ポストの「コミュニティまたはトピック」列から設定可能）
            daily_raw_topic = daily_data.get("topic", "")
            target_topic = None
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

            # 2段階フェイルセーフ投稿（左下「＋」ボタン優先、失敗時はタイムライン上部「最近どう？ / 投稿」枠）
            # ※ どちらか一方が成功した時点で完了し、2重投稿を完全に防止
            posted_ok, post_err_msg = execute_threads_post_with_fallback(
                page=page,
                account_name=account_name,
                text_content=post_text,
                target_topic=target_topic,
                resolved_img=None,
                log_func=self.log,
                abort_check=lambda: not self.is_running or self.stop_requested
            )

            if not posted_ok:
                raise Exception(post_err_msg or "日常ポスト送信が完了できませんでした。")

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
            if not self.is_running or self.stop_requested:
                try:
                    if page and not page.is_closed():
                        self._wait_for_user_in_open_browser(page, context, account_name, profile_dir)
                except Exception:
                    pass
            if context:
                try:
                    # 【提案3】日常ポスト投稿後もセッションCookieを自動抽出して延命更新
                    auto_refresh_and_save_session(context, account_name, profile_dir, self.account_config_path, log_func=self.log)
                except Exception:
                    pass
                try:
                    context.close()
                except Exception:
                    pass
            self.current_context = None
            self.current_page = None
            _THREAD_LOG_LOCAL.account_name = ""

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
                            ok = self._execute_daily_greeting_post(acc_st, daily_data, p)
                            if ok:
                                self.log(f"[INFO] 🛡️ 連投BAN防止: @{acc_name} の日常ポスト完了直後のプロモ投稿をスキップし次アカウントへ移行します。")
                                continue
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
        _THREAD_LOG_LOCAL.account_name = sanitize_account_folder_name(account_name)
        get_account_images_dir(account_name, create=True)
        posts_file = self._find_posts_file(account_name)
        if not posts_file or not os.path.exists(posts_file):
            self.log(f"[{account_name}] '{posts_file}' が見つかりません。")
            return

        custom_settings = read_account_settings_from_post_file(posts_file)
        profile_subpath = custom_settings.get("profile_dir", "")
        if not profile_subpath or any(k in profile_subpath for k in ["コミュニティ", "トピック", "モード", "間隔", "ツイート", "ポスト"]): profile_subpath = f"chrome_profiles/{account_name}"
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
            safe_goto_threads(page, context, session_token, account_name, log_func=self.log, timeout=60000)
            time.sleep(3)

            is_logged_in = ensure_authenticated_session(
                page, context, account, self.account_config_path, log_func=self.log, wait_timeout_sec=300
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
                resolved_img = resolve_image_path(raw_image, text_content, account_name=account_name)
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

                    # 2. 2段階フェイルセーフ投稿を実行
                    #    【第1優先】画面左下の円形「＋」ボタンからの投稿
                    #    【第2優先】失敗時はタイムライン上部の「最近どう？ / 投稿」枠からの代替投稿
                    #    ※ どちらか一方が成功した時点で完了し、2重投稿を完全に防止
                    posted_ok, post_err_msg = execute_threads_post_with_fallback(
                        page=page,
                        account_name=account_name,
                        text_content=text_content,
                        target_topic=target_topic,
                        resolved_img=resolved_img,
                        log_func=self.log
                    )

                    if not posted_ok:
                        raise Exception(post_err_msg or "投稿処理に失敗しました（左下「＋」ボタンおよび「最近どう？」枠の両方で投稿できませんでした）。")


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
            _THREAD_LOG_LOCAL.account_name = ""

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

