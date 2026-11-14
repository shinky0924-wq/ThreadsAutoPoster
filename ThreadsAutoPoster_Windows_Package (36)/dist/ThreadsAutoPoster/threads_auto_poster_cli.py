# -*- coding: utf-8 -*-
"""
Meta Threads 自動投稿システム (Playwright CLI版)
================================================
openpyxlによる軽量・安全なExcel直接読み込み＆更新（NumPy/Pandas不要）
- アカウント情報.xlsx: アカウントID, auth_token
- 投稿エクセル: ポスト文章, 画像ファイルパス (設定セルのimage.png, ポスト文章内のimage2.pngも自動検出)
"""
import os, sys, time, random, re
from datetime import datetime
import openpyxl
from playwright.sync_api import sync_playwright

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

LOG_DIR = os.path.join(BASE_DIR, "log")
IMAGES_DIR = os.path.join(BASE_DIR, "images")
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(IMAGES_DIR, exist_ok=True)

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

def make_account_logger(account_name):
    log_path = get_account_log_path(account_name)
    def _log(msg):
        print(msg)
        try:
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"[{now_str}] {msg}\n")
                f.flush()
        except Exception:
            pass
    return _log

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

ACCOUNT_CONFIG_FILE = find_account_file()

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
        for sname in wb.sheetnames:
            ws = wb[sname]
            all_rows = list(ws.iter_rows(values_only=True))
            for r_idx, row in enumerate(all_rows):
                if not row:
                    continue
                row_str = " ".join([str(c).strip() for c in row if c is not None])
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
                            min_v, max_v = int(nums[0]), int(nums[1])
                            settings["min_interval_sec"] = int(min_v * mult)
                            settings["max_interval_sec"] = int(max_v * mult)
                            settings["interval_str"] = f"{min_v}〜{max_v}{'分' if mult == 60 else '秒'}"
                            break
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
                            elif re.match(r'^d{1,2}:d{2}', c_s):
                                times.append(c_s)
                            elif c_s.isdigit() and len(c_s) <= 2:
                                times.append(f"{c_s}:00")
                        if len(times) >= 2:
                            settings["time_slot"] = f"{times[0]}-{times[1]}"
                            break
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

                # 4. Chromeプロファイル保存パス (または プロファイル)
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
        wb.close()
    except Exception as e:
        print(f"[WARN] 設定読み込みエラー: {e}")
    return settings

def update_post_status(filepath, row_idx, status="投稿完了", finish_time=None, error_msg="", sheet_name=None):
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

        status_col, time_col, err_col = None, None, None
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
            if status_col:
                break

        if status_col:
            ws.cell(row=row_idx, column=status_col, value=status)
        if time_col and finish_time:
            ws.cell(row=row_idx, column=time_col, value=finish_time)
        if err_col:
            ws.cell(row=row_idx, column=err_col, value=error_msg)
        wb.save(filepath)
        wb.close()
        return True
    except Exception as e:
        print(f"[WARN] Excel書き込みエラー: {e}")
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
            cand_base = os.path.join(d, os.path.basename(cleaned))
            if os.path.exists(cand_base):
                return os.path.abspath(cand_base)

    if post_text:
        matches = re.findall(r'([a-zA-Z0-9_-./\:]+.(?:png|jpg|jpeg|webp|gif))', str(post_text), re.IGNORECASE)
        for m in matches:
            m_clean = m.strip()
            for d in search_dirs:
                cand = os.path.join(d, m_clean)
                if os.path.exists(cand):
                    return os.path.abspath(cand)
                cand_base = os.path.join(d, os.path.basename(m_clean))
                if os.path.exists(cand_base):
                    return os.path.abspath(cand_base)
    return ""

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

def inject_session_cookies(context, raw_token, account_name, log_func=print):
    token = str(raw_token or "").strip()
    if not token or token.lower() == "nan" or token.startswith("ここに"):
        log_func(f"[WARN] [{account_name}] Excelの auth_token (sessionid) が空です。")
        return False
    if token.startswith("sessionid="):
        token = token[len("sessionid="):].strip()
    token = token.strip("\"\x27 ;\t\r\n")
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

    cookies_by_domain = []
    for domain in [".threads.net", "www.threads.net", ".instagram.com", "www.instagram.com"]:
        cookies_by_domain.append({
            "name": "sessionid",
            "value": token,
            "domain": domain,
            "path": "/",
            "httpOnly": True,
            "secure": True,
            "sameSite": "Lax"
        })
        if ds_user_id:
            cookies_by_domain.append({
                "name": "ds_user_id",
                "value": str(ds_user_id),
                "domain": domain,
                "path": "/",
                "secure": True,
                "sameSite": "Lax"
            })
    cookies_by_url = []
    for u in ["https://www.threads.net/", "https://www.instagram.com/"]:
        cookies_by_url.append({
            "name": "sessionid",
            "value": token,
            "url": u,
            "httpOnly": True,
            "secure": True,
            "sameSite": "Lax"
        })
        if ds_user_id:
            cookies_by_url.append({
                "name": "ds_user_id",
                "value": str(ds_user_id),
                "url": u,
                "secure": True,
                "sameSite": "Lax"
            })
    try:
        context.add_cookies(cookies_by_domain)
    except Exception:
        pass
    try:
        context.add_cookies(cookies_by_url)
    except Exception:
        pass
    log_func(f"[SUCCESS] [{account_name}] エクセルの sessionid をブラウザに自動注入しました。")
    return True

def detect_threads_logged_in(page):
    cur_u = (page.url or "").lower()
    if "/login" in cur_u or "about:blank" in cur_u:
        return False, "ログイン画面表示中"

    try:
        content = page.content()
        if any(w in content for w in ["Instagramアカウントでログイン", "Log in with Instagram", "パスワードを忘れた場合"]):
            return False, "ログイン画面表示中"

        if '"viewer":{' in content and ('"username":"' in content or '"id":"' in content):
            if '"viewer":null' not in content:
                import re
                m_user = re.search(r'"viewer":\{[^}]*"username":"([^"]+)"', content)
                user_found = m_user.group(1) if m_user else ""
                return True, f"セッション認証済み (ユーザー: @{user_found or '確認済'})"
    except Exception:
        pass

    success_selectors = [
        'svg[aria-label="新しい投稿"]',
        'svg[aria-label="New post"]',
        'svg[aria-label="投稿を作成"]',
        'svg[aria-label="Create"]',
        'svg[aria-label="作成"]',
        'div[role="button"]:has-text("投稿")',
        'div[role="button"]:has-text("Post")',
        'a[href*="/@"]',
        'svg[aria-label="プロフィール"]',
        'svg[aria-label="ホーム"]',
        'div[role="textbox"]'
    ]
    for sel in success_selectors:
        try:
            if page.locator(sel).first.is_visible(timeout=500):
                return True, f"ログイン済み要素 ({sel}) を検出"
        except Exception:
            pass

    unauth_selectors = [
        'span:has-text("Instagramでログイン")',
        'div[role="button"]:has-text("Instagramでログイン")',
        'button:has-text("Instagramでログイン")',
        'h1:has-text("Threadsでもっと発信しよう")',
        'h2:has-text("Threadsでもっと発信しよう")',
        'div:has-text("Threadsでもっと発信しよう")',
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
                log_func(f"[INFO] [{account_name}] セッション状態を「{os.path.basename(filepath)}」と自動同期しました。")
            return new_sessionid
    except Exception as e:
        log_func(f"[WARN] sessionid 自動保存エラー: {e}")
    return ""

def find_login_form_inputs(page, timeout_sec=8):
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
        for p_sel in p_selectors:
            try:
                loc = page.locator(p_sel).first
                if loc.is_visible(timeout=200):
                    pass_loc = loc
                    break
            except Exception:
                pass

        for u_sel in u_selectors:
            try:
                loc = page.locator(u_sel).first
                if loc.is_visible(timeout=200):
                    user_loc = loc
                    break
            except Exception:
                pass

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

        if "about:blank" in cur_u or ("threads.net" not in cur_u and "instagram.com" not in cur_u):
            log_func(f"[INFO] 🌐 [{account_name}] Threads公式ログイン画面 (https://www.threads.net/login) へアクセスします...")
            try:
                page.goto("https://www.threads.net/login", timeout=35000, wait_until="domcontentloaded")
                time.sleep(2.5)
            except Exception as e_goto:
                log_func(f"[WARN] [{account_name}] ログイン画面アクセス警告: {e_goto}")

        user_loc, pass_loc = find_login_form_inputs(page, timeout_sec=10)

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

        log_func(f"[INFO] ⏳ [{account_name}] Instagramの認証処理・通信ラグを待機しています（最大60秒）...")
        instagram_logged_in = False
        recaptcha_announced = False
        
        for wait_i in range(1, 61):
            time.sleep(1.0)
            cur_u = (page.url or "").lower()

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
                continue

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

            try:
                all_cookies = context.cookies()
                has_sess = any(c.get("name") == "sessionid" and len(str(c.get("value") or "")) > 15 for c in all_cookies)
                if has_sess and not any(k in cur_u for k in ["/login", "recaptcha", "auth_platform", "challenge", "checkpoint"]):
                    log_func(f"[SUCCESS] 🔑 [{account_name}] Instagram正規セッションCookie (sessionid) の生成を確認しました！")
                    instagram_logged_in = True
                    break
            except Exception:
                pass

            if "threads.net" in cur_u and "/login" not in cur_u:
                is_ok, _ = detect_threads_logged_in(page)
                if is_ok:
                    instagram_logged_in = True
                    break

            if "instagram.com" in cur_u and not any(k in cur_u for k in ["/accounts/login", "recaptcha", "auth_platform", "challenge", "checkpoint", "about:blank"]):
                log_func(f"[SUCCESS] 🚪 [{account_name}] Instagram正規画面へのログイン完了を確認しました (URL: {page.url})。")
                instagram_logged_in = True
                break

            if wait_i % 5 == 0 and not is_recaptcha_or_challenge:
                log_func(f"[INFO] ⏳ [{account_name}] 認証待機中... ({wait_i}/60秒)")

        time.sleep(2.0)

        cur_u = (page.url or "").lower()
        if "threads.net" not in cur_u or "/login" in cur_u:
            log_func(f"[INFO] 🌐 [{account_name}] Instagramログイン完了を確認！Threadsホーム画面へ遷移してセッションを確立します...")
            try:
                page.goto("https://www.threads.net/", timeout=45000, wait_until="domcontentloaded")
                time.sleep(3.5)
            except Exception as e_th:
                log_func(f"[WARN] [{account_name}] Threadsホーム遷移警告: {e_th}")

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

def main():
    if not os.path.exists(ACCOUNT_CONFIG_FILE):
        print(f"[ERROR] '{ACCOUNT_CONFIG_FILE}' が見つかりません。")
        return
    all_accounts = read_excel_records(ACCOUNT_CONFIG_FILE)
    accounts = [
        acc for acc in all_accounts
        if str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip() != ""
    ]

    print(f"[*] 有効なアカウント設定数: {len(accounts)}")

    with sync_playwright() as p:
        for acc in accounts:
            name = str(acc.get("アカウントID") or acc.get("アカウント名") or "").strip()
            posts_file = os.path.join(BASE_DIR, f"{name}.xlsx")
            if not os.path.exists(posts_file):
                print(f"[{name}] '{posts_file}' が見つかりません。")
                continue

            # <アカウント名>.xlsx からポスト間隔・モード・実行時間・Chromeプロファイル設定をすべて反映
            custom_settings = read_account_settings_from_post_file(posts_file)

            time_slot = custom_settings.get("time_slot", "0:00-0:00")
            now_hour = datetime.now().hour
            if not is_within_target_hours(time_slot, now_hour):
                print(f"[SKIP] [{name}] 現在時刻({now_hour}時)は{os.path.basename(posts_file)}の指定実行時間 ({time_slot}) 外です。")
                continue

            if "min_interval_sec" in custom_settings:
                min_d = custom_settings["min_interval_sec"]
                max_d = custom_settings["max_interval_sec"]
                print(f"[{name}] 【ポスト間隔適用】{custom_settings.get('interval_str', '')} ({min_d}〜{max_d}秒) のランダム間隔で待機します。")
            else:
                min_d = 7200
                max_d = 21600
                print(f"[{name}] 【デフォルト間隔適用】120〜360分 ({min_d}〜{max_d}秒) のランダム間隔で待機します。")

            post_mode = custom_settings.get("post_mode", 1)
            prof_sub = custom_settings.get("profile_dir") or f"chrome_profiles/{name}"
            prof_dir = os.path.join(BASE_DIR, prof_sub)
            os.makedirs(prof_dir, exist_ok=True)

            all_posts = read_excel_records(posts_file)
            unposted = [p for p in all_posts if str(p.get("ステータス", "")).strip() == "未投稿"]
            if not unposted:
                print(f"[{name}] 未投稿のポストはありません。")
                continue

            if post_mode == 1:
                random.shuffle(unposted)
                print(f"[{name}] 【モード1: ランダム】未投稿 {len(unposted)} 件をランダム順序で投稿します。")
            else:
                print(f"[{name}] 【モード0: 上から順】未投稿 {len(unposted)} 件を先頭から順に投稿します。")

            acc_log = make_account_logger(name)
            get_account_images_dir(name, create=True)
            prof_dir = os.path.join(BASE_DIR, f"chrome_profiles/{sanitize_account_folder_name(name)}")
            os.makedirs(prof_dir, exist_ok=True)
            context = p.chromium.launch_persistent_context(user_data_dir=prof_dir, headless=False, channel="chrome", args=["--disable-blink-features=AutomationControlled"])
            
            # auth_token (sessionid) 注入
            session_token = str(acc.get("auth_token") or acc.get("Authトークン_sessionid") or acc.get("sessionid") or "").strip()
            inject_session_cookies(context, session_token, name, log_func=acc_log)

            page = context.pages[0] if context.pages else context.new_page()
            page.goto("https://www.threads.net/", timeout=60000)
            time.sleep(3)

            is_ok = ensure_authenticated_session(page, context, acc, ACCOUNT_CONFIG_FILE, log_func=acc_log, wait_timeout_sec=300)
            if not is_ok:
                acc_log(f"[ERROR] [{name}] ログイン認証が未完了のためスキップします。")
                continue

            for post_item in unposted:
                row_idx = post_item["_row_idx"]
                post_id = post_item.get("ID") or post_item.get("id") or row_idx - 1
                
                # ポスト文章
                text = ""
                for k in ["ポスト文章", "投稿本文", "ツイート文章", "本文", "テキスト", "text"]:
                    if k in post_item and post_item[k] is not None and str(post_item[k]).strip() != "":
                        text = str(post_item[k]).strip()
                        break
                
                # 画像
                raw_img = ""
                for k in ["画像ファイルパス", "設定セル", "画像", "画像パス", "image"]:
                    if k in post_item and post_item[k] is not None and str(post_item[k]).strip() != "":
                        raw_img = str(post_item[k]).strip()
                        break
                resolved_img = resolve_image_path(raw_img, text, account_name=name)

                acc_log(f"[{name}] Post #{post_id} ({len(text)}文字) 投稿中...")

                for close_s in ['svg[aria-label="閉じる"]', 'div[role="button"][aria-label="閉じる"]']:
                    try:
                        if page.locator(close_s).first.is_visible(timeout=500):
                            page.locator(close_s).first.click()
                            time.sleep(1)
                    except Exception:
                        pass

                for s in ['svg[aria-label="新しい投稿"]', 'svg[aria-label="投稿を作成"]', 'div[role="button"]:has-text("投稿")', 'div[role="button"]:has-text("新しいスレッド")']:
                    try:
                        if page.locator(s).first.is_visible(timeout=1000):
                            page.locator(s).first.click()
                            break
                    except Exception:
                        pass
                time.sleep(2)

                tb = page.locator('div[role="textbox"], div[contenteditable="true"]').first
                tb.click()
                for line in text.split("\n"):
                    tb.press_sequentially(line, delay=random.randint(15, 30))
                    tb.press("Enter")
                    time.sleep(0.1)
                time.sleep(1)

                if resolved_img:
                    acc_log(f"[{name}] 画像添付: {os.path.basename(resolved_img)}")
                    file_input = page.locator('input[type="file"]').first
                    if file_input:
                        file_input.set_input_files(resolved_img)
                        time.sleep(2.5)

                time.sleep(1.5)
                page.locator('div[role="button"]:has-text("投稿する"), div[role="button"]:has-text("投稿")').last.click()
                time.sleep(4)

                finish_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                update_post_status(posts_file, row_idx, status="投稿完了", finish_time=finish_time, sheet_name=post_item.get("_sheet_name"))
                acc_log(f"[SUCCESS] [{name}] Post #{post_id} 投稿完了！Excel更新")

                wait_sec = random.randint(min_d, max_d)
                if wait_sec >= 60:
                    acc_log(f"[*] BAN防止のため {wait_sec} 秒 (約 {wait_sec // 60} 分) 待機中...")
                else:
                    acc_log(f"[*] BAN防止のため {wait_sec} 秒待機中...")
                time.sleep(wait_sec)

            context.close()

if __name__ == "__main__":
    main()
