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
        headers = [str(h).strip() if h is not None else f"col_{i}" for i, h in enumerate(rows[target_header_row_idx - 1])]
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
    token = token.strip(""' ;\t\r\n")
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
    success_selectors = [
        'svg[aria-label="新しい投稿"]',
        'svg[aria-label="投稿を作成"]',
        'div[role="button"]:has-text("投稿")',
        'a[href*="/@"]',
        'svg[aria-label="プロフィール"]',
        'svg[aria-label="ホーム"]',
        'div[role="textbox"]'
    ]
    for sel in success_selectors:
        try:
            if page.locator(sel).first.is_visible(timeout=800):
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
            if page.locator(sel).first.is_visible(timeout=800):
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

    log_func(f"[ERROR] [{account_name}] エクセルの sessionid で自動ログインできませんでした。Excelの「auth_token」列の値が正しいかご確認ください。")
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

            prof_dir = os.path.join(BASE_DIR, str(acc.get("Chromeプロファイル保存パス", f"chrome_profiles/{name}")))
            os.makedirs(prof_dir, exist_ok=True)
            context = p.chromium.launch_persistent_context(user_data_dir=prof_dir, headless=False, channel="chrome", args=["--disable-blink-features=AutomationControlled"])
            
            # auth_token (sessionid) 注入
            session_token = str(acc.get("auth_token") or acc.get("Authトークン_sessionid") or acc.get("sessionid") or "").strip()
            inject_session_cookies(context, session_token, name, log_func=print)

            page = context.pages[0] if context.pages else context.new_page()
            page.goto("https://www.threads.net/", timeout=60000)
            time.sleep(3)

            is_ok = ensure_authenticated_session(page, context, name, ACCOUNT_CONFIG_FILE, log_func=print, wait_timeout_sec=300)
            if not is_ok:
                print(f"[ERROR] [{name}] ログイン認証が未完了のためスキップします。")
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
                resolved_img = resolve_image_path(raw_img, text)

                print(f"[{name}] Post #{post_id} ({len(text)}文字) 投稿中...")

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
                    print(f"[{name}] 画像添付: {os.path.basename(resolved_img)}")
                    file_input = page.locator('input[type="file"]').first
                    if file_input:
                        file_input.set_input_files(resolved_img)
                        time.sleep(2.5)

                time.sleep(1.5)
                page.locator('div[role="button"]:has-text("投稿する"), div[role="button"]:has-text("投稿")').last.click()
                time.sleep(4)

                finish_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                update_post_status(posts_file, row_idx, status="投稿完了", finish_time=finish_time, sheet_name=post_item.get("_sheet_name"))
                print(f"[SUCCESS] [{name}] Post #{post_id} 投稿完了！Excel更新")

                wait_sec = random.randint(min_d, max_d)
                if wait_sec >= 60:
                    print(f"[*] BAN防止のため {wait_sec} 秒 (約 {wait_sec // 60} 分) 待機中...")
                else:
                    print(f"[*] BAN防止のため {wait_sec} 秒待機中...")
                time.sleep(wait_sec)

            context.close()

if __name__ == "__main__":
    main()
