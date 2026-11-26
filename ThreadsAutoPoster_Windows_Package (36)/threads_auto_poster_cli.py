# -*- coding: utf-8 -*-
"""
Meta Threads 自動投稿システム (Playwright CLI版)
================================================
openpyxlによる軽量・安全なExcel直接読み込み＆更新（NumPy/Pandas不要）
- アカウント情報.xlsx: アカウントID, auth_token
- 投稿エクセル: ポスト文章, 画像ファイルパス (設定セルのimage.png, ポスト文章内のimage2.pngも自動検出)
"""
import os, sys, time, random, re, unicodedata, json, hashlib, platform, urllib.parse, threading
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

AI_ENGINE_CHOICES = (
    "groq",
    "gemini",
    "duckai",
    "venice",
    "deepai",
    "talkai",
    "chatgpt",
    "qwen",
    "mistral",
    "grok",
    "deepseek",
    "claude",
    "perplexity",
    "copilot",
    "openrouter",
    "ollama",
    "lmstudio",
    "chrome_builtin",
)

AI_WINDOW_PROFILE_NAME = "_ai_window"
AI_WINDOW_CDP_PORT = 9333

AI_ENGINE_DEFAULTS = {
    "groq": {"model": "llama-3.3-70b-versatile", "api_url": "https://api.groq.com/openai/v1/chat/completions", "web_url": "https://console.groq.com/", "mode": "openai"},
    "gemini": {"model": "gemini-2.5-flash", "api_url": "", "web_url": "https://gemini.google.com/app", "mode": "gemini"},
    "duckai": {"model": "gpt-4o-mini", "api_url": "", "web_url": "https://duckduckgo.com/?q=DuckDuckGo+AI+Chat&ia=chat", "mode": "web"},
    "venice": {"model": "llama-3.3-70b", "api_url": "https://api.venice.ai/api/v1/chat/completions", "web_url": "https://venice.ai/chat", "mode": "web_or_openai"},
    "deepai": {"model": "standard", "api_url": "", "web_url": "https://deepai.org/chat", "mode": "web"},
    "talkai": {"model": "gpt-4o-mini", "api_url": "", "web_url": "https://talkai.info/", "mode": "web"},
    "chatgpt": {"model": "gpt-4o-mini", "api_url": "https://api.openai.com/v1/chat/completions", "web_url": "https://chatgpt.com/", "mode": "web_or_openai"},
    "qwen": {"model": "qwen-plus", "api_url": "https://dashscope-intl.aliyuncs.com/compatible-mode/v1/chat/completions", "web_url": "https://chat.qwen.ai/", "mode": "web_or_openai"},
    "mistral": {"model": "mistral-small-latest", "api_url": "https://api.mistral.ai/v1/chat/completions", "web_url": "https://chat.mistral.ai/chat", "mode": "web_or_openai"},
    "grok": {"model": "grok-3-mini", "api_url": "https://api.x.ai/v1/chat/completions", "web_url": "https://grok.com/", "mode": "web_or_openai"},
    "deepseek": {"model": "deepseek-chat", "api_url": "https://api.deepseek.com/chat/completions", "web_url": "https://chat.deepseek.com/", "mode": "web_or_openai"},
    "claude": {"model": "claude-3-5-haiku-latest", "api_url": "https://api.anthropic.com/v1/messages", "web_url": "https://claude.ai/new", "mode": "claude"},
    "perplexity": {"model": "sonar", "api_url": "https://api.perplexity.ai/chat/completions", "web_url": "https://www.perplexity.ai/", "mode": "web_or_openai"},
    "copilot": {"model": "copilot", "api_url": "", "web_url": "https://copilot.microsoft.com/", "mode": "web"},
    "openrouter": {"model": "google/gemini-2.0-flash-exp:free", "api_url": "https://openrouter.ai/api/v1/chat/completions", "web_url": "https://openrouter.ai/chat", "mode": "openai"},
    "ollama": {"model": "llama3.2", "api_url": "http://127.0.0.1:11434/api/chat", "web_url": "http://127.0.0.1:11434", "mode": "ollama"},
    "lmstudio": {"model": "local-model", "api_url": "http://127.0.0.1:1234/v1/chat/completions", "web_url": "http://127.0.0.1:1234", "mode": "openai_nokey"},
    "chrome_builtin": {"model": "gemini-nano", "api_url": "", "web_url": "about:blank", "mode": "chrome_builtin"},
}

FIXED_AI_SYSTEM_PROMPT = (
    "以下の内容を確認したうえで、設定されたキャラ・指示に沿った魅力的なスレッズのポストのメッセージを作成してください。\n"
    "【作成方針（自動選定）】\n"
    "設定されたキャラ・指示に沿った切り口を内部で3パターン（共感・補足・視点の共有など）検討し、対象の投稿に最も自然で噛み合っている1つだけを選んでください。\n"
    "※検討過程や他案・比較・番号（1. 2. 3. 等）は一切書かず、選んだ1つのポストの本文のみを出力してください。"
)

AI_PROMPT_SECTION_META = [
    ("character", "キャラ設定", "あなたはThreads(スレッズ)で役立つ発信をする親しみやすく魅力的なユーザーです。"),
    ("target", "ターゲット層", "日々の暮らしや仕事を少し良くしたいと思っている20〜40代のユーザー。小さな悩みや気づきに共感しやすい層。"),
    ("purpose", "投稿の目的", "読者の共感や「なるほど」を引き出し、いいね・保存・プロフィール閲覧やリプライにつなげること。"),
    ("tone_rule", "口調・文体ルール", "一人称は使わず自然なです・ます調で、親しみやすく柔らかい文体にしてください。絵文字は文末に0〜1個まで、感嘆符（！）の連続は控えてください。"),
    ("prompt", "プロンプト", "ターゲット層の悩みや関心に寄り添った、思わず共感や保存をしたくなる自然な日本語のThreadsポストを作成してください。ハッシュタグや絵文字の多用は控えてください。"),
    ("samples", "お手本サンプル", "例1: 朝のたった5分で今日やることを3つだけ書き出すようにしたら、1日の焦りが驚くほど減りました。全部完璧にやろうとしない方が案外うまくいきますよね✨\n例2: 「あとでやろう」はだいたい忘れるので、2分で終わることはその場で片付けるようにしています。小さなストレスが溜まらなくておすすめです。"),
    ("ng_rules", "NGルール", "・「素晴らしいですね！」「応援しています！」などの定型的なお世辞やAIっぽい表現は禁止\n・宣伝・営業・勧誘・過度なハッシュタグ（#）は禁止"),
    ("output_rule", "出力ルール", "※出力は投稿するメッセージ本文のみ（140〜400文字以内）で返してください。挨拶や前置き（例：「了解しました」「投稿案:」）や引用符などの余計な注釈は一切含めないでください。"),
]



def _is_ai_post_sheet_name(sname):
    s = unicodedata.normalize("NFKC", str(sname or "")).strip().lower().replace(" ", "").replace("　", "").replace("_", "")
    return s in {"aiポスト", "ai投稿", "aipost", "aiposts", "ai"} or s.startswith("aiポスト") or s.startswith("aipost")


def _parse_ai_interval_range(raw):
    s = unicodedata.normalize("NFKC", str(raw or "")).strip()
    if not s:
        return None, None, ""
    is_sec = "秒" in s
    nums = [float(x) for x in re.findall(r'\d+(?:\.\d+)?', s)]
    if len(nums) >= 2:
        mult = 1 if is_sec else 60
        lo = max(1, int(min(nums[0], nums[1]) * mult))
        hi = max(lo, int(max(nums[0], nums[1]) * mult))
        return lo, hi, s
    if len(nums) == 1:
        mult = 1 if is_sec else 60
        v = max(1, int(nums[0] * mult))
        return v, v, s
    return None, None, ""


def get_ai_window_profile_dir(base_dir=BASE_DIR):
    p_dir = os.path.join(base_dir, "chrome_profiles", AI_WINDOW_PROFILE_NAME)
    os.makedirs(p_dir, exist_ok=True)
    return p_dir


def _find_chrome_executable_for_ai():
    import shutil
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    exe = next((c for c in candidates if c and os.path.exists(c)), None)
    if exe:
        return exe
    return shutil.which("chrome") or shutil.which("google-chrome") or shutil.which("chromium") or shutil.which("msedge")


_AI_WINDOW_LAUNCH_LOCK = threading.Lock()


def _cleanup_profile_locks(profile_dir):
    for lk in ("SingletonLock", "SingletonCookie", "SingletonSocket", "lockfile"):
        try:
            lp = os.path.join(profile_dir, lk)
            if os.path.exists(lp):
                os.remove(lp)
        except Exception:
            pass


def _find_chrome_pids_for_profile(profile_dir, pipe_only=False):
    if not profile_dir:
        return []
    abs_norm = os.path.abspath(profile_dir).replace("/", "\\").lower()
    abs_posix = os.path.abspath(profile_dir).lower()
    pids = []
    if sys.platform == "win32":
        try:
            import subprocess
            ps_script = (
                "Get-CimInstance Win32_Process -Filter \"Name='chrome.exe' OR Name='msedge.exe'\" "
                "| Select-Object ProcessId,CommandLine | ConvertTo-Json -Compress"
            )
            creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            res = subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps_script],
                capture_output=True, text=True, timeout=5, creationflags=creation_flags
            )
            raw = (res.stdout or "").strip()
            if raw:
                items = json.loads(raw)
                if isinstance(items, dict):
                    items = [items]
                for it in (items or []):
                    cmd = str(it.get("CommandLine") or "").replace("/", "\\").lower()
                    pid = it.get("ProcessId")
                    if not pid or not cmd:
                        continue
                    if abs_norm in cmd and "--user-data-dir=" in cmd:
                        if pipe_only and "--remote-debugging-pipe" not in cmd:
                            continue
                        pids.append(int(pid))
        except Exception:
            pass
    else:
        try:
            for ent in os.listdir("/proc"):
                if not ent.isdigit():
                    continue
                try:
                    with open(f"/proc/{ent}/cmdline", "rb") as f:
                        cmd = f.read().replace(b"\x00", b" ").decode("utf-8", errors="ignore").lower()
                    if ("chrome" in cmd or "msedge" in cmd) and "--user-data-dir=" in cmd and (abs_posix in cmd or abs_norm in cmd):
                        if pipe_only and "--remote-debugging-pipe" not in cmd:
                            continue
                        pids.append(int(ent))
                except Exception:
                    pass
        except Exception:
            pass
    return pids


def _terminate_chrome_pids(pids, wait_sec=3.0):
    if not pids:
        return
    import subprocess
    if sys.platform == "win32":
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        for pid in pids:
            try:
                subprocess.run(
                    ["taskkill", "/PID", str(int(pid)), "/T", "/F"],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=4, creationflags=creation_flags
                )
            except Exception:
                pass
    else:
        import signal
        for pid in pids:
            try:
                os.kill(int(pid), signal.SIGTERM)
            except Exception:
                pass
    time.sleep(min(1.5, max(0.3, wait_sec)))


def _prepare_ai_window_prefs(profile_dir):
    for pref_sub in ["Default/Preferences", "Preferences"]:
        pref_path = os.path.join(profile_dir, pref_sub)
        if os.path.exists(pref_path):
            try:
                with open(pref_path, "r", encoding="utf-8") as pf:
                    data = json.load(pf) or {}
                if not isinstance(data, dict):
                    data = {}
                prof = data.setdefault("profile", {})
                if isinstance(prof, dict):
                    prof["exit_type"] = "Normal"
                    prof["exited_cleanly"] = True
                sess = data.setdefault("session", {})
                if isinstance(sess, dict):
                    sess["restore_on_startup"] = 5
                with open(pref_path, "w", encoding="utf-8") as pf:
                    json.dump(data, pf, ensure_ascii=False)
            except Exception:
                try:
                    with open(pref_path, "r", encoding="utf-8") as pf:
                        p_text = pf.read()
                    p_text = p_text.replace('"exit_type":"Crashed"', '"exit_type":"Normal"')
                    p_text = p_text.replace('"exited_cleanly":false', '"exited_cleanly":true')
                    with open(pref_path, "w", encoding="utf-8") as pf:
                        pf.write(p_text)
                except Exception:
                    pass
    sess_dir = os.path.join(profile_dir, "Default", "Sessions")
    if os.path.isdir(sess_dir):
        try:
            for fn in os.listdir(sess_dir):
                if fn.startswith(("Session_", "Tabs_")):
                    try:
                        os.remove(os.path.join(sess_dir, fn))
                    except Exception:
                        pass
        except Exception:
            pass


def is_ai_cdp_port_open(port=AI_WINDOW_CDP_PORT):
    import urllib.request
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{int(port)}/json/version", timeout=1.2) as res:
            return res.status == 200
    except Exception:
        return False


def ensure_fixed_ai_window_launched(base_dir=BASE_DIR, cdp_port=AI_WINDOW_CDP_PORT, initial_url=None, log_func=print):
    with _AI_WINDOW_LAUNCH_LOCK:
        profile_dir = get_ai_window_profile_dir(base_dir)
        if is_ai_cdp_port_open(cdp_port):
            return True, profile_dir
        existing_pids = _find_chrome_pids_for_profile(profile_dir, pipe_only=False)
        if existing_pids:
            for _ in range(4):
                time.sleep(0.4)
                if is_ai_cdp_port_open(cdp_port):
                    return True, profile_dir
            if log_func:
                log_func("[INFO] 🧠 既存の非CDP AIウィンドウを整理してから1窓のみで再起動します...")
            _terminate_chrome_pids(existing_pids, wait_sec=2.0)
        _cleanup_profile_locks(profile_dir)
        _prepare_ai_window_prefs(profile_dir)
        chrome_exe = _find_chrome_executable_for_ai()
        if not chrome_exe:
            raise RuntimeError("AI専用ウィンドウ起動用の Google Chrome が見つかりませんでした。")
        import subprocess
        cmd = [
            chrome_exe,
            f"--user-data-dir={profile_dir}",
            f"--remote-debugging-port={int(cdp_port)}",
            "--no-first-run",
            "--no-default-browser-check",
            "--lang=ja-JP",
            "--disable-session-crashed-bubble",
            "--hide-crash-restore-bubble",
            "--disable-features=InfiniteSessionRestore",
            "--window-size=1180,820",
        ]
        if initial_url:
            cmd.append(str(initial_url))
        creation_flags = 0
        if sys.platform == "win32":
            creation_flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        if log_func:
            log_func(f"[INFO] 🧠 【固定AI専用Window起動】プロファイル: chrome_profiles/{AI_WINDOW_PROFILE_NAME} (CDP:{cdp_port})")
        subprocess.Popen(cmd, cwd=base_dir, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=creation_flags)
        for _ in range(24):
            if is_ai_cdp_port_open(cdp_port):
                return True, profile_dir
            time.sleep(0.5)
        return is_ai_cdp_port_open(cdp_port), profile_dir


_AI_ENGINE_ALIASES = {
    "openai": "chatgpt",
    "gpt": "chatgpt",
    "google": "gemini",
    "anthropic": "claude",
    "duck": "duckai",
    "ddg": "duckai",
    "or": "openrouter",
    "chrome": "chrome_builtin",
}


def _split_multi_tokens(raw_val):
    s = str(raw_val or "").strip()
    if not s:
        return []
    return [t.strip() for t in re.split(r'[,，、;；]+', s) if t.strip()]


def _normalize_engine_token(tok):
    t = str(tok or "").strip().lower()
    t = _AI_ENGINE_ALIASES.get(t, t)
    return t if t in AI_ENGINE_CHOICES else None


def _parse_engine_list(raw_val, default_fallback="duckai"):
    s = str(raw_val or "").strip()
    if not s:
        return [default_fallback] if default_fallback else []
    res = []
    for tok in re.split(r'[,，、;；/\s]+', s):
        norm = _normalize_engine_token(tok)
        if norm and norm not in res:
            res.append(norm)
    if not res and default_fallback:
        return [default_fallback]
    return res


def _normalize_engine_list_str(raw_val, default_fallback="duckai"):
    return ", ".join(_parse_engine_list(raw_val, default_fallback=default_fallback))


def _extract_engine_tag(token):
    s = str(token or "").strip()
    if ":" in s:
        head, tail = s.split(":", 1)
        h_low = head.strip().lower()
        if h_low not in ("http", "https"):
            norm = _normalize_engine_token(h_low)
            if norm and tail.strip():
                return norm, tail.strip()
    return None, s


def _infer_engine_for_key_or_url(val_str):
    v = str(val_str or "").strip()
    vl = v.lower()
    if v.startswith("gsk_") or "api.groq.com" in vl or "console.groq.com" in vl:
        return "groq"
    if v.startswith("AIza") or "generativelanguage.googleapis.com" in vl or "gemini.google.com" in vl:
        return "gemini"
    if v.startswith("sk-ant-") or "api.anthropic.com" in vl or "claude.ai" in vl:
        return "claude"
    if v.startswith("sk-or-") or "openrouter.ai" in vl:
        return "openrouter"
    if v.startswith("xai-") or "api.x.ai" in vl:
        return "grok"
    if v.startswith("pplx-") or "api.perplexity.ai" in vl:
        return "perplexity"
    if v.startswith("sk-proj-") or "api.openai.com" in vl:
        return "chatgpt"
    if ":11434" in vl or "ollama" in vl:
        return "ollama"
    if ":1234" in vl or "lmstudio" in vl:
        return "lmstudio"
    if "api.deepseek.com" in vl:
        return "deepseek"
    if "api.mistral.ai" in vl:
        return "mistral"
    if "dashscope" in vl:
        return "qwen"
    if "api.venice.ai" in vl:
        return "venice"
    return None


def _parse_api_key_or_url(raw_val, default_url=""):
    s = str(raw_val or "").strip()
    if not s or s.lower() in ("none", "nan", "-"):
        return "", default_url
    if "|" in s:
        p1, p2 = [x.strip() for x in s.split("|", 1)]
        if p1.lower().startswith(("http://", "https://")):
            return p2, p1
        return p1, (p2 or default_url)
    if s.lower().startswith(("http://", "https://")):
        return "", s
    return s, default_url


def _resolve_ai_engine_chain(primary_engine, fallback_raw):
    primaries = _parse_engine_list(primary_engine, default_fallback="duckai")
    fb_s = str(fallback_raw or "").strip()
    if fb_s.lower() in {"off", "none", "no", "0", "false", "なし", "無効", "切替なし"}:
        return primaries
    chain = list(primaries)
    if fb_s:
        for t in _parse_engine_list(fb_s, default_fallback=None):
            if t not in chain:
                chain.append(t)
        return chain
    for cand in [
        "duckai", "talkai", "deepai", "venice",
        "groq", "gemini", "chatgpt", "grok", "claude", "copilot",
        "perplexity", "deepseek", "qwen", "mistral", "openrouter",
        "ollama", "lmstudio", "chrome_builtin",
    ]:
        if cand not in chain:
            chain.append(cand)
    return chain


def _resolve_engine_models_for_chain(model_raw, primary_engines, chain):
    tokens = _split_multi_tokens(model_raw)
    per_eng = {e: [] for e in chain}
    untagged = []
    for tok in tokens:
        tag, val = _extract_engine_tag(tok)
        if tag:
            if tag in per_eng and val not in per_eng[tag]:
                per_eng[tag].append(val)
        else:
            untagged.append(val)
    if untagged:
        if len(primary_engines) <= 1 and primary_engines:
            target = primary_engines[0]
            if target in per_eng:
                for u in untagged:
                    if u not in per_eng[target]:
                        per_eng[target].append(u)
        else:
            for idx, u in enumerate(untagged):
                if idx < len(primary_engines):
                    target = primary_engines[idx]
                    if target in per_eng and u not in per_eng[target]:
                        per_eng[target].append(u)
    return per_eng


def _resolve_engine_credentials_for_chain(raw_key_or_url, primary_engines, chain):
    tokens = _split_multi_tokens(raw_key_or_url)
    per_eng = {e: [] for e in chain}
    untagged = []
    for tok in tokens:
        tag, val = _extract_engine_tag(tok)
        if tag:
            if tag in per_eng and val not in per_eng[tag]:
                per_eng[tag].append(val)
            continue
        inferred = _infer_engine_for_key_or_url(val)
        if inferred and inferred in per_eng:
            if val not in per_eng[inferred]:
                per_eng[inferred].append(val)
            continue
        untagged.append(tok)
    if untagged:
        if len(primary_engines) <= 1 and primary_engines:
            target = primary_engines[0]
            if target in per_eng:
                for u in untagged:
                    if u not in per_eng[target]:
                        per_eng[target].append(u)
        else:
            for idx, u in enumerate(untagged):
                if idx < len(primary_engines):
                    target = primary_engines[idx]
                    if target in per_eng and u not in per_eng[target]:
                        per_eng[target].append(u)
    return per_eng


def _strip_ai_ui_noise_lines(text):
    t = str(text or "").strip()
    if not t:
        return ""
    exact_noise = {
        "duck.ai", "duck.aiの発言", "duck.ai said", "chatgptの発言", "chatgpt said",
        "geminiの回答", "claudeの回答", "grokの発言", "deepseekの発言", "aiの発言", "aiの回答",
        "アシスタント", "assistant", "下書き", "draft", "思考中", "考え中", "生成中", "検索中",
        "thinking", "thinking...", "generating", "generating...", "searching", "searching...",
        "最新の回答へ移動", "下にスクロール", "トップへ戻る", "scroll to bottom", "jump to latest response",
        "コピー", "コピーしました", "copy", "copied", "再生成", "regenerate",
        "良い回答", "悪い回答", "good response", "bad response",
        "共有", "share", "編集", "edit", "新しいチャット", "new chat",
        "音声で読み上げる", "read aloud", "停止", "stop", "送信", "send",
        "他のaiモデルに質問", "他のAIモデルに質問", "ask another ai model", "ask another model",
        "ダウンロード", "download", "やり直す", "retry",
        "投稿文を作成しました", "投稿文を作成しました。", "投稿を作成しました", "投稿を作成しました。",
        "こちらが投稿文です", "こちらが投稿文です。", "以上です", "以上です。",
    }
    model_badge_re = re.compile(
        r'^(?:gpt-[\w.-]+|claude[\s-][\w.-]+|gemini[\s-][\w.-]+|llama[\s-][\w.-]+|mixtral[\s-][\w.-]+|mistral[\s-][\w.-]+|qwen[\s-]?[\w.-]+|deepseek[\s-]?[\w.-]+|o[134](?:-[\w.-]+)?)(?:\s+[a-z0-9._-]+){0,3}$',
        re.IGNORECASE,
    )
    disclaimer_re = re.compile(
        r'すべてのチャットは非公開|aiは間違える可能性|all chats are private|ai can make mistakes|check important info',
        re.IGNORECASE,
    )
    out = []
    for line in t.splitlines():
        s = line.strip()
        if not s:
            if out and out[-1] != "":
                out.append("")
            continue
        sl = s.lower()
        if sl in exact_noise or s in exact_noise:
            continue
        if disclaimer_re.search(s):
            continue
        if len(s) <= 36 and model_badge_re.match(s):
            continue
        out.append(line.rstrip())
    return "\n".join(out).strip()


def _strip_prompt_fragments(raw_text, prompt_text=""):
    t = str(raw_text or "").strip()
    if not t:
        return ""
    prompt_header_re = re.compile(
        r'^【(?:作成方針（自動選定）|キャラ設定|ターゲット層|投稿の目的|口調・文体ルール|プロンプト|お手本サンプル|NGルール|出力ルール|今回の個別作成指示・テーマ)】'
    )
    compact_p = re.sub(r'\s+', '', str(prompt_text or ""))
    if prompt_text:
        end_marker = "選んだ1つのポストの本文のみを出力してください。"
        prompt_lines = [ln.strip() for ln in str(prompt_text).splitlines() if ln.strip()]
        cut_pos = -1
        m_idx = t.rfind(end_marker)
        if m_idx != -1:
            p = m_idx + len(end_marker)
            if len(t[p:].strip()) >= 4:
                cut_pos = max(cut_pos, p)
        for pl in prompt_lines[-4:]:
            if len(pl) >= 8:
                p_idx = t.rfind(pl)
                if p_idx != -1:
                    p = p_idx + len(pl)
                    if len(t[p:].strip()) >= 4:
                        cut_pos = max(cut_pos, p)
        if cut_pos > 0:
            t = t[cut_pos:].strip()
    out_lines = []
    for ln in t.splitlines():
        s = ln.strip()
        if not s:
            if out_lines and out_lines[-1] != "":
                out_lines.append("")
            continue
        if prompt_header_re.match(s):
            continue
        s_comp = re.sub(r'\s+', '', s)
        if compact_p and len(s_comp) >= 6 and s_comp in compact_p:
            continue
        if "※出力は投稿するメッセージ本文のみ" in s or "選んだ1つのポストの本文のみを出力してください" in s:
            continue
        out_lines.append(ln.rstrip())
    res = "\n".join(out_lines).strip()
    if not res:
        return ""
    res_comp = re.sub(r'\s+', '', res)
    if compact_p and len(res_comp) >= 6 and res_comp in compact_p:
        return ""
    return res


def _clean_ai_generated_post_text(raw_text, max_len=495, prompt_text=""):
    t = str(raw_text or "").strip()
    if not t:
        return ""
    t = re.sub(r'^```[a-zA-Z0-9_-]*\s*', '', t)
    t = re.sub(r'\s*```$', '', t).strip()
    t = _strip_prompt_fragments(t, prompt_text=prompt_text)
    t = _strip_ai_ui_noise_lines(t)
    if not t:
        return ""
    lines = t.splitlines()
    while len(lines) > 1 and any(lines[0].strip().startswith(prefix) for prefix in [
        "以下に", "ご提案", "投稿案", "ツイート案", "Threads用", "はい、", "承知しました", "了解しました",
        "投稿文を作成しました", "投稿を作成しました", "こちらが投稿文", "Sure", "Here is"
    ]):
        lines = lines[1:]
    while len(lines) > 1 and lines[-1].strip() in [
        "投稿文を作成しました", "投稿文を作成しました。", "投稿を作成しました", "投稿を作成しました。",
        "以上です", "以上です。", "いかがでしょうか", "いかがでしょうか？", "いかがでしょうか。"
    ]:
        lines = lines[:-1]
    t = "\n".join(lines).strip()
    t = re.sub(r'^(?:以下[、:]|投稿文[:：]?|本文[:：]|ポスト案[:：]|ポスト[:：]|案\d*[:：]|Threads投稿[:：])\s*\n?', '', t, flags=re.IGNORECASE).strip()
    if len(t) >= 2 and ((t[0] == "「" and t[-1] == "」") or (t[0] == '"' and t[-1] == '"')):
        inner = t[1:-1].strip()
        if inner:
            t = inner
    t = re.sub(r'\n{3,}', '\n\n', t).strip()
    return (t[:max_len - 1].rstrip() + "…") if len(t) > max_len else t


def _http_post_json(url, payload, headers=None, timeout=45):
    import urllib.request
    import urllib.error
    req_headers = {"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"}
    if headers:
        req_headers.update(headers)
    req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers=req_headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as res:
            return json.loads(res.read().decode("utf-8", errors="replace"))
    except urllib.error.HTTPError as he:
        raise RuntimeError(f"HTTP {he.code}: {he.reason}")


def _build_ai_generation_prompt(account_name, ai_cfg, prompt_row=None):
    blocks = [FIXED_AI_SYSTEM_PROMPT]
    sections = (ai_cfg or {}).get("prompt_sections") or {}

    has_any_section = False
    for sec_key, sec_label, _def_val in AI_PROMPT_SECTION_META:
        s_val = str(sections.get(sec_key) or "").strip()
        if s_val:
            has_any_section = True
            blocks.append(f"【{sec_label}】\n{s_val}")

    custom_inst = ""
    if isinstance(prompt_row, dict):
        custom_inst = str(prompt_row.get("prompt") or "").strip()

    if custom_inst and custom_inst != str(sections.get("prompt") or "").strip():
        blocks.append(f"【今回の個別作成指示・テーマ】\n{custom_inst}")
    elif not has_any_section:
        blocks.append("【プロンプト】\n自然な日本語のThreads投稿文を1つ作成してください。")

    return "\n\n".join(blocks)


_AI_WINDOW_DOMAIN_MATCHERS = {
    "duckai": ("duck.ai", "duckduckgo.com"),
    "talkai": ("talkai.info",),
    "deepai": ("deepai.org",),
    "venice": ("venice.ai",),
    "chatgpt": ("chatgpt.com", "chat.openai.com"),
    "gemini": ("gemini.google.com",),
    "claude": ("claude.ai",),
    "grok": ("grok.com", "x.com/i/grok"),
    "deepseek": ("chat.deepseek.com", "deepseek.com"),
    "qwen": ("chat.qwen.ai", "qwen.ai"),
    "mistral": ("chat.mistral.ai",),
    "perplexity": ("perplexity.ai",),
    "copilot": ("copilot.microsoft.com",),
}


def _dismiss_ai_window_modals(page, engine):
    consent_selectors = [
        'button:has-text("Get Started")',
        'button:has-text("Get started")',
        'button:has-text("Start chatting")',
        'button:has-text("Start")',
        'button:has-text("はじめる")',
        'button:has-text("チャットを開始")',
        'button:has-text("同意して続ける")',
        'button:has-text("Agree and Continue")',
        'button:has-text("Accept all")',
        'button:has-text("Accept All")',
        'button:has-text("Consent")',
        'button:has-text("I agree")',
        'button:has-text("I Agree")',
        'button:has-text("Agree")',
        'button:has-text("同意")',
        'button:has-text("承諾")',
        'button:has-text("続行")',
        'button:has-text("Accept")',
        'button:has-text("Continue")',
        'button:has-text("Next")',
        'button:has-text("次へ")',
        'button:has-text("Got it")',
        'button:has-text("OK")',
        'button:has-text("閉じる")',
    ]
    for _pass in range(3 if engine == "duckai" else 2):
        clicked_any = False
        for ok_sel in consent_selectors:
            try:
                loc = page.locator(ok_sel).first
                if loc.is_visible(timeout=400):
                    loc.click(force=True)
                    clicked_any = True
                    time.sleep(0.5)
            except Exception:
                pass
        if not clicked_any:
            break


def _generate_via_fixed_ai_window(playwright_obj, engine, prompt_text, base_dir=BASE_DIR, log_func=print):
    ensure_fixed_ai_window_launched(
        base_dir=base_dir,
        cdp_port=AI_WINDOW_CDP_PORT,
        initial_url=AI_ENGINE_DEFAULTS.get(engine, {}).get("web_url") or "https://duckduckgo.com/?q=DuckDuckGo+AI+Chat&ia=chat",
        log_func=log_func,
    )
    if not playwright_obj:
        raise RuntimeError("Playwright未初期化")
    browser = playwright_obj.chromium.connect_over_cdp(f"http://127.0.0.1:{AI_WINDOW_CDP_PORT}")
    try:
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()
        if engine == "chrome_builtin":
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            out = page.evaluate("""async (prompt) => {
                const lmApi = window.LanguageModel || (window.ai && window.ai.languageModel);
                if (!lmApi || typeof lmApi.create !== 'function') throw new Error('LanguageModel API無効');
                const sess = await lmApi.create();
                return String(await sess.prompt(prompt) || '').trim();
            }""", prompt_text)
            if out:
                if log_func:
                    log_func(f"[INFO] 🤖 [AI専用Window] {engine} の出力を取得しました ({len(out)}文字)")
                return out
            raise RuntimeError("chrome_builtin から空応答")
        target_url = AI_ENGINE_DEFAULTS.get(engine, {}).get("web_url") or "https://chatgpt.com/"
        domain_keys = _AI_WINDOW_DOMAIN_MATCHERS.get(engine) or (
            (target_url.split("/")[2].replace("www.", ""),) if "://" in target_url else ()
        )
        page = None
        blank_page = None
        for ep in ctx.pages:
            try:
                cur_u = (ep.url or "").lower()
                if any(dk in cur_u for dk in domain_keys if dk):
                    page = ep
                    break
                if blank_page is None and cur_u in ("", "about:blank", "chrome://newtab/", "edge://newtab/"):
                    blank_page = ep
            except Exception:
                pass
        if page is None:
            page = blank_page if blank_page is not None else ctx.new_page()
            try:
                page.bring_to_front()
            except Exception:
                pass
            page.goto(target_url, timeout=45000, wait_until="domcontentloaded")
            time.sleep(2.0)
        else:
            try:
                page.bring_to_front()
            except Exception:
                pass
            time.sleep(0.8)

        input_selectors = [
            'textarea[name="prompt"]',
            'textarea[name="user-prompt"]',
            'textarea#prompt-textarea',
            'div#prompt-textarea[contenteditable="true"]',
            'div[contenteditable="true"][role="textbox"]',
            '[contenteditable="plaintext-only"][role="textbox"]',
            'rich-textarea div[contenteditable="true"]',
            'textarea[placeholder]',
            'textarea',
            '[contenteditable="true"]',
            '[contenteditable="plaintext-only"]',
            '[role="textbox"]',
            'input[type="text"]:not([readonly]):not([disabled])',
            'input:not([type]):not([readonly]):not([disabled])',
        ]

        target_ctx = page

        def _find_input_box():
            nonlocal target_ctx
            search_targets = [page]
            try:
                for fr in page.frames:
                    if fr != page.main_frame:
                        search_targets.append(fr)
            except Exception:
                pass
            for ctx_node in search_targets:
                for sel in input_selectors:
                    try:
                        cand = ctx_node.locator(sel).last
                        if cand.is_visible(timeout=400):
                            target_ctx = ctx_node
                            return cand
                    except Exception:
                        pass
            return None

        def _is_login_wall():
            try:
                cur_u_low = (page.url or "").lower()
                if any(ak in cur_u_low for ak in ("accounts.google.com", "login.microsoftonline.com", "appleid.apple.com", "/auth/", "/signin", "/login")):
                    return True
                if engine in ("chatgpt", "gemini", "claude", "grok", "deepseek", "qwen", "mistral", "perplexity"):
                    for l_sel in ['button:has-text("Log in")', 'button:has-text("Sign in")', 'a:has-text("Log in")', 'a:has-text("Sign in")', 'button:has-text("ログイン")', 'a:has-text("ログイン")']:
                        try:
                            if page.locator(l_sel).first.is_visible(timeout=250):
                                return True
                        except Exception:
                            pass
            except Exception:
                pass
            return False

        _dismiss_ai_window_modals(page, engine)
        input_box = _find_input_box()
        if not input_box:
            for _wait_inp in range(6):
                if _wait_inp >= 2 and _is_login_wall():
                    raise RuntimeError(f"AI専用ウィンドウ({engine})はログインが必要な画面のため自動スキップします")
                time.sleep(1.0)
                _dismiss_ai_window_modals(page, engine)
                input_box = _find_input_box()
                if input_box:
                    break
        if not input_box:
            raise RuntimeError(f"AI専用ウィンドウ({engine})で入力欄が見つかりません")

        try:
            ctx.grant_permissions(["clipboard-read", "clipboard-write"])
        except Exception:
            pass
        for hook_ctx in ([page] if target_ctx == page else [target_ctx, page]):
            try:
                hook_ctx.evaluate("""() => {
                    window.__aiAutoCapturedText = '';
                    if (window.__aiAutoClipboardHooked) return;
                    window.__aiAutoClipboardHooked = true;
                    try {
                        const clipProto = (typeof Clipboard !== 'undefined' && Clipboard.prototype) ? Clipboard.prototype : null;
                        const origProtoWrite = clipProto && clipProto.writeText ? clipProto.writeText : null;
                        if (clipProto && origProtoWrite) {
                            clipProto.writeText = async function(text) {
                                window.__aiAutoCapturedText = String(text || '');
                                try { return await origProtoWrite.call(this, text); } catch (e) {}
                            };
                        } else if (navigator.clipboard && navigator.clipboard.writeText) {
                            const origWrite = navigator.clipboard.writeText.bind(navigator.clipboard);
                            navigator.clipboard.writeText = async function(text) {
                                window.__aiAutoCapturedText = String(text || '');
                                try { return await origWrite(text); } catch (e) {}
                            };
                        }
                        const origProtoWriteItems = clipProto && clipProto.write ? clipProto.write : null;
                        if (clipProto && origProtoWriteItems) {
                            clipProto.write = async function(items) {
                                try {
                                    for (const item of (items || [])) {
                                        if (item && item.types && item.types.includes('text/plain')) {
                                            const blob = await item.getType('text/plain');
                                            window.__aiAutoCapturedText = await blob.text();
                                            break;
                                        }
                                    }
                                } catch (e) {}
                                try { return await origProtoWriteItems.call(this, items); } catch (e) {}
                            };
                        }
                    } catch (e) {}
                }""")
            except Exception:
                pass

        before_snap = target_ctx.evaluate("""() => {
            const SELS = [
                '[data-message-author-role="assistant"] .markdown', '[data-message-author-role="assistant"] .prose',
                '[data-message-author-role="assistant"]', '[data-content="ai-message"]',
                '.ds-markdown', '.response-content-markdown', '.model-response-text',
                'model-response .markdown', 'model-response', 'message-content', '.markdown-body', '.markdown', '.prose',
                '.outputBox', '.bot-message', '.ai-response',
                '[data-testid*="assistant"]', '[data-testid*="response"]', '[data-testid="chat-message"]',
                '[data-testid*="message"]', '[class*="assistant"]', '[class*="response"]',
                '[class*="answer"]', 'article', '[role="article"]'
            ];
            const SKIP_TAGS = new Set(['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEMPLATE', 'SVG', 'TEXTAREA', 'INPUT', 'SELECT', 'OPTION', 'CANVAS', 'IFRAME']);
            const BLOCK_TAGS = new Set(['P', 'DIV', 'LI', 'BR', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE', 'PRE', 'SECTION', 'ARTICLE', 'TR', 'DT', 'DD']);

            function extractLiveText(rootNode) {
                if (!rootNode) return '';
                const parts = [];
                function walk(node, inPre) {
                    if (!node) return;
                    const nt = node.nodeType;
                    if (nt === 3) {
                        const rawVal = node.nodeValue || '';
                        const v = inPre ? rawVal : rawVal.replace(/[\\r\\n\\t ]+/g, ' ');
                        if (v && (inPre || v.trim())) parts.push(v);
                        return;
                    }
                    if (nt !== 1 && nt !== 11) return;
                    if (nt === 1) {
                        const tag = (node.tagName || '').toUpperCase();
                        if (SKIP_TAGS.has(tag)) return;
                        if (tag === 'BUTTON' || tag === 'NAV' || tag === 'HEADER' || tag === 'FOOTER' || tag === 'ASIDE') return;
                        if (node.hidden || (node.getAttribute && node.getAttribute('aria-hidden') === 'true')) return;
                        const cls = node.classList;
                        if (cls && (cls.contains('sr-only') || cls.contains('visually-hidden'))) return;
                        const role = node.getAttribute ? (node.getAttribute('role') || '').toLowerCase() : '';
                        if (role === 'button' && !(node.querySelector && node.querySelector('p, article, .markdown, .prose, blockquote'))) return;
                        try {
                            const cs = window.getComputedStyle(node);
                            if (cs && (cs.display === 'none' || cs.visibility === 'hidden')) return;
                        } catch (e) {}
                        const isBlock = BLOCK_TAGS.has(tag);
                        const childPre = inPre || tag === 'PRE';
                        if (isBlock) parts.push('\\n');
                        if (node.shadowRoot) walk(node.shadowRoot, childPre);
                        for (let c = node.firstChild; c; c = c.nextSibling) {
                            walk(c, childPre);
                        }
                        if (isBlock) parts.push('\\n');
                        return;
                    }
                    for (let c = node.firstChild; c; c = c.nextSibling) {
                        walk(c, inPre);
                    }
                }
                walk(rootNode, false);
                let res = parts.join('').replace(/[ \\t\\r\\f\\v]+/g, ' ').replace(/ *\\n */g, '\\n').replace(/\\n{3,}/g, '\\n\\n').trim();
                if (!res && rootNode.innerText) {
                    res = String(rootNode.innerText || '').trim();
                }
                return res;
            }

            function collectNodes(root) {
                const texts = [];
                for (const s of SELS) {
                    try {
                        root.querySelectorAll(s).forEach(el => {
                            const role = (el.getAttribute && (el.getAttribute('data-message-author-role') || '').toLowerCase()) || '';
                            if (role === 'user') return;
                            if (el.querySelector && el.querySelector('textarea, input')) return;
                            const t = extractLiveText(el);
                            if (t) texts.push(t);
                        });
                    } catch (e) {}
                }
                try {
                    root.querySelectorAll('*').forEach(el => { if (el.shadowRoot) texts.push(...collectNodes(el.shadowRoot)); });
                } catch (e) {}
                return texts;
            }
            function countCopyBtns(root) {
                let cnt = 0;
                const copySel = 'button[aria-label*="Copy" i], button[aria-label*="コピー"], button[aria-label*="クリップボード"], button[aria-label*="clipboard" i], button[title*="Copy" i], button[title*="コピー"], button[data-testid*="copy" i]';
                try { cnt += root.querySelectorAll(copySel).length; } catch (e) {}
                try { root.querySelectorAll('*').forEach(el => { if (el.shadowRoot) cnt += countCopyBtns(el.shadowRoot); }); } catch (e) {}
                return cnt;
            }
            const bodyTxt = document.body ? extractLiveText(document.body) : '';
            return {
                selTexts: collectNodes(document),
                bodyLines: bodyTxt.split(/\\r?\\n/).map(x => x.trim()).filter(Boolean),
                copyBtnCount: countCopyBtns(document),
            };
        }""")

        input_box.click()
        time.sleep(0.25)
        try:
            tag_name = input_box.evaluate("el => (el.tagName || '').toLowerCase()")
        except Exception:
            tag_name = "textarea"

        if tag_name in ("textarea", "input"):
            try:
                input_box.fill(prompt_text)
                cur_val = input_box.input_value()
                if not cur_val or len(cur_val.strip()) < 10:
                    raise ValueError("empty after fill")
            except Exception:
                input_box.click()
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                page.keyboard.insert_text(prompt_text)
        else:
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            page.keyboard.insert_text(prompt_text)
        time.sleep(0.5)

        submitted = False
        for send_sel in [
            'button[data-testid="send-button"]', 'button[aria-label*="Send" i]',
            'button[aria-label*="送信"]', 'button[type="submit"]',
            'button:has-text("送信")', 'button:has-text("Send")',
        ]:
            try:
                sb = target_ctx.locator(send_sel).last
                if sb.is_visible(timeout=500) and sb.is_enabled(timeout=500):
                    sb.click()
                    submitted = True
                    break
            except Exception:
                pass
        if not submitted:
            page.keyboard.press("Enter")
            time.sleep(0.5)
            try:
                still_val = input_box.input_value() if tag_name in ("textarea", "input") else input_box.inner_text()
                if still_val and len(still_val.strip()) > 20:
                    page.keyboard.press("Control+Enter")
            except Exception:
                pass

        def _evaluate_ai_output(allow_body_diff=False):
            return target_ctx.evaluate("""(args) => {
                const beforeSelCounts = Object.create(null);
                for (const t of (args.beforeSelTexts || [])) beforeSelCounts[t] = (beforeSelCounts[t] || 0) + 1;
                const beforeBodyCounts = Object.create(null);
                for (const ln of (args.beforeBodyLines || [])) beforeBodyCounts[ln] = (beforeBodyCounts[ln] || 0) + 1;
                const promptFull = String(args.promptText || '');
                const promptCompact = promptFull.replace(/\\s+/g, '');
                const promptLines = promptFull.split(/\\r?\\n/).map(x => x.trim()).filter(Boolean);
                const promptLineSet = new Set(promptLines);
                const promptHeaderRe = /^【(?:作成方針（自動選定）|キャラ設定|ターゲット層|投稿の目的|口調・文体ルール|プロンプト|お手本サンプル|NGルール|出力ルール|今回の個別作成指示・テーマ)】/;
                const endMarker = '選んだ1つのポストの本文のみを出力してください。';
                const uiNoise = new Set([
                    'copy', 'コピー', 'コピーしました', 'copied', 'regenerate', '再生成',
                    'like', 'dislike', 'share', '共有', 'edit', '編集', 'send', '送信',
                    'stop', '停止', '中止', 'new chat', '新しいチャット',
                    '他のaiモデルに質問', 'ask another ai model', 'ask another model',
                    'ダウンロード', 'download', 'やり直す', 'retry',
                    'duck.ai', 'duck.aiの発言', 'duck.ai said', 'chatgptの発言', 'chatgpt said',
                    'geminiの回答', 'claudeの回答', 'grokの発言', 'deepseekの発言',
                    'aiの発言', 'aiの回答', 'アシスタント', 'assistant',
                    'talkai', 'deepai', 'venice', 'chatgpt', 'gemini', 'claude', 'grok',
                    'read aloud', '音声で読み上げる', '読み上げる',
                    '良い回答', '悪い回答', 'good response', 'bad response',
                    '下書き', 'draft', '思考中', '考え中', '生成中', '検索中',
                    'thinking', 'thinking...', 'generating', 'generating...', 'searching', 'searching...',
                    '最新の回答へ移動', '下にスクロール', 'トップへ戻る',
                    'scroll to bottom', 'jump to latest response',
                    '投稿文を作成しました', '投稿文を作成しました。',
                    '投稿を作成しました', '投稿を作成しました。'
                ]);
                const modelBadgeRe = /^(?:gpt-[\\w.-]+|claude[\\s-][\\w.-]+|gemini[\\s-][\\w.-]+|llama[\\s-][\\w.-]+|mixtral[\\s-][\\w.-]+|mistral[\\s-][\\w.-]+|qwen[\\s-]?[\\w.-]+|deepseek[\\s-]?[\\w.-]+|o[134](?:-[\\w.-]+)?)(?:\\s+[a-z0-9._-]+){0,3}$/i;
                const disclaimerRe = /すべてのチャットは非公開|aiは間違える可能性|all chats are private|ai can make mistakes|check important info/i;
                const SKIP_TAGS = new Set(['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEMPLATE', 'SVG', 'TEXTAREA', 'INPUT', 'SELECT', 'OPTION', 'CANVAS', 'IFRAME']);
                const BLOCK_TAGS = new Set(['P', 'DIV', 'LI', 'BR', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE', 'PRE', 'SECTION', 'ARTICLE', 'TR', 'DT', 'DD']);

                function extractLiveText(rootNode) {
                    if (!rootNode) return '';
                    const parts = [];
                    function walk(node, inPre) {
                        if (!node) return;
                        const nt = node.nodeType;
                        if (nt === 3) {
                            const rawVal = node.nodeValue || '';
                            const v = inPre ? rawVal : rawVal.replace(/[\\r\\n\\t ]+/g, ' ');
                            if (v && (inPre || v.trim())) parts.push(v);
                            return;
                        }
                        if (nt !== 1 && nt !== 11) return;
                        if (nt === 1) {
                            const tag = (node.tagName || '').toUpperCase();
                            if (SKIP_TAGS.has(tag)) return;
                            if (tag === 'BUTTON' || tag === 'NAV' || tag === 'HEADER' || tag === 'FOOTER' || tag === 'ASIDE') return;
                            if (node.hidden || (node.getAttribute && node.getAttribute('aria-hidden') === 'true')) return;
                            const cls = node.classList;
                            if (cls && (cls.contains('sr-only') || cls.contains('visually-hidden'))) return;
                            const role = node.getAttribute ? (node.getAttribute('role') || '').toLowerCase() : '';
                            if (role === 'button' && !(node.querySelector && node.querySelector('p, article, .markdown, .prose, blockquote'))) return;
                            try {
                                const cs = window.getComputedStyle(node);
                                if (cs && (cs.display === 'none' || cs.visibility === 'hidden')) return;
                            } catch (e) {}
                            const isBlock = BLOCK_TAGS.has(tag);
                            const childPre = inPre || tag === 'PRE';
                            if (isBlock) parts.push('\\n');
                            if (node.shadowRoot) walk(node.shadowRoot, childPre);
                            for (let c = node.firstChild; c; c = c.nextSibling) {
                                walk(c, childPre);
                            }
                            if (isBlock) parts.push('\\n');
                            return;
                        }
                        for (let c = node.firstChild; c; c = c.nextSibling) {
                            walk(c, inPre);
                        }
                    }
                    walk(rootNode, false);
                    let res = parts.join('').replace(/[ \\t\\r\\f\\v]+/g, ' ').replace(/ *\\n */g, '\\n').replace(/\\n{3,}/g, '\\n\\n').trim();
                    if (!res && rootNode.innerText) {
                        res = String(rootNode.innerText || '').trim();
                    }
                    return res;
                }

                function cleanCandidate(raw) {
                    let t = String(raw || '').trim();
                    if (!t) return '';
                    const tCompact = t.replace(/\\s+/g, '');
                    if (tCompact.length >= 6 && promptCompact.includes(tCompact)) return '';
                    let cutPos = -1;
                    const markerIdx = t.lastIndexOf(endMarker);
                    if (markerIdx !== -1) {
                        const p = markerIdx + endMarker.length;
                        cutPos = Math.max(cutPos, p);
                    }
                    for (let k = promptLines.length - 1; k >= Math.max(0, promptLines.length - 6); k--) {
                        const pl = promptLines[k];
                        if (pl && pl.length >= 8) {
                            const lpIdx = t.lastIndexOf(pl);
                            if (lpIdx !== -1) {
                                const p = lpIdx + pl.length;
                                cutPos = Math.max(cutPos, p);
                            }
                        }
                    }
                    if (cutPos > 0) {
                        const after = t.slice(cutPos).trim();
                        if (after.length < 4) return '';
                        t = after;
                    }
                    const outLines = [];
                    let nonEmptyCount = 0;
                    for (const line of t.split(/\\r?\\n/)) {
                        const s = line.trim();
                        if (!s) { if (outLines.length > 0 && outLines[outLines.length - 1] !== '') outLines.push(''); continue; }
                        nonEmptyCount++;
                        const sCompact = s.replace(/\\s+/g, '');
                        if (promptLineSet.has(s) || (sCompact.length >= 6 && promptCompact.includes(sCompact)) || promptHeaderRe.test(s) || s.includes('※出力は投稿するメッセージ本文のみ')) continue;
                        if (uiNoise.has(s.toLowerCase()) || uiNoise.has(s)) continue;
                        if (disclaimerRe.test(s)) continue;
                        if (s.length <= 36 && modelBadgeRe.test(s)) continue;
                        outLines.push(line.trimEnd());
                    }
                    const joined = outLines.join('\\n').trim();
                    if (!joined || (nonEmptyCount > 0 && outLines.length === 0)) return '';
                    const jCompact = joined.replace(/\\s+/g, '');
                    if (jCompact.length >= 6 && promptCompact.includes(jCompact)) return '';
                    return joined;
                }
                const SPECIFIC_SELS = [
                    '[data-message-author-role="assistant"] .markdown', '[data-message-author-role="assistant"] .prose',
                    '[data-content="ai-message"]', '.ds-markdown', '.response-content-markdown',
                    '.model-response-text', 'model-response .markdown', 'message-content', '.markdown-body', '.markdown',
                    '.prose', '.outputBox', '.bot-message', '.ai-response',
                    '[data-message-author-role="assistant"]', '[data-testid*="assistant"]', '[data-testid*="response"]',
                    '[data-testid="chat-message"]', '[data-testid*="message"]', '[class*="assistant"]', '[class*="response"]',
                    '[class*="answer"]', 'article', '[role="article"]'
                ];
                function getSearchRoots() {
                    const roots = [document];
                    const seen = new Set(roots);
                    for (let i = 0; i < roots.length; i++) {
                        try {
                            const r = roots[i];
                            const all = r.querySelectorAll('*');
                            for (let j = 0; j < all.length; j++) {
                                const sr = all[j].shadowRoot;
                                if (sr && !seen.has(sr)) {
                                    seen.add(sr);
                                    roots.push(sr);
                                }
                            }
                        } catch (e) {}
                    }
                    return roots;
                }
                const searchRoots = getSearchRoots();
                function queryAllDeep(root, sel, out) {
                    const rootsToScan = (root === document) ? searchRoots : [root];
                    for (const r of rootsToScan) {
                        try { r.querySelectorAll(sel).forEach(el => out.push(el)); } catch (e) {}
                    }
                }
                let bestText = '';
                for (const sel of SPECIFIC_SELS) {
                    const els = [];
                    queryAllDeep(document, sel, els);
                    const candList = [];
                    for (const el of els) {
                        const role = (el.getAttribute && (el.getAttribute('data-message-author-role') || '').toLowerCase()) || '';
                        const tid = (el.getAttribute && (el.getAttribute('data-testid') || '').toLowerCase()) || '';
                        if (role === 'user' || (tid.includes('user') && !tid.includes('assistant'))) continue;
                        if (el.querySelector && el.querySelector('textarea, input')) continue;
                        const raw = extractLiveText(el);
                        if (!raw) continue;
                        if (beforeSelCounts[raw] && beforeSelCounts[raw] > 0) continue;
                        const cleaned = cleanCandidate(raw);
                        if (cleaned.length >= 10) candList.push(cleaned);
                    }
                    if (candList.length > 0) { bestText = candList[candList.length - 1]; break; }
                }
                if (!bestText && args.allowBodyDiff) {
                    let combinedBody = '';
                    if (document.body) {
                        combinedBody = extractLiveText(document.body);
                    }
                    for (const r of searchRoots) {
                        if (r !== document) combinedBody += '\\n' + extractLiveText(r);
                    }
                    const curLines = combinedBody.split(/\\r?\\n/).map(x => x.trim()).filter(Boolean);
                    const usedBefore = Object.assign(Object.create(null), beforeBodyCounts);
                    const diffLines = [];
                    for (const ln of curLines) {
                        if (usedBefore[ln] && usedBefore[ln] > 0) { usedBefore[ln] -= 1; continue; }
                        diffLines.push(ln);
                    }
                    const diffClean = cleanCandidate(diffLines.join('\\n'));
                    if (diffClean.length >= 10) bestText = diffClean;
                }
                function isTrulyActiveStop(el) {
                    try {
                        if (el.disabled || el.getAttribute('aria-disabled') === 'true') return false;
                        if (el.closest && el.closest('[aria-hidden="true"], [hidden], [inert], nav, header, aside')) return false;
                        const cs = window.getComputedStyle(el);
                        if (!cs || cs.display === 'none' || cs.visibility === 'hidden' || parseFloat(cs.opacity || '1') <= 0.05) return false;
                        const r = el.getBoundingClientRect();
                        return !!(r && r.width > 2 && r.height > 2);
                    } catch (e) { return false; }
                }
                let isGenerating = false;
                const stopEls = [];
                queryAllDeep(document, 'button[aria-label*="Stop" i], button[aria-label*="停止"], button[aria-label*="中止"], button[title*="Stop" i], button[title*="停止"], button[data-testid="stop-button"], button[data-testid*="stop" i], [role="button"][data-testid*="stop" i], [data-streaming="true"], [data-is-streaming="true"], .typing-indicator', stopEls);
                for (const se of stopEls) {
                    if (isTrulyActiveStop(se)) { isGenerating = true; break; }
                }
                if (!isGenerating) {
                    const btns = [];
                    queryAllDeep(document, 'button', btns);
                    for (const b of btns) {
                        try {
                            const bt = (b.innerText || '').trim().toLowerCase();
                            if (bt === 'stop' || bt === 'stop generating' || bt === '停止' || bt === '中止' || bt === '生成を停止') {
                                if (isTrulyActiveStop(b)) { isGenerating = true; break; }
                            }
                        } catch (e) {}
                    }
                }
                function isUserMessageCopyBtn(b) {
                    try {
                        let p = b.parentElement;
                        for (let d = 0; d < 4 && p; d++, p = p.parentElement) {
                            const rAttr = (p.getAttribute && (p.getAttribute('data-message-author-role') || '')).toLowerCase();
                            const tAttr = (p.getAttribute && (p.getAttribute('data-testid') || '')).toLowerCase();
                            if (rAttr === 'user' || (tAttr.includes('user') && !tAttr.includes('assistant'))) return true;
                            if (p.querySelector && p.querySelector('button[aria-label*="Edit" i], button[aria-label*="編集"]')) {
                                const hasReplyMarker = p.querySelector('[data-message-author-role="assistant"], button[aria-label*="Bad" i], button[aria-label*="Thumbs" i], button[aria-label*="Good" i]');
                                if (!hasReplyMarker && !String(p.innerText || '').includes('他のAIモデル')) return true;
                            }
                        }
                    } catch (e) {}
                    return false;
                }
                const copyBtns = [];
                queryAllDeep(document, 'button[aria-label*="Copy" i], button[aria-label*="コピー"], button[aria-label*="クリップボード"], button[aria-label*="clipboard" i], button[title*="Copy" i], button[title*="コピー"], button[data-testid*="copy" i]', copyBtns);
                const visibleCopyBtns = copyBtns.filter(b => {
                    try {
                        if (b.disabled || b.getAttribute('aria-disabled') === 'true') return false;
                        const cs = window.getComputedStyle(b);
                        if (!cs || cs.display === 'none' || cs.visibility === 'hidden' || parseFloat(cs.opacity || '1') <= 0.05) return false;
                        const r = b.getBoundingClientRect();
                        if (!r || r.width <= 0 || r.height <= 0) return false;
                        if (isUserMessageCopyBtn(b)) return false;
                        return true;
                    } catch (e) { return false; }
                });
                const hasNewCopyBtn = visibleCopyBtns.length > (args.beforeCopyBtnCount || 0);
                return { text: bestText, isGenerating, hasNewCopyBtn, copyBtnCount: visibleCopyBtns.length };
            }""", {
                "beforeSelTexts": before_snap.get("selTexts", []) if isinstance(before_snap, dict) else [],
                "beforeBodyLines": before_snap.get("bodyLines", []) if isinstance(before_snap, dict) else [],
                "beforeCopyBtnCount": before_snap.get("copyBtnCount", 0) if isinstance(before_snap, dict) else 0,
                "promptText": prompt_text,
                "allowBodyDiff": allow_body_diff,
            })

        def _filter_prompt_echo(raw_copied, fallback_text):
            clean_fb = _strip_ai_ui_noise_lines(_strip_prompt_fragments(fallback_text or "", prompt_text=prompt_text))
            c_txt = _strip_ai_ui_noise_lines(_strip_prompt_fragments(raw_copied or "", prompt_text=prompt_text))
            if not c_txt or len(c_txt) < 10:
                return clean_fb
            return c_txt

        def _try_copy_via_button(fallback_text):
            clean_fb = _strip_ai_ui_noise_lines(_strip_prompt_fragments(fallback_text or "", prompt_text=prompt_text))
            try:
                copied_raw = target_ctx.evaluate("""async () => {
                    window.__aiAutoCapturedText = '';
                    const copySel = 'button[aria-label*="Copy" i], button[aria-label*="コピー"], button[aria-label*="クリップボード"], button[aria-label*="clipboard" i], button[title*="Copy" i], button[title*="コピー"], button[data-testid*="copy" i]';
                    function collectCopyBtns(root, out) {
                        try { root.querySelectorAll(copySel).forEach(el => out.push(el)); } catch (e) {}
                        try { root.querySelectorAll('*').forEach(el => { if (el.shadowRoot) collectCopyBtns(el.shadowRoot, out); }); } catch (e) {}
                    }
                    function isUserCopyBtn(b) {
                        try {
                            let p = b.parentElement;
                            for (let d = 0; d < 4 && p; d++, p = p.parentElement) {
                                const rAttr = (p.getAttribute && (p.getAttribute('data-message-author-role') || '')).toLowerCase();
                                const tAttr = (p.getAttribute && (p.getAttribute('data-testid') || '')).toLowerCase();
                                if (rAttr === 'user' || (tAttr.includes('user') && !tAttr.includes('assistant'))) return true;
                                if (p.querySelector && p.querySelector('button[aria-label*="Edit" i], button[aria-label*="編集"]')) {
                                    const hasReply = p.querySelector('[data-message-author-role="assistant"], button[aria-label*="Bad" i], button[aria-label*="Thumbs" i], button[aria-label*="Good" i]');
                                    if (!hasReply && !String(p.innerText || '').includes('他のAIモデル')) return true;
                                }
                            }
                        } catch (e) {}
                        return false;
                    }
                    const allBtns = [];
                    collectCopyBtns(document, allBtns);
                    const visBtns = allBtns.filter(b => {
                        try {
                            if (b.disabled || b.getAttribute('aria-disabled') === 'true') return false;
                            const cs = window.getComputedStyle(b);
                            if (!cs || cs.display === 'none' || cs.visibility === 'hidden' || parseFloat(cs.opacity || '1') <= 0.05) return false;
                            const r = b.getBoundingClientRect();
                            if (!r || r.width <= 0 || r.height <= 0) return false;
                            if (b.closest && b.closest('pre, code, [data-message-author-role="user"]')) return false;
                            if (isUserCopyBtn(b)) return false;
                            return true;
                        } catch (e) { return false; }
                    });
                    const targetBtn = visBtns.length > 0 ? visBtns[visBtns.length - 1] : null;
                    if (!targetBtn) return '';
                    targetBtn.scrollIntoView({ block: 'center' });
                    targetBtn.click();
                    await new Promise(res => setTimeout(res, 220));
                    if (window.__aiAutoCapturedText && window.__aiAutoCapturedText.trim()) {
                        return window.__aiAutoCapturedText.trim();
                    }
                    try {
                        if (navigator.clipboard && navigator.clipboard.readText) {
                            const clipTxt = await Promise.race([
                                navigator.clipboard.readText(),
                                new Promise(res => setTimeout(() => res(''), 500))
                            ]);
                            if (clipTxt && String(clipTxt).trim()) return String(clipTxt).trim();
                        }
                    } catch (e) {}
                    return '';
                }""")
                resolved = _filter_prompt_echo(copied_raw, clean_fb)
                if resolved and len(resolved) >= 10:
                    return resolved
            except Exception:
                pass
            return clean_fb

        last_txt = ""
        clean_stable_ticks = 0
        unchanged_ticks = 0
        eval_err_logged = False
        copy_probe_ticks = 0
        time.sleep(0.35)
        for tick in range(240):
            time.sleep(0.5)
            try:
                ext = _evaluate_ai_output(allow_body_diff=(tick >= 4))
            except Exception as e_eval:
                if not eval_err_logged and log_func:
                    log_func(f"[WARN] [AI専用Window:{engine}] 応答監視中に一時エラー: {e_eval}")
                    eval_err_logged = True
                ext = {"text": "", "isGenerating": False, "hasNewCopyBtn": False, "copyBtnCount": 0}
            cur_txt = _strip_ai_ui_noise_lines(str((ext or {}).get("text") or "").strip()) if isinstance(ext, dict) else ""
            is_gen = bool((ext or {}).get("isGenerating")) if isinstance(ext, dict) else False
            has_new_copy = bool((ext or {}).get("hasNewCopyBtn")) if isinstance(ext, dict) else False
            if (not cur_txt or len(cur_txt) < 10) and not is_gen and has_new_copy and tick >= 10:
                copy_probe_ticks += 1
                if copy_probe_ticks in (1, 6, 14):
                    probed = _try_copy_via_button("")
                    if probed and len(probed) >= 10:
                        if log_func:
                            log_func(f"[INFO] 🤖 [AI専用Window] {engine} のコピーボタンから本文を取得しました ({len(probed)}文字)")
                        return probed
            if cur_txt and len(cur_txt) >= 10:
                if cur_txt == last_txt:
                    unchanged_ticks += 1
                    clean_stable_ticks = 0 if is_gen else (clean_stable_ticks + 1)
                    ready_clean = (not is_gen) and (
                        (has_new_copy and clean_stable_ticks >= 2) or (clean_stable_ticks >= 4)
                    )
                    ready_cap = (has_new_copy and unchanged_ticks >= 5) or (unchanged_ticks >= 10)
                    if ready_clean or ready_cap:
                        time.sleep(0.35)
                        try:
                            ext2 = _evaluate_ai_output(allow_body_diff=(tick >= 4))
                        except Exception:
                            ext2 = ext
                        cur_txt2 = _strip_ai_ui_noise_lines(str((ext2 or {}).get("text") or "").strip()) if isinstance(ext2, dict) else cur_txt
                        is_gen2 = bool((ext2 or {}).get("isGenerating")) if isinstance(ext2, dict) else False
                        if cur_txt2 == cur_txt and (not is_gen2 or ready_cap):
                            final_txt = _try_copy_via_button(cur_txt2)
                            if final_txt:
                                if log_func:
                                    log_func(f"[INFO] 🤖 [AI専用Window] {engine} の出力完了を確認し本文のみを取得しました ({len(final_txt)}文字)")
                                return final_txt
                        else:
                            if cur_txt2:
                                last_txt = cur_txt2
                            clean_stable_ticks = 0
                            unchanged_ticks = 0
                else:
                    last_txt = cur_txt
                    clean_stable_ticks = 0
                    unchanged_ticks = 0
        if last_txt and len(last_txt) >= 10:
            final_txt = _try_copy_via_button(last_txt) or last_txt
            if log_func:
                log_func(f"[INFO] 🤖 [AI専用Window] {engine} の出力を取得しました ({len(final_txt)}文字)")
            return final_txt
        raise RuntimeError(f"{engine} (固定AIウィンドウ) タイムアウト")
    finally:
        try:
            browser.close()
        except Exception:
            pass


def _generate_with_single_engine(engine, prompt_text, ai_cfg, playwright_obj=None, base_dir=BASE_DIR, log_func=print):
    meta = AI_ENGINE_DEFAULTS.get(engine, {})
    model_name = str(ai_cfg.get("model_name") or "").strip() or meta.get("model", "")
    api_key, base_url = _parse_api_key_or_url(ai_cfg.get("api_key_or_url"), default_url=meta.get("api_url", ""))
    ai_win_sync = bool(ai_cfg.get("ai_window_sync", True))
    mode = meta.get("mode", "web")
    if engine == "chrome_builtin":
        return _generate_via_fixed_ai_window(playwright_obj, engine, prompt_text, base_dir=base_dir, log_func=log_func)
    if engine == "gemini":
        if api_key or (base_url and base_url != meta.get("api_url", "")):
            url = base_url or f"https://generativelanguage.googleapis.com/v1beta/models/{model_name or 'gemini-2.5-flash'}:generateContent?key={api_key}"
            data = _http_post_json(url, {"contents": [{"parts": [{"text": prompt_text}]}], "generationConfig": {"temperature": 0.9, "maxOutputTokens": 700}})
            parts = (((data.get("candidates") or [{}])[0].get("content") or {}).get("parts")) or []
            txt = "".join(str(pt.get("text") or "") for pt in parts).strip()
            if txt:
                return txt
            raise RuntimeError("Gemini API空応答")
        if ai_win_sync:
            return _generate_via_fixed_ai_window(playwright_obj, engine, prompt_text, base_dir=base_dir, log_func=log_func)
        raise RuntimeError("gemini: APIキー未設定")
    if engine == "claude":
        if api_key:
            data = _http_post_json(base_url or meta["api_url"], {"model": model_name, "max_tokens": 700, "messages": [{"role": "user", "content": prompt_text}]}, headers={"x-api-key": api_key, "anthropic-version": "2023-06-01"})
            txt = "".join(str(b.get("text") or "") for b in (data.get("content") or []) if isinstance(b, dict)).strip()
            if txt:
                return txt
            raise RuntimeError("Claude API空応答")
        if ai_win_sync:
            return _generate_via_fixed_ai_window(playwright_obj, engine, prompt_text, base_dir=base_dir, log_func=log_func)
        raise RuntimeError("claude: APIキー未設定")
    if engine == "ollama":
        url = base_url or meta["api_url"]
        data = _http_post_json(url, {"model": model_name or "llama3.2", "messages": [{"role": "user", "content": prompt_text}], "stream": False}, timeout=60)
        txt = str(((data.get("message") or {}).get("content")) or data.get("response") or "").strip()
        if txt:
            return txt
        raise RuntimeError("Ollama空応答")
    if mode in ("openai", "openai_nokey", "web_or_openai"):
        is_custom = bool(base_url and base_url != meta.get("api_url", ""))
        if api_key or mode == "openai_nokey" or is_custom:
            url = base_url or meta.get("api_url")
            hdrs = {"Authorization": f"Bearer {api_key}"} if api_key else {}
            data = _http_post_json(url, {"model": model_name or "gpt-4o-mini", "messages": [{"role": "user", "content": prompt_text}], "temperature": 0.9}, headers=hdrs, timeout=45)
            txt = ((((data.get("choices") or [{}])[0]).get("message") or {}).get("content") or "").strip()
            if txt:
                return txt
            raise RuntimeError(f"{engine} API空応答")
    if ai_win_sync:
        return _generate_via_fixed_ai_window(playwright_obj, engine, prompt_text, base_dir=base_dir, log_func=log_func)
    raise RuntimeError(f"{engine}: API未設定かつAI専用ウィンドウ連携=0")


def generate_ai_post_text(account_name, ai_cfg, prompt_row=None, playwright_obj=None, base_dir=BASE_DIR, log_func=print):
    prompt_text = _build_ai_generation_prompt(account_name, ai_cfg, prompt_row=prompt_row)
    primaries = _parse_engine_list(ai_cfg.get("engine", "duckai"), default_fallback="duckai")
    raw_fallback = str(ai_cfg.get("fallback_candidates") or "").strip()
    chain = _resolve_ai_engine_chain(ai_cfg.get("engine", "duckai"), raw_fallback)
    model_map = _resolve_engine_models_for_chain(ai_cfg.get("model_name", ""), primaries, chain)
    cred_map = _resolve_engine_credentials_for_chain(ai_cfg.get("api_key_or_url", ""), primaries, chain)

    if not raw_fallback:
        login_free_webs = {"duckai", "talkai", "deepai", "venice"}
        filtered_chain = []
        for e in chain:
            has_cred = any(str(c or "").strip() for c in (cred_map.get(e) or []))
            if e in primaries or e in login_free_webs or has_cred:
                filtered_chain.append(e)
        chain = filtered_chain or primaries

    errors = []
    for eng_idx, eng in enumerate(chain):
        eng_models = model_map.get(eng) or [""]
        eng_creds = cred_map.get(eng) or [""]
        for mod_val in eng_models:
            for cred_val in eng_creds:
                try:
                    local_cfg = dict(ai_cfg)
                    local_cfg["model_name"] = mod_val
                    local_cfg["api_key_or_url"] = cred_val
                    parsed_key, parsed_url = _parse_api_key_or_url(cred_val, default_url=AI_ENGINE_DEFAULTS.get(eng, {}).get("api_url", ""))
                    is_custom_url = bool(parsed_url and parsed_url != AI_ENGINE_DEFAULTS.get(eng, {}).get("api_url", ""))
                    eng_mode = AI_ENGINE_DEFAULTS.get(eng, {}).get("mode", "")
                    if eng_mode == "openai" and not parsed_key and not is_custom_url:
                        continue
                    raw_out = _generate_with_single_engine(eng, prompt_text, local_cfg, playwright_obj=playwright_obj, base_dir=base_dir, log_func=log_func)
                    cleaned = _clean_ai_generated_post_text(raw_out, prompt_text=prompt_text)
                    if cleaned:
                        if eng_idx > 0 and log_func:
                            log_func(f"[INFO] [{account_name}] 🔄 自動切替エンジン ({eng}) で投稿文を生成しました")
                        return cleaned, eng, None
                    errors.append(f"{eng}: 空の応答")
                except Exception as e_eng:
                    errors.append(f"{eng}: {e_eng}")
                    if log_func:
                        next_hint = f" ➔ 次({chain[eng_idx + 1]})へ自動切替" if eng_idx + 1 < len(chain) else ""
                        log_func(f"[WARN] [{account_name}] AI生成試行 ({eng}) 失敗: {e_eng}{next_hint}")
    return "", chain[0] if chain else "duckai", " / ".join(errors[:4])


def read_ai_post_config_from_xlsx(filepath):
    default_sections = {k: d for k, _lbl, d in AI_PROMPT_SECTION_META}
    cfg = {
        "exists": False,
        "sheet_name": "AIポスト",
        "time_slot": "0:00-0:00",
        "weekdays": "毎日",
        "interval_str": "120〜360分",
        "min_interval_sec": 7200,
        "max_interval_sec": 21600,
        "enabled": False,
        "memo": "",
        "engine": "duckai",
        "model_name": "",
        "fallback_candidates": "",
        "ai_window_sync": True,
        "api_key_or_url": "",
        "prompt_sections": dict(default_sections),
        "rows": [],
    }
    if not filepath or not os.path.exists(filepath):
        return cfg
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = next((wb[s] for s in wb.sheetnames if _is_ai_post_sheet_name(s)), None)
        if ws is None:
            wb.close()
            return cfg
        cfg["exists"] = True
        cfg["sheet_name"] = ws.title
        all_rows = list(ws.iter_rows(values_only=True))

        def _apply_pair(h_row, v_row):
            hdr_list = [str(c or "").strip() for c in (h_row or [])[:10]]
            for idx, lbl in enumerate(hdr_list):
                raw_v = v_row[idx] if v_row and idx < len(v_row) else ""
                val = raw_v.strftime("%H:%M") if hasattr(raw_v, "strftime") else ("" if raw_v is None else str(raw_v).strip())
                if not lbl:
                    continue
                if "実行時間" in lbl or "時間帯" in lbl:
                    if val:
                        cfg["time_slot"] = val
                elif "曜日" in lbl:
                    if val:
                        cfg["weekdays"] = val
                elif "間隔" in lbl:
                    mn, mx, disp = _parse_ai_interval_range(val)
                    if mn is not None:
                        cfg["min_interval_sec"] = mn
                        cfg["max_interval_sec"] = mx
                        cfg["interval_str"] = disp
                elif "有効" in lbl or "状態" in lbl:
                    cfg["enabled"] = bool(val) and (val not in ["無効", "停止", "OFF", "off", "0", "false"])
                elif "メモ" in lbl:
                    cfg["memo"] = val
                elif "エンジン" in lbl:
                    if val:
                        cfg["engine"] = _normalize_engine_list_str(val, default_fallback="duckai")
                elif "モデル" in lbl:
                    cfg["model_name"] = val
                elif "フォールバック" in lbl:
                    cfg["fallback_candidates"] = val
                elif "ウィンドウ" in lbl or "window" in lbl.lower():
                    cfg["ai_window_sync"] = val.lower() not in ["0", "off", "false", "無効", "なし"]
                elif "API" in lbl.upper() or "URL" in lbl.upper() or "キー" in lbl:
                    cfg["api_key_or_url"] = val

        if len(all_rows) >= 2:
            _apply_pair(all_rows[0], all_rows[1])
        row3_str = " ".join(str(c or "") for c in (all_rows[2] or [])) if len(all_rows) >= 4 else ""
        if len(all_rows) >= 4 and ("AIエンジン" in row3_str or "AIモデル" in row3_str):
            _apply_pair(all_rows[2], all_rows[3])

        label_to_key = {s_lbl: s_key for s_key, s_lbl, _d in AI_PROMPT_SECTION_META}
        for r_cand in all_rows[4:20]:
            if not r_cand:
                continue
            s_key = label_to_key.get(str(r_cand[0] or "").strip())
            if s_key:
                cfg["prompt_sections"][s_key] = "" if len(r_cand) < 2 or r_cand[1] is None else str(r_cand[1]).strip()

        tbl_hdr_idx = next(
            (r_i for r_i, r in enumerate(all_rows[:25], start=1) if r and ("生成されたポスト" in " ".join(str(c or "") for c in r) or "AI指示プロンプト・テーマ" in " ".join(str(c or "") for c in r))),
            None,
        )
        if tbl_hdr_idx and tbl_hdr_idx < len(all_rows):
            th = [str(c or "").strip() for c in all_rows[tbl_hdr_idx - 1]]
            col_id = next((i for i, h in enumerate(th) if h.upper() == "ID"), 0)
            col_pr = next((i for i, h in enumerate(th) if "プロンプト" in h or "指示" in h), -1)
            col_tp = next((i for i, h in enumerate(th) if "トピック" in h), -1)
            col_dl = next((i for i, h in enumerate(th) if "削除時間" in h), -1)
            col_gn = next((i for i, h in enumerate(th) if "生成された" in h), 1)
            col_st = next((i for i, h in enumerate(th) if "ステータス" in h), 3)
            for r_num, row in enumerate(all_rows[tbl_hdr_idx:], start=tbl_hdr_idx + 1):
                pr_v = str(row[col_pr] or "").strip() if row and col_pr != -1 and col_pr < len(row) else ""
                gn_v = str(row[col_gn] or "").strip() if row and col_gn != -1 and col_gn < len(row) else ""
                if not pr_v and not gn_v:
                    continue
                cfg["rows"].append({
                    "_row_idx": r_num,
                    "id": str(row[col_id]).strip() if col_id < len(row) and row[col_id] is not None else str(r_num - tbl_hdr_idx),
                    "prompt": pr_v,
                    "generated_text": gn_v,
                    "topic": str(row[col_tp] or "").strip() if col_tp != -1 and col_tp < len(row) else "",
                    "delete_after": str(row[col_dl] or "").strip() if col_dl != -1 and col_dl < len(row) else "",
                    "status": (str(row[col_st] or "").strip() if col_st != -1 and col_st < len(row) else "") or "指示待機",
                })
        wb.close()
    except Exception:
        pass
    return cfg


def record_ai_post_result_in_xlsx(filepath, row_idx=None, prompt_text="", topic="", delete_after="", generated_text="", engine_used="", status="投稿完了", finish_time=None, error_msg=""):
    if not filepath or not os.path.exists(filepath):
        return None
    try:
        wb = openpyxl.load_workbook(filepath)
        ws = next((wb[s] for s in wb.sheetnames if _is_ai_post_sheet_name(s)), wb.create_sheet("AIポスト"))
        r3_txt = " ".join(str(ws.cell(row=3, column=c).value or "") for c in range(1, 6))
        r7_txt = " ".join(str(ws.cell(row=7, column=c).value or "") for c in range(1, 6))
        is_new = ("AIエンジン" in r3_txt) or ("プロンプト項目" in r7_txt)
        if is_new:
            last_used = 17
            max_id = 0
            for r in range(18, ws.max_row + 1):
                if str(ws.cell(row=r, column=2).value or "").strip() or str(ws.cell(row=r, column=4).value or "").strip():
                    last_used = r
                    try:
                        max_id = max(max_id, int(str(ws.cell(row=r, column=1).value).strip()))
                    except Exception:
                        pass
            target_row = row_idx if (row_idx and row_idx >= 18) else (last_used + 1)
            if not row_idx or row_idx < 18:
                ws.cell(row=target_row, column=1, value=max_id + 1)
            ws.cell(row=target_row, column=2, value=generated_text or prompt_text)
            ws.cell(row=target_row, column=3, value=engine_used)
            ws.cell(row=target_row, column=4, value=status)
            meta_s = " / ".join(x for x in [finish_time or "", (f"削除:{delete_after}" if delete_after else ""), (error_msg or "")] if x)
            ws.cell(row=target_row, column=5, value=meta_s)
        else:
            target_row = row_idx
            if not target_row or target_row < 5:
                last_used = 4
                max_id = 0
                for r in range(5, ws.max_row + 1):
                    if str(ws.cell(row=r, column=2).value or "").strip() or str(ws.cell(row=r, column=5).value or "").strip():
                        last_used = r
                        try:
                            max_id = max(max_id, int(str(ws.cell(row=r, column=1).value).strip()))
                        except Exception:
                            pass
                target_row = last_used + 1
                ws.cell(row=target_row, column=1, value=max_id + 1)
                ws.cell(row=target_row, column=2, value=prompt_text)
                ws.cell(row=target_row, column=3, value=topic)
                ws.cell(row=target_row, column=4, value=delete_after)
            ws.cell(row=target_row, column=5, value=generated_text)
            ws.cell(row=target_row, column=6, value=engine_used)
            ws.cell(row=target_row, column=7, value=status)
            if finish_time:
                ws.cell(row=target_row, column=8, value=finish_time)
            ws.cell(row=target_row, column=9, value=error_msg or "")
        wb.save(filepath)
        wb.close()
        return target_row
    except Exception:
        return None

def _find_post_header_row_in_sheet_cli(rows):
    post_keys = ["ポスト文章", "投稿本文", "ツイート文章", "posttext", "投稿文章", "ポスト本文", "ポスト内容", "投稿内容", "本文"]
    aux_keys = ["id", "ステータス", "status", "状態", "画像", "削除時間", "トピック", "投稿時間", "投稿完了日時"]
    fallback_row = None
    for r_idx, row in enumerate((rows or [])[:15], start=1):
        if not row:
            continue
        row_cells_norm = [re.sub(r'[\s　_\-:=]+', '', str(c).strip().lower()) for c in row if c is not None]
        has_post_hdr = any(len(cell) <= 18 and any(k in cell for k in post_keys) for cell in row_cells_norm)
        if not has_post_hdr:
            continue
        has_aux_hdr = any(len(cell) <= 18 and any(k in cell for k in aux_keys) for cell in row_cells_norm)
        if has_aux_hdr:
            return r_idx
        if fallback_row is None:
            fallback_row = r_idx
    return fallback_row


def _is_system_sheet_name_cli(sname, ws=None):
    s = str(sname or "").strip()
    sl = s.lower()
    if not s or _is_ai_post_sheet_name(s):
        return True
    if any(k in sl for k in ["日常ポスト", "日常挨拶", "日常あいさつ", "日常枠", "アカウント情報", "accountinfo", "account_info"]):
        return True
    exact_sys = {"設定", "基本設定", "テーマ設定", "アカウント設定", "config", "setting", "settings", "themesetting", "themeconfig"}
    if sl in exact_sys or "設計" in s or any(k in sl for k in ["設定", "config", "setting"]):
        if ws is not None:
            try:
                rows = list(ws.iter_rows(max_row=15, values_only=True))
                if _find_post_header_row_in_sheet_cli(rows) is not None:
                    return False
            except Exception:
                pass
        return True
    return False


def _read_sheet_theme_meta_cli(rows, header_row_idx, sheet_name):
    meta = {
        "name": str(sheet_name or "").strip(),
        "sheet_name": str(sheet_name or "").strip(),
        "time_slot": "0:00-0:00",
        "weekdays": "毎日",
        "enabled": True,
    }
    if not rows or not header_row_idx or header_row_idx <= 1:
        return meta
    meta_rows = rows[:header_row_idx - 1]
    label_keywords = ["実行時間", "時間帯", "投稿時間", "配信時間", "時間枠", "曜日", "有効", "状態"]
    def _is_lbl(cv):
        nl = re.sub(r'[\s　_\-:=()（）\[\]【】/]+', '', str(cv or "").strip().lower())
        return bool(nl) and len(nl) <= 20 and any(k in nl for k in label_keywords)
    def _fmt(cv):
        if cv is None:
            return ""
        return cv.strftime("%H:%M") if hasattr(cv, "strftime") else str(cv).strip()
    for r_i, row in enumerate(meta_rows):
        if not row:
            continue
        for c_i, cell in enumerate(row):
            nl = re.sub(r'[\s　_\-:=()（）\[\]【】/]+', '', str(cell or "").strip().lower())
            if not nl or len(nl) > 24:
                continue
            below_v = _fmt(meta_rows[r_i + 1][c_i]) if (r_i + 1 < len(meta_rows) and meta_rows[r_i + 1] and c_i < len(meta_rows[r_i + 1]) and not _is_lbl(meta_rows[r_i + 1][c_i])) else ""
            right_v = _fmt(row[c_i + 1]) if (c_i + 1 < len(row) and not _is_lbl(row[c_i + 1])) else ""
            val = below_v or right_v
            if any(k in nl for k in ["実行時間", "時間帯", "投稿時間", "配信時間", "時間枠"]) and not any(ex in nl for ex in ["完了", "最終"]):
                if val and not val.startswith(("(", "（", "※")):
                    meta["time_slot"] = val
            elif "曜日" in nl:
                if val and not val.startswith(("(", "（", "※")):
                    meta["weekdays"] = val
            elif any(k in nl for k in ["有効", "状態", "status"]):
                if val and not val.startswith(("(", "（", "※")):
                    val_n = unicodedata.normalize("NFKC", val).strip().lower()
                    meta["enabled"] = not any(off_k in val_n for off_k in ["無効", "停止", "off", "false"]) and val_n != "0"
    return meta


def read_excel_records(filepath):
    if not filepath or not os.path.exists(filepath):
        return []

    wb = None
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
        sheets_to_parse = []

        if is_account_file:
            target_ws = None
            target_header_row_idx = 1
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
            if not target_ws:
                target_ws = wb.active
                target_header_row_idx = 1
            sheets_to_parse.append((target_ws.title, list(target_ws.iter_rows(values_only=True)), target_header_row_idx, None))
        else:
            theme_sheets = [wb[s] for s in wb.sheetnames if not _is_system_sheet_name_cli(s, wb[s])]
            if not theme_sheets:
                non_ai = [wb[s] for s in wb.sheetnames if not _is_ai_post_sheet_name(s) and not any(k in s.lower() for k in ["設定", "config", "setting"])]
                theme_sheets = non_ai[:1] if non_ai else [wb.active]

            for ws in theme_sheets:
                if _is_ai_post_sheet_name(ws.title):
                    continue
                all_rows = list(ws.iter_rows(values_only=True))
                hdr_idx = _find_post_header_row_in_sheet_cli(all_rows)
                if hdr_idx is None:
                    continue
                s_meta = _read_sheet_theme_meta_cli(all_rows, hdr_idx, ws.title)
                sheets_to_parse.append((ws.title, all_rows, hdr_idx, s_meta))

        wb.close()

        records = []
        for sheet_name, rows, target_header_row_idx, sheet_meta in sheets_to_parse:
            if not rows or len(rows) < target_header_row_idx:
                continue

            header_row = rows[target_header_row_idx - 1]
            headers = [str(h).strip() if h is not None else f"col_{i}" for i, h in enumerate(header_row)]

            for r_idx, row in enumerate(rows[target_header_row_idx:], start=target_header_row_idx + 1):
                if not row or all(c is None or str(c).strip() == "" for c in row):
                    continue
                r_dict = {"_row_idx": r_idx, "_sheet_name": sheet_name}
                if sheet_meta is not None:
                    r_dict["_sheet_theme_meta"] = sheet_meta
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
                    elif any(k in norm_h for k in ["投稿時間", "投稿時刻", "投稿時間帯", "実行時間", "時間帯", "配信時間", "timeslot", "時間枠"]) and not any(ex in norm_h for ex in ["完了", "最終", "posted", "削除"]):
                        r_dict["投稿時間"] = val_clean
                        r_dict["時間帯"] = val_clean
                        r_dict["実行時間"] = val_clean
                    elif any(k in norm_h for k in ["曜日", "対象曜日", "weekday", "weekdays"]):
                        r_dict["曜日"] = val_clean
                    elif any(k in norm_h for k in ["削除時間", "自動削除", "削除タイマー", "削除まで", "deleteafter"]) and not any(ex in norm_h for ex in ["完了", "済", "日時"]):
                        r_dict["削除時間"] = val_clean
                        r_dict["delete_after"] = val_clean
                    elif any(k in norm_h for k in ["プロキシ", "proxy"]):
                        if val_clean or not r_dict.get("プロキシ"):
                            r_dict["プロキシ"] = val_clean
                            r_dict["proxy"] = val_clean
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
            if _is_ai_post_sheet_name(sname):
                continue
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

def mark_post_deleted_in_excel(filepath, row_idx, deleted_time_str=None, sheet_name=None):
    if not filepath or not os.path.exists(filepath) or not row_idx:
        return False
    try:
        wb = openpyxl.load_workbook(filepath)
        ws = None
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            for sname in wb.sheetnames:
                c_ws = wb[sname]
                for r in c_ws.iter_rows(max_row=15, values_only=True):
                    if r and any(str(c).strip() in ["ステータス", "status", "状態"] for c in r if c is not None):
                        ws = c_ws
                        break
                if ws:
                    break
            if not ws:
                ws = wb.active

        status_col, time_col = None, None
        for r_num, row_cells in enumerate(ws.iter_rows(max_row=15), start=1):
            vals = [str(cell.value).strip() if cell.value is not None else "" for cell in row_cells]
            for cand in ["ステータス", "status", "状態"]:
                if cand in vals:
                    status_col = vals.index(cand) + 1
                    break
            for cand in ["投稿完了日時", "完了日時", "日時", "posted_at"]:
                if cand in vals:
                    time_col = vals.index(cand) + 1
                    break
            if status_col:
                break

        del_ts = deleted_time_str or datetime.now().strftime("%H:%M:%S")
        if status_col:
            cur_st = str(ws.cell(row=row_idx, column=status_col).value or "").strip()
            if cur_st not in ("未投稿", "指示待機"):
                ws.cell(row=row_idx, column=status_col, value="投稿完了(削除済)")
        if time_col:
            cur_t = str(ws.cell(row=row_idx, column=time_col).value or "").strip()
            if "削除済" not in cur_t:
                ws.cell(row=row_idx, column=time_col, value=f"{cur_t} (削除済 {del_ts})".strip())
        wb.save(filepath)
        wb.close()
        return True
    except Exception as e:
        print(f"[WARN] Excel削除ステータス更新エラー: {e}")
        return False

_PENDING_DELETES_LOCK = threading.Lock()

def parse_delete_delay_spec(raw_spec):
    if raw_spec is None:
        return None, None, ""
    s = unicodedata.normalize("NFKC", str(raw_spec)).strip()
    if not s or s.startswith(("(", "（", "※")):
        return None, None, ""
    low_compact = s.lower().replace(" ", "").replace("　", "")
    if low_compact in ["nan", "none", "null", "0", "なし", "無し", "削除しない", "残す", "off", "no", "false", "-", "ー"]:
        return None, None, ""

    def _detect_unit(tok_str, default_mult=60):
        t_low = tok_str.lower()
        if any(u in tok_str for u in ["時間"]) or re.search(r'\d\s*(?:hr|hrs|hour|hours|h)\b', t_low):
            return 3600, "時間"
        if any(u in tok_str for u in ["秒"]) or re.search(r'\d\s*(?:sec|secs|second|seconds|s)\b', t_low):
            return 1, "秒"
        if any(u in tok_str for u in ["分"]) or re.search(r'\d\s*(?:min|mins|minute|minutes|m)\b', t_low):
            return 60, "分"
        return default_mult, ("時間" if default_mult == 3600 else ("秒" if default_mult == 1 else "分"))

    normalized = re.sub(r'\s*(?:[〜～~ー−―‐–—\-]|から|\bto\b)\s*', '-', s, flags=re.IGNORECASE)
    if "-" in normalized:
        p1, p2 = normalized.split("-", 1)
        m1_all = re.findall(r'\d+(?:\.\d+)?', p1)
        m2_all = re.findall(r'\d+(?:\.\d+)?', p2)
        if m1_all and m2_all:
            mult2, unit2 = _detect_unit(p2, default_mult=60)
            mult1, unit1 = _detect_unit(p1, default_mult=mult2)
            v1 = float(m1_all[0]) * mult1
            v2 = float(m2_all[0]) * mult2
            min_s = max(1, int(round(min(v1, v2))))
            max_s = max(min_s, int(round(max(v1, v2))))
            n1_disp = m1_all[0].rstrip('0').rstrip('.') if '.' in m1_all[0] else m1_all[0]
            n2_disp = m2_all[0].rstrip('0').rstrip('.') if '.' in m2_all[0] else m2_all[0]
            disp = f"{n1_disp}〜{n2_disp}{unit2}" if unit1 == unit2 else f"{n1_disp}{unit1}〜{n2_disp}{unit2}"
            return min_s, max_s, disp

    m_single = re.findall(r'\d+(?:\.\d+)?', s)
    if m_single:
        mult, unit = _detect_unit(s, default_mult=60)
        val_f = float(m_single[0])
        if val_f <= 0:
            return None, None, ""
        sec_val = max(1, int(round(val_f * mult)))
        n_disp = m_single[0].rstrip('0').rstrip('.') if '.' in m_single[0] else m_single[0]
        return sec_val, sec_val, f"{n_disp}{unit}"

    return None, None, ""

def _pending_deletes_path(profile_dir):
    return os.path.join(profile_dir, "pending_deletes.json")

def load_pending_deletes(profile_dir):
    if not profile_dir:
        return []
    path = _pending_deletes_path(profile_dir)
    with _PENDING_DELETES_LOCK:
        if not os.path.exists(path):
            return []
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, list) else []
        except Exception:
            return []

def save_pending_deletes(profile_dir, jobs):
    if not profile_dir:
        return False
    os.makedirs(profile_dir, exist_ok=True)
    path = _pending_deletes_path(profile_dir)
    with _PENDING_DELETES_LOCK:
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(jobs or [], f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False

def enqueue_pending_delete_job(profile_dir, account_name, posts_file, row_idx, sheet_name, post_id, delete_spec, post_url="", post_code="", profile_handle="", text_content=""):
    min_s, max_s, disp = parse_delete_delay_spec(delete_spec)
    if min_s is None or max_s is None:
        return None
    delay_sec = random.randint(min_s, max_s)
    now_ts = time.time()
    job = {
        "job_id": f"{int(now_ts * 1000)}_{row_idx}_{random.randint(100, 999)}",
        "account_name": str(account_name or "").strip(),
        "profile_dir": profile_dir,
        "posts_file": posts_file or "",
        "row_idx": int(row_idx) if row_idx else 0,
        "sheet_name": str(sheet_name or ""),
        "post_id": str(post_id or ""),
        "text_content": str(text_content or ""),
        "post_url": str(post_url or ""),
        "post_code": str(post_code or ""),
        "profile_handle": str(profile_handle or account_name or "").strip().lstrip("@"),
        "delete_spec": disp or str(delete_spec),
        "delay_sec": delay_sec,
        "posted_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "delete_at_ts": now_ts + delay_sec,
        "attempts": 0,
    }
    jobs = load_pending_deletes(profile_dir)
    jobs.append(job)
    save_pending_deletes(profile_dir, jobs)
    return job

def get_due_pending_deletes(profile_dir, now_ts=None):
    cur = time.time() if now_ts is None else now_ts
    return [j for j in load_pending_deletes(profile_dir) if float(j.get("delete_at_ts") or 0) <= cur]

def remove_pending_delete_job(profile_dir, job_id):
    jobs = load_pending_deletes(profile_dir)
    save_pending_deletes(profile_dir, [j for j in jobs if j.get("job_id") != job_id])

def update_pending_delete_job(profile_dir, updated_job):
    jobs = load_pending_deletes(profile_dir)
    save_pending_deletes(profile_dir, [updated_job if j.get("job_id") == updated_job.get("job_id") else j for j in jobs])

def _format_duration_jp(seconds):
    s = max(0, int(round(seconds)))
    h, m, rem_s = s // 3600, (s % 3600) // 60, s % 60
    if h > 0:
        return f"{h}時間{m}分" if m > 0 else f"{h}時間"
    if m > 0:
        return f"{m}分{rem_s}秒" if rem_s > 0 else f"{m}分"
    return f"{rem_s}秒"

def _extract_verification_snippets_cli(text_content):
    snippets = []
    if not text_content:
        return snippets
    for line in str(text_content).split("\n"):
        cleaned = unicodedata.normalize("NFKC", line).strip()
        cleaned = re.sub(r'[\s\u200b-\u200d\ufeff]+', '', cleaned)
        if not cleaned:
            continue
        if cleaned.startswith(("http://", "https://")):
            body_url = re.sub(r'^https?://', '', cleaned)
            if len(body_url) >= 8:
                snippets.append(body_url[:24])
            continue
        snippets.append(cleaned[:22])
    if not snippets:
        compact = re.sub(r'[\s\u200b-\u200d\ufeff]+', '', unicodedata.normalize("NFKC", str(text_content)))
        if compact:
            snippets.append(compact[:20])
    seen = set()
    uniq = []
    for sn in snippets:
        if len(sn) >= 4 and sn not in seen:
            seen.add(sn)
            uniq.append(sn)
    return uniq[:5]

def _extract_delete_after_cli(post_item):
    if not isinstance(post_item, dict):
        return ""
    for k in ["削除時間", "自動削除時間", "自動削除", "削除タイマー", "削除まで", "delete_after", "deleteAfter"]:
        v = post_item.get(k)
        if v is not None and str(v).strip() and str(v).strip().lower() not in ["none", "nan"]:
            return str(v).strip()
    return ""

def delete_threads_post_on_browser(page, context, account_name, job, log_func=print):
    handle = str(job.get("profile_handle") or account_name or "").strip().lstrip("@")
    post_url = str(job.get("post_url") or "").strip()
    post_code = str(job.get("post_code") or "").strip()
    snippets = _extract_verification_snippets_cli(job.get("text_content") or "")

    target_urls = []
    if post_url:
        u_norm = post_url.replace("https://www.threads.net", "https://www.threads.com")
        if u_norm.startswith("/"):
            u_norm = "https://www.threads.com" + u_norm
        target_urls.append(u_norm)
    elif post_code and handle:
        target_urls.append(f"https://www.threads.com/@{handle}/post/{post_code}")
    if handle:
        prof_u = f"https://www.threads.com/@{handle}"
        if prof_u not in target_urls:
            target_urls.append(prof_u)

    if not target_urls:
        return False, "削除対象URL/アカウント名が特定できません"

    for t_url in target_urls:
        is_permalink = "/post/" in t_url
        log_func(f"[{account_name}] 🗑️ 削除対象ページを開きます: {t_url}")
        try:
            page.goto(t_url, wait_until="domcontentloaded", timeout=30000)
            time.sleep(2.5)
        except Exception:
            continue

        cur_u = (page.url or "").lower()
        if "/login" in cur_u or "accounts/login" in cur_u:
            return False, "login_required"
        is_cp, cp_why = detect_phone_checkpoint_screen(page)
        if is_cp:
            return False, f"checkpoint:{cp_why}"

        try:
            body_txt_check = page.inner_text("body")[:2000]
        except Exception:
            body_txt_check = ""
        if is_permalink and any(nf in body_txt_check for nf in [
            "このページはご利用いただけません", "ページが見つかりません",
            "Sorry, this page isn't available", "This page isn't available"
        ]):
            return True, "already_deleted"

        more_clicked = False
        for scroll_try in range(1 if is_permalink else 6):
            try:
                found_box = page.evaluate("""({ snippets, isPermalink }) => {
                    const norm = (t) => (t || '').normalize('NFKC').replace(/[\\s\\u200b-\\u200d\\ufeff]+/g, '');
                    const root = document.querySelector('main') || document.querySelector('div[role="main"]') || document.body;
                    const matchMoreAria = (txt) => {
                        const t = (txt || '').trim();
                        if (!t) return false;
                        return ['もっと見る', 'More', 'オプション', 'Options', 'メニュー', 'Menu'].some(k => t === k || t.includes(k));
                    };
                    const btns = Array.from(root.querySelectorAll('div[role="button"], button')).filter(el => {
                        if (el.closest('nav')) return false;
                        const r = el.getBoundingClientRect();
                        if (r.width <= 0 || r.height <= 0 || r.width > 120 || r.height > 120) return false;
                        if (matchMoreAria(el.getAttribute('aria-label'))) return true;
                        const svg = el.querySelector('svg[aria-label], svg title');
                        if (svg) {
                            const sLabel = svg.getAttribute('aria-label') || svg.textContent || '';
                            if (matchMoreAria(sLabel)) return true;
                        }
                        const circles = el.querySelectorAll('svg circle');
                        if (circles.length === 3 && r.width <= 48 && r.height <= 48) return true;
                        return false;
                    });
                    if (!btns.length) return null;
                    let targetBtn = null;
                    if (snippets && snippets.length > 0) {
                        let bestBtn = null, smallestLen = 9999999;
                        for (const b of btns) {
                            let cur = b.parentElement;
                            for (let depth = 0; depth < 8 && cur && cur !== root; depth++) {
                                const cNorm = norm(cur.innerText || '');
                                if (snippets.some(sn => sn && cNorm.includes(sn))) {
                                    if (cNorm.length < smallestLen) { smallestLen = cNorm.length; bestBtn = b; }
                                    break;
                                }
                                cur = cur.parentElement;
                            }
                        }
                        targetBtn = bestBtn;
                    }
                    if (!targetBtn && isPermalink) targetBtn = btns[0];
                    if (!targetBtn) return null;
                    targetBtn.scrollIntoView({ block: 'center' });
                    const rect = targetBtn.getBoundingClientRect();
                    return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
                }""", {"snippets": snippets, "isPermalink": is_permalink})
                if found_box:
                    page.mouse.click(found_box["x"], found_box["y"])
                    time.sleep(1.2)
                    more_clicked = True
                    break
            except Exception:
                pass
            if not is_permalink and scroll_try < 5:
                try:
                    page.evaluate("window.scrollBy(0, 650)")
                    time.sleep(1.2)
                except Exception:
                    pass

        if not more_clicked:
            continue

        del_menu_clicked = False
        for d_sel in [
            'div[role="menuitem"]:has-text("削除")',
            'div[role="menu"] div[role="button"]:has-text("削除")',
            'div[role="menu"] span:text-is("削除")',
            'button:text-is("削除")', 'div[role="button"]:has-text("削除")', 'span:text-is("削除")',
            'div[role="menuitem"]:has-text("Delete")', 'div[role="menu"] span:text-is("Delete")', 'button:text-is("Delete")', 'span:text-is("Delete")',
        ]:
            try:
                loc = page.locator(d_sel).last
                if loc.is_visible(timeout=900) and len((loc.inner_text() or "").strip()) <= 20:
                    loc.click(force=True)
                    del_menu_clicked = True
                    time.sleep(1.2)
                    break
            except Exception:
                pass

        if not del_menu_clicked:
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            continue

        confirm_clicked = False
        for c_sel in [
            'div[role="dialog"] button:has-text("削除")', 'div[role="dialog"] div[role="button"]:has-text("削除")',
            'div[aria-modal="true"] button:has-text("削除")', 'div[aria-modal="true"] div[role="button"]:has-text("削除")',
            'div[role="dialog"] span:text-is("削除")', 'div[role="dialog"] button:has-text("Delete")',
            'div[role="dialog"] div[role="button"]:has-text("Delete")', 'div[aria-modal="true"] span:text-is("Delete")',
            'button:text-is("削除")', 'button:text-is("Delete")',
        ]:
            try:
                loc_c = page.locator(c_sel).last
                if loc_c.is_visible(timeout=2000) and len((loc_c.inner_text() or "").strip()) <= 20:
                    loc_c.click(force=True)
                    confirm_clicked = True
                    break
            except Exception:
                pass

        if not confirm_clicked:
            continue

        for _ in range(12):
            time.sleep(0.5)
            try:
                if page.evaluate("() => { const t = document.body.innerText || ''; return t.includes('削除しました') || t.includes('投稿を削除しました') || t.includes('Post deleted') || t.includes('Deleted'); }"):
                    return True, "deleted"
            except Exception:
                pass
            if is_permalink and "/post/" not in (page.url or ""):
                return True, "deleted"

        return True, "delete_confirmed"

    return False, "対象投稿が画面上に見つかりませんでした"

def _cli_process_due_deletes(page, context, account_name, prof_dir, log_func, restore_home=False):
    due_jobs = get_due_pending_deletes(prof_dir)
    if not due_jobs:
        return 0
    done_cnt = 0
    for job in due_jobs:
        pid = job.get("post_id") or job.get("row_idx") or "?"
        log_func(f"[INFO] 🗑️ [{account_name}] Post #{pid} の削除予定時刻に到達したため削除を実行します...")
        ok_del, detail = delete_threads_post_on_browser(page, context, account_name, job, log_func=log_func)
        if ok_del:
            if job.get("posts_file") and job.get("row_idx"):
                mark_post_deleted_in_excel(job["posts_file"], int(job["row_idx"]), deleted_time_str=datetime.now().strftime("%H:%M:%S"), sheet_name=job.get("sheet_name") or None)
            remove_pending_delete_job(prof_dir, job["job_id"])
            log_func(f"[SUCCESS] 🗑️ [{account_name}] Post #{pid} を自動削除しました！ ({detail})")
            done_cnt += 1
        else:
            if str(detail).startswith("checkpoint:"):
                log_func(f"[WARN] 🚨 [{account_name}] 自動削除中に本人確認画面を検知したため停止します ({detail})。")
                return -1
            attempts = int(job.get("attempts") or 0) + 1
            if attempts >= 3:
                remove_pending_delete_job(prof_dir, job["job_id"])
                log_func(f"[WARN] ⚠️ [{account_name}] Post #{pid} の自動削除を3回試行しましたが失敗したため予約解除しました ({detail})。")
            else:
                job["attempts"] = attempts
                job["delete_at_ts"] = time.time() + 180
                update_pending_delete_job(prof_dir, job)
                log_func(f"[WARN] ⏳ [{account_name}] Post #{pid} の自動削除を3分後に再試行します ({attempts}/3): {detail}")
    if restore_home:
        try:
            page.goto("https://www.threads.net/", timeout=30000, wait_until="domcontentloaded")
            time.sleep(1.5)
        except Exception:
            pass
    return done_cnt

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

def _normalize_time_text(val):
    s = unicodedata.normalize("NFKC", str(val or "")).strip()
    s = (
        s.replace("〜", "-")
        .replace("～", "-")
        .replace("ー", "-")
        .replace("−", "-")
        .replace("―", "-")
        .replace("‐", "-")
        .replace("–", "-")
        .replace("—", "-")
        .replace("から", "-")
        .replace("~", "-")
    )
    return re.sub(r'\bto\b', '-', s, flags=re.IGNORECASE)

def _parse_clock_token_to_min(tok, bound_kind=None):
    s = _normalize_time_text(tok).replace("翌", "").replace("午前", "").strip()
    is_pm = False
    if "午後" in s or "pm" in s.lower():
        is_pm = True
        s = re.sub(r'午後|pm', '', s, flags=re.IGNORECASE).strip()
    s = s.replace("半", "30").replace(" ", "")
    if not s:
        return None
    m_jp = re.fullmatch(r'(\d{1,2})時(?:(\d{1,2})分?)?', s)
    if m_jp:
        h = int(m_jp.group(1))
        m = int(m_jp.group(2)) if m_jp.group(2) is not None else (59 if bound_kind == "single_end" else 0)
        if is_pm and h < 12:
            h += 12
        if 0 <= h <= 24 and 0 <= m <= 59 and not (h == 24 and m > 0):
            return h * 60 + m
        return None
    m_col = re.fullmatch(r'(\d{1,2}):(\d{1,2})(?::\d{1,2})?', s)
    if m_col:
        h, m = int(m_col.group(1)), int(m_col.group(2))
        if is_pm and h < 12:
            h += 12
        if 0 <= h <= 24 and 0 <= m <= 59 and not (h == 24 and m > 0):
            return h * 60 + m
        return None
    m_num = re.fullmatch(r'(\d{1,2})', s)
    if m_num:
        h = int(m_num.group(1))
        if is_pm and h < 12:
            h += 12
        if 0 <= h <= 24:
            return 1440 if h == 24 else h * 60 + (59 if bound_kind == "single_end" else 0)
    return None

def is_within_target_hours(time_setting, current_hour=None):
    if current_hour is None:
        now_d = datetime.now()
        cur_min = now_d.hour * 60 + now_d.minute
        has_exact_min = True
    else:
        cur_min = int(current_hour) * 60
        has_exact_min = False
    if time_setting is None:
        return True
    raw = str(time_setting).strip()
    if not raw or raw.lower() in ["nan", "none", "終日", "24時間", "いつでも", "all"]:
        return True
    parts = [p.strip() for p in _normalize_time_text(raw).replace("、", ",").replace(";", ",").split(",") if p.strip()]
    if not parts:
        return True
    for part in parts:
        norm_p = part.replace(" ", "")
        if norm_p in ["0:00-0:00", "0-24", "0:00-24:00", "0-0", "0時-0時", "0時-24時", "終日", "24時間"]:
            return True
        if "-" in part:
            tokens = part.split("-", 1)
            s_min = _parse_clock_token_to_min(tokens[0])
            e_min = _parse_clock_token_to_min(tokens[1])
            if s_min is not None and e_min is not None:
                if s_min == e_min or (s_min == 0 and e_min >= 1440):
                    return True
                if has_exact_min:
                    if (s_min < e_min and s_min <= cur_min <= e_min) or (s_min > e_min and (cur_min >= s_min or cur_min <= e_min)):
                        return True
                else:
                    s_h, e_h, cur_h = s_min // 60, e_min // 60, cur_min // 60
                    if (s_h < e_h and s_h <= cur_h < e_h) or (s_h > e_h and (cur_h >= s_h or cur_h < e_h)) or s_h == e_h:
                        return True
        else:
            s_min = _parse_clock_token_to_min(part, bound_kind="single_start")
            e_min = _parse_clock_token_to_min(part, bound_kind="single_end")
            if s_min is not None and e_min is not None:
                if has_exact_min and s_min <= cur_min <= e_min:
                    return True
                if not has_exact_min and (s_min // 60) == (cur_min // 60):
                    return True
    return False

_WEEKDAY_TOKEN_MAP = {
    "月": 0, "月曜": 0, "月曜日": 0, "mon": 0, "monday": 0,
    "火": 1, "火曜": 1, "火曜日": 1, "tue": 1, "tues": 1, "tuesday": 1,
    "水": 2, "水曜": 2, "水曜日": 2, "wed": 2, "wednesday": 2,
    "木": 3, "木曜": 3, "木曜日": 3, "thu": 3, "thur": 3, "thurs": 3, "thursday": 3,
    "金": 4, "金曜": 4, "金曜日": 4, "fri": 4, "friday": 4,
    "土": 5, "土曜": 5, "土曜日": 5, "sat": 5, "saturday": 5,
    "日": 6, "日曜": 6, "日曜日": 6, "sun": 6, "sunday": 6,
}

def is_now_within_weekdays(now_dt, weekday_str):
    if now_dt is None:
        now_dt = datetime.now()
    if weekday_str is None:
        return True
    raw = unicodedata.normalize("NFKC", str(weekday_str)).strip()
    norm_all = raw.lower().replace(" ", "").replace("　", "")
    if not norm_all or norm_all in ["nan", "none", "毎日", "全日", "すべて", "全て", "全曜日", "all", "everyday", "指定なし"]:
        return True
    cleaned = re.sub(r'(?:毎週|曜日指定|のみ|だけ)', '', raw)
    cleaned = re.sub(r'\s*(?:[〜～~ー−―‐–—\-]|から|\bto\b)\s*', '-', cleaned, flags=re.IGNORECASE)
    allowed_days = set()
    tokens = [t.strip() for t in re.split(r'[,，、/;；\s　・&＆+]+', cleaned) if t.strip()]
    for tok in tokens:
        tl = tok.lower()
        if tl in ["毎日", "全日", "すべて", "全て", "all", "everyday"]:
            return True
        if tl in ["平日", "weekdays", "weekday"]:
            allowed_days.update([0, 1, 2, 3, 4])
            continue
        if tl in ["土日", "週末", "土日祝", "weekend", "weekends"]:
            allowed_days.update([5, 6])
            continue
        if "-" in tl:
            parts = tl.split("-", 1)
            d1 = _WEEKDAY_TOKEN_MAP.get(parts[0].strip())
            d2 = _WEEKDAY_TOKEN_MAP.get(parts[1].strip())
            if d1 is not None and d2 is not None:
                cur = d1
                for _ in range(7):
                    allowed_days.add(cur)
                    if cur == d2:
                        break
                    cur = (cur + 1) % 7
                continue
        if tl in _WEEKDAY_TOKEN_MAP:
            allowed_days.add(_WEEKDAY_TOKEN_MAP[tl])
            continue
        if tl and all(ch in "月火水木金土日" for ch in tl):
            for ch in tl:
                allowed_days.add(_WEEKDAY_TOKEN_MAP[ch])
    if not allowed_days:
        return True
    return now_dt.weekday() in allowed_days

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

def _get_current_machine_id():
    raw = "|".join([
        platform.node() or os.environ.get("COMPUTERNAME", ""),
        os.environ.get("USERNAME", "") or os.environ.get("USER", ""),
        platform.machine() or "",
        sys.platform or "",
    ])
    return hashlib.sha256(raw.encode("utf-8", errors="ignore")).hexdigest()[:16]

def extract_proxy_from_account(account_rec):
    if not isinstance(account_rec, dict):
        return ""
    for k in ["プロキシ", "proxy", "Proxy", "PROXY", "プロキシ設定"]:
        val = str(account_rec.get(k) or "").strip()
        if val and val.lower() not in ["nan", "none", "-", "なし"] and not val.startswith(("(", "（", "※")):
            return val
    return ""

def parse_playwright_proxy(raw_proxy):
    s = str(raw_proxy or "").strip()
    if not s or s.lower() in ["nan", "none", "-", "なし"]:
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
            log_func(f"[INFO] 🧹 端末保護Cookie(datr/mid/ig_did)を維持しつつ、認証Cookieのみ {removed} 件リセットしました。")
    except Exception:
        try:
            for d_target in target_domains:
                context.clear_cookies(domain=d_target)
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
    has_chrome_default = os.path.exists(os.path.join(profile_dir, "Default"))
    host_changed = bool(prev_host and prev_host != cur_host)
    legacy_migrated = bool(not prev_host and has_chrome_default and has_portable)
    if host_changed or legacy_migrated:
        if has_portable:
            for stale_rel in ["Default/Cookies", "Default/Cookies-journal", "Default/Network/Cookies", "Default/Network/Cookies-journal"]:
                sp = os.path.join(profile_dir, stale_rel)
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
                log_func(f"[WARN] 🌐 [{account_name}] 別PC/VPS起動を検知したためDPAPI再バインド＆移行保護を有効化します。")
    fp["last_host_id"] = cur_host
    try:
        with open(fp_path, "w", encoding="utf-8") as f:
            json.dump(fp, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return host_changed

def restore_portable_session_if_needed(context, profile_dir, account_name, log_func=print):
    if not profile_dir or not os.path.exists(profile_dir):
        return False
    portable_file = os.path.join(profile_dir, "portable_session.json")
    if not os.path.exists(portable_file) or os.path.getsize(portable_file) < 5:
        portable_file = os.path.join(profile_dir, "cookies.json")
    if not os.path.exists(portable_file) or os.path.getsize(portable_file) < 5:
        return False
    rebind_marker = os.path.join(profile_dir, ".portable_needs_rebind")
    force_rebind = os.path.exists(rebind_marker)
    try:
        with open(portable_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        cookies_to_add = data.get("cookies", []) if isinstance(data, dict) else (data if isinstance(data, list) else [])
        cur_cookies = context.cookies(["https://www.threads.net", "https://www.threads.com", "https://www.instagram.com"])
        cur_names = {c.get("name") for c in cur_cookies if c.get("value")}
        has_session = "sessionid" in cur_names
        missing_device_cookies = any(dn not in cur_names for dn in ("datr", "mid", "ig_did"))
        if cookies_to_add and (force_rebind or not has_session or missing_device_cookies):
            valid_cks = [c for c in cookies_to_add if isinstance(c, dict) and c.get("name") and c.get("value")]
            for c_item in valid_cks:
                if "sameSite" in c_item and c_item["sameSite"] not in ["Strict", "Lax", "None"]:
                    c_item["sameSite"] = "Lax"
            try:
                context.add_cookies(valid_cks)
            except Exception:
                for c_item in valid_cks:
                    try:
                        context.add_cookies([c_item])
                    except Exception:
                        pass
            if log_func:
                log_func(f"[INFO] ♻️ [{account_name}] portable_session.json から全 {len(valid_cks)} 件の正規Cookieを復元・再バインドしました。")
        if force_rebind and os.path.exists(rebind_marker):
            try:
                os.remove(rebind_marker)
            except Exception:
                pass
        return True
    except Exception:
        return False

def save_portable_session_file(context, profile_dir, account_name="", log_func=print):
    if not profile_dir:
        return False
    try:
        for pg in getattr(context, "pages", []):
            try:
                if not pg.is_closed():
                    pg_u = (pg.url or "").lower()
                    if any(k in pg_u for k in ["/challenge/", "/checkpoint/", "auth_platform", "two_step_verification", "/accounts/suspended"]):
                        return False
            except Exception:
                pass
        all_cks = context.cookies(["https://www.threads.net", "https://www.threads.com", "https://www.instagram.com"])
        has_sess = any(c.get("name") == "sessionid" and len(str(c.get("value") or "")) > 15 for c in all_cks)
        if not has_sess:
            return False
        os.makedirs(profile_dir, exist_ok=True)
        context.storage_state(path=os.path.join(profile_dir, "portable_session.json"))
        with open(os.path.join(profile_dir, "cookies.json"), "w", encoding="utf-8") as cf:
            json.dump(all_cks, cf, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False

def safe_launch_persistent_context(p, profile_dir, headless=False, log_func=print, proxy_str=""):
    os.makedirs(profile_dir, exist_ok=True)
    _cleanup_profile_locks(profile_dir)
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
    detect_and_mark_host_migration(profile_dir, account_name=os.path.basename(profile_dir), log_func=log_func)
    fp_path = os.path.join(profile_dir, "fingerprint.json")
    fp = {}
    if os.path.exists(fp_path):
        try:
            with open(fp_path, "r", encoding="utf-8") as f:
                fp = json.load(f) or {}
        except Exception:
            fp = {}
    proxy_cfg = parse_playwright_proxy(proxy_str)
    if proxy_cfg and log_func:
        log_func(f"[INFO] 🌐 [{os.path.basename(profile_dir)}] 専用プロキシ経由で接続: {proxy_cfg['server']}")
    launch_kwargs = dict(
        user_data_dir=profile_dir,
        headless=headless,
        channel="chrome",
        locale=fp.get("locale", "ja-JP"),
        timezone_id=fp.get("timezone_id", "Asia/Tokyo"),
        viewport=fp.get("viewport", {"width": 1280, "height": 850}),
        args=["--lang=ja-JP", "--disable-dev-shm-usage", "--test-type", "--no-first-run", "--no-default-browser-check", "--disable-session-crashed-bubble", "--hide-crash-restore-bubble"],
        ignore_default_args=["--enable-automation"],
    )
    if proxy_cfg:
        launch_kwargs["proxy"] = proxy_cfg
    for attempt in range(1, 5):
        try:
            try:
                context = p.chromium.launch_persistent_context(**launch_kwargs)
            except Exception as chrome_err:
                if "Executable doesn't exist" in str(chrome_err) or "Failed to launch chrome" in str(chrome_err):
                    if log_func:
                        log_func(f"[WARN] Chrome起動失敗のため内蔵Chromiumで起動します: {chrome_err}")
                    fallback_kwargs = {k: v for k, v in launch_kwargs.items() if k != "channel"}
                    context = p.chromium.launch_persistent_context(**fallback_kwargs)
                else:
                    raise
            try:
                context.add_init_script("Object.defineProperty(navigator, 'webdriver', { get: () => undefined });")
            except Exception:
                pass
            restore_portable_session_if_needed(context, profile_dir, os.path.basename(profile_dir), log_func=log_func)
            return context
        except Exception as e:
            err_str = str(e)
            is_lock_err = any(k in err_str for k in ["Target page, context or browser has been closed", "closed", "ProcessSingleton", "Singleton", "user data directory is already in use"])
            if is_lock_err and attempt < 4:
                zombie_pids = _find_chrome_pids_for_profile(profile_dir, pipe_only=True)
                if zombie_pids:
                    if log_func:
                        log_func(f"[WARN] [{os.path.basename(profile_dir)}] 残留Chromeプロセス ({len(zombie_pids)}件) をクリーンアップして再起動します... ({attempt}/3)")
                    _terminate_chrome_pids(zombie_pids, wait_sec=2.0)
                else:
                    if log_func:
                        log_func(f"[WARN] [{os.path.basename(profile_dir)}] Chromeロック残骸または手動Chromeとの競合を検知しました（開いている場合は閉じてください）。再試行します... ({attempt}/3)")
                    time.sleep(1.8 + attempt * 0.5)
                _cleanup_profile_locks(profile_dir)
                continue
            raise

def inject_session_cookies(context, raw_token, account_name, log_func=print):
    token = str(raw_token or "").strip()
    if not token or token.lower() in ["nan", "none"] or token.startswith("ここに"):
        return False
    if token.startswith("sessionid="):
        token = token[len("sessionid="):].strip()
    token = token.strip("\"\x27 ;\t\r\n")
    if not token:
        return False
    try:
        cur_cks = context.cookies(["https://www.threads.net", "https://www.threads.com", "https://www.instagram.com"])
        if any(c.get("name") == "sessionid" and str(c.get("value") or "").strip() == token for c in cur_cks):
            return True
    except Exception:
        pass
    decoded_token = urllib.parse.unquote(token)
    ds_user_id = ""
    if ":" in decoded_token:
        potential_id = decoded_token.split(":")[0].strip()
        if potential_id.isdigit():
            ds_user_id = potential_id
    cookie_expires = int(time.time()) + 31536000
    cookies_by_domain = []
    for domain in [".threads.net", "www.threads.net", ".threads.com", "www.threads.com", ".instagram.com", "www.instagram.com"]:
        cookies_by_domain.append({
            "name": "sessionid",
            "value": token,
            "domain": domain,
            "path": "/",
            "httpOnly": True,
            "secure": True,
            "sameSite": "Lax",
            "expires": cookie_expires,
        })
        if ds_user_id:
            cookies_by_domain.append({
                "name": "ds_user_id",
                "value": str(ds_user_id),
                "domain": domain,
                "path": "/",
                "secure": True,
                "sameSite": "Lax",
                "expires": cookie_expires,
            })
    try:
        context.add_cookies(cookies_by_domain)
    except Exception:
        pass
    log_func(f"[SUCCESS] [{account_name}] エクセルの sessionid (端末保護Cookieは変更せず保持) を注入しました。")
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

            is_cp_pre, cp_why_pre = detect_phone_checkpoint_screen(page)
            if is_cp_pre:
                log_func(f"[WARN] 🚨 [{account_name}] 電話番号/本人確認画面を検知したためID/PWログインを即時中断します（{cp_why_pre}）。")
                return False

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

        # 端末識別Cookie(datr/mid/ig_did)は維持したまま、壊れた認証Cookieのみリセット
        clear_auth_cookies_preserve_device(context, log_func=log_func)
        time.sleep(0.5)

        ok = execute_id_pw_login(page, context, account_name, password, log_func=log_func)
        if ok:
            p_dir = os.path.join(BASE_DIR, f"chrome_profiles/{account_name}")
            save_portable_session_file(context, p_dir, account_name, log_func)
            if acc_state is not None:
                acc_state["login_cooldown_until"] = 0
                acc_state["login_cooldown_started"] = False
                acc_state["pure_profile_mode"] = True
            # 新しいセッションIDをアカウント情報.xlsxに同期
            save_context_cookies_to_excel(context, filepath, account_name, log_func=log_func)
            return True
        else:
            is_cp_end, _ = detect_phone_checkpoint_screen(page)
            if is_cp_end and acc_state is not None:
                acc_state["checkpoint_paused"] = True
            log_func(f"[WARN] ❌ [{account_name}] ID/PW自動ログインに失敗しました。電話番号ロック画面やパスワードをご確認ください。")
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
            global_time_slot = custom_settings.get("time_slot", "0:00-0:00")

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
            unposted = [p_it for p_it in all_posts if str(p_it.get("ステータス", "")).strip() == "未投稿"]
            ai_cfg_cli = read_ai_post_config_from_xlsx(posts_file)
            now_dt_check = datetime.now()

            ai_pending_rows = [
                r for r in (ai_cfg_cli.get("rows") or [])
                if not any(x in str(r.get("status") or "").strip() for x in ("完了", "投稿済", "削除"))
                and str(r.get("prompt") or "").strip()
            ]
            ai_has_sections = any(str(v or "").strip() for v in (ai_cfg_cli.get("prompt_sections") or {}).values())
            ai_runnable_now = bool(
                ai_cfg_cli.get("enabled")
                and is_now_within_weekdays(now_dt_check, ai_cfg_cli.get("weekdays", "毎日"))
                and is_within_target_hours(ai_cfg_cli.get("time_slot", "0:00-0:00"))
                and (ai_pending_rows or ai_has_sections or ai_cfg_cli.get("memo"))
            )

            eligible_unposted = []
            for p_it in unposted:
                s_meta = p_it.get("_sheet_theme_meta") if isinstance(p_it.get("_sheet_theme_meta"), dict) else {}
                if s_meta and not s_meta.get("enabled", True):
                    continue
                row_slot = ""
                for k_sl in ["投稿時間", "投稿時間帯", "投稿時刻", "実行時間", "時間帯"]:
                    v_sl = str(p_it.get(k_sl) or "").strip()
                    if v_sl and v_sl.lower() not in ["none", "nan"]:
                        row_slot = v_sl
                        break
                row_wd = ""
                for k_wd in ["曜日", "対象曜日", "weekdays", "weekday"]:
                    v_wd = str(p_it.get(k_wd) or "").strip()
                    if v_wd and v_wd.lower() not in ["none", "nan"]:
                        row_wd = v_wd
                        break
                eff_wd = row_wd or str(s_meta.get("weekdays") or "").strip() or "毎日"
                eff_slot = row_slot or str(s_meta.get("time_slot") or "").strip() or global_time_slot or "0:00-0:00"
                if is_now_within_weekdays(now_dt_check, eff_wd) and is_within_target_hours(eff_slot):
                    eligible_unposted.append(p_it)

            if not eligible_unposted and not ai_runnable_now:
                print(f"[SKIP] [{name}] 実行可能な固定ポストおよびAIポストがありません。")
                continue
            unposted = eligible_unposted

            if unposted:
                if post_mode == 1:
                    random.shuffle(unposted)
                    print(f"[{name}] 【モード1: ランダム】未投稿 {len(unposted)} 件をランダム順序で投稿します。")
                else:
                    print(f"[{name}] 【モード0: 上から順】未投稿 {len(unposted)} 件を先頭から順に投稿します。")

            acc_log = make_account_logger(name)
            get_account_images_dir(name, create=True)

            # AIポストが有効なら固定AI専用Window(_ai_window)またはAPIで投稿文を事前生成（例外発生時も固定ポスト窓起動は止めない）
            ai_generated_payload = None
            if ai_runnable_now:
                try:
                    chosen_ai_row = ai_pending_rows[0] if ai_pending_rows else None
                    ai_txt, ai_used_eng, ai_gen_err = generate_ai_post_text(
                        name, ai_cfg_cli, prompt_row=chosen_ai_row, playwright_obj=p, base_dir=BASE_DIR, log_func=acc_log
                    )
                    if ai_txt:
                        ai_generated_payload = {
                            "text": ai_txt,
                            "engine": ai_used_eng,
                            "row": chosen_ai_row,
                        }
                        acc_log(f"[INFO] 🤖 [{name}] AIポスト事前生成完了 (使用AI: {ai_used_eng} / {len(ai_txt)}文字)")
                    else:
                        acc_log(f"[ERROR] [{name}] AIポスト生成失敗 ({ai_used_eng}): {ai_gen_err}")
                except Exception as e_ai_gen_cli:
                    acc_log(f"[ERROR] [{name}] AIポスト事前生成例外（固定ポストの実行は継続します）: {e_ai_gen_cli}")

            if not unposted and not ai_generated_payload:
                continue

            prof_folder = sanitize_account_folder_name(name)
            if prof_folder == AI_WINDOW_PROFILE_NAME:
                prof_folder = f"account_{prof_folder}"
            prof_dir = os.path.join(BASE_DIR, f"chrome_profiles/{prof_folder}")
            os.makedirs(prof_dir, exist_ok=True)
            proxy_str = extract_proxy_from_account(acc)
            acc_log(f"[INFO] 🪟 [{name}] アカウント別Window (chrome_profiles/{prof_folder}) を起動します...")
            context = safe_launch_persistent_context(p, prof_dir, headless=False, log_func=acc_log, proxy_str=proxy_str)
            
            # auth_token (sessionid) 注入
            session_token = str(acc.get("auth_token") or acc.get("Authトークン_sessionid") or acc.get("sessionid") or "").strip()
            if session_token:
                inject_session_cookies(context, session_token, name, log_func=acc_log)

            page = context.pages[0] if context.pages else context.new_page()
            page.goto("https://www.threads.net/", timeout=60000)
            time.sleep(3)

            is_cp_cli, cp_why_cli = detect_phone_checkpoint_screen(page)
            if is_cp_cli:
                acc_log(f"[WARN] 🚨 [{name}] 携帯電話番号（SMS）確認画面を検知したためアカウント保護停止します（{cp_why_cli}）。")
                context.close()
                continue

            cli_acc_state = {"name": name}
            is_ok = ensure_authenticated_session(page, context, acc, ACCOUNT_CONFIG_FILE, log_func=acc_log, wait_timeout_sec=300, acc_state=cli_acc_state)
            if not is_ok or cli_acc_state.get("checkpoint_paused"):
                acc_log(f"[ERROR] [{name}] ログイン認証が未完了またはロック検知のためスキップします。")
                context.close()
                continue

            _cli_process_due_deletes(page, context, name, prof_dir, acc_log, restore_home=True)

            if ai_generated_payload:
                try:
                    ai_txt_p = ai_generated_payload["text"]
                    ai_row_p = ai_generated_payload["row"]
                    ai_eng_p = ai_generated_payload["engine"]
                    for close_s in ['svg[aria-label="閉じる"]', 'div[role="button"][aria-label="閉じる"]']:
                        try:
                            if page.locator(close_s).first.is_visible(timeout=500):
                                page.locator(close_s).first.click()
                                time.sleep(0.8)
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
                    tb_ai = page.locator('div[role="textbox"], div[contenteditable="true"]').first
                    tb_ai.click()
                    for line_ai in ai_txt_p.split("\n"):
                        tb_ai.press_sequentially(line_ai, delay=random.randint(15, 30))
                        tb_ai.press("Enter")
                        time.sleep(0.1)
                    time.sleep(1.5)
                    page.locator('div[role="button"]:has-text("投稿する"), div[role="button"]:has-text("投稿")').last.click()
                    time.sleep(4)
                    fin_ai = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    w_row_ai = record_ai_post_result_in_xlsx(
                        posts_file,
                        row_idx=(ai_row_p.get("_row_idx") if ai_row_p else None),
                        prompt_text=((ai_row_p.get("prompt") if ai_row_p else ai_cfg_cli.get("memo")) or ""),
                        topic=(ai_row_p.get("topic") if ai_row_p else ""),
                        delete_after=(ai_row_p.get("delete_after") if ai_row_p else ""),
                        generated_text=ai_txt_p,
                        engine_used=ai_eng_p,
                        status="投稿完了",
                        finish_time=fin_ai,
                    )
                    acc_log(f"[SUCCESS] [{name}] 🤖 AIポスト投稿完了 (使用AI: {ai_eng_p})！AIポストシート更新")
                    del_ai = (ai_row_p.get("delete_after") if ai_row_p else "")
                    if del_ai and (w_row_ai or (ai_row_p and ai_row_p.get("_row_idx"))):
                        enqueue_pending_delete_job(
                            profile_dir=prof_dir,
                            account_name=name,
                            posts_file=posts_file,
                            row_idx=int(w_row_ai or ai_row_p["_row_idx"]),
                            sheet_name=ai_cfg_cli.get("sheet_name") or "AIポスト",
                            post_id=(ai_row_p.get("id") if ai_row_p else "AI"),
                            delete_spec=del_ai,
                            profile_handle=name,
                            text_content=ai_txt_p,
                        )
                except Exception as e_ai_cli:
                    acc_log(f"[ERROR] [{name}] 🤖 AIポスト投稿エラー: {e_ai_cli}")

            for post_item in unposted:
                if get_due_pending_deletes(prof_dir):
                    if _cli_process_due_deletes(page, context, name, prof_dir, acc_log, restore_home=True) == -1:
                        break
                s_meta = post_item.get("_sheet_theme_meta") if isinstance(post_item.get("_sheet_theme_meta"), dict) else {}
                row_slot = ""
                for k_sl in ["投稿時間", "投稿時間帯", "投稿時刻", "実行時間", "時間帯"]:
                    v_sl = str(post_item.get(k_sl) or "").strip()
                    if v_sl and v_sl.lower() not in ["none", "nan"]:
                        row_slot = v_sl
                        break
                row_wd = ""
                for k_wd in ["曜日", "対象曜日", "weekdays", "weekday"]:
                    v_wd = str(post_item.get(k_wd) or "").strip()
                    if v_wd and v_wd.lower() not in ["none", "nan"]:
                        row_wd = v_wd
                        break
                eff_wd = row_wd or str(s_meta.get("weekdays") or "").strip() or "毎日"
                eff_slot = row_slot or str(s_meta.get("time_slot") or "").strip() or global_time_slot or "0:00-0:00"
                post_id_skip = post_item.get("ID") or post_item.get("id") or post_item["_row_idx"] - 1
                if not is_now_within_weekdays(datetime.now(), eff_wd):
                    acc_log(f"[INFO] ⏭️ [{name}] Post #{post_id_skip} は対象曜日 ({eff_wd}) 外のため即スキップして次のポストへ進みます。")
                    continue
                if not is_within_target_hours(eff_slot):
                    acc_log(f"[INFO] ⏭️ [{name}] Post #{post_id_skip} は投稿時間 ({eff_slot}) 外のため即スキップして次のポストへ進みます。")
                    continue

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

                captured_url = ""
                try:
                    captured_url = page.evaluate("""(handle) => {
                        const hLower = (handle || '').toLowerCase();
                        const toastLinks = Array.from(document.querySelectorAll('[role="status"] a[href*="/post/"], [role="alert"] a[href*="/post/"], div[aria-live] a[href*="/post/"]'));
                        for (const a of toastLinks) {
                            const href = (a.getAttribute('href') || '').trim();
                            if (href) return href.startsWith('http') ? href : ('https://www.threads.com' + href);
                        }
                        if (hLower) {
                            const links = Array.from(document.querySelectorAll('a[href*="/post/"]'));
                            for (const a of links) {
                                const href = (a.getAttribute('href') || '').trim();
                                if (href.toLowerCase().includes('/@' + hLower + '/post/')) {
                                    return href.startsWith('http') ? href : ('https://www.threads.com' + href);
                                }
                            }
                        }
                        return '';
                    }""", name or "") or ""
                except Exception:
                    captured_url = ""

                finish_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                update_post_status(posts_file, row_idx, status="投稿完了", finish_time=finish_time, sheet_name=post_item.get("_sheet_name"))
                acc_log(f"[SUCCESS] [{name}] Post #{post_id} 投稿完了！Excel更新")

                del_spec_cli = _extract_delete_after_cli(post_item)
                if del_spec_cli:
                    del_job = enqueue_pending_delete_job(
                        profile_dir=prof_dir,
                        account_name=name,
                        posts_file=posts_file,
                        row_idx=row_idx,
                        sheet_name=post_item.get("_sheet_name") or "",
                        post_id=post_id,
                        delete_spec=del_spec_cli,
                        post_url=captured_url,
                        profile_handle=name,
                        text_content=text,
                    )
                    if del_job:
                        acc_log(f"[INFO] 🗑️ [{name}] Post #{post_id} の自動削除を予約: {_format_duration_jp(del_job['delay_sec'])}後 ({datetime.fromtimestamp(del_job['delete_at_ts']).strftime('%H:%M:%S')} 予定)")

                wait_sec = random.randint(min_d, max_d)
                if wait_sec >= 60:
                    acc_log(f"[*] BAN防止のため {wait_sec} 秒 (約 {wait_sec // 60} 分) 待機中...")
                else:
                    acc_log(f"[*] BAN防止のため {wait_sec} 秒待機中...")
                for _w_i in range(wait_sec):
                    if _w_i % 5 == 0 and get_due_pending_deletes(prof_dir):
                        if _cli_process_due_deletes(page, context, name, prof_dir, acc_log, restore_home=True) == -1:
                            break
                    time.sleep(1)

            while True:
                rem_jobs_cli = load_pending_deletes(prof_dir)
                if not rem_jobs_cli:
                    break
                if get_due_pending_deletes(prof_dir):
                    if _cli_process_due_deletes(page, context, name, prof_dir, acc_log, restore_home=False) == -1:
                        break
                    continue
                next_ts = min(float(j.get("delete_at_ts") or time.time()) for j in rem_jobs_cli)
                time.sleep(max(1, min(5, int(next_ts - time.time()))))

            context.close()

if __name__ == "__main__":
    main()
