/**
 * Spintax and dynamic content processing utilities
 */

export function parseSpintax(text: string): string {
  if (!text) return '';
  let result = text;
  const spintaxRegex = /\{([^{}]+)\}/g;
  
  // Recursively resolve nested or consecutive spintax blocks
  let safetyLimit = 10;
  while (spintaxRegex.test(result) && safetyLimit > 0) {
    result = result.replace(spintaxRegex, (_match, group) => {
      const choices = group.split('|');
      const pick = choices[Math.floor(Math.random() * choices.length)];
      return pick.trim();
    });
    safetyLimit--;
  }

  // Dynamic date/time placeholders
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const date = now.getDate();
  const weekdays = ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'];
  const weekday = weekdays[now.getDay()];

  result = result
    .replace(/\{today\}/gi, `${year}年${month}月${date}日(${weekday.charAt(0)})`)
    .replace(/\{date\}/gi, `${month}月${date}日`)
    .replace(/\{month\}/gi, `${month}月`)
    .replace(/\{day\}/gi, `${date}日`)
    .replace(/\{weekday\}/gi, weekday);

  return result;
}

/**
 * Extracts numeric user ID from auth_token / sessionid
 * Supports formats: "47094044745:HP477WPO..." or "59443198637%3AoWW..."
 */
export function extractUserIdFromToken(token: string): string {
  if (!token) return '';
  const decoded = decodeURIComponent(token.trim());
  if (decoded.includes(':')) {
    const parts = decoded.split(':');
    if (/^\d+$/.test(parts[0])) return parts[0];
  }
  if (token.includes('%3A') || token.includes('%3a')) {
    const parts = token.split(/%3[Aa]/);
    if (/^\d+$/.test(parts[0])) return parts[0];
  }
  return '';
}

/**
 * Resolves image path from cell or detects embedded image filename in post text
 */
export function resolveImagePath(
  rawImgPath?: string,
  postText?: string,
  customImages?: Record<string, string>
): string {
  if (rawImgPath && rawImgPath.trim()) {
    const clean = rawImgPath.trim();
    if (customImages) {
      const base = clean.split(/[\\/]/).pop() || clean;
      if (customImages[clean]) return customImages[clean];
      if (customImages[base]) return customImages[base];
    }
    if (clean.includes('image2.png')) return '/images/image2.png';
    if (clean.includes('image.png')) return '/images/image.png';
    if (clean.startsWith('http://') || clean.startsWith('https://') || clean.startsWith('data:')) {
      return clean;
    }
    return clean;
  }

  if (postText) {
    if (postText.includes('image2.png')) return '/images/image2.png';
    if (postText.includes('image.png')) return '/images/image.png';
  }

  return '';
}

export interface ParsedDeleteDelay {
  minSec: number;
  maxSec: number;
  label: string;
}

/**
 * Parses per-post 削除時間 spec strings like:
 * - "30～60", "30-60", "30分～60分" -> { minSec: 1800, maxSec: 3600, label: "30〜60分" }
 * - "1～3時間", "10～30秒", "45分", "2時間"
 * - Blank / "0" / "削除しない" / "残す" -> null
 */
export function parseDeleteDelaySpec(raw?: string): ParsedDeleteDelay | null {
  if (!raw) return null;
  const s = raw.normalize('NFKC').trim();
  if (!s || s.startsWith('(') || s.startsWith('（') || s.startsWith('※')) return null;

  const lowCompact = s.toLowerCase().replace(/[\s　]+/g, '');
  if (['nan', 'none', 'null', '0', 'なし', '無し', '削除しない', '残す', 'off', 'no', 'false', '-', 'ー'].includes(lowCompact)) {
    return null;
  }

  const detectUnit = (tok: string, defaultMult = 60): { mult: number; unit: string } => {
    const low = tok.toLowerCase();
    if (tok.includes('時間') || /\d\s*(?:hr|hrs|hour|hours|h)\b/i.test(low)) {
      return { mult: 3600, unit: '時間' };
    }
    if (tok.includes('秒') || /\d\s*(?:sec|secs|second|seconds|s)\b/i.test(low)) {
      return { mult: 1, unit: '秒' };
    }
    if (tok.includes('分') || /\d\s*(?:min|mins|minute|minutes|m)\b/i.test(low)) {
      return { mult: 60, unit: '分' };
    }
    return { mult: defaultMult, unit: defaultMult === 3600 ? '時間' : defaultMult === 1 ? '秒' : '分' };
  };

  const normalized = s.replace(/\s*(?:[〜～~ー−―‐–—\-]|から|\bto\b)\s*/gi, '-');
  if (normalized.includes('-')) {
    const [p1, p2] = normalized.split('-', 2);
    const m1 = p1.match(/\d+(?:\.\d+)?/);
    const m2 = p2.match(/\d+(?:\.\d+)?/);
    if (m1 && m2) {
      const u2 = detectUnit(p2, 60);
      const u1 = detectUnit(p1, u2.mult);
      const v1 = parseFloat(m1[0]) * u1.mult;
      const v2 = parseFloat(m2[0]) * u2.mult;
      const minSec = Math.max(1, Math.round(Math.min(v1, v2)));
      const maxSec = Math.max(minSec, Math.round(Math.max(v1, v2)));
      const label = u1.unit === u2.unit ? `${m1[0]}〜${m2[0]}${u2.unit}` : `${m1[0]}${u1.unit}〜${m2[0]}${u2.unit}`;
      return { minSec, maxSec, label };
    }
  }

  const mSingle = s.match(/\d+(?:\.\d+)?/);
  if (mSingle) {
    const u = detectUnit(s, 60);
    const val = parseFloat(mSingle[0]);
    if (val <= 0) return null;
    const sec = Math.max(1, Math.round(val * u.mult));
    return { minSec: sec, maxSec: sec, label: `${mSingle[0]}${u.unit}` };
  }

  return null;
}

export function formatDurationJp(seconds: number): string {
  const s = Math.max(0, Math.round(seconds));
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const rem = s % 60;
  if (h > 0) return m > 0 ? `${h}時間${m}分` : `${h}時間`;
  if (m > 0) return rem > 0 ? `${m}分${rem}秒` : `${m}分`;
  return `${rem}秒`;
}

import type { AiEngineId, AiPostConfig, AiPromptSections } from '../types';

export interface AiEngineMeta {
  id: AiEngineId;
  label: string;
  defaultModel: string;
  modeNote: string;
  placeholderKeyOrUrl: string;
}

export const AI_ENGINES_META: AiEngineMeta[] = [
  { id: 'groq', label: 'groq (Groq高速API)', defaultModel: 'llama-3.3-70b-versatile', modeNote: 'API / 自動フォールバック対応', placeholderKeyOrUrl: 'gsk_... (Groq APIキー)' },
  { id: 'gemini', label: 'gemini (Google Gemini)', defaultModel: 'gemini-2.5-flash', modeNote: 'API または AI専用Window', placeholderKeyOrUrl: 'AIza... (空欄でAI専用Window利用可)' },
  { id: 'duckai', label: 'duckai (Duck.ai 無料チャット)', defaultModel: 'gpt-4o-mini', modeNote: 'AI専用Window推奨 / 無料', placeholderKeyOrUrl: '空欄可 (AI専用Windowで自動実行)' },
  { id: 'venice', label: 'venice (Venice AI)', defaultModel: 'llama-3.3-70b', modeNote: 'API または AI専用Window', placeholderKeyOrUrl: 'Venice APIキー (空欄でWindow利用可)' },
  { id: 'deepai', label: 'deepai (DeepAI)', defaultModel: 'standard', modeNote: 'API または AI専用Window', placeholderKeyOrUrl: 'DeepAI APIキー (空欄でWindow利用可)' },
  { id: 'talkai', label: 'talkai (TalkAI 無料Web)', defaultModel: 'gpt-4o-mini', modeNote: 'AI専用Window連携', placeholderKeyOrUrl: '空欄可 (AI専用Windowで自動実行)' },
  { id: 'chatgpt', label: 'chatgpt (OpenAI ChatGPT)', defaultModel: 'gpt-4o-mini', modeNote: 'API または AI専用Window', placeholderKeyOrUrl: 'sk-... (空欄でAI専用Window利用可)' },
  { id: 'qwen', label: 'qwen (Qwen Chat / DashScope)', defaultModel: 'qwen-plus', modeNote: 'API または AI専用Window', placeholderKeyOrUrl: 'DashScope APIキー (空欄でWindow利用可)' },
  { id: 'mistral', label: 'mistral (Mistral / Le Chat)', defaultModel: 'mistral-small-latest', modeNote: 'API または AI専用Window', placeholderKeyOrUrl: 'Mistral APIキー (空欄でWindow利用可)' },
  { id: 'grok', label: 'grok (xAI Grok)', defaultModel: 'grok-3-mini', modeNote: 'API または AI専用Window', placeholderKeyOrUrl: 'xai-... (空欄でAI専用Window利用可)' },
  { id: 'deepseek', label: 'deepseek (DeepSeek)', defaultModel: 'deepseek-chat', modeNote: 'API または AI専用Window', placeholderKeyOrUrl: 'sk-... (空欄でAI専用Window利用可)' },
  { id: 'claude', label: 'claude (Anthropic Claude)', defaultModel: 'claude-3-5-haiku-latest', modeNote: 'API または AI専用Window', placeholderKeyOrUrl: 'sk-ant-... (空欄でAI専用Window利用可)' },
  { id: 'perplexity', label: 'perplexity (Perplexity Sonar)', defaultModel: 'sonar', modeNote: 'API または AI専用Window', placeholderKeyOrUrl: 'pplx-... (空欄でAI専用Window利用可)' },
  { id: 'copilot', label: 'copilot (Microsoft Copilot)', defaultModel: 'copilot-web', modeNote: 'AI専用Window / 互換URL', placeholderKeyOrUrl: '空欄可 (AI専用Window または 接続URL)' },
  { id: 'openrouter', label: 'openrouter (OpenRouter)', defaultModel: 'google/gemini-2.0-flash-exp:free', modeNote: 'API専用 (統合ゲートウェイ)', placeholderKeyOrUrl: 'sk-or-v1-... (OpenRouter APIキー)' },
  { id: 'ollama', label: 'ollama (ローカルOllama)', defaultModel: 'llama3.2', modeNote: 'ローカル接続URL (完全無料)', placeholderKeyOrUrl: 'http://127.0.0.1:11434 (空欄=既定URL)' },
  { id: 'lmstudio', label: 'lmstudio (LM Studioローカル)', defaultModel: 'local-model', modeNote: 'ローカルOpenAI互換URL', placeholderKeyOrUrl: 'http://127.0.0.1:1234 (空欄=既定URL)' },
  { id: 'chrome_builtin', label: 'chrome_builtin (Chrome内蔵AI)', defaultModel: 'gemini-nano', modeNote: '固定AI専用Window内蔵AI専用', placeholderKeyOrUrl: '空欄可 (固定AI専用Window内で実行)' },
];

export const FIXED_AI_SYSTEM_PROMPT = `以下の内容を確認したうえで、設定されたキャラ・指示に沿った魅力的なスレッズのポストのメッセージを作成してください。
【作成方針（自動選定）】
設定されたキャラ・指示に沿った切り口を内部で3パターン（共感・補足・視点の共有など）検討し、対象の投稿に最も自然で噛み合っている1つだけを選んでください。
※検討過程や他案・比較・番号（1. 2. 3. 等）は一切書かず、選んだ1つのポストの本文のみを出力してください。`;

export interface AiPromptSectionDef {
  key: keyof AiPromptSections;
  label: string;
  defaultValue: string;
  description: string;
}

export const AI_PROMPT_SECTION_META: AiPromptSectionDef[] = [
  {
    key: 'character',
    label: 'キャラ設定',
    defaultValue: 'あなたはThreads(スレッズ)で役立つ発信をする親しみやすく魅力的なユーザーです。',
    description: '年齢層・立場・発信ジャンル・運用目的など、アカウントの人物像を設定します（空欄で省略可）',
  },
  {
    key: 'target',
    label: 'ターゲット層',
    defaultValue: '日々の暮らしや仕事を少し良くしたいと思っている20〜40代のユーザー。小さな悩みや気づきに共感しやすい層。',
    description: '「誰に届けるか」の指定です。発信側の「キャラ設定」とは別に、読者の属性・悩み・興味関心をAIに伝えます（空欄で省略可）',
  },
  {
    key: 'purpose',
    label: '投稿の目的',
    defaultValue: '読者の共感や「なるほど」を引き出し、いいね・保存・プロフィール閲覧やリプライにつなげること。',
    description: 'プロフィール画面への誘導、リンクのクリック、リプライの獲得など、読者に最終的にどう行動してほしいかの着地点です（空欄で省略可）',
  },
  {
    key: 'toneRule',
    label: '口調・文体ルール',
    defaultValue: '一人称は使わず自然なです・ます調で、親しみやすく柔らかい文体にしてください。絵文字は文末に0〜1個まで、感嘆符（！）の連続は控えてください。',
    description: '一人称・敬語/タメ口・絵文字や記号の癖など、文章のトーンと口調を指定します（空欄で省略可）',
  },
  {
    key: 'prompt',
    label: 'プロンプト',
    defaultValue: 'ターゲット層の悩みや関心に寄り添った、思わず共感や保存をしたくなる自然な日本語のThreadsポストを作成してください。ハッシュタグや絵文字の多用は控えてください。',
    description: '内容の作り方・文字数目安など具体的な作成指示を記述してください（改行を保持したままAIへ送信されます）',
  },
  {
    key: 'samples',
    label: 'お手本サンプル',
    defaultValue:
      '例1: 朝のたった5分で今日やることを3つだけ書き出すようにしたら、1日の焦りが驚くほど減りました。全部完璧にやろうとしない方が案外うまくいきますよね✨\n例2: 「あとでやろう」はだいたい忘れるので、2分で終わることはその場で片付けるようにしています。小さなストレスが溜まらなくておすすめです。',
    description: 'このアカウントらしい理想の投稿例を2〜3個記述してください。AIがトーンと距離感の参考にします（空欄で省略可）',
  },
  {
    key: 'ngRules',
    label: 'NGルール',
    defaultValue:
      '・「素晴らしいですね！」「応援しています！」などの定型的なお世辞やAIっぽい表現は禁止\n・宣伝・営業・勧誘・過度なハッシュタグ（#）は禁止',
    description: '禁止ワード・AIっぽい定型句の禁止・避ける話題など、アカウント別のNGルールを記述してください（空欄で省略可）',
  },
  {
    key: 'outputRule',
    label: '出力ルール',
    defaultValue:
      '※出力は投稿するメッセージ本文のみ（140〜400文字以内）で返してください。挨拶や前置き（例：「了解しました」「投稿案:」）や引用符などの余計な注釈は一切含めないでください。',
    description: '本文のみ出力・前置きや「」禁止・文字数上限などの最終整形ルールです（空欄で省略可）',
  },
];

export function createDefaultAiPromptSections(
  overrides?: Partial<AiPromptSections>
): AiPromptSections {
  return {
    character: AI_PROMPT_SECTION_META[0].defaultValue,
    target: AI_PROMPT_SECTION_META[1].defaultValue,
    purpose: AI_PROMPT_SECTION_META[2].defaultValue,
    toneRule: AI_PROMPT_SECTION_META[3].defaultValue,
    prompt: AI_PROMPT_SECTION_META[4].defaultValue,
    samples: AI_PROMPT_SECTION_META[5].defaultValue,
    ngRules: AI_PROMPT_SECTION_META[6].defaultValue,
    outputRule: AI_PROMPT_SECTION_META[7].defaultValue,
    ...(overrides || {}),
  };
}

export const AI_LOGIN_FREE_ENGINES: AiEngineId[] = [
  'duckai',
  'talkai',
  'deepai',
  'venice',
];

export const DEFAULT_AI_FAILOVER_ORDER: AiEngineId[] = [
  'duckai',
  'talkai',
  'deepai',
  'venice',
  'groq',
  'gemini',
  'chatgpt',
  'grok',
  'claude',
  'copilot',
  'perplexity',
  'deepseek',
  'qwen',
  'mistral',
  'openrouter',
  'ollama',
  'lmstudio',
  'chrome_builtin',
];

const VALID_AI_ENGINE_SET = new Set<string>(AI_ENGINES_META.map((m) => m.id));

export function parseCommaTokens(raw?: string): string[] {
  if (!raw) return [];
  return String(raw)
    .split(/[,，、;；]+/)
    .map((t) => t.trim())
    .filter(Boolean);
}

export function parseAiEngineList(raw?: string): AiEngineId[] {
  if (!raw) return ['duckai'];
  const tokens = String(raw)
    .split(/[,，、;；/\s]+/)
    .map((t) => t.trim().toLowerCase())
    .filter(Boolean);
  const res: AiEngineId[] = [];
  for (const tok of tokens) {
    if (VALID_AI_ENGINE_SET.has(tok) && !res.includes(tok as AiEngineId)) {
      res.push(tok as AiEngineId);
    }
  }
  return res.length > 0 ? res : ['duckai'];
}

export function normalizeAiEngineList(raw?: string): string {
  return parseAiEngineList(raw).join(', ');
}

export function getAiFailoverChainPreview(
  engineRaw?: string,
  fallbackRaw?: string
): AiEngineId[] {
  const primaries = parseAiEngineList(engineRaw);
  const fbTrim = String(fallbackRaw || '').trim();
  if (['off', 'none', 'no', '0', 'false', 'なし', '無効', '切替なし'].includes(fbTrim.toLowerCase())) {
    return primaries;
  }
  const chain = [...primaries];
  if (fbTrim) {
    const fbTokens = fbTrim
      .split(/[,，、;；/\s]+/)
      .map((t) => t.trim().toLowerCase())
      .filter(Boolean);
    for (const tok of fbTokens) {
      if (VALID_AI_ENGINE_SET.has(tok) && !chain.includes(tok as AiEngineId)) {
        chain.push(tok as AiEngineId);
      }
    }
    return chain;
  }
  for (const cand of DEFAULT_AI_FAILOVER_ORDER) {
    if (!chain.includes(cand)) chain.push(cand);
  }
  return chain;
}

export function createDefaultAiPostConfig(overrides?: Partial<AiPostConfig>): AiPostConfig {
  return {
    enabled: false,
    activeHours: '0:00-0:00',
    weekdays: '毎日',
    intervalStr: '120〜360分',
    memo: '例: フレンドリーで共感されやすい口調・絵文字1〜2個・ハッシュタグは1個まで（カンマ区切りでエンジン・モデル・API複数指定＆失敗時自動切替に対応）',
    engine: 'duckai',
    modelName: '',
    fallbackCandidates: '',
    aiWindowSync: true,
    apiKeyOrUrl: '',
    promptTemplate: '',
    topic: '',
    deleteAfter: '',
    ...(overrides || {}),
    promptSections: createDefaultAiPromptSections(overrides?.promptSections),
  };
}




