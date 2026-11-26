export interface AccountDesignProfile {
  genre?: string;
  target?: string;
  tone?: 'polite' | 'friendly' | 'expert' | 'empathetic';
  painPoint?: string;
  benefitPoints?: string;
  cta?: string;
  url?: string;
  topic?: string;
  imageStyle?:
    | 'auto'
    | 'photo_overlay'
    | 'magazine_clean'
    | 'chat_bubble'
    | 'neon_night'
    | 'step_roadmap'
    | 'qa_card'
    | 'diagram_card'
    | 'checklist'
    | 'mix'
    | 'all_mix';
}

export type AiEngineId =
  | 'groq'
  | 'gemini'
  | 'duckai'
  | 'venice'
  | 'deepai'
  | 'talkai'
  | 'chatgpt'
  | 'qwen'
  | 'mistral'
  | 'grok'
  | 'deepseek'
  | 'claude'
  | 'perplexity'
  | 'copilot'
  | 'openrouter'
  | 'ollama'
  | 'lmstudio'
  | 'chrome_builtin';

export interface AiPromptSections {
  character: string; // キャラ設定
  target: string; // ターゲット層
  usp: string; // 自店の強み・アピールポイント（USP）
  theme: string; // 投稿テーマ（切り口）
  cta: string; // 具体的なCTA（行動喚起）
  toneRule: string; // 口調・文体ルール
  prompt: string; // プロンプト
  samples: string; // お手本サンプル
  ngRules: string; // NGルール
  outputRule: string; // 出力ルール
}

export interface AiPostConfig {
  enabled: boolean; // 有効 / 無効
  activeHours: string; // 実行時間 (例: 0:00-0:00)
  weekdays: string; // 曜日 (例: 毎日)
  intervalStr: string; // ポスト間隔 (例: 120〜360分)
  memo: string; // メモ (AI生成方針・キャラ設定・補足メモ)
  engine: string; // AIエンジン (単体 または カンマ区切りで複数指定可 例: duckai, talkai, deepai)
  modelName: string; // AIモデル名(空欄=既定モデル自動 / カンマ区切り・engine:model形式可)
  fallbackCandidates: string; // AIフォールバック候補(空欄=自動切替 / off=切替なし / カンマ区切り複数指定可)
  aiWindowSync: boolean; // AI専用ウィンドウ連携(1=ON/0=OFF)
  apiKeyOrUrl: string; // AI APIキー / 接続URL (カンマ区切り・engine:key形式可)
  promptSections?: AiPromptSections; // 10項目構造化プロンプト (【】付送信)
  promptTemplate?: string; // 指示プロンプト
  topic?: string; // デフォルトトピック
  deleteAfter?: string; // 削除時間 (ランダム自動削除)
}

export interface AiPostHistoryItem {
  id: number;
  prompt: string;
  generatedText?: string;
  topic?: string;
  deleteAfter?: string;
  engineUsed?: string;
  status: '指示待機' | '投稿完了' | '投稿完了(削除済)' | 'エラー';
  postedAt?: string;
  errorMessage?: string;
}

export interface AccountInfo {
  id: string;
  authToken: string;
  userId: string;
  password?: string; // ID/PW自動ログイン用
  notes: string;
  minIntervalMinutes: number;
  maxIntervalMinutes: number;
  activeHours: string;
  mode: 0 | 1; // 0: 上から順, 1: ランダム
  profileDir: string;
  topics: string[];
  currentTopicIndex: number;
  enabled: boolean; // 有効 / 無効 (電話番号認証時等に自動で無効・停止)
  checkpointReason?: string; // 停止理由
  proxy?: string; // プロキシ設定 (例: http://user:pass@ip:port / ip:port:user:pass)
  design?: AccountDesignProfile; // アカウント別の自動生成用設計データ
  aiPostConfig?: AiPostConfig; // 固定ポストとは完全別枠のAIポスト設定
  aiPostHistory?: AiPostHistoryItem[]; // AIポスト指示・生成履歴
}

export interface PostItem {
  id: number;
  text: string;
  theme?: string;
  themeTimeSlot?: string;
  themeWeekdays?: string;
  timeSlot?: string;
  weekdays?: string;
  deleteAfter?: string;
  intervalStr?: string;
  mode?: 0 | 1;
  topic?: string;
  imagePath?: string;
  status: '未投稿' | '投稿中' | '投稿完了' | '投稿完了(削除済)' | 'エラー';
  completedAt?: string;
  errorMessage?: string;
}

export interface DailyGreetingSlot {
  id: string;
  slotName: string; // '朝枠' | '昼枠' | '夜枠'
  timeRange: string;
  text: string;
  enabled: boolean;
  lastPostedDate?: string;
  lastPostedTime?: string;
}

export interface SystemSettings {
  autoLoop: boolean;
  singleWindowRotation: boolean;
  blockMedia: boolean;
  stealthPatrol: boolean;
  autoCleanProfile: boolean;
  topicAbTest: boolean;
  dailyGreeting: boolean;
  speedMode: 'real' | 'fast' | 'instant';
}

export interface LogMessage {
  id: string;
  timestamp: string;
  level: 'INFO' | 'SUCCESS' | 'WARN' | 'ERROR' | 'STEALTH';
  message: string;
}
