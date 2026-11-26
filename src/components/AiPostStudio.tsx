import React, { useState, useEffect, useMemo } from 'react';
import {
  AccountInfo,
  AiPostConfig,
  AiPostHistoryItem,
  AiPromptSections,
} from '../types';
import {
  AI_ENGINES_META,
  AI_LOGIN_FREE_ENGINES,
  AI_PROMPT_SECTION_META,
  FIXED_AI_SYSTEM_PROMPT,
  createDefaultAiPostConfig,
  createDefaultAiPromptSections,
  parseAiEngineList,
  getAiFailoverChainPreview,
} from '../utils/spintax';
import {
  Bot,
  Power,
  Clock,
  CalendarDays,
  Timer,
  Cpu,
  KeyRound,
  AppWindow,
  Layers,
  Plus,
  Trash2,
  FileSpreadsheet,
  Save,
  Sparkles,
  Eye,
  EyeOff,
  CheckCircle2,
  AlertCircle,
  Lock,
  RotateCcw,
  FileText,
} from 'lucide-react';

interface AiPostStudioProps {
  currentAccount: AccountInfo;
  onUpdateAccountAiConfig: (
    accountId: string,
    config: AiPostConfig,
    history: AiPostHistoryItem[]
  ) => void;
  onExportAccountExcel: () => void;
  onLog: (level: 'INFO' | 'SUCCESS' | 'WARN' | 'ERROR', message: string) => void;
}

const SECTION_BADGE_STYLES: Record<keyof AiPromptSections, string> = {
  character: 'bg-blue-950/70 text-blue-200 border-blue-700/70',
  target: 'bg-sky-950/70 text-sky-200 border-sky-700/70',
  usp: 'bg-teal-950/70 text-teal-200 border-teal-700/70',
  theme: 'bg-indigo-950/70 text-indigo-200 border-indigo-700/70',
  cta: 'bg-emerald-950/70 text-emerald-200 border-emerald-700/70',
  toneRule: 'bg-amber-950/70 text-amber-200 border-amber-700/70',
  prompt: 'bg-violet-950/70 text-violet-200 border-violet-700/70',
  samples: 'bg-pink-950/70 text-pink-200 border-pink-700/70',
  ngRules: 'bg-rose-950/70 text-rose-200 border-rose-700/70',
  outputRule: 'bg-slate-800/90 text-slate-200 border-slate-600/70',
};

export const AiPostStudio: React.FC<AiPostStudioProps> = ({
  currentAccount,
  onUpdateAccountAiConfig,
  onExportAccountExcel,
  onLog,
}) => {
  const [config, setConfig] = useState<AiPostConfig>(() =>
    createDefaultAiPostConfig(currentAccount.aiPostConfig)
  );
  const [history, setHistory] = useState<AiPostHistoryItem[]>(
    () => currentAccount.aiPostHistory || []
  );
  const [showKey, setShowKey] = useState(false);
  const [showCompiledPreview, setShowCompiledPreview] = useState(false);
  const [newPrompt, setNewPrompt] = useState('');
  const [newPromptTopic, setNewPromptTopic] = useState('');
  const [newPromptDeleteAfter, setNewPromptDeleteAfter] = useState('');
  const [savedNotice, setSavedNotice] = useState(false);

  useEffect(() => {
    setConfig(createDefaultAiPostConfig(currentAccount.aiPostConfig));
    setHistory(currentAccount.aiPostHistory || []);
    setSavedNotice(false);
  }, [currentAccount.id, currentAccount.aiPostConfig, currentAccount.aiPostHistory]);

  const sections: AiPromptSections = useMemo(
    () => createDefaultAiPromptSections(config.promptSections),
    [config.promptSections]
  );

  const selectedEngines = parseAiEngineList(config.engine);
  const engineMeta =
    AI_ENGINES_META.find((e) => e.id === selectedEngines[0]) ||
    AI_ENGINES_META.find((e) => e.id === 'duckai') ||
    AI_ENGINES_META[0];
  const failoverChain = getAiFailoverChainPreview(
    config.engine,
    config.fallbackCandidates
  );

  const loginFreeEngines = AI_ENGINES_META.filter((e) =>
    AI_LOGIN_FREE_ENGINES.includes(e.id)
  );
  const otherEngines = AI_ENGINES_META.filter(
    (e) => !AI_LOGIN_FREE_ENGINES.includes(e.id)
  );

  const compiledPromptPreview = useMemo(() => {
    const blocks: string[] = [FIXED_AI_SYSTEM_PROMPT];
    AI_PROMPT_SECTION_META.forEach((sec) => {
      const val = (sections[sec.key] || '').trim();
      if (val) {
        blocks.push(`【${sec.label}】\n${val}`);
      }
    });
    if ((config.memo || '').trim()) {
      blocks.push(`【メモ】\n${config.memo.trim()}`);
    }
    return blocks.join('\n\n');
  }, [sections, config.memo]);

  const toggleEngineInList = (engId: string) => {
    const current = parseAiEngineList(config.engine);
    const exists = current.includes(engId as (typeof current)[number]);
    let nextList: string[];
    if (exists) {
      nextList = current.filter((x) => x !== engId);
      if (nextList.length === 0) nextList = [engId];
    } else {
      nextList = [...current, engId];
    }
    setConfig({ ...config, engine: nextList.join(', ') });
  };

  const updateSectionValue = (key: keyof AiPromptSections, val: string) => {
    const nextSections: AiPromptSections = { ...sections, [key]: val };
    setConfig({
      ...config,
      promptSections: nextSections,
      ...(key === 'prompt' ? { promptTemplate: val } : {}),
    });
  };

  const handleResetSections = () => {
    const def = createDefaultAiPromptSections();
    setConfig({
      ...config,
      promptSections: def,
      promptTemplate: def.prompt,
    });
  };

  const handleSave = (nextCfg = config, nextHist = history) => {
    const normalizedCfg = createDefaultAiPostConfig(nextCfg);
    onUpdateAccountAiConfig(currentAccount.id, normalizedCfg, nextHist);
    setSavedNotice(true);
    setTimeout(() => setSavedNotice(false), 2500);
    onLog(
      'SUCCESS',
      `[@${currentAccount.id}] 🤖 AIポスト設定を保存しました (エンジン: ${normalizedCfg.engine} / 専用Window連携: ${
        normalizedCfg.aiWindowSync ? '1 (固定_ai_window)' : '0 (OFF)'
      })`
    );
  };

  const handleAddPromptRow = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newPrompt.trim()) return;
    const nextId =
      history.length > 0 ? Math.max(...history.map((h) => h.id || 0)) + 1 : 1;
    const added: AiPostHistoryItem = {
      id: nextId,
      prompt: newPrompt.trim(),
      topic: newPromptTopic.trim() || undefined,
      deleteAfter: newPromptDeleteAfter.trim() || undefined,
      status: '指示待機',
    };
    const updatedHist = [...history, added];
    setHistory(updatedHist);
    setNewPrompt('');
    setNewPromptTopic('');
    setNewPromptDeleteAfter('');
    handleSave(config, updatedHist);
  };

  const handleDeleteRow = (id: number) => {
    const updatedHist = history.filter((h) => h.id !== id);
    setHistory(updatedHist);
    handleSave(config, updatedHist);
  };

  return (
    <div className="space-y-4">
      {/* Architecture Banner: 完全分離 & Window分離 */}
      <div className="bg-gradient-to-r from-indigo-950/90 via-zinc-900 to-zinc-900 border border-indigo-800/70 rounded-xl p-4 shadow-sm">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div className="space-y-1.5">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="inline-flex items-center gap-1.5 text-sm font-bold text-white">
                <Bot className="w-4 h-4 text-indigo-400" />
                @{currentAccount.id} 専用 AIポスト設定
              </span>
              <span className="px-2 py-0.5 rounded bg-indigo-900/70 text-indigo-200 border border-indigo-700 text-[11px] font-semibold">
                固定ポストとは完全別枠 (AIポスト シート)
              </span>
              <span
                className={`px-2 py-0.5 rounded text-[11px] font-bold border ${
                  config.enabled
                    ? 'bg-emerald-950/80 text-emerald-300 border-emerald-700'
                    : 'bg-zinc-950 text-zinc-400 border-zinc-700'
                }`}
              >
                {config.enabled ? '● 有効 (独立スケジュール稼働)' : '○ 無効 (停止中)'}
              </span>
            </div>
            <p className="text-xs text-zinc-300 leading-relaxed">
              固定ポスト（基本・各テーマシート）とは独立したスケジュール＆10項目【】プロンプトで動作します。
              <span className="text-indigo-300 font-semibold">
                「AI専用ウィンドウ連携 = 1 (ON)」
              </span>
              のときAI操作は固定プロファイル（
              <code className="text-amber-300 font-mono">
                chrome_profiles/_ai_window
              </code>
              ）のみを使い、各Threadsアカウント窓（
              <code className="text-emerald-300 font-mono">
                chrome_profiles/@アカウント名
              </code>
              ）とは完全分離で動作します。
            </p>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <button
              onClick={() => handleSave()}
              className="flex items-center gap-1.5 px-3 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold shadow cursor-pointer"
            >
              <Save className="w-3.5 h-3.5" />
              <span>AI設定を保存</span>
            </button>
            <button
              onClick={onExportAccountExcel}
              className="flex items-center gap-1.5 px-3 py-2 rounded-lg bg-emerald-700 hover:bg-emerald-600 text-white text-xs font-bold shadow cursor-pointer"
              title="現在の固定ポストシート＋AIポストシートを1つのExcelに出力"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>{currentAccount.id}.xlsx 書出</span>
            </button>
          </div>
        </div>
        {savedNotice && (
          <div className="mt-2.5 flex items-center gap-1.5 text-xs text-emerald-300 font-semibold">
            <CheckCircle2 className="w-3.5 h-3.5" />
            <span>@{currentAccount.id} の AIポスト設定を保存しました。Excel書出でWindows版にも即反映できます。</span>
          </div>
        )}
      </div>

      {/* Fixed System Prompt Display (先頭固定プロンプト) */}
      <div className="bg-zinc-900 border border-violet-900/60 rounded-xl p-4 space-y-2">
        <div className="flex items-center justify-between flex-wrap gap-2">
          <div className="flex items-center gap-1.5 text-xs font-bold text-violet-300">
            <Lock className="w-3.5 h-3.5" />
            <span>固定プロンプト（AI生成時に必ず先頭へ送信）</span>
          </div>
          <button
            type="button"
            onClick={() => setShowCompiledPreview((v) => !v)}
            className="flex items-center gap-1 px-2.5 py-1 rounded bg-zinc-800 hover:bg-zinc-700 text-[11px] text-zinc-200 border border-zinc-700 cursor-pointer"
          >
            <FileText className="w-3 h-3" />
            <span>{showCompiledPreview ? '送信全体プレビューを閉じる' : 'AIへ実際に送るプロンプト全体を見る'}</span>
          </button>
        </div>
        <pre className="text-[11px] leading-relaxed text-zinc-200 bg-zinc-950 border border-zinc-800 rounded-lg p-3 whitespace-pre-wrap font-sans">
          {FIXED_AI_SYSTEM_PROMPT}
        </pre>
        {showCompiledPreview && (
          <div className="pt-1">
            <div className="text-[11px] font-semibold text-amber-300 mb-1">
              送信全体プレビュー（固定プロンプト ＋ 下の非空【】項目）:
            </div>
            <pre className="text-[11px] leading-relaxed text-emerald-200/90 bg-black/70 border border-emerald-900/60 rounded-lg p-3 whitespace-pre-wrap font-sans max-h-72 overflow-y-auto">
              {compiledPromptPreview}
            </pre>
          </div>
        )}
      </div>

      {/* Top 2 Rows matching Excel A1:E4 */}
      <div className="bg-zinc-900 border border-zinc-800 rounded-xl p-4 space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-3 pb-3 border-b border-zinc-800">
          <div className="flex items-center gap-2 text-xs font-bold text-zinc-200">
            <Cpu className="w-4 h-4 text-indigo-400" />
            <span>AIポスト シート上段設定（1〜2行目：スケジュール / 3〜4行目：AIエンジン）</span>
          </div>

          <button
            type="button"
            onClick={() => {
              const next = { ...config, enabled: !config.enabled };
              setConfig(next);
              handleSave(next, history);
            }}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold border transition-all cursor-pointer ${
              config.enabled
                ? 'bg-emerald-950/80 text-emerald-300 border-emerald-700 hover:bg-emerald-900/70'
                : 'bg-rose-950/70 text-rose-300 border-rose-800 hover:bg-rose-900/60'
            }`}
          >
            <Power className="w-3.5 h-3.5" />
            <span>有効/無効: {config.enabled ? '有効' : '無効'}</span>
          </button>
        </div>

        {/* Excel Row 1-2: 実行時間 | 曜日 | ポスト間隔 | 有効/無効 | メモ */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-3 text-xs">
          <div>
            <label className="block text-zinc-400 font-semibold mb-1 flex items-center gap-1">
              <Clock className="w-3.5 h-3.5 text-sky-400" />
              実行時間
            </label>
            <input
              type="text"
              value={config.activeHours}
              onChange={(e) =>
                setConfig({ ...config, activeHours: e.target.value })
              }
              placeholder="0:00-0:00"
              className="w-full px-3 py-2 bg-zinc-950 border border-zinc-800 rounded-lg text-zinc-100 font-mono focus:outline-none focus:border-indigo-500"
            />
            <div className="flex flex-wrap gap-1 mt-1">
              {['0:00-0:00', '7:00-10:00', '19:00-23:30'].map((slot) => (
                <button
                  key={slot}
                  type="button"
                  onClick={() => setConfig({ ...config, activeHours: slot })}
                  className="px-1.5 py-0.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-300 rounded text-[10px] cursor-pointer"
                >
                  {slot === '0:00-0:00' ? '終日' : slot}
                </button>
              ))}
            </div>
          </div>

          <div>
            <label className="block text-zinc-400 font-semibold mb-1 flex items-center gap-1">
              <CalendarDays className="w-3.5 h-3.5 text-amber-400" />
              曜日
            </label>
            <input
              type="text"
              value={config.weekdays}
              onChange={(e) =>
                setConfig({ ...config, weekdays: e.target.value })
              }
              placeholder="毎日"
              className="w-full px-3 py-2 bg-zinc-950 border border-zinc-800 rounded-lg text-zinc-100 font-mono focus:outline-none focus:border-indigo-500"
            />
            <div className="flex flex-wrap gap-1 mt-1">
              {['毎日', '月～金', '土日'].map((wd) => (
                <button
                  key={wd}
                  type="button"
                  onClick={() => setConfig({ ...config, weekdays: wd })}
                  className="px-1.5 py-0.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-300 rounded text-[10px] cursor-pointer"
                >
                  {wd}
                </button>
              ))}
            </div>
          </div>

          <div>
            <label className="block text-zinc-400 font-semibold mb-1 flex items-center gap-1">
              <Timer className="w-3.5 h-3.5 text-emerald-400" />
              ポスト間隔
            </label>
            <input
              type="text"
              value={config.intervalStr}
              onChange={(e) =>
                setConfig({ ...config, intervalStr: e.target.value })
              }
              placeholder="120〜360分"
              className="w-full px-3 py-2 bg-zinc-950 border border-zinc-800 rounded-lg text-zinc-100 font-mono focus:outline-none focus:border-indigo-500"
            />
            <div className="flex flex-wrap gap-1 mt-1">
              {['30〜60分', '60〜120分', '120〜360分'].map((iv) => (
                <button
                  key={iv}
                  type="button"
                  onClick={() => setConfig({ ...config, intervalStr: iv })}
                  className="px-1.5 py-0.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-300 rounded text-[10px] cursor-pointer"
                >
                  {iv}
                </button>
              ))}
            </div>
          </div>

          <div>
            <label className="block text-zinc-400 font-semibold mb-1">
              有効/無効
            </label>
            <select
              value={config.enabled ? '有効' : '無効'}
              onChange={(e) =>
                setConfig({ ...config, enabled: e.target.value === '有効' })
              }
              className="w-full px-3 py-2 bg-zinc-950 border border-zinc-800 rounded-lg text-zinc-100 focus:outline-none focus:border-indigo-500 cursor-pointer"
            >
              <option value="有効">有効</option>
              <option value="無効">無効</option>
            </select>
            <p className="text-[10px] text-zinc-500 mt-1">
              無効時は固定ポストのみ稼働
            </p>
          </div>

          <div>
            <label className="block text-zinc-400 font-semibold mb-1">
              メモ
            </label>
            <input
              type="text"
              value={config.memo}
              onChange={(e) => setConfig({ ...config, memo: e.target.value })}
              placeholder="補足・運用メモ"
              className="w-full px-3 py-2 bg-zinc-950 border border-zinc-800 rounded-lg text-zinc-100 focus:outline-none focus:border-indigo-500"
            />
          </div>
        </div>

        {/* Excel Row 3-6: AI Engine selector + login-free priority + comma multi-select */}
        <div className="pt-3 border-t border-zinc-800/80 space-y-3 text-xs">
          <div className="p-3 rounded-lg bg-zinc-950 border border-zinc-800 space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <label className="text-indigo-300 font-semibold flex items-center gap-1.5">
                <Cpu className="w-3.5 h-3.5" />
                AIエンジン（カンマ「,」区切りで複数選択＆失敗時はログイン不要サイト優先で自動切替）
              </label>
              <span className="text-[11px] text-zinc-400">
                選択順:{' '}
                <span className="text-indigo-300 font-mono">
                  {selectedEngines.join(' ➔ ')}
                </span>
              </span>
            </div>

            <input
              type="text"
              value={config.engine}
              onChange={(e) => setConfig({ ...config, engine: e.target.value })}
              placeholder="例: duckai, talkai, deepai (カンマ区切りで複数指定可)"
              className="w-full px-3 py-2 bg-zinc-900 border border-zinc-700 rounded-lg text-white font-mono focus:outline-none focus:border-indigo-500"
            />

            <div className="space-y-1.5">
              <div className="flex flex-wrap items-center gap-1.5">
                <span className="text-[10px] font-bold text-emerald-300 px-1.5 py-0.5 rounded bg-emerald-950/60 border border-emerald-800/70">
                  ログイン不要(優先)
                </span>
                {loginFreeEngines.map((eng) => {
                  const orderIdx = selectedEngines.indexOf(eng.id);
                  const isSelected = orderIdx !== -1;
                  return (
                    <button
                      key={eng.id}
                      type="button"
                      onClick={() => toggleEngineInList(eng.id)}
                      title={`${eng.label}｜既定: ${eng.defaultModel}｜${eng.modeNote}`}
                      className={`px-2 py-1 rounded-md border text-[11px] font-mono transition-colors cursor-pointer ${
                        isSelected
                          ? 'bg-emerald-600/90 border-emerald-400 text-white font-bold'
                          : 'bg-zinc-900 border-zinc-800 text-emerald-200 hover:bg-zinc-800 hover:border-zinc-700'
                      }`}
                    >
                      {isSelected ? `${orderIdx + 1}. ` : ''}
                      {eng.id}
                    </button>
                  );
                })}
              </div>

              <div className="flex flex-wrap items-center gap-1.5">
                <span className="text-[10px] font-bold text-indigo-300 px-1.5 py-0.5 rounded bg-indigo-950/60 border border-indigo-800/70">
                  API・要ログイン(14種)
                </span>
                {otherEngines.map((eng) => {
                  const orderIdx = selectedEngines.indexOf(eng.id);
                  const isSelected = orderIdx !== -1;
                  return (
                    <button
                      key={eng.id}
                      type="button"
                      onClick={() => toggleEngineInList(eng.id)}
                      title={`${eng.label}｜既定: ${eng.defaultModel}｜${eng.modeNote}`}
                      className={`px-2 py-1 rounded-md border text-[11px] font-mono transition-colors cursor-pointer ${
                        isSelected
                          ? 'bg-indigo-600/90 border-indigo-400 text-white font-bold'
                          : 'bg-zinc-900 border-zinc-800 text-zinc-300 hover:bg-zinc-800 hover:border-zinc-700'
                      }`}
                    >
                      {isSelected ? `${orderIdx + 1}. ` : ''}
                      {eng.id}
                    </button>
                  );
                })}
              </div>
            </div>

            <div className="p-2 rounded border border-dashed border-indigo-700/70 bg-indigo-950/20 text-[11px] text-zinc-300 leading-relaxed">
              先頭から順に試行し、失敗時は自動で次のAIへ切り替わります（自動切替優先順:{' '}
              <span className="text-emerald-300 font-mono">
                {failoverChain.slice(0, 8).join(' ➔ ')}
                {failoverChain.length > 8 ? ` …他${failoverChain.length - 8}種` : ''}
              </span>
              ）。モデル名・APIキー/URLも「,」区切りで複数指定できます。
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-4 gap-3">
            <div>
              <label className="block text-zinc-300 font-semibold mb-1">
                AIモデル名 (空欄=既定自動 / 「,」可)
              </label>
              <input
                type="text"
                value={config.modelName}
                onChange={(e) =>
                  setConfig({ ...config, modelName: e.target.value })
                }
                placeholder={`空欄で自動 (${engineMeta.defaultModel})`}
                className="w-full px-3 py-2 bg-zinc-950 border border-zinc-800 rounded-lg text-zinc-100 font-mono focus:outline-none focus:border-indigo-500"
              />
              <p className="text-[10px] text-zinc-500 mt-1">
                例: groq:llama-3.3-70b-versatile, gemini:gemini-2.5-flash
              </p>
            </div>

            <div>
              <label className="block text-zinc-300 font-semibold mb-1 flex items-center gap-1">
                <Layers className="w-3.5 h-3.5 text-amber-400" />
                AIフォールバック候補
              </label>
              <input
                type="text"
                value={config.fallbackCandidates}
                onChange={(e) =>
                  setConfig({ ...config, fallbackCandidates: e.target.value })
                }
                placeholder="空欄=自動切替 / off=切替なし"
                className="w-full px-3 py-2 bg-zinc-950 border border-zinc-800 rounded-lg text-zinc-100 font-mono focus:outline-none focus:border-indigo-500"
              />
              <div className="flex flex-wrap gap-1 mt-1">
                <button
                  type="button"
                  onClick={() => setConfig({ ...config, fallbackCandidates: '' })}
                  className="px-1.5 py-0.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-300 rounded text-[10px] cursor-pointer"
                >
                  自動切替(ログイン不要優先)
                </button>
                <button
                  type="button"
                  onClick={() => setConfig({ ...config, fallbackCandidates: 'off' })}
                  className="px-1.5 py-0.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-300 rounded text-[10px] cursor-pointer"
                >
                  off
                </button>
              </div>
            </div>

            <div>
              <label className="block text-indigo-300 font-semibold mb-1 flex items-center gap-1">
                <AppWindow className="w-3.5 h-3.5" />
                AI専用ウィンドウ連携 (1=ON/0=OFF)
              </label>
              <select
                value={config.aiWindowSync ? '1' : '0'}
                onChange={(e) =>
                  setConfig({ ...config, aiWindowSync: e.target.value === '1' })
                }
                className="w-full px-3 py-2 bg-zinc-950 border border-zinc-800 rounded-lg text-white font-semibold focus:outline-none focus:border-indigo-500 cursor-pointer"
              >
                <option value="1">1: ON (_ai_window使用)</option>
                <option value="0">0: OFF (API/URLのみ)</option>
              </select>
              <p className="text-[10px] text-zinc-500 mt-1">
                duckai等Webチャットは 1 (ON) で自動動作
              </p>
            </div>

            <div>
              <label className="block text-zinc-300 font-semibold mb-1 flex items-center gap-1">
                <KeyRound className="w-3.5 h-3.5 text-emerald-400" />
                AI APIキー / 接続URL (「,」複数可)
              </label>
              <div className="flex items-center gap-1.5">
                <input
                  type={showKey ? 'text' : 'password'}
                  value={config.apiKeyOrUrl}
                  onChange={(e) =>
                    setConfig({ ...config, apiKeyOrUrl: e.target.value })
                  }
                  placeholder="空欄可 / groq:gsk_xxx, gemini:AIza..."
                  className="flex-1 px-3 py-2 bg-zinc-950 border border-zinc-800 rounded-lg text-zinc-100 font-mono focus:outline-none focus:border-indigo-500"
                />
                <button
                  type="button"
                  onClick={() => setShowKey(!showKey)}
                  className="p-2 bg-zinc-800 hover:bg-zinc-700 text-zinc-300 rounded-lg border border-zinc-700 cursor-pointer"
                  title={showKey ? '隠す' : '表示'}
                >
                  {showKey ? (
                    <EyeOff className="w-4 h-4" />
                  ) : (
                    <Eye className="w-4 h-4" />
                  )}
                </button>
              </div>
              <p className="text-[10px] text-zinc-500 mt-1">
                「URL|APIキー」や「,」複数キー指定に対応
              </p>
            </div>
          </div>
        </div>
      </div>

      {/* 10-Row Structured Prompt Table matching Excel Rows 7-17 */}
      <div className="bg-zinc-900 border border-zinc-800 rounded-xl overflow-hidden">
        <div className="flex items-center justify-between flex-wrap gap-2 px-4 py-3 bg-indigo-950/80 border-b border-indigo-900">
          <div>
            <h4 className="text-xs font-bold text-white flex items-center gap-1.5">
              <Sparkles className="w-3.5 h-3.5 text-amber-400" />
              プロンプト項目 (【】付送信) — AIポストシート 7〜17行目と完全同期
            </h4>
            <p className="text-[11px] text-indigo-200/80 mt-0.5">
              入力されている項目だけが「【項目名】内容」として固定プロンプトの下へ順番に送信されます（空欄項目は省略）。
            </p>
          </div>
          <button
            type="button"
            onClick={handleResetSections}
            className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-zinc-900/90 hover:bg-zinc-800 text-zinc-200 border border-zinc-700 text-[11px] font-semibold cursor-pointer"
          >
            <RotateCcw className="w-3 h-3" />
            <span>標準テンプレートに戻す</span>
          </button>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-indigo-950/60 text-indigo-100 border-b border-zinc-800">
                <th className="py-2.5 px-3 w-44 font-bold">
                  プロンプト項目 (【】付送信)
                </th>
                <th className="py-2.5 px-3 font-bold">
                  AIへ送る設定内容 (ここを編集)
                </th>
                <th className="py-2.5 px-3 w-80 font-bold hidden lg:table-cell">
                  各項目の説明
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-zinc-800">
              {AI_PROMPT_SECTION_META.map((sec) => {
                const val = sections[sec.key] ?? '';
                const lineCount = Math.max(2, val.split('\n').length + (val.length > 75 ? 1 : 0));
                return (
                  <tr key={sec.key} className="hover:bg-zinc-800/30 align-top">
                    <td className="py-3 px-3">
                      <span
                        className={`inline-block px-2.5 py-1 rounded-md border text-[11px] font-bold ${SECTION_BADGE_STYLES[sec.key]}`}
                      >
                        {sec.label}
                      </span>
                      <div className="text-[10px] text-zinc-500 mt-1.5 lg:hidden">
                        {sec.description}
                      </div>
                    </td>
                    <td className="py-2.5 px-3">
                      <textarea
                        rows={Math.min(6, lineCount)}
                        value={val}
                        onChange={(e) =>
                          updateSectionValue(sec.key, e.target.value)
                        }
                        placeholder={`【${sec.label}】を設定（空欄なら送信時に省略）`}
                        className="w-full px-3 py-2 bg-zinc-950 border border-zinc-800 rounded-lg text-zinc-100 text-xs leading-relaxed focus:outline-none focus:border-indigo-500 resize-y"
                      />
                    </td>
                    <td className="py-3 px-3 text-[11px] text-zinc-400 leading-relaxed hidden lg:table-cell">
                      {sec.description}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* Optional Spot Instruction & Generation History (Excel Row 19+) */}
      <div className="bg-zinc-900 border border-zinc-800 rounded-xl p-4 space-y-3">
        <div className="flex items-center justify-between flex-wrap gap-2">
          <div>
            <h4 className="text-xs font-bold text-white flex items-center gap-1.5">
              <Sparkles className="w-3.5 h-3.5 text-emerald-400" />
              生成ポスト履歴 ＆ 個別指示（AIポストシート 19行目〜）
            </h4>
            <p className="text-[11px] text-zinc-400 mt-0.5">
              上の10項目が設定されていれば追加指示なしでも自動生成＆投稿でき、生成された文章と使用エンジンはここに自動記録されます。
            </p>
          </div>
        </div>

        <form
          onSubmit={handleAddPromptRow}
          className="grid grid-cols-1 md:grid-cols-12 gap-2 text-xs bg-zinc-950/80 p-3 rounded-lg border border-zinc-800"
        >
          <input
            type="text"
            value={newPrompt}
            onChange={(e) => setNewPrompt(e.target.value)}
            placeholder="個別追加テーマ / 指示 (任意・空のままでも上の10項目で自動稼働します)"
            className="md:col-span-6 px-3 py-1.5 bg-zinc-900 border border-zinc-700 rounded-lg text-white focus:outline-none focus:border-indigo-500"
          />
          <input
            type="text"
            value={newPromptTopic}
            onChange={(e) => setNewPromptTopic(e.target.value)}
            placeholder="トピック (任意)"
            className="md:col-span-2 px-2.5 py-1.5 bg-zinc-900 border border-zinc-700 rounded-lg text-white focus:outline-none focus:border-indigo-500"
          />
          <input
            type="text"
            value={newPromptDeleteAfter}
            onChange={(e) => setNewPromptDeleteAfter(e.target.value)}
            placeholder="削除時間 例:30～60"
            className="md:col-span-2 px-2.5 py-1.5 bg-zinc-900 border border-zinc-700 rounded-lg text-white font-mono focus:outline-none focus:border-indigo-500"
          />
          <button
            type="submit"
            className="md:col-span-2 flex items-center justify-center gap-1 px-3 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white font-bold rounded-lg cursor-pointer"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>スポット追加</span>
          </button>
        </form>

        <div className="overflow-x-auto border border-zinc-800 rounded-lg">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-zinc-950 text-zinc-400 border-b border-zinc-800">
                <th className="py-2 px-2.5 w-12">ID</th>
                <th className="py-2 px-2.5">生成されたポスト文章 / 個別指示</th>
                <th className="py-2 px-2.5 w-28">使用AIエンジン</th>
                <th className="py-2 px-2.5 w-24">ステータス</th>
                <th className="py-2 px-2.5 w-44">投稿完了日時 / エラー</th>
                <th className="py-2 px-2.5 w-12 text-center">操作</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-zinc-800/70">
              {history.length === 0 ? (
                <tr>
                  <td
                    colSpan={6}
                    className="py-6 text-center text-zinc-500 text-xs"
                  >
                    <div className="flex flex-col items-center gap-1">
                      <AlertCircle className="w-4 h-4 text-zinc-600" />
                      <span>
                        履歴はまだありません（有効にすると上の10項目をもとに自動生成・投稿＆ここへ記録されます）
                      </span>
                    </div>
                  </td>
                </tr>
              ) : (
                history.map((item) => (
                  <tr key={item.id} className="hover:bg-zinc-800/40">
                    <td className="py-2 px-2.5 font-mono text-zinc-400">
                      #{item.id}
                    </td>
                    <td className="py-2 px-2.5 text-zinc-200 max-w-md break-words whitespace-pre-wrap">
                      {item.generatedText || item.prompt || '―'}
                    </td>
                    <td className="py-2 px-2.5 font-mono text-indigo-300">
                      {item.engineUsed || config.engine}
                    </td>
                    <td className="py-2 px-2.5">
                      <span
                        className={`inline-block px-2 py-0.5 rounded text-[10px] font-bold ${
                          item.status.includes('完了')
                            ? 'bg-emerald-950 text-emerald-300 border border-emerald-800'
                            : item.status === 'エラー'
                            ? 'bg-rose-950 text-rose-300 border border-rose-800'
                            : 'bg-amber-950 text-amber-300 border border-amber-800'
                        }`}
                      >
                        {item.status}
                      </span>
                    </td>
                    <td className="py-2 px-2.5 text-zinc-400 text-[11px]">
                      {[
                        item.postedAt,
                        item.deleteAfter ? `削除:${item.deleteAfter}` : '',
                        item.errorMessage,
                      ]
                        .filter(Boolean)
                        .join(' / ') || '―'}
                    </td>
                    <td className="py-2 px-2.5 text-center">
                      <button
                        type="button"
                        onClick={() => handleDeleteRow(item.id)}
                        className="p-1 text-zinc-500 hover:text-rose-400 cursor-pointer"
                        title="削除"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
