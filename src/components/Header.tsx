import React, { useRef } from 'react';
import { 
  Play, 
  Square, 
  Zap, 
  ShieldCheck, 
  RotateCcw, 
  Download, 
  Upload, 
  Trash2,
  Clock,
  Sparkles,
  FileCode2
} from 'lucide-react';

interface HeaderProps {
  isRunning: boolean;
  isPausedInterval: boolean;
  isStealthing: boolean;
  statusText: string;
  countdownSeconds: number;
  onStart: () => void;
  onStop: () => void;
  onTestPost: () => void;
  onOpenSessionModal: () => void;
  onResetToUnposted: () => void;
  onReloadDefaultData: () => void;
  onExportExcel: () => void;
  onImportExcel: (file: File) => void;
  onClearLogs: () => void;
  onDownloadPyScript: () => void;
}

export const Header: React.FC<HeaderProps> = ({
  isRunning,
  isPausedInterval,
  isStealthing,
  statusText,
  countdownSeconds,
  onStart,
  onStop,
  onTestPost,
  onOpenSessionModal,
  onResetToUnposted,
  onReloadDefaultData,
  onExportExcel,
  onImportExcel,
  onClearLogs,
  onDownloadPyScript,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);

  const formatCountdown = (secs: number) => {
    const m = Math.floor(secs / 60);
    const s = secs % 60;
    return `${m}分${s < 10 ? '0' : ''}${s}秒`;
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      onImportExcel(e.target.files[0]);
      e.target.value = '';
    }
  };

  return (
    <header className="border-b border-zinc-800 bg-zinc-950/80 backdrop-blur sticky top-0 z-40">
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        className="hidden"
        onChange={handleFileChange}
      />

      {/* Top Brand Bar */}
      <div className="max-w-7xl mx-auto px-4 sm:px-6 py-3 flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-zinc-800 via-zinc-900 to-emerald-600 flex items-center justify-center text-white font-black text-xl shadow-lg border border-zinc-700">
            @
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h1 className="text-lg font-bold text-white tracking-tight flex items-center gap-1.5">
                Threads AutoPoster <span className="text-emerald-400 font-semibold text-sm">Studio</span>
              </h1>
              <span className="text-[10px] uppercase font-mono px-2 py-0.5 rounded-full bg-emerald-950/80 text-emerald-300 border border-emerald-800/60 font-semibold">
                v2.5 Web
              </span>
            </div>
            <p className="text-xs text-zinc-400 hidden sm:block">
              Meta Threads 自動投稿システム・1窓ローテーション & ステルスBAN防止エンジン
            </p>
          </div>
        </div>

        {/* Live Status Badge */}
        <div className="flex items-center gap-2">
          <div className={`flex items-center gap-2 px-3.5 py-1.5 rounded-full text-xs font-medium border shadow-inner ${
            isRunning
              ? isStealthing
                ? 'bg-amber-950/40 text-amber-300 border-amber-700/60'
                : isPausedInterval
                ? 'bg-purple-950/40 text-purple-300 border-purple-700/60'
                : 'bg-emerald-950/40 text-emerald-300 border-emerald-700/60'
              : 'bg-zinc-900 text-zinc-300 border-zinc-800'
          }`}>
            <span className="relative flex h-2.5 w-2.5">
              {isRunning && (
                <span className={`animate-ping absolute inline-flex h-full w-full rounded-full opacity-75 ${
                  isStealthing ? 'bg-amber-400' : isPausedInterval ? 'bg-purple-400' : 'bg-emerald-400'
                }`} />
              )}
              <span className={`relative inline-flex rounded-full h-2.5 w-2.5 ${
                isRunning
                  ? isStealthing ? 'bg-amber-400' : isPausedInterval ? 'bg-purple-400' : 'bg-emerald-400'
                  : 'bg-zinc-500'
              }`} />
            </span>
            <span className="font-medium tracking-wide">{statusText}</span>
            {isRunning && isPausedInterval && countdownSeconds > 0 && (
              <span className="flex items-center gap-1 text-[11px] font-mono text-purple-200 bg-purple-900/40 px-2 py-0.5 rounded">
                <Clock className="w-3 h-3 animate-spin" />
                {formatCountdown(countdownSeconds)}
              </span>
            )}
          </div>
        </div>
      </div>

      {/* Main Command Actions Toolbar */}
      <div className="bg-zinc-900/90 border-t border-zinc-800/80 px-4 sm:px-6 py-2.5">
        <div className="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-2">
          {/* Execution Controls */}
          <div className="flex items-center gap-2 flex-wrap">
            {!isRunning ? (
              <button
                onClick={onStart}
                className="flex items-center gap-1.5 px-4 py-2 bg-emerald-600 hover:bg-emerald-500 active:scale-95 text-white text-sm font-semibold rounded-lg shadow-md hover:shadow-emerald-900/30 transition-all cursor-pointer"
              >
                <Play className="w-4 h-4 fill-white" />
                <span>自動投稿を開始</span>
              </button>
            ) : (
              <button
                onClick={onStop}
                className="flex items-center gap-1.5 px-4 py-2 bg-rose-600 hover:bg-rose-500 active:scale-95 text-white text-sm font-semibold rounded-lg shadow-md hover:shadow-rose-900/30 transition-all cursor-pointer"
              >
                <Square className="w-4 h-4 fill-white" />
                <span>停止</span>
              </button>
            )}

            <button
              onClick={onTestPost}
              disabled={isRunning}
              className="flex items-center gap-1.5 px-3 py-2 bg-amber-500/20 hover:bg-amber-500/30 text-amber-300 border border-amber-500/40 text-xs font-semibold rounded-lg disabled:opacity-40 transition-all cursor-pointer"
              title="待機時間をスキップして未投稿の1件を即座に投稿テスト"
            >
              <Zap className="w-3.5 h-3.5 text-amber-400" />
              <span>1件即時テスト</span>
            </button>

            <button
              onClick={onOpenSessionModal}
              className="flex items-center gap-1.5 px-3 py-2 bg-violet-600/20 hover:bg-violet-600/30 text-violet-300 border border-violet-500/40 text-xs font-semibold rounded-lg transition-all cursor-pointer"
            >
              <ShieldCheck className="w-3.5 h-3.5 text-violet-400" />
              <span>セッション検証</span>
            </button>
          </div>

          {/* Excel & Management Controls */}
          <div className="flex items-center gap-1.5 flex-wrap">
            <button
              onClick={onDownloadPyScript}
              className="flex items-center gap-1 px-3 py-1.5 bg-gradient-to-r from-indigo-600 to-purple-600 hover:from-indigo-500 hover:to-purple-500 text-white text-xs font-bold rounded-md shadow transition-all cursor-pointer border border-indigo-400/40"
              title="最新の threads_gui_app.py を直接ダウンロード (HTML混入なし)"
            >
              <FileCode2 className="w-3.5 h-3.5" />
              <span>⬇️ threads_gui_app.py</span>
            </button>

            <button
              onClick={onResetToUnposted}
              className="flex items-center gap-1 px-2.5 py-1.5 bg-indigo-600/20 hover:bg-indigo-600/30 text-indigo-300 border border-indigo-500/30 text-xs font-medium rounded-md transition-colors cursor-pointer"
              title="完了した投稿をすべて『未投稿』ステータスに戻します"
            >
              <RotateCcw className="w-3 h-3" />
              <span>未投稿にリセット</span>
            </button>

            <button
              onClick={() => fileInputRef.current?.click()}
              className="flex items-center gap-1 px-2.5 py-1.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 border border-zinc-700 text-xs font-medium rounded-md transition-colors cursor-pointer"
              title="Excel (.xlsx) をアップロードして読み込み"
            >
              <Upload className="w-3 h-3" />
              <span>Excel読込</span>
            </button>

            <button
              onClick={onExportExcel}
              className="flex items-center gap-1 px-2.5 py-1.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 border border-zinc-700 text-xs font-medium rounded-md transition-colors cursor-pointer"
              title="現在の投稿状態を .xlsx ファイルとして出力"
            >
              <Download className="w-3 h-3" />
              <span>Excel出力</span>
            </button>

            <button
              onClick={onReloadDefaultData}
              className="flex items-center gap-1 px-2.5 py-1.5 bg-zinc-800/80 hover:bg-zinc-700 text-zinc-300 text-xs font-medium rounded-md transition-colors cursor-pointer"
              title="初期プリセットデータに復元"
            >
              <Sparkles className="w-3 h-3 text-sky-400" />
              <span>初期化</span>
            </button>

            <button
              onClick={onClearLogs}
              className="flex items-center gap-1 px-2 py-1.5 text-zinc-400 hover:text-zinc-200 text-xs rounded-md transition-colors cursor-pointer"
              title="実行ログをクリア"
            >
              <Trash2 className="w-3 h-3" />
            </button>
          </div>
        </div>
      </div>
    </header>
  );
};
