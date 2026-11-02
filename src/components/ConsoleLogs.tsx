import React, { useRef, useEffect, useState } from 'react';
import { LogMessage } from '../types';
import { Terminal, Download, Trash2, ArrowDown, Filter } from 'lucide-react';

interface ConsoleLogsProps {
  logs: LogMessage[];
  onClear: () => void;
}

export const ConsoleLogs: React.FC<ConsoleLogsProps> = ({ logs, onClear }) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const [autoScroll, setAutoScroll] = useState(true);
  const [filterLevel, setFilterLevel] = useState<string>('ALL');

  useEffect(() => {
    if (autoScroll && containerRef.current) {
      containerRef.current.scrollTop = containerRef.current.scrollHeight;
    }
  }, [logs, autoScroll]);

  const handleDownloadLogs = (type: 'all' | 'error') => {
    const targetLogs = type === 'error' ? logs.filter((l) => l.level === 'ERROR' || l.level === 'WARN') : logs;
    const content = targetLogs.map((l) => `[${l.timestamp}] [${l.level}] ${l.message}`).join('\n');
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = type === 'error' ? 'error_log.txt' : 'execution_log.txt';
    a.click();
    URL.revokeObjectURL(url);
  };

  const filteredLogs = logs.filter((l) => {
    if (filterLevel === 'ALL') return true;
    return l.level === filterLevel;
  });

  return (
    <div className="bg-zinc-950 border border-zinc-800 rounded-xl overflow-hidden shadow-lg flex flex-col h-[480px]">
      {/* Terminal Title Bar */}
      <div className="bg-zinc-900 px-4 py-2.5 border-b border-zinc-800 flex flex-wrap items-center justify-between gap-2 text-xs">
        <div className="flex items-center gap-2">
          <Terminal className="w-4 h-4 text-emerald-400" />
          <span className="font-mono font-bold text-zinc-200">
            実行ログ / コンソール出力 (Threads AutoPoster Engine)
          </span>
          <span className="text-[10px] px-1.5 py-0.5 rounded bg-zinc-800 text-zinc-400 font-mono">
            {logs.length} 行
          </span>
        </div>

        <div className="flex items-center gap-2 flex-wrap">
          {/* Level Filter */}
          <div className="flex items-center gap-1 text-[11px] text-zinc-400">
            <Filter className="w-3 h-3 text-zinc-500" />
            <select
              value={filterLevel}
              onChange={(e) => setFilterLevel(e.target.value)}
              className="bg-zinc-950 border border-zinc-700 text-zinc-300 rounded px-1.5 py-0.5 font-mono cursor-pointer"
            >
              <option value="ALL">ALL (全ログ)</option>
              <option value="INFO">INFOのみ</option>
              <option value="SUCCESS">SUCCESSのみ</option>
              <option value="WARN">WARNのみ</option>
              <option value="ERROR">ERRORのみ</option>
            </select>
          </div>

          <label className="flex items-center gap-1 text-[11px] text-zinc-400 cursor-pointer">
            <input
              type="checkbox"
              checked={autoScroll}
              onChange={(e) => setAutoScroll(e.target.checked)}
              className="accent-emerald-500 rounded"
            />
            <span>自動スクロール</span>
          </label>

          <button
            onClick={() => handleDownloadLogs('all')}
            className="flex items-center gap-1 px-2 py-1 bg-zinc-800 hover:bg-zinc-700 text-zinc-300 rounded text-[11px] cursor-pointer"
            title="execution_log.txt をダウンロード"
          >
            <Download className="w-3 h-3" />
            <span>全ログ(txt)</span>
          </button>

          <button
            onClick={() => handleDownloadLogs('error')}
            className="flex items-center gap-1 px-2 py-1 bg-rose-950/60 hover:bg-rose-900/60 text-rose-300 border border-rose-800/60 rounded text-[11px] cursor-pointer"
            title="error_log.txt をダウンロード"
          >
            <Download className="w-3 h-3" />
            <span>エラーログ(txt)</span>
          </button>

          <button
            onClick={onClear}
            className="p-1 text-zinc-500 hover:text-zinc-300 rounded cursor-pointer"
            title="コンソールをクリア"
          >
            <Trash2 className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* Terminal Content */}
      <div
        ref={containerRef}
        className="flex-1 p-3.5 overflow-y-auto font-mono text-[12px] leading-relaxed space-y-1 bg-black text-zinc-300 selection:bg-emerald-900 selection:text-white"
      >
        {filteredLogs.length === 0 ? (
          <div className="text-zinc-600 text-center py-10 font-sans">
            ログはありません。「自動投稿を開始」または「1件即時テスト」をクリックしてください。
          </div>
        ) : (
          filteredLogs.map((log) => {
            let badgeClass = 'text-sky-400';
            if (log.level === 'SUCCESS') badgeClass = 'text-emerald-400 font-bold';
            if (log.level === 'WARN') badgeClass = 'text-amber-400 font-bold';
            if (log.level === 'ERROR') badgeClass = 'text-rose-400 font-bold';
            if (log.level === 'STEALTH') badgeClass = 'text-purple-400 font-bold';

            return (
              <div key={log.id} className="flex items-start gap-2 break-all hover:bg-zinc-900/40 py-0.5 rounded px-1">
                <span className="text-zinc-500 select-none text-[11px]">[{log.timestamp}]</span>
                <span className={`select-none ${badgeClass}`}>[{log.level}]</span>
                <span className="text-zinc-200 whitespace-pre-wrap">{log.message}</span>
              </div>
            );
          })
        )}
      </div>
    </div>
  );
};
