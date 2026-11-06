import React, { useState } from 'react';
import { ShieldCheck, CheckCircle2, AlertTriangle, Download, Rocket, ExternalLink, X, Info } from 'lucide-react';
import { extractUserIdFromToken } from '../utils/spintax';
import { RAW_PROFILE_LAUNCHER_PY } from '../data/rawProfileLauncher';

interface SessionModalProps {
  isOpen: boolean;
  onClose: () => void;
  currentAuthToken: string;
  onSaveToken: (token: string) => void;
  accountId: string;
}

export const SessionModal: React.FC<SessionModalProps> = ({
  isOpen,
  onClose,
  currentAuthToken,
  onSaveToken,
  accountId,
}) => {
  const [tokenInput, setTokenInput] = useState(currentAuthToken);

  if (!isOpen) return null;

  const handleDownloadLauncher = () => {
    try {
      const blob = new Blob([RAW_PROFILE_LAUNCHER_PY], { type: 'text/x-python;charset=utf-8' });
      const blobUrl = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = blobUrl;
      a.download = 'profile_launcher.py';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(blobUrl);
    } catch (err: any) {
      console.error(err);
    }
  };

  return (
    <div className="fixed inset-0 bg-black/80 backdrop-blur-xs flex items-center justify-center p-4 z-50">
      <div className="bg-zinc-900 border border-violet-500/60 rounded-2xl max-w-xl w-full p-6 space-y-4 shadow-2xl relative text-xs">
        {/* Close Button */}
        <button
          onClick={onClose}
          className="absolute top-4 right-4 text-zinc-400 hover:text-white"
        >
          <X className="w-5 h-5" />
        </button>

        {/* Header */}
        <div className="flex items-center gap-2.5 pb-2 border-b border-zinc-800">
          <div className="p-2 rounded-xl bg-violet-950 text-violet-400 border border-violet-800">
            <Rocket className="w-5 h-5" />
          </div>
          <div>
            <h2 className="text-sm font-bold text-white flex items-center gap-2">
              <span>ログイン永続化 & プロファイル管理 (profile_launcher)</span>
              <span className="text-[10px] px-2 py-0.5 rounded-full bg-emerald-950 text-emerald-300 border border-emerald-800">
                推奨方式
              </span>
            </h2>
            <p className="text-zinc-400 text-[11px]">
              対象アカウント: <span className="font-mono text-emerald-400">@{accountId}</span>
            </p>
          </div>
        </div>

        {/* Notice about HTTP 429 & SessionID to Profile */}
        <div className="p-3.5 rounded-xl bg-violet-950/40 border border-violet-800/80 space-y-1.5 text-zinc-200">
          <div className="font-bold text-violet-300 flex items-center gap-1.5">
            <ShieldCheck className="w-4 h-4 text-violet-400 shrink-0" />
            <span>【HTTP ERROR 429 回避】sessionidからプロファイル自動生成・永続保存</span>
          </div>
          <p className="text-zinc-300 text-[11px] leading-relaxed">
            VPS等の同一IPからID/PWで何度もログインしようとすると、Metaのレートリミットにより「<strong>HTTP ERROR 429</strong>」や「後ほどもう一度実行してください」が発生します。<br />
            ブラウザで取得した <code className="text-emerald-400 font-mono">sessionid</code> を入力してプロファイル化することで、ログイン画面を一切介さずThreadsにログイン済み状態で永続接続できます。
          </p>
        </div>

        {/* Quick sessionid input */}
        <div className="p-3.5 rounded-xl bg-zinc-950 border border-zinc-800 space-y-2">
          <label className="font-bold text-zinc-300 block text-[11px]">
            sessionid（セッションCookie）:
          </label>
          <div className="flex gap-2">
            <input
              type="text"
              value={tokenInput}
              onChange={(e) => setTokenInput(e.target.value)}
              placeholder="例: 47094044745%3AXyZ..."
              className="flex-1 bg-zinc-900 border border-zinc-700 rounded-lg px-3 py-2 text-white font-mono text-xs focus:outline-none focus:border-violet-500"
            />
            <button
              onClick={() => {
                onSaveToken(tokenInput);
                onClose();
              }}
              className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold rounded-lg cursor-pointer text-xs flex items-center gap-1.5 shrink-0"
            >
              <CheckCircle2 className="w-3.5 h-3.5" />
              <span>保存して適用</span>
            </button>
          </div>
          <p className="text-[10px] text-zinc-500">
            ※保存すると「アカウント情報.xlsx」に即座に反映され、threads_gui_appおよびprofile_launcherと同期されます。
          </p>
        </div>

        {/* profile_launcher download */}
        <div className="flex items-center justify-between p-3 rounded-xl bg-zinc-950 border border-zinc-800">
          <div>
            <div className="font-bold text-white text-xs">profile_launcher.py (推奨ツール)</div>
            <div className="text-[11px] text-zinc-400">ワンクリックでsessionidからChromeプロファイルを生成・永続化</div>
          </div>
          <button
            onClick={handleDownloadLauncher}
            className="flex items-center gap-1.5 px-4 py-2 bg-violet-600 hover:bg-violet-500 text-white font-bold rounded-lg cursor-pointer text-xs shrink-0"
          >
            <Download className="w-3.5 h-3.5" />
            <span>ツールDL</span>
          </button>
        </div>

        {/* Footer Actions */}
        <div className="flex items-center justify-end gap-2 pt-2 border-t border-zinc-800">
          <button
            onClick={onClose}
            className="px-5 py-2 bg-zinc-800 hover:bg-zinc-700 text-white rounded-lg cursor-pointer text-xs"
          >
            閉じる
          </button>
        </div>
      </div>
    </div>
  );
};
