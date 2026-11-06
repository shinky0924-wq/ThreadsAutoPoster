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

        {/* Notice about SessionID Logout issues */}
        <div className="p-3.5 rounded-xl bg-amber-950/40 border border-amber-800/80 space-y-1.5 text-zinc-200">
          <div className="font-bold text-amber-300 flex items-center gap-1.5">
            <AlertTriangle className="w-4 h-4 text-amber-400 shrink-0" />
            <span>なぜsessionid入力ではなく「プロファイル永続化」が必要なのか？</span>
          </div>
          <p className="text-zinc-300 text-[11px] leading-relaxed">
            Meta（Instagram/Threads）は定期的にセッションCookie（sessionid）を強制ローテーション・無効化するため、手動コピペしたsessionidは数時間〜数日で使えなくなり強制ログアウトされます。<br />
            <strong>「profile_launcher」で作成した専用Chromeプロファイル</strong>は、IndexedDB・ローカルストレージ・正規端末Cookieを完全保存するため、<strong>ID・PWで一度ログインするだけで永続的にログイン状態を維持</strong>できます。
          </p>
        </div>

        {/* 3-Step Simple Guide */}
        <div className="p-3.5 rounded-xl bg-zinc-950 border border-zinc-800 space-y-2">
          <span className="font-bold text-violet-300 block text-[11px] flex items-center gap-1.5">
            <Info className="w-4 h-4 text-violet-400" />
            【かんたん3ステップ】profile_launcherの使い方:
          </span>
          <ol className="list-decimal list-inside space-y-1.5 text-zinc-300 leading-relaxed pl-1 text-[11px]">
            <li>
              下のボタンから <code className="text-violet-300 font-mono">profile_launcher.py</code> をダウンロードして起動。
            </li>
            <li>
              プロファイル名に <code className="text-emerald-400 font-mono">{accountId}</code> を指定して「<strong>🚀 Chromeを起動</strong>」をクリック。
            </li>
            <li>
              開いたChromeでThreadsにID/PWで1回通常ログインし、ウィンドウを閉じるだけで保存完了！
            </li>
          </ol>
        </div>

        {/* Download profile_launcher.py button */}
        <div className="flex items-center justify-between p-3 rounded-xl bg-violet-950/30 border border-violet-800/50">
          <div>
            <div className="font-bold text-white text-xs">profile_launcher.py</div>
            <div className="text-[11px] text-zinc-400">プロファイル作成・ログイン永続化GUIツール</div>
          </div>
          <button
            onClick={handleDownloadLauncher}
            className="flex items-center gap-1.5 px-4 py-2 bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-violet-500 hover:to-indigo-500 text-white font-bold rounded-lg shadow cursor-pointer text-xs"
          >
            <Download className="w-3.5 h-3.5" />
            <span>ツールをダウンロード</span>
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
