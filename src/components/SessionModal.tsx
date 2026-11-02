import React, { useState } from 'react';
import { ShieldCheck, CheckCircle2, AlertTriangle, Copy, ExternalLink, X } from 'lucide-react';
import { extractUserIdFromToken } from '../utils/spintax';

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
  const [testResult, setTestResult] = useState<{
    valid: boolean;
    userId: string;
    message: string;
  } | null>(null);

  if (!isOpen) return null;

  const handleTestToken = () => {
    const clean = tokenInput.trim();
    if (!clean) {
      setTestResult({
        valid: false,
        userId: '',
        message: 'トークンが入力されていません。',
      });
      return;
    }

    const userId = extractUserIdFromToken(clean);
    if (userId) {
      setTestResult({
        valid: true,
        userId,
        message: `✅ 有効なsessionid形式です！ユーザーID: [${userId}] を検出しました。Threads / Instagramのクッキー注入に対応しています。`,
      });
    } else {
      setTestResult({
        valid: true,
        userId: '不明',
        message: `⚠️ sessionid文字列を認識しました（ユーザーID分離形式ではありませんが認証試行可能です）。`,
      });
    }
  };

  const handleSave = () => {
    onSaveToken(tokenInput.trim());
    onClose();
  };

  return (
    <div className="fixed inset-0 bg-black/70 backdrop-blur-xs flex items-center justify-center p-4 z-50">
      <div className="bg-zinc-900 border border-zinc-700 rounded-2xl max-w-xl w-full p-6 space-y-4 shadow-2xl relative text-xs">
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
            <ShieldCheck className="w-5 h-5" />
          </div>
          <div>
            <h2 className="text-sm font-bold text-white">
              Authトークン (sessionid) による自動ログイン設定
            </h2>
            <p className="text-zinc-400 text-[11px]">
              対象アカウント: <span className="font-mono text-emerald-400">@{accountId}</span>
            </p>
          </div>
        </div>

        {/* 30-sec Guide */}
        <div className="p-3.5 rounded-xl bg-zinc-950 border border-zinc-800 space-y-2">
          <span className="font-bold text-zinc-300 block text-[11px]">
            【簡単30秒！】sessionidの取得手順:
          </span>
          <ol className="list-decimal list-inside space-y-1 text-zinc-400 leading-relaxed pl-1">
            <li>
              Google Chrome等のブラウザで{' '}
              <a
                href="https://www.threads.net/"
                target="_blank"
                rel="noreferrer"
                className="text-sky-400 hover:underline inline-flex items-center gap-0.5"
              >
                Threads.net <ExternalLink className="w-2.5 h-2.5" />
              </a>{' '}
              を開き、ログインします。
            </li>
            <li>キーボードの <kbd className="px-1 py-0.5 rounded bg-zinc-800 font-mono text-zinc-200">F12</kbd> キーを押してデベロッパーツールを開きます。</li>
            <li>上部タブの「Application（アプリケーション）」をクリックします。</li>
            <li>左メニュー「Cookies」内の「https://www.threads.net」をクリックします。</li>
            <li>一覧の「sessionid」の値をコピーして、下の欄に貼り付けます。</li>
          </ol>
        </div>

        {/* Input & Tester */}
        <div className="space-y-2">
          <label className="text-zinc-300 font-semibold block">
            sessionid / Authトークン:
          </label>
          <textarea
            value={tokenInput}
            onChange={(e) => {
              setTokenInput(e.target.value);
              setTestResult(null);
            }}
            rows={3}
            placeholder="例: 47094044745:HP477WPO3v01NR:4:AYlIDUGTeie5QpgJ04lsSGIceblgEb_1RYr_AuqSwg"
            className="w-full bg-zinc-950 border border-zinc-700 text-white p-2.5 rounded-lg focus:outline-none focus:border-violet-500 font-mono text-xs"
          />

          <div className="flex items-center justify-between gap-2 pt-1">
            <button
              type="button"
              onClick={handleTestToken}
              className="px-3 py-1.5 bg-violet-600/30 hover:bg-violet-600/40 text-violet-300 border border-violet-500/50 rounded-lg font-medium cursor-pointer"
            >
              トークン形式テスト
            </button>

            {testResult && (
              <span className={`text-[11px] font-mono ${testResult.valid ? 'text-emerald-400' : 'text-rose-400'}`}>
                {testResult.message}
              </span>
            )}
          </div>
        </div>

        {/* Footer Actions */}
        <div className="flex items-center justify-end gap-2 pt-3 border-t border-zinc-800">
          <button
            onClick={onClose}
            className="px-4 py-2 text-zinc-400 hover:text-white cursor-pointer"
          >
            キャンセル
          </button>
          <button
            onClick={handleSave}
            className="px-5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold rounded-lg shadow cursor-pointer"
          >
            トークンを保存して適用
          </button>
        </div>
      </div>
    </div>
  );
};
