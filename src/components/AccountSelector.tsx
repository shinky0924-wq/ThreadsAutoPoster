import React, { useState } from 'react';
import { AccountInfo } from '../types';
import { User, Key, Sliders, Hash, Edit3, Check, Plus, AlertCircle } from 'lucide-react';

interface AccountSelectorProps {
  accounts: AccountInfo[];
  selectedAccountId: string;
  onSelectAccount: (id: string) => void;
  onUpdateAccount: (updated: AccountInfo) => void;
  onAddAccount: (newAcc: AccountInfo) => void;
}

export const AccountSelector: React.FC<AccountSelectorProps> = ({
  accounts,
  selectedAccountId,
  onSelectAccount,
  onUpdateAccount,
  onAddAccount,
}) => {
  const currentAccount = accounts.find((a) => a.id === selectedAccountId) || accounts[0];
  const [isEditingSettings, setIsEditingSettings] = useState(false);
  const [tempAccount, setTempAccount] = useState<AccountInfo>(currentAccount);
  const [isAddingAccount, setIsAddingAccount] = useState(false);
  const [newAccountId, setNewAccountId] = useState('');
  const [newAuthToken, setNewAuthToken] = useState('');

  // Sync tempAccount when selected account changes
  React.useEffect(() => {
    if (currentAccount) {
      setTempAccount(currentAccount);
      setIsEditingSettings(false);
    }
  }, [selectedAccountId, currentAccount]);

  const handleSaveSettings = () => {
    onUpdateAccount(tempAccount);
    setIsEditingSettings(false);
  };

  const handleCreateAccount = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newAccountId.trim()) return;
    const cleanId = newAccountId.trim().replace(/^@/, '');
    const cleanToken = newAuthToken.trim();
    const newAcc: AccountInfo = {
      id: cleanId,
      authToken: cleanToken,
      userId: cleanToken.includes(':') ? cleanToken.split(':')[0] : '',
      notes: '手動追加アカウント',
      minIntervalMinutes: 120,
      maxIntervalMinutes: 360,
      activeHours: '0:00-0:00',
      mode: 1,
      profileDir: `chrome_profiles/${cleanId}`,
      topics: ['信太山新地', '大阪求人'],
      currentTopicIndex: 0,
    };
    onAddAccount(newAcc);
    setNewAccountId('');
    setNewAuthToken('');
    setIsAddingAccount(false);
  };

  if (!currentAccount) return null;

  const hasToken = !!currentAccount.authToken && !currentAccount.authToken.startsWith('ここに');

  return (
    <div className="bg-zinc-900 border border-zinc-800 rounded-xl p-4 shadow-sm mb-4">
      {/* Account Tabs and Status */}
      <div className="flex flex-wrap items-center justify-between gap-3 pb-3 border-b border-zinc-800/80">
        <div className="flex items-center gap-2 flex-wrap">
          <span className="text-xs font-semibold text-zinc-400 flex items-center gap-1">
            <User className="w-3.5 h-3.5 text-zinc-500" /> アカウント切替:
          </span>
          <div className="flex items-center gap-1.5 flex-wrap">
            {accounts.map((acc) => {
              const isSelected = acc.id === selectedAccountId;
              return (
                <button
                  key={acc.id}
                  onClick={() => onSelectAccount(acc.id)}
                  className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium transition-all cursor-pointer ${
                    isSelected
                      ? 'bg-zinc-800 text-white shadow border border-zinc-700 font-semibold'
                      : 'bg-zinc-950/60 text-zinc-400 hover:text-zinc-200 border border-transparent'
                  }`}
                >
                  <span className="w-2 h-2 rounded-full bg-emerald-500" />
                  <span>@{acc.id}</span>
                  {acc.id === 'sll15802' && (
                    <span className="text-[10px] px-1 py-0.2 rounded bg-emerald-950 text-emerald-300 font-mono">
                      Main
                    </span>
                  )}
                </button>
              );
            })}
            <button
              onClick={() => setIsAddingAccount(!isAddingAccount)}
              className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs text-zinc-400 hover:text-zinc-200 bg-zinc-950/40 hover:bg-zinc-800 border border-zinc-800 cursor-pointer"
              title="アカウントを追加"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>追加</span>
            </button>
          </div>
        </div>

        {/* Edit Setting Button */}
        <button
          onClick={() => {
            if (isEditingSettings) handleSaveSettings();
            else setIsEditingSettings(true);
          }}
          className="flex items-center gap-1.5 text-xs px-2.5 py-1 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 border border-zinc-700 rounded-md cursor-pointer"
        >
          {isEditingSettings ? (
            <>
              <Check className="w-3.5 h-3.5 text-emerald-400" />
              <span className="font-semibold text-emerald-300">設定を保存</span>
            </>
          ) : (
            <>
              <Edit3 className="w-3.5 h-3.5 text-zinc-400" />
              <span>設定を編集</span>
            </>
          )}
        </button>
      </div>

      {/* Add Account Panel */}
      {isAddingAccount && (
        <form onSubmit={handleCreateAccount} className="p-3 my-3 bg-zinc-950 rounded-lg border border-zinc-800 flex flex-wrap gap-2 items-center">
          <input
            type="text"
            placeholder="アカウントID (例: my_threads_acc)"
            value={newAccountId}
            onChange={(e) => setNewAccountId(e.target.value)}
            className="bg-zinc-900 border border-zinc-700 text-xs text-white px-2.5 py-1.5 rounded focus:outline-none focus:border-emerald-500 flex-1 min-w-[160px]"
          />
          <input
            type="text"
            placeholder="auth_token / sessionid (省略可)"
            value={newAuthToken}
            onChange={(e) => setNewAuthToken(e.target.value)}
            className="bg-zinc-900 border border-zinc-700 text-xs text-white px-2.5 py-1.5 rounded focus:outline-none focus:border-emerald-500 flex-1 min-w-[200px]"
          />
          <button
            type="submit"
            className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded cursor-pointer"
          >
            登録
          </button>
          <button
            type="button"
            onClick={() => setIsAddingAccount(false)}
            className="px-2 py-1.5 text-zinc-400 hover:text-zinc-200 text-xs cursor-pointer"
          >
            キャンセル
          </button>
        </form>
      )}

      {/* Details Row */}
      {!isEditingSettings ? (
        <div className="grid grid-cols-1 md:grid-cols-4 gap-3 mt-3 text-xs">
          {/* Auth Token Status */}
          <div className="bg-zinc-950/60 p-2.5 rounded-lg border border-zinc-800/80">
            <div className="text-zinc-500 flex items-center gap-1 mb-1">
              <Key className="w-3.5 h-3.5 text-violet-400" />
              <span>Authトークン (sessionid)</span>
            </div>
            <div className="flex items-center gap-1.5">
              <span className={`w-2 h-2 rounded-full ${hasToken ? 'bg-emerald-400' : 'bg-rose-400'}`} />
              <span className="font-mono text-zinc-200 truncate">
                {hasToken ? `${currentAccount.authToken.slice(0, 18)}...` : '未設定'}
              </span>
            </div>
            {currentAccount.userId && (
              <div className="text-[11px] text-zinc-500 mt-0.5 font-mono">
                ユーザーID: {currentAccount.userId}
              </div>
            )}
          </div>

          {/* Interval */}
          <div className="bg-zinc-950/60 p-2.5 rounded-lg border border-zinc-800/80">
            <div className="text-zinc-500 flex items-center gap-1 mb-1">
              <Sliders className="w-3.5 h-3.5 text-sky-400" />
              <span>ポスト間隔</span>
            </div>
            <div className="font-medium text-zinc-200">
              {currentAccount.minIntervalMinutes}分 〜 {currentAccount.maxIntervalMinutes}分
            </div>
            <div className="text-[11px] text-zinc-500 mt-0.5">毎回ランダムで待機 (BAN強力防止)</div>
          </div>

          {/* Active hours & mode */}
          <div className="bg-zinc-950/60 p-2.5 rounded-lg border border-zinc-800/80">
            <div className="text-zinc-500 flex items-center gap-1 mb-1">
              <Sliders className="w-3.5 h-3.5 text-amber-400" />
              <span>実行時間帯 / モード</span>
            </div>
            <div className="font-medium text-zinc-200">
              {currentAccount.activeHours} ({currentAccount.mode === 1 ? 'ランダム投稿' : '上から順'})
            </div>
            <div className="text-[11px] text-zinc-500 mt-0.5 font-mono truncate">
              {currentAccount.profileDir}
            </div>
          </div>

          {/* Topic A/B rotation */}
          <div className="bg-zinc-950/60 p-2.5 rounded-lg border border-zinc-800/80">
            <div className="text-zinc-500 flex items-center gap-1 mb-1">
              <Hash className="w-3.5 h-3.5 text-teal-400" />
              <span>トピック A/Bテスト</span>
            </div>
            <div className="font-medium text-teal-300 truncate">
              {currentAccount.topics && currentAccount.topics.length > 0
                ? currentAccount.topics.join(', ')
                : '未指定'}
            </div>
            <div className="text-[11px] text-zinc-500 mt-0.5">投稿ごとに均等ローテーション</div>
          </div>
        </div>
      ) : (
        /* Edit Settings Panel */
        <div className="mt-3 bg-zinc-950 p-3 rounded-lg border border-zinc-700 space-y-3 text-xs">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <div>
              <label className="text-zinc-400 font-medium block mb-1">Authトークン (sessionid)</label>
              <input
                type="text"
                value={tempAccount.authToken}
                onChange={(e) => setTempAccount({ ...tempAccount, authToken: e.target.value })}
                className="w-full bg-zinc-900 border border-zinc-700 text-white text-xs px-2.5 py-1.5 rounded focus:outline-none focus:border-emerald-500 font-mono"
              />
            </div>
            <div>
              <label className="text-zinc-400 font-medium block mb-1">トピック (カンマ区切りでA/Bテスト)</label>
              <input
                type="text"
                value={tempAccount.topics.join(', ')}
                onChange={(e) =>
                  setTempAccount({
                    ...tempAccount,
                    topics: e.target.value.split(',').map((t) => t.trim()).filter(Boolean),
                  })
                }
                className="w-full bg-zinc-900 border border-zinc-700 text-white text-xs px-2.5 py-1.5 rounded focus:outline-none focus:border-emerald-500"
              />
            </div>
          </div>

          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            <div>
              <label className="text-zinc-400 font-medium block mb-1">最小間隔 (分)</label>
              <input
                type="number"
                value={tempAccount.minIntervalMinutes}
                onChange={(e) =>
                  setTempAccount({ ...tempAccount, minIntervalMinutes: Number(e.target.value) || 60 })
                }
                className="w-full bg-zinc-900 border border-zinc-700 text-white text-xs px-2.5 py-1.5 rounded focus:outline-none"
              />
            </div>
            <div>
              <label className="text-zinc-400 font-medium block mb-1">最大間隔 (分)</label>
              <input
                type="number"
                value={tempAccount.maxIntervalMinutes}
                onChange={(e) =>
                  setTempAccount({ ...tempAccount, maxIntervalMinutes: Number(e.target.value) || 180 })
                }
                className="w-full bg-zinc-900 border border-zinc-700 text-white text-xs px-2.5 py-1.5 rounded focus:outline-none"
              />
            </div>
            <div>
              <label className="text-zinc-400 font-medium block mb-1">実行時間帯</label>
              <input
                type="text"
                value={tempAccount.activeHours}
                placeholder="0:00-0:00 (終日)"
                onChange={(e) => setTempAccount({ ...tempAccount, activeHours: e.target.value })}
                className="w-full bg-zinc-900 border border-zinc-700 text-white text-xs px-2.5 py-1.5 rounded focus:outline-none"
              />
            </div>
            <div>
              <label className="text-zinc-400 font-medium block mb-1">投稿モード</label>
              <select
                value={tempAccount.mode}
                onChange={(e) =>
                  setTempAccount({ ...tempAccount, mode: Number(e.target.value) as 0 | 1 })
                }
                className="w-full bg-zinc-900 border border-zinc-700 text-white text-xs px-2.5 py-1.5 rounded focus:outline-none"
              >
                <option value={1}>1: ランダム (推奨)</option>
                <option value={0}>0: 上から順</option>
              </select>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
