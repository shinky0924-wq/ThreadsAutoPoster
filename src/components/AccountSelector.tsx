import React, { useState } from 'react';
import { AccountInfo } from '../types';
import { 
  User, 
  Key, 
  Sliders, 
  Hash, 
  Edit3, 
  Check, 
  Plus, 
  AlertTriangle, 
  ShieldCheck, 
  ShieldAlert, 
  Power, 
  RefreshCw,
  Smartphone
} from 'lucide-react';

interface AccountSelectorProps {
  accounts: AccountInfo[];
  selectedAccountId: string;
  onSelectAccount: (id: string) => void;
  onUpdateAccount: (updated: AccountInfo) => void;
  onAddAccount: (newAcc: AccountInfo) => void;
  onSimulateCheckpoint?: (accId: string) => void;
}

export const AccountSelector: React.FC<AccountSelectorProps> = ({
  accounts,
  selectedAccountId,
  onSelectAccount,
  onUpdateAccount,
  onAddAccount,
  onSimulateCheckpoint,
}) => {
  const currentAccount = accounts.find((a) => a.id === selectedAccountId) || accounts[0];
  const [isEditingSettings, setIsEditingSettings] = useState(false);
  const [tempAccount, setTempAccount] = useState<AccountInfo>(currentAccount);
  const [isAddingAccount, setIsAddingAccount] = useState(false);
  const [newAccountId, setNewAccountId] = useState('');
  const [newProfileName, setNewProfileName] = useState('');
  const [showCheckpointModal, setShowCheckpointModal] = useState(false);
  const [showRedirectModal, setShowRedirectModal] = useState(false);

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

  const handleToggleAccountEnabled = () => {
    const nextEnabled = currentAccount.enabled === false;
    const updated: AccountInfo = {
      ...currentAccount,
      enabled: nextEnabled,
      checkpointReason: nextEnabled ? undefined : (currentAccount.checkpointReason || '手動停止'),
    };
    onUpdateAccount(updated);
  };

  const handleCreateAccount = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newAccountId.trim()) return;
    const cleanId = newAccountId.trim().replace(/^@/, '');
    const cleanProfile = newProfileName.trim() || cleanId;
    const newAcc: AccountInfo = {
      id: cleanId,
      authToken: '',
      userId: '',
      notes: 'profile_launcher対応アカウント',
      minIntervalMinutes: 120,
      maxIntervalMinutes: 360,
      activeHours: '0:00-0:00',
      mode: 1,
      profileDir: `chrome_profiles/${cleanProfile}`,
      topics: ['信太山新地', '大阪求人'],
      currentTopicIndex: 0,
      enabled: true,
    };
    onAddAccount(newAcc);
    setNewAccountId('');
    setNewProfileName('');
    setIsAddingAccount(false);
  };

  if (!currentAccount) return null;

  const isEnabled = currentAccount.enabled !== false;
  const hasToken = !!currentAccount.authToken && !currentAccount.authToken.startsWith('ここに');

  return (
    <div className="bg-zinc-900 border border-zinc-800 rounded-xl p-4 shadow-sm mb-4 space-y-3">
      {/* Account Tabs and Status */}
      <div className="flex flex-wrap items-center justify-between gap-3 pb-3 border-b border-zinc-800/80">
        <div className="flex items-center gap-2 flex-wrap">
          <span className="text-xs font-semibold text-zinc-400 flex items-center gap-1">
            <User className="w-3.5 h-3.5 text-zinc-500" /> アカウント切替:
          </span>
          <div className="flex items-center gap-1.5 flex-wrap">
            {accounts.map((acc) => {
              const isSelected = acc.id === selectedAccountId;
              const accActive = acc.enabled !== false;
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
                  <span className={`w-2 h-2 rounded-full ${accActive ? 'bg-emerald-500' : 'bg-rose-500 animate-pulse'}`} />
                  <span>@{acc.id}</span>
                  {!accActive && (
                    <span className="text-[10px] px-1 py-0.2 rounded bg-rose-950 text-rose-300 border border-rose-800 font-mono">
                      停止
                    </span>
                  )}
                  {acc.id === 'sll15802' && accActive && (
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

        {/* Action Controls: Enable/Disable Toggle + Edit Setting Button */}
        <div className="flex items-center gap-2">
          <button
            onClick={handleToggleAccountEnabled}
            className={`flex items-center gap-1 text-xs px-2.5 py-1 rounded-md font-semibold border transition-all cursor-pointer ${
              isEnabled
                ? 'bg-emerald-950/60 hover:bg-rose-950/80 text-emerald-300 hover:text-rose-300 border-emerald-800 hover:border-rose-800'
                : 'bg-rose-950 text-rose-300 border-rose-800 hover:bg-emerald-950 hover:text-emerald-300 hover:border-emerald-800'
            }`}
            title="クリックでこのアカウントの有効/無効を切替（アカウント情報.xlsxに即時反映）"
          >
            <Power className="w-3.5 h-3.5" />
            <span>{isEnabled ? '稼働中 (クリックで停止)' : '停止中 (クリックで再開)'}</span>
          </button>

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
      </div>

      {/* Checkpoint Warning Banner (Shown if account is stopped due to phone verification) */}
      {!isEnabled && (
        <div className="p-3 rounded-xl bg-rose-950/60 border border-rose-800 flex items-start justify-between gap-3 text-xs">
          <div className="flex items-start gap-2.5">
            <AlertTriangle className="w-5 h-5 text-rose-400 shrink-0 mt-0.5" />
            <div className="space-y-1">
              <div className="font-bold text-rose-200 flex items-center gap-2">
                <span>@{currentAccount.id} は現在「停止」状態です</span>
                <span className="px-1.5 py-0.5 rounded bg-rose-900/80 text-[10px] font-mono text-rose-200 border border-rose-700">
                  {currentAccount.checkpointReason || '要認証・自動停止'}
                </span>
              </div>
              <p className="text-rose-300/90 text-[11px] leading-relaxed">
                携帯電話番号認証（SMS確認）画面が検知されたため、アカウント保護のために自動停止されました。
                「アカウント情報.xlsx」の【有効/無効】列は「無効」に更新されており、手動でSMS認証を完了するまでこのアカウントの自動投稿はスキップされます（他のアカウントは継続して安全に動作します）。
              </p>
            </div>
          </div>

          <button
            onClick={handleToggleAccountEnabled}
            className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold rounded-lg shadow shrink-0 text-xs cursor-pointer"
          >
            認証完了したので再開する
          </button>
        </div>
      )}

      {/* Phone Checkpoint Safety & Simulation Info Box */}
      <div className="p-3.5 rounded-xl bg-zinc-950 border border-zinc-800/90 flex flex-wrap items-center justify-between gap-3 text-xs">
        <div className="flex items-center gap-2.5">
          <div className="p-1.5 rounded-lg bg-amber-950 text-amber-400 border border-amber-800/80">
            <Smartphone className="w-4 h-4" />
          </div>
          <div>
            <span className="font-bold text-white text-xs flex items-center gap-1.5">
              携帯電話番号認証（SMS確認）画面検知セーフティガード
              <span className="text-[10px] font-normal px-2 py-0.5 rounded bg-emerald-950 text-emerald-300 border border-emerald-800">
                自動防御ON
              </span>
            </span>
            <p className="text-zinc-400 text-[11px] mt-0.5">
              「携帯電話番号を入力」画面を検知すると、即座にブラウザを終了し「アカウント情報.xlsx」の【有効/無効】を「無効」に自動更新して安全停止します。
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={() => setShowCheckpointModal(true)}
            className="px-2.5 py-1 bg-zinc-800 hover:bg-zinc-700 text-zinc-300 border border-zinc-700 rounded-lg text-xs cursor-pointer"
          >
            検知対象画面を確認
          </button>

          {onSimulateCheckpoint && (
            <button
              onClick={() => onSimulateCheckpoint(currentAccount.id)}
              className="px-2.5 py-1 bg-amber-950/80 hover:bg-amber-900/80 text-amber-300 border border-amber-800 rounded-lg text-xs font-medium cursor-pointer"
              title="この画面を検知した際の自動停止フローをテスト実行"
            >
              🧪 自動停止テスト
            </button>
          )}
        </div>
      </div>

      {/* Redirect Loop Auto-Repair Info Box */}
      <div className="p-3.5 rounded-xl bg-gradient-to-r from-sky-950/50 via-zinc-950 to-indigo-950/40 border border-sky-800/60 flex flex-wrap items-center justify-between gap-3 text-xs">
        <div className="flex items-center gap-2.5">
          <div className="p-1.5 rounded-lg bg-sky-950 text-sky-400 border border-sky-800">
            <ShieldCheck className="w-4 h-4" />
          </div>
          <div>
            <span className="font-bold text-white text-xs flex items-center gap-1.5">
              リダイレクトループ (ERR_TOO_MANY_REDIRECTS) 自己修復エンジン
              <span className="text-[10px] font-normal px-2 py-0.5 rounded bg-emerald-950 text-emerald-300 border border-emerald-800 font-semibold">
                修復後投稿ON
              </span>
            </span>
            <p className="text-zinc-300 text-[11px] mt-0.5">
              「Cookieを削除してみてください」画面が出ても次アカウントへスキップせず、破損Cookie・キャッシュを自動クリアして直ちに自己修復＆投稿を完了させてから次のアカウントへ進みます。
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={() => setShowRedirectModal(true)}
            className="px-2.5 py-1 bg-sky-900/40 hover:bg-sky-800/60 text-sky-200 border border-sky-700/60 rounded-lg text-xs cursor-pointer"
          >
            修復対象画面を確認
          </button>
        </div>
      </div>

      {/* Checkpoint Screen Preview Modal */}
      {showCheckpointModal && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-zinc-900 border border-zinc-700 rounded-2xl max-w-lg w-full p-5 space-y-3 shadow-2xl text-xs">
            <div className="flex items-center justify-between border-b border-zinc-800 pb-2">
              <span className="font-bold text-sm text-white flex items-center gap-1.5">
                <AlertTriangle className="w-4 h-4 text-amber-400" />
                検知対象: Meta 携帯電話番号認証要求画面
              </span>
              <button
                onClick={() => setShowCheckpointModal(false)}
                className="text-zinc-400 hover:text-white font-mono text-sm"
              >
                ✕
              </button>
            </div>

            <div className="rounded-lg overflow-hidden border border-zinc-800 bg-white flex justify-center p-2">
              <img
                src="/assets/aistudio/image_fe10c3de.png"
                alt="携帯電話番号を入力画面"
                className="max-h-[320px] object-contain"
              />
            </div>

            <p className="text-zinc-300 leading-relaxed text-[11px]">
              この「携帯電話番号を入力」「SMSまたはWhatsAppで送信されたコードでこの携帯電話番号を認証する必要があります」という画面が表示された場合、
              <strong>アカウントを凍結から守るため自動で即座にブラウザを終了し、Excel上で【無効】に切り替えて安全に停止</strong>します。
            </p>

            <div className="flex justify-end pt-2 border-t border-zinc-800">
              <button
                onClick={() => setShowCheckpointModal(false)}
                className="px-4 py-1.5 bg-zinc-800 hover:bg-zinc-700 text-white rounded-lg cursor-pointer"
              >
                閉じる
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Redirect Error (ERR_TOO_MANY_REDIRECTS) Screen Preview Modal */}
      {showRedirectModal && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-zinc-900 border border-sky-700/80 rounded-2xl max-w-lg w-full p-5 space-y-3 shadow-2xl text-xs">
            <div className="flex items-center justify-between border-b border-zinc-800 pb-2">
              <span className="font-bold text-sm text-white flex items-center gap-1.5">
                <ShieldCheck className="w-4 h-4 text-sky-400" />
                自己修復対象: ERR_TOO_MANY_REDIRECTS 画面
              </span>
              <button
                onClick={() => setShowRedirectModal(false)}
                className="text-zinc-400 hover:text-white font-mono text-sm"
              >
                ✕
              </button>
            </div>

            <div className="rounded-lg overflow-hidden border border-zinc-800 bg-white flex justify-center p-2">
              <img
                src="/assets/aistudio/error_sll15802_20261103_182012_adbe9bf4.png"
                alt="このページは動作していません リダイレクトが繰り返し行われました"
                className="max-h-[320px] object-contain"
              />
            </div>

            <div className="space-y-1.5 text-zinc-300 leading-relaxed text-[11px]">
              <p>
                「www.threads.net でリダイレクトが繰り返し行われました。Cookie を削除してみてください」と表示されて停止していた問題の修復機能です。
              </p>
              <div className="p-2.5 rounded-lg bg-sky-950/60 border border-sky-800/60 text-sky-200">
                <strong>✨ 自己修復フロー:</strong><br />
                ① エラーを検知しても次アカウントへ飛ばずに即時捕捉<br />
                ② ブラウザ内の競合Cookie・循環キャッシュを完全クリア<br />
                ③ クリーンなセッションCookieを再注入してThreadsホームを再読込<br />
                ④ 該当ポスト（Post #2など）を正常に投稿完了させ、Excelを更新してから次アカウントへ進みます
              </div>
            </div>

            <div className="flex justify-end pt-2 border-t border-zinc-800">
              <button
                onClick={() => setShowRedirectModal(false)}
                className="px-4 py-1.5 bg-zinc-800 hover:bg-zinc-700 text-white rounded-lg cursor-pointer"
              >
                閉じる
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Add Account Panel */}
      {isAddingAccount && (
        <form onSubmit={handleCreateAccount} className="p-3 my-3 bg-zinc-950 rounded-lg border border-zinc-800 flex flex-wrap gap-2 items-center">
          <input
            type="text"
            placeholder="アカウントID (例: sll15802)"
            value={newAccountId}
            onChange={(e) => setNewAccountId(e.target.value)}
            className="bg-zinc-900 border border-zinc-700 text-xs text-white px-2.5 py-1.5 rounded focus:outline-none focus:border-emerald-500 flex-1 min-w-[160px]"
          />
          <input
            type="text"
            placeholder="プロファイル名 (省略時はID名: sll15802)"
            value={newProfileName}
            onChange={(e) => setNewProfileName(e.target.value)}
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
        <div className="grid grid-cols-1 md:grid-cols-4 gap-3 text-xs">
          {/* Chrome Profile & Persistent Login Status */}
          <div className="bg-zinc-950/60 p-2.5 rounded-lg border border-zinc-800/80">
            <div className="text-zinc-500 flex items-center justify-between gap-1 mb-1">
              <span className="flex items-center gap-1">
                <ShieldCheck className="w-3.5 h-3.5 text-emerald-400" />
                <span>永続Chromeプロファイル</span>
              </span>
              <span className="text-[10px] text-emerald-400 font-semibold">
                sessionid不要
              </span>
            </div>
            <div className="flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse" />
              <span className="font-mono text-zinc-200 truncate">
                {currentAccount.profileDir || `chrome_profiles/${currentAccount.id}`}
              </span>
            </div>
            <div className="text-[11px] text-zinc-400 mt-1">
              ID・PWログイン情報を永続保持 (強制ログアウト防止)
            </div>
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
        <div className="bg-zinc-950 p-3 rounded-lg border border-zinc-700 space-y-3 text-xs">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <div>
              <label className="text-zinc-400 font-medium block mb-1">Chromeプロファイルフォルダ名 (profile_launcher対応)</label>
              <input
                type="text"
                value={tempAccount.profileDir}
                onChange={(e) => setTempAccount({ ...tempAccount, profileDir: e.target.value })}
                placeholder="例: chrome_profiles/sll15802"
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
            <div>
              <label className="text-zinc-400 font-medium block mb-1">実行時間帯 (例: 9:00-23:00 または 0:00-0:00)</label>
              <input
                type="text"
                value={tempAccount.activeHours}
                onChange={(e) => setTempAccount({ ...tempAccount, activeHours: e.target.value })}
                className="w-full bg-zinc-900 border border-zinc-700 text-white text-xs px-2.5 py-1.5 rounded focus:outline-none focus:border-emerald-500 font-mono"
              />
            </div>
            <div>
              <label className="text-zinc-400 font-medium block mb-1">
                投稿モード
              </label>
              <select
                value={tempAccount.mode}
                onChange={(e) => setTempAccount({ ...tempAccount, mode: Number(e.target.value) as 0 | 1 })}
                className="w-full bg-zinc-900 border border-zinc-700 text-white text-xs px-2.5 py-1.5 rounded focus:outline-none"
              >
                <option value={1}>1: ランダム順序で投稿</option>
                <option value={0}>0: 投稿リストの上から順に投稿</option>
              </select>
            </div>
          </div>
          <div className="flex justify-end gap-2 pt-1 border-t border-zinc-800">
            <button
              type="button"
              onClick={() => setIsEditingSettings(false)}
              className="px-3 py-1.5 text-zinc-400 hover:text-white cursor-pointer"
            >
              キャンセル
            </button>
            <button
              type="button"
              onClick={handleSaveSettings}
              className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold rounded cursor-pointer"
            >
              変更を保存
            </button>
          </div>
        </div>
      )}
    </div>
  );
};
