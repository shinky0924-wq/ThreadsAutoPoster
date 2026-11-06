import React, { useState } from 'react';
import { 
  Download, 
  FileCode2, 
  FileSpreadsheet, 
  Image as ImageIcon, 
  Check, 
  Copy, 
  ArrowRight, 
  Sparkles,
  HelpCircle,
  FolderOpen,
  CheckCircle2,
  ExternalLink,
  ShieldCheck,
  AlertTriangle
} from 'lucide-react';
import { RAW_THREADS_GUI_APP_PY } from '../data/rawScript';
import { RAW_PROFILE_LAUNCHER_PY } from '../data/rawProfileLauncher';
import { AccountInfo } from '../types';
import { exportAccountsConfigWorkbook } from '../utils/excelUtils';

interface DownloadCenterProps {
  accounts?: AccountInfo[];
}

export const DownloadCenter: React.FC<DownloadCenterProps> = ({ accounts = [] }) => {
  const [downloadSuccess, setDownloadSuccess] = useState<string | null>(null);
  const [copied, setCopied] = useState(false);
  const [copiedLauncher, setCopiedLauncher] = useState(false);

  // Instant 100% clean Blob download for threads_gui_app.py
  const downloadPythonDirectly = () => {
    try {
      const blob = new Blob([RAW_THREADS_GUI_APP_PY], { type: 'text/x-python;charset=utf-8' });
      const blobUrl = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = blobUrl;
      a.download = 'threads_gui_app.py';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(blobUrl);
      setDownloadSuccess('threads_gui_app.py (最新版) のダウンロードを開始しました！');
      setTimeout(() => setDownloadSuccess(null), 4000);
    } catch (err: any) {
      console.error(err);
    }
  };

  // Instant 100% clean Blob download for profile_launcher.py
  const downloadProfileLauncherDirectly = () => {
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
      setDownloadSuccess('profile_launcher.py (ログイン・プロファイル生成ツール) のダウンロードを開始しました！');
      setTimeout(() => setDownloadSuccess(null), 4000);
    } catch (err: any) {
      console.error(err);
    }
  };

  const copyScriptToClipboard = () => {
    navigator.clipboard.writeText(RAW_THREADS_GUI_APP_PY);
    setCopied(true);
    setTimeout(() => setCopied(false), 3000);
  };

  const copyLauncherToClipboard = () => {
    navigator.clipboard.writeText(RAW_PROFILE_LAUNCHER_PY);
    setCopiedLauncher(true);
    setTimeout(() => setCopiedLauncher(false), 3000);
  };

  const downloadExcelFile = (url: string, filename: string) => {
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  const downloadAccountsExcelDirectly = () => {
    try {
      if (accounts && accounts.length > 0) {
        exportAccountsConfigWorkbook(accounts);
      } else {
        downloadExcelFile('/download/アカウント情報.xlsx', 'アカウント情報.xlsx');
      }
      setDownloadSuccess('アカウント情報.xlsx (ID/PW・sessionid不要・プロファイル連携版) のダウンロードを開始しました！');
      setTimeout(() => setDownloadSuccess(null), 4000);
    } catch (err: any) {
      console.error(err);
      downloadExcelFile('/download/アカウント情報.xlsx', 'アカウント情報.xlsx');
    }
  };

  return (
    <div className="bg-zinc-900 border border-zinc-800 rounded-xl p-5 space-y-6 shadow-sm text-xs">
      {/* Header Banner */}
      <div className="pb-4 border-b border-zinc-800 flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-2">
          <div className="p-2 rounded-xl bg-indigo-950 text-indigo-400 border border-indigo-800">
            <FolderOpen className="w-5 h-5" />
          </div>
          <div>
            <h3 className="text-sm font-bold text-white flex items-center gap-2">
              Windows既存フォルダ更新用ファイル (threads_gui_app.py & エクセル)
            </h3>
            <p className="text-zinc-400 text-xs mt-0.5">
              この画面から最新の <span className="font-mono text-indigo-300 font-bold">threads_gui_app.py</span> を直接ダウンロードできます
            </p>
          </div>
        </div>
      </div>

      {downloadSuccess && (
        <div className="p-3.5 rounded-xl bg-emerald-950/80 border border-emerald-800 text-emerald-300 flex items-center gap-2 font-medium">
          <CheckCircle2 className="w-5 h-5 text-emerald-400 shrink-0" />
          <span>{downloadSuccess}</span>
        </div>
      )}

      {/* 🌟 Primary Feature Card: profile_launcher.py (NEW TOOL) */}
      <div className="p-5 rounded-2xl bg-gradient-to-br from-violet-950/80 via-zinc-950 to-zinc-950 border-2 border-violet-500/80 shadow-2xl space-y-4">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div className="space-y-1.5">
            <div className="flex items-center gap-2">
              <span className="p-2 rounded-lg bg-violet-600 text-white shadow-lg">
                <Sparkles className="w-6 h-6" />
              </span>
              <div>
                <h4 className="text-base font-bold text-white flex items-center gap-2">
                  profile_launcher.py (アプリ名: profile_luncher)
                  <span className="text-[11px] font-bold px-2.5 py-0.5 rounded-full bg-emerald-950 text-emerald-300 border border-emerald-700 font-mono">
                    NEW 永続化ツール
                  </span>
                </h4>
                <p className="text-zinc-300 text-xs">
                  exe/スクリプト起動 ➔ プロファイル名入力 ➔ 専用Chrome起動 ➔ ID/PWでログインして閉じるだけ！
                </p>
              </div>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={copyLauncherToClipboard}
              className="flex items-center gap-1.5 px-3 py-2 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 border border-zinc-700 rounded-xl font-medium cursor-pointer transition-all"
              title="クリップボードに全コードをコピー"
            >
              {copiedLauncher ? <Check className="w-4 h-4 text-emerald-400" /> : <Copy className="w-4 h-4" />}
              <span>{copiedLauncher ? 'コピー完了！' : 'コードをコピー'}</span>
            </button>

            <button
              onClick={downloadProfileLauncherDirectly}
              className="flex items-center gap-2 px-5 py-2.5 bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-violet-500 hover:to-indigo-500 text-white font-bold rounded-xl shadow-lg hover:shadow-violet-900/40 text-sm cursor-pointer transition-all active:scale-95 border border-violet-400/50 animate-pulse"
            >
              <Download className="w-4 h-4" />
              <span>🚀 profile_launcher.py をダウンロード</span>
            </button>
          </div>
        </div>

        {/* Feature Highlights Grid */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-2.5 pt-2">
          <div className="p-2.5 rounded-lg bg-zinc-900/90 border border-zinc-800 text-zinc-300">
            <div className="font-bold text-violet-300 mb-0.5">① sessionid不要・永久保存</div>
            <div className="text-[11px] text-zinc-400">
              Cookie・IndexedDBトークンがローカルフォルダに完全保持され、Metaの定期強制ログアウトを防ぎます。
            </div>
          </div>
          <div className="p-2.5 rounded-lg bg-zinc-900/90 border border-zinc-800 text-zinc-300">
            <div className="font-bold text-sky-300 mb-0.5">② 複数アカウント完全分離</div>
            <div className="text-[11px] text-zinc-400">
              <code className="text-white">chrome_profiles/sll15802</code> 等、アカウントごとに独立した環境で誤爆や混ざりを防止。
            </div>
          </div>
          <div className="p-2.5 rounded-lg bg-zinc-900/90 border border-zinc-800 text-zinc-300">
            <div className="font-bold text-emerald-300 mb-0.5">③ Excelワンクリック登録</div>
            <div className="text-[11px] text-zinc-400">
              ツール内の「📋 アカウント情報.xlsxに登録」ボタンで、ID・PWを一切書かずに安全にExcel連携。
            </div>
          </div>
        </div>
      </div>

      {/* Main Feature Download Box: threads_gui_app.py */}
      <div className="p-5 rounded-2xl bg-gradient-to-br from-indigo-950/60 via-zinc-950 to-zinc-950 border-2 border-indigo-600/80 shadow-xl space-y-4">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div className="space-y-1.5">
            <div className="flex items-center gap-2">
              <span className="p-2 rounded-lg bg-indigo-600 text-white">
                <FileCode2 className="w-6 h-6" />
              </span>
              <div>
                <h4 className="text-base font-bold text-white flex items-center gap-2">
                  threads_gui_app.py <span className="text-[11px] font-normal px-2 py-0.5 rounded-full bg-emerald-950 text-emerald-300 border border-emerald-800 font-mono">最新版 (231 KB)</span>
                </h4>
                <p className="text-zinc-400 text-xs">
                  左下「＋」＆タイムライン上部「最近どう？/投稿」の2段階フェイルセーフ、2重投稿防止、永続プロファイル完全対応
                </p>
              </div>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={copyScriptToClipboard}
              className="flex items-center gap-1.5 px-3.5 py-2.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 border border-zinc-700 rounded-xl font-medium cursor-pointer transition-all"
              title="クリップボードに全コードをコピー"
            >
              {copied ? <Check className="w-4 h-4 text-emerald-400" /> : <Copy className="w-4 h-4" />}
              <span>{copied ? 'コピー完了！' : 'コードを全文コピー'}</span>
            </button>

            <button
              onClick={downloadPythonDirectly}
              className="flex items-center gap-2 px-5 py-2.5 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-bold rounded-xl shadow-lg hover:shadow-emerald-900/40 text-sm cursor-pointer transition-all active:scale-95"
            >
              <Download className="w-4 h-4" />
              <span>🐍 threads_gui_app.py をダウンロード</span>
            </button>
          </div>
        </div>

        <div className="p-3 rounded-xl bg-zinc-900/80 border border-zinc-800 flex items-start gap-2.5 text-zinc-300 text-[11px] leading-relaxed">
          <ShieldCheck className="w-4 h-4 text-emerald-400 shrink-0 mt-0.5" />
          <div>
            <span className="font-bold text-emerald-300">HTML混入防止済（純粋なPythonスクリプト）:</span><br />
            ブラウザメモリから直接バイナリ生成するため、Cookie認証HTMLが混入することなく、純粋なPythonコードとしてダウンロードされます。
          </div>
        </div>
      </div>

      {/* Excel Files Download Grid */}
      <div className="space-y-3">
        <h4 className="font-bold text-zinc-300 text-xs">エクセル設定ファイル一覧 (更新用):</h4>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
          <div className="p-4 rounded-xl border bg-zinc-950/70 border-zinc-800 flex flex-col justify-between gap-3">
            <div className="space-y-1.5">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <FileSpreadsheet className="w-5 h-5 text-emerald-400" />
                  <span className="font-mono font-bold text-white text-xs">アカウント情報.xlsx</span>
                </div>
                <span className="px-2 py-0.5 rounded text-[10px] font-mono border border-emerald-800 bg-emerald-950/30 text-emerald-300">
                  ID/PW・sessionid削除済
                </span>
              </div>
              <p className="text-zinc-400 text-[11px]">
                列構成: <code className="text-white">['アカウントID', 'プロファイル名', '有効/無効', '備考']</code>。<br />
                パスワードやsessionidは一切記載不要で安全に運用できます。
              </p>
            </div>
            <div className="flex justify-end pt-2 border-t border-zinc-800/80">
              <button
                onClick={downloadAccountsExcelDirectly}
                className="flex items-center gap-1.5 px-3.5 py-2 bg-gradient-to-r from-sky-600 to-indigo-600 hover:from-sky-500 hover:to-indigo-500 text-white font-bold rounded-lg text-xs cursor-pointer shadow-md transition-all active:scale-95"
              >
                <Download className="w-3.5 h-3.5" />
                <span>📊 アカウント情報.xlsx をダウンロード</span>
              </button>
            </div>
          </div>

          <div className="p-4 rounded-xl border bg-zinc-950/70 border-zinc-800 flex flex-col justify-between gap-3">
            <div className="space-y-1.5">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <FileSpreadsheet className="w-5 h-5 text-sky-400" />
                  <span className="font-mono font-bold text-white text-xs">sll15802.xlsx</span>
                </div>
                <span className="px-2 py-0.5 rounded text-[10px] font-mono border border-sky-800 bg-sky-950/30 text-sky-300">
                  メイン投稿
                </span>
              </div>
              <p className="text-zinc-400 text-[11px]">
                メインアカウントの設定・投稿リスト・日常ポスト枠が入ったエクセルファイル。
              </p>
            </div>
            <div className="flex justify-end pt-2 border-t border-zinc-800/80">
              <a
                href="/download/sll15802.xlsx"
                download="sll15802.xlsx"
                className="flex items-center gap-1.5 px-3 py-1.5 bg-zinc-800 hover:bg-zinc-700 text-white rounded-lg text-xs"
              >
                <Download className="w-3.5 h-3.5" />
                <span>ダウンロード</span>
              </a>
            </div>
          </div>
        </div>
      </div>

      {/* Guide: Step by step workflow */}
      <div className="p-4 rounded-xl bg-zinc-950 border border-zinc-800 space-y-3">
        <h4 className="font-bold text-zinc-200 text-xs flex items-center gap-1.5">
          <HelpCircle className="w-4 h-4 text-emerald-400" />
          ✨ 新しい初期設定＆ログイン手順（強制ログアウト防止）:
        </h4>
        <ol className="list-decimal list-inside space-y-1.5 text-zinc-300 text-[11px] leading-relaxed pl-1">
          <li>
            ダウンロードした <code className="text-violet-300 bg-zinc-900 px-1.5 py-0.5 rounded font-mono">profile_launcher.py</code> と <code className="text-emerald-400 bg-zinc-900 px-1.5 py-0.5 rounded font-mono">threads_gui_app.py</code>、<code className="text-sky-400 bg-zinc-900 px-1.5 py-0.5 rounded font-mono">アカウント情報.xlsx</code> を同じフォルダに配置します。
          </li>
          <li>
            <code className="text-violet-300 font-mono">profile_launcher.py</code>（または生成されたEXE）を起動します。
          </li>
          <li>
            プロファイル名にアカウント名（例: <code className="text-white font-mono">sll15802</code>）を入力し、緑色の「<strong>🚀 このプロファイルでChromeを起動（ログイン用）</strong>」ボタンをクリックします。
          </li>
          <li>
            開いた専用Chromeウィンドウで、ThreadsにIDとパスワードで1回通常ログインします。
          </li>
          <li>
            ログイン完了後、Chromeウィンドウをそのまま閉じます。（これで <code className="text-zinc-300 font-mono">chrome_profiles/sll15802</code> にログイン状態が永久保存されます）
          </li>
          <li>
            「<strong>📋 アカウント情報.xlsxに登録</strong>」を押すと、ID・PW不要でExcelに自動登録されます。
          </li>
          <li>
            あとは通常どおり <code className="text-emerald-400 font-mono">threads_gui_app.py</code>（またはThreadsAutoPoster.exe）を起動して自動投稿を開始すれば、永続プロファイルで安全に巡回・投稿されます！
          </li>
        </ol>
      </div>
    </div>
  );
};
