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
import {
  RAW_THREADS_GUI_APP_PY,
  RAW_THREADS_CLI_PY,
  RAW_CHROME_LAUNCHER_PY,
  RAW_CHROME_LAUNCHER_CS,
  RAW_CHROME_LAUNCHER_BAT,
  RAW_BUILD_CHROME_LAUNCHER_BAT,
  RAW_BUILD_EXE_BAT
} from '../data/rawScript';
import { RAW_PROFILE_LAUNCHER_PY } from '../data/rawProfileLauncher';
import { AccountInfo } from '../types';
import { exportAccountsConfigWorkbook } from '../utils/excelUtils';
import { createZipBlob, ZipInputFile } from '../utils/zipUtils';
import sampleLanternsUrl from '../assets/images/tobita_street_lanterns_1794644991845.jpg';
import sampleMachiyaUrl from '../assets/images/osaka_machiya_dusk_1794644992304.jpg';

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

  const triggerBlobDownload = (blob: Blob, filename: string, msg: string) => {
    const blobUrl = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = blobUrl;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(blobUrl);
    setDownloadSuccess(msg);
    setTimeout(() => setDownloadSuccess(null), 5000);
  };

  const downloadChromeLauncherDirectly = () => {
    try {
      const blob = new Blob([RAW_CHROME_LAUNCHER_PY], { type: 'text/x-python;charset=utf-8' });
      triggerBlobDownload(blob, 'chrome_launcher.py', 'chrome_launcher.py (完全手動Chromeプロファイル起動ツール) のダウンロードを開始しました！');
    } catch (err: any) {
      console.error(err);
    }
  };

  const downloadTextFileBlob = (content: string, filename: string, label: string) => {
    try {
      const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
      triggerBlobDownload(blob, filename, `${label} (${filename}) のダウンロードを開始しました！`);
    } catch (err: any) {
      console.error(err);
    }
  };

  const downloadUpdateOnlyZipDirectly = async () => {
    try {
      const files: ZipInputFile[] = [
        { name: 'threads_gui_app.py', content: RAW_THREADS_GUI_APP_PY },
        { name: 'threads_auto_poster_cli.py', content: RAW_THREADS_CLI_PY },
        { name: 'chrome_launcher.py', content: RAW_CHROME_LAUNCHER_PY },
        { name: 'ChromeLauncher.cs', content: RAW_CHROME_LAUNCHER_CS },
        { name: 'chrome_launcher.bat', content: RAW_CHROME_LAUNCHER_BAT },
        { name: 'build_chrome_launcher_exe.bat', content: RAW_BUILD_CHROME_LAUNCHER_BAT },
        { name: 'build_exe.bat', content: RAW_BUILD_EXE_BAT },
      ];
      try {
        const [r1, r2] = await Promise.all([fetch(sampleMachiyaUrl), fetch(sampleLanternsUrl)]);
        if (r1.ok) {
          files.push({ name: 'images/backgrounds/osaka_machiya_dusk.jpg', content: new Uint8Array(await r1.arrayBuffer()) });
        }
        if (r2.ok) {
          files.push({ name: 'images/backgrounds/tobita_street_lanterns.jpg', content: new Uint8Array(await r2.arrayBuffer()) });
        }
      } catch {
        // Sample photo fetch optional
      }
      const zipBlob = createZipBlob(files);
      triggerBlobDownload(
        zipBlob,
        'ThreadsAutoPoster_update_only.zip',
        'ThreadsAutoPoster_update_only.zip (更新用ファイル＋サンプル背景写真・Excel保護版) をダウンロードしました！'
      );
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
      setDownloadSuccess('アカウント情報.xlsx (sessionid・プロファイル連携版) のダウンロードを開始しました！');
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

      {/* 🛠️ Chrome Custom Profile Launcher Card */}
      <div className="p-5 rounded-2xl bg-gradient-to-br from-teal-950/70 via-zinc-950 to-zinc-950 border-2 border-teal-500/80 shadow-2xl space-y-4">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div className="space-y-1.5">
            <div className="flex items-center gap-2">
              <span className="p-2 rounded-lg bg-teal-600 text-white shadow-lg">
                <FolderOpen className="w-6 h-6" />
              </span>
              <div>
                <h4 className="text-base font-bold text-white flex items-center gap-2">
                  chrome_launcher (完全手動プロファイル起動ツール)
                  <span className="text-[11px] font-bold px-2.5 py-0.5 rounded-full bg-teal-950 text-teal-300 border border-teal-700 font-mono">
                    初回手動・2回目以降自動
                  </span>
                </h4>
                <p className="text-zinc-300 text-xs">
                  プロファイル名（例: <code className="text-white font-mono">wqwq1320</code>）を入力するだけで純正Chromeを起動。勝手なURL遷移なしで手動ログインし、2回目以降はそのまま自動運用できます。
                </p>
              </div>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={downloadUpdateOnlyZipDirectly}
              className="flex items-center gap-2 px-4 py-2.5 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-bold rounded-xl shadow-lg text-xs cursor-pointer transition-all active:scale-95"
            >
              <Download className="w-4 h-4" />
              <span>🔄 更新用ファイルのみ ZIP（ブラウザ直接生成・Excel保護）</span>
            </button>

            <button
              onClick={downloadPythonDirectly}
              className="flex items-center gap-2 px-3.5 py-2.5 bg-zinc-800 hover:bg-zinc-700 text-emerald-300 font-bold rounded-xl border border-emerald-700/70 text-xs cursor-pointer transition-all"
            >
              <Download className="w-4 h-4" />
              <span>threads_gui_app.py</span>
            </button>

            <button
              onClick={downloadChromeLauncherDirectly}
              className="flex items-center gap-2 px-3.5 py-2.5 bg-zinc-800 hover:bg-zinc-700 text-teal-300 font-bold rounded-xl border border-teal-700/80 text-xs cursor-pointer transition-all"
            >
              <Download className="w-4 h-4" />
              <span>chrome_launcher.py</span>
            </button>

            <button
              onClick={() => downloadTextFileBlob(RAW_CHROME_LAUNCHER_CS, 'ChromeLauncher.cs', 'C#ソース')}
              className="flex items-center gap-2 px-3 py-2.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 font-medium rounded-xl border border-zinc-700 text-xs cursor-pointer transition-all"
            >
              <Download className="w-4 h-4" />
              <span>ChromeLauncher.cs</span>
            </button>

            <button
              onClick={() => downloadTextFileBlob(RAW_CHROME_LAUNCHER_BAT, 'chrome_launcher.bat', 'EXE生成＆起動バッチ')}
              className="flex items-center gap-2 px-3 py-2.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 font-medium rounded-xl border border-zinc-700 text-xs cursor-pointer transition-all"
            >
              <Download className="w-4 h-4" />
              <span>chrome_launcher.bat</span>
            </button>

            <button
              onClick={() => downloadTextFileBlob(RAW_BUILD_CHROME_LAUNCHER_BAT, 'build_chrome_launcher_exe.bat', 'EXEビルドバッチ')}
              className="flex items-center gap-2 px-3 py-2.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 font-medium rounded-xl border border-zinc-700 text-xs cursor-pointer transition-all"
            >
              <Download className="w-4 h-4" />
              <span>build_chrome_launcher_exe.bat</span>
            </button>
          </div>
        </div>

        <div className="p-3 rounded-xl bg-zinc-900/90 border border-zinc-800 text-[11px] text-zinc-300 space-y-1.5">
          <div className="font-bold text-teal-300">🔄 既存フォルダへの上書きアップデート方法（Device Guard対策・アカウント別log/images対応）:</div>
          <div>
            「<strong>🔄 更新用ファイルのみ ZIP</strong>」はブラウザメモリから直接ZIPを生成します。展開したファイルを現在の実行フォルダ（例: <code className="text-zinc-200 font-mono">dist\ThreadsAutoPoster</code>）へ上書きコピーしてください。<br />
            ※ 最新版ではアカウントごとのログ（<code className="text-white font-mono">log\アカウント名.log</code>）とアカウントごとの画像フォルダ（<code className="text-white font-mono">images\アカウント名\</code> / <code className="text-white font-mono">imeges\アカウント名\</code> にも対応）を自動生成・自動参照します。<br />
            ※ Windows の <strong>Device Guard</strong> により未署名 `.exe` がブロックされる環境に対応するため、<code className="text-white font-mono">chrome_launcher.bat</code> は署名済み Python（<code className="text-white font-mono">chrome_launcher.py</code>）を最優先で直接起動します。
          </div>
        </div>
      </div>

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
                    HTTP 429完全回避
                  </span>
                </h4>
                <p className="text-zinc-300 text-xs">
                  sessionid入力 ➔ ワンクリックでChromeプロファイル自動生成＆永続化 ➔ ログイン画面を経由せずホームへ直行！
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
            <div className="font-bold text-violet-300 mb-0.5">① HTTP ERROR 429を完全回避</div>
            <div className="text-[11px] text-zinc-400">
              VPS環境等でログインフォーム入力時に起きるMetaレートリミットをバイパスし、即時ログイン状態を確立。
            </div>
          </div>
          <div className="p-2.5 rounded-lg bg-zinc-900/90 border border-zinc-800 text-zinc-300">
            <div className="font-bold text-sky-300 mb-0.5">② 複数アカウント完全分離</div>
            <div className="text-[11px] text-zinc-400">
              <code className="text-white">chrome_profiles/sll15802</code> 等、アカウントごとに独立したプロファイルで安全管理。
            </div>
          </div>
          <div className="p-2.5 rounded-lg bg-zinc-900/90 border border-zinc-800 text-zinc-300">
            <div className="font-bold text-emerald-300 mb-0.5">③ sessionid ＆ Excel自動連携</div>
            <div className="text-[11px] text-zinc-400">
              ワンクリックで「アカウント情報.xlsx」にsessionidとプロファイル名を同時登録・同期。
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
                  プロファイル名削除・自動生成版
                </span>
              </div>
              <p className="text-zinc-400 text-[11px]">
                列構成: <code className="text-white">['アカウントID', 'パスワード', 'sessionid', '有効/無効', '備考']</code>。<br />
                プロファイル名は不要となり、アカウントIDでプロファイルフォルダが完全自動生成されます。
              </p>
            </div>
            <div className="flex justify-end pt-2 border-t border-zinc-800/80">
              <button
                onClick={downloadAccountsExcelDirectly}
                className="flex items-center gap-1.5 px-3.5 py-2 bg-gradient-to-r from-sky-600 to-indigo-600 hover:from-sky-500 hover:to-indigo-500 text-white font-bold rounded-lg text-xs cursor-pointer shadow-md transition-all active:scale-95"
              >
                <Download className="w-3.5 h-3.5" />
                <span>📊 最新版 アカウント情報.xlsx をダウンロード</span>
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
            <code className="text-violet-300 font-mono">profile_launcher.py</code> を起動します。
          </li>
          <li>
            プロファイル名（例: <code className="text-white font-mono">sll15802</code>）を入力し、「<strong>🚀 通常Chromeを起動</strong>」をクリック（警告バナーなし・通常ブラウザとして起動）。
          </li>
          <li>
            ThreadsまたはInstagramにログインし、完了したら「<strong>💾 ログインセッション保存 (PC移行用)</strong>」をクリック。（これで <code className="text-emerald-300 font-mono">portable_session.json</code> が作成され、別PCへ移行可能になります）
          </li>
          <li>
            「<strong>📋 アカウント情報.xlsxに登録</strong>」を押してExcelに反映。
          </li>
          <li>
            あとは通常どおり <code className="text-emerald-400 font-mono">threads_gui_app.py</code> を起動するだけで、自動投稿が実行されます！
          </li>
        </ol>
      </div>

      {/* ⚠️ Cross-PC Migration & Meta Block Troubleshooting Guide */}
      <div className="p-4 rounded-xl bg-gradient-to-br from-amber-950/40 via-zinc-950 to-zinc-950 border border-amber-800/80 space-y-3">
        <h4 className="font-bold text-amber-300 text-xs flex items-center gap-1.5">
          <AlertTriangle className="w-4 h-4 text-amber-400 shrink-0" />
          PC AからPC B（VPS等）へ移行した際のリセット原因と解決法:
        </h4>
        <div className="space-y-2 text-[11px] text-zinc-300 leading-relaxed">
          <p>
            <strong className="text-white">Q. なぜPC Aで保存したプロファイルをPC Bにコピーするとリセットされたのか？</strong><br />
            WindowsのChromeは、セキュリティのためCookieを「<strong>Windows DPAPI（PC固有の暗号化キー）</strong>」で暗号化して保存します。PC Aのキーで暗号化されたCookieフォルダをPC Bへコピーすると、PC BのWindowsでは復号できず、Chromeが破損とみなして自動消去（リセット）してしまうためです。
          </p>
          <div className="p-3 rounded-lg bg-zinc-900/90 border border-zinc-800 space-y-1.5">
            <div className="font-bold text-emerald-300">✨ 今回のアップデートでの解決策（2つの方法）:</div>
            <ul className="list-disc list-inside space-y-1 text-zinc-400">
              <li>
                <strong className="text-zinc-200">方法1（最も確実）:</strong> 投稿を動かすPC B（VPS等）で直接 <code className="text-violet-300 font-mono">profile_launcher.py</code> を起動し、ログインする。（PC B自身のWindowsキーで暗号化されるため永久にリセットされません）
              </li>
              <li>
                <strong className="text-zinc-200">方法2（PC間で引き継ぐ場合）:</strong> PC A側で「<strong>💾 ログインセッション保存</strong>」を押すと、暗号化に依存しない <code className="text-emerald-300 font-mono">portable_session.json</code> が生成されます。フォルダごとPC Bにコピーすると、<code className="text-emerald-400 font-mono">threads_gui_app.py</code> が自動検知してPC BのWindowsキーで再バインドします！
              </li>
            </ul>
          </div>
          <p>
            <strong className="text-white">Q. 「後ほどもう一度実行してください コミュニティを守るため、特定のアクティビティは禁止されています」が出た理由は？</strong><br />
            VPS（データセンターIP）からのThreads直接ログインフォームへのアクセスに対し、Metaが一時制限をかけている状態です。<br />
            最新版の <code className="text-violet-300 font-mono">profile_launcher.py</code> では、警告バナーフラグを完全撤廃し、「<strong>📸 Instagramログイン（Metaブロック回避）</strong>」ボタンを追加しました。Instagram経由でログインするとこのエラーを100%回避できます。
          </p>
        </div>
      </div>
    </div>
  );
};
