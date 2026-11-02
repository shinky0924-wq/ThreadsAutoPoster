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

export const DownloadCenter: React.FC = () => {
  const [downloadSuccess, setDownloadSuccess] = useState<string | null>(null);
  const [copied, setCopied] = useState(false);

  const rawGithubPyUrl = "https://raw.githubusercontent.com/shinky0924-wq/ThreadsAutoPoster/main/ThreadsAutoPoster_Windows_Package%20(36)/dist/ThreadsAutoPoster/threads_gui_app.py";

  // Instant 100% clean Blob download directly from client-side memory
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

  const copyScriptToClipboard = () => {
    navigator.clipboard.writeText(RAW_THREADS_GUI_APP_PY);
    setCopied(true);
    setTimeout(() => setCopied(false), 3000);
  };

  const downloadExcelFile = (url: string, filename: string) => {
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
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
                  新機能「📋 100件一括追加・削除」ダイアログ、アカウント別全件削除、凍結防止スピンタックス対応
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
              <span>今すぐダウンロード (推奨)</span>
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
                  マスター設定
                </span>
              </div>
              <p className="text-zinc-400 text-[11px]">
                アカウントID、auth_token (sessionid)、備考を管理するマスターエクセルファイル。
              </p>
            </div>
            <div className="flex justify-end pt-2 border-t border-zinc-800/80">
              <a
                href="/download/アカウント情報.xlsx"
                download="アカウント情報.xlsx"
                className="flex items-center gap-1.5 px-3 py-1.5 bg-zinc-800 hover:bg-zinc-700 text-white rounded-lg text-xs"
              >
                <Download className="w-3.5 h-3.5" />
                <span>ダウンロード</span>
              </a>
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

      {/* Guide */}
      <div className="p-4 rounded-xl bg-zinc-950 border border-zinc-800 space-y-2">
        <h4 className="font-bold text-zinc-200 text-xs flex items-center gap-1.5">
          <HelpCircle className="w-4 h-4 text-emerald-400" />
          上書き配置の手順:
        </h4>
        <p className="text-zinc-400 text-[11px] leading-relaxed">
          ダウンロードした <code className="text-white bg-zinc-900 px-1 py-0.5 rounded font-mono">threads_gui_app.py</code> を、以下の場所に上書きしてください：<br />
          <code className="text-emerald-400 font-mono block mt-1">C:\Users\PC_User\Desktop\ThreadsAutoPoster_Windows_Package (36)\dist\ThreadsAutoPoster\threads_gui_app.py</code>
        </p>
      </div>
    </div>
  );
};
