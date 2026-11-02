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
  AlertTriangle
} from 'lucide-react';

export const DownloadCenter: React.FC = () => {
  const [downloadingId, setDownloadingId] = useState<string | null>(null);
  const [downloadError, setDownloadError] = useState<string | null>(null);

  const rawGithubPyUrl = "https://raw.githubusercontent.com/shinky0924-wq/ThreadsAutoPoster/main/ThreadsAutoPoster_Windows_Package%20(36)/dist/ThreadsAutoPoster/threads_gui_app.py";
  const githubZipUrl = "https://github.com/shinky0924-wq/ThreadsAutoPoster/archive/refs/heads/main.zip";

  const downloadBlobSecurely = async (url: string, filename: string, id: string) => {
    setDownloadingId(id);
    setDownloadError(null);
    try {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`HTTP Error: ${res.status}`);
      const text = await res.text();

      // Verify that the response is NOT an HTML error or Cookie Check page
      if (text.includes('<!doctype html>') || text.includes('<title>Cookie check</title>')) {
        throw new Error('ブラウザのCookie認証ページが混入しました。GitHub直接ダウンロードリンクをご利用ください。');
      }

      const blob = new Blob([text], { type: filename.endsWith('.py') ? 'text/x-python;charset=utf-8' : 'application/octet-stream' });
      const blobUrl = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = blobUrl;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(blobUrl);
    } catch (err: any) {
      console.error(err);
      setDownloadError(`${filename} のダウンロード中にエラーが発生しました: ${err.message}`);
    } finally {
      setDownloadingId(null);
    }
  };

  const files = [
    {
      id: 'py',
      name: 'threads_gui_app.py',
      size: '226 KB',
      desc: '【EXE更新用】新機能「📋 100件一括追加・削除」ダイアログ、アカウント別全件削除、凍結防止スピンタックス対応の完全版GUIスクリプト。',
      url: '/download/threads_gui_app.py',
      icon: <FileCode2 className="w-5 h-5 text-indigo-400" />,
      color: 'border-indigo-800 bg-indigo-950/30 text-indigo-300',
      badge: '超重要・最新版',
      isPy: true,
    },
    {
      id: 'acc',
      name: 'アカウント情報.xlsx',
      size: '5.4 KB',
      desc: 'アカウントID、auth_token (sessionid)、備考を管理するマスターエクセルファイル。',
      url: '/download/アカウント情報.xlsx',
      icon: <FileSpreadsheet className="w-5 h-5 text-emerald-400" />,
      color: 'border-emerald-800 bg-emerald-950/30 text-emerald-300',
      badge: 'マスター設定',
      isPy: false,
    },
    {
      id: 'sll',
      name: 'sll15802.xlsx',
      size: '12 KB',
      desc: 'メインアカウントの設定・投稿リスト・日常ポスト枠が入ったエクセルファイル。',
      url: '/download/sll15802.xlsx',
      icon: <FileSpreadsheet className="w-5 h-5 text-sky-400" />,
      color: 'border-sky-800 bg-sky-950/30 text-sky-300',
      badge: 'メイン投稿',
      isPy: false,
    },
    {
      id: 'jht',
      name: 'jhtwj2.xlsx',
      size: '9.3 KB',
      desc: 'サブアカウント1用の個別投稿・設定エクセルファイル。',
      url: '/download/jhtwj2.xlsx',
      icon: <FileSpreadsheet className="w-5 h-5 text-teal-400" />,
      color: 'border-teal-800 bg-teal-950/30 text-teal-300',
      badge: 'サブアカウント1',
      isPy: false,
    },
    {
      id: 'wql',
      name: 'wqlopakpskapkoq.xlsx',
      size: '9.3 KB',
      desc: 'サブアカウント2用の個別投稿・設定エクセルファイル。',
      url: '/download/wqlopakpskapkoq.xlsx',
      icon: <FileSpreadsheet className="w-5 h-5 text-purple-400" />,
      color: 'border-purple-800 bg-purple-950/30 text-purple-300',
      badge: 'サブアカウント2',
      isPy: false,
    },
  ];

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
              既存のWindowsフォルダ（ThreadsAutoPoster.exeがある場所）に上書きするだけで更新できます
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <a
            href={githubZipUrl}
            target="_blank"
            rel="noreferrer"
            className="flex items-center gap-1 px-3 py-1.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 border border-zinc-700 rounded-lg text-xs"
          >
            <Download className="w-3.5 h-3.5" />
            <span>GitHub一式ZIPダウンロード</span>
          </a>
        </div>
      </div>

      {/* Warning Box explaining why HTML got into the py file */}
      <div className="p-4 rounded-xl bg-amber-950/30 border border-amber-800/80 space-y-2">
        <div className="flex items-center gap-2 text-amber-300 font-bold">
          <AlertTriangle className="w-4 h-4 text-amber-400 shrink-0" />
          <span>【重要】ブラウザによる「Cookie check HTML」の誤混入を完全防止</span>
        </div>
        <p className="text-amber-200/90 leading-relaxed text-[11px]">
          ブラウザによってはダウンロード時にAI Studioの認証HTML（&lt;!doctype html&gt;...）が混ざり、Pythonが起動エラーになる場合があります。
          本画面では<strong>中身を自動検査する安全ダウンロード</strong>を実装しているほか、<strong>GitHub公式の直接生リンク（HTML混入リスクゼロ）</strong>からもワンクリックで取得できます。
        </p>
      </div>

      {/* Direct GitHub Link for threads_gui_app.py */}
      <div className="p-4 rounded-xl bg-indigo-950/50 border border-indigo-700/80 flex flex-wrap items-center justify-between gap-3">
        <div className="space-y-1">
          <div className="flex items-center gap-2">
            <FileCode2 className="w-5 h-5 text-indigo-400" />
            <span className="font-bold text-white text-xs">【推奨・HTML混入なし】GitHub Raw 直接ダウンロード</span>
          </div>
          <p className="text-indigo-200/80 text-[11px]">
            GitHubサーバーから直接純粋なPythonスクリプト（threads_gui_app.py）を保存します。
          </p>
        </div>

        <a
          href={rawGithubPyUrl}
          target="_blank"
          rel="noreferrer"
          download="threads_gui_app.py"
          className="flex items-center gap-1.5 px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white font-bold rounded-lg shadow-md text-xs cursor-pointer"
        >
          <ExternalLink className="w-3.5 h-3.5" />
          <span>GitHub直接ダウンロード (threads_gui_app.py)</span>
        </a>
      </div>

      {downloadError && (
        <div className="p-3 rounded-lg bg-rose-950 border border-rose-800 text-rose-300 text-xs">
          {downloadError}
        </div>
      )}

      {/* Files Download Grid */}
      <div className="space-y-3">
        <h4 className="font-bold text-zinc-300 text-xs">個別ファイル安全ダウンロード:</h4>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
          {files.map((file) => (
            <div
              key={file.id}
              className="p-4 rounded-xl border flex flex-col justify-between gap-3 transition-all bg-zinc-950/70 border-zinc-800 hover:border-zinc-700"
            >
              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    {file.icon}
                    <span className="font-mono font-bold text-white text-xs">{file.name}</span>
                  </div>
                  <span className={`px-2 py-0.5 rounded text-[10px] font-mono border ${file.color}`}>
                    {file.badge}
                  </span>
                </div>
                <p className="text-zinc-400 text-[11px] leading-relaxed">
                  {file.desc}
                </p>
              </div>

              <div className="flex items-center justify-between pt-2 border-t border-zinc-800/80">
                <span className="font-mono text-[11px] text-zinc-500">{file.size}</span>
                <button
                  type="button"
                  onClick={() => downloadBlobSecurely(file.url, file.name, file.id)}
                  disabled={downloadingId === file.id}
                  className="flex items-center gap-1.5 px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold rounded-lg shadow transition-all cursor-pointer text-xs disabled:opacity-40"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>{downloadingId === file.id ? 'ダウンロード中...' : '安全ダウンロード'}</span>
                </button>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* 3-Step Simple Guide */}
      <div className="p-4 rounded-xl bg-zinc-950 border border-zinc-800 space-y-3">
        <h4 className="font-bold text-zinc-200 text-xs flex items-center gap-1.5">
          <HelpCircle className="w-4 h-4 text-emerald-400" />
          Windows既存フォルダでの正しい配置場所:
        </h4>
        <div className="space-y-2 text-zinc-400 text-[11px] leading-relaxed">
          <p>
            ダウンロードした <code className="text-white bg-zinc-900 px-1 py-0.5 rounded font-mono">threads_gui_app.py</code> は、以下のいずれか（または両方）の場所に上書きコピーしてください：
          </p>
          <ul className="list-disc list-inside space-y-1 pl-2 text-zinc-300 font-mono text-[11px]">
            <li><code>C:\Users\PC_User\Desktop\ThreadsAutoPoster_Windows_Package (36)\dist\ThreadsAutoPoster\threads_gui_app.py</code></li>
            <li><code>C:\Users\PC_User\Desktop\ThreadsAutoPoster_Windows_Package (36)\threads_gui_app.py</code></li>
          </ul>
        </div>
      </div>
    </div>
  );
};
