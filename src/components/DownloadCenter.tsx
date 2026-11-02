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
  CheckCircle2
} from 'lucide-react';

export const DownloadCenter: React.FC = () => {
  const [copiedText, setCopiedText] = useState<string | null>(null);

  const handleCopy = (text: string, id: string) => {
    navigator.clipboard.writeText(text);
    setCopiedText(id);
    setTimeout(() => setCopiedText(null), 2000);
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
    },
  ];

  return (
    <div className="bg-zinc-900 border border-zinc-800 rounded-xl p-5 space-y-6 shadow-sm text-xs">
      {/* Header Banner */}
      <div className="pb-4 border-b border-zinc-800 flex flex-wrap items-center justify-between gap-3">
        <div>
          <div className="flex items-center gap-2">
            <div className="p-2 rounded-xl bg-indigo-950 text-indigo-400 border border-indigo-800">
              <FolderOpen className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-sm font-bold text-white flex items-center gap-2">
                Windows既存フォルダ更新用ファイル (threads_gui_app.py & エクセル)
              </h3>
              <p className="text-zinc-400 text-xs mt-0.5">
                既存のフォルダに以下のファイルを上書きするだけで、EXE側でも新機能がそのまま動作します
              </p>
            </div>
          </div>
        </div>
      </div>

      {/* 3-Step Simple Guide */}
      <div className="p-4 rounded-xl bg-zinc-950 border border-zinc-800 space-y-3">
        <h4 className="font-bold text-zinc-200 text-xs flex items-center gap-1.5">
          <HelpCircle className="w-4 h-4 text-emerald-400" />
          既存のWindowsフォルダでの差し替え手順（わずか3ステップ）:
        </h4>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-3 pt-1">
          <div className="p-3 rounded-lg bg-zinc-900/80 border border-zinc-800 space-y-1.5">
            <div className="font-bold text-emerald-400 flex items-center gap-1.5">
              <span className="w-5 h-5 rounded-full bg-emerald-950 text-emerald-300 flex items-center justify-center font-mono text-xs border border-emerald-800">1</span>
              <span>ダウンロード</span>
            </div>
            <p className="text-zinc-400 leading-relaxed text-[11px]">
              下のリストから <code className="text-white bg-zinc-950 px-1 py-0.5 rounded font-mono">threads_gui_app.py</code> と使いたい <code className="text-white bg-zinc-950 px-1 py-0.5 rounded font-mono">.xlsx</code> をダウンロードします。
            </p>
          </div>

          <div className="p-3 rounded-lg bg-zinc-900/80 border border-zinc-800 space-y-1.5">
            <div className="font-bold text-sky-400 flex items-center gap-1.5">
              <span className="w-5 h-5 rounded-full bg-sky-950 text-sky-300 flex items-center justify-center font-mono text-xs border border-sky-800">2</span>
              <span>既存フォルダへ上書き</span>
            </div>
            <p className="text-zinc-400 leading-relaxed text-[11px]">
              Windows上で現在使っているフォルダへダウンロードしたファイルをドラッグ＆ドロップし、「ファイルを置き換える」を選択します。
            </p>
          </div>

          <div className="p-3 rounded-lg bg-zinc-900/80 border border-zinc-800 space-y-1.5">
            <div className="font-bold text-purple-400 flex items-center gap-1.5">
              <span className="w-5 h-5 rounded-full bg-purple-950 text-purple-300 flex items-center justify-center font-mono text-xs border border-purple-800">3</span>
              <span>EXEをダブルクリック起動</span>
            </div>
            <p className="text-zinc-400 leading-relaxed text-[11px]">
              いつも通り <code className="text-white bg-zinc-950 px-1 py-0.5 rounded font-mono">ThreadsAutoPoster.exe</code> を起動。ツールバーに新ボタンが出現します！
            </p>
          </div>
        </div>
      </div>

      {/* Files Download Grid */}
      <div className="space-y-3">
        <h4 className="font-bold text-zinc-300 text-xs">ダウンロード対象ファイル一覧:</h4>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
          {files.map((file) => (
            <div
              key={file.id}
              className={`p-4 rounded-xl border flex flex-col justify-between gap-3 transition-all bg-zinc-950/70 border-zinc-800 hover:border-zinc-700`}
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
                <a
                  href={file.url}
                  download={file.name}
                  className="flex items-center gap-1.5 px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold rounded-lg shadow transition-all cursor-pointer text-xs"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>ダウンロード</span>
                </a>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Image Assets Section */}
      <div className="p-4 rounded-xl bg-zinc-950 border border-zinc-800 flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-2">
          <ImageIcon className="w-5 h-5 text-violet-400" />
          <div>
            <div className="font-bold text-white text-xs">添付用画像ファイル (image.png & image2.png)</div>
            <div className="text-[11px] text-zinc-400">
              ポスト文章で指定される添付画像のオリジナルファイルです
            </div>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <a
            href="/images/image.png"
            download="image.png"
            className="flex items-center gap-1 px-3 py-1.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 border border-zinc-700 rounded-lg text-xs"
          >
            <Download className="w-3 h-3" />
            <span>image.png</span>
          </a>
          <a
            href="/images/image2.png"
            download="image2.png"
            className="flex items-center gap-1 px-3 py-1.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 border border-zinc-700 rounded-lg text-xs"
          >
            <Download className="w-3 h-3" />
            <span>image2.png</span>
          </a>
        </div>
      </div>
    </div>
  );
};
