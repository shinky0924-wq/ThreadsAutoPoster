import React, { useState, useEffect, useRef } from 'react';
import { AccountInfo, AccountDesignProfile, DailyGreetingSlot, PostItem } from '../types';
import {
  ACCOUNT_DESIGN_PRESETS,
  AccountDesignInput,
  AccountTone,
  BUILTIN_SAMPLE_BACKGROUNDS,
  BackgroundPhotoItem,
  ColorPaletteId,
  ConcreteImageTemplateStyle,
  GeneratedPostBundle,
  ImageAspectRatio,
  ImageTemplateStyle,
  generatePostSpecs,
  loadImageElement,
  renderCanvasToDataUrlAndBytes,
  renderPostImageOnCanvas,
} from '../utils/postAndImageGenerator';
import { buildAccountWorkbookBytes } from '../utils/excelUtils';
import { createZipBlob, ZipInputFile } from '../utils/zipUtils';
import {
  Wand2,
  ImagePlus,
  Download,
  CheckCircle2,
  Trash2,
  FileSpreadsheet,
  Package,
  RefreshCw,
  Layers,
  Eye,
  Save,
} from 'lucide-react';

interface GeneratorStudioProps {
  currentAccount: AccountInfo;
  dailySlots: DailyGreetingSlot[];
  onApplyToAccountPosts: (
    newItems: Omit<PostItem, 'id'>[],
    replaceExisting: boolean,
    generatedImages: Record<string, string>
  ) => void;
  onSaveDesignToAccount: (accountId: string, designProfile: AccountDesignProfile) => void;
  onSwitchToTab: (tab: 'posts' | 'preview') => void;
  onLog: (level: 'INFO' | 'SUCCESS' | 'WARN' | 'ERROR', message: string) => void;
}

function buildDesignFromAccount(acc: AccountInfo): Omit<AccountDesignInput, 'count' | 'enableSpintax' | 'generateImages'> {
  const base = ACCOUNT_DESIGN_PRESETS[0].design;
  const d = acc.design;
  return {
    ...base,
    genre: d?.genre ?? base.genre,
    target: d?.target ?? base.target,
    tone: d?.tone ?? base.tone,
    painPoint: d?.painPoint ?? base.painPoint,
    benefitPoints: d?.benefitPoints ?? base.benefitPoints,
    cta: d?.cta ?? base.cta,
    url: d?.url ?? base.url,
    topic: d?.topic ?? acc.topics?.[0] ?? base.topic,
    imageStyle: d?.imageStyle ?? base.imageStyle,
  };
}

export const GeneratorStudio: React.FC<GeneratorStudioProps> = ({
  currentAccount,
  dailySlots,
  onApplyToAccountPosts,
  onSaveDesignToAccount,
  onSwitchToTab,
  onLog,
}) => {
  const [design, setDesign] = useState<AccountDesignInput>(() => ({
    ...buildDesignFromAccount(currentAccount),
    count: 10,
    enableSpintax: true,
    generateImages: true,
  }));

  const lastLoadedAccRef = useRef<string>('');

  // Auto-load design fields whenever switching accounts or importing updated account design from Excel
  useEffect(() => {
    const sig = `${currentAccount.id}:${JSON.stringify(currentAccount.design || {})}`;
    if (lastLoadedAccRef.current !== sig) {
      lastLoadedAccRef.current = sig;
      setDesign((prev) => ({
        ...prev,
        ...buildDesignFromAccount(currentAccount),
      }));
      setPreviewHeadlineOverride('');
    }
  }, [currentAccount.id, currentAccount.design, currentAccount.topics]);

  const [backgrounds, setBackgrounds] = useState<BackgroundPhotoItem[]>(BUILTIN_SAMPLE_BACKGROUNDS);
  const [selectedBgId, setSelectedBgId] = useState<string>(BUILTIN_SAMPLE_BACKGROUNDS[0].id);
  const [previewHeadlineOverride, setPreviewHeadlineOverride] = useState<string>('');
  const [generatedBundles, setGeneratedBundles] = useState<GeneratedPostBundle[]>([]);
  const [isGenerating, setIsGenerating] = useState(false);
  const [progressCount, setProgressCount] = useState(0);
  const [lastElapsedMs, setLastElapsedMs] = useState<number | null>(null);
  const [replaceOnImport, setReplaceOnImport] = useState(true);
  const [inspectedBundle, setInspectedBundle] = useState<GeneratedPostBundle | null>(null);

  const liveCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const bgUploadInputRef = useRef<HTMLInputElement | null>(null);

  const activeBackground =
    backgrounds.find((b) => b.id === selectedBgId) || backgrounds[0] || null;

  // Render live single-card preview whenever design/background changes
  useEffect(() => {
    let cancelled = false;
    const renderLive = async () => {
      const canvas = liveCanvasRef.current;
      if (!canvas) return;

      const hasPhotos = backgrounds.length > 0;
      const sampleSpecs = generatePostSpecs({ ...design, count: 1 }, currentAccount.id, hasPhotos);
      const spec = sampleSpecs[0];
      if (!spec) return;

      let loadedImg: HTMLImageElement | null = null;
      if (activeBackground?.src) {
        try {
          loadedImg = await loadImageElement(activeBackground.src);
        } catch {
          loadedImg = null;
        }
      }
      if (cancelled) return;

      const previewStyle: ConcreteImageTemplateStyle = spec.imageStyleUsed;

      renderPostImageOnCanvas(canvas, {
        aspectRatio: design.aspectRatio,
        style: previewStyle,
        palette: design.palette,
        scrimOpacity: design.scrimOpacity,
        topic: design.topic || 'GUIDE',
        headline: previewHeadlineOverride.trim() || spec.headline,
        subhead: spec.subhead,
        painPoint: spec.painPoint,
        bullets: spec.bullets,
        accountId: currentAccount.id,
        showWatermark: design.showAccountWatermark,
        backgroundImg: loadedImg,
        postIndex: 1,
      });
    };

    renderLive();
    return () => {
      cancelled = true;
    };
  }, [design, activeBackground, currentAccount.id, previewHeadlineOverride]);

  const handleApplyPreset = (presetId: string) => {
    const found = ACCOUNT_DESIGN_PRESETS.find((p) => p.id === presetId);
    if (!found) return;
    setDesign((prev) => ({
      ...prev,
      ...found.design,
    }));
    setPreviewHeadlineOverride('');
    onLog('INFO', `プリセット「${found.label}」をアカウント設計に読み込みました。`);
  };

  const handleUploadBackgrounds = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const added: BackgroundPhotoItem[] = [];
    Array.from(files).forEach((file, idx) => {
      if (!file.type.startsWith('image/')) return;
      const objUrl = URL.createObjectURL(file);
      added.push({
        id: `user-${Date.now()}-${idx}`,
        name: file.name,
        src: objUrl,
        isBuiltin: false,
      });
    });

    if (added.length > 0) {
      setBackgrounds((prev) => [...prev, ...added]);
      setSelectedBgId(added[0].id);
      onLog('SUCCESS', `手持ち背景写真を ${added.length} 枚追加しました。`);
    }
    e.target.value = '';
  };

  const handleRemoveBackground = (id: string) => {
    setBackgrounds((prev) => {
      const target = prev.find((b) => b.id === id);
      if (target && !target.isBuiltin && target.src.startsWith('blob:')) {
        URL.revokeObjectURL(target.src);
      }
      const next = prev.filter((b) => b.id !== id);
      if (selectedBgId === id && next.length > 0) {
        setSelectedBgId(next[0].id);
      }
      return next;
    });
  };

  const handleDownloadSingleLivePng = () => {
    const canvas = liveCanvasRef.current;
    if (!canvas) return;
    const dataUrl = canvas.toDataURL('image/png');
    const a = document.createElement('a');
    a.href = dataUrl;
    a.download = `${currentAccount.id}_sample_preview.png`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    onLog('SUCCESS', `テスト画像PNG (${currentAccount.id}_sample_preview.png) を保存しました。`);
  };

  const handleBatchGenerate = async () => {
    if (isGenerating) return;
    setIsGenerating(true);
    setProgressCount(0);
    const t0 = performance.now();

    try {
      const specs = generatePostSpecs(design, currentAccount.id, backgrounds.length > 0);
      const offscreen = document.createElement('canvas');

      // Preload all backgrounds once for rapid rendering
      const loadedBgImgs: (HTMLImageElement | null)[] = [];
      if (design.generateImages && backgrounds.length > 0) {
        for (const bg of backgrounds) {
          try {
            loadedBgImgs.push(await loadImageElement(bg.src));
          } catch {
            loadedBgImgs.push(null);
          }
        }
      }

      const results: GeneratedPostBundle[] = [];
      for (let i = 0; i < specs.length; i++) {
        const spec = specs[i];
        let imageDataUrl = '';
        let imageBytes: Uint8Array | undefined;

        if (design.generateImages) {
          const bgImg =
            loadedBgImgs.length > 0 ? loadedBgImgs[i % loadedBgImgs.length] : null;

          renderPostImageOnCanvas(offscreen, {
            aspectRatio: design.aspectRatio,
            style: spec.imageStyleUsed,
            palette: design.palette,
            scrimOpacity: design.scrimOpacity,
            topic: spec.topic,
            headline: i === 0 && previewHeadlineOverride.trim() ? previewHeadlineOverride.trim() : spec.headline,
            subhead: spec.subhead,
            painPoint: spec.painPoint,
            bullets: spec.bullets,
            accountId: currentAccount.id,
            showWatermark: design.showAccountWatermark,
            backgroundImg: bgImg,
            postIndex: spec.index,
          });

          const rendered = await renderCanvasToDataUrlAndBytes(offscreen);
          imageDataUrl = rendered.dataUrl;
          imageBytes = rendered.bytes;
        }

        results.push({
          ...spec,
          imageDataUrl,
          imageBytes,
        });

        if ((i + 1) % 5 === 0 || i === specs.length - 1) {
          setProgressCount(i + 1);
          // Yield to UI thread so progress counter updates smoothly
          await new Promise((r) => setTimeout(r, 0));
        }
      }

      const elapsed = Math.round(performance.now() - t0);
      setLastElapsedMs(elapsed);
      setGeneratedBundles(results);
      onLog(
        'SUCCESS',
        `⚡ [@${currentAccount.id}] ${results.length} 件のポスト文章${design.generateImages ? '＋Canvas画像PNG' : ''} を ${(elapsed / 1000).toFixed(2)}秒 (完全0円) で生成しました！`
      );
    } catch (err: any) {
      onLog('ERROR', `一括生成エラー: ${err?.message || err}`);
    } finally {
      setIsGenerating(false);
    }
  };

  const handleRegisterToAccountPosts = () => {
    if (generatedBundles.length === 0) return;
    const newItems: Omit<PostItem, 'id'>[] = generatedBundles.map((b) => ({
      text: b.text,
      topic: b.topic,
      imagePath: b.imageFilename,
      status: '未投稿',
    }));

    const imgMap: Record<string, string> = {};
    generatedBundles.forEach((b) => {
      if (b.imageFilename && b.imageDataUrl) {
        imgMap[b.imageFilename] = b.imageDataUrl;
      }
    });

    onApplyToAccountPosts(newItems, replaceOnImport, imgMap);
  };

  const handleDownloadPackZip = () => {
    if (generatedBundles.length === 0) return;

    const postsForExcel: PostItem[] = generatedBundles.map((b) => ({
      id: b.index,
      text: b.text,
      topic: b.topic,
      imagePath: b.imageFilename,
      status: '未投稿',
    }));

    const xlsxBytes = buildAccountWorkbookBytes(currentAccount, postsForExcel, dailySlots);
    const zipFiles: ZipInputFile[] = [
      {
        name: `${currentAccount.id}.xlsx`,
        content: xlsxBytes,
      },
    ];

    generatedBundles.forEach((b) => {
      if (b.imageFilename && b.imageBytes) {
        zipFiles.push({
          name: `images/${currentAccount.id}/${b.imageFilename}`,
          content: b.imageBytes,
        });
      }
    });

    const blob = createZipBlob(zipFiles);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${currentAccount.id}_generated_pack.zip`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    onLog(
      'SUCCESS',
      `📦 ${currentAccount.id}_generated_pack.zip (${currentAccount.id}.xlsx + 画像${zipFiles.length - 1}枚) をダウンロードしました！`
    );
  };

  const handleDownloadSingleBundlePng = (b: GeneratedPostBundle) => {
    if (!b.imageDataUrl) return;
    const a = document.createElement('a');
    a.href = b.imageDataUrl;
    a.download = b.imageFilename || `post_${b.index}.png`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  return (
    <div className="space-y-4">
      {/* Header Bar */}
      <div className="bg-zinc-900 border border-zinc-800 rounded-xl p-5 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
        <div className="space-y-1">
          <div className="flex items-center gap-2 text-xs text-zinc-400">
            <span>対象アカウント: <strong className="text-white font-mono">@{currentAccount.id}</strong></span>
            <span aria-hidden="true">·</span>
            <span>HTML5 Canvas + スピンタックス生成</span>
            <span aria-hidden="true">·</span>
            <span className="text-emerald-400 font-medium">APIキー不要 / 完全無料</span>
          </div>
          <h2 className="text-base font-bold text-white">
            アカウント設計 × ワンクリック投稿・画像テスト生成スタジオ
          </h2>
        </div>

        {/* Presets */}
        <div className="flex items-center gap-2 flex-wrap">
          <span className="text-xs text-zinc-400">テスト用プリセット:</span>
          {ACCOUNT_DESIGN_PRESETS.map((preset) => (
            <button
              key={preset.id}
              type="button"
              onClick={() => handleApplyPreset(preset.id)}
              className="px-3 py-1.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 border border-zinc-700 rounded-lg text-xs font-medium transition-colors cursor-pointer whitespace-nowrap shrink-0"
            >
              {preset.label}
            </button>
          ))}
        </div>
      </div>

      {/* Main 2-Column Workspace */}
      <div className="grid grid-cols-1 xl:grid-cols-12 gap-4 items-start">
        {/* Left: Account Design & Canvas Controls (5 cols) */}
        <div className="xl:col-span-5 bg-zinc-900 border border-zinc-800 rounded-xl p-5 space-y-5 text-xs">
          {/* Section 1: Account Copy Design */}
          <div className="space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-2 pb-2 border-b border-zinc-800">
              <h3 className="text-sm font-bold text-white">
                1. アカウント設計（@{currentAccount.id} から読込済）
              </h3>
              <button
                type="button"
                onClick={() =>
                  onSaveDesignToAccount(currentAccount.id, {
                    genre: design.genre,
                    target: design.target,
                    tone: design.tone,
                    painPoint: design.painPoint,
                    benefitPoints: design.benefitPoints,
                    cta: design.cta,
                    url: design.url,
                    topic: design.topic,
                    imageStyle: design.imageStyle,
                  })
                }
                className="flex items-center gap-1 px-2.5 py-1 bg-sky-600/20 hover:bg-sky-600/30 text-sky-300 border border-sky-500/40 rounded-md text-[11px] font-semibold cursor-pointer transition-colors whitespace-nowrap shrink-0"
                title="入力中のアカウント設計をこのアカウントに紐づけて保存（Excel出力にも反映）"
              >
                <Save className="w-3 h-3" />
                <span>@{currentAccount.id} に保存</span>
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label className="block text-zinc-400 mb-1">ジャンル・発信テーマ</label>
                <input
                  type="text"
                  value={design.genre}
                  onChange={(e) => setDesign({ ...design, genre: e.target.value })}
                  className="w-full bg-zinc-950 border border-zinc-700 text-white px-3 py-2 rounded-lg focus:outline-none focus:border-emerald-500"
                />
              </div>
              <div>
                <label className="block text-zinc-400 mb-1">トピック（Excel設定用）</label>
                <input
                  type="text"
                  value={design.topic}
                  onChange={(e) => setDesign({ ...design, topic: e.target.value })}
                  className="w-full bg-zinc-950 border border-zinc-700 text-white px-3 py-2 rounded-lg focus:outline-none focus:border-emerald-500"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label className="block text-zinc-400 mb-1">ターゲット層</label>
                <input
                  type="text"
                  value={design.target}
                  onChange={(e) => setDesign({ ...design, target: e.target.value })}
                  className="w-full bg-zinc-950 border border-zinc-700 text-white px-3 py-2 rounded-lg focus:outline-none focus:border-emerald-500"
                />
              </div>
              <div>
                <label className="block text-zinc-400 mb-1">アカウント口調</label>
                <select
                  value={design.tone}
                  onChange={(e) => setDesign({ ...design, tone: e.target.value as AccountTone })}
                  className="w-full bg-zinc-950 border border-zinc-700 text-white px-3 py-2 rounded-lg focus:outline-none focus:border-emerald-500"
                >
                  <option value="polite">丁寧・誠実（信頼感重視）</option>
                  <option value="friendly">フレンドリー・親近感</option>
                  <option value="expert">プロ解説・ノウハウ特化</option>
                  <option value="empathetic">エモい・悩み共感スタイル</option>
                </select>
              </div>
            </div>

            <div>
              <label className="block text-zinc-400 mb-1">読者の悩み・不安ポイント</label>
              <input
                type="text"
                value={design.painPoint}
                onChange={(e) => setDesign({ ...design, painPoint: e.target.value })}
                className="w-full bg-zinc-950 border border-zinc-700 text-white px-3 py-2 rounded-lg focus:outline-none focus:border-emerald-500"
              />
            </div>

            <div>
              <label className="block text-zinc-400 mb-1">
                強み・ポイント（1行1項目 · 図解/チェックリスト画像の項目にも自動反映）
              </label>
              <textarea
                rows={4}
                value={design.benefitPoints}
                onChange={(e) => setDesign({ ...design, benefitPoints: e.target.value })}
                className="w-full bg-zinc-950 border border-zinc-700 text-white p-2.5 rounded-lg focus:outline-none focus:border-emerald-500 leading-relaxed"
              />
            </div>

            <div>
              <label className="block text-zinc-400 mb-1">誘導文（CTA・結び）</label>
              <input
                type="text"
                value={design.cta}
                onChange={(e) => setDesign({ ...design, cta: e.target.value })}
                className="w-full bg-zinc-950 border border-zinc-700 text-white px-3 py-2 rounded-lg focus:outline-none focus:border-emerald-500"
              />
            </div>

            <div>
              <label className="block text-zinc-400 mb-1">リンクURL（任意）</label>
              <input
                type="text"
                value={design.url}
                onChange={(e) => setDesign({ ...design, url: e.target.value })}
                placeholder="https://..."
                className="w-full bg-zinc-950 border border-zinc-700 text-white px-3 py-2 rounded-lg font-mono focus:outline-none focus:border-emerald-500"
              />
            </div>
          </div>

          {/* Section 2: Canvas Image Studio Settings */}
          <div className="space-y-3 pt-2 border-t border-zinc-800">
            <div className="flex items-center justify-between">
              <h3 className="text-sm font-bold text-white">
                2. Canvas 画像生成デザイン設定
              </h3>
              <label className="flex items-center gap-1.5 cursor-pointer text-zinc-300">
                <input
                  type="checkbox"
                  checked={design.generateImages}
                  onChange={(e) => setDesign({ ...design, generateImages: e.target.checked })}
                  className="accent-emerald-500"
                />
                <span>画像も同時に生成</span>
              </label>
            </div>

            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-zinc-400 mb-1">画像スタイル</label>
                <select
                  value={design.imageStyle}
                  onChange={(e) =>
                    setDesign({ ...design, imageStyle: e.target.value as ImageTemplateStyle })
                  }
                  className="w-full bg-zinc-950 border border-zinc-700 text-white px-3 py-2 rounded-lg focus:outline-none"
                >
                  <option value="auto">自動おまかせ（設計に合わせて最適化）</option>
                  <option value="magazine_clean">雑誌エディトリアル風</option>
                  <option value="chat_bubble">吹き出し会話SNS風</option>
                  <option value="neon_night">ネオン夜風</option>
                  <option value="step_roadmap">ステップ図解</option>
                  <option value="qa_card">Q&A神回答カード</option>
                  <option value="photo_overlay">背景写真＋自動文字入れ</option>
                  <option value="diagram_card">SNS図解カード（01/02/03）</option>
                  <option value="checklist">保存用チェックリスト</option>
                  <option value="all_mix">全テンプレートミックス</option>
                  <option value="mix">3パターンミックス</option>
                </select>
              </div>
              <div>
                <label className="block text-zinc-400 mb-1">画像サイズ</label>
                <select
                  value={design.aspectRatio}
                  onChange={(e) =>
                    setDesign({ ...design, aspectRatio: e.target.value as ImageAspectRatio })
                  }
                  className="w-full bg-zinc-950 border border-zinc-700 text-white px-3 py-2 rounded-lg focus:outline-none"
                >
                  <option value="1:1">1:1 正方形 (1080×1080)</option>
                  <option value="4:5">4:5 縦長推奨 (1080×1350)</option>
                </select>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-zinc-400 mb-1">配色テーマ</label>
                <select
                  value={design.palette}
                  onChange={(e) =>
                    setDesign({ ...design, palette: e.target.value as ColorPaletteId })
                  }
                  className="w-full bg-zinc-950 border border-zinc-700 text-white px-3 py-2 rounded-lg focus:outline-none"
                >
                  <option value="amber_night">提灯アンバー × 黒</option>
                  <option value="emerald_slate">エメラルド × 濃紺</option>
                  <option value="crimson_gold">深紅アクセント × 黒</option>
                  <option value="ivory_ink">生成り白 × 墨文字</option>
                  <option value="magazine_ivory_navy">雑誌アイボリー × 紺</option>
                  <option value="neon_cyan_night">ネオンシアン × 漆黒</option>
                </select>
              </div>
              <div>
                <label className="block text-zinc-400 mb-1 tabular-nums">
                  文字の読みやすさ暗幕 ({Math.round(design.scrimOpacity * 100)}%)
                </label>
                <input
                  type="range"
                  min={0.3}
                  max={0.85}
                  step={0.05}
                  value={design.scrimOpacity}
                  onChange={(e) =>
                    setDesign({ ...design, scrimOpacity: Number(e.target.value) })
                  }
                  className="w-full accent-emerald-500 mt-2"
                />
              </div>
            </div>

            <div>
              <label className="block text-zinc-400 mb-1">
                テスト用キャッチコピー上書き（空欄なら自動生成）
              </label>
              <input
                type="text"
                value={previewHeadlineOverride}
                onChange={(e) => setPreviewHeadlineOverride(e.target.value)}
                placeholder="例: 飛田新地・初めてでも失敗しない選び方"
                className="w-full bg-zinc-950 border border-zinc-700 text-white px-3 py-2 rounded-lg focus:outline-none focus:border-emerald-500"
              />
            </div>

            {/* Background Photo Picker & Upload */}
            <div className="space-y-2 pt-1">
              <div className="flex items-center justify-between">
                <span className="text-zinc-300 font-semibold">
                  背景写真（クリックで切替 · 一括生成時は順番に使用）
                </span>
                <button
                  type="button"
                  onClick={() => bgUploadInputRef.current?.click()}
                  className="flex items-center gap-1 px-2.5 py-1 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 border border-zinc-700 rounded-md cursor-pointer transition-colors"
                >
                  <ImagePlus className="w-3.5 h-3.5 text-sky-400" />
                  <span>手持ち写真を追加</span>
                </button>
                <input
                  ref={bgUploadInputRef}
                  type="file"
                  accept="image/*"
                  multiple
                  className="hidden"
                  onChange={handleUploadBackgrounds}
                />
              </div>

              <div className="grid grid-cols-3 sm:grid-cols-4 gap-2">
                {backgrounds.map((bg) => {
                  const isSelected = bg.id === activeBackground?.id;
                  return (
                    <div
                      key={bg.id}
                      onClick={() => setSelectedBgId(bg.id)}
                      className={`relative group rounded-lg overflow-hidden border cursor-pointer aspect-square bg-zinc-950 transition-all ${
                        isSelected
                          ? 'border-emerald-500 ring-2 ring-emerald-500/40'
                          : 'border-zinc-800 hover:border-zinc-600'
                      }`}
                      title={bg.name}
                    >
                      <img
                        src={bg.src}
                        alt={bg.name}
                        className="w-full h-full object-cover"
                      />
                      <div className="absolute inset-x-0 bottom-0 bg-black/75 px-1.5 py-0.5 text-[10px] text-zinc-200 truncate">
                        {bg.name}
                      </div>
                      {!bg.isBuiltin && (
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            handleRemoveBackground(bg.id);
                          }}
                          className="absolute top-1 right-1 p-1 rounded bg-black/80 text-zinc-300 hover:text-rose-400 opacity-0 group-hover:opacity-100 transition-opacity cursor-pointer"
                          title="削除"
                        >
                          <Trash2 className="w-3 h-3" />
                        </button>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          </div>

          {/* Section 3: Batch Options & Main CTA */}
          <div className="space-y-3 pt-2 border-t border-zinc-800">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex items-center gap-2">
                <label className="text-zinc-400">生成件数:</label>
                <select
                  value={design.count}
                  onChange={(e) => setDesign({ ...design, count: Number(e.target.value) })}
                  className="bg-zinc-950 border border-zinc-700 text-white px-2.5 py-1.5 rounded-lg font-mono tabular-nums focus:outline-none"
                >
                  <option value={5}>5件 (テスト)</option>
                  <option value={10}>10件</option>
                  <option value={30}>30件</option>
                  <option value={50}>50件</option>
                  <option value={100}>100件 (大量)</option>
                </select>
              </div>

              <label className="flex items-center gap-1.5 cursor-pointer text-zinc-300">
                <input
                  type="checkbox"
                  checked={design.enableSpintax}
                  onChange={(e) => setDesign({ ...design, enableSpintax: e.target.checked })}
                  className="accent-emerald-500"
                />
                <span>スピンタックス {'{A|B|C}'} 挿入</span>
              </label>

              <label className="flex items-center gap-1.5 cursor-pointer text-zinc-300">
                <input
                  type="checkbox"
                  checked={design.showAccountWatermark}
                  onChange={(e) =>
                    setDesign({ ...design, showAccountWatermark: e.target.checked })
                  }
                  className="accent-emerald-500"
                />
                <span>@ID透かし</span>
              </label>
            </div>

            <button
              type="button"
              onClick={handleBatchGenerate}
              disabled={isGenerating}
              className="w-full flex items-center justify-center gap-2 px-4 py-3 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 text-white font-bold text-sm rounded-xl shadow-md transition-all cursor-pointer"
            >
              {isGenerating ? (
                <>
                  <RefreshCw className="w-4 h-4 animate-spin" />
                  <span className="tabular-nums">
                    生成中... ({progressCount} / {design.count} 件)
                  </span>
                </>
              ) : (
                <>
                  <Wand2 className="w-4 h-4" />
                  <span>
                    ワンクリックで {design.count} 件を一括生成（完全無料）
                  </span>
                </>
              )}
            </button>
          </div>
        </div>

        {/* Right: Live Canvas Preview & Generated Output List (7 cols) */}
        <div className="xl:col-span-7 space-y-4">
          {/* Live Single Card Canvas Preview */}
          <div className="bg-zinc-900 border border-zinc-800 rounded-xl p-5 space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-2 pb-3 border-b border-zinc-800">
              <div>
                <h3 className="text-sm font-bold text-white">
                  リアルタイム画像プレビュー（文字入れ・図解仕上がり確認）
                </h3>
                <p className="text-xs text-zinc-400">
                  左の設定や背景写真を変えると即座に1080px解像度でレンダリングされます
                </p>
              </div>
              <button
                type="button"
                onClick={handleDownloadSingleLivePng}
                className="flex items-center gap-1.5 px-3 py-1.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 border border-zinc-700 rounded-lg text-xs font-medium cursor-pointer transition-colors whitespace-nowrap shrink-0"
              >
                <Download className="w-3.5 h-3.5" />
                <span>この1枚をPNG保存</span>
              </button>
            </div>

            <div className="flex justify-center bg-zinc-950 border border-zinc-800 rounded-xl p-4">
              <canvas
                ref={liveCanvasRef}
                className="max-h-[420px] w-auto rounded-lg shadow-lg border border-zinc-800"
              />
            </div>
          </div>

          {/* Batch Results */}
          <div className="bg-zinc-900 border border-zinc-800 rounded-xl p-5 space-y-4 text-xs">
            <div className="flex flex-wrap items-center justify-between gap-3 pb-3 border-b border-zinc-800">
              <div>
                <h3 className="text-sm font-bold text-white flex items-center gap-2">
                  <span>一括生成結果</span>
                  {generatedBundles.length > 0 && (
                    <span className="font-mono tabular-nums text-emerald-400">
                      {generatedBundles.length} 件
                      {lastElapsedMs !== null ? ` · ${(lastElapsedMs / 1000).toFixed(2)}秒` : ''}
                    </span>
                  )}
                </h3>
                <p className="text-zinc-400 text-xs mt-0.5">
                  生成したセットはそのままアプリ内の投稿リストへ登録、またはWindows版用のZIPで一括出力できます
                </p>
              </div>

              {generatedBundles.length > 0 && (
                <div className="flex flex-wrap items-center gap-2">
                  <label className="flex items-center gap-1.5 text-zinc-300 cursor-pointer mr-1">
                    <input
                      type="checkbox"
                      checked={replaceOnImport}
                      onChange={(e) => setReplaceOnImport(e.target.checked)}
                      className="accent-emerald-500"
                    />
                    <span>既存リストを入替</span>
                  </label>

                  <button
                    type="button"
                    onClick={handleRegisterToAccountPosts}
                    className="flex items-center gap-1.5 px-3 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold rounded-lg cursor-pointer transition-colors whitespace-nowrap shrink-0"
                  >
                    <Layers className="w-3.5 h-3.5" />
                    <span>投稿リストへ反映</span>
                  </button>

                  <button
                    type="button"
                    onClick={handleDownloadPackZip}
                    className="flex items-center gap-1.5 px-3 py-2 bg-indigo-600 hover:bg-indigo-500 text-white font-semibold rounded-lg cursor-pointer transition-colors whitespace-nowrap shrink-0"
                    title={`${currentAccount.id}.xlsx と images/${currentAccount.id}/ をまとめたZIPをダウンロード`}
                  >
                    <Package className="w-3.5 h-3.5" />
                    <span>Excel＋画像ZIP出力</span>
                  </button>
                </div>
              )}
            </div>

            {generatedBundles.length === 0 ? (
              <div className="py-10 text-center space-y-2 text-zinc-400">
                <p className="font-medium text-zinc-300">
                  まだ一括生成データはありません
                </p>
                <p className="text-zinc-500">
                  左下の「ワンクリックで一括生成」ボタンを押すと、文章と画像PNGのペアが即座に生成されます。
                </p>
              </div>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 max-h-[540px] overflow-y-auto pr-1">
                {generatedBundles.map((bundle) => (
                  <div
                    key={bundle.index}
                    className="bg-zinc-950 border border-zinc-800 rounded-xl p-3 flex gap-3 hover:border-zinc-700 transition-colors"
                  >
                    {bundle.imageDataUrl && (
                      <div className="w-24 shrink-0 space-y-1.5">
                        <img
                          src={bundle.imageDataUrl}
                          alt={bundle.imageFilename}
                          onClick={() => setInspectedBundle(bundle)}
                          className="w-24 h-24 object-cover rounded-lg border border-zinc-800 cursor-pointer hover:opacity-90"
                        />
                        <div className="flex items-center justify-between text-[10px] font-mono text-zinc-400">
                          <span className="truncate">{bundle.imageFilename}</span>
                          <button
                            type="button"
                            onClick={() => handleDownloadSingleBundlePng(bundle)}
                            className="text-sky-400 hover:text-sky-300 cursor-pointer"
                            title="このPNGを保存"
                          >
                            <Download className="w-3 h-3" />
                          </button>
                        </div>
                      </div>
                    )}

                    <div className="flex-1 min-w-0 flex flex-col justify-between space-y-2">
                      <div>
                        <div className="flex items-center justify-between text-[11px] text-zinc-400 font-mono tabular-nums mb-1">
                          <span>#{String(bundle.index).padStart(3, '0')} · {bundle.topic}</span>
                          <span>{bundle.text.length}文字</span>
                        </div>
                        <p className="text-zinc-200 whitespace-pre-wrap line-clamp-4 leading-relaxed">
                          {bundle.text}
                        </p>
                      </div>

                      <div className="flex items-center justify-end gap-2 pt-1 border-t border-zinc-900">
                        <button
                          type="button"
                          onClick={() => setInspectedBundle(bundle)}
                          className="inline-flex items-center gap-1 text-[11px] text-zinc-400 hover:text-white cursor-pointer"
                        >
                          <Eye className="w-3 h-3" />
                          <span>拡大確認</span>
                        </button>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {generatedBundles.length > 0 && (
              <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-zinc-800 text-zinc-400">
                <div className="flex items-center gap-1.5">
                  <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                  <span>
                    ZIPには <code className="text-zinc-200 font-mono">{currentAccount.id}.xlsx</code> と{' '}
                    <code className="text-zinc-200 font-mono">images/{currentAccount.id}/</code> が同梱されます
                  </span>
                </div>
                <button
                  type="button"
                  onClick={() => onSwitchToTab('preview')}
                  className="inline-flex items-center gap-1 text-sky-400 hover:text-sky-300 cursor-pointer font-medium"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5" />
                  <span>実機プレビューで確認する</span>
                </button>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Modal to inspect single generated bundle */}
      {inspectedBundle && (
        <div
          onClick={() => setInspectedBundle(null)}
          className="fixed inset-0 bg-black/80 backdrop-blur-xs flex items-center justify-center p-4 z-50"
        >
          <div
            onClick={(e) => e.stopPropagation()}
            className="bg-zinc-900 border border-zinc-700 rounded-2xl max-w-2xl w-full p-5 space-y-4 shadow-2xl text-xs"
          >
            <div className="flex items-center justify-between pb-2 border-b border-zinc-800">
              <span className="text-sm font-bold text-white font-mono">
                生成データ #{String(inspectedBundle.index).padStart(3, '0')} ({inspectedBundle.imageFilename || 'テキストのみ'})
              </span>
              <button
                type="button"
                onClick={() => setInspectedBundle(null)}
                className="px-2.5 py-1 text-zinc-400 hover:text-white cursor-pointer"
              >
                閉じる
              </button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 items-start">
              {inspectedBundle.imageDataUrl && (
                <div className="space-y-2">
                  <img
                    src={inspectedBundle.imageDataUrl}
                    alt={inspectedBundle.imageFilename}
                    className="w-full rounded-xl border border-zinc-800"
                  />
                  <button
                    type="button"
                    onClick={() => handleDownloadSingleBundlePng(inspectedBundle)}
                    className="w-full flex items-center justify-center gap-1.5 py-2 bg-zinc-800 hover:bg-zinc-700 text-white rounded-lg cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>{inspectedBundle.imageFilename} をダウンロード</span>
                  </button>
                </div>
              )}

              <div className="bg-zinc-950 border border-zinc-800 rounded-xl p-3.5 space-y-2">
                <div className="text-zinc-400 font-medium">Threads 投稿文（スピンタックス含む）</div>
                <p className="text-zinc-100 whitespace-pre-wrap leading-relaxed">
                  {inspectedBundle.text}
                </p>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
