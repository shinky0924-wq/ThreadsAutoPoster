import tobitaLanternsBg from '../assets/images/tobita_street_lanterns_1794644991845.jpg';
import osakaMachiyaBg from '../assets/images/osaka_machiya_dusk_1794644992304.jpg';

export type AccountTone = 'polite' | 'friendly' | 'expert' | 'empathetic';
export type ImageTemplateStyle = 'photo_overlay' | 'diagram_card' | 'checklist' | 'mix';
export type ImageAspectRatio = '1:1' | '4:5';
export type ColorPaletteId = 'amber_night' | 'emerald_slate' | 'crimson_gold' | 'ivory_ink';

export interface AccountDesignInput {
  genre: string;
  target: string;
  tone: AccountTone;
  painPoint: string;
  benefitPoints: string;
  cta: string;
  url: string;
  topic: string;
  count: number;
  enableSpintax: boolean;
  generateImages: boolean;
  imageStyle: ImageTemplateStyle;
  aspectRatio: ImageAspectRatio;
  palette: ColorPaletteId;
  scrimOpacity: number; // 0.35 - 0.85
  showAccountWatermark: boolean;
}

export interface BackgroundPhotoItem {
  id: string;
  name: string;
  src: string;
  isBuiltin?: boolean;
}

export interface GeneratedPostBundle {
  index: number;
  text: string;
  topic: string;
  headline: string;
  subhead: string;
  bullets: string[];
  imageFilename: string;
  imageStyleUsed: Exclude<ImageTemplateStyle, 'mix'>;
  imageDataUrl: string;
  imageBytes?: Uint8Array;
}

export const BUILTIN_SAMPLE_BACKGROUNDS: BackgroundPhotoItem[] = [
  {
    id: 'builtin-tobita-lanterns',
    name: '飛田新地風_提灯風景1.jpg',
    src: tobitaLanternsBg,
    isBuiltin: true,
  },
  {
    id: 'builtin-osaka-machiya',
    name: '飛田新地風_町家路地2.jpg',
    src: osakaMachiyaBg,
    isBuiltin: true,
  },
];

export interface AccountPreset {
  id: string;
  label: string;
  design: Omit<AccountDesignInput, 'count' | 'enableSpintax' | 'generateImages'>;
}

export const ACCOUNT_DESIGN_PRESETS: AccountPreset[] = [
  {
    id: 'tobita_test',
    label: '飛田新地・風景×求人案内 (テスト)',
    design: {
      genre: '飛田新地・大阪エリアの安心ワーク紹介',
      target: '短期集中でしっかり稼ぎたい未経験・経験者の方',
      tone: 'polite',
      painPoint: '雰囲気や給与システムが事前に分からず一歩踏み出せない',
      benefitPoints: '見学のみOK・日払い対応\n1日2〜3時間の短時間シフト\nプライバシー徹底・個室待機完備\n専任スタッフの事前サポート',
      cta: '雰囲気や条件の詳細はプロフィールの案内ページからご確認ください',
      url: 'https://osaka-shinodayamashinchivip.pages.dev/blog/osaka-shinodayama-high-guarantee-guide',
      topic: '飛田新地',
      imageStyle: 'photo_overlay',
      aspectRatio: '1:1',
      palette: 'amber_night',
      scrimOpacity: 0.62,
      showAccountWatermark: true,
    },
  },
  {
    id: 'shinodayama_guide',
    label: '信太山新地・安心システム解説',
    design: {
      genre: '信太山新地の働き方・システム解説',
      target: '身バレ防止・自由出勤を重視したい方',
      tone: 'empathetic',
      painPoint: '身バレや強引な勧誘がないか不安で相談できない',
      benefitPoints: 'アリバイ・送迎・寮完備\nお友達同士の見学も大歓迎\nノルマ・罰金いっさい無し\n自分のペースで週1から勤務OK',
      cta: '不安な点の質問だけでも気軽にプロフのリンクからどうぞ',
      url: 'https://osaka-shinodayamashinchivip.pages.dev/blog/shinodayamashinchi-privacy-guide',
      topic: '信太山新地',
      imageStyle: 'mix',
      aspectRatio: '1:1',
      palette: 'emerald_slate',
      scrimOpacity: 0.58,
      showAccountWatermark: true,
    },
  },
  {
    id: 'checklist_tips',
    label: 'SNS運用・チェックリスト図解',
    design: {
      genre: 'Threads運用・自動化ノウハウ',
      target: '投稿ネタ切れや画像作成の手間に悩む運用者',
      tone: 'expert',
      painPoint: '毎日投稿と画像作りに時間がかかりすぎて続かない',
      benefitPoints: 'フック×本文×CTAの型を固定\nスピンタックスで重複判定を回避\n背景写真＋自動文字入れで時短\n1日3分で100件分のストック完成',
      cta: '保存して次回の投稿作成のチェックリストに活用してください',
      url: '',
      topic: 'SNS運用',
      imageStyle: 'checklist',
      aspectRatio: '4:5',
      palette: 'ivory_ink',
      scrimOpacity: 0.55,
      showAccountWatermark: true,
    },
  },
];

interface PaletteTheme {
  bgPrimary: string;
  bgSecondary: string;
  accent: string;
  accentSoft: string;
  textMain: string;
  textMuted: string;
  border: string;
  isLight: boolean;
}

const PALETTES: Record<ColorPaletteId, PaletteTheme> = {
  amber_night: {
    bgPrimary: '#0b0c10',
    bgSecondary: '#161922',
    accent: '#f59e0b',
    accentSoft: 'rgba(245, 158, 11, 0.16)',
    textMain: '#f8fafc',
    textMuted: '#cbd5e1',
    border: 'rgba(245, 158, 11, 0.35)',
    isLight: false,
  },
  emerald_slate: {
    bgPrimary: '#091114',
    bgSecondary: '#102026',
    accent: '#10b981',
    accentSoft: 'rgba(16, 185, 129, 0.16)',
    textMain: '#f8fafc',
    textMuted: '#cbd5e1',
    border: 'rgba(16, 185, 129, 0.35)',
    isLight: false,
  },
  crimson_gold: {
    bgPrimary: '#12090c',
    bgSecondary: '#1f1015',
    accent: '#f43f5e',
    accentSoft: 'rgba(244, 63, 94, 0.16)',
    textMain: '#fff1f2',
    textMuted: '#fda4af',
    border: 'rgba(244, 63, 94, 0.35)',
    isLight: false,
  },
  ivory_ink: {
    bgPrimary: '#f6f4ef',
    bgSecondary: '#ffffff',
    accent: '#0f766e',
    accentSoft: 'rgba(15, 118, 110, 0.10)',
    textMain: '#18181b',
    textMuted: '#52525b',
    border: '#d6d3cd',
    isLight: true,
  },
};

const HOOK_PATTERNS: Record<AccountTone, string[]> = {
  polite: [
    '【保存版】{topic}で失敗しないために最初に知っておくべきこと',
    '「{pain}」と迷われている方へ、現場のリアルをお伝えします。',
    '{target}から特にご相談が多い疑問を分かりやすく整理しました。',
    '実はあまり知られていませんが、{genre}には安心できる共通点があります。',
    '初めての方ほど大切にしてほしい、{topic}の3つのチェック基準。',
    '雰囲気を見てから決めたい方へ。事前相談でよく聞かれるポイントまとめ。',
    '無駄な不安を減らして納得して選ぶための{genre}ガイド。',
    '短時間で効率よく結果を出したい方が、最初に見直すべきポイント。',
  ],
  friendly: [
    'ぶっちゃけ「{pain}」って思ってる人、かなり多いですよね…！',
    '{target}にこそ知ってほしい、{topic}のリアルな話✨',
    'これ知らないまま動くのはもったいない！{genre}のポイントまとめ🙌',
    'よく質問をもらう「実際どうなの？」に正直にお答えします💬',
    '最初の一歩がいちばん緊張するけど、実はこんな雰囲気なんです🌿',
    '気になるけど聞きづらいこと、ここでまとめてスッキリ解決しましょう☕️',
    '自分のペースを大事にしたい人ほど、この選び方がしっくりきます✨',
    '保存してあとで見返せるように、大事なポイントだけギュッとまとめました📌',
  ],
  expert: [
    '【結論】{topic}で後悔しない人は、必ず「この基準」を先に確認しています。',
    '「{pain}」を放置したまま選ぶと、時間と労力をロスする理由。',
    '{target}向けの{genre}：成果と安心感を両立する設計ルール。',
    '曖昧な情報に振り回されないための、チェック項目リスト公開。',
    '効率よく条件を整える人が実践している3ステップ手順。',
    'ミスマッチを防ぐために、事前に必ず確認すべきポイント一覧。',
    '再現性高く進めるための、{topic}・実践テンプレート整理。',
    '迷う時間をゼロにする、シンプルな判断軸をまとめました。',
  ],
  empathetic: [
    '「{pain}」…その気持ち、すごくよく分かります。',
    '無理をしてまで頑張らなくて大丈夫。{target}に合った選び方があります。',
    '不安があるまま動かなくてOK。まずは雰囲気を知るところから始めてみませんか？',
    '誰にも相談できずに一人で悩んでいた方に、そっと届けたいお話。',
    '{genre}は「知ってから決める」だけで、気持ちの負担がぐっと軽くなります。',
    '大切なのは、あなた自身の安心とペースを守れる環境かどうか。',
    '「思っていたより落ち着いた雰囲気だった」という声が多い理由とは。',
    '焦らず自分のタイミングで納得できるよう、ポイントをまとめました。',
  ],
};

const BODY_CONNECTORS_SPINTAX = [
  '{まずは雰囲気を知るだけでもOKです|最初の一歩は情報収集からで十分です|事前確認だけで不安の大半は解消できます}。',
  '{条件やペースを重視したい方ほど|無理なく続けたい方ほど|初めて検討される方ほど}、事前のすり合わせが鍵になります。',
  '{見学や質問だけでも歓迎される環境|プライバシー配慮が整っている環境|サポート体制が明確な環境}を選ぶのがおすすめです。',
  '{1人で抱え込まずに|焦って決めずに|比較しながら}、自分に合う選択肢を確かめてみてください。',
];

const BODY_CONNECTORS_PLAIN = [
  'まずは雰囲気を知るだけでも十分な前進になります。',
  '条件やペースを重視したい方ほど、事前のすり合わせが大切です。',
  'サポート体制やプライバシー配慮が明確な環境を選ぶと安心です。',
  '焦って決めずに、自分に合う選択肢を一つずつ確かめてみてください。',
];

const SPINTAX_CLOSERS = [
  '{気になる点があればいつでも参考にしてみてください|まずはブックマークして見返してみてください|一歩目の判断材料になれば嬉しいです}',
  '{無理のないペースで進めていきましょう|自分に合った環境選びを応援しています|納得できる選択ができるよう願っています}',
];

function shortText(text: string, maxLen: number): string {
  const clean = text.replace(/\s+/g, ' ').trim();
  if (clean.length <= maxLen) return clean;
  return `${clean.slice(0, maxLen - 1)}…`;
}

function buildCatchHeadline(hookRaw: string, topic: string, genre: string, idx: number): string {
  // Generate crisp 12-26 char Japanese catchcopy suitable for canvas typography
  const customHeadlines = [
    `${topic}・失敗しない選び方`,
    `知らないと損する${topic}の基本`,
    `初めてでも安心の${topic}ガイド`,
    `${topic}のリアルな雰囲気とは`,
    `後悔しないための3つの基準`,
    `無理なく続くワークスタイル`,
    `保存版・事前チェック項目`,
    `短時間×安心環境のポイント`,
  ];
  if (idx % 2 === 0) {
    return customHeadlines[(idx / 2) % customHeadlines.length];
  }
  const stripped = hookRaw.replace(/【[^】]+】/g, '').replace(/[「」]/g, '').trim();
  return shortText(stripped || genre, 26);
}

export function generatePostSpecs(
  design: AccountDesignInput,
  _accountId: string
): Omit<GeneratedPostBundle, 'imageDataUrl' | 'imageBytes'>[] {
  const topic = design.topic.trim() || 'おすすめ';
  const genre = design.genre.trim() || topic;
  const target = design.target.trim() || 'これから始める方';
  const pain = design.painPoint.trim() || '何から確認すればいいか分からない';
  const cta = design.cta.trim() || '詳細はプロフィールの案内をご覧ください。';
  const url = design.url.trim();

  const rawBullets = design.benefitPoints
    .split(/\r?\n/)
    .map((s) => s.trim())
    .filter(Boolean);
  const bullets =
    rawBullets.length > 0
      ? rawBullets.slice(0, 4)
      : [
          '事前の条件・システム確認で安心',
          '自分のペースに合わせた柔軟な選択',
          'プライバシーとサポート体制を重視',
          'まずは見学・相談からスタート可能',
        ];

  const hooks = HOOK_PATTERNS[design.tone] || HOOK_PATTERNS.polite;
  const count = Math.max(1, Math.min(200, design.count || 10));
  const items: Omit<GeneratedPostBundle, 'imageDataUrl' | 'imageBytes'>[] = [];
  const styleCycle: Exclude<ImageTemplateStyle, 'mix'>[] = ['photo_overlay', 'diagram_card', 'checklist'];

  for (let i = 0; i < count; i++) {
    const hookTemplate = hooks[i % hooks.length];
    const hook = hookTemplate
      .replace(/\{topic\}/g, topic)
      .replace(/\{genre\}/g, genre)
      .replace(/\{target\}/g, target)
      .replace(/\{pain\}/g, pain);

    // Rotate 2 bullets per post so every post focuses on a fresh angle
    const b1 = bullets[i % bullets.length];
    const b2 = bullets[(i + 1) % bullets.length];
    const connector = design.enableSpintax
      ? BODY_CONNECTORS_SPINTAX[i % BODY_CONNECTORS_SPINTAX.length]
      : BODY_CONNECTORS_PLAIN[i % BODY_CONNECTORS_PLAIN.length];

    const closer = design.enableSpintax
      ? SPINTAX_CLOSERS[i % SPINTAX_CLOSERS.length]
      : 'ぜひ参考にしてみてください。';

    const bulletLines =
      design.tone === 'expert'
        ? `① ${b1}\n② ${b2}`
        : `・${b1}\n・${b2}`;

    const ctaLine = design.enableSpintax
      ? `{${cta}|${closer}。${cta}}`
      : cta;

    const parts = [hook, '', bulletLines, connector, '', ctaLine];
    if (url) {
      parts.push(url);
    }

    const fullText = parts.join('\n').trim();
    const numStr = String(i + 1).padStart(3, '0');
    const imageFilename = design.generateImages ? `gen_${numStr}.png` : '';
    const imageStyleUsed: Exclude<ImageTemplateStyle, 'mix'> =
      design.imageStyle === 'mix' ? styleCycle[i % styleCycle.length] : design.imageStyle;

    items.push({
      index: i + 1,
      text: fullText,
      topic,
      headline: buildCatchHeadline(hook, topic, genre, i),
      subhead: shortText(`${target}向け · ${b1}`, 34),
      bullets,
      imageFilename,
      imageStyleUsed,
    });
  }

  return items;
}

// =====================================================================
// HTML5 Canvas Japanese Typography & SNS Graphic Rendering Engine
// =====================================================================

const imageCache = new Map<string, HTMLImageElement>();

export function loadImageElement(src: string): Promise<HTMLImageElement> {
  if (!src) return Promise.reject(new Error('Empty image src'));
  const cached = imageCache.get(src);
  if (cached && cached.complete && cached.naturalWidth > 0) {
    return Promise.resolve(cached);
  }
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.onload = () => {
      imageCache.set(src, img);
      resolve(img);
    };
    img.onerror = (err) => reject(err);
    img.src = src;
  });
}

function wrapJapaneseLines(
  ctx: CanvasRenderingContext2D,
  text: string,
  maxWidth: number,
  maxLines = 4
): string[] {
  const lines: string[] = [];
  const paragraphs = text.split('\n');

  for (const para of paragraphs) {
    let current = '';
    for (const ch of Array.from(para)) {
      const next = current + ch;
      if (ctx.measureText(next).width > maxWidth && current.length > 0) {
        lines.push(current);
        current = ch;
        if (lines.length >= maxLines) break;
      } else {
        current = next;
      }
    }
    if (current && lines.length < maxLines) {
      lines.push(current);
    }
    if (lines.length >= maxLines) break;
  }

  if (lines.length === maxLines) {
    const last = lines[maxLines - 1];
    if (ctx.measureText(last).width > maxWidth - 20) {
      lines[maxLines - 1] = `${last.slice(0, Math.max(1, last.length - 1))}…`;
    }
  }
  return lines;
}

function drawRoundedRect(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  r: number
) {
  const radius = Math.min(r, w / 2, h / 2);
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.arcTo(x + w, y, x + w, y + h, radius);
  ctx.arcTo(x + w, y + h, x, y + h, radius);
  ctx.arcTo(x, y + h, x, y, radius);
  ctx.arcTo(x, y, x + w, y, radius);
  ctx.closePath();
}

function drawImageCover(
  ctx: CanvasRenderingContext2D,
  img: HTMLImageElement,
  width: number,
  height: number
) {
  const imgRatio = img.naturalWidth / Math.max(1, img.naturalHeight);
  const canvasRatio = width / height;
  let drawW = width;
  let drawH = height;
  let offsetX = 0;
  let offsetY = 0;

  if (imgRatio > canvasRatio) {
    drawH = height;
    drawW = height * imgRatio;
    offsetX = (width - drawW) / 2;
  } else {
    drawW = width;
    drawH = width / imgRatio;
    offsetY = (height - drawH) / 2;
  }
  ctx.drawImage(img, offsetX, offsetY, drawW, drawH);
}

const JP_FONT =
  '"Hiragino Kaku Gothic ProN", "Hiragino Sans", "Yu Gothic", "Meiryo", "Noto Sans JP", sans-serif';

export interface RenderImageOptions {
  aspectRatio: ImageAspectRatio;
  style: Exclude<ImageTemplateStyle, 'mix'>;
  palette: ColorPaletteId;
  scrimOpacity: number;
  topic: string;
  headline: string;
  subhead: string;
  bullets: string[];
  accountId: string;
  showWatermark: boolean;
  backgroundImg?: HTMLImageElement | null;
  postIndex: number;
}

export function renderPostImageOnCanvas(
  canvas: HTMLCanvasElement,
  opts: RenderImageOptions
): void {
  const width = 1080;
  const height = opts.aspectRatio === '4:5' ? 1350 : 1080;
  canvas.width = width;
  canvas.height = height;

  const ctx = canvas.getContext('2d');
  if (!ctx) return;

  const theme = PALETTES[opts.palette] || PALETTES.amber_night;
  const pad = 76;

  ctx.clearRect(0, 0, width, height);

  if (opts.style === 'photo_overlay') {
    // 1. Background Photo or Rich Atmospheric Fallback Gradient
    if (opts.backgroundImg && opts.backgroundImg.naturalWidth > 0) {
      drawImageCover(ctx, opts.backgroundImg, width, height);
    } else {
      const grad = ctx.createLinearGradient(0, 0, width, height);
      grad.addColorStop(0, '#18181b');
      grad.addColorStop(0.5, '#27272a');
      grad.addColorStop(1, '#09090b');
      ctx.fillStyle = grad;
      ctx.fillRect(0, 0, width, height);
    }

    // 2. Measured Editorial Scrim (stronger at bottom/center for >= 4.5:1 legibility)
    const scrimAlpha = Math.max(0.25, Math.min(0.9, opts.scrimOpacity));
    const scrim = ctx.createLinearGradient(0, 0, 0, height);
    scrim.addColorStop(0, `rgba(8, 10, 14, ${(scrimAlpha * 0.45).toFixed(2)})`);
    scrim.addColorStop(0.45, `rgba(8, 10, 14, ${(scrimAlpha * 0.72).toFixed(2)})`);
    scrim.addColorStop(1, `rgba(6, 8, 12, ${Math.min(0.94, scrimAlpha + 0.22).toFixed(2)})`);
    ctx.fillStyle = scrim;
    ctx.fillRect(0, 0, width, height);

    // Subtle inset hairline frame
    ctx.strokeStyle = 'rgba(255, 255, 255, 0.16)';
    ctx.lineWidth = 2;
    ctx.strokeRect(38, 38, width - 76, height - 76);

    // Top-left editorial kicker bar + Topic label
    ctx.fillStyle = theme.accent;
    ctx.fillRect(pad, pad + 6, 42, 6);

    ctx.font = `700 28px ${JP_FONT}`;
    ctx.fillStyle = '#f8fafc';
    ctx.textBaseline = 'middle';
    ctx.fillText(opts.topic || 'GUIDE', pad + 58, pad + 10);

    // Bottom editorial text composition
    const maxTextWidth = width - pad * 2;
    ctx.font = `800 68px ${JP_FONT}`;
    const headlineLines = wrapJapaneseLines(ctx, opts.headline, maxTextWidth, 3);
    const lineHeight = 90;
    const subheadHeight = opts.subhead ? 54 : 0;
    const footerHeight = opts.showWatermark ? 56 : 16;

    const totalTextBlockH = headlineLines.length * lineHeight + subheadHeight + footerHeight + 36;
    let cursorY = height - pad - totalTextBlockH;

    // Accent bar above headline
    ctx.fillStyle = theme.accent;
    ctx.fillRect(pad, cursorY, 72, 8);
    cursorY += 40;

    // Main Japanese catchcopy with crisp shadow
    ctx.font = `800 68px ${JP_FONT}`;
    ctx.fillStyle = '#ffffff';
    ctx.textBaseline = 'top';
    ctx.shadowColor = 'rgba(0, 0, 0, 0.75)';
    ctx.shadowBlur = 18;
    ctx.shadowOffsetY = 4;

    for (const line of headlineLines) {
      ctx.fillText(line, pad, cursorY);
      cursorY += lineHeight;
    }

    // Reset shadow
    ctx.shadowColor = 'transparent';
    ctx.shadowBlur = 0;
    ctx.shadowOffsetY = 0;

    if (opts.subhead) {
      cursorY += 12;
      ctx.font = `500 32px ${JP_FONT}`;
      ctx.fillStyle = 'rgba(241, 245, 249, 0.90)';
      const subLines = wrapJapaneseLines(ctx, opts.subhead, maxTextWidth, 2);
      for (const sline of subLines) {
        ctx.fillText(sline, pad, cursorY);
        cursorY += 44;
      }
    }

    if (opts.showWatermark) {
      ctx.strokeStyle = 'rgba(255, 255, 255, 0.22)';
      ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(pad, height - pad - 44);
      ctx.lineTo(width - pad, height - pad - 44);
      ctx.stroke();

      ctx.font = `600 26px ${JP_FONT}`;
      ctx.fillStyle = 'rgba(226, 232, 240, 0.85)';
      ctx.textBaseline = 'bottom';
      ctx.fillText(`@${opts.accountId || 'threads'}`, pad, height - pad);

      ctx.textAlign = 'right';
      ctx.fillStyle = theme.accent;
      ctx.fillText(`NO.${String(opts.postIndex).padStart(2, '0')}`, width - pad, height - pad);
      ctx.textAlign = 'left';
    }
    return;
  }

  // 2. Diagram Card / Checklist Canvas Background
  ctx.fillStyle = theme.bgPrimary;
  ctx.fillRect(0, 0, width, height);

  // Subtle decorative grid lines for clean editorial diagram look
  ctx.strokeStyle = theme.isLight ? 'rgba(0, 0, 0, 0.04)' : 'rgba(255, 255, 255, 0.035)';
  ctx.lineWidth = 1;
  const gridStep = 72;
  for (let gx = gridStep; gx < width; gx += gridStep) {
    ctx.beginPath();
    ctx.moveTo(gx, 0);
    ctx.lineTo(gx, height);
    ctx.stroke();
  }
  for (let gy = gridStep; gy < height; gy += gridStep) {
    ctx.beginPath();
    ctx.moveTo(0, gy);
    ctx.lineTo(width, gy);
    ctx.stroke();
  }

  // Optional subtle background photo tint inside diagram header
  if (opts.backgroundImg && opts.backgroundImg.naturalWidth > 0) {
    ctx.save();
    ctx.globalAlpha = theme.isLight ? 0.08 : 0.14;
    drawImageCover(ctx, opts.backgroundImg, width, height);
    ctx.restore();
  }

  const maxW = width - pad * 2;
  let y = pad;

  // Header kicker
  ctx.font = `700 28px ${JP_FONT}`;
  ctx.fillStyle = theme.accent;
  ctx.textBaseline = 'top';
  const modeLabel = opts.style === 'checklist' ? `${opts.topic} · 保存用チェックリスト` : `${opts.topic} · ポイント図解`;
  ctx.fillText(modeLabel, pad, y);
  y += 50;

  // Headline
  ctx.font = `800 54px ${JP_FONT}`;
  ctx.fillStyle = theme.textMain;
  const titleLines = wrapJapaneseLines(ctx, opts.headline, maxW, 2);
  for (const line of titleLines) {
    ctx.fillText(line, pad, y);
    y += 72;
  }

  y += 14;
  ctx.strokeStyle = theme.border;
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(pad, y);
  ctx.lineTo(width - pad, y);
  ctx.stroke();
  y += 38;

  // Rows (up to 4 items)
  const listItems = (opts.bullets.length > 0 ? opts.bullets : ['事前確認でミスマッチ防止', '無理のない時間設定', '安心のサポート体制']).slice(0, 4);
  const availableH = height - y - pad - 76;
  const gap = 22;
  const cardH = Math.min(165, Math.floor((availableH - gap * (listItems.length - 1)) / listItems.length));

  listItems.forEach((item, idx) => {
    const cy = y + idx * (cardH + gap);

    drawRoundedRect(ctx, pad, cy, maxW, cardH, 16);
    ctx.fillStyle = theme.bgSecondary;
    ctx.fill();
    ctx.strokeStyle = theme.border;
    ctx.lineWidth = 1.5;
    ctx.stroke();

    if (opts.style === 'checklist') {
      // Checkbox icon box
      const boxSize = 46;
      const bx = pad + 28;
      const by = cy + (cardH - boxSize) / 2;
      drawRoundedRect(ctx, bx, by, boxSize, boxSize, 10);
      ctx.fillStyle = theme.accentSoft;
      ctx.fill();
      ctx.strokeStyle = theme.accent;
      ctx.lineWidth = 2.5;
      ctx.stroke();

      // Checkmark
      ctx.strokeStyle = theme.accent;
      ctx.lineWidth = 4;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.beginPath();
      ctx.moveTo(bx + 11, by + 23);
      ctx.lineTo(bx + 20, by + 32);
      ctx.lineTo(bx + 35, by + 14);
      ctx.stroke();
    } else {
      // Numbered step index 01..04
      ctx.font = `800 36px "JetBrains Mono", monospace`;
      ctx.fillStyle = theme.accent;
      ctx.textBaseline = 'middle';
      ctx.fillText(String(idx + 1).padStart(2, '0'), pad + 28, cy + cardH / 2);
    }

    // Item text
    ctx.font = `600 33px ${JP_FONT}`;
    ctx.fillStyle = theme.textMain;
    ctx.textBaseline = 'middle';
    const textX = pad + 102;
    const itemLines = wrapJapaneseLines(ctx, item, maxW - 130, 2);
    if (itemLines.length === 1) {
      ctx.fillText(itemLines[0], textX, cy + cardH / 2);
    } else {
      ctx.fillText(itemLines[0], textX, cy + cardH / 2 - 22);
      ctx.fillText(itemLines[1], textX, cy + cardH / 2 + 22);
    }
  });

  // Footer
  if (opts.showWatermark) {
    ctx.font = `600 26px ${JP_FONT}`;
    ctx.fillStyle = theme.textMuted;
    ctx.textBaseline = 'bottom';
    ctx.fillText(`@${opts.accountId || 'threads'}`, pad, height - pad);

    ctx.textAlign = 'right';
    ctx.fillStyle = theme.accent;
    ctx.fillText(opts.subhead ? shortText(opts.subhead, 24) : `${opts.topic} GUIDE`, width - pad, height - pad);
    ctx.textAlign = 'left';
  }
}

export async function renderCanvasToDataUrlAndBytes(
  canvas: HTMLCanvasElement
): Promise<{ dataUrl: string; bytes: Uint8Array }> {
  const dataUrl = canvas.toDataURL('image/png');
  const blob: Blob = await new Promise((resolve) => {
    canvas.toBlob((b) => resolve(b || new Blob()), 'image/png');
  });
  const buf = await blob.arrayBuffer();
  return {
    dataUrl,
    bytes: new Uint8Array(buf),
  };
}
