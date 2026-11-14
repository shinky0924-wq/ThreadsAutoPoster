import tobitaLanternsBg from '../assets/images/tobita_street_lanterns_1794644991845.jpg';
import osakaMachiyaBg from '../assets/images/osaka_machiya_dusk_1794644992304.jpg';

export type AccountTone = 'polite' | 'friendly' | 'expert' | 'empathetic';
export type ConcreteImageTemplateStyle =
  | 'photo_overlay'
  | 'magazine_clean'
  | 'chat_bubble'
  | 'neon_night'
  | 'step_roadmap'
  | 'qa_card'
  | 'diagram_card'
  | 'checklist';
export type ImageTemplateStyle = ConcreteImageTemplateStyle | 'auto' | 'mix' | 'all_mix';
export type ImageAspectRatio = '1:1' | '4:5';
export type ColorPaletteId =
  | 'amber_night'
  | 'emerald_slate'
  | 'crimson_gold'
  | 'ivory_ink'
  | 'magazine_ivory_navy'
  | 'neon_cyan_night';

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
  scrimOpacity: number; // 0.30 - 0.85
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
  painPoint: string;
  bullets: string[];
  imageFilename: string;
  imageStyleUsed: ConcreteImageTemplateStyle;
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
      imageStyle: 'auto',
      aspectRatio: '1:1',
      palette: 'amber_night',
      scrimOpacity: 0.48,
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
      imageStyle: 'auto',
      aspectRatio: '1:1',
      palette: 'emerald_slate',
      scrimOpacity: 0.48,
      showAccountWatermark: true,
    },
  },
  {
    id: 'checklist_tips',
    label: 'SNS運用・ステップ図解',
    design: {
      genre: 'Threads運用・自動化ノウハウ',
      target: '投稿ネタ切れや画像作成の手間に悩む運用者',
      tone: 'expert',
      painPoint: '毎日投稿と画像作りに時間がかかりすぎて続かない',
      benefitPoints: 'フック×本文×CTAの型を固定\nスピンタックスで重複判定を回避\nアカウント設計に合わせて自動画像化\n1日3分で100件分のストック完成',
      cta: '保存して次回の投稿作成のチェックリストに活用してください',
      url: '',
      topic: 'SNS運用',
      imageStyle: 'step_roadmap',
      aspectRatio: '4:5',
      palette: 'ivory_ink',
      scrimOpacity: 0.48,
      showAccountWatermark: true,
    },
  },
  {
    id: 'beauty_magazine',
    label: '美容・カフェ・雑誌マガジン風',
    design: {
      genre: '大人女子の週末カフェ・美容習慣',
      target: '無理なく垢抜けたい20〜30代女性',
      tone: 'friendly',
      painPoint: '忙しくて自分のケアやリフレッシュの時間が取れない',
      benefitPoints: '朝10分で気分が整うルーティン\n肌と心に優しいシンプル習慣\n週末だけのご褒美カフェ時間\nプチプラでも続く厳選アイテム',
      cta: 'あとで見返せるように保存＆フォローしてもらえると嬉しいです',
      url: '',
      topic: '美容習慣',
      imageStyle: 'magazine_clean',
      aspectRatio: '4:5',
      palette: 'magazine_ivory_navy',
      scrimOpacity: 0.45,
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
  magazine_ivory_navy: {
    bgPrimary: '#faf7f2',
    bgSecondary: '#fffefb',
    accent: '#1e3a8a',
    accentSoft: 'rgba(30, 58, 138, 0.10)',
    textMain: '#18181b',
    textMuted: '#52525b',
    border: '#d4cfc6',
    isLight: true,
  },
  neon_cyan_night: {
    bgPrimary: '#070910',
    bgSecondary: '#0e1422',
    accent: '#22d3ee',
    accentSoft: 'rgba(34, 211, 238, 0.16)',
    textMain: '#f8fafc',
    textMuted: '#bae6fd',
    border: 'rgba(34, 211, 238, 0.38)',
    isLight: false,
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

export function resolveDesignImageTemplates(
  design: Pick<AccountDesignInput, 'imageStyle' | 'genre' | 'topic' | 'target' | 'painPoint' | 'tone'>,
  hasPhotos = true
): ConcreteImageTemplateStyle[] {
  const pref = design.imageStyle || 'auto';
  if (
    pref !== 'auto' &&
    pref !== 'mix' &&
    pref !== 'all_mix'
  ) {
    return [pref];
  }

  const allStyles: ConcreteImageTemplateStyle[] = hasPhotos
    ? ['photo_overlay', 'magazine_clean', 'chat_bubble', 'neon_night', 'step_roadmap', 'qa_card', 'diagram_card', 'checklist']
    : ['magazine_clean', 'chat_bubble', 'neon_night', 'step_roadmap', 'qa_card', 'diagram_card', 'checklist'];

  if (pref === 'all_mix') {
    return allStyles;
  }
  if (pref === 'mix') {
    return hasPhotos
      ? ['photo_overlay', 'diagram_card', 'checklist']
      : ['magazine_clean', 'diagram_card', 'checklist'];
  }

  const joined = `${design.genre || ''} ${design.topic || ''} ${design.target || ''} ${design.painPoint || ''}`.toLowerCase();
  const tone = design.tone || 'polite';

  let seq: ConcreteImageTemplateStyle[];
  if (/新地|ナイト|高収入|キャバ|夜職|ガールズ|ラウンジ|体験入店|短期集中/.test(joined)) {
    if (tone === 'friendly') {
      seq = ['chat_bubble', 'neon_night', 'qa_card', 'magazine_clean'];
    } else if (tone === 'empathetic') {
      seq = ['chat_bubble', 'qa_card', 'magazine_clean', 'neon_night'];
    } else {
      seq = ['neon_night', 'magazine_clean', 'qa_card', 'chat_bubble'];
    }
  } else if (/副業|キャリア|転職|sns|threads|ai|ビジネス|稼|運用|ノウハウ|投資/.test(joined)) {
    seq =
      tone === 'expert'
        ? ['step_roadmap', 'diagram_card', 'qa_card', 'checklist']
        : ['step_roadmap', 'magazine_clean', 'chat_bubble', 'checklist'];
  } else if (/美容|カフェ|暮|ダイエット|コスメ|女性|育児|習慣/.test(joined)) {
    seq = ['magazine_clean', 'checklist', 'chat_bubble', 'qa_card'];
  } else {
    const toneDefaults: Record<AccountTone, ConcreteImageTemplateStyle[]> = {
      expert: ['step_roadmap', 'diagram_card', 'qa_card', 'checklist'],
      friendly: ['chat_bubble', 'magazine_clean', 'qa_card', 'checklist'],
      empathetic: ['chat_bubble', 'qa_card', 'magazine_clean', 'checklist'],
      polite: ['magazine_clean', 'qa_card', 'step_roadmap', 'checklist'],
    };
    seq = toneDefaults[tone] || ['magazine_clean', 'qa_card', 'step_roadmap', 'checklist'];
  }

  if (hasPhotos && !seq.includes('photo_overlay')) {
    return ['photo_overlay', seq[0], seq[1], 'photo_overlay', ...seq.slice(2)];
  }
  return seq;
}

export function generatePostSpecs(
  design: AccountDesignInput,
  _accountId: string,
  hasPhotos = true
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
  const styleCycle = resolveDesignImageTemplates(design, hasPhotos);

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
        : `・{${b1}|${b1}}\n・${b2}`.replace(/\{([^|]+)\|\1\}/g, '$1');

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
    const imageStyleUsed = styleCycle[i % styleCycle.length];

    items.push({
      index: i + 1,
      text: fullText,
      topic,
      headline: buildCatchHeadline(hook, topic, genre, i),
      subhead: shortText(`${target}向け · ${b1}`, 34),
      painPoint: pain,
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
const JP_MINCHO_FONT =
  '"Yu Mincho", "Hiragino Mincho ProN", "Noto Serif JP", Georgia, serif';

export interface RenderImageOptions {
  aspectRatio: ImageAspectRatio;
  style: ConcreteImageTemplateStyle;
  palette: ColorPaletteId;
  scrimOpacity: number;
  topic: string;
  headline: string;
  subhead: string;
  painPoint?: string;
  bullets: string[];
  accountId: string;
  showWatermark: boolean;
  backgroundImg?: HTMLImageElement | null;
  postIndex: number;
}

function fillSmoothAmbientBackground(
  ctx: CanvasRenderingContext2D,
  theme: PaletteTheme,
  width: number,
  height: number,
  glowAlpha = 0.22
) {
  const grad = ctx.createLinearGradient(0, 0, width * 0.4, height);
  grad.addColorStop(0, theme.bgSecondary);
  grad.addColorStop(1, theme.bgPrimary);
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, width, height);

  const glow = ctx.createRadialGradient(width * 0.78, height * 0.16, 40, width * 0.78, height * 0.16, width * 0.75);
  glow.addColorStop(0, theme.accentSoft);
  glow.addColorStop(1, 'rgba(0, 0, 0, 0)');
  ctx.save();
  ctx.globalAlpha = Math.min(1, glowAlpha * 3.5);
  ctx.fillStyle = glow;
  ctx.fillRect(0, 0, width, height);
  ctx.restore();
}

function isLightHex(hex: string): boolean {
  const h = hex.replace('#', '');
  if (h.length !== 6) return false;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return r + g + b > 460;
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
  const painText = (opts.painPoint || '何から確認すればいいか分からず不安…').trim();
  const listItems = (
    opts.bullets.length > 0 ? opts.bullets : ['事前確認でミスマッチ防止', '無理のない時間設定', '安心のサポート体制']
  ).slice(0, 4);

  ctx.clearRect(0, 0, width, height);

  if (opts.style === 'photo_overlay') {
    const hasPhoto = Boolean(opts.backgroundImg && opts.backgroundImg.naturalWidth > 0);
    if (hasPhoto && opts.backgroundImg) {
      drawImageCover(ctx, opts.backgroundImg, width, height);
      const scrimAlpha = Math.max(0.15, Math.min(0.85, opts.scrimOpacity));
      const scrim = ctx.createLinearGradient(0, 0, 0, height);
      scrim.addColorStop(0, `rgba(8, 10, 14, ${(scrimAlpha * 0.18).toFixed(2)})`);
      scrim.addColorStop(0.42, `rgba(8, 10, 14, ${(scrimAlpha * 0.43).toFixed(2)})`);
      scrim.addColorStop(1, `rgba(6, 8, 12, ${Math.min(0.88, scrimAlpha + 0.34).toFixed(2)})`);
      ctx.fillStyle = scrim;
      ctx.fillRect(0, 0, width, height);
    } else {
      fillSmoothAmbientBackground(ctx, theme, width, height, 0.35);
    }

    const useLightText = hasPhoto || !theme.isLight;
    ctx.strokeStyle = useLightText ? 'rgba(255, 255, 255, 0.18)' : theme.border;
    ctx.lineWidth = 2;
    ctx.strokeRect(38, 38, width - 76, height - 76);

    ctx.fillStyle = theme.accent;
    ctx.fillRect(pad, pad + 6, 42, 6);

    ctx.font = `700 28px ${JP_FONT}`;
    ctx.fillStyle = useLightText ? '#f8fafc' : theme.textMain;
    ctx.textBaseline = 'middle';
    ctx.fillText(opts.topic || 'GUIDE', pad + 58, pad + 10);

    const maxTextWidth = width - pad * 2;
    ctx.font = `800 66px ${JP_FONT}`;
    const headlineLines = wrapJapaneseLines(ctx, opts.headline, maxTextWidth, 3);
    const lineHeight = 88;

    ctx.font = `500 32px ${JP_FONT}`;
    const subLines = opts.subhead ? wrapJapaneseLines(ctx, opts.subhead, maxTextWidth, 2) : [];
    const subLineHeight = 44;
    const subheadHeight = subLines.length > 0 ? 14 + subLines.length * subLineHeight + 18 : 0;
    const footerHeight = opts.showWatermark ? 66 : 16;

    const totalTextBlockH = 40 + headlineLines.length * lineHeight + subheadHeight + footerHeight;
    let cursorY = height - pad - totalTextBlockH;

    ctx.fillStyle = theme.accent;
    ctx.fillRect(pad, cursorY, 72, 8);
    cursorY += 40;

    ctx.font = `800 66px ${JP_FONT}`;
    ctx.fillStyle = useLightText ? '#ffffff' : theme.textMain;
    ctx.textBaseline = 'top';
    if (useLightText) {
      ctx.shadowColor = 'rgba(0, 0, 0, 0.75)';
      ctx.shadowBlur = 18;
      ctx.shadowOffsetY = 4;
    }
    for (const line of headlineLines) {
      ctx.fillText(line, pad, cursorY);
      cursorY += lineHeight;
    }
    ctx.shadowColor = 'transparent';
    ctx.shadowBlur = 0;
    ctx.shadowOffsetY = 0;

    if (subLines.length > 0) {
      cursorY += 14;
      ctx.font = `500 32px ${JP_FONT}`;
      ctx.fillStyle = useLightText ? 'rgba(241, 245, 249, 0.90)' : theme.textMuted;
      for (const sline of subLines) {
        ctx.fillText(sline, pad, cursorY);
        cursorY += subLineHeight;
      }
    }

    if (opts.showWatermark) {
      ctx.strokeStyle = useLightText ? 'rgba(255, 255, 255, 0.24)' : theme.border;
      ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(pad, height - pad - 44);
      ctx.lineTo(width - pad, height - pad - 44);
      ctx.stroke();

      ctx.font = `600 26px ${JP_FONT}`;
      ctx.fillStyle = useLightText ? 'rgba(226, 232, 240, 0.85)' : theme.textMuted;
      ctx.textBaseline = 'bottom';
      ctx.fillText(`@${opts.accountId || 'threads'}`, pad, height - pad);

      ctx.textAlign = 'right';
      ctx.fillStyle = theme.accent;
      ctx.fillText(`NO.${String(opts.postIndex).padStart(2, '0')}`, width - pad, height - pad);
      ctx.textAlign = 'left';
    }
    return;
  }

  if (opts.style === 'magazine_clean') {
    const paperBg = theme.isLight ? '#faf7f2' : '#121318';
    const cardBg = theme.isLight ? '#ffffff' : '#1a1c24';
    const inkCol = theme.isLight ? '#18181b' : '#f8fafc';
    const subCol = theme.isLight ? '#52525b' : '#b4bcc8';
    const lineCol = theme.isLight ? '#dcd6cc' : '#303440';

    ctx.fillStyle = paperBg;
    ctx.fillRect(0, 0, width, height);
    if (opts.backgroundImg && opts.backgroundImg.naturalWidth > 0) {
      ctx.save();
      ctx.globalAlpha = theme.isLight ? 0.08 : 0.14;
      drawImageCover(ctx, opts.backgroundImg, width, height);
      ctx.restore();
    }
    ctx.strokeStyle = lineCol;
    ctx.lineWidth = 2;
    ctx.strokeRect(42, 42, width - 84, height - 84);

    let y = pad + 6;
    ctx.font = `700 24px ${JP_FONT}`;
    ctx.textBaseline = 'top';
    ctx.fillStyle = theme.accent;
    ctx.fillText(`ISSUE NO.${String(opts.postIndex).padStart(2, '0')}`, pad, y);
    ctx.textAlign = 'right';
    ctx.fillStyle = subCol;
    ctx.fillText((opts.topic || 'FEATURE').toUpperCase(), width - pad, y);
    ctx.textAlign = 'left';
    y += 42;

    ctx.fillStyle = theme.accent;
    ctx.fillRect(pad, y, 88, 6);
    ctx.strokeStyle = lineCol;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(pad + 104, y + 3);
    ctx.lineTo(width - pad, y + 3);
    ctx.stroke();
    y += 42;

    const maxW = width - pad * 2;
    ctx.font = `700 58px ${JP_MINCHO_FONT}`;
    ctx.fillStyle = inkCol;
    const hLines = wrapJapaneseLines(ctx, opts.headline, maxW, 3);
    for (const hl of hLines) {
      ctx.fillText(hl, pad, y);
      y += 78;
    }

    if (opts.subhead) {
      y += 6;
      ctx.font = `400 28px ${JP_FONT}`;
      ctx.fillStyle = subCol;
      const sLines = wrapJapaneseLines(ctx, opts.subhead, maxW, 2);
      for (const sl of sLines) {
        ctx.fillText(sl, pad, y);
        y += 40;
      }
    }

    y += 28;
    const boxBottom = height - pad - 68;
    drawRoundedRect(ctx, pad, y, maxW, boxBottom - y, 16);
    ctx.fillStyle = cardBg;
    ctx.fill();
    ctx.strokeStyle = lineCol;
    ctx.lineWidth = 2;
    ctx.stroke();

    const items = listItems.slice(0, 3);
    let innerY = y + 34;
    const slotH = Math.max(92, Math.floor((boxBottom - y - 56) / Math.max(1, items.length)));
    items.forEach((itTxt, idx) => {
      ctx.font = `700 32px ${JP_MINCHO_FONT}`;
      ctx.fillStyle = theme.accent;
      ctx.fillText(`${String(idx + 1).padStart(2, '0')}.`, pad + 34, innerY + 4);

      ctx.font = `600 30px ${JP_FONT}`;
      ctx.fillStyle = inkCol;
      const wrapped = wrapJapaneseLines(ctx, itTxt, maxW - 140, 2);
      wrapped.forEach((wl, li) => {
        ctx.fillText(wl, pad + 112, innerY + li * 40);
      });

      if (idx < items.length - 1) {
        ctx.strokeStyle = lineCol;
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(pad + 34, innerY + slotH - 14);
        ctx.lineTo(width - pad - 34, innerY + slotH - 14);
        ctx.stroke();
      }
      innerY += slotH;
    });

    if (opts.showWatermark) {
      ctx.font = `600 24px ${JP_FONT}`;
      ctx.fillStyle = subCol;
      ctx.textBaseline = 'bottom';
      ctx.fillText(`@${opts.accountId || 'threads'}`, pad, height - pad);
      ctx.textAlign = 'right';
      ctx.fillStyle = theme.accent;
      ctx.fillText('EDITORIAL PICK', width - pad, height - pad);
      ctx.textAlign = 'left';
    }
    return;
  }

  if (opts.style === 'chat_bubble') {
    fillSmoothAmbientBackground(ctx, theme, width, height, 0.18);
    const qBg = theme.isLight ? '#e8ecf2' : '#202430';
    const qBorder = theme.isLight ? '#cbd5e1' : '#3a4154';
    const aBg = theme.isLight ? '#ffffff' : '#181e2a';
    const headFg = isLightHex(theme.accent) ? '#0c0c0e' : '#ffffff';

    let y = pad;
    ctx.font = `700 26px ${JP_FONT}`;
    ctx.textBaseline = 'top';
    ctx.fillStyle = theme.accent;
    ctx.fillText(`💬 ${opts.topic || '相談'} · よくあるお悩み＆回答`, pad, y);
    y += 52;
    ctx.strokeStyle = theme.border;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(pad, y);
    ctx.lineTo(width - pad, y);
    ctx.stroke();
    y += 32;

    const avS = 72;
    drawRoundedRect(ctx, pad, y, avS, avS, 36);
    ctx.fillStyle = qBg;
    ctx.fill();
    ctx.strokeStyle = qBorder;
    ctx.lineWidth = 2;
    ctx.stroke();
    ctx.font = `700 28px ${JP_FONT}`;
    ctx.fillStyle = theme.textMuted;
    ctx.fillText('Q', pad + 24, y + 20);

    const qbX1 = pad + avS + 20;
    const qbX2 = width - pad - 70;
    ctx.font = `700 32px ${JP_FONT}`;
    const qLines = wrapJapaneseLines(ctx, `「${painText}」`, qbX2 - qbX1 - 56, 3);
    const qbH = Math.max(110, 44 + qLines.length * 44);
    drawRoundedRect(ctx, qbX1, y, qbX2 - qbX1, qbH, 24);
    ctx.fillStyle = qBg;
    ctx.fill();
    ctx.strokeStyle = qBorder;
    ctx.stroke();
    ctx.fillStyle = theme.textMain;
    qLines.forEach((ql, iq) => {
      ctx.fillText(ql, qbX1 + 28, y + 24 + iq * 44);
    });
    y += qbH + 38;

    const abX1 = pad + 56;
    const abX2 = width - pad - avS - 20;
    drawRoundedRect(ctx, width - pad - avS, y, avS, avS, 36);
    ctx.fillStyle = theme.accent;
    ctx.fill();
    ctx.font = `700 28px ${JP_FONT}`;
    ctx.fillStyle = headFg;
    ctx.fillText('A', width - pad - avS + 25, y + 20);

    ctx.font = `800 40px ${JP_FONT}`;
    const aLines = wrapJapaneseLines(ctx, opts.headline, abX2 - abX1 - 56, 2);
    const abH = Math.max(106, 44 + aLines.length * 54);
    drawRoundedRect(ctx, abX1, y, abX2 - abX1, abH, 24);
    ctx.fillStyle = theme.accent;
    ctx.fill();
    ctx.fillStyle = headFg;
    aLines.forEach((al, ia) => {
      ctx.fillText(al, abX1 + 28, y + 24 + ia * 54);
    });
    y += abH + 26;

    const availBot = height - pad - 60 - y;
    const cItems = listItems.slice(0, 3);
    drawRoundedRect(ctx, abX1, y, abX2 - abX1, availBot, 24);
    ctx.fillStyle = aBg;
    ctx.fill();
    ctx.strokeStyle = theme.accent;
    ctx.lineWidth = 2;
    ctx.stroke();

    const rowStep = Math.max(66, Math.floor((availBot - 40) / Math.max(1, cItems.length)));
    let iy = y + 26;
    cItems.forEach((bt, bi) => {
      drawRoundedRect(ctx, abX1 + 26, iy + 6, 34, 34, 10);
      ctx.fillStyle = theme.accentSoft;
      ctx.fill();
      ctx.strokeStyle = theme.accent;
      ctx.lineWidth = 2;
      ctx.stroke();
      ctx.font = `700 22px ${JP_FONT}`;
      ctx.fillStyle = theme.accent;
      ctx.fillText(String(bi + 1), abX1 + 36, iy + 11);

      ctx.font = `600 30px ${JP_FONT}`;
      ctx.fillStyle = theme.textMain;
      const bLines = wrapJapaneseLines(ctx, bt, abX2 - abX1 - 108, 2);
      bLines.forEach((bl, li) => {
        ctx.fillText(bl, abX1 + 78, iy + 4 + li * 38);
      });
      iy += rowStep;
    });

    if (opts.showWatermark) {
      ctx.font = `600 25px ${JP_FONT}`;
      ctx.fillStyle = theme.textMuted;
      ctx.textBaseline = 'bottom';
      ctx.fillText(`@${opts.accountId || 'threads'}`, pad, height - pad);
      ctx.textAlign = 'right';
      ctx.fillStyle = theme.accent;
      ctx.fillText(`Q&A #${String(opts.postIndex).padStart(2, '0')}`, width - pad, height - pad);
      ctx.textAlign = 'left';
    }
    return;
  }

  if (opts.style === 'neon_night') {
    const ac = theme.isLight ? '#38bdf8' : theme.accent;
    ctx.fillStyle = '#07080e';
    ctx.fillRect(0, 0, width, height);
    const ng = ctx.createRadialGradient(width * 0.75, height * 0.2, 40, width * 0.75, height * 0.2, width * 0.8);
    ng.addColorStop(0, 'rgba(34, 211, 238, 0.16)');
    ng.addColorStop(1, 'rgba(7, 8, 14, 0)');
    ctx.fillStyle = ng;
    ctx.fillRect(0, 0, width, height);

    if (opts.backgroundImg && opts.backgroundImg.naturalWidth > 0) {
      ctx.save();
      ctx.globalAlpha = 0.18;
      drawImageCover(ctx, opts.backgroundImg, width, height);
      ctx.restore();
    }

    ctx.save();
    drawRoundedRect(ctx, 44, 44, width - 88, height - 88, 26);
    ctx.shadowColor = ac;
    ctx.shadowBlur = 24;
    ctx.strokeStyle = ac;
    ctx.lineWidth = 3;
    ctx.stroke();
    ctx.restore();

    let y = pad + 8;
    const tagTxt = `◆ ${(opts.topic || 'NIGHT').toUpperCase()}  //  PICK UP ${String(opts.postIndex).padStart(2, '0')}`;
    ctx.font = `700 24px ${JP_FONT}`;
    ctx.textBaseline = 'top';
    const tagBoxW = Math.min(width - pad * 2, Math.max(320, Math.ceil(ctx.measureText(tagTxt).width) + 42));
    drawRoundedRect(ctx, pad, y, tagBoxW, 48, 10);
    ctx.fillStyle = 'rgba(20, 24, 38, 0.88)';
    ctx.fill();
    ctx.strokeStyle = ac;
    ctx.lineWidth = 2;
    ctx.stroke();
    ctx.fillStyle = ac;
    ctx.fillText(tagTxt, pad + 18, y + 11);
    y += 82;

    const maxW = width - pad * 2;
    ctx.font = `800 58px ${JP_FONT}`;
    ctx.fillStyle = '#ffffff';
    ctx.shadowColor = 'rgba(0, 0, 0, 0.8)';
    ctx.shadowBlur = 14;
    const hLines = wrapJapaneseLines(ctx, opts.headline, maxW, 2);
    for (const hl of hLines) {
      ctx.fillText(hl, pad, y);
      y += 78;
    }
    ctx.shadowBlur = 0;

    y += 14;
    ctx.fillStyle = ac;
    ctx.fillRect(pad, y, 160, 6);
    ctx.strokeStyle = 'rgba(255, 255, 255, 0.2)';
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(pad + 176, y + 3);
    ctx.lineTo(width - pad, y + 3);
    ctx.stroke();
    y += 42;

    const nItems = listItems.slice(0, 3);
    const availH = height - pad - 68 - y;
    const gap = 24;
    const cardH = Math.min(176, Math.max(108, Math.floor((availH - gap * (nItems.length - 1)) / Math.max(1, nItems.length))));
    nItems.forEach((itTxt, idx) => {
      const cy = y + idx * (cardH + gap);
      drawRoundedRect(ctx, pad, cy, maxW, cardH, 16);
      ctx.fillStyle = 'rgba(16, 20, 32, 0.90)';
      ctx.fill();
      ctx.strokeStyle = ac;
      ctx.lineWidth = 2;
      ctx.stroke();

      ctx.fillStyle = ac;
      ctx.fillRect(pad + 22, cy + 24, 8, cardH - 48);

      ctx.font = `800 32px ${JP_FONT}`;
      ctx.fillStyle = ac;
      ctx.textBaseline = 'middle';
      ctx.fillText(String(idx + 1).padStart(2, '0'), pad + 52, cy + cardH / 2);

      ctx.font = `700 32px ${JP_FONT}`;
      ctx.fillStyle = '#f8fafc';
      const ilines = wrapJapaneseLines(ctx, itTxt, maxW - 150, 2);
      if (ilines.length === 1) {
        ctx.fillText(ilines[0], pad + 122, cy + cardH / 2);
      } else {
        ctx.fillText(ilines[0], pad + 122, cy + cardH / 2 - 22);
        ctx.fillText(ilines[1], pad + 122, cy + cardH / 2 + 22);
      }
    });

    if (opts.showWatermark) {
      ctx.font = `600 25px ${JP_FONT}`;
      ctx.fillStyle = 'rgba(203, 213, 225, 0.88)';
      ctx.textBaseline = 'bottom';
      ctx.fillText(`@${opts.accountId || 'threads'}`, pad, height - pad);
      ctx.textAlign = 'right';
      ctx.fillStyle = ac;
      ctx.fillText(`NIGHT EDITION #${String(opts.postIndex).padStart(2, '0')}`, width - pad, height - pad);
      ctx.textAlign = 'left';
    }
    return;
  }

  if (opts.style === 'step_roadmap') {
    fillSmoothAmbientBackground(ctx, theme, width, height, 0.16);
    const numFg = isLightHex(theme.accent) ? '#0c0c0e' : '#ffffff';

    let y = pad;
    ctx.font = `700 26px ${JP_FONT}`;
    ctx.textBaseline = 'top';
    ctx.fillStyle = theme.accent;
    ctx.fillText(`🧭 ${opts.topic || 'GUIDE'} · 実践ステップ図解`, pad, y);
    y += 46;

    const maxW = width - pad * 2;
    ctx.font = `800 50px ${JP_FONT}`;
    ctx.fillStyle = theme.textMain;
    const hLines = wrapJapaneseLines(ctx, opts.headline, maxW, 2);
    for (const hl of hLines) {
      ctx.fillText(hl, pad, y);
      y += 66;
    }
    y += 14;
    ctx.strokeStyle = theme.border;
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(pad, y);
    ctx.lineTo(width - pad, y);
    ctx.stroke();
    y += 38;

    const sItems = listItems.slice(0, 3);
    const availH = height - pad - 68 - y;
    const gap = 26;
    const cardH = Math.min(192, Math.max(112, Math.floor((availH - gap * (sItems.length - 1)) / Math.max(1, sItems.length))));
    const nodeX = pad + 34;

    if (sItems.length > 1) {
      const topCy = y + cardH / 2;
      const botCy = y + (sItems.length - 1) * (cardH + gap) + cardH / 2;
      ctx.strokeStyle = theme.accent;
      ctx.lineWidth = 5;
      ctx.beginPath();
      ctx.moveTo(nodeX, topCy);
      ctx.lineTo(nodeX, botCy);
      ctx.stroke();
    }

    sItems.forEach((itTxt, idx) => {
      const cy = y + idx * (cardH + gap);
      const midY = cy + cardH / 2;

      ctx.beginPath();
      ctx.arc(nodeX, midY, 28, 0, Math.PI * 2);
      ctx.fillStyle = theme.accent;
      ctx.fill();
      ctx.strokeStyle = theme.bgPrimary;
      ctx.lineWidth = 4;
      ctx.stroke();

      ctx.font = `800 26px ${JP_FONT}`;
      ctx.fillStyle = numFg;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(String(idx + 1), nodeX, midY + 1);
      ctx.textAlign = 'left';

      const cx1 = pad + 88;
      drawRoundedRect(ctx, cx1, cy, width - pad - cx1, cardH, 16);
      ctx.fillStyle = theme.bgSecondary;
      ctx.fill();
      ctx.strokeStyle = theme.border;
      ctx.lineWidth = 2;
      ctx.stroke();

      ctx.font = `700 22px ${JP_FONT}`;
      ctx.fillStyle = theme.accent;
      ctx.textBaseline = 'top';
      ctx.fillText(`STEP 0${idx + 1}`, cx1 + 26, cy + 18);

      ctx.font = `700 31px ${JP_FONT}`;
      ctx.fillStyle = theme.textMain;
      const ilines = wrapJapaneseLines(ctx, itTxt, width - pad - cx1 - 52, 2);
      ilines.forEach((il, li) => {
        ctx.fillText(il, cx1 + 26, cy + 54 + li * 42);
      });
    });

    if (opts.showWatermark) {
      ctx.font = `600 25px ${JP_FONT}`;
      ctx.fillStyle = theme.textMuted;
      ctx.textBaseline = 'bottom';
      ctx.fillText(`@${opts.accountId || 'threads'}`, pad, height - pad);
      ctx.textAlign = 'right';
      ctx.fillStyle = theme.accent;
      ctx.fillText('3-STEP ROADMAP', width - pad, height - pad);
      ctx.textAlign = 'left';
    }
    return;
  }

  if (opts.style === 'qa_card') {
    fillSmoothAmbientBackground(ctx, theme, width, height, 0.16);
    let y = pad;
    ctx.font = `700 26px ${JP_FONT}`;
    ctx.textBaseline = 'top';
    ctx.fillStyle = theme.accent;
    ctx.fillText(`❓ ${opts.topic || 'GUIDE'} · Q&A よくいただく質問`, pad, y);
    y += 48;

    const maxW = width - pad * 2;
    ctx.font = `700 34px ${JP_FONT}`;
    const qLines = wrapJapaneseLines(ctx, painText, maxW - 130, 3);
    const qBoxH = Math.max(136, 48 + qLines.length * 46);
    drawRoundedRect(ctx, pad, y, maxW, qBoxH, 18);
    ctx.fillStyle = theme.bgSecondary;
    ctx.fill();
    ctx.strokeStyle = theme.accent;
    ctx.lineWidth = 3;
    ctx.stroke();

    ctx.font = `800 54px ${JP_FONT}`;
    ctx.fillStyle = theme.accent;
    ctx.textBaseline = 'middle';
    ctx.fillText('Q.', pad + 30, y + qBoxH / 2);

    ctx.font = `700 34px ${JP_FONT}`;
    ctx.fillStyle = theme.textMain;
    ctx.textBaseline = 'top';
    let qy = y + (qBoxH - qLines.length * 44) / 2;
    for (const ql of qLines) {
      ctx.fillText(ql, pad + 108, qy);
      qy += 44;
    }
    y += qBoxH + 30;

    ctx.font = `700 26px ${JP_FONT}`;
    ctx.fillStyle = theme.accent;
    ctx.fillText('A. 結論・押さえるべきポイント', pad, y);
    y += 42;

    ctx.font = `800 44px ${JP_FONT}`;
    ctx.fillStyle = theme.textMain;
    const aLines = wrapJapaneseLines(ctx, opts.headline, maxW, 2);
    for (const al of aLines) {
      ctx.fillText(al, pad, y);
      y += 58;
    }
    y += 22;

    const qItems = listItems.slice(0, 3);
    const availH = height - pad - 68 - y;
    const gap = 18;
    const cardH = Math.min(132, Math.max(82, Math.floor((availH - gap * (qItems.length - 1)) / Math.max(1, qItems.length))));
    qItems.forEach((itTxt, idx) => {
      const cy = y + idx * (cardH + gap);
      drawRoundedRect(ctx, pad, cy, maxW, cardH, 14);
      ctx.fillStyle = theme.bgSecondary;
      ctx.fill();
      ctx.strokeStyle = theme.border;
      ctx.lineWidth = 2;
      ctx.stroke();

      drawRoundedRect(ctx, pad + 24, cy + (cardH - 36) / 2, 36, 36, 8);
      ctx.fillStyle = theme.accentSoft;
      ctx.fill();
      ctx.strokeStyle = theme.accent;
      ctx.lineWidth = 2;
      ctx.stroke();

      ctx.font = `700 24px ${JP_FONT}`;
      ctx.fillStyle = theme.accent;
      ctx.textBaseline = 'middle';
      ctx.fillText('✓', pad + 32, cy + cardH / 2);

      ctx.font = `700 30px ${JP_FONT}`;
      ctx.fillStyle = theme.textMain;
      const ilines = wrapJapaneseLines(ctx, itTxt, maxW - 116, 2);
      if (ilines.length === 1) {
        ctx.fillText(ilines[0], pad + 82, cy + cardH / 2);
      } else {
        ctx.fillText(ilines[0], pad + 82, cy + cardH / 2 - 20);
        ctx.fillText(ilines[1], pad + 82, cy + cardH / 2 + 20);
      }
    });

    if (opts.showWatermark) {
      ctx.font = `600 25px ${JP_FONT}`;
      ctx.fillStyle = theme.textMuted;
      ctx.textBaseline = 'bottom';
      ctx.fillText(`@${opts.accountId || 'threads'}`, pad, height - pad);
      ctx.textAlign = 'right';
      ctx.fillStyle = theme.accent;
      ctx.fillText(`Q&A ANSWER #${String(opts.postIndex).padStart(2, '0')}`, width - pad, height - pad);
      ctx.textAlign = 'left';
    }
    return;
  }

  // diagram_card / checklist with smooth ambient gradient
  fillSmoothAmbientBackground(ctx, theme, width, height, 0.14);

  if (opts.backgroundImg && opts.backgroundImg.naturalWidth > 0) {
    ctx.save();
    ctx.globalAlpha = theme.isLight ? 0.08 : 0.14;
    drawImageCover(ctx, opts.backgroundImg, width, height);
    ctx.restore();
  }

  const maxW = width - pad * 2;
  let y = pad;

  ctx.font = `700 28px ${JP_FONT}`;
  ctx.fillStyle = theme.accent;
  ctx.textBaseline = 'top';
  const modeLabel = opts.style === 'checklist' ? `${opts.topic} · 保存用チェックリスト` : `${opts.topic} · ポイント図解`;
  ctx.fillText(modeLabel, pad, y);
  y += 50;

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
      const boxSize = 46;
      const bx = pad + 28;
      const by = cy + (cardH - boxSize) / 2;
      drawRoundedRect(ctx, bx, by, boxSize, boxSize, 10);
      ctx.fillStyle = theme.accentSoft;
      ctx.fill();
      ctx.strokeStyle = theme.accent;
      ctx.lineWidth = 2.5;
      ctx.stroke();

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
      ctx.font = `800 36px "JetBrains Mono", monospace`;
      ctx.fillStyle = theme.accent;
      ctx.textBaseline = 'middle';
      ctx.fillText(String(idx + 1).padStart(2, '0'), pad + 28, cy + cardH / 2);
    }

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
