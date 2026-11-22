import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { AccountInfo, AccountDesignProfile, PostItem, DailyGreetingSlot } from '../types';
import { extractUserIdFromToken } from './spintax';

function normalizeTone(raw?: string): AccountDesignProfile['tone'] | undefined {
  if (!raw) return undefined;
  const v = raw.trim().toLowerCase();
  if (v === 'polite' || v.includes('丁寧') || v.includes('誠実')) return 'polite';
  if (v === 'friendly' || v.includes('フレンド') || v.includes('親近')) return 'friendly';
  if (v === 'expert' || v.includes('プロ') || v.includes('解説')) return 'expert';
  if (v === 'empathetic' || v.includes('共感') || v.includes('エモ')) return 'empathetic';
  return undefined;
}

function normalizeImageStyle(raw?: string): AccountDesignProfile['imageStyle'] | undefined {
  if (!raw) return undefined;
  const v = raw.trim().toLowerCase();
  if (v === 'auto' || v.includes('自動おまかせ') || v.includes('最適化')) return 'auto';
  if (v === 'magazine_clean' || v.includes('雑誌') || v.includes('エディトリアル')) return 'magazine_clean';
  if (v === 'chat_bubble' || v.includes('吹き出し') || v.includes('会話')) return 'chat_bubble';
  if (v === 'neon_night' || v.includes('ネオン')) return 'neon_night';
  if (v === 'step_roadmap' || v.includes('ステップ') || v.includes('ロードマップ')) return 'step_roadmap';
  if (v === 'qa_card' || v.includes('q&a') || v.includes('神回答')) return 'qa_card';
  if (v === 'photo_overlay' || v.includes('写真') || v.includes('文字入れ')) return 'photo_overlay';
  if (v === 'diagram_card' || v.includes('図解')) return 'diagram_card';
  if (v === 'checklist' || v.includes('チェック')) return 'checklist';
  if (v === 'all_mix' || v.includes('全テンプレート')) return 'all_mix';
  if (v === 'mix' || v.includes('ミックス')) return 'mix';
  return undefined;
}

function sanitizeExcelSheetName(name: string): string {
  const cleaned = String(name || '')
    .trim()
    .replace(/[\\/?*[\]:]+/g, '_')
    .replace(/^'+|'+$/g, '')
    .trim();
  return (cleaned || '基本').slice(0, 31);
}

function isDailyGreetingSheet(sname: string, rawRows?: any[][]): boolean {
  const s = String(sname || '').trim();
  const sl = s.toLowerCase();
  if (!s) return false;
  if (rawRows && rawRows.length > 0) {
    const r0 = (rawRows[0] || []).map((c) => String(c ?? ''));
    if (r0.some((c) => c.includes('枠名'))) return true;
    for (let r = 0; r < Math.min(10, rawRows.length); r++) {
      const cells = (rawRows[r] || []).map((c) => String(c ?? '').trim());
      if (cells.some((c) => c === 'ID' || c === 'ステータス' || c === '投稿モード' || c === 'ポスト間隔' || c === 'ツイート間隔')) {
        return false;
      }
    }
  }
  if (['日常ポスト', '日常挨拶', '日常あいさつ', '日常枠', 'あいさつ', '挨拶', 'daily', 'daily_posts', 'dailygreeting'].includes(sl)) {
    return true;
  }
  if (s.includes('日常') && (s.includes('挨拶') || s.includes('あいさつ') || s.includes('枠'))) {
    return true;
  }
  return sl.includes('daily') && sl.includes('post');
}

function isSystemSheetName(sname: string, rawRows?: any[][]): boolean {
  const s = String(sname || '').trim();
  const sl = s.toLowerCase();
  if (!s) return true;
  if (['設定', 'config', 'setting', '基本設定', 'テーマ設定', 'themesetting', 'themeconfig'].some((k) => sl.includes(k))) return true;
  if (isDailyGreetingSheet(s, rawRows)) return true;
  if (s.includes('設計')) return true;
  if (['アカウント情報', 'accountinfo', 'account_info'].some((k) => sl.includes(k))) return true;
  return false;
}

function isLegacyFlatPostSheetName(sname: string): boolean {
  const sl = String(sname || '').trim().toLowerCase();
  return ['投稿リスト', '投稿', 'ポスト', 'post', 'posts', 'tweet', 'tweets', 'postlist'].includes(sl);
}

const FONT_NAME = 'Meiryo UI';

const solidFill = (argb: string): ExcelJS.FillPattern => ({
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb },
  bgColor: { argb },
});

const thinBorder = (argb: string): Partial<ExcelJS.Borders> => ({
  top: { style: 'thin', color: { argb } },
  left: { style: 'thin', color: { argb } },
  bottom: { style: 'thin', color: { argb } },
  right: { style: 'thin', color: { argb } },
});

function estimateRowHeight(text: string, charsPerLine = 32, minH = 26, maxH = 140): number {
  const lines = String(text || '')
    .split('\n')
    .reduce((acc, seg) => acc + Math.max(1, Math.ceil(seg.length / charsPerLine)), 0);
  return Math.max(minH, Math.min(maxH, lines * 18 + 10));
}

function triggerBrowserExcelDownload(buffer: ArrayBuffer, filename: string) {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

async function buildStyledAccountWorkbook(
  account: AccountInfo,
  posts: PostItem[],
  _dailySlots?: DailyGreetingSlot[]
): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Threads AutoPoster';
  wb.created = new Date();

  const metaHdrFill = solidFill('FF1E3A8A');
  const metaValFill = solidFill('FFDBEAFE');
  const metaMemoFill = solidFill('FFEFF6FF');
  const postHdrFill = solidFill('FF0F766E');
  const idColFill = solidFill('FFF1F5F9');
  const schedColFill = solidFill('FFF0F9FF');
  const oddFill = solidFill('FFFFFFFF');
  const evenFill = solidFill('FFF8FAFC');
  const greenFill = solidFill('FFDCFCE7');
  const amberFill = solidFill('FFFEF3C7');
  const redFill = solidFill('FFFEE2E2');

  const metaBorder = thinBorder('FF60A5FA');
  const hdrBorder = thinBorder('FF334155');
  const dataBorder = thinBorder('FFCBD5E1');

  const hdrFont: Partial<ExcelJS.Font> = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  const metaValFont: Partial<ExcelJS.Font> = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FF1E3A8A' } };
  const metaMemoFont: Partial<ExcelJS.Font> = { name: FONT_NAME, size: 9, bold: false, color: { argb: 'FF334155' } };
  const bodyFont: Partial<ExcelJS.Font> = { name: FONT_NAME, size: 10, bold: false, color: { argb: 'FF0F172A' } };
  const idFont: Partial<ExcelJS.Font> = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FF475569' } };
  const greenFont: Partial<ExcelJS.Font> = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FF166534' } };
  const amberFont: Partial<ExcelJS.Font> = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FF92400E' } };
  const redFont: Partial<ExcelJS.Font> = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FF991B1B' } };

  // 1. テーマ別投稿シート群 (Excel下部シートタブでテーマ・実行時間・曜日・ポスト間隔・投稿モード・文章を直接管理。設定シートは廃止)
  const themeBuckets = new Map<string, PostItem[]>();
  posts.forEach((p) => {
    const tName = sanitizeExcelSheetName(p.theme || '基本');
    if (!themeBuckets.has(tName)) {
      themeBuckets.set(tName, []);
    }
    themeBuckets.get(tName)!.push(p);
  });
  if (themeBuckets.size === 0) {
    themeBuckets.set('基本', []);
  }

  const postHeaders = [
    'ID',
    'ポスト文章',
    '投稿時間',
    '曜日',
    'コミュニティまたはトピック',
    '画像ファイルパス',
    'ステータス',
    '投稿完了日時',
    'エラー内容',
  ];

  themeBuckets.forEach((themePosts, sheetTitle) => {
    const ws = wb.addWorksheet(sheetTitle, {
      views: [{ state: 'frozen', ySplit: 4, showGridLines: true }],
    });

    ws.columns = [
      { width: 16 },
      { width: 56 },
      { width: 24 },
      { width: 20 },
      { width: 26 },
      { width: 26 },
      { width: 16 },
      { width: 22 },
      { width: 30 },
    ];

    const firstSlot = themePosts.find((p) => p.themeTimeSlot || p.timeSlot)?.themeTimeSlot || themePosts.find((p) => p.timeSlot)?.timeSlot || account.activeHours || '0:00-0:00';
    const firstWeekdays = themePosts.find((p) => p.themeWeekdays || p.weekdays)?.themeWeekdays || themePosts.find((p) => p.weekdays)?.weekdays || '毎日';
    const intervalStr =
      themePosts.find((p) => p.intervalStr)?.intervalStr ||
      `${account.minIntervalMinutes}〜${account.maxIntervalMinutes}分`;
    const firstMode = themePosts.find((p) => p.mode !== undefined)?.mode ?? account.mode ?? 1;
    const modeStr = firstMode === 0 ? '0: 上から順' : '1: ランダム';

    // Row 1: 上部テーマ設定ヘッダー (紺)
    const metaHeaders = ['実行時間', '曜日', 'ポスト間隔', '投稿モード', '有効/無効', 'メモ'];
    const r1 = ws.getRow(1);
    r1.height = 26;
    metaHeaders.forEach((lbl, idx) => {
      const cell = r1.getCell(idx + 1);
      cell.value = lbl;
      cell.fill = metaHdrFill;
      cell.font = hdrFont;
      cell.border = hdrBorder;
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      if (lbl === '投稿モード') {
        cell.note = '【投稿モード】数字だけでもOK\n0 (または 0: 上から順) : シートの上から順番に投稿\n1 (または 1: ランダム) : 未投稿からランダムに投稿';
      }
    });

    // Row 2: 上部テーマ設定値 (アイスブルー + 有効緑)
    const metaValues = [
      firstSlot,
      firstWeekdays,
      intervalStr,
      modeStr,
      '有効',
      '投稿モード: 0=上から順 / 1=ランダム (数字だけでもOK) ｜ 下のシートを追加・複製するとテーマ別に管理できます',
    ];
    const r2 = ws.getRow(2);
    r2.height = 36;
    metaValues.forEach((val, idx) => {
      const cell = r2.getCell(idx + 1);
      cell.value = val;
      cell.border = metaBorder;
      if (idx === 4) {
        cell.fill = greenFill;
        cell.font = greenFont;
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      } else if (idx === 5) {
        cell.fill = metaMemoFill;
        cell.font = metaMemoFont;
        cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      } else {
        cell.fill = metaValFill;
        cell.font = metaValFont;
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      }
    });

    // Row 3: 見やすい空白セパレーター
    ws.getRow(3).height = 12;

    // Row 4: 投稿一覧テーブルヘッダー (ディープティール)
    const r4 = ws.getRow(4);
    r4.height = 28;
    postHeaders.forEach((lbl, idx) => {
      const cell = r4.getCell(idx + 1);
      cell.value = lbl;
      cell.fill = postHdrFill;
      cell.font = hdrFont;
      cell.border = hdrBorder;
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      if (lbl === '投稿時間') {
        cell.note = '【各ポストの投稿時間指定（上の「実行時間」より優先）】\n例: 7～10 / 9時～12時 / 18:00-22:00 / 終日\n空欄の場合はシート上部の「実行時間」設定に従います。\n（※上の実行時間が「12:00-15:00」でも、ここが「18～21」なら18～21時に投稿されます）';
      }
      if (lbl === '曜日') {
        cell.note = '【各ポストの曜日指定（上の「曜日」より優先）】\n例: 月 / 月～木 / 月～金 / 月,木,金\n空欄の場合はシート上部の「曜日」設定に従います。\n（※上の曜日が「火」でも、ここが「水」なら水曜日に投稿されます）';
      }
    });

    // Row 5〜バッファ行 (120行以上まで色・枠線を事前配置して削除時も消えない)
    const endRow = Math.max(4 + themePosts.length + 60, 124);
    for (let rIdx = 5; rIdx <= endRow; rIdx++) {
      const pIdx = rIdx - 5;
      const p = pIdx < themePosts.length ? themePosts[pIdx] : null;
      const row = ws.getRow(rIdx);
      const isEven = (rIdx - 4) % 2 === 0;
      const baseFill = isEven ? evenFill : oddFill;

      const vals: (string | number)[] = p
        ? [
            p.id ?? pIdx + 1,
            p.text,
            p.timeSlot || '',
            p.weekdays || '',
            p.topic || '',
            p.imagePath || '',
            p.status,
            p.completedAt || '',
            p.errorMessage || '',
          ]
        : ['', '', '', '', '', '', '', '', ''];

      for (let c = 1; c <= 9; c++) {
        const cell = row.getCell(c);
        cell.value = vals[c - 1];
        cell.border = dataBorder;
        if (c === 3 || c === 4) {
          cell.numFmt = '@';
        }

        if (c === 1 || c === 3 || c === 4 || c === 7) {
          cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        } else if (c === 2 || c === 9) {
          cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        } else {
          cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        }

        if (c === 7 && vals[6]) {
          const st = String(vals[6]);
          if (st.includes('完了')) {
            cell.fill = greenFill;
            cell.font = greenFont;
          } else if (st.includes('エラー')) {
            cell.fill = redFill;
            cell.font = redFont;
          } else {
            cell.fill = amberFill;
            cell.font = amberFont;
          }
        } else if (c === 1) {
          cell.fill = idColFill;
          cell.font = idFont;
        } else if (c === 3 || c === 4) {
          cell.fill = schedColFill;
          cell.font = bodyFont;
        } else {
          cell.fill = baseFill;
          cell.font = bodyFont;
        }
      }

      row.height = p ? estimateRowHeight(p.text, 34, 34, 140) : 24;
    }

    // Excelネイティブ条件付き書式（セル内容を変更・削除しても色が追従）
    const cfEndRow = Math.max(endRow + 200, 500);
    ws.addConditionalFormatting({
      ref: `G5:G${cfEndRow}`,
      rules: [
        { type: 'cellIs', operator: 'equal', priority: 1, formulae: ['"投稿完了"'], style: { fill: greenFill, font: greenFont } },
        { type: 'cellIs', operator: 'equal', priority: 2, formulae: ['"未投稿"'], style: { fill: amberFill, font: amberFont } },
        { type: 'cellIs', operator: 'equal', priority: 3, formulae: ['"エラー"'], style: { fill: redFill, font: redFont } },
      ],
    });
    ws.addConditionalFormatting({
      ref: 'E2',
      rules: [
        { type: 'cellIs', operator: 'equal', priority: 1, formulae: ['"有効"'], style: { fill: greenFill, font: greenFont } },
        { type: 'cellIs', operator: 'equal', priority: 2, formulae: ['"無効"'], style: { fill: redFill, font: redFont } },
      ],
    });
  });

  return wb;
}

/**
 * Exports current account and post data into styled .xlsx format
 */
export async function exportAccountWorkbook(
  account: AccountInfo,
  posts: PostItem[],
  dailySlots?: DailyGreetingSlot[]
): Promise<void> {
  const wb = await buildStyledAccountWorkbook(account, posts, dailySlots);
  const buffer = await wb.xlsx.writeBuffer();
  triggerBrowserExcelDownload(buffer as ArrayBuffer, `${account.id}.xlsx`);
}

/**
 * Builds styled .xlsx bytes (Uint8Array) for bundling inside a ZIP archive
 */
export async function buildAccountWorkbookBytes(
  account: AccountInfo,
  posts: PostItem[],
  dailySlots?: DailyGreetingSlot[]
): Promise<Uint8Array> {
  const wb = await buildStyledAccountWorkbook(account, posts, dailySlots);
  const buffer = await wb.xlsx.writeBuffer();
  return new Uint8Array(buffer as ArrayBuffer);
}

/**
 * Exports all accounts list to styled アカウント情報.xlsx
 */
export async function exportAccountsConfigWorkbook(accounts: AccountInfo[]): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Threads AutoPoster';
  const ws = wb.addWorksheet('アカウント設定', {
    views: [{ state: 'frozen', ySplit: 1, showGridLines: true }],
  });

  ws.columns = [
    { width: 20 },
    { width: 22 },
    { width: 45 },
    { width: 14 },
    { width: 40 },
  ];

  const hdrFill = solidFill('FF1E293B');
  const oddFill = solidFill('FFFFFFFF');
  const evenFill = solidFill('FFF8FAFC');
  const greenFill = solidFill('FFDCFCE7');
  const redFill = solidFill('FFFEE2E2');
  const hdrBorder = thinBorder('FF334155');
  const dataBorder = thinBorder('FFCBD5E1');
  const hdrFont: Partial<ExcelJS.Font> = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  const bodyFont: Partial<ExcelJS.Font> = { name: FONT_NAME, size: 10, bold: false, color: { argb: 'FF0F172A' } };
  const greenFont: Partial<ExcelJS.Font> = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FF166534' } };
  const redFont: Partial<ExcelJS.Font> = { name: FONT_NAME, size: 10, bold: true, color: { argb: 'FF991B1B' } };

  const headers = ['アカウントID', 'パスワード', 'sessionid', '有効/無効', '備考'];
  const r1 = ws.getRow(1);
  r1.height = 28;
  headers.forEach((lbl, idx) => {
    const cell = r1.getCell(idx + 1);
    cell.value = lbl;
    cell.fill = hdrFill;
    cell.font = hdrFont;
    cell.border = hdrBorder;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  const endRow = Math.max(1 + accounts.length + 30, 80);
  for (let rIdx = 2; rIdx <= endRow; rIdx++) {
    const aIdx = rIdx - 2;
    const acc = aIdx < accounts.length ? accounts[aIdx] : null;
    const row = ws.getRow(rIdx);
    const baseFill = rIdx % 2 === 0 ? evenFill : oddFill;
    const vals = acc
      ? [
          acc.id,
          acc.password || '',
          acc.authToken || '',
          acc.enabled !== false ? '有効' : '無効',
          acc.notes || '',
        ]
      : ['', '', '', '', ''];

    for (let c = 1; c <= 5; c++) {
      const cell = row.getCell(c);
      cell.value = vals[c - 1];
      cell.border = dataBorder;
      cell.alignment = { horizontal: c === 4 ? 'center' : 'left', vertical: 'middle', wrapText: true };
      if (c === 4 && vals[3]) {
        const isOff = vals[3] === '無効';
        cell.fill = isOff ? redFill : greenFill;
        cell.font = isOff ? redFont : greenFont;
      } else {
        cell.fill = baseFill;
        cell.font = bodyFont;
      }
    }
    row.height = acc ? 26 : 22;
  }

  ws.addConditionalFormatting({
    ref: `D2:D${Math.max(endRow + 100, 300)}`,
    rules: [
      { type: 'cellIs', operator: 'equal', priority: 1, formulae: ['"有効"'], style: { fill: greenFill, font: greenFont } },
      { type: 'cellIs', operator: 'equal', priority: 2, formulae: ['"無効"'], style: { fill: redFill, font: redFont } },
    ],
  });

  const buffer = await wb.xlsx.writeBuffer();
  triggerBrowserExcelDownload(buffer as ArrayBuffer, 'アカウント情報.xlsx');
}

/**
 * Parses uploaded Excel workbook and extracts accounts, posts, or settings
 */
export async function parseExcelFile(
  file: File
): Promise<{
  accounts?: Partial<AccountInfo>[];
  posts?: PostItem[];
  dailySlots?: DailyGreetingSlot[];
  settings?: Partial<AccountInfo>;
  fileName: string;
}> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const result: {
          accounts?: Partial<AccountInfo>[];
          posts?: PostItem[];
          dailySlots?: DailyGreetingSlot[];
          settings?: Partial<AccountInfo>;
          fileName: string;
        } = { fileName: file.name };

        // Check if this is an account config file
        const accountSheet = wb.SheetNames.find(
          (name) => name.includes('アカウント') || name.toLowerCase().includes('account')
        ) || (wb.SheetNames.length === 1 && file.name.includes('アカウント') ? wb.SheetNames[0] : null);

        if (accountSheet) {
          const ws = wb.Sheets[accountSheet];
          const rawRows = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1 });
          if (rawRows.length > 0) {
            let headerIdx = 0;
            for (let i = 0; i < Math.min(10, rawRows.length); i++) {
              const rowStr = (rawRows[i] || []).join(' ');
              if (rowStr.includes('アカウントID') || rowStr.includes('アカウント名') || rowStr.includes('パスワード') || rowStr.includes('sessionid')) {
                headerIdx = i;
                break;
              }
            }
            const headers = (rawRows[headerIdx] || []).map((h) => String(h || '').trim());
            const accIdCol = headers.findIndex((h) => h.includes('アカウントID') || h.includes('アカウント名') || h === 'アカウント' || h === 'ID');
            const passCol = headers.findIndex((h) => h.includes('パスワード') || h.toLowerCase().includes('password') || h.toLowerCase() === 'pw');
            const sessionCol = headers.findIndex((h) => h.toLowerCase().includes('sessionid') || h.includes('auth_token') || h.includes('トークン'));
            const enabledCol = headers.findIndex((h) => h.includes('有効') || h.includes('ステータス') || h.includes('状態'));
            const notesCol = headers.findIndex((h) => h.includes('備考') || h.includes('note'));

            const getCell = (row: any[], idx: number) =>
              idx !== -1 && row[idx] !== undefined && row[idx] !== null ? String(row[idx]).trim() : '';

            const parsedAccounts: Partial<AccountInfo>[] = [];
            for (let r = headerIdx + 1; r < rawRows.length; r++) {
              const row = rawRows[r];
              if (!row || !row[accIdCol]) continue;
              const accId = String(row[accIdCol]).trim().replace(/^@/, '');
              if (!accId) continue;
              const passVal = getCell(row, passCol);
              const tokenVal = getCell(row, sessionCol);
              const statusStr = getCell(row, enabledCol) || '有効';
              const isEnabled = !['無効', '停止', 'off', 'OFF', '0', 'false', '停止 (要電話番号認証)'].includes(statusStr);

              parsedAccounts.push({
                id: accId,
                password: passVal,
                authToken: tokenVal,
                userId: extractUserIdFromToken(tokenVal) || accId,
                // プロファイル名はアカウントIDで自動生成 (chrome_profiles/<アカウントID>)
                profileDir: `chrome_profiles/${accId}`,
                enabled: isEnabled,
                notes: getCell(row, notesCol),
              });
            }
            if (parsedAccounts.length > 0) {
              result.accounts = parsedAccounts;
            }
          }
        }

        // Also check for standalone アカウント設計 sheet inside <account>.xlsx
        const designSheetName = wb.SheetNames.find((name) => name.includes('設計'));
        if (designSheetName) {
          const ws = wb.Sheets[designSheetName];
          const rawRows = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1 });
          const parsedDesign: AccountDesignProfile = {};
          for (const row of rawRows) {
            if (!row || row.length < 2) continue;
            const k = String(row[0] || '').trim();
            const v = String(row[1] || '').trim();
            if (!k || !v) continue;
            if (k.includes('ジャンル')) parsedDesign.genre = v;
            else if (k.includes('ターゲット')) parsedDesign.target = v;
            else if (k.includes('口調')) parsedDesign.tone = normalizeTone(v);
            else if (k.includes('悩み')) parsedDesign.painPoint = v;
            else if (k.includes('強み') || k.includes('ポイント')) parsedDesign.benefitPoints = v;
            else if (k.includes('誘導文') || k.includes('CTA')) parsedDesign.cta = v;
            else if (k.includes('URL')) parsedDesign.url = v;
            else if (k.includes('トピック')) parsedDesign.topic = v;
            else if (k.includes('スタイル')) parsedDesign.imageStyle = normalizeImageStyle(v);
          }
          if (Object.keys(parsedDesign).length > 0) {
            result.settings = { ...(result.settings || {}), design: parsedDesign };
          }
        }

        // Check for テーマ別投稿シート群 (および互換用の旧「投稿リスト」シート)
        const allParsedPosts: PostItem[] = [];
        for (const sName of wb.SheetNames) {
          const ws = wb.Sheets[sName];
          const rawRows = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1 });
          if (isSystemSheetName(sName, rawRows)) continue;
          if (!rawRows || rawRows.length < 2) continue;

          let hdrIdx = -1;
          for (let r = 0; r < Math.min(15, rawRows.length); r++) {
            const cellsNorm = (rawRows[r] || []).map((c) =>
              String(c ?? '').trim().toLowerCase().replace(/[\s　_\-:=]+/g, '')
            );
            if (cellsNorm.some((c) => ['ポスト文章', '投稿本文', 'ツイート文章', 'posttext', '投稿文章', 'ポスト本文', 'ポスト内容', '投稿内容', '本文'].some((k) => c.includes(k)))) {
              hdrIdx = r;
              break;
            }
          }
          if (hdrIdx === -1) continue;

          // シート上部のテーマメタ設定（実行時間・曜日・ポスト間隔・投稿モードなど）を読み取り
          let sheetTimeSlot = '';
          let sheetWeekdays = '';
          let sheetIntervalStr = '';
          let sheetMode: 0 | 1 | undefined = undefined;
          const metaRows = rawRows.slice(0, hdrIdx);
          for (let mr = 0; mr < metaRows.length; mr++) {
            const mRow = metaRows[mr] || [];
            for (let mc = 0; mc < mRow.length; mc++) {
              const lbl = String(mRow[mc] ?? '').trim();
              if (!lbl || lbl.length > 20) continue;
              const below = mr + 1 < metaRows.length && metaRows[mr + 1] ? String(metaRows[mr + 1][mc] ?? '').trim() : '';
              const right = mc + 1 < mRow.length ? String(mRow[mc + 1] ?? '').trim() : '';
              const cand = below || right;
              if (!cand || cand.startsWith('(') || cand.startsWith('（') || cand.startsWith('※')) continue;

              if (!sheetTimeSlot && (lbl.includes('実行時間') || lbl.includes('時間帯')) && !lbl.includes('完了') && !lbl.includes('最終')) {
                sheetTimeSlot = cand;
              } else if (!sheetWeekdays && lbl.includes('曜日')) {
                sheetWeekdays = cand;
              } else if (!sheetIntervalStr && (lbl.includes('ツイート間隔') || lbl.includes('ポスト間隔') || lbl.includes('投稿間隔') || lbl === '間隔')) {
                sheetIntervalStr = cand;
              } else if (sheetMode === undefined && (lbl.includes('投稿モード') || lbl.includes('モード') || lbl.toLowerCase() === 'mode')) {
                const cl = cand.normalize('NFKC').trim().toLowerCase();
                if (cl.startsWith('0') || cand.includes('順番') || (cand.includes('上') && cand.includes('順')) || cl === 'sequential') {
                  sheetMode = 0;
                } else if (cl.startsWith('1') || cand.includes('ランダム') || cl === 'random') {
                  sheetMode = 1;
                }
              }
            }
          }

          const headers = (rawRows[hdrIdx] || []).map((h) => String(h || '').trim());
          const idCol = headers.findIndex((h) => h.toUpperCase() === 'ID');
          const themeCol = headers.findIndex((h) => h.includes('テーマ') && !h.includes('配色'));
          const textCol = headers.findIndex((h) => h.includes('ポスト文章') || h.includes('投稿本文') || h.includes('ツイート文章') || h.includes('ポスト内容') || h.includes('投稿内容') || h.includes('本文'));
          const topicCol = headers.findIndex((h) => (h.includes('トピック') || h.includes('コミュニティ')) && !h.includes('適用'));
          const imgCol = headers.findIndex((h) => h.includes('画像') || h.toLowerCase().includes('image'));
          const slotCol = headers.findIndex((h) => (h.includes('投稿時間') || h.includes('投稿時刻') || h.includes('実行時間') || h.includes('時間帯')) && !h.includes('完了') && !h.includes('最終'));
          const wdCol = headers.findIndex((h) => h.includes('曜日') || h.toLowerCase().includes('weekday'));
          const statusCol = headers.findIndex((h) => h.includes('ステータス') || h.toLowerCase().includes('status'));
          const finishCol = headers.findIndex((h) => h.includes('完了日時') || h.includes('投稿完了日時'));
          const errCol = headers.findIndex((h) => h.includes('エラー'));

          for (let r = hdrIdx + 1; r < rawRows.length; r++) {
            const row = rawRows[r];
            if (!row) continue;
            const text = textCol !== -1 && row[textCol] ? String(row[textCol]).trim() : '';
            if (!text) continue;

            const statusRaw = statusCol !== -1 && row[statusCol] ? String(row[statusCol]).trim() : '未投稿';
            let status: PostItem['status'] = '未投稿';
            if (statusRaw.includes('完了')) status = '投稿完了';
            else if (statusRaw.includes('中')) status = '投稿中';
            else if (statusRaw.includes('エラー')) status = 'エラー';

            const rowTheme = themeCol !== -1 && row[themeCol] ? String(row[themeCol]).trim() : '';
            const resolvedTheme = rowTheme || (isLegacyFlatPostSheetName(sName) ? '基本' : sName);

            allParsedPosts.push({
              id: idCol !== -1 && row[idCol] ? Number(row[idCol]) || (r - hdrIdx) : (r - hdrIdx),
              text,
              theme: resolvedTheme,
              themeTimeSlot: sheetTimeSlot || undefined,
              themeWeekdays: sheetWeekdays || undefined,
              timeSlot: slotCol !== -1 && row[slotCol] ? String(row[slotCol]).trim() : undefined,
              weekdays: wdCol !== -1 && row[wdCol] ? String(row[wdCol]).trim() : undefined,
              intervalStr: sheetIntervalStr || undefined,
              mode: sheetMode,
              topic: topicCol !== -1 && row[topicCol] ? String(row[topicCol]).trim() : '',
              imagePath: imgCol !== -1 && row[imgCol] ? String(row[imgCol]).trim() : '',
              status,
              completedAt: finishCol !== -1 && row[finishCol] ? String(row[finishCol]).trim() : undefined,
              errorMessage: errCol !== -1 && row[errCol] ? String(row[errCol]).trim() : undefined,
            });
          }
        }
        if (allParsedPosts.length > 0) {
          result.posts = allParsedPosts;
        }

        // Check for 日常ポスト
        const dailySheetName = wb.SheetNames.find((name) => {
          const ws = wb.Sheets[name];
          const rawRows = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1 });
          return isDailyGreetingSheet(name, rawRows);
        });
        if (dailySheetName) {
          const ws = wb.Sheets[dailySheetName];
          const rawRows = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1 });
          if (rawRows.length > 1) {
            const slots: DailyGreetingSlot[] = [];
            for (let r = 1; r < rawRows.length; r++) {
              const row = rawRows[r];
              if (!row || !row[2]) continue;
              slots.push({
                id: `imported_daily_${r}`,
                slotName: String(row[0] || '枠名').trim(),
                timeRange: String(row[1] || '0:00-0:00').trim(),
                text: String(row[2]).trim(),
                enabled: String(row[3] || '有効').includes('有効'),
                lastPostedDate: row[4] ? String(row[4]).trim() : undefined,
                lastPostedTime: row[5] ? String(row[5]).trim() : undefined,
              });
            }
            if (slots.length > 0) {
              result.dailySlots = slots;
            }
          }
        }

        resolve(result);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}
