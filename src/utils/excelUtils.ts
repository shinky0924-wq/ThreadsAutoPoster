import * as XLSX from 'xlsx';
import { AccountInfo, PostItem, DailyGreetingSlot } from '../types';
import { extractUserIdFromToken } from './spintax';

/**
 * Exports current account and post data into authentic .xlsx format
 */
export function exportAccountWorkbook(
  account: AccountInfo,
  posts: PostItem[],
  dailySlots: DailyGreetingSlot[]
) {
  const wb = XLSX.utils.book_new();

  // 1. 設定 Sheet
  const settingRows: any[][] = [
    ['ツイート間隔'],
    ['※指定の範囲内から毎回ランダム'],
    [account.minIntervalMinutes, '分', '〜', account.maxIntervalMinutes, '分'],
    ['ツイート実行時間'],
    ['(時:分〜時:分)'],
    [account.activeHours.split('-')[0] || '0:00', '〜', account.activeHours.split('-')[1] || '0:00'],
    ['モード'],
    ['(0:上から順, 1:ランダム)'],
    [account.mode],
    ['Chromeプロファイル保存パス'],
    ['(空欄の場合は chrome_profiles/<アカウントID> が自動使用されます)'],
    [account.profileDir || `chrome_profiles/${account.id}`],
    ['コミュニティまたはトピック', account.topics.join(', ')],
    ['(カンマ区切りで「信太山新地, 大阪求人, 高収入バイト」と書くと投稿ごとに自動ローテーションA/Bテスト)'],
    [account.topics[0] || '信太山新地'],
  ];
  const wsSetting = XLSX.utils.aoa_to_sheet(settingRows);
  XLSX.utils.book_append_sheet(wb, wsSetting, '設定');

  // 2. 投稿リスト Sheet
  const postHeaders = ['ID', 'ポスト文章', 'コミュニティまたはトピック', '画像ファイルパス', 'ステータス', '投稿完了日時', 'エラー内容'];
  const postRows: any[][] = [postHeaders];
  posts.forEach((p) => {
    postRows.push([
      p.id,
      p.text,
      p.topic || '',
      p.imagePath || '',
      p.status,
      p.completedAt || '',
      p.errorMessage || '',
    ]);
  });
  const wsPosts = XLSX.utils.aoa_to_sheet(postRows);
  XLSX.utils.book_append_sheet(wb, wsPosts, '投稿リスト');

  // 3. 日常ポスト Sheet
  const dailyHeaders = ['枠名', '時間帯', 'ポスト文章', '有効無効', '最終投稿日', '最終投稿日時'];
  const dailyRows: any[][] = [dailyHeaders];
  dailySlots.forEach((slot) => {
    dailyRows.push([
      slot.slotName,
      slot.timeRange,
      slot.text,
      slot.enabled ? '有効' : '無効',
      slot.lastPostedDate || '',
      slot.lastPostedTime || '',
    ]);
  });
  const wsDaily = XLSX.utils.aoa_to_sheet(dailyRows);
  XLSX.utils.book_append_sheet(wb, wsDaily, '日常ポスト');

  XLSX.writeFile(wb, `${account.id}.xlsx`);
}

/**
 * Exports all accounts list to アカウント情報.xlsx
 */
export function exportAccountsConfigWorkbook(accounts: AccountInfo[]) {
  const wb = XLSX.utils.book_new();
  const headers = ['アカウントID', 'auth_token', '備考'];
  const rows: any[][] = [headers];

  accounts.forEach((acc) => {
    rows.push([acc.id, acc.authToken, acc.notes || '']);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'アカウント設定');
  XLSX.writeFile(wb, 'アカウント情報.xlsx');
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
              if (rowStr.includes('アカウントID') || rowStr.includes('アカウント名') || rowStr.includes('auth_token')) {
                headerIdx = i;
                break;
              }
            }
            const headers = (rawRows[headerIdx] || []).map((h) => String(h || '').trim());
            const accIdCol = headers.findIndex((h) => h.includes('アカウントID') || h.includes('アカウント名') || h === 'アカウント');
            const tokenCol = headers.findIndex((h) => h.includes('auth_token') || h.includes('sessionid'));
            const notesCol = headers.findIndex((h) => h.includes('備考') || h.includes('note'));

            const parsedAccounts: Partial<AccountInfo>[] = [];
            for (let r = headerIdx + 1; r < rawRows.length; r++) {
              const row = rawRows[r];
              if (!row || !row[accIdCol]) continue;
              const accId = String(row[accIdCol]).trim();
              if (!accId) continue;
              const token = tokenCol !== -1 && row[tokenCol] ? String(row[tokenCol]).trim() : '';
              parsedAccounts.push({
                id: accId,
                authToken: token,
                userId: extractUserIdFromToken(token),
                notes: notesCol !== -1 && row[notesCol] ? String(row[notesCol]).trim() : '',
              });
            }
            if (parsedAccounts.length > 0) {
              result.accounts = parsedAccounts;
            }
          }
        }

        // Check for 投稿リスト
        const postSheetName = wb.SheetNames.find(
          (name) => name.includes('投稿') || name.includes('ポスト') || name.includes('post')
        );
        if (postSheetName) {
          const ws = wb.Sheets[postSheetName];
          const rawRows = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1 });
          if (rawRows.length > 1) {
            const headers = (rawRows[0] || []).map((h) => String(h || '').trim());
            const idCol = headers.findIndex((h) => h.toUpperCase() === 'ID');
            const textCol = headers.findIndex((h) => h.includes('ポスト文章') || h.includes('投稿本文') || h.includes('ツイート文章') || h.includes('本文'));
            const topicCol = headers.findIndex((h) => h.includes('トピック') || h.includes('コミュニティ'));
            const imgCol = headers.findIndex((h) => h.includes('画像') || h.includes('image'));
            const statusCol = headers.findIndex((h) => h.includes('ステータス') || h.includes('status'));
            const finishCol = headers.findIndex((h) => h.includes('完了日時') || h.includes('投稿完了日時'));
            const errCol = headers.findIndex((h) => h.includes('エラー'));

            const posts: PostItem[] = [];
            for (let r = 1; r < rawRows.length; r++) {
              const row = rawRows[r];
              if (!row) continue;
              const text = textCol !== -1 && row[textCol] ? String(row[textCol]).trim() : '';
              if (!text) continue;

              const statusRaw = statusCol !== -1 && row[statusCol] ? String(row[statusCol]).trim() : '未投稿';
              let status: PostItem['status'] = '未投稿';
              if (statusRaw.includes('完了')) status = '投稿完了';
              else if (statusRaw.includes('中')) status = '投稿中';
              else if (statusRaw.includes('エラー')) status = 'エラー';

              posts.push({
                id: idCol !== -1 && row[idCol] ? Number(row[idCol]) || r : r,
                text,
                topic: topicCol !== -1 && row[topicCol] ? String(row[topicCol]).trim() : '',
                imagePath: imgCol !== -1 && row[imgCol] ? String(row[imgCol]).trim() : '',
                status,
                completedAt: finishCol !== -1 && row[finishCol] ? String(row[finishCol]).trim() : undefined,
                errorMessage: errCol !== -1 && row[errCol] ? String(row[errCol]).trim() : undefined,
              });
            }
            if (posts.length > 0) {
              result.posts = posts;
            }
          }
        }

        // Check for 日常ポスト
        const dailySheetName = wb.SheetNames.find(
          (name) => name.includes('日常') || name.includes('挨拶') || name.includes('daily')
        );
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
