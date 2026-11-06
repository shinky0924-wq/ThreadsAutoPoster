/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from 'react';
import { AccountInfo, PostItem, DailyGreetingSlot, SystemSettings, LogMessage } from './types';
import { initialAccounts, initialPosts, initialDailySlots, initialSettings } from './data/initialData';
import { Header } from './components/Header';
import { AccountSelector } from './components/AccountSelector';
import { FeatureToggles } from './components/FeatureToggles';
import { PostsTable } from './components/PostsTable';
import { ThreadsPreview } from './components/ThreadsPreview';
import { DailyGreetings } from './components/DailyGreetings';
import { ConsoleLogs } from './components/ConsoleLogs';
import { SessionModal } from './components/SessionModal';
import { exportAccountWorkbook, exportAccountsConfigWorkbook, parseExcelFile } from './utils/excelUtils';
import { parseSpintax, resolveImagePath, extractUserIdFromToken } from './utils/spintax';
import { DownloadCenter } from './components/DownloadCenter';
import { RAW_THREADS_GUI_APP_PY } from './data/rawScript';
import { RAW_PROFILE_LAUNCHER_PY } from './data/rawProfileLauncher';
import { 
  FileSpreadsheet, 
  Smartphone, 
  Sun, 
  Terminal, 
  Users, 
  Download, 
  FileText,
  AlertCircle,
  Package,
  Rocket
} from 'lucide-react';

export default function App() {
  // Master state
  const [accounts, setAccounts] = useState<AccountInfo[]>(() => {
    const saved = localStorage.getItem('threads_accounts');
    return saved ? JSON.parse(saved) : initialAccounts;
  });

  const [selectedAccountId, setSelectedAccountId] = useState<string>(() => {
    return initialAccounts[0]?.id || 'sll15802';
  });

  // Posts partitioned per account id
  const [accountPosts, setAccountPosts] = useState<Record<string, PostItem[]>>(() => {
    const saved = localStorage.getItem('threads_account_posts');
    if (saved) return JSON.parse(saved);
    const initialMap: Record<string, PostItem[]> = {};
    initialAccounts.forEach((acc) => {
      initialMap[acc.id] = JSON.parse(JSON.stringify(initialPosts));
    });
    return initialMap;
  });

  const [dailySlots, setDailySlots] = useState<DailyGreetingSlot[]>(() => {
    const saved = localStorage.getItem('threads_daily_slots');
    return saved ? JSON.parse(saved) : initialDailySlots;
  });

  const [settings, setSettings] = useState<SystemSettings>(() => {
    const saved = localStorage.getItem('threads_settings');
    return saved ? JSON.parse(saved) : initialSettings;
  });

  const [logs, setLogs] = useState<LogMessage[]>([
    {
      id: 'init-1',
      timestamp: new Date().toLocaleTimeString('ja-JP'),
      level: 'INFO',
      message: 'Threads AutoPoster Studio を起動しました。Excel設定読み込み完了。',
    },
    {
      id: 'init-2',
      timestamp: new Date().toLocaleTimeString('ja-JP'),
      level: 'INFO',
      message: `検出アカウント: @sll15802 (メイン), @jhtwj2, @wqlopakpskapkoq | 未投稿 24 件待機中`,
    },
  ]);

  // Execution state
  const [isRunning, setIsRunning] = useState(false);
  const [isPausedInterval, setIsPausedInterval] = useState(false);
  const [isStealthing, setIsStealthing] = useState(false);
  const [statusText, setStatusText] = useState('待機中 (Excel読込済)');
  const [countdownSeconds, setCountdownSeconds] = useState(0);
  const [activeTab, setActiveTab] = useState<'posts' | 'preview' | 'daily' | 'logs' | 'accounts' | 'downloads'>('posts');
  const [sessionModalOpen, setSessionModalOpen] = useState(false);

  // Sync to localStorage
  useEffect(() => {
    localStorage.setItem('threads_accounts', JSON.stringify(accounts));
  }, [accounts]);

  useEffect(() => {
    localStorage.setItem('threads_account_posts', JSON.stringify(accountPosts));
  }, [accountPosts]);

  useEffect(() => {
    localStorage.setItem('threads_daily_slots', JSON.stringify(dailySlots));
  }, [dailySlots]);

  useEffect(() => {
    localStorage.setItem('threads_settings', JSON.stringify(settings));
  }, [settings]);

  // Current selected account & posts
  const currentAccount = accounts.find((a) => a.id === selectedAccountId) || accounts[0] || initialAccounts[0];
  const currentPosts = accountPosts[currentAccount.id] || [];

  // Helper to append log
  const addLog = (level: LogMessage['level'], message: string) => {
    const newLog: LogMessage = {
      id: `log-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
      timestamp: new Date().toLocaleTimeString('ja-JP'),
      level,
      message,
    };
    setLogs((prev) => [...prev.slice(-300), newLog]);
  };

  // Identify next post
  const unpostedList = currentPosts.filter((p) => p.status === '未投稿');
  const nextPost = unpostedList.length > 0
    ? currentAccount.mode === 1
      ? unpostedList[0] // preview top candidate
      : unpostedList[0]
    : null;

  // Next topic for A/B testing
  const currentTopics = currentAccount.topics || ['信太山新地'];
  const nextTopic = currentTopics.length > 0
    ? currentTopics[currentAccount.currentTopicIndex % currentTopics.length]
    : undefined;

  // Single post execution worker
  const executeSinglePost = async (postToRun?: PostItem, accountToUse?: AccountInfo): Promise<boolean> => {
    const acc = accountToUse || currentAccount;
    const postsList = accountPosts[acc.id] || [];
    const targetPost = postToRun || (
      acc.mode === 1
        ? postsList.filter((p) => p.status === '未投稿')[Math.floor(Math.random() * postsList.filter((p) => p.status === '未投稿').length)]
        : postsList.find((p) => p.status === '未投稿')
    );

    if (!targetPost) {
      if (settings.autoLoop) {
        addLog('INFO', `[${acc.id}] 未投稿ポストがなくなりました。🔁【無限ループ】全ポストを『未投稿』に自動リセットします。`);
        setAccountPosts((prev) => ({
          ...prev,
          [acc.id]: (prev[acc.id] || []).map((p) => ({ ...p, status: '未投稿', completedAt: undefined })),
        }));
        return true;
      } else {
        addLog('WARN', `[${acc.id}] すべてのポストの投稿が完了しました。`);
        return false;
      }
    }

    // 1. Browser launch & media block simulation
    if (settings.blockMedia) {
      addLog('INFO', `[${acc.id}] ⚡【提案2: メディア通信遮断】有効: タイムライン画像・動画をブロック（メモリ60%削減）`);
    }

    // 2. Persistent Profile & Session verification
    const profileFolder = acc.profileDir || `chrome_profiles/${acc.id}`;
    addLog('SUCCESS', `[${acc.id}] 永続プロファイルフォルダ（${profileFolder}）のChromeセッションを読込完了。`);
    addLog('SUCCESS', `[${acc.id}] 認証済みセッションでThreadsに接続しました！（ユーザー: @${acc.id}）`);

    // 3. Stealth Patrol simulation (if enabled)
    if (settings.stealthPatrol) {
      setIsStealthing(true);
      setStatusText(`@${acc.id}: ステルス巡回中 (人間行動模倣)`);
      addLog('STEALTH', `[${acc.id}] 🕵️‍♂️【ステルス巡回】おすすめタイムラインを閲覧中（人間らしい自然な操作ログを生成）...`);

      const stealthDuration = settings.speedMode === 'real' ? 4000 : settings.speedMode === 'fast' ? 2000 : 800;
      await new Promise((r) => setTimeout(r, stealthDuration));

      addLog('STEALTH', `[${acc.id}] 👁️‍🗨️ ステルス巡回完了。投稿フローに移行します。`);
      setIsStealthing(false);
    }

    // 4. Topic A/B rotation & Spintax expansion
    let activeTopic = targetPost.topic || acc.topics[0] || '信太山新地';
    if (settings.topicAbTest && acc.topics.length > 0) {
      activeTopic = acc.topics[acc.currentTopicIndex % acc.topics.length];
      addLog('INFO', `[${acc.id}] 🎯【トピックA/Bテスト】今回トピック: 「${activeTopic}」を適用`);
      // Advance topic index
      setAccounts((prev) =>
        prev.map((a) =>
          a.id === acc.id ? { ...a, currentTopicIndex: (a.currentTopicIndex + 1) % a.topics.length } : a
        )
      );
    }

    // 5. Post text & image resolution
    const expandedText = parseSpintax(targetPost.text);
    const resolvedImg = resolveImagePath(targetPost.imagePath, targetPost.text);

    // Mark as 投稿中
    setAccountPosts((prev) => ({
      ...prev,
      [acc.id]: (prev[acc.id] || []).map((p) =>
        p.id === targetPost.id ? { ...p, status: '投稿中' } : p
      ),
    }));
    setStatusText(`@${acc.id}: Post #${targetPost.id} 投稿中...`);

    addLog('INFO', `[${acc.id}] 「＋ 新しいスレッド」を開き、Post #${targetPost.id} (${expandedText.length}文字) を入力中...`);
    if (resolvedImg) {
      addLog('INFO', `[${acc.id}] 添付画像: ${resolvedImg.split('/').pop()} をThreadsにアップロード`);
    }

    const typingDuration = settings.speedMode === 'real' ? 3000 : settings.speedMode === 'fast' ? 1500 : 600;
    await new Promise((r) => setTimeout(r, typingDuration));

    // Post complete!
    const finishTime = new Date().toLocaleString('ja-JP');
    setAccountPosts((prev) => ({
      ...prev,
      [acc.id]: (prev[acc.id] || []).map((p) =>
        p.id === targetPost.id
          ? { ...p, status: '投稿完了', completedAt: finishTime, errorMessage: undefined }
          : p
      ),
    }));

    addLog('SUCCESS', `[${acc.id}] 🎉 Post #${targetPost.id} 投稿完了！Excel行を自動同期しました。`);

    if (settings.singleWindowRotation) {
      addLog('INFO', `[${acc.id}] 🪟【1窓ローテーション】ブラウザを完全終了しメモリ解放。次回待機に移行`);
    }

    return true;
  };

  // Main automation loop runner
  const runnerRef = useRef<boolean>(false);
  useEffect(() => {
    runnerRef.current = isRunning;
  }, [isRunning]);

  const startAutoPosting = async () => {
    if (isRunning) return;
    setIsRunning(true);
    runnerRef.current = true;
    addLog('INFO', `▶ 自動投稿エンジンを開始しました。(速度モード: ${settings.speedMode})`);

    let accountIndex = accounts.findIndex((a) => a.id === selectedAccountId);
    if (accountIndex === -1) accountIndex = 0;

    while (runnerRef.current) {
      const enabledAccounts = accounts.filter((a) => a.enabled !== false);
      if (enabledAccounts.length === 0) {
        addLog('WARN', '⚠️ 有効なアカウントがありません（全アカウントが停止または要認証です）。');
        stopAutoPosting();
        break;
      }

      accountIndex = accountIndex % enabledAccounts.length;
      const activeAccount = enabledAccounts[accountIndex];
      setSelectedAccountId(activeAccount.id);

      // Check active hours
      // Execute post for this account
      const success = await executeSinglePost(undefined, activeAccount);
      if (!runnerRef.current) break;

      if (!success) {
        addLog('INFO', `[${activeAccount.id}] 未投稿データがありません。`);
      }

      // Rotate account if 1-window rotation is enabled
      if (settings.singleWindowRotation && enabledAccounts.length > 1) {
        accountIndex = (accountIndex + 1) % enabledAccounts.length;
        const nextAcc = enabledAccounts[accountIndex];
        addLog('INFO', `🪟【1窓ローテーション】次回番: @${nextAcc.id} にバトンタッチ`);
      }

      // Calculate interval wait
      let waitSeconds = 10;
      if (settings.speedMode === 'real') {
        const minSec = activeAccount.minIntervalMinutes * 60;
        const maxSec = activeAccount.maxIntervalMinutes * 60;
        waitSeconds = Math.floor(Math.random() * (maxSec - minSec + 1)) + minSec;
      } else if (settings.speedMode === 'fast') {
        waitSeconds = Math.floor(Math.random() * 6) + 8; // 8〜13秒
      } else {
        waitSeconds = 2; // 2秒
      }

      addLog('INFO', `[*] BAN防止のため ${waitSeconds >= 60 ? `${Math.floor(waitSeconds / 60)}分` : `${waitSeconds}秒`} のランダム間隔で安全待機します...`);

      setIsPausedInterval(true);
      setStatusText(`待機中 (BAN防止インターバル)`);
      setCountdownSeconds(waitSeconds);

      // Countdown ticker
      for (let s = waitSeconds; s > 0; s--) {
        if (!runnerRef.current) break;
        setCountdownSeconds(s);
        await new Promise((r) => setTimeout(r, 1000));
      }

      setIsPausedInterval(false);
      setCountdownSeconds(0);
    }

    setIsRunning(false);
    setIsPausedInterval(false);
    setIsStealthing(false);
    setStatusText('待機中 (停止済)');
    addLog('WARN', '⏹ 自動投稿を停止しました。');
  };

  const stopAutoPosting = () => {
    runnerRef.current = false;
    setIsRunning(false);
    setIsPausedInterval(false);
    setIsStealthing(false);
    setStatusText('停止処理中...');
    setCountdownSeconds(0);
  };

  // Immediate single test post
  const handleImmediateTestPost = async () => {
    if (isRunning) return;
    addLog('INFO', `⚡ [${currentAccount.id}] 1件テスト投稿を即時実行します...`);
    await executeSinglePost();
    setStatusText('待機中 (テスト完了)');
  };

  // Reset all posts for current account to '未投稿'
  const handleResetToUnposted = () => {
    setAccountPosts((prev) => ({
      ...prev,
      [currentAccount.id]: (prev[currentAccount.id] || []).map((p) => ({
        ...p,
        status: '未投稿',
        completedAt: undefined,
        errorMessage: undefined,
      })),
    }));
    addLog('INFO', `🔁 @${currentAccount.id} の全投稿ステータスを『未投稿』にリセットしました。`);
  };

  // Clear all posts for current account
  const handleClearAllPosts = () => {
    const prevCount = (accountPosts[currentAccount.id] || []).length;
    setAccountPosts((prev) => ({
      ...prev,
      [currentAccount.id]: [],
    }));
    addLog('WARN', `🗑️ @${currentAccount.id} の既存ポスト（${prevCount}件）を全件削除しました。`);
  };

  // Bulk add posts (supporting 100+ items at once)
  const handleBulkAddPosts = (newItems: Omit<PostItem, 'id'>[], replaceExisting: boolean) => {
    if (replaceExisting) {
      const formatted: PostItem[] = newItems.map((p, idx) => ({
        ...p,
        id: idx + 1,
      }));
      setAccountPosts((prev) => ({
        ...prev,
        [currentAccount.id]: formatted,
      }));
      addLog('SUCCESS', `🚀 @${currentAccount.id} のポストを全件入れ替えし、新規 ${formatted.length} 件を一括登録しました！`);
    } else {
      const existing = accountPosts[currentAccount.id] || [];
      const startId = existing.reduce((m, p) => Math.max(m, p.id), 0);
      const formatted: PostItem[] = newItems.map((p, idx) => ({
        ...p,
        id: startId + idx + 1,
      }));
      setAccountPosts((prev) => ({
        ...prev,
        [currentAccount.id]: [...existing, ...formatted],
      }));
      addLog('SUCCESS', `🚀 @${currentAccount.id} に新規 ${formatted.length} 件のポストを追加しました！（合計 ${existing.length + formatted.length} 件）`);
    }
  };

  // Apply Anti-Freeze Spintax to all daily slots
  const handleApplyAntiFreezeSpintax = () => {
    const secureSlots: DailyGreetingSlot[] = [
      {
        id: 'secure_morning_1',
        slotName: '朝枠',
        timeRange: '7:00-9:00',
        text: '{おはようございます☀️|皆さんおはようございます！☕️|良い朝ですね🌿|今朝もすっきり目覚めました✨} {今日もいい天気！|本日も元気に|ぼちぼちマイペースで|気合い入れて}頑張りましょう〜\n\n素敵な{1日|一日|日}になりますように✨ ({today})',
        enabled: true,
      },
      {
        id: 'secure_morning_2',
        slotName: '朝枠',
        timeRange: '7:00-9:00',
        text: '{おはようございます|朝のコーヒータイム☕️|朝活スタート！} 朝から{しっかりエネルギー補給|ストレッチして|準備万端で}気合い入れてます💪 {今日も良い出会いがありますように|本日も無理せずいきましょう}！',
        enabled: true,
      },
      {
        id: 'secure_lunch_1',
        slotName: '昼枠',
        timeRange: '12:00-13:30',
        text: '{やっとお昼休憩！🍙|お昼ご飯タイムです✨|午後からも頑張ろう〜} 今日は{一段とバタバタしてる💦|比較的落ち着いてるかも🍵|あっという間に時間が過ぎる！}\n\n美味しいもの食べて午後も乗り切ろう〜！皆さんはお昼何食べましたか？',
        enabled: true,
      },
      {
        id: 'secure_night_1',
        slotName: '夜枠',
        timeRange: '21:00-23:30',
        text: '{今日も1日本当に忙しかった〜！|本日もお疲れ様でした🌙|あっという間に夜ですね🍵} お風呂入ってゆっくり{温まります|リフレッシュします|休みます}。\n\n皆さんも{暖かくして|無理せず}ゆっくり休んでくださいね✨',
        enabled: true,
      },
      {
        id: 'secure_night_2',
        slotName: '夜枠',
        timeRange: '21:00-23:30',
        text: '{今日のタスクがひと段落しました🌙|夜風が心地いい時間帯ですね|今日も充実した1日でした✨}\n\n明日も{いい日になりますように|素敵な1日にしていきましょう}。温かくしておやすみなさい🛌 ({date})',
        enabled: true,
      },
    ];
    setDailySlots(secureSlots);
    addLog('SUCCESS', '🛡️【凍結防止】全日常挨拶枠に何百通りも自動分散する高精度スピンタックスと動的日付を一括適用しました！');
  };

  // Restore defaults
  const handleReloadDefaultData = () => {
    setAccounts(initialAccounts);
    setSelectedAccountId(initialAccounts[0].id);
    const initialMap: Record<string, PostItem[]> = {};
    initialAccounts.forEach((acc) => {
      initialMap[acc.id] = JSON.parse(JSON.stringify(initialPosts));
    });
    setAccountPosts(initialMap);
    setDailySlots(initialDailySlots);
    setSettings(initialSettings);
    addLog('SUCCESS', '🔄 初期プリセットデータ（アカウント情報.xlsx, sll15802.xlsx）を再読み込みしました。');
  };

  // Export current Excel
  const handleExportExcel = () => {
    exportAccountWorkbook(currentAccount, currentPosts, dailySlots);
    addLog('SUCCESS', `📥 『${currentAccount.id}.xlsx』を出力しました。(設定・投稿リスト・日常ポスト)`);
  };

  const handleExportAccountsConfig = () => {
    exportAccountsConfigWorkbook(accounts);
    addLog('SUCCESS', `📥 『アカウント情報.xlsx』を出力しました。`);
  };

  // Import Excel file
  const handleImportExcel = async (file: File) => {
    try {
      addLog('INFO', `📤 Excelファイル解析中: ${file.name}...`);
      const parsed = await parseExcelFile(file);

      if (parsed.accounts && parsed.accounts.length > 0) {
        const merged: AccountInfo[] = parsed.accounts.map((a, idx) => ({
          id: a.id || `acc_${idx + 1}`,
          authToken: a.authToken || '',
          userId: a.userId || extractUserIdFromToken(a.authToken || ''),
          notes: a.notes || '',
          minIntervalMinutes: a.minIntervalMinutes || 120,
          maxIntervalMinutes: a.maxIntervalMinutes || 360,
          activeHours: a.activeHours || '0:00-0:00',
          mode: a.mode ?? 1,
          profileDir: a.profileDir || `chrome_profiles/${a.id}`,
          topics: a.topics || ['信太山新地', '大阪求人'],
          currentTopicIndex: 0,
          enabled: a.enabled !== false,
        }));
        setAccounts(merged);
        setSelectedAccountId(merged[0].id);
        addLog('SUCCESS', `✅ アカウント情報.xlsx から ${merged.length} 件のアカウントを読み込みました！`);
      }

      if (parsed.posts && parsed.posts.length > 0) {
        setAccountPosts((prev) => ({
          ...prev,
          [currentAccount.id]: parsed.posts!,
        }));
        addLog('SUCCESS', `✅ 投稿リストから ${parsed.posts.length} 件のポストデータを読み込みました！`);
      }

      if (parsed.dailySlots && parsed.dailySlots.length > 0) {
        setDailySlots(parsed.dailySlots);
        addLog('SUCCESS', `✅ 日常ポスト枠から ${parsed.dailySlots.length} 件の枠データを読み込みました！`);
      }
    } catch (err: any) {
      addLog('ERROR', `Excel読み込みエラー: ${err.message || err}`);
    }
  };

  // Direct client-side download of pure Python script (zero HTML proxy interference)
  const handleDirectDownloadPyScript = () => {
    try {
      const blob = new Blob([RAW_THREADS_GUI_APP_PY], { type: 'text/x-python;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'threads_gui_app.py';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      addLog('SUCCESS', '📥 threads_gui_app.py (最新版 231KB) のクリーンダウンロードを実行しました！');
    } catch (err: any) {
      addLog('ERROR', `ダウンロード失敗: ${err.message}`);
    }
  };

  // Direct client-side download of profile_launcher.py
  const handleDirectDownloadProfileLauncher = () => {
    try {
      const blob = new Blob([RAW_PROFILE_LAUNCHER_PY], { type: 'text/x-python;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'profile_launcher.py';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      addLog('SUCCESS', '📥 profile_launcher.py (プロファイル生成・ログイン永続化ツール) をダウンロードしました！');
    } catch (err: any) {
      addLog('ERROR', `ダウンロード失敗: ${err.message}`);
    }
  };

  // Simulate detecting the phone number verification screen (SMS checkpoint)
  const handleSimulateCheckpoint = (accId: string) => {
    setAccounts((prev) =>
      prev.map((a) =>
        a.id === accId
          ? {
              ...a,
              enabled: false,
              checkpointReason: '携帯電話番号認証要求画面検知',
              notes: `${a.notes || ''} [${new Date().toLocaleString('ja-JP')} 自動停止: 携帯電話番号認証画面検知]`.trim(),
            }
          : a
      )
    );
    addLog('ERROR', `🚨 [CRITICAL] [@${accId}] 携帯電話番号認証（SMS確認）画面を検知しました！（シミュレーションテスト）`);
    addLog('WARN', `🛡️ [@${accId}] アカウント情報.xlsx の【有効/無効】を「無効」に自動更新し、直ちに自動投稿を停止しました。`);
    addLog('INFO', `📸 認証要求画面のスクリーンショットを error_screenshots/phone_checkpoint_${accId}.png に保存しました。`);
    addLog('SUCCESS', `✅ 他の有効なアカウントの自動投稿はそのまま安全に継続されます。`);
  };

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100 flex flex-col font-sans selection:bg-emerald-900 selection:text-white">
      {/* Top Header & Toolbar */}
      <Header
        isRunning={isRunning}
        isPausedInterval={isPausedInterval}
        isStealthing={isStealthing}
        statusText={statusText}
        countdownSeconds={countdownSeconds}
        onStart={startAutoPosting}
        onStop={stopAutoPosting}
        onTestPost={handleImmediateTestPost}
        onOpenSessionModal={() => setSessionModalOpen(true)}
        onResetToUnposted={handleResetToUnposted}
        onReloadDefaultData={handleReloadDefaultData}
        onExportExcel={handleExportExcel}
        onImportExcel={handleImportExcel}
        onClearLogs={() => setLogs([])}
        onDownloadPyScript={handleDirectDownloadPyScript}
        onDownloadProfileLauncher={handleDirectDownloadProfileLauncher}
        onDownloadAccountsExcel={handleExportAccountsConfig}
      />

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 space-y-4">
        {/* 🌟 Prominent Top Download Card - Always Visible */}
        <div className="p-4 sm:p-5 rounded-2xl bg-gradient-to-r from-zinc-900 via-indigo-950/70 to-zinc-900 border-2 border-indigo-500/80 shadow-2xl flex flex-col xl:flex-row items-center justify-between gap-4">
          <div className="space-y-1.5 text-center xl:text-left">
            <div className="flex items-center gap-2 justify-center xl:justify-start flex-wrap">
              <span className="px-2.5 py-0.5 rounded-full bg-violet-500/20 text-violet-300 border border-violet-500/40 font-mono text-[11px] font-bold">
                🎉 新ツール: profile_launcher 付属
              </span>
              <span className="px-2.5 py-0.5 rounded-full bg-emerald-500/20 text-emerald-300 border border-emerald-500/40 font-mono text-[11px] font-bold">
                強制ログアウト根絶・sessionid不要
              </span>
            </div>
            <h2 className="text-base sm:text-lg font-bold text-white flex items-center justify-center xl:justify-start gap-2">
              <Rocket className="w-5 h-5 text-indigo-400" />
              <span>Windows版 更新ファイル & プロファイル生成ツール</span>
            </h2>
            <p className="text-zinc-300 text-xs">
              定期ログアウトの原因だったsessionid入力は完全撤廃。新ツール <strong className="text-violet-300 font-mono">profile_launcher.py</strong> でChromeを開き、IDとPWで1回ログインするだけで永続保存されます。
            </p>
          </div>

          <div className="flex items-center gap-2.5 flex-wrap justify-center shrink-0">
            {/* New Profile Launcher Button */}
            <button
              onClick={handleDirectDownloadProfileLauncher}
              className="flex items-center gap-2 px-4 py-3 bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-violet-500 hover:to-indigo-500 text-white font-bold rounded-xl shadow-lg hover:shadow-violet-900/50 text-sm cursor-pointer transition-all active:scale-95 border border-violet-400/50 animate-pulse"
              title="新アプリ: profile_launcher.py（プロファイル生成・ログイン永続化ツール）をダウンロード"
            >
              <Rocket className="w-4 h-4 text-violet-200" />
              <span>🚀 profile_launcher.py</span>
            </button>

            {/* Download Button 1 */}
            <button
              onClick={handleDirectDownloadPyScript}
              className="flex items-center gap-2 px-4 py-3 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-bold rounded-xl shadow-lg hover:shadow-emerald-900/50 text-sm cursor-pointer transition-all active:scale-95 border border-emerald-400/50"
              title="最新の threads_gui_app.py (231KB) を直接ダウンロード"
            >
              <Download className="w-4 h-4" />
              <span>🐍 threads_gui_app.py</span>
            </button>

            <button
              onClick={handleExportAccountsConfig}
              className="flex items-center gap-2 px-4 py-3 bg-gradient-to-r from-sky-600 to-blue-600 hover:from-sky-500 hover:to-blue-500 text-white font-bold rounded-xl shadow-lg hover:shadow-sky-900/50 text-sm cursor-pointer transition-all active:scale-95 border border-sky-400/50"
              title="sessionid・プロファイル永続化対応の最新のアカウント情報.xlsx を直接ダウンロード"
            >
              <Download className="w-4 h-4" />
              <span>📊 アカウント情報.xlsx</span>
            </button>

            {/* Guide Shortcut */}
            <button
              onClick={() => setActiveTab('downloads')}
              className="px-3.5 py-3 bg-zinc-800 hover:bg-zinc-700 text-zinc-300 hover:text-white rounded-xl text-xs font-semibold cursor-pointer border border-zinc-700 transition-colors"
              title="配置手順マニュアルタブへ移動"
            >
              <span>📁 使い方・配置手順</span>
            </button>
          </div>
        </div>

        {/* Account Selector & Settings Bar */}
        <AccountSelector
          accounts={accounts}
          selectedAccountId={selectedAccountId}
          onSelectAccount={(id) => setSelectedAccountId(id)}
          onUpdateAccount={(updated) => {
            setAccounts((prev) => prev.map((a) => (a.id === updated.id ? updated : a)));
            addLog('SUCCESS', `[${updated.id}] アカウント設定を更新しました。(有効/無効: ${updated.enabled !== false ? '有効' : '無効'})`);
          }}
          onAddAccount={(newAcc) => {
            setAccounts((prev) => [...prev, newAcc]);
            setAccountPosts((prev) => ({
              ...prev,
              [newAcc.id]: JSON.parse(JSON.stringify(initialPosts)),
            }));
            setSelectedAccountId(newAcc.id);
            addLog('SUCCESS', `新規アカウント @${newAcc.id} を追加しました。`);
          }}
          onSimulateCheckpoint={handleSimulateCheckpoint}
        />

        {/* 6 Key BAN Prevention Master Switches */}
        <FeatureToggles
          settings={settings}
          onToggle={(key, val) => {
            setSettings((prev) => ({ ...prev, [key]: val }));
            addLog('INFO', `設定切替: [${String(key)}] = ${val}`);
          }}
        />

        {/* Navigation Tabs */}
        <div className="flex items-center gap-1.5 border-b border-zinc-800 pb-2 text-xs overflow-x-auto">
          <button
            onClick={() => setActiveTab('posts')}
            className={`flex items-center gap-1.5 px-3.5 py-2 rounded-lg font-medium transition-all cursor-pointer whitespace-nowrap ${
              activeTab === 'posts'
                ? 'bg-zinc-800 text-white font-bold border border-zinc-700 shadow-sm'
                : 'text-zinc-400 hover:text-zinc-200 hover:bg-zinc-900'
            }`}
          >
            <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-400" />
            <span>投稿リスト & Excel ({currentPosts.length})</span>
          </button>

          <button
            onClick={() => setActiveTab('preview')}
            className={`flex items-center gap-1.5 px-3.5 py-2 rounded-lg font-medium transition-all cursor-pointer whitespace-nowrap ${
              activeTab === 'preview'
                ? 'bg-zinc-800 text-white font-bold border border-zinc-700 shadow-sm'
                : 'text-zinc-400 hover:text-zinc-200 hover:bg-zinc-900'
            }`}
          >
            <Smartphone className="w-3.5 h-3.5 text-sky-400" />
            <span>実機プレビュー & 次回投稿</span>
          </button>

          <button
            onClick={() => setActiveTab('daily')}
            className={`flex items-center gap-1.5 px-3.5 py-2 rounded-lg font-medium transition-all cursor-pointer whitespace-nowrap ${
              activeTab === 'daily'
                ? 'bg-zinc-800 text-white font-bold border border-zinc-700 shadow-sm'
                : 'text-zinc-400 hover:text-zinc-200 hover:bg-zinc-900'
            }`}
          >
            <Sun className="w-3.5 h-3.5 text-amber-400" />
            <span>日常つぶやき・挨拶枠 ({dailySlots.length})</span>
          </button>

          <button
            onClick={() => setActiveTab('logs')}
            className={`flex items-center gap-1.5 px-3.5 py-2 rounded-lg font-medium transition-all cursor-pointer whitespace-nowrap ${
              activeTab === 'logs'
                ? 'bg-zinc-800 text-white font-bold border border-zinc-700 shadow-sm'
                : 'text-zinc-400 hover:text-zinc-200 hover:bg-zinc-900'
            }`}
          >
            <Terminal className="w-3.5 h-3.5 text-purple-400" />
            <span>実行コンソール ({logs.length})</span>
          </button>

          <button
            onClick={() => setActiveTab('accounts')}
            className={`flex items-center gap-1.5 px-3.5 py-2 rounded-lg font-medium transition-all cursor-pointer whitespace-nowrap ${
              activeTab === 'accounts'
                ? 'bg-zinc-800 text-white font-bold border border-zinc-700 shadow-sm'
                : 'text-zinc-400 hover:text-zinc-200 hover:bg-zinc-900'
            }`}
          >
            <Users className="w-3.5 h-3.5 text-teal-400" />
            <span>アカウント & Excel一括</span>
          </button>

          <button
            onClick={() => setActiveTab('downloads')}
            className={`flex items-center gap-1.5 px-3.5 py-2 rounded-lg font-medium transition-all cursor-pointer whitespace-nowrap ${
              activeTab === 'downloads'
                ? 'bg-indigo-600 text-white font-bold shadow-md border border-indigo-500'
                : 'text-indigo-300 hover:text-white hover:bg-indigo-950/60 border border-indigo-500/40'
            }`}
          >
            <Package className="w-3.5 h-3.5 text-indigo-300" />
            <span>📦 Windows更新ダウンロード</span>
          </button>
        </div>

        {/* Tab Content Views */}
        <div>
          {activeTab === 'posts' && (
            <PostsTable
              posts={currentPosts}
              currentAccountId={currentAccount.id}
              onUpdatePost={(updated) => {
                setAccountPosts((prev) => ({
                  ...prev,
                  [currentAccount.id]: (prev[currentAccount.id] || []).map((p) =>
                    p.id === updated.id ? updated : p
                  ),
                }));
                addLog('INFO', `[${currentAccount.id}] Post #${updated.id} を更新しました。`);
              }}
              onDeletePost={(id) => {
                setAccountPosts((prev) => ({
                  ...prev,
                  [currentAccount.id]: (prev[currentAccount.id] || []).filter((p) => p.id !== id),
                }));
                addLog('INFO', `[${currentAccount.id}] Post #${id} を削除しました。`);
              }}
              onAddPost={(newP) => {
                const maxId = currentPosts.reduce((m, p) => Math.max(m, p.id), 0);
                const post: PostItem = { ...newP, id: maxId + 1 };
                setAccountPosts((prev) => ({
                  ...prev,
                  [currentAccount.id]: [...(prev[currentAccount.id] || []), post],
                }));
                addLog('SUCCESS', `[${currentAccount.id}] 新規 Post #${post.id} を追加しました。`);
              }}
              onClearAllPosts={handleClearAllPosts}
              onBulkAddPosts={handleBulkAddPosts}
              onPostSingle={async (p) => {
                await executeSinglePost(p);
              }}
            />
          )}

          {activeTab === 'preview' && (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              <ThreadsPreview
                currentAccount={currentAccount}
                nextPost={nextPost}
                isRunning={isRunning}
                isStealthing={isStealthing}
                countdownSeconds={countdownSeconds}
                onPostNow={async () => {
                  if (nextPost) await executeSinglePost(nextPost);
                }}
                selectedTopic={nextTopic}
              />
              <ConsoleLogs logs={logs} onClear={() => setLogs([])} />
            </div>
          )}

          {activeTab === 'daily' && (
            <DailyGreetings
              slots={dailySlots}
              onUpdateSlot={(updated) => {
                setDailySlots((prev) => prev.map((s) => (s.id === updated.id ? updated : s)));
                addLog('INFO', `日常枠 [${updated.slotName}] を更新しました。`);
              }}
              onAddSlot={(newSlot) => {
                const id = `daily_${Date.now()}`;
                setDailySlots((prev) => [...prev, { ...newSlot, id }]);
                addLog('SUCCESS', `新しい日常枠 [${newSlot.slotName}] を追加しました。`);
              }}
              onDeleteSlot={(id) => {
                setDailySlots((prev) => prev.filter((s) => s.id !== id));
                addLog('INFO', `日常枠を削除しました。`);
              }}
              onApplyAntiFreezeSpintax={handleApplyAntiFreezeSpintax}
            />
          )}

          {activeTab === 'logs' && (
            <ConsoleLogs logs={logs} onClear={() => setLogs([])} />
          )}

          {activeTab === 'accounts' && (
            <div className="bg-zinc-900 border border-zinc-800 rounded-xl p-5 space-y-4 shadow-sm text-xs">
              <div className="flex items-center justify-between pb-3 border-b border-zinc-800">
                <div>
                  <h3 className="text-sm font-bold text-white flex items-center gap-1.5">
                    <Users className="w-4 h-4 text-teal-400" />
                    マルチアカウント一括管理 & Excel入出力
                  </h3>
                  <p className="text-zinc-400 text-xs mt-0.5">
                    「アカウント情報.xlsx」と各アカウントの個別投稿Excelを一括エクスポート・インポート
                  </p>
                </div>

                <button
                  onClick={handleExportAccountsConfig}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 border border-zinc-700 rounded-lg cursor-pointer"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>アカウント情報.xlsx 出力</span>
                </button>
              </div>

              {/* Accounts table */}
              <div className="overflow-x-auto rounded-lg border border-zinc-800">
                <table className="w-full text-left">
                  <thead className="bg-zinc-950 text-zinc-400 font-semibold border-b border-zinc-800">
                    <tr>
                      <th className="py-2.5 px-3">アカウントID</th>
                      <th className="py-2.5 px-3">状態 (有効/無効)</th>
                      <th className="py-2.5 px-3">sessionid (セッションCookie)</th>
                      <th className="py-2.5 px-3">Chromeプロファイル (永続化)</th>
                      <th className="py-2.5 px-3">待機間隔</th>
                      <th className="py-2.5 px-3">未投稿 / 総数</th>
                      <th className="py-2.5 px-3 text-right">Excel出力</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-zinc-800/60">
                    {accounts.map((acc) => {
                      const pList = accountPosts[acc.id] || [];
                      const unposted = pList.filter((p) => p.status === '未投稿').length;
                      const isAccActive = acc.enabled !== false;
                      const profName = acc.profileDir ? acc.profileDir.replace('chrome_profiles/', '') : acc.id;
                      return (
                        <tr key={acc.id} className="hover:bg-zinc-800/30">
                          <td className="py-2.5 px-3 font-semibold text-white">
                            @{acc.id}
                          </td>
                          <td className="py-2.5 px-3">
                            <button
                              onClick={() => {
                                const nextState = !isAccActive;
                                setAccounts((prev) =>
                                  prev.map((a) => (a.id === acc.id ? { ...a, enabled: nextState } : a))
                                );
                                addLog('INFO', `[@${acc.id}] 状態を ${nextState ? '有効' : '無効'} に変更しました。`);
                              }}
                              className={`px-2 py-0.5 rounded text-[10px] font-semibold cursor-pointer border ${
                                isAccActive
                                  ? 'bg-emerald-950 text-emerald-300 border-emerald-800'
                                  : 'bg-rose-950 text-rose-300 border-rose-800'
                              }`}
                              title="クリックで有効/無効を切替"
                            >
                              {isAccActive ? '● 有効' : '🔴 停止中'}
                            </button>
                          </td>
                          <td className="py-2.5 px-3 font-mono text-[11px] text-zinc-400">
                            {acc.authToken && !acc.authToken.startsWith('ここに') ? (
                              <span className="text-emerald-400 font-semibold" title={acc.authToken}>
                                {acc.authToken.length > 20 ? `${acc.authToken.slice(0, 18)}...` : acc.authToken}
                              </span>
                            ) : (
                              <span className="text-zinc-600">未設定</span>
                            )}
                          </td>
                          <td className="py-2.5 px-3 font-mono text-zinc-300">
                            <span className="px-2 py-0.5 rounded bg-zinc-950 border border-zinc-700 text-indigo-300 text-[11px]">
                              📁 chrome_profiles/{profName}
                            </span>
                            <span className="ml-2 text-[10px] text-emerald-400 font-sans">
                              (profile_launcher対応)
                            </span>
                          </td>
                          <td className="py-2.5 px-3 text-zinc-300">
                            {acc.minIntervalMinutes}〜{acc.maxIntervalMinutes}分
                          </td>
                          <td className="py-2.5 px-3">
                            <span className="font-bold text-sky-400">{unposted}</span> / {pList.length}
                          </td>
                          <td className="py-2.5 px-3 text-right">
                            <button
                              onClick={() => exportAccountWorkbook(acc, pList, dailySlots)}
                              className="px-2.5 py-1 bg-zinc-800 hover:bg-zinc-700 text-zinc-300 rounded text-[11px] cursor-pointer inline-flex items-center gap-1"
                            >
                              <Download className="w-3 h-3" />
                              <span>{acc.id}.xlsx</span>
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* Sample Files Download Banner */}
              <div className="p-4 rounded-xl bg-zinc-950 border border-zinc-800 flex flex-wrap items-center justify-between gap-3">
                <div className="flex items-center gap-2">
                  <FileText className="w-5 h-5 text-emerald-400" />
                  <div>
                    <div className="font-bold text-white text-xs">リポジトリ同梱の実機サンプルExcelファイル</div>
                    <div className="text-[11px] text-zinc-400">
                      GitHubリポジトリ（ThreadsAutoPoster）に同梱されているオリジナルの.xlsxファイルをダウンロードできます
                    </div>
                  </div>
                </div>

                <div className="flex items-center gap-2 flex-wrap">
                  <a
                    href="/data/sll15802.xlsx"
                    download="sll15802.xlsx"
                    className="px-2.5 py-1 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 rounded text-xs flex items-center gap-1"
                  >
                    <Download className="w-3 h-3" />
                    <span>sll15802.xlsx</span>
                  </a>
                  <a
                    href="/data/アカウント情報.xlsx"
                    download="アカウント情報.xlsx"
                    className="px-2.5 py-1 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 rounded text-xs flex items-center gap-1"
                  >
                    <Download className="w-3 h-3" />
                    <span>アカウント情報.xlsx</span>
                  </a>
                  <a
                    href="/data/jhtwj2.xlsx"
                    download="jhtwj2.xlsx"
                    className="px-2.5 py-1 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 rounded text-xs flex items-center gap-1"
                  >
                    <Download className="w-3 h-3" />
                    <span>jhtwj2.xlsx</span>
                  </a>
                </div>
              </div>
            </div>
          )}

          {activeTab === 'downloads' && (
            <DownloadCenter accounts={accounts} />
          )}
        </div>
      </main>

      {/* Session ID / Auth Token Modal */}
      <SessionModal
        isOpen={sessionModalOpen}
        onClose={() => setSessionModalOpen(false)}
        currentAuthToken={currentAccount.authToken}
        onSaveToken={(token) => {
          const userId = extractUserIdFromToken(token);
          setAccounts((prev) =>
            prev.map((a) =>
              a.id === currentAccount.id
                ? { ...a, authToken: token, userId: userId || a.userId }
                : a
            )
          );
          addLog('SUCCESS', `[@${currentAccount.id}] Authトークン (sessionid) を保存しました。`);
        }}
        accountId={currentAccount.id}
      />
    </div>
  );
}
