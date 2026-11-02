import React from 'react';
import { SystemSettings } from '../types';
import { 
  Repeat, 
  AppWindow, 
  Zap, 
  Eye, 
  ShieldAlert, 
  Shuffle, 
  SunMedium, 
  Gauge 
} from 'lucide-react';

interface FeatureTogglesProps {
  settings: SystemSettings;
  onToggle: <K extends keyof SystemSettings>(key: K, value: SystemSettings[K]) => void;
}

export const FeatureToggles: React.FC<FeatureTogglesProps> = ({ settings, onToggle }) => {
  return (
    <div className="bg-zinc-900/90 border border-zinc-800 rounded-xl p-4 shadow-sm mb-4">
      <div className="flex items-center justify-between pb-3 border-b border-zinc-800 mb-3">
        <h3 className="text-xs font-bold text-zinc-300 uppercase tracking-wider flex items-center gap-1.5">
          <Zap className="w-3.5 h-3.5 text-emerald-400" />
          高度なBAN防止 & 最適化スイッチ (6大マスター機能)
        </h3>

        {/* Speed multiplier for browser simulation */}
        <div className="flex items-center gap-1.5 bg-zinc-950 px-2.5 py-1 rounded-lg border border-zinc-800 text-xs">
          <Gauge className="w-3.5 h-3.5 text-sky-400" />
          <span className="text-zinc-400 text-[11px] hidden sm:inline">待機インターバル速度:</span>
          <select
            value={settings.speedMode}
            onChange={(e) =>
              onToggle('speedMode', e.target.value as SystemSettings['speedMode'])
            }
            className="bg-transparent text-white font-medium text-xs focus:outline-none cursor-pointer"
          >
            <option value="real" className="bg-zinc-900 text-zinc-200">
              実運用 (1x: 120〜360分)
            </option>
            <option value="fast" className="bg-zinc-900 text-zinc-200">
              高速テスト (10秒)
            </option>
            <option value="instant" className="bg-zinc-900 text-zinc-200">
              即時テスト (2秒)
            </option>
          </select>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-2 text-xs">
        {/* Toggle 1: Auto Loop */}
        <label className="flex items-start gap-2.5 p-2 rounded-lg hover:bg-zinc-800/40 cursor-pointer transition-colors border border-transparent hover:border-zinc-800">
          <input
            type="checkbox"
            checked={settings.autoLoop}
            onChange={(e) => onToggle('autoLoop', e.target.checked)}
            className="mt-0.5 rounded border-zinc-700 bg-zinc-950 text-emerald-500 focus:ring-emerald-500 cursor-pointer accent-emerald-500"
          />
          <div className="flex-1">
            <span className="text-emerald-400 font-semibold flex items-center gap-1">
              <Repeat className="w-3.5 h-3.5" />
              無限ループ投稿
            </span>
            <p className="text-zinc-400 text-[11px] mt-0.5">
              全ポスト投稿完了後に自動で『未投稿』にリセットして最初から周回
            </p>
          </div>
        </label>

        {/* Toggle 2: Single Window Rotation */}
        <label className="flex items-start gap-2.5 p-2 rounded-lg hover:bg-zinc-800/40 cursor-pointer transition-colors border border-transparent hover:border-zinc-800">
          <input
            type="checkbox"
            checked={settings.singleWindowRotation}
            onChange={(e) => onToggle('singleWindowRotation', e.target.checked)}
            className="mt-0.5 rounded border-zinc-700 bg-zinc-950 text-sky-500 focus:ring-sky-500 cursor-pointer accent-sky-500"
          />
          <div className="flex-1">
            <span className="text-sky-400 font-semibold flex items-center gap-1">
              <AppWindow className="w-3.5 h-3.5" />
              【提案1】1窓・順番交代ローテーション
            </span>
            <p className="text-zinc-400 text-[11px] mt-0.5">
              常に開くブラウザは1つだけ。投稿後に完全終了してメモリ解放＆待機順に切替
            </p>
          </div>
        </label>

        {/* Toggle 3: Block Media */}
        <label className="flex items-start gap-2.5 p-2 rounded-lg hover:bg-zinc-800/40 cursor-pointer transition-colors border border-transparent hover:border-zinc-800">
          <input
            type="checkbox"
            checked={settings.blockMedia}
            onChange={(e) => onToggle('blockMedia', e.target.checked)}
            className="mt-0.5 rounded border-zinc-700 bg-zinc-950 text-purple-500 focus:ring-purple-500 cursor-pointer accent-purple-500"
          />
          <div className="flex-1">
            <span className="text-purple-400 font-semibold flex items-center gap-1">
              <Zap className="w-3.5 h-3.5" />
              【提案2】画像・動画通信ブロック
            </span>
            <p className="text-zinc-400 text-[11px] mt-0.5">
              タイムライン画像・動画・リールを遮断（メモリ50-70%削減・CPU静音化）
            </p>
          </div>
        </label>

        {/* Toggle 4: Stealth Patrol */}
        <label className="flex items-start gap-2.5 p-2 rounded-lg hover:bg-zinc-800/40 cursor-pointer transition-colors border border-transparent hover:border-zinc-800">
          <input
            type="checkbox"
            checked={settings.stealthPatrol}
            onChange={(e) => onToggle('stealthPatrol', e.target.checked)}
            className="mt-0.5 rounded border-zinc-700 bg-zinc-950 text-amber-500 focus:ring-amber-500 cursor-pointer accent-amber-500"
          />
          <div className="flex-1">
            <span className="text-amber-400 font-semibold flex items-center gap-1">
              <Eye className="w-3.5 h-3.5" />
              【新機能1】ステルス巡回
            </span>
            <p className="text-zinc-400 text-[11px] mt-0.5">
              投稿前におすすめタイムラインを数秒閲覧（人間らしい行動模倣・BANリスク回避）
            </p>
          </div>
        </label>

        {/* Toggle 5: Auto Clean Profile */}
        <label className="flex items-start gap-2.5 p-2 rounded-lg hover:bg-zinc-800/40 cursor-pointer transition-colors border border-transparent hover:border-zinc-800">
          <input
            type="checkbox"
            checked={settings.autoCleanProfile}
            onChange={(e) => onToggle('autoCleanProfile', e.target.checked)}
            className="mt-0.5 rounded border-zinc-700 bg-zinc-950 text-rose-500 focus:ring-rose-500 cursor-pointer accent-rose-500"
          />
          <div className="flex-1">
            <span className="text-rose-400 font-semibold flex items-center gap-1">
              <ShieldAlert className="w-3.5 h-3.5" />
              【新機能2】自動プロファイル初期化
            </span>
            <p className="text-zinc-400 text-[11px] mt-0.5">
              Meta警告や連続エラー検知時、手動操作不要でCookie/指紋を安全リセット
            </p>
          </div>
        </label>

        {/* Toggle 6: Topic A/B Test */}
        <label className="flex items-start gap-2.5 p-2 rounded-lg hover:bg-zinc-800/40 cursor-pointer transition-colors border border-transparent hover:border-zinc-800">
          <input
            type="checkbox"
            checked={settings.topicAbTest}
            onChange={(e) => onToggle('topicAbTest', e.target.checked)}
            className="mt-0.5 rounded border-zinc-700 bg-zinc-950 text-teal-500 focus:ring-teal-500 cursor-pointer accent-teal-500"
          />
          <div className="flex-1">
            <span className="text-teal-400 font-semibold flex items-center gap-1">
              <Shuffle className="w-3.5 h-3.5" />
              【新機能3】トピックA/Bテスト・自動切替
            </span>
            <p className="text-zinc-400 text-[11px] mt-0.5">
              Excelの複数トピック（信太山新地, 大阪求人, 高収入バイト等）を投稿ごとに自動均等切替
            </p>
          </div>
        </label>

        {/* Toggle 7: Daily Greeting */}
        <label className="flex items-start gap-2.5 p-2 rounded-lg hover:bg-zinc-800/40 cursor-pointer transition-colors border border-transparent hover:border-zinc-800 md:col-span-2">
          <input
            type="checkbox"
            checked={settings.dailyGreeting}
            onChange={(e) => onToggle('dailyGreeting', e.target.checked)}
            className="mt-0.5 rounded border-zinc-700 bg-zinc-950 text-orange-500 focus:ring-orange-500 cursor-pointer accent-orange-500"
          />
          <div className="flex-1">
            <span className="text-orange-400 font-semibold flex items-center gap-1">
              <SunMedium className="w-3.5 h-3.5" />
              【日常つぶやき・あいさつ枠】優先投稿
            </span>
            <p className="text-zinc-400 text-[11px] mt-0.5">
              朝(7-9時)・昼(12-13:30)・夜(21-23:30)に人間らしい日常文を優先投稿（スピンタックス自動展開・動的日付差し込み・30日間クールダウン）
            </p>
          </div>
        </label>
      </div>
    </div>
  );
};
