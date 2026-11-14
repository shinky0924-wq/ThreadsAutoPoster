import React from 'react';
import { AccountInfo, PostItem } from '../types';
import { 
  Heart, 
  MessageCircle, 
  Repeat2, 
  Send, 
  MoreHorizontal, 
  Sparkles,
  Clock,
  Shuffle,
  ShieldAlert,
  ArrowRight
} from 'lucide-react';
import { resolveImagePath, parseSpintax } from '../utils/spintax';

interface ThreadsPreviewProps {
  currentAccount: AccountInfo;
  nextPost: PostItem | null;
  isRunning: boolean;
  isStealthing: boolean;
  countdownSeconds: number;
  onPostNow: () => void;
  selectedTopic?: string;
  customImages?: Record<string, string>;
}

export const ThreadsPreview: React.FC<ThreadsPreviewProps> = ({
  currentAccount,
  nextPost,
  isRunning,
  isStealthing,
  countdownSeconds,
  onPostNow,
  selectedTopic,
  customImages,
}) => {
  const formatCountdown = (secs: number) => {
    const m = Math.floor(secs / 60);
    const s = secs % 60;
    return `${m}分${s < 10 ? '0' : ''}${s}秒`;
  };

  const imageSrc = nextPost ? resolveImagePath(nextPost.imagePath, nextPost.text, customImages) : '';
  const displayText = nextPost ? parseSpintax(nextPost.text) : '未投稿のポストがありません。';

  return (
    <div className="bg-zinc-900 border border-zinc-800 rounded-xl p-5 shadow-sm space-y-4">
      {/* Title & Status */}
      <div className="flex items-center justify-between pb-3 border-b border-zinc-800">
        <div className="flex items-center gap-2">
          <div className="w-5 h-5 rounded-full bg-white flex items-center justify-center text-black font-black text-xs">
            @
          </div>
          <h3 className="text-sm font-bold text-white tracking-wide">
            Threads 実機プレビュー & 次回投稿枠
          </h3>
        </div>

        {selectedTopic && (
          <span className="flex items-center gap-1 text-xs px-2.5 py-0.5 rounded-full bg-teal-950 text-teal-300 border border-teal-800 font-mono">
            <Shuffle className="w-3 h-3" />
            次回トピック: {selectedTopic}
          </span>
        )}
      </div>

      {/* Countdown Card (if running) */}
      {isRunning && (
        <div className="p-3 rounded-lg bg-zinc-950 border border-zinc-800 flex items-center justify-between gap-3">
          <div className="flex items-center gap-2.5">
            <div className={`p-2 rounded-lg ${isStealthing ? 'bg-amber-950 text-amber-400' : 'bg-purple-950 text-purple-400'}`}>
              <Clock className="w-4 h-4 animate-spin" />
            </div>
            <div>
              <div className="text-xs font-semibold text-zinc-300">
                {isStealthing ? '🕵️‍♂️ ステルス巡回実行中...' : 'BAN防止待機インターバル中'}
              </div>
              <div className="text-[11px] text-zinc-500 font-mono">
                次回投稿まで残り: <span className="text-emerald-400 font-bold">{formatCountdown(countdownSeconds)}</span>
              </div>
            </div>
          </div>

          <button
            onClick={onPostNow}
            className="flex items-center gap-1 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded-lg shadow cursor-pointer transition-all"
          >
            <span>待機スキップで即時投稿</span>
            <ArrowRight className="w-3 h-3" />
          </button>
        </div>
      )}

      {/* Authentic Threads Post Card */}
      <div className="max-w-md mx-auto bg-black rounded-2xl border border-zinc-800 p-4 shadow-xl text-white font-sans">
        {/* Header: User avatar + handle + menu */}
        <div className="flex items-start justify-between gap-3 mb-2.5">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-full bg-gradient-to-tr from-zinc-700 to-zinc-500 flex items-center justify-center font-bold text-sm text-zinc-200 border border-zinc-700 shadow-sm">
              {currentAccount.id.charAt(0).toUpperCase()}
            </div>
            <div>
              <div className="flex items-center gap-1.5">
                <span className="font-semibold text-sm hover:underline cursor-pointer">
                  {currentAccount.id}
                </span>
                <span className="text-zinc-500 text-xs">・今</span>
              </div>
              <div className="text-[11px] text-zinc-500 font-mono">
                {currentAccount.userId ? `ID: ${currentAccount.userId}` : '認証済アカウント'}
              </div>
            </div>
          </div>
          <button className="text-zinc-500 hover:text-zinc-300">
            <MoreHorizontal className="w-4 h-4" />
          </button>
        </div>

        {/* Post Content */}
        <div className="pl-11 space-y-3">
          <div className="text-sm leading-relaxed whitespace-pre-wrap break-words text-zinc-100">
            {displayText}
          </div>

          {/* Attached Image Preview */}
          {imageSrc && (
            <div className="rounded-xl overflow-hidden border border-zinc-800 bg-zinc-950 relative group">
              <div className="h-48 w-full bg-gradient-to-br from-zinc-900 via-zinc-950 to-zinc-900 flex items-center justify-center relative">
                <img
                  src={imageSrc}
                  alt="Post Attachment"
                  className="max-h-full max-w-full object-contain"
                  onError={(e) => {
                    // Fallback to placeholder if local image isn't loaded
                    (e.target as HTMLElement).style.display = 'none';
                  }}
                />
                <div className="absolute bottom-2 left-2 px-2 py-0.5 rounded bg-black/70 text-[10px] font-mono text-zinc-300 backdrop-blur-xs flex items-center gap-1">
                  <span>📎 {nextPost?.imagePath || 'image.png'}</span>
                </div>
              </div>
            </div>
          )}

          {/* Threads Action Icons */}
          <div className="flex items-center gap-4 text-zinc-400 pt-1">
            <button className="hover:text-rose-500 transition-colors cursor-pointer">
              <Heart className="w-5 h-5" />
            </button>
            <button className="hover:text-sky-400 transition-colors cursor-pointer">
              <MessageCircle className="w-5 h-5" />
            </button>
            <button className="hover:text-emerald-400 transition-colors cursor-pointer">
              <Repeat2 className="w-5 h-5" />
            </button>
            <button className="hover:text-zinc-200 transition-colors cursor-pointer">
              <Send className="w-5 h-5" />
            </button>
          </div>

          {/* Bottom thread preview line */}
          <div className="text-[11px] text-zinc-500 flex items-center gap-1.5 pt-1">
            <span className="w-1.5 h-1.5 rounded-full bg-zinc-600" />
            <span>返信0件・いいね0件</span>
          </div>
        </div>
      </div>
    </div>
  );
};
