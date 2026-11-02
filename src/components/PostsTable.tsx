import React, { useState } from 'react';
import { PostItem } from '../types';
import { 
  CheckCircle2, 
  Clock, 
  AlertCircle, 
  Search, 
  Plus, 
  Trash2, 
  Send, 
  Image as ImageIcon,
  ExternalLink,
  Edit2,
  FileText,
  Sparkles,
  Layers,
  X,
  AlertTriangle
} from 'lucide-react';
import { resolveImagePath } from '../utils/spintax';

interface PostsTableProps {
  posts: PostItem[];
  currentAccountId: string;
  onUpdatePost: (updated: PostItem) => void;
  onDeletePost: (id: number) => void;
  onAddPost: (post: Omit<PostItem, 'id'>) => void;
  onPostSingle: (post: PostItem) => void;
  onClearAllPosts: () => void;
  onBulkAddPosts: (newPosts: Omit<PostItem, 'id'>[], replaceExisting: boolean) => void;
}

export const PostsTable: React.FC<PostsTableProps> = ({
  posts,
  currentAccountId,
  onUpdatePost,
  onDeletePost,
  onAddPost,
  onPostSingle,
  onClearAllPosts,
  onBulkAddPosts,
}) => {
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState<'all' | '未投稿' | '投稿完了' | 'エラー'>('all');
  
  // Single Add
  const [isAdding, setIsAdding] = useState(false);
  const [newText, setNewText] = useState('');
  const [newTopic, setNewTopic] = useState('信太山新地');
  const [newImagePath, setNewImagePath] = useState('image.png');
  const [editingPost, setEditingPost] = useState<PostItem | null>(null);

  // Bulk Add Modal
  const [isBulkAdding, setIsBulkAdding] = useState(false);
  const [bulkTextInput, setBulkTextInput] = useState('');
  const [bulkTopic, setBulkTopic] = useState('信太山新地');
  const [bulkImageOption, setBulkImageOption] = useState<'none' | 'alternate' | 'image1' | 'image2'>('alternate');
  const [bulkReplace, setBulkReplace] = useState(false);

  // Clear confirmation modal
  const [showClearConfirm, setShowClearConfirm] = useState(false);

  // Parse bulk text input to count items
  const parsedBulkItems = bulkTextInput
    .split(/\n\s*\n|\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  const filteredPosts = posts.filter((p) => {
    if (statusFilter !== 'all' && p.status !== statusFilter) return false;
    if (search.trim()) {
      const q = search.toLowerCase();
      return (
        p.text.toLowerCase().includes(q) ||
        (p.topic && p.topic.toLowerCase().includes(q)) ||
        (p.imagePath && p.imagePath.toLowerCase().includes(q))
      );
    }
    return true;
  });

  const countTotal = posts.length;
  const countUnposted = posts.filter((p) => p.status === '未投稿').length;
  const countCompleted = posts.filter((p) => p.status === '投稿完了').length;
  const countError = posts.filter((p) => p.status === 'エラー').length;

  const handleCreatePost = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newText.trim()) return;
    onAddPost({
      text: newText.trim(),
      topic: newTopic.trim(),
      imagePath: newImagePath.trim(),
      status: '未投稿',
    });
    setNewText('');
    setIsAdding(false);
  };

  const handleSaveEdit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingPost) return;
    onUpdatePost(editingPost);
    setEditingPost(null);
  };

  const handleExecuteBulkAdd = () => {
    if (parsedBulkItems.length === 0) return;

    const newItems: Omit<PostItem, 'id'>[] = parsedBulkItems.map((text, idx) => {
      let imagePath = '';
      if (bulkImageOption === 'image1') imagePath = 'image.png';
      else if (bulkImageOption === 'image2') imagePath = 'image2.png';
      else if (bulkImageOption === 'alternate') {
        imagePath = idx % 3 === 0 ? 'image.png' : idx % 3 === 1 ? 'image2.png' : '';
      }

      return {
        text,
        topic: bulkTopic.trim() || '信太山新地',
        imagePath,
        status: '未投稿',
      };
    });

    onBulkAddPosts(newItems, bulkReplace);
    setBulkTextInput('');
    setIsBulkAdding(false);
  };

  // Helper: Generate 100 sample posts
  const handleGenerate100Samples = () => {
    const topics = ['信太山新地', '大阪求人', '高収入バイト', 'ナイトワーク安心ナビ'];
    const hooks = [
      '【完全保証】安心して働ける環境を徹底紹介！',
      '高収入とプライバシー保護を両立できる働き方ガイド✨',
      '未経験から始めるための面接・体験入社Q&Aまとめ🍙',
      'シフトの自由度と給与システムを徹底比較！',
      '寮完備・即日入居サポートありの人気店舗特集🏢',
      '働きやすい環境づくりとスタッフのサポート体制について🍵',
      '短期集中・週末のみでもOK！自分に合ったワークスタイル見学🌿',
      'よくある質問と安心ルールの徹底解説まとめ💡',
    ];
    const urls = [
      'https://osaka-shinodayamashinchivip.pages.dev/blog/osaka-shinodayama-high-guarantee-guide',
      'https://osaka-shinodayamashinchivip.pages.dev/blog/shinodayamashinchi-salary-system',
      'https://osaka-shinodayamashinchivip.pages.dev/blog/shinodayamashinchi-privacy-guide',
      'https://osaka-shinodayamashinchivip.pages.dev/blog/shinodayamashinchi-housing-support',
      'https://osaka-shinodayamashinchivip.pages.dev/blog/shinodayamashinchi-interview-guide',
      'https://osaka-shinodayamashinchivip.pages.dev/blog/shinodayamashinchi-trial-guide',
    ];

    const lines: string[] = [];
    for (let i = 1; i <= 100; i++) {
      const topic = topics[(i - 1) % topics.length];
      const hook = hooks[(i - 1) % hooks.length];
      const url = urls[(i - 1) % urls.length];
      lines.push(`${topic} #${i}\n${hook}\n${url}`);
    }
    setBulkTextInput(lines.join('\n\n'));
  };

  return (
    <div className="bg-zinc-900 border border-zinc-800 rounded-xl overflow-hidden shadow-sm">
      {/* Top Controls & Metrics Bar */}
      <div className="p-4 border-b border-zinc-800 flex flex-wrap items-center justify-between gap-3">
        {/* Metric Badges */}
        <div className="flex items-center gap-2 flex-wrap text-xs">
          <button
            onClick={() => setStatusFilter('all')}
            className={`px-3 py-1 rounded-md font-medium transition-colors cursor-pointer ${
              statusFilter === 'all'
                ? 'bg-zinc-800 text-white border border-zinc-700'
                : 'text-zinc-400 hover:text-white'
            }`}
          >
            すべて: <span className="font-bold text-white">{countTotal}</span>
          </button>
          <button
            onClick={() => setStatusFilter('未投稿')}
            className={`px-3 py-1 rounded-md font-medium transition-colors cursor-pointer ${
              statusFilter === '未投稿'
                ? 'bg-sky-950/80 text-sky-300 border border-sky-800'
                : 'text-zinc-400 hover:text-sky-300'
            }`}
          >
            未投稿: <span className="font-bold text-sky-400">{countUnposted}</span>
          </button>
          <button
            onClick={() => setStatusFilter('投稿完了')}
            className={`px-3 py-1 rounded-md font-medium transition-colors cursor-pointer ${
              statusFilter === '投稿完了'
                ? 'bg-emerald-950/80 text-emerald-300 border border-emerald-800'
                : 'text-zinc-400 hover:text-emerald-300'
            }`}
          >
            投稿完了: <span className="font-bold text-emerald-400">{countCompleted}</span>
          </button>
          {countError > 0 && (
            <button
              onClick={() => setStatusFilter('エラー')}
              className={`px-3 py-1 rounded-md font-medium transition-colors cursor-pointer ${
                statusFilter === 'エラー'
                  ? 'bg-rose-950/80 text-rose-300 border border-rose-800'
                  : 'text-zinc-400 hover:text-rose-300'
              }`}
            >
              エラー: <span className="font-bold text-rose-400">{countError}</span>
            </button>
          )}
        </div>

        {/* Action Buttons & Search */}
        <div className="flex items-center gap-2 flex-wrap flex-1 max-w-xl justify-end">
          <div className="relative flex-1 min-w-[140px] max-w-xs">
            <Search className="w-3.5 h-3.5 text-zinc-500 absolute left-2.5 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              placeholder="本文やトピックを検索..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="w-full bg-zinc-950 border border-zinc-800 text-xs text-white pl-8 pr-3 py-1.5 rounded-lg focus:outline-none focus:border-emerald-500"
            />
          </div>

          {/* Bulk Add Button */}
          <button
            onClick={() => setIsBulkAdding(true)}
            className="flex items-center gap-1.5 px-3 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold rounded-lg shadow cursor-pointer transition-all"
            title="100件以上のツイートをテキストエリアから一括コピペ登録"
          >
            <Layers className="w-3.5 h-3.5" />
            <span>100件一括追加</span>
          </button>

          {/* Add 1 Post Button */}
          <button
            onClick={() => setIsAdding(!isAdding)}
            className="flex items-center gap-1 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded-lg shadow cursor-pointer transition-all"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>1件追加</span>
          </button>

          {/* Clear All Posts Button */}
          {posts.length > 0 && (
            <button
              onClick={() => setShowClearConfirm(true)}
              className="flex items-center gap-1 px-2.5 py-1.5 bg-rose-950/60 hover:bg-rose-900/60 text-rose-300 border border-rose-800/60 text-xs font-medium rounded-lg cursor-pointer transition-colors"
              title="選択中のアカウントのツイートを全件削除"
            >
              <Trash2 className="w-3.5 h-3.5" />
              <span>全件削除</span>
            </button>
          )}
        </div>
      </div>

      {/* Clear Confirmation Modal */}
      {showClearConfirm && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-zinc-900 border border-rose-800 rounded-2xl max-w-md w-full p-6 space-y-4 shadow-2xl text-xs">
            <div className="flex items-center gap-3 pb-2 border-b border-zinc-800 text-rose-400">
              <div className="p-2.5 rounded-xl bg-rose-950/80 border border-rose-800">
                <AlertTriangle className="w-6 h-6" />
              </div>
              <div>
                <h3 className="text-sm font-bold text-white">既存ツイートを全件削除しますか？</h3>
                <p className="text-zinc-400 text-[11px]">
                  対象: <span className="font-mono text-emerald-400">@{currentAccountId}</span> (登録 {posts.length} 件)
                </p>
              </div>
            </div>

            <p className="text-zinc-300 leading-relaxed">
              現在登録されている <span className="text-rose-400 font-bold">{posts.length} 件</span> の投稿リストがすべて消去されます。
              削除後に「100件一括追加」で新しいツイート群に入れ替えることができます。
            </p>

            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-800">
              <button
                type="button"
                onClick={() => setShowClearConfirm(false)}
                className="px-4 py-2 text-zinc-400 hover:text-white cursor-pointer"
              >
                キャンセル
              </button>
              <button
                type="button"
                onClick={() => {
                  onClearAllPosts();
                  setShowClearConfirm(false);
                }}
                className="px-5 py-2 bg-rose-600 hover:bg-rose-500 text-white font-semibold rounded-lg shadow cursor-pointer"
              >
                全件削除する
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Bulk Add Modal (100+ Posts) */}
      {isBulkAdding && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <div className="bg-zinc-900 border border-zinc-700 rounded-2xl max-w-2xl w-full p-6 space-y-4 shadow-2xl text-xs relative max-h-[92vh] flex flex-col">
            <button
              onClick={() => setIsBulkAdding(false)}
              className="absolute top-4 right-4 text-zinc-400 hover:text-white"
            >
              <X className="w-5 h-5" />
            </button>

            <div className="flex items-center justify-between pb-2 border-b border-zinc-800">
              <div className="flex items-center gap-2">
                <div className="p-2 rounded-xl bg-indigo-950 text-indigo-400 border border-indigo-800">
                  <Layers className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-sm font-bold text-white">
                    新規ツイート一括登録 (100件以上対応)
                  </h3>
                  <p className="text-zinc-400 text-[11px]">
                    対象アカウント: <span className="font-mono text-emerald-400">@{currentAccountId}</span>
                  </p>
                </div>
              </div>

              <button
                type="button"
                onClick={handleGenerate100Samples}
                className="flex items-center gap-1 px-3 py-1.5 bg-zinc-800 hover:bg-zinc-700 text-amber-300 border border-amber-600/50 rounded-lg cursor-pointer"
                title="テスト用に100件のユニーク投稿を自動生成"
              >
                <Sparkles className="w-3.5 h-3.5" />
                <span>100件サンプル自動入力</span>
              </button>
            </div>

            <div className="space-y-1">
              <div className="flex items-center justify-between">
                <label className="text-zinc-300 font-semibold">
                  投稿テキスト貼り付け (空行または改行区切りで1ポストずつ認識):
                </label>
                <span className="font-mono text-indigo-300 font-bold bg-indigo-950/80 px-2.5 py-0.5 rounded border border-indigo-800">
                  検出件数: {parsedBulkItems.length} 件
                </span>
              </div>
              <textarea
                value={bulkTextInput}
                onChange={(e) => setBulkTextInput(e.target.value)}
                rows={9}
                placeholder="ここにツイート本文を100個コピペしてください。&#10;&#10;例:&#10;信太山新地 #1&#10;https://example.com/guide&#10;&#10;信太山新地 #2&#10;https://example.com/salary&#10;&#10;信太山新地 #3&#10;https://example.com/privacy"
                className="w-full bg-zinc-950 border border-zinc-700 text-white p-3 rounded-lg focus:outline-none focus:border-indigo-500 font-mono text-xs leading-relaxed"
              />
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div>
                <label className="text-zinc-400 block mb-1">一括設定トピック (カテゴリ)</label>
                <input
                  type="text"
                  value={bulkTopic}
                  onChange={(e) => setBulkTopic(e.target.value)}
                  placeholder="例: 信太山新地"
                  className="w-full bg-zinc-950 border border-zinc-700 text-white px-2.5 py-1.5 rounded focus:outline-none"
                />
              </div>

              <div>
                <label className="text-zinc-400 block mb-1">画像添付ルール</label>
                <select
                  value={bulkImageOption}
                  onChange={(e) => setBulkImageOption(e.target.value as any)}
                  className="w-full bg-zinc-950 border border-zinc-700 text-white px-2.5 py-1.5 rounded focus:outline-none"
                >
                  <option value="alternate">image.png と image2.png を交互・ランダムに添付</option>
                  <option value="image1">すべて image.png を添付</option>
                  <option value="image2">すべて image2.png を添付</option>
                  <option value="none">添付なし (テキストのみ)</option>
                </select>
              </div>
            </div>

            <div className="p-3 rounded-lg bg-zinc-950 border border-zinc-800 flex items-center justify-between">
              <label className="flex items-center gap-2 cursor-pointer">
                <input
                  type="checkbox"
                  checked={bulkReplace}
                  onChange={(e) => setBulkReplace(e.target.checked)}
                  className="accent-rose-500 rounded"
                />
                <span className="text-zinc-300 font-medium">
                  既存のツイート ({posts.length} 件) を削除して、この {parsedBulkItems.length} 件に完全入れ替えする
                </span>
              </label>
            </div>

            <div className="flex items-center justify-end gap-2 pt-2 border-t border-zinc-800">
              <button
                type="button"
                onClick={() => setIsBulkAdding(false)}
                className="px-4 py-2 text-zinc-400 hover:text-white cursor-pointer"
              >
                キャンセル
              </button>
              <button
                type="button"
                onClick={handleExecuteBulkAdd}
                disabled={parsedBulkItems.length === 0}
                className="px-5 py-2 bg-indigo-600 hover:bg-indigo-500 disabled:opacity-40 text-white font-semibold rounded-lg shadow cursor-pointer transition-all"
              >
                {parsedBulkItems.length} 件を一括追加
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Add Single Post Form */}
      {isAdding && (
        <form onSubmit={handleCreatePost} className="p-4 bg-zinc-950/90 border-b border-zinc-800 space-y-3">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-emerald-400">新しい投稿を1件追加</span>
            <span className={`text-[11px] font-mono ${newText.length > 500 ? 'text-rose-400 font-bold' : 'text-zinc-400'}`}>
              {newText.length} / 500文字
            </span>
          </div>
          <textarea
            value={newText}
            onChange={(e) => setNewText(e.target.value)}
            rows={3}
            placeholder="Threads投稿文章 (最大500文字)。URLや {おはよう|こんにちは} のスピンタックス、image2.pngの記載もOK"
            className="w-full bg-zinc-900 border border-zinc-700 text-white text-xs p-2.5 rounded-lg focus:outline-none focus:border-emerald-500"
            required
          />
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
            <div>
              <label className="text-zinc-400 block mb-1">トピック / カテゴリ</label>
              <input
                type="text"
                value={newTopic}
                onChange={(e) => setNewTopic(e.target.value)}
                placeholder="例: 信太山新地"
                className="w-full bg-zinc-900 border border-zinc-700 text-white text-xs px-2.5 py-1.5 rounded focus:outline-none"
              />
            </div>
            <div>
              <label className="text-zinc-400 block mb-1">添付画像 (image.png または画像URL)</label>
              <select
                value={newImagePath}
                onChange={(e) => setNewImagePath(e.target.value)}
                className="w-full bg-zinc-900 border border-zinc-700 text-white text-xs px-2.5 py-1.5 rounded focus:outline-none"
              >
                <option value="">添付なし</option>
                <option value="image.png">image.png (サンプル1)</option>
                <option value="image2.png">image2.png (サンプル2)</option>
              </select>
            </div>
          </div>
          <div className="flex justify-end gap-2 pt-1">
            <button
              type="button"
              onClick={() => setIsAdding(false)}
              className="px-3 py-1.5 text-zinc-400 hover:text-white text-xs cursor-pointer"
            >
              キャンセル
            </button>
            <button
              type="submit"
              className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded cursor-pointer"
            >
              投稿リストに追加
            </button>
          </div>
        </form>
      )}

      {/* Edit Post Modal / Inline */}
      {editingPost && (
        <div className="fixed inset-0 bg-black/70 backdrop-blur-xs flex items-center justify-center p-4 z-50">
          <form onSubmit={handleSaveEdit} className="bg-zinc-900 border border-zinc-700 rounded-xl p-5 max-w-lg w-full space-y-3 shadow-2xl">
            <div className="flex items-center justify-between border-b border-zinc-800 pb-2">
              <span className="text-sm font-bold text-white">投稿 #{editingPost.id} を編集</span>
              <span className={`text-xs font-mono ${editingPost.text.length > 500 ? 'text-rose-400 font-bold' : 'text-zinc-400'}`}>
                {editingPost.text.length} / 500文字
              </span>
            </div>
            <textarea
              value={editingPost.text}
              onChange={(e) => setEditingPost({ ...editingPost, text: e.target.value })}
              rows={4}
              className="w-full bg-zinc-950 border border-zinc-700 text-white text-xs p-2.5 rounded-lg focus:outline-none focus:border-emerald-500"
              required
            />
            <div className="grid grid-cols-2 gap-3 text-xs">
              <div>
                <label className="text-zinc-400 block mb-1">トピック</label>
                <input
                  type="text"
                  value={editingPost.topic || ''}
                  onChange={(e) => setEditingPost({ ...editingPost, topic: e.target.value })}
                  className="w-full bg-zinc-950 border border-zinc-700 text-white text-xs px-2.5 py-1.5 rounded focus:outline-none"
                />
              </div>
              <div>
                <label className="text-zinc-400 block mb-1">添付画像</label>
                <input
                  type="text"
                  value={editingPost.imagePath || ''}
                  placeholder="image.png または URL"
                  onChange={(e) => setEditingPost({ ...editingPost, imagePath: e.target.value })}
                  className="w-full bg-zinc-950 border border-zinc-700 text-white text-xs px-2.5 py-1.5 rounded focus:outline-none"
                />
              </div>
            </div>
            <div>
              <label className="text-zinc-400 text-xs block mb-1">ステータス</label>
              <select
                value={editingPost.status}
                onChange={(e) => setEditingPost({ ...editingPost, status: e.target.value as PostItem['status'] })}
                className="w-full bg-zinc-950 border border-zinc-700 text-white text-xs px-2.5 py-1.5 rounded focus:outline-none"
              >
                <option value="未投稿">未投稿</option>
                <option value="投稿完了">投稿完了</option>
                <option value="エラー">エラー</option>
              </select>
            </div>
            <div className="flex justify-end gap-2 pt-2 border-t border-zinc-800">
              <button
                type="button"
                onClick={() => setEditingPost(null)}
                className="px-3 py-1.5 text-zinc-400 hover:text-white text-xs cursor-pointer"
              >
                キャンセル
              </button>
              <button
                type="submit"
                className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded cursor-pointer"
              >
                変更を反映
              </button>
            </div>
          </form>
        </div>
      )}

      {/* Table */}
      <div className="overflow-x-auto max-h-[520px]">
        <table className="w-full text-left text-xs border-collapse">
          <thead className="bg-zinc-950 text-zinc-400 font-semibold sticky top-0 z-10 border-b border-zinc-800">
            <tr>
              <th className="py-2.5 px-3 w-12 text-center">ID</th>
              <th className="py-2.5 px-3">ポスト文章 / 投稿内容</th>
              <th className="py-2.5 px-3 w-28">トピック</th>
              <th className="py-2.5 px-3 w-24">画像</th>
              <th className="py-2.5 px-3 w-28">ステータス</th>
              <th className="py-2.5 px-3 w-36">完了日時</th>
              <th className="py-2.5 px-3 w-28 text-right">操作</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-zinc-800/60 font-sans">
            {filteredPosts.length === 0 ? (
              <tr>
                <td colSpan={7} className="py-12 text-center text-zinc-500">
                  <div className="max-w-sm mx-auto space-y-2">
                    <p className="font-medium text-zinc-400">登録されているツイートはありません。</p>
                    <p className="text-[11px] text-zinc-500">
                      「100件一括追加」ボタンを押して新規テキストを一発コピペ投入するか、「1件追加」をクリックしてください。
                    </p>
                  </div>
                </td>
              </tr>
            ) : (
              filteredPosts.map((post) => {
                const img = resolveImagePath(post.imagePath, post.text);
                const isUnposted = post.status === '未投稿';
                const isCompleted = post.status === '投稿完了';
                const isError = post.status === 'エラー';

                return (
                  <tr
                    key={post.id}
                    className="hover:bg-zinc-800/30 transition-colors group"
                  >
                    <td className="py-2.5 px-3 text-center font-mono text-zinc-400 font-medium">
                      #{post.id}
                    </td>

                    {/* Post Text */}
                    <td className="py-2.5 px-3 text-zinc-200">
                      <div className="line-clamp-2 font-medium max-w-md break-all">
                        {post.text}
                      </div>
                      <div className="flex items-center gap-2 mt-0.5 text-[10px] text-zinc-500 font-mono">
                        <span>{post.text.length} 文字</span>
                        {post.text.includes('http') && (
                          <span className="flex items-center gap-0.5 text-sky-400">
                            <ExternalLink className="w-2.5 h-2.5" /> リンク有
                          </span>
                        )}
                      </div>
                    </td>

                    {/* Topic */}
                    <td className="py-2.5 px-3 text-zinc-400">
                      {post.topic ? (
                        <span className="px-2 py-0.5 rounded bg-zinc-800 text-teal-300 font-mono text-[11px] border border-zinc-700/60">
                          {post.topic}
                        </span>
                      ) : (
                        <span className="text-zinc-600">―</span>
                      )}
                    </td>

                    {/* Image */}
                    <td className="py-2.5 px-3">
                      {img ? (
                        <div className="flex items-center gap-1 text-[11px] text-violet-300">
                          <ImageIcon className="w-3.5 h-3.5 text-violet-400" />
                          <span className="font-mono truncate max-w-[80px]">
                            {post.imagePath || '添付有'}
                          </span>
                        </div>
                      ) : (
                        <span className="text-zinc-600">なし</span>
                      )}
                    </td>

                    {/* Status Badge */}
                    <td className="py-2.5 px-3">
                      <button
                        onClick={() =>
                          onUpdatePost({
                            ...post,
                            status: isUnposted ? '投稿完了' : '未投稿',
                            completedAt: isUnposted ? new Date().toLocaleString('ja-JP') : undefined,
                          })
                        }
                        className="cursor-pointer"
                        title="クリックでステータスを切替"
                      >
                        {isUnposted && (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-sky-950/80 text-sky-300 border border-sky-800 text-[11px]">
                            <Clock className="w-3 h-3" /> 未投稿
                          </span>
                        )}
                        {isCompleted && (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-950/80 text-emerald-300 border border-emerald-800 text-[11px]">
                            <CheckCircle2 className="w-3 h-3" /> 投稿完了
                          </span>
                        )}
                        {isError && (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-rose-950/80 text-rose-300 border border-rose-800 text-[11px]">
                            <AlertCircle className="w-3 h-3" /> エラー
                          </span>
                        )}
                        {post.status === '投稿中' && (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-amber-950/80 text-amber-300 border border-amber-800 text-[11px] animate-pulse">
                            投稿中...
                          </span>
                        )}
                      </button>
                    </td>

                    {/* Completion Date */}
                    <td className="py-2.5 px-3 text-[11px] font-mono text-zinc-400">
                      {post.completedAt || '―'}
                    </td>

                    {/* Actions */}
                    <td className="py-2.5 px-3 text-right">
                      <div className="flex items-center justify-end gap-1">
                        <button
                          onClick={() => onPostSingle(post)}
                          className="p-1 rounded text-zinc-400 hover:text-emerald-400 hover:bg-zinc-800 transition-colors cursor-pointer"
                          title="この投稿を単体でテスト実行"
                        >
                          <Send className="w-3.5 h-3.5" />
                        </button>
                        <button
                          onClick={() => setEditingPost(post)}
                          className="p-1 rounded text-zinc-400 hover:text-sky-400 hover:bg-zinc-800 transition-colors cursor-pointer"
                          title="編集"
                        >
                          <Edit2 className="w-3.5 h-3.5" />
                        </button>
                        <button
                          onClick={() => onDeletePost(post.id)}
                          className="p-1 rounded text-zinc-400 hover:text-rose-400 hover:bg-zinc-800 transition-colors cursor-pointer"
                          title="削除"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
};
