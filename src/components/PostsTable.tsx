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
  Edit2
} from 'lucide-react';
import { resolveImagePath } from '../utils/spintax';

interface PostsTableProps {
  posts: PostItem[];
  onUpdatePost: (updated: PostItem) => void;
  onDeletePost: (id: number) => void;
  onAddPost: (post: Omit<PostItem, 'id'>) => void;
  onPostSingle: (post: PostItem) => void;
}

export const PostsTable: React.FC<PostsTableProps> = ({
  posts,
  onUpdatePost,
  onDeletePost,
  onAddPost,
  onPostSingle,
}) => {
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState<'all' | '未投稿' | '投稿完了' | 'エラー'>('all');
  const [isAdding, setIsAdding] = useState(false);
  const [newText, setNewText] = useState('');
  const [newTopic, setNewTopic] = useState('信太山新地');
  const [newImagePath, setNewImagePath] = useState('image.png');
  const [editingPost, setEditingPost] = useState<PostItem | null>(null);

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

        {/* Search & Add Post Button */}
        <div className="flex items-center gap-2 flex-1 max-w-md justify-end">
          <div className="relative flex-1 max-w-xs">
            <Search className="w-3.5 h-3.5 text-zinc-500 absolute left-2.5 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              placeholder="本文やトピックを検索..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="w-full bg-zinc-950 border border-zinc-800 text-xs text-white pl-8 pr-3 py-1.5 rounded-lg focus:outline-none focus:border-emerald-500"
            />
          </div>

          <button
            onClick={() => setIsAdding(!isAdding)}
            className="flex items-center gap-1 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded-lg shadow cursor-pointer transition-all"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>新規投稿</span>
          </button>
        </div>
      </div>

      {/* Add Post Form */}
      {isAdding && (
        <form onSubmit={handleCreatePost} className="p-4 bg-zinc-950/90 border-b border-zinc-800 space-y-3">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-emerald-400">新しい投稿を追加</span>
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
                <td colSpan={7} className="py-8 text-center text-zinc-500">
                  該当する投稿が見つかりませんでした。
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
