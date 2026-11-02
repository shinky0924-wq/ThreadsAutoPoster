import React, { useState } from 'react';
import { DailyGreetingSlot } from '../types';
import { Sun, Coffee, Moon, Sparkles, Plus, Trash2, CheckCircle2 } from 'lucide-react';
import { parseSpintax } from '../utils/spintax';

interface DailyGreetingsProps {
  slots: DailyGreetingSlot[];
  onUpdateSlot: (updated: DailyGreetingSlot) => void;
  onAddSlot: (slot: Omit<DailyGreetingSlot, 'id'>) => void;
  onDeleteSlot: (id: string) => void;
}

export const DailyGreetings: React.FC<DailyGreetingsProps> = ({
  slots,
  onUpdateSlot,
  onAddSlot,
  onDeleteSlot,
}) => {
  const [spintaxInput, setSpintaxInput] = useState('{おはようございます☀️|皆さんおはようございます！☕️} 今日もいい天気！頑張りましょう〜\n\n素敵な1日になりますように✨');
  const [spintaxResult, setSpintaxResult] = useState('');
  const [isAdding, setIsAdding] = useState(false);
  const [newSlotName, setNewSlotName] = useState('朝枠');
  const [newTimeRange, setNewTimeRange] = useState('7:00-9:00');
  const [newText, setNewText] = useState('');

  const handleTestSpintax = () => {
    setSpintaxResult(parseSpintax(spintaxInput));
  };

  const handleCreateSlot = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newText.trim()) return;
    onAddSlot({
      slotName: newSlotName,
      timeRange: newTimeRange,
      text: newText.trim(),
      enabled: true,
    });
    setNewText('');
    setIsAdding(false);
  };

  const getSlotIcon = (name: string) => {
    if (name.includes('朝')) return <Sun className="w-4 h-4 text-amber-400" />;
    if (name.includes('昼')) return <Coffee className="w-4 h-4 text-orange-400" />;
    return <Moon className="w-4 h-4 text-indigo-400" />;
  };

  return (
    <div className="bg-zinc-900 border border-zinc-800 rounded-xl p-5 shadow-sm space-y-5">
      <div className="flex items-center justify-between pb-3 border-b border-zinc-800">
        <div>
          <h3 className="text-sm font-bold text-white flex items-center gap-2">
            <Sun className="w-4 h-4 text-amber-400" />
            日常つぶやき・あいさつ枠 (BAN防止最優先投稿枠)
          </h3>
          <p className="text-xs text-zinc-400 mt-0.5">
            朝・昼・夕・夜に人間らしい日常のつぶやきを自動投稿。30日間クールダウン & スピンタックスで凍結を完全回避
          </p>
        </div>

        <button
          onClick={() => setIsAdding(!isAdding)}
          className="flex items-center gap-1 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded-lg shadow cursor-pointer transition-all"
        >
          <Plus className="w-3.5 h-3.5" />
          <span>挨拶枠追加</span>
        </button>
      </div>

      {/* Add Slot Form */}
      {isAdding && (
        <form onSubmit={handleCreateSlot} className="p-4 bg-zinc-950 rounded-lg border border-zinc-800 space-y-3">
          <div className="grid grid-cols-2 gap-3 text-xs">
            <div>
              <label className="text-zinc-400 block mb-1">枠名</label>
              <select
                value={newSlotName}
                onChange={(e) => {
                  setNewSlotName(e.target.value);
                  if (e.target.value === '朝枠') setNewTimeRange('7:00-9:00');
                  else if (e.target.value === '昼枠') setNewTimeRange('12:00-13:30');
                  else setNewTimeRange('21:00-23:30');
                }}
                className="w-full bg-zinc-900 border border-zinc-700 text-white text-xs px-2.5 py-1.5 rounded focus:outline-none"
              >
                <option value="朝枠">朝枠 (7:00-9:00)</option>
                <option value="昼枠">昼枠 (12:00-13:30)</option>
                <option value="夜枠">夜枠 (21:00-23:30)</option>
              </select>
            </div>
            <div>
              <label className="text-zinc-400 block mb-1">実行時間帯</label>
              <input
                type="text"
                value={newTimeRange}
                onChange={(e) => setNewTimeRange(e.target.value)}
                className="w-full bg-zinc-900 border border-zinc-700 text-white text-xs px-2.5 py-1.5 rounded focus:outline-none font-mono"
              />
            </div>
          </div>
          <div>
            <label className="text-zinc-400 text-xs block mb-1">
              投稿テキスト (スピンタックス {'{選択A|選択B}'} や {'{today}'} 対応)
            </label>
            <textarea
              value={newText}
              onChange={(e) => setNewText(e.target.value)}
              rows={3}
              placeholder="{おはようございます|皆さんおはよう} 今日も元気にいきましょう！"
              className="w-full bg-zinc-900 border border-zinc-700 text-white text-xs p-2.5 rounded-lg focus:outline-none focus:border-emerald-500"
              required
            />
          </div>
          <div className="flex justify-end gap-2">
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
              登録する
            </button>
          </div>
        </form>
      )}

      {/* Slots List */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
        {slots.map((slot) => (
          <div
            key={slot.id}
            className={`p-3.5 rounded-xl border transition-all ${
              slot.enabled
                ? 'bg-zinc-950/80 border-zinc-800 hover:border-zinc-700'
                : 'bg-zinc-950/40 border-zinc-900 opacity-60'
            }`}
          >
            <div className="flex items-center justify-between pb-2 mb-2 border-b border-zinc-800/80">
              <div className="flex items-center gap-2">
                {getSlotIcon(slot.slotName)}
                <span className="font-bold text-xs text-white">{slot.slotName}</span>
                <span className="text-[11px] font-mono px-2 py-0.5 rounded bg-zinc-900 text-zinc-400 border border-zinc-800">
                  {slot.timeRange}
                </span>
              </div>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => onUpdateSlot({ ...slot, enabled: !slot.enabled })}
                  className={`px-2 py-0.5 rounded text-[10px] font-semibold transition-colors cursor-pointer ${
                    slot.enabled
                      ? 'bg-emerald-950 text-emerald-300 border border-emerald-800'
                      : 'bg-zinc-800 text-zinc-500'
                  }`}
                >
                  {slot.enabled ? '有効' : '無効'}
                </button>
                <button
                  onClick={() => onDeleteSlot(slot.id)}
                  className="p-1 text-zinc-600 hover:text-rose-400 transition-colors cursor-pointer"
                  title="削除"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                </button>
              </div>
            </div>

            <p className="text-xs text-zinc-300 whitespace-pre-wrap leading-relaxed">
              {slot.text}
            </p>

            <div className="mt-2.5 pt-2 border-t border-zinc-800/60 flex items-center justify-between text-[10px] text-zinc-500 font-mono">
              <span>最終投稿: {slot.lastPostedTime || '未実行'}</span>
              <span>30日クールダウン保護済</span>
            </div>
          </div>
        ))}
      </div>

      {/* Spintax Tester Box */}
      <div className="bg-zinc-950 p-4 rounded-xl border border-zinc-800 space-y-3">
        <div className="flex items-center justify-between">
          <span className="text-xs font-bold text-zinc-300 flex items-center gap-1.5">
            <Sparkles className="w-3.5 h-3.5 text-sky-400" />
            スピンタックス & 動的日付 リアルタイムテスター
          </span>
          <button
            onClick={handleTestSpintax}
            className="flex items-center gap-1 px-3 py-1 bg-sky-600 hover:bg-sky-500 text-white text-xs font-semibold rounded cursor-pointer"
          >
            <span>展開テスト</span>
          </button>
        </div>

        <textarea
          value={spintaxInput}
          onChange={(e) => setSpintaxInput(e.target.value)}
          rows={2}
          className="w-full bg-zinc-900 border border-zinc-700 text-white text-xs p-2.5 rounded-lg focus:outline-none font-mono"
        />

        {spintaxResult && (
          <div className="p-2.5 rounded bg-zinc-900/90 border border-sky-800/60 text-xs text-sky-200 whitespace-pre-wrap">
            <span className="text-[10px] text-sky-400 uppercase font-mono block mb-1">展開後のプレビュー:</span>
            {spintaxResult}
          </div>
        )}
      </div>
    </div>
  );
};
