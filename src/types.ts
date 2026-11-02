export interface AccountInfo {
  id: string;
  authToken: string;
  userId: string;
  notes: string;
  minIntervalMinutes: number;
  maxIntervalMinutes: number;
  activeHours: string;
  mode: 0 | 1; // 0: 上から順, 1: ランダム
  profileDir: string;
  topics: string[];
  currentTopicIndex: number;
}

export interface PostItem {
  id: number;
  text: string;
  topic?: string;
  imagePath?: string;
  status: '未投稿' | '投稿中' | '投稿完了' | 'エラー';
  completedAt?: string;
  errorMessage?: string;
}

export interface DailyGreetingSlot {
  id: string;
  slotName: string; // '朝枠' | '昼枠' | '夜枠'
  timeRange: string;
  text: string;
  enabled: boolean;
  lastPostedDate?: string;
  lastPostedTime?: string;
}

export interface SystemSettings {
  autoLoop: boolean;
  singleWindowRotation: boolean;
  blockMedia: boolean;
  stealthPatrol: boolean;
  autoCleanProfile: boolean;
  topicAbTest: boolean;
  dailyGreeting: boolean;
  speedMode: 'real' | 'fast' | 'instant';
}

export interface LogMessage {
  id: string;
  timestamp: string;
  level: 'INFO' | 'SUCCESS' | 'WARN' | 'ERROR' | 'STEALTH';
  message: string;
}
