// KATA Navi アプリA データモデル（実装指示書 §4 準拠）

export type WorkType = "作業" | "感覚" | "混在" | "未判定";
export type PlanStatus = "未着手" | "進行中" | "完了";
export type Freq = "日次" | "週次" | "月次" | "不定期";
export type Ecrs = "E" | "C" | "R" | "S";

export interface TaskScores {
  impact: number; // 影響度（既定軸）
  freq: number; // 頻度（既定軸）
  quality?: number; // 任意：品質（Phase2）
  difficulty?: number; // 任意：難易度（Phase2）
}

export interface Task {
  id: string;
  group: string; // 業務名（ワークグループ）★切替・保存の基準
  category: string; // プロセス区分（受注/製造…）
  workType: WorkType;
  content: string; // 作業内容（必須）
  person: string; // 担当（必須）
  time: number; // 所要時間（必須・分）
  freq: Freq; // 頻度（必須）
  freqCount?: number;
  target?: string;
  method?: string;
  tools?: string;
  mmm: { muri: boolean; muda: boolean; mura: boolean }; // ムリムダムラ
  mmmDetails?: { muri?: string; muda?: string; mura?: string };
  problems: string[]; // 問題カテゴリ
  scores: TaskScores; // 既定2軸
  ecrs?: Ecrs; // A-3で選択
}

export interface Plan {
  id: string;
  taskId: string;
  taskName: string;
  content: string; // 必須
  person: string; // 必須
  date: string; // 完了目標日 必須
  method?: string;
  effect?: string;
  status: PlanStatus;
  afterMonthly?: number | null; // A-5 効果検証（改善後の月間時間）
}

export interface ProblemCategory {
  id: string;
  name: string;
  desc?: string;
}

export interface Masters {
  processCategories: string[];
  problemCategories: ProblemCategory[];
}

export interface AppState {
  tasks: Task[];
  plans: Plan[];
  masters: Masters;
  currentGroup: string | null; // null = すべての業務名（全体一覧）
}
