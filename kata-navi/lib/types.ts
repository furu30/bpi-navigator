// KATA Navi アプリA データモデル（実装指示書 §4 準拠）

export type WorkType = "作業" | "感覚" | "混在" | "未判定";
export type PlanStatus = "未着手" | "進行中" | "完了";
export type Freq = "日次" | "週次" | "月次" | "不定期";
export type Ecrs = "E" | "C" | "R" | "S";
// アプリB：育成・定着計画の受け継ぎ方（作業→標準化／感覚→形式知化・教育）
export type ReproType = "標準化" | "形式知化" | "教育";

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
  tacit?: Tacit | null; // アプリB：暗黙知（5つの問い）
}

/** アプリB：暗黙知の言語化（5つの問い） */
export interface Tacit {
  q1?: string; // 何を見て判断しているか
  q2?: string; // OK/NG を分ける基準
  q3?: string; // 迷う・例外的なケース
  q4?: string; // ベテランが無意識にやっている工夫
  q5?: string; // 若手への教え方
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

/** アプリB：育成・定着計画（A の改善計画 plans とは別配列で持つ） */
export interface ReproPlan {
  id: string;
  taskId: string;
  taskName: string;
  type: ReproType; // 作業→標準化／感覚→形式知化 or 教育
  content: string; // 必須
  person: string; // 進める担当 必須
  target?: string; // 対象者（誰に）任意
  date: string; // 完了目標日 必須
  method?: string; // 手段・教材（マニュアル化/OJT/動画 等）任意
  status: PlanStatus;
  able?: number | null; // できる人：現在の人数（B-5）
  targetCount?: number | null; // できる人：目標人数（B-5）
}

/** アプリB：育成・定着計画モーダルの入力項目（id/status/able等はstore側で補完） */
export interface ReproPlanInput {
  taskId: string;
  taskName: string;
  type: ReproType;
  content: string;
  person: string;
  target?: string;
  date: string;
  method?: string;
}

export interface ProblemCategory {
  id: string;
  name: string;
  desc?: string;
}

/** 改善計画モーダルが扱う入力項目（id/status/afterMonthlyはstore側で補完） */
export interface PlanInput {
  taskId: string;
  taskName: string;
  content: string;
  person: string;
  date: string;
  method?: string;
  effect?: string;
}

/** 業務追加・編集モーダルが扱う入力項目（id/problems/scores等はstore側で補完） */
export interface TaskInput {
  group: string;
  content: string;
  person: string;
  time: number;
  freq: Freq;
  workType: WorkType;
  category: string;
  mmm: { muri: boolean; muda: boolean; mura: boolean };
}

export interface Masters {
  processCategories: string[];
  problemCategories: ProblemCategory[];
}

export interface AppState {
  company: string; // 会社名（クライアント名）。設定・マスタ画面で編集
  tasks: Task[]; // A/B 共有
  plans: Plan[]; // アプリA：改善計画
  reproPlans: ReproPlan[]; // アプリB：育成・定着計画（plansとは別配列）
  masters: Masters;
  currentGroup: string | null; // null = すべての業務名（全体一覧）
}
