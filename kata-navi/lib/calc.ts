// 計算ロジック（実装指示書 §4／アプリB §4）

import type { Freq, Task, TaskScores, WorkType } from "./types";

const FREQ_MULTIPLIER: Record<Freq, number> = {
  日次: 20,
  週次: 4,
  月次: 1,
  不定期: 1,
};

/** 月間の所要時間（分）= 所要時間 × 頻度係数 */
export function monthly(task: Pick<Task, "time" | "freq">): number {
  return task.time * (FREQ_MULTIPLIER[task.freq] ?? 1);
}

/** 優先度スコア = 影響度×0.6 + 頻度×0.4（小数1桁） */
export function totalScore(scores: TaskScores): number {
  return Math.round((scores.impact * 0.6 + scores.freq * 0.4) * 10) / 10;
}

/** 表示用：小数1桁固定の文字列（例: "3.0"） */
export function formatScore(scores: TaskScores): string {
  return totalScore(scores).toFixed(1);
}

// ===== アプリB（再現性ナビ）=====

const ATTR_TYPE: Record<WorkType, number> = { 感覚: 3, 混在: 2, 作業: 1, 未判定: 1 };
const ATTR_FREQ: Record<Freq, number> = { 日次: 3, 週次: 2, 月次: 1, 不定期: 1 };

/** 属人性リスク = 種別係数 × 頻度係数（B-1の並び順に使用） */
export function attrition(task: Pick<Task, "workType" | "freq">): number {
  return (ATTR_TYPE[task.workType] ?? 1) * (ATTR_FREQ[task.freq] ?? 1);
}

/** 暗黙知が言語化済みか（5つの問いの主要項目が埋まっているか） */
export function tacitDone(task: Pick<Task, "tacit">): boolean {
  const t = task.tacit;
  return !!(t && (t.q1 || t.q2 || t.q5));
}
