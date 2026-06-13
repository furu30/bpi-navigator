// 計算ロジック（実装指示書 §4）

import type { Freq, Task, TaskScores } from "./types";

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
