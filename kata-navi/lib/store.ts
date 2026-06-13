"use client";

// グローバル状態（localStorage 永続化）のコンテキストとフック。
// 実体のProviderは components/StoreProvider.tsx。CRUDはステップ4以降で拡張する。

import { createContext, useContext } from "react";
import type { AppState, Plan, Task } from "./types";

export const LS_KEY = "kata-navi-appA";

export interface StoreContextValue {
  state: AppState;
  /** localStorageからの復元完了フラグ（SSRハイドレーション対策） */
  hydrated: boolean;
  /** 表示中の業務名を切り替える（null = すべての業務名／全体一覧） */
  setCurrentGroup: (group: string | null) => void;
  /** 新しい業務名を作成し、その業務名を選択状態にする */
  addGroup: (name: string) => void;
  /** 状態を丸ごと差し替える（JSON読込・デモ初期化などで使用） */
  replaceState: (next: AppState) => void;
  /** デモデータに戻す */
  resetDemo: () => void;
}

export const StoreContext = createContext<StoreContextValue | null>(null);

export function useStore(): StoreContextValue {
  const ctx = useContext(StoreContext);
  if (!ctx) {
    throw new Error("useStore は StoreProvider の内側で呼び出してください");
  }
  return ctx;
}

// ===== 派生ヘルパー（純粋関数） =====

/** 登録済みの業務名一覧（出現順・重複排除）。選択中の空グループも含める。 */
export function groupsOf(state: AppState): string[] {
  const groups = [...new Set(state.tasks.map((t) => t.group))];
  if (state.currentGroup && !groups.includes(state.currentGroup)) {
    groups.push(state.currentGroup);
  }
  return groups;
}

/** 現在の業務名で絞り込んだ作業（null のときは全件）。 */
export function curTasks(state: AppState): Task[] {
  return state.currentGroup
    ? state.tasks.filter((t) => t.group === state.currentGroup)
    : state.tasks;
}

/** 現在の業務名に紐づく改善計画。 */
export function curPlans(state: AppState): Plan[] {
  const ids = new Set(curTasks(state).map((t) => t.id));
  return state.plans.filter((p) => ids.has(p.taskId));
}
