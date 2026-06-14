"use client";

// グローバル状態（localStorage 永続化）のコンテキストとフック。
// 実体のProviderは components/StoreProvider.tsx。CRUDはステップ4以降で拡張する。

import { createContext, useContext } from "react";
import type {
  AppState,
  Ecrs,
  Plan,
  PlanInput,
  PlanStatus,
  Task,
  TaskInput,
} from "./types";

export const LS_KEY = "kata-navi-appA";

/** 業務名ファイル（業務名単位の保存／読込）の形 */
export interface GroupFile {
  group: string;
  tasks: Task[];
  plans: Plan[];
}

export interface StoreContextValue {
  state: AppState;
  /** localStorageからの復元完了フラグ（SSRハイドレーション対策） */
  hydrated: boolean;
  /** 会社名（クライアント名）を設定 */
  setCompany: (company: string) => void;
  /** 表示中の業務名を切り替える（null = すべての業務名／全体一覧） */
  setCurrentGroup: (group: string | null) => void;
  /** 新しい業務名を作成し、その業務名を選択状態にする */
  addGroup: (name: string) => void;
  /** 業務（作業）を追加 */
  addTask: (input: TaskInput) => void;
  /** 業務（作業）を編集（入力項目のみ更新。problems/scores/ecrsは保持） */
  updateTask: (id: string, input: TaskInput) => void;
  /** 業務（作業）を削除（関連する改善計画も併せて削除） */
  deleteTask: (id: string) => void;
  /** スコア（影響度／頻度）を更新（A-2 スライダー） */
  setScore: (id: string, key: "impact" | "freq", value: number) => void;
  /** 問題カテゴリのトグル（A-2） */
  toggleProblem: (id: string, problem: string) => void;
  /** ECRS区分の選択／解除（A-3。同じ値の再選択で解除） */
  setEcrs: (id: string, key: Ecrs) => void;
  /** 改善計画を追加（A-3「計画に追加」） */
  addPlan: (input: PlanInput) => void;
  /** 改善計画を編集（status/afterMonthlyは保持） */
  updatePlan: (id: string, input: PlanInput) => void;
  /** 改善計画を削除 */
  deletePlan: (id: string) => void;
  /** ステータス変更（A-4） */
  setPlanStatus: (id: string, status: PlanStatus) => void;
  /** 改善後の月間時間を設定（A-5。null=未入力） */
  setPlanAfter: (id: string, afterMonthly: number | null) => void;
  /** プロセス区分マスタ：追加（重複・空は無視） */
  addProcessCategory: (name: string) => void;
  /** プロセス区分マスタ：名称変更（既存作業のcategoryにも反映） */
  renameProcessCategory: (oldName: string, newName: string) => void;
  /** プロセス区分マスタ：削除（マスタからのみ削除） */
  deleteProcessCategory: (name: string) => void;
  /** 問題カテゴリマスタ：追加（重複・空は無視） */
  addProblemCategory: (name: string, desc?: string) => void;
  /** 問題カテゴリマスタ：編集（名称変更は既存作業のproblemsにも反映） */
  updateProblemCategory: (id: string, name: string, desc?: string) => void;
  /** 問題カテゴリマスタ：削除（既存作業のproblemsからも除去） */
  deleteProblemCategory: (id: string) => void;
  /** 業務名の名称変更（その業務名の全作業を付け替え） */
  renameGroup: (oldName: string, newName: string) => void;
  /** 業務名の削除（その業務名の作業＋関連plansを削除） */
  deleteGroup: (name: string) => void;
  /** 業務名ファイルを取り込む（同名業務名は置換し、その業務名を選択） */
  mergeGroupFile: (file: GroupFile) => void;
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

/** 指定した業務名に紐づく改善計画。 */
export function groupPlans(state: AppState, group: string): Plan[] {
  const ids = new Set(
    state.tasks.filter((t) => t.group === group).map((t) => t.id),
  );
  return state.plans.filter((p) => ids.has(p.taskId));
}
