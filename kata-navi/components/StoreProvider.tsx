"use client";

import { useCallback, useEffect, useState } from "react";
import { demoState, uid } from "@/lib/demo";
import {
  type GroupFile,
  LS_KEY,
  StoreContext,
  type StoreContextValue,
} from "@/lib/store";
import type { AppState, TaskInput } from "@/lib/types";

/** localStorageから状態を復元。壊れている／無いときはデモデータ。 */
function loadState(): AppState {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (raw) {
      const data = JSON.parse(raw);
      if (data && Array.isArray(data.tasks) && Array.isArray(data.plans)) {
        // masters が欠けている古い保存にも備える
        if (!data.masters) data.masters = demoState().masters;
        if (data.currentGroup === undefined) data.currentGroup = null;
        return data as AppState;
      }
    }
  } catch {
    /* 破損データは無視してデモにフォールバック */
  }
  return demoState();
}

export function StoreProvider({ children }: { children: React.ReactNode }) {
  // SSRとの不一致を避けるため、初期値はデモ（決定的）にし、
  // マウント後に localStorage の内容へ置き換える。
  const [state, setState] = useState<AppState>(() => demoState());
  const [hydrated, setHydrated] = useState(false);

  // 初回マウント時に localStorage から復元する。
  // SSR出力と初期描画を一致させるため、復元はマウント後のeffectで行う
  // （初期描画は決定的なデモ値、その後クライアントでlocalStorageへ差し替え）。
  useEffect(() => {
    // eslint-disable-next-line react-hooks/set-state-in-effect -- ハイドレーション後の一度きりの復元
    setState(loadState());
    setHydrated(true);
  }, []);

  // 変更を localStorage へ自動保存（復元完了後のみ）
  useEffect(() => {
    if (!hydrated) return;
    try {
      localStorage.setItem(LS_KEY, JSON.stringify(state));
    } catch {
      /* 容量超過などは無視 */
    }
  }, [state, hydrated]);

  const setCurrentGroup = useCallback((group: string | null) => {
    setState((prev) => ({ ...prev, currentGroup: group }));
  }, []);

  const addGroup = useCallback((name: string) => {
    const g = name.trim();
    if (!g) return;
    setState((prev) => ({ ...prev, currentGroup: g }));
  }, []);

  const addTask = useCallback((input: TaskInput) => {
    setState((prev) => ({
      ...prev,
      tasks: [
        ...prev.tasks,
        {
          id: uid(),
          ...input,
          problems: [],
          scores: { impact: 3, freq: 3 },
        },
      ],
    }));
  }, []);

  const updateTask = useCallback((id: string, input: TaskInput) => {
    setState((prev) => ({
      ...prev,
      // problems / scores / ecrs は編集対象外なので保持する
      tasks: prev.tasks.map((t) => (t.id === id ? { ...t, ...input } : t)),
    }));
  }, []);

  const deleteTask = useCallback((id: string) => {
    setState((prev) => ({
      ...prev,
      tasks: prev.tasks.filter((t) => t.id !== id),
      plans: prev.plans.filter((p) => p.taskId !== id),
    }));
  }, []);

  const renameGroup = useCallback((oldName: string, newName: string) => {
    const next = newName.trim();
    if (!next || next === oldName) return;
    setState((prev) => ({
      ...prev,
      tasks: prev.tasks.map((t) =>
        t.group === oldName ? { ...t, group: next } : t,
      ),
      currentGroup: prev.currentGroup === oldName ? next : prev.currentGroup,
    }));
  }, []);

  const deleteGroup = useCallback((name: string) => {
    setState((prev) => {
      const removedIds = new Set(
        prev.tasks.filter((t) => t.group === name).map((t) => t.id),
      );
      return {
        ...prev,
        tasks: prev.tasks.filter((t) => t.group !== name),
        plans: prev.plans.filter((p) => !removedIds.has(p.taskId)),
        currentGroup: prev.currentGroup === name ? null : prev.currentGroup,
      };
    });
  }, []);

  const mergeGroupFile = useCallback((file: GroupFile) => {
    setState((prev) => {
      // 同名業務名の作業・関連plansを除去し、ファイル側で置換する
      const incomingTaskIds = new Set(file.tasks.map((t) => t.id));
      const keptTasks = prev.tasks.filter((t) => t.group !== file.group);
      const keptPlans = prev.plans.filter(
        (p) => !incomingTaskIds.has(p.taskId),
      );
      return {
        ...prev,
        tasks: [...keptTasks, ...file.tasks],
        plans: [...keptPlans, ...(file.plans ?? [])],
        currentGroup: file.group,
      };
    });
  }, []);

  const replaceState = useCallback((next: AppState) => {
    setState(next);
  }, []);

  const resetDemo = useCallback(() => {
    setState(demoState());
  }, []);

  const value: StoreContextValue = {
    state,
    hydrated,
    setCurrentGroup,
    addGroup,
    addTask,
    updateTask,
    deleteTask,
    renameGroup,
    deleteGroup,
    mergeGroupFile,
    replaceState,
    resetDemo,
  };

  return (
    <StoreContext.Provider value={value}>{children}</StoreContext.Provider>
  );
}
