"use client";

import { useCallback, useEffect, useState } from "react";
import { demoState, uid } from "@/lib/demo";
import {
  type GroupFile,
  LS_KEY,
  StoreContext,
  type StoreContextValue,
} from "@/lib/store";
import type {
  AppState,
  Ecrs,
  PlanInput,
  PlanStatus,
  TaskInput,
} from "@/lib/types";

/** localStorageから状態を復元。壊れている／無いときはデモデータ。 */
function loadState(): AppState {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (raw) {
      const data = JSON.parse(raw);
      if (data && Array.isArray(data.tasks) && Array.isArray(data.plans)) {
        // 古い保存にも備えて欠けている項目を補完する
        if (!data.masters) data.masters = demoState().masters;
        if (data.currentGroup === undefined) data.currentGroup = null;
        if (typeof data.company !== "string") data.company = "";
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

  const setCompany = useCallback((company: string) => {
    setState((prev) => ({ ...prev, company }));
  }, []);

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

  const setScore = useCallback(
    (id: string, key: "impact" | "freq", value: number) => {
      setState((prev) => ({
        ...prev,
        tasks: prev.tasks.map((t) =>
          t.id === id ? { ...t, scores: { ...t.scores, [key]: value } } : t,
        ),
      }));
    },
    [],
  );

  const toggleProblem = useCallback((id: string, problem: string) => {
    setState((prev) => ({
      ...prev,
      tasks: prev.tasks.map((t) => {
        if (t.id !== id) return t;
        const has = t.problems.includes(problem);
        return {
          ...t,
          problems: has
            ? t.problems.filter((p) => p !== problem)
            : [...t.problems, problem],
        };
      }),
    }));
  }, []);

  const setEcrs = useCallback((id: string, key: Ecrs) => {
    setState((prev) => ({
      ...prev,
      tasks: prev.tasks.map((t) =>
        t.id === id ? { ...t, ecrs: t.ecrs === key ? undefined : key } : t,
      ),
    }));
  }, []);

  const addPlan = useCallback((input: PlanInput) => {
    setState((prev) => ({
      ...prev,
      plans: [
        ...prev.plans,
        { id: uid(), ...input, status: "未着手", afterMonthly: null },
      ],
    }));
  }, []);

  const updatePlan = useCallback((id: string, input: PlanInput) => {
    setState((prev) => ({
      ...prev,
      // status / afterMonthly は編集対象外なので保持する
      plans: prev.plans.map((p) => (p.id === id ? { ...p, ...input } : p)),
    }));
  }, []);

  const deletePlan = useCallback((id: string) => {
    setState((prev) => ({
      ...prev,
      plans: prev.plans.filter((p) => p.id !== id),
    }));
  }, []);

  const setPlanStatus = useCallback((id: string, status: PlanStatus) => {
    setState((prev) => ({
      ...prev,
      plans: prev.plans.map((p) => (p.id === id ? { ...p, status } : p)),
    }));
  }, []);

  const setPlanAfter = useCallback(
    (id: string, afterMonthly: number | null) => {
      setState((prev) => ({
        ...prev,
        plans: prev.plans.map((p) =>
          p.id === id ? { ...p, afterMonthly } : p,
        ),
      }));
    },
    [],
  );

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
    setCompany,
    setCurrentGroup,
    addGroup,
    addTask,
    updateTask,
    deleteTask,
    setScore,
    toggleProblem,
    setEcrs,
    addPlan,
    updatePlan,
    deletePlan,
    setPlanStatus,
    setPlanAfter,
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
