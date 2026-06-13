// JSON保存／読込ユーティリティ（ブラウザ専用：クライアントから呼ぶこと）

import type { GroupFile } from "./store";
import type { AppState } from "./types";

/** 任意のデータをJSONファイルとしてダウンロードさせる */
function downloadJSON(filename: string, data: unknown) {
  const blob = new Blob([JSON.stringify(data, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

/** 全体（全業務名）をエクスポート */
export function exportAll(state: AppState) {
  downloadJSON("kata-navi-appA-all.json", {
    tasks: state.tasks,
    plans: state.plans,
    masters: state.masters,
    currentGroup: state.currentGroup,
  } satisfies AppState);
}

/** 指定した業務名だけをエクスポート（業務名ファイル） */
export function exportGroup(state: AppState, group: string) {
  const tasks = state.tasks.filter((t) => t.group === group);
  const ids = new Set(tasks.map((t) => t.id));
  const plans = state.plans.filter((p) => ids.has(p.taskId));
  const file: GroupFile = { group, tasks, plans };
  downloadJSON(`kata-navi_${group}.json`, file);
}

export type ParseResult =
  | { kind: "group"; data: GroupFile }
  | { kind: "full"; data: AppState }
  | { kind: "error"; message: string };

/** 読み込んだテキストを判定して全体ファイル／業務名ファイルに振り分ける */
export function parseImport(text: string): ParseResult {
  let data: unknown;
  try {
    data = JSON.parse(text);
  } catch {
    return { kind: "error", message: "JSONとして読み込めませんでした。" };
  }
  if (!data || typeof data !== "object") {
    return { kind: "error", message: "読み込めない形式です。" };
  }
  const obj = data as Record<string, unknown>;
  // 業務名ファイル（group プロパティ付き）
  if (typeof obj.group === "string" && Array.isArray(obj.tasks)) {
    return {
      kind: "group",
      data: {
        group: obj.group,
        tasks: obj.tasks as GroupFile["tasks"],
        plans: Array.isArray(obj.plans) ? (obj.plans as GroupFile["plans"]) : [],
      },
    };
  }
  // 全体ファイル
  if (Array.isArray(obj.tasks) && Array.isArray(obj.plans)) {
    return { kind: "full", data: obj as unknown as AppState };
  }
  return { kind: "error", message: "読み込めない形式です。" };
}
