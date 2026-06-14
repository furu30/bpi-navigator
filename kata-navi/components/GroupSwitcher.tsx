"use client";

import { groupsOf, useStore } from "@/lib/store";

const ALL = "__all__";

/** トップバー常設の業務名（ワークグループ）切替（実装指示書 §5）。 */
export function GroupSwitcher() {
  const { state, hydrated, setCurrentGroup, addGroup } = useStore();
  const groups = groupsOf(state);

  function onChange(value: string) {
    setCurrentGroup(value === ALL ? null : value);
  }

  function onAdd() {
    const name = window.prompt("新しい業務名を入力してください");
    if (name && name.trim()) addGroup(name.trim());
  }

  return (
    <div className="flex min-w-0 items-center gap-2">
      <span className="hidden text-xs font-bold text-[var(--muted)] sm:inline">
        業務名
      </span>
      <select
        value={state.currentGroup ?? ALL}
        onChange={(e) => onChange(e.target.value)}
        // hydration前は復元前のデモ値が出るため操作を抑止
        disabled={!hydrated}
        className="min-w-0 max-w-[150px] rounded-lg border-[1.5px] border-[var(--line)] px-2.5 py-1.5 text-[13px] focus:border-[var(--blue)] focus:outline-none sm:max-w-none"
      >
        <option value={ALL}>すべての業務名（全体一覧）</option>
        {groups.map((g) => (
          <option key={g} value={g}>
            {g}
          </option>
        ))}
      </select>
      <button
        type="button"
        onClick={onAdd}
        title="業務名を追加"
        className="flex-shrink-0 rounded-lg border-[1.5px] border-[var(--line)] bg-white px-3 py-1.5 text-xs font-bold text-[var(--navy)] hover:border-[#94a3b8]"
      >
        ＋<span className="hidden sm:inline"> 業務名</span>
      </button>
    </div>
  );
}
