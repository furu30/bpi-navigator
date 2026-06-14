"use client";

import { curReproPlans, curTasks, useStore } from "@/lib/store";

// アプリB 各画面のステップ実装前プレースホルダ（ティール）。
// currentGroup の絞り込みが効いていることを件数で可視化しておく。
export function BPlaceholder({
  step,
  lead,
  next,
}: {
  step: string;
  lead: string;
  next: string;
}) {
  const { state, hydrated } = useStore();
  const scope = state.currentGroup ?? "すべての業務名（全体一覧）";
  const taskCount = hydrated ? curTasks(state).length : 0;
  const planCount = hydrated ? curReproPlans(state).length : 0;

  return (
    <>
      <div className="mb-5 flex gap-2.5 rounded-xl border border-[var(--teal-line)] bg-[var(--teal-soft)] px-4 py-3 text-[13px] text-[#134e4a]">
        💡<div>{lead}</div>
      </div>
      <div className="rounded-xl border border-dashed border-[var(--line)] bg-white p-8 text-center">
        <div className="text-[15px] font-extrabold">{step}</div>
        <p className="mt-2 text-[13px] text-[var(--muted)]">{next}</p>
        <p className="mt-3 text-[12.5px] text-[var(--muted)]">
          表示中の業務名：<b className="text-[var(--teal-d)]">{scope}</b>
          {" ／ "}対象の作業 <b>{taskCount}</b>件 ・ 育成計画 <b>{planCount}</b>件
        </p>
      </div>
    </>
  );
}
