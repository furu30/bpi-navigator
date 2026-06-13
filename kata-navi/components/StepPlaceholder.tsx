"use client";

import { curPlans, curTasks, useStore } from "@/lib/store";

// ステップ5以降で本実装する画面の暫定プレースホルダ。
// currentGroup の絞り込みが各画面に効いていることを件数で可視化しておく。
export function StepPlaceholder({
  step,
  lead,
}: {
  step: string;
  lead: string;
}) {
  const { state, hydrated } = useStore();
  const scope = state.currentGroup ?? "すべての業務名（全体一覧）";
  const taskCount = hydrated ? curTasks(state).length : 0;
  const planCount = hydrated ? curPlans(state).length : 0;

  return (
    <>
      <div className="mb-5 flex gap-2.5 rounded-xl border border-[var(--blue-line)] bg-[var(--blue-soft)] px-4 py-3 text-[13px] text-[#1e3a8a]">
        💡<div>{lead}</div>
      </div>
      <div className="rounded-xl border border-dashed border-[var(--line)] bg-white p-8 text-center">
        <div className="text-[15px] font-extrabold">{step}</div>
        <p className="mt-2 text-[13px] text-[var(--muted)]">
          この画面はステップ5で実装します（基盤のみ構築済み）。
        </p>
        <p className="mt-3 text-[12.5px] text-[var(--muted)]">
          表示中の業務名：<b className="text-[var(--blue-d)]">{scope}</b>
          {" ／ "}対象の作業 <b>{taskCount}</b>件 ・ 改善計画 <b>{planCount}</b>件
        </p>
      </div>
    </>
  );
}
