"use client";

import Link from "next/link";
import { useState } from "react";
import { AiAssistModal } from "@/components/AiAssistModal";
import { PlanModal } from "@/components/PlanModal";
import { monthly, totalScore } from "@/lib/calc";
import { curTasks, curPlans, useStore } from "@/lib/store";
import type { Plan, PlanInput, PlanStatus, Task } from "@/lib/types";

const STATUSES: PlanStatus[] = ["未着手", "進行中", "完了"];

const PLAN_AI_SYSTEM =
  "あなたは中小製造業の業務改善コンサルタントです。以下の業務（優先度・ECRS区分・問題点つき）に対し、具体的な改善計画案（改善内容／想定手段／期待効果の目安）を業務ごとに簡潔な日本語で提案してください。ECRS（排除・結合・交換・簡素化）の観点を活かしてください。";

function planAiContext(tasks: Task[]): string {
  const targets = tasks.filter(
    (t) => t.workType === "作業" || t.workType === "混在",
  );
  if (targets.length === 0) return "（作業・混在の業務がありません）";
  return targets
    .map((t) => {
      const ecrs = t.ecrs ? ` / ECRS:${t.ecrs}` : "";
      const probs = t.problems.length ? ` / 問題:${t.problems.join("・")}` : "";
      return `- ${t.content}（${monthly(t)}分/月 / 優先度${totalScore(t.scores).toFixed(1)}${ecrs}${probs}）`;
    })
    .join("\n");
}

// A-4 改善計画（計画の追加・編集・削除、ステータス変更）
export default function PlanPage() {
  const {
    state,
    hydrated,
    updatePlan,
    deletePlan,
    setPlanStatus,
  } = useStore();
  const [editing, setEditing] = useState<Plan | null>(null);
  const [aiOpen, setAiOpen] = useState(false);
  if (!hydrated) return <div className="text-[var(--muted)]">読み込み中…</div>;

  const ps = curPlans(state);
  const ts = curTasks(state);

  function handleSubmit(input: PlanInput) {
    if (editing) updatePlan(editing.id, input);
    setEditing(null);
  }

  return (
    <>
      <div className="mb-5 flex gap-2.5 rounded-xl border border-[var(--blue-line)] bg-[var(--blue-soft)] px-4 py-3 text-[13px] text-[#1e3a8a]">
        🛠
        <div>
          <b className="text-[var(--blue-d)]">気づきポイント</b>
          ：改善候補を実行プランに。<b>必須は「内容・担当・期日」だけ</b>
          。手段や効果は任意（詳細を開いて記入）。
        </div>
      </div>

      <div className="mb-4 flex items-center justify-between gap-3">
        <h2 className="text-[20px] font-extrabold">改善計画（{ps.length}件）</h2>
        <button
          type="button"
          onClick={() => setAiOpen(true)}
          className="rounded-[9px] border border-[#ddd6fe] bg-[#ede9fe] px-3 py-1.5 text-[12px] font-bold text-[#6d28d9]"
        >
          🤖 AIアシスト
        </button>
      </div>

      {ps.length ? (
        <div className="overflow-x-auto rounded-xl shadow-[0_1px_3px_rgba(0,0,0,.05)]">
        <table className="w-full min-w-[620px] border-collapse text-[13px]">
          <thead>
            <tr>
              <th className={th}>改善内容</th>
              <th className={`${th} w-[80px]`}>担当</th>
              <th className={`${th} w-[115px]`}>期日</th>
              <th className={`${th} w-[96px]`}>状況</th>
              <th className={`${th} w-[110px]`} />
            </tr>
          </thead>
          <tbody>
            {ps.map((p) => (
              <tr key={p.id}>
                <td className={td}>
                  <b>{p.content}</b>
                  <div className="text-[12px] text-[var(--muted)]">
                    対象：{p.taskName}
                    {p.method ? ` ／ ${p.method}` : ""}
                  </div>
                </td>
                <td className={td}>{p.person || "-"}</td>
                <td className={td}>{p.date || "-"}</td>
                <td className={td}>
                  <select
                    value={p.status}
                    onChange={(e) =>
                      setPlanStatus(p.id, e.target.value as PlanStatus)
                    }
                    className={`rounded-full px-2 py-[3px] text-[11px] font-bold ${statusCls(
                      p.status,
                    )}`}
                  >
                    {STATUSES.map((s) => (
                      <option key={s}>{s}</option>
                    ))}
                  </select>
                </td>
                <td className={td}>
                  <div className="flex gap-1">
                    <button
                      type="button"
                      onClick={() => setEditing(p)}
                      className={btnOutSm}
                    >
                      編集
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        if (window.confirm("この改善計画を削除しますか？"))
                          deletePlan(p.id);
                      }}
                      className={btnOutSm}
                    >
                      削除
                    </button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        </div>
      ) : (
        <div className="rounded-xl border border-dashed border-[var(--line)] bg-white p-[34px] text-center text-[var(--muted)]">
          A-3「ECRSで改善案」から計画を追加してください。
        </div>
      )}

      <div className="mt-6 flex justify-between border-t border-[var(--line)] pt-[18px]">
        <Link href="/a/ecrs" className={btnOut}>
          ← ECRS
        </Link>
        <Link href="/a/result" className={btnPri}>
          効果検証 →
        </Link>
      </div>

      {editing && (
        <PlanModal
          taskId={editing.taskId}
          taskName={editing.taskName}
          initial={editing}
          onClose={() => setEditing(null)}
          onSubmit={handleSubmit}
        />
      )}

      {aiOpen && (
        <AiAssistModal
          title="改善計画のAIアシスト"
          system={PLAN_AI_SYSTEM}
          context={`# 改善対象の業務\n${planAiContext(ts)}`}
          onClose={() => setAiOpen(false)}
        />
      )}
    </>
  );
}

function statusCls(status: PlanStatus): string {
  if (status === "完了") return "bg-[#dcfce7] text-[#15803d]";
  if (status === "進行中") return "bg-[#dbeafe] text-[#1d4ed8]";
  return "bg-[#f1f5f9] text-[#64748b]";
}

const btnPri =
  "rounded-[9px] bg-[var(--blue)] px-[18px] py-2.5 text-[13.5px] font-bold text-white hover:bg-[var(--blue-d)]";
const btnOut =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-[18px] py-2.5 text-[13.5px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
const btnOutSm =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-3 py-1.5 text-[12px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
const th =
  "border-b border-[var(--line)] bg-[#f1f5f9] px-3 py-2.5 text-left text-[12px] font-bold text-[#475569]";
const td = "border-b border-[#f1f5f9] bg-white px-3 py-2.5 align-middle";
