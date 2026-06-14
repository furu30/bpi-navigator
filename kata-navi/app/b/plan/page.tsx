"use client";

import Link from "next/link";
import { useState } from "react";
import { AiAssistModal } from "@/components/AiAssistModal";
import { ReproPlanModal } from "@/components/ReproPlanModal";
import { tacitDone } from "@/lib/calc";
import { curReproPlans, curTasks, useStore } from "@/lib/store";
import type {
  PlanStatus,
  ReproPlan,
  ReproPlanInput,
  ReproType,
  Task,
} from "@/lib/types";

const STATUSES: PlanStatus[] = ["未着手", "進行中", "完了"];

const REPRO_AI_SYSTEM =
  "あなたは中小製造業の人材育成・技能伝承のコンサルタントです。以下の業務（種別・暗黙知の言語化状況つき）について、属人性を再現性に変えるための育成・定着計画（標準化／形式知化／教育・OJTのいずれか・内容・対象者・手段の目安）を業務ごとに簡潔な日本語で提案してください。";

function reproAiContext(tasks: Task[]): string {
  if (tasks.length === 0) return "（業務がありません）";
  return tasks
    .map((t) => {
      const tacit =
        t.workType === "感覚" || t.workType === "混在"
          ? tacitDone(t)
            ? " / 暗黙知:言語化済"
            : " / 暗黙知:未"
          : "";
      return `- ${t.content}（種別:${t.workType} / 担当:${t.person} / ${t.freq}${tacit}）`;
    })
    .join("\n");
}

// B-4 育成・定着計画（reproPlans のCRUD・ステータス・種別バッジ）
export default function BPlanPage() {
  const {
    state,
    hydrated,
    updateReproPlan,
    deleteReproPlan,
    setReproPlanStatus,
  } = useStore();
  const [editing, setEditing] = useState<ReproPlan | null>(null);
  const [aiOpen, setAiOpen] = useState(false);
  if (!hydrated) return <div className="text-[var(--muted)]">読み込み中…</div>;

  const ps = curReproPlans(state);
  const ts = curTasks(state);

  function handleSubmit(input: ReproPlanInput) {
    if (editing) updateReproPlan(editing.id, input);
    setEditing(null);
  }

  return (
    <>
      <div className="mb-5 flex gap-2.5 rounded-xl border border-[var(--teal-line)] bg-[var(--teal-soft)] px-4 py-3 text-[13px] text-[#134e4a]">
        📋
        <div>
          <b className="text-[var(--teal-d)]">気づきポイント</b>
          ：「誰が・いつまでに・どう」受け継ぐかを計画に。標準化／形式知化／教育を同じ仕組みで管理します。
        </div>
      </div>

      <div className="mb-4 flex items-center justify-between gap-3">
        <h2 className="text-[20px] font-extrabold">育成・定着計画（{ps.length}件）</h2>
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
          <table className="w-full min-w-[680px] border-collapse text-[13px]">
            <thead>
              <tr>
                <th className={th}>内容</th>
                <th className={`${th} w-[84px]`}>種別</th>
                <th className={`${th} w-[80px]`}>担当</th>
                <th className={`${th} w-[112px]`}>期日</th>
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
                      {p.target ? ` ／ 対象者：${p.target}` : ""}
                    </div>
                  </td>
                  <td className={td}>
                    <TypeBadge type={p.type} />
                  </td>
                  <td className={td}>{p.person || "-"}</td>
                  <td className={td}>{p.date || "-"}</td>
                  <td className={td}>
                    <select
                      value={p.status}
                      onChange={(e) =>
                        setReproPlanStatus(p.id, e.target.value as PlanStatus)
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
                          if (window.confirm("この育成・定着計画を削除しますか？"))
                            deleteReproPlan(p.id);
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
          B-3「標準化／形式知化・教育」から計画を追加してください。
        </div>
      )}

      <div className="mt-6 flex justify-between border-t border-[var(--line)] pt-[18px]">
        <Link href="/b/route" className={btnOut}>
          ← 標準化／形式知化・教育
        </Link>
        <Link href="/b/result" className={btnPri}>
          定着の確認 →
        </Link>
      </div>

      {editing && (
        <ReproPlanModal
          taskId={editing.taskId}
          taskName={editing.taskName}
          type={editing.type}
          initial={editing}
          onClose={() => setEditing(null)}
          onSubmit={handleSubmit}
        />
      )}

      {aiOpen && (
        <AiAssistModal
          title="育成・定着計画のAIアシスト"
          system={REPRO_AI_SYSTEM}
          context={`# 対象の業務\n${reproAiContext(ts)}`}
          onClose={() => setAiOpen(false)}
        />
      )}
    </>
  );
}

function TypeBadge({ type }: { type: ReproType }) {
  const cls =
    type === "標準化"
      ? "bg-[#ccfbf1] text-[#0f766e]"
      : type === "形式知化"
        ? "bg-[#dcfce7] text-[#15803d]"
        : "bg-[#f1f5f9] text-[#64748b]";
  return (
    <span className={`rounded-full px-2 py-0.5 text-[11px] font-bold ${cls}`}>
      {type}
    </span>
  );
}

function statusCls(status: PlanStatus): string {
  if (status === "完了") return "bg-[#dcfce7] text-[#15803d]";
  if (status === "進行中") return "bg-[#ccfbf1] text-[#0f766e]";
  return "bg-[#f1f5f9] text-[#64748b]";
}

const btnPri =
  "rounded-[9px] bg-[var(--teal)] px-[18px] py-2.5 text-[13.5px] font-bold text-white hover:bg-[var(--teal-d)]";
const btnOut =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-[18px] py-2.5 text-[13.5px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
const btnOutSm =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-3 py-1.5 text-[12px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
const th =
  "border-b border-[var(--line)] bg-[#f1f5f9] px-3 py-2.5 text-left text-[12px] font-bold text-[#475569]";
const td = "border-b border-[#f1f5f9] bg-white px-3 py-2.5 align-middle";
