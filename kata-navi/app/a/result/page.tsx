"use client";

import Link from "next/link";
import { monthly } from "@/lib/calc";
import { curPlans, curTasks, useStore } from "@/lib/store";

// A-5 効果検証・レポート（改善後の月間時間から削減分・年換算、業務名別サマリ）
export default function ResultPage() {
  const { state, hydrated, setPlanAfter } = useStore();
  if (!hydrated) return <div className="text-[var(--muted)]">読み込み中…</div>;

  const ps = curPlans(state);
  const ts = curTasks(state);
  const done = ps.filter((p) => p.status === "完了").length;

  const beforeOf = (taskId: string) => {
    const t = state.tasks.find((x) => x.id === taskId);
    return t ? monthly(t) : 0;
  };

  const totalCut = ps.reduce((s, p) => {
    if (p.afterMonthly != null) return s + (beforeOf(p.taskId) - p.afterMonthly);
    return s;
  }, 0);
  const cutPos = Math.max(totalCut, 0);
  const yearHours = ((cutPos * 12) / 60).toFixed(0);

  return (
    <>
      <div className="mb-5 flex gap-2.5 rounded-xl border border-[var(--blue-line)] bg-[var(--blue-soft)] px-4 py-3 text-[13px] text-[#1e3a8a]">
        📈
        <div>
          <b className="text-[var(--blue-d)]">気づきポイント</b>
          ：改善前後を比べて、効果を数値で確認。<b>「これだけ減った」</b>
          が次のやる気になります。
        </div>
      </div>

      <div className="mb-[18px] grid grid-cols-2 gap-3 sm:grid-cols-4">
        <Stat value={ps.length} label="改善計画" />
        <Stat value={done} label="完了" />
        <Stat value={cutPos} label="削減(分/月)" green />
        <Stat value={yearHours} label="年換算(時間)" green />
      </div>

      {ps.length ? (
        <div className="overflow-x-auto rounded-xl shadow-[0_1px_3px_rgba(0,0,0,.05)]">
        <table className="w-full min-w-[560px] border-collapse text-[13px]">
          <thead>
            <tr>
              <th className={th}>改善</th>
              <th className={`${th} w-[90px]`}>改善前</th>
              <th className={`${th} w-[160px]`}>改善後</th>
              <th className={`${th} w-[90px]`}>効果</th>
            </tr>
          </thead>
          <tbody>
            {ps.map((p) => {
              const before = beforeOf(p.taskId);
              const after = p.afterMonthly;
              const cut = after != null ? before - after : null;
              return (
                <tr key={p.id}>
                  <td className={td}>
                    <b>{p.content}</b>
                    <div className="text-[12px] text-[var(--muted)]">
                      {p.taskName}
                    </div>
                  </td>
                  <td className={td}>{before}分/月</td>
                  <td className={td}>
                    <input
                      type="number"
                      value={after ?? ""}
                      placeholder="改善後"
                      onChange={(e) =>
                        setPlanAfter(
                          p.id,
                          e.target.value === "" ? null : Number(e.target.value),
                        )
                      }
                      className="w-[90px] rounded-lg border-[1.5px] border-[var(--line)] px-2 py-1 text-[13px] focus:border-[var(--blue)] focus:outline-none"
                    />
                    分/月
                  </td>
                  <td className={td}>
                    {cut != null ? (
                      <b
                        className={
                          cut > 0 ? "text-[var(--green)]" : "text-[var(--red)]"
                        }
                      >
                        {cut > 0 ? "−" : "+"}
                        {Math.abs(cut)}分
                      </b>
                    ) : (
                      <span className="text-[var(--muted)]">未入力</span>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        </div>
      ) : (
        <div className="rounded-xl border border-dashed border-[var(--line)] bg-white p-[34px] text-center text-[var(--muted)]">
          改善計画がありません。
        </div>
      )}

      <div className="mt-4 rounded-xl border border-[var(--line)] bg-white p-[18px] shadow-[0_1px_3px_rgba(0,0,0,.04)]">
        <h3 className="text-[15px] font-extrabold">
          📄 ふりかえりレポート（サマリ
          {state.currentGroup ? `：${state.currentGroup}` : ""}）
        </h3>
        <p className="mt-1.5 text-[13px] text-[var(--muted)]">
          {ts.length}件の業務を棚卸しし、{ps.length}件の改善に着手。月あたり{" "}
          <b className="text-[var(--green)]">
            {cutPos}分（約{yearHours}時間/年）
          </b>
          の削減見込み。次サイクルでは未着手の業務に取り組みます。
        </p>
        <div className="mt-2.5">
          <button
            type="button"
            onClick={() =>
              window.alert("試作版：本実装ではPDF/Excel出力に対応します")
            }
            className={btnOutSm}
          >
            レポート出力
          </button>
        </div>
      </div>

      <div className="mt-6 flex justify-between border-t border-[var(--line)] pt-[18px]">
        <Link href="/a/plan" className={btnOut}>
          ← 改善計画
        </Link>
        <Link href="/stock" className={btnPri}>
          次の改善サイクルへ ↻
        </Link>
      </div>
    </>
  );
}

function Stat({
  value,
  label,
  green,
}: {
  value: React.ReactNode;
  label: string;
  green?: boolean;
}) {
  return (
    <div className="rounded-xl border border-[var(--line)] bg-white p-4 text-center">
      <div
        className={`text-[26px] font-extrabold ${
          green ? "text-[var(--green)]" : "text-[var(--blue)]"
        }`}
      >
        {value}
      </div>
      <div className="mt-0.5 text-[12px] text-[var(--muted)]">{label}</div>
    </div>
  );
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
