"use client";

import Link from "next/link";
import { curReproPlans, curTasks, useStore } from "@/lib/store";
import type { ReproPlan } from "@/lib/types";

// B-5 定着の確認・レポート（できる人 現在/目標で定着度、業務名別サマリ）
export default function BResultPage() {
  const { state, hydrated, setReproAble, setReproTargetCount } = useStore();
  if (!hydrated) return <div className="text-[var(--muted)]">読み込み中…</div>;

  const ps = curReproPlans(state);
  const ts = curTasks(state);
  const kankaku = ts.filter(
    (t) => t.workType === "感覚" || t.workType === "混在",
  ).length;
  const done = ps.filter((p) => isSettled(p)).length;
  const ableTotal = ps.reduce((s, p) => s + (p.able ?? 0), 0);

  return (
    <>
      <div className="mb-5 flex gap-2.5 rounded-xl border border-[var(--teal-line)] bg-[var(--teal-soft)] px-4 py-3 text-[13px] text-[#134e4a]">
        📈
        <div>
          <b className="text-[var(--teal-d)]">気づきポイント</b>
          ：受け継げたか＝<b>「できる人が増えたか」</b>
          で確認します。属人性が再現性に変わった手応えが、次の継承につながります。
        </div>
      </div>

      <div className="mb-[18px] grid grid-cols-2 gap-3 sm:grid-cols-4">
        <Stat value={ps.length} label="育成計画" />
        <Stat value={done} label="定着（完了）" green />
        <Stat value={ableTotal} label="できる人(合計)" green />
        <Stat value={kankaku} label="感覚業務" />
      </div>

      {ps.length ? (
        <div className="overflow-x-auto rounded-xl shadow-[0_1px_3px_rgba(0,0,0,.05)]">
          <table className="w-full min-w-[560px] border-collapse text-[13px]">
            <thead>
              <tr>
                <th className={th}>計画</th>
                <th className={`${th} w-[180px]`}>できる人（現在/目標）</th>
                <th className={`${th} w-[160px]`}>定着度</th>
              </tr>
            </thead>
            <tbody>
              {ps.map((p) => {
                const ratio =
                  p.able != null && p.targetCount
                    ? Math.min(100, Math.round((p.able / p.targetCount) * 100))
                    : 0;
                return (
                  <tr key={p.id}>
                    <td className={td}>
                      <b>{p.content}</b>
                      <div className="text-[12px] text-[var(--muted)]">
                        {p.taskName} ／ {p.type}
                      </div>
                    </td>
                    <td className={td}>
                      <input
                        type="number"
                        min={0}
                        value={p.able ?? ""}
                        placeholder="現在"
                        onChange={(e) =>
                          setReproAble(
                            p.id,
                            e.target.value === "" ? null : Number(e.target.value),
                          )
                        }
                        className={numCls}
                      />{" "}
                      /{" "}
                      <input
                        type="number"
                        min={0}
                        value={p.targetCount ?? ""}
                        placeholder="目標"
                        onChange={(e) =>
                          setReproTargetCount(
                            p.id,
                            e.target.value === "" ? null : Number(e.target.value),
                          )
                        }
                        className={numCls}
                      />{" "}
                      人
                    </td>
                    <td className={td}>
                      {p.targetCount ? (
                        <>
                          <div className="h-2 overflow-hidden rounded bg-[#f1f5f9]">
                            <span
                              className="block h-full bg-[var(--green)]"
                              style={{ width: `${ratio}%` }}
                            />
                          </div>
                          <div className="mt-0.5 text-[11px] text-[var(--muted)]">
                            {ratio}% できる
                          </div>
                        </>
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
          育成計画がありません。
        </div>
      )}

      <div className="mt-4 rounded-xl border border-[var(--line)] bg-white p-[18px] shadow-[0_1px_3px_rgba(0,0,0,.04)]">
        <h3 className="text-[15px] font-extrabold">
          📄 振り返りレポート（サマリ
          {state.currentGroup ? `：${state.currentGroup}` : ""}）
        </h3>
        <p className="mt-1.5 text-[13px] text-[var(--muted)]">
          感覚業務 {kankaku}件に対し、{ps.length}件の育成・定着計画を実施。
          <b className="text-[var(--teal-d)]">「できる人」</b>
          を増やし、属人性を再現性に変えていきます。
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
        <Link href="/b/plan" className={btnOut}>
          ← 育成・定着計画
        </Link>
        <Link href="/stock" className={btnPri}>
          次の継承サイクルへ ↻
        </Link>
      </div>
    </>
  );
}

/** 定着＝完了、または できる人が目標に到達 */
function isSettled(p: ReproPlan): boolean {
  if (p.status === "完了") return true;
  return (
    p.able != null &&
    p.targetCount != null &&
    p.targetCount > 0 &&
    p.able >= p.targetCount
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
          green ? "text-[var(--green)]" : "text-[var(--teal)]"
        }`}
      >
        {value}
      </div>
      <div className="mt-0.5 text-[12px] text-[var(--muted)]">{label}</div>
    </div>
  );
}

const numCls =
  "w-[60px] rounded-lg border-[1.5px] border-[var(--line)] px-2 py-1 text-[13px] focus:border-[var(--teal)] focus:outline-none";
const btnPri =
  "rounded-[9px] bg-[var(--teal)] px-[18px] py-2.5 text-[13.5px] font-bold text-white hover:bg-[var(--teal-d)]";
const btnOut =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-[18px] py-2.5 text-[13.5px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
const btnOutSm =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-3 py-1.5 text-[12px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
const th =
  "border-b border-[var(--line)] bg-[#f1f5f9] px-3 py-2.5 text-left text-[12px] font-bold text-[#475569]";
const td = "border-b border-[#f1f5f9] bg-white px-3 py-2.5 align-middle";
