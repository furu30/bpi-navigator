"use client";

import Link from "next/link";
import { useRouter } from "next/navigation";
import { useState } from "react";
import { PlanModal } from "@/components/PlanModal";
import { monthly, totalScore } from "@/lib/calc";
import { curTasks, useStore } from "@/lib/store";
import type { Ecrs, PlanInput, Task } from "@/lib/types";

const ECRS_OPTIONS: { key: Ecrs; name: string; desc: string }[] = [
  { key: "E", name: "排除", desc: "やめられないか" },
  { key: "C", name: "結合", desc: "まとめられないか" },
  { key: "R", name: "交換", desc: "順序や場所を変えられないか" },
  { key: "S", name: "簡素化", desc: "もっと簡単にできないか" },
];

// A-3 ECRSで改善案（作業/混在のみ対象、感覚業務は「アプリBで扱います」案内）
export default function EcrsPage() {
  const { state, hydrated, setEcrs, addPlan } = useStore();
  const router = useRouter();
  const [planTask, setPlanTask] = useState<Task | null>(null);
  if (!hydrated) return <div className="text-[var(--muted)]">読み込み中…</div>;

  const ts = curTasks(state);
  const targets = ts.filter(
    (t) => t.workType === "作業" || t.workType === "混在",
  );
  const kankaku = ts.filter((t) => t.workType === "感覚");

  function handleSubmitPlan(input: PlanInput) {
    addPlan(input);
    setPlanTask(null);
    router.push("/a/plan");
  }

  return (
    <>
      <div className="mb-5 flex gap-2.5 rounded-xl border border-[var(--blue-line)] bg-[var(--blue-soft)] px-4 py-3 text-[13px] text-[#1e3a8a]">
        🎯
        <div>
          <b className="text-[var(--blue-d)]">気づきポイント</b>
          ：作業は<b>ECRS</b>
          でムダ取り。排除→結合→交換→簡素化の順で考えると効果が大きい改善から見つかります。
        </div>
      </div>

      {kankaku.length > 0 && (
        <div className="mb-3.5 rounded-lg border border-[#fde68a] bg-[#fffbeb] px-3 py-2 text-[11.5px] text-[var(--amber)]">
          🧠 感覚業務が {kankaku.length}件 あります（例：
          {kankaku
            .slice(0, 2)
            .map((t) => t.content)
            .join("、")}
          {kankaku.length > 2 ? " ほか" : ""}）。これらは効率化ではなく
          <b>再現性ナビ（アプリB）</b>で形式知化・教育として扱います。
        </div>
      )}

      {targets.length ? (
        targets.map((t) => (
          <div
            key={t.id}
            className="mb-3.5 rounded-xl border border-[var(--line)] bg-white p-[18px] shadow-[0_1px_3px_rgba(0,0,0,.04)]"
          >
            <div className="mb-1.5">
              <h3 className="text-[15px] font-extrabold">{t.content}</h3>
              <div className="text-[12px] text-[var(--muted)]">
                {t.workType} ・ {monthly(t)}分/月 ・ 優先度{" "}
                {totalScore(t.scores).toFixed(1)}
              </div>
            </div>
            <div className="mb-1.5 text-[12px] text-[var(--muted)]">
              どのECRSで攻めますか？
            </div>
            <div>
              {ECRS_OPTIONS.map((o) => {
                const on = t.ecrs === o.key;
                return (
                  <button
                    key={o.key}
                    type="button"
                    title={o.desc}
                    onClick={() => setEcrs(t.id, o.key)}
                    className={`m-[3px] inline-block rounded-lg border-[1.5px] px-3 py-[5px] text-[12px] ${
                      on
                        ? "border-[var(--blue)] bg-[var(--blue)] text-white"
                        : "border-[var(--line)] text-[var(--ink)]"
                    }`}
                  >
                    {o.key}：{o.name}
                  </button>
                );
              })}
            </div>
            {t.ecrs && (
              <div className="mt-2.5">
                <button
                  type="button"
                  onClick={() => setPlanTask(t)}
                  className="rounded-[9px] bg-[var(--blue)] px-3 py-1.5 text-[12px] font-bold text-white hover:bg-[var(--blue-d)]"
                >
                  この改善を計画に追加 →
                </button>
              </div>
            )}
          </div>
        ))
      ) : (
        <div className="rounded-xl border border-dashed border-[var(--line)] bg-white p-[34px] text-center text-[var(--muted)]">
          作業・混在の業務がありません。棚卸しで登録してください。
        </div>
      )}

      <div className="mt-6 flex justify-between border-t border-[var(--line)] pt-[18px]">
        <Link href="/a/analyze" className={btnOut}>
          ← 問題の見える化
        </Link>
        <Link href="/a/plan" className={btnPri}>
          改善計画 →
        </Link>
      </div>

      {planTask && (
        <PlanModal
          task={planTask}
          onClose={() => setPlanTask(null)}
          onSubmit={handleSubmitPlan}
        />
      )}
    </>
  );
}

const btnPri =
  "rounded-[9px] bg-[var(--blue)] px-[18px] py-2.5 text-[13.5px] font-bold text-white hover:bg-[var(--blue-d)]";
const btnOut =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-[18px] py-2.5 text-[13.5px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
