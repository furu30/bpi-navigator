"use client";

import Link from "next/link";
import { useState } from "react";
import { ReproPlanModal } from "@/components/ReproPlanModal";
import { curTasks, useStore } from "@/lib/store";
import type { ReproPlanInput, ReproType, Task } from "@/lib/types";

type ModalState = { task: Task; type: ReproType } | null;

// B-3 標準化／形式知化・教育（種別別の受け継ぎ方ボタン → ReproPlanModal）
export default function BRoutePage() {
  const { state, hydrated, addReproPlan } = useStore();
  const [modal, setModal] = useState<ModalState>(null);
  if (!hydrated) return <div className="text-[var(--muted)]">読み込み中…</div>;

  const ts = curTasks(state);

  function handleSubmit(input: ReproPlanInput) {
    addReproPlan(input);
    setModal(null);
  }

  return (
    <>
      <div className="mb-5 flex gap-2.5 rounded-xl border border-[var(--teal-line)] bg-[var(--teal-soft)] px-4 py-3 text-[13px] text-[#134e4a]">
        🛠
        <div>
          <b className="text-[var(--teal-d)]">気づきポイント</b>：
          <b>作業＝標準化</b>（誰でもできる手順に）。
          <b>感覚＝形式知化（ルートA）or 教育（ルートB）</b>
          。業務の性質でルートを選びます。
        </div>
      </div>

      {ts.length ? (
        ts.map((t) => {
          const myPlans = state.reproPlans.filter((p) => p.taskId === t.id);
          return (
            <div
              key={t.id}
              className="mb-3.5 rounded-xl border border-[var(--line)] bg-white p-[18px] shadow-[0_1px_3px_rgba(0,0,0,.04)]"
            >
              <div className="mb-1.5">
                <h3 className="text-[15px] font-extrabold">{t.content}</h3>
                <div className="text-[12px] text-[var(--muted)]">
                  {t.workType} ・ {t.person} ・ {t.freq}
                </div>
              </div>
              <div className="mb-2 text-[12px] text-[var(--muted)]">
                受け継ぎ方のボタンを押すと、計画の入力に進みます。
              </div>

              <div className="flex flex-wrap gap-2">
                {t.workType === "作業" && (
                  <RouteBtn onClick={() => setModal({ task: t, type: "標準化" })}>
                    📘 標準化（手順化・マニュアル化）を計画 →
                  </RouteBtn>
                )}
                {(t.workType === "感覚" || t.workType === "混在") && (
                  <>
                    <RouteBtn
                      onClick={() => setModal({ task: t, type: "形式知化" })}
                    >
                      🅰 形式知化を計画（判断基準の言語化）→
                    </RouteBtn>
                    <RouteBtn onClick={() => setModal({ task: t, type: "教育" })}>
                      🅱 教育・OJTを計画 →
                    </RouteBtn>
                  </>
                )}
                {t.workType === "未判定" && (
                  <span className="text-[12px] text-[var(--muted)]">
                    ⚠ 未判定です。先に B-2 で「作業／感覚」を決めてください。
                  </span>
                )}
              </div>

              {myPlans.length > 0 && (
                <div className="mt-2.5 text-[12.5px] text-[var(--teal-d)]">
                  ✓ 計画あり（{myPlans.length}件）
                  <Link href="/b/plan" className="underline">
                    B-4で確認
                  </Link>
                </div>
              )}
            </div>
          );
        })
      ) : (
        <div className="rounded-xl border border-dashed border-[var(--line)] bg-white p-[34px] text-center text-[var(--muted)]">
          業務がありません。棚卸しで登録してください。
        </div>
      )}

      <div className="mt-6 flex justify-between border-t border-[var(--line)] pt-[18px]">
        <Link href="/b/identify" className={btnOut}>
          ← 切り分け・暗黙知
        </Link>
        <Link href="/b/plan" className={btnPri}>
          育成・定着計画 →
        </Link>
      </div>

      {modal && (
        <ReproPlanModal
          taskId={modal.task.id}
          taskName={modal.task.content}
          type={modal.type}
          onClose={() => setModal(null)}
          onSubmit={handleSubmit}
        />
      )}
    </>
  );
}

function RouteBtn({
  onClick,
  children,
}: {
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="rounded-lg border-[1.5px] border-[var(--teal-line)] bg-white px-3.5 py-2 text-[12.5px] font-bold text-[var(--teal-d)] hover:border-[var(--teal)] hover:bg-[var(--teal)] hover:text-white"
    >
      {children}
    </button>
  );
}

const btnPri =
  "rounded-[9px] bg-[var(--teal)] px-[18px] py-2.5 text-[13.5px] font-bold text-white hover:bg-[var(--teal-d)]";
const btnOut =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-[18px] py-2.5 text-[13.5px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
