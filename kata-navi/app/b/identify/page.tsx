"use client";

import Link from "next/link";
import { useState } from "react";
import { TacitModal } from "@/components/TacitModal";
import { tacitDone } from "@/lib/calc";
import { curTasks, useStore } from "@/lib/store";
import type { Task, WorkType } from "@/lib/types";

const WORK_TYPES: WorkType[] = ["作業", "感覚", "混在", "未判定"];

// B-2 切り分け・暗黙知（種別の確認/変更＋感覚業務の「5つの問い」）
export default function BIdentifyPage() {
  const { state, hydrated, setWorkType, setTacit } = useStore();
  const [tacitTask, setTacitTask] = useState<Task | null>(null);
  if (!hydrated) return <div className="text-[var(--muted)]">読み込み中…</div>;

  const ts = curTasks(state);

  return (
    <>
      <div className="mb-5 flex gap-2.5 rounded-xl border border-[var(--teal-line)] bg-[var(--teal-soft)] px-4 py-3 text-[13px] text-[#134e4a]">
        🧠
        <div>
          <b className="text-[var(--teal-d)]">気づきポイント</b>
          ：感覚業務は、ベテランの頭の中を<b>「5つの問い」</b>
          で言葉にします。ここで引き出した内容が、形式知化・教育の元になります。
        </div>
      </div>

      {ts.length ? (
        ts.map((t) => {
          const isKankaku = t.workType === "感覚" || t.workType === "混在";
          return (
            <div
              key={t.id}
              className="mb-3.5 rounded-xl border border-[var(--line)] bg-white p-[18px] shadow-[0_1px_3px_rgba(0,0,0,.04)]"
            >
              <div className="mb-1.5 flex flex-wrap items-start justify-between gap-3">
                <div>
                  <h3 className="text-[15px] font-extrabold">{t.content}</h3>
                  <div className="text-[12px] text-[var(--muted)]">
                    {t.category} / {t.person}
                  </div>
                </div>
                <select
                  value={t.workType}
                  onChange={(e) => setWorkType(t.id, e.target.value as WorkType)}
                  className="rounded-lg border-[1.5px] border-[var(--line)] px-2.5 py-1.5 text-[13px] focus:border-[var(--teal)] focus:outline-none"
                >
                  {WORK_TYPES.map((w) => (
                    <option key={w}>{w}</option>
                  ))}
                </select>
              </div>

              {isKankaku ? (
                <div className="flex flex-wrap items-center justify-between gap-2 rounded-[10px] border border-[var(--teal-line)] bg-[var(--teal-soft)] px-3 py-2.5">
                  <span className="text-[12.5px] text-[var(--teal-d)]">
                    🧠{" "}
                    {tacitDone(t)
                      ? "判断軸を言語化済み"
                      : "判断軸がまだ言語化されていません"}
                  </span>
                  <button
                    type="button"
                    onClick={() => setTacitTask(t)}
                    className="rounded-[9px] bg-[var(--teal)] px-3 py-1.5 text-[12px] font-bold text-white hover:bg-[var(--teal-d)]"
                  >
                    {tacitDone(t) ? "5つの問いを編集" : "5つの問いを引き出す"}
                  </button>
                </div>
              ) : (
                <div className="text-[12px] text-[var(--muted)]">
                  作業（標準化対象）。B-3で手順化します。
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
        <Link href="/b/dashboard" className={btnOut}>
          ← ダッシュボード
        </Link>
        <Link href="/b/route" className={btnPri}>
          標準化／形式知化・教育 →
        </Link>
      </div>

      {tacitTask && (
        <TacitModal
          task={tacitTask}
          onClose={() => setTacitTask(null)}
          onSubmit={(tacit) => {
            setTacit(tacitTask.id, tacit);
            setTacitTask(null);
          }}
        />
      )}
    </>
  );
}

const btnPri =
  "rounded-[9px] bg-[var(--teal)] px-[18px] py-2.5 text-[13.5px] font-bold text-white hover:bg-[var(--teal-d)]";
const btnOut =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-[18px] py-2.5 text-[13.5px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
