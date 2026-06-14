"use client";

import { useState } from "react";
import type { Tacit, Task } from "@/lib/types";

const QUESTIONS: { key: keyof Tacit; label: string; placeholder: string }[] = [
  { key: "q1", label: "① 何を見て判断していますか？", placeholder: "例：刃先の摩耗、表面の光沢、削り音" },
  { key: "q2", label: "② OK / NG を分ける基準は？", placeholder: "例：刃欠け0.05mm以上はNG" },
  { key: "q3", label: "③ 迷うとき・例外的なケースは？", placeholder: "例：材料が硬めの時" },
  { key: "q4", label: "④ ベテランが無意識にやっている工夫は？", placeholder: "例：切粉の色で温度を判断" },
  { key: "q5", label: "⑤ 若手にはどう教えていますか？", placeholder: "例：3ヶ月OJTで隣で見せる" },
];

interface Props {
  task: Task;
  onClose: () => void;
  onSubmit: (tacit: Tacit) => void;
}

// 暗黙知（5つの問い）モーダル。開いている時だけマウントして使う。
export function TacitModal({ task, onClose, onSubmit }: Props) {
  const [vals, setVals] = useState<Tacit>(() => ({ ...(task.tacit ?? {}) }));

  function handleSubmit() {
    // 空欄は undefined に正規化して保存（言語化済み判定が空文字で誤判定しないように）
    const tacit: Tacit = {};
    for (const { key } of QUESTIONS) {
      const v = vals[key]?.trim();
      if (v) tacit[key] = v;
    }
    onSubmit(tacit);
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-start justify-center overflow-auto bg-[rgba(15,23,42,.5)] px-4 py-10"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="w-full max-w-[580px] rounded-2xl bg-white shadow-[0_20px_50px_rgba(0,0,0,.3)]">
        <h3 className="flex items-center justify-between border-b border-[var(--line)] px-5 py-4 text-[16px] font-extrabold">
          🧠 判断ポイントの引き出し（5つの問い）
          <button
            type="button"
            onClick={onClose}
            className="text-[22px] leading-none text-[#94a3b8] hover:text-[var(--muted)]"
            aria-label="閉じる"
          >
            ×
          </button>
        </h3>

        <div className="px-5 py-[18px]">
          <div className="mb-3">
            <label className="mb-1 block text-[12.5px] font-bold text-[#475569]">
              対象業務
            </label>
            <input
              value={task.content}
              readOnly
              className="w-full rounded-lg border-[1.5px] border-[var(--line)] bg-[#f8fafc] px-[11px] py-[9px] text-[13.5px]"
            />
          </div>

          {QUESTIONS.map((q) => (
            <div key={q.key} className="mb-2">
              <label className="mb-0.5 block text-[12px] font-bold text-[var(--teal-d)]">
                {q.label}
              </label>
              <input
                value={vals[q.key] ?? ""}
                onChange={(e) =>
                  setVals((p) => ({ ...p, [q.key]: e.target.value }))
                }
                placeholder={q.placeholder}
                className="w-full rounded-lg border-[1.5px] border-[var(--line)] px-[11px] py-[9px] text-[13.5px] focus:border-[var(--teal)] focus:outline-none"
              />
            </div>
          ))}
        </div>

        <div className="flex justify-end gap-2.5 border-t border-[var(--line)] px-5 py-3.5">
          <button
            type="button"
            onClick={onClose}
            className="rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-[18px] py-2.5 text-[13.5px] font-bold text-[var(--navy)] hover:border-[#94a3b8]"
          >
            キャンセル
          </button>
          <button
            type="button"
            onClick={handleSubmit}
            className="rounded-[9px] bg-[var(--teal)] px-[18px] py-2.5 text-[13.5px] font-bold text-white hover:bg-[var(--teal-d)]"
          >
            保存
          </button>
        </div>
      </div>
    </div>
  );
}
