"use client";

import { useState } from "react";
import type { Plan, PlanInput } from "@/lib/types";

const METHODS = [
  "手順変更",
  "ツール導入",
  "システム化",
  "廃止",
  "マニュアル化",
  "その他",
];

interface Props {
  /** 対象業務のID（A-3：対象タスク／A-4：既存planのtaskId） */
  taskId: string;
  /** 対象業務の表示名 */
  taskName: string;
  /** 編集時の既存値（null＝新規追加） */
  initial?: Plan | null;
  onClose: () => void;
  onSubmit: (input: PlanInput) => void;
}

// 開いているときだけマウントして使う（マウント時の初期値でリセット）。
export function PlanModal({
  taskId,
  taskName,
  initial,
  onClose,
  onSubmit,
}: Props) {
  const [content, setContent] = useState(initial?.content ?? "");
  const [person, setPerson] = useState(initial?.person ?? "");
  const [date, setDate] = useState(initial?.date ?? "");
  const [method, setMethod] = useState(initial?.method ?? METHODS[0]);
  const [effect, setEffect] = useState(initial?.effect ?? "");
  const [showDetail, setShowDetail] = useState(!!initial);
  const [error, setError] = useState("");

  function handleSubmit() {
    const c = content.trim();
    const p = person.trim();
    if (!c || !p || !date) {
      setError("内容・担当・期日は必須です。");
      return;
    }
    onSubmit({
      taskId,
      taskName,
      content: c,
      person: p,
      date,
      method,
      effect: effect.trim() || undefined,
    });
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-start justify-center overflow-auto bg-[rgba(15,23,42,.5)] px-4 py-10"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="w-full max-w-[560px] rounded-2xl bg-white shadow-[0_20px_50px_rgba(0,0,0,.3)]">
        <h3 className="flex items-center justify-between border-b border-[var(--line)] px-5 py-4 text-[16px] font-extrabold">
          {initial ? "改善計画を編集" : "改善計画"}
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
          <Field label="対象業務">
            <input value={taskName} readOnly className={`${inputCls} bg-[#f8fafc]`} />
          </Field>
          <Field label="改善内容" required>
            <textarea
              value={content}
              onChange={(e) => setContent(e.target.value)}
              rows={2}
              placeholder="何をどう変えるか"
              className={inputCls}
            />
          </Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label="担当者" required>
              <input
                value={person}
                onChange={(e) => setPerson(e.target.value)}
                placeholder="責任者"
                className={inputCls}
              />
            </Field>
            <Field label="完了目標日" required>
              <input
                type="date"
                value={date}
                onChange={(e) => setDate(e.target.value)}
                className={inputCls}
              />
            </Field>
          </div>

          <button
            type="button"
            onClick={() => setShowDetail((v) => !v)}
            className="py-1 text-[12.5px] font-bold text-[var(--blue)]"
          >
            {showDetail ? "－" : "＋"} 詳細（任意：手段・期待効果）
          </button>
          {showDetail && (
            <div className="mt-2.5 rounded-[10px] border border-dashed border-[var(--line)] bg-[#f8fafc] p-3.5">
              <Field label="改善手段">
                <select
                  value={method}
                  onChange={(e) => setMethod(e.target.value)}
                  className={inputCls}
                >
                  {METHODS.map((m) => (
                    <option key={m}>{m}</option>
                  ))}
                </select>
              </Field>
              <Field label="期待効果（定量）">
                <input
                  value={effect}
                  onChange={(e) => setEffect(e.target.value)}
                  placeholder="例：月60分削減"
                  className={inputCls}
                />
              </Field>
            </div>
          )}

          {error && (
            <p className="mt-3 text-[12.5px] font-bold text-[var(--red)]">
              {error}
            </p>
          )}
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
            className="rounded-[9px] bg-[var(--blue)] px-[18px] py-2.5 text-[13.5px] font-bold text-white hover:bg-[var(--blue-d)]"
          >
            保存
          </button>
        </div>
      </div>
    </div>
  );
}

const inputCls =
  "w-full rounded-lg border-[1.5px] border-[var(--line)] px-[11px] py-[9px] text-[13.5px] focus:border-[var(--blue)] focus:outline-none";

function Field({
  label,
  required,
  children,
}: {
  label: string;
  required?: boolean;
  children: React.ReactNode;
}) {
  return (
    <div className="mb-3">
      <label className="mb-1 block text-[12.5px] font-bold text-[#475569]">
        {label}
        {required && <span className="text-[var(--red)]"> *</span>}
      </label>
      {children}
    </div>
  );
}
