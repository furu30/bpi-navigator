"use client";

import { useState } from "react";
import type { ReproPlan, ReproPlanInput, ReproType } from "@/lib/types";

const METHODS = [
  "マニュアル化",
  "手順書作成",
  "判断基準の言語化",
  "事例DB構築",
  "OJTシナリオ整備",
  "動画教材作成",
  "その他",
];

interface Props {
  taskId: string;
  taskName: string;
  type: ReproType; // 受け継ぎ方（ボタン由来／既存planの種別）
  initial?: ReproPlan | null; // 編集時の既存値
  onClose: () => void;
  onSubmit: (input: ReproPlanInput) => void;
}

// 育成・定着計画モーダル。開いている時だけマウントして使う。
export function ReproPlanModal({
  taskId,
  taskName,
  type,
  initial,
  onClose,
  onSubmit,
}: Props) {
  const [content, setContent] = useState(initial?.content ?? "");
  const [person, setPerson] = useState(initial?.person ?? "");
  const [target, setTarget] = useState(initial?.target ?? "");
  const [date, setDate] = useState(initial?.date ?? "");
  const [method, setMethod] = useState(initial?.method ?? METHODS[0]);
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
      type,
      content: c,
      person: p,
      target: target.trim() || undefined,
      date,
      method,
    });
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
          {initial ? "育成・定着計画を編集" : "育成・定着計画"}
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
          <div className="mb-3 grid grid-cols-2 gap-3">
            <Field label="対象業務">
              <input value={taskName} readOnly className={`${inputCls} bg-[#f8fafc]`} />
            </Field>
            <Field label="種別">
              <input value={type} readOnly className={`${inputCls} bg-[#f8fafc]`} />
            </Field>
          </div>

          <Field label="内容" required>
            <textarea
              value={content}
              onChange={(e) => setContent(e.target.value)}
              rows={2}
              placeholder="どう標準化／形式知化／教育するか"
              className={inputCls}
            />
          </Field>

          <div className="grid grid-cols-3 gap-3">
            <Field label="担当者" required>
              <input
                value={person}
                onChange={(e) => setPerson(e.target.value)}
                placeholder="進める人"
                className={inputCls}
              />
            </Field>
            <Field label="対象者（誰に）">
              <input
                value={target}
                onChange={(e) => setTarget(e.target.value)}
                placeholder="若手・新人など"
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
            className="py-1 text-[12.5px] font-bold text-[var(--teal-d)]"
          >
            {showDetail ? "－" : "＋"} 詳細（任意：手段・教材）
          </button>
          {showDetail && (
            <div className="mt-2.5 rounded-[10px] border border-dashed border-[var(--line)] bg-[#f8fafc] p-3.5">
              <Field label="手段・教材">
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
            className="rounded-[9px] bg-[var(--teal)] px-[18px] py-2.5 text-[13.5px] font-bold text-white hover:bg-[var(--teal-d)]"
          >
            保存
          </button>
        </div>
      </div>
    </div>
  );
}

const inputCls =
  "w-full rounded-lg border-[1.5px] border-[var(--line)] px-[11px] py-[9px] text-[13.5px] focus:border-[var(--teal)] focus:outline-none";

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
