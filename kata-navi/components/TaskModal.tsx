"use client";

import { useId, useState } from "react";
import type { Freq, Task, TaskInput, WorkType } from "@/lib/types";

const FREQS: Freq[] = ["日次", "週次", "月次", "不定期"];
const WORK_TYPES: { value: WorkType; label: string }[] = [
  { value: "作業", label: "作業（標準化できる）" },
  { value: "感覚", label: "感覚（高度な判断）" },
  { value: "混在", label: "混在" },
  { value: "未判定", label: "未判定" },
];

interface Props {
  /** 編集対象（null＝新規追加） */
  editingTask: Task | null;
  /** 新規追加時の業務名初期値（現在選択中の業務名など） */
  defaultGroup?: string | null;
  existingGroups: string[];
  processCategories: string[];
  onClose: () => void;
  onSubmit: (input: TaskInput, editingId: string | null) => void;
}

// このコンポーネントは「開いているときだけ」マウントして使う。
// （マウント時の初期値で各入力を初期化する＝開くたびにリセットされる）
export function TaskModal({
  editingTask,
  defaultGroup,
  existingGroups,
  processCategories,
  onClose,
  onSubmit,
}: Props) {
  const listId = useId();
  const isEdit = !!editingTask;

  const [group, setGroup] = useState(
    editingTask?.group ?? defaultGroup ?? "",
  );
  const [content, setContent] = useState(editingTask?.content ?? "");
  const [person, setPerson] = useState(editingTask?.person ?? "");
  const [time, setTime] = useState<string>(String(editingTask?.time ?? 30));
  const [freq, setFreq] = useState<Freq>(editingTask?.freq ?? "日次");
  const [workType, setWorkType] = useState<WorkType>(
    editingTask?.workType ?? "作業",
  );
  const [category, setCategory] = useState(
    editingTask?.category ?? processCategories[0] ?? "",
  );
  const [muri, setMuri] = useState(!!editingTask?.mmm.muri);
  const [muda, setMuda] = useState(!!editingTask?.mmm.muda);
  const [mura, setMura] = useState(!!editingTask?.mmm.mura);
  const [showDetail, setShowDetail] = useState(false);
  const [error, setError] = useState("");

  function handleSubmit() {
    const g = group.trim();
    const c = content.trim();
    const p = person.trim();
    if (!g || !c || !p) {
      setError("業務名・作業内容・担当者は必須です。");
      return;
    }
    onSubmit(
      {
        group: g,
        content: c,
        person: p,
        time: Number(time) || 0,
        freq,
        workType,
        category,
        mmm: { muri, muda, mura },
      },
      editingTask?.id ?? null,
    );
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
          {isEdit ? "業務を編集" : "業務を追加"}
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
          <Field label="業務名（まとめ）" required>
            <input
              list={listId}
              value={group}
              onChange={(e) => setGroup(e.target.value)}
              placeholder="例：受注業務（既存から選ぶ／新規入力）"
              className={inputCls}
            />
            <datalist id={listId}>
              {existingGroups.map((g) => (
                <option key={g} value={g} />
              ))}
            </datalist>
          </Field>

          <Field label="作業内容（何を）" required>
            <input
              value={content}
              onChange={(e) => setContent(e.target.value)}
              placeholder="例：受注データの入力"
              className={inputCls}
            />
          </Field>

          <div className="grid grid-cols-3 gap-3">
            <Field label="担当者" required>
              <input
                value={person}
                onChange={(e) => setPerson(e.target.value)}
                placeholder="例：事務"
                className={inputCls}
              />
            </Field>
            <Field label="所要時間(分)" required>
              <input
                type="number"
                value={time}
                onChange={(e) => setTime(e.target.value)}
                className={inputCls}
              />
            </Field>
            <Field label="頻度" required>
              <select
                value={freq}
                onChange={(e) => setFreq(e.target.value as Freq)}
                className={inputCls}
              >
                {FREQS.map((f) => (
                  <option key={f}>{f}</option>
                ))}
              </select>
            </Field>
          </div>

          <Field label="業務種別" required>
            <select
              value={workType}
              onChange={(e) => setWorkType(e.target.value as WorkType)}
              className={inputCls}
            >
              {WORK_TYPES.map((w) => (
                <option key={w.value} value={w.value}>
                  {w.label}
                </option>
              ))}
            </select>
          </Field>

          <button
            type="button"
            onClick={() => setShowDetail((v) => !v)}
            className="py-1 text-[12.5px] font-bold text-[var(--blue)]"
          >
            {showDetail ? "－" : "＋"} 詳細（任意：プロセス区分・ムリムダムラ）
          </button>
          {showDetail && (
            <div className="mt-2.5 rounded-[10px] border border-dashed border-[var(--line)] bg-[#f8fafc] p-3.5">
              <Field label="プロセス区分">
                <select
                  value={category}
                  onChange={(e) => setCategory(e.target.value)}
                  className={inputCls}
                >
                  {processCategories.map((c) => (
                    <option key={c}>{c}</option>
                  ))}
                </select>
              </Field>
              <Field label="ムリ・ムダ・ムラ">
                <div className="flex gap-4">
                  <Check label="ムリ" checked={muri} onChange={setMuri} />
                  <Check label="ムダ" checked={muda} onChange={setMuda} />
                  <Check label="ムラ" checked={mura} onChange={setMura} />
                </div>
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

function Check({
  label,
  checked,
  onChange,
}: {
  label: string;
  checked: boolean;
  onChange: (v: boolean) => void;
}) {
  return (
    <label className="flex items-center gap-1.5 text-[13px]">
      <input
        type="checkbox"
        checked={checked}
        onChange={(e) => onChange(e.target.checked)}
        className="h-auto w-auto"
      />
      {label}
    </label>
  );
}
