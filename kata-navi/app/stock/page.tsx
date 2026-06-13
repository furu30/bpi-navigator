"use client";

import Link from "next/link";
import { useState } from "react";
import { MmmTags, TypeTag } from "@/components/Tags";
import { TaskModal } from "@/components/TaskModal";
import { monthly } from "@/lib/calc";
import { exportGroup } from "@/lib/io";
import { curTasks, groupPlans, groupsOf, useStore } from "@/lib/store";
import type { Task, TaskInput } from "@/lib/types";

// S-C1 業務の棚卸し（共通）— 全体一覧／単一業務名の2モード。
// 業務の追加・編集・削除、業務名の名称変更・削除・保存。

type ModalState = { mode: "new" } | { mode: "edit"; task: Task } | null;

export default function StockPage() {
  const store = useStore();
  const { state, hydrated, setCurrentGroup, addTask, updateTask } = store;
  const [modal, setModal] = useState<ModalState>(null);

  if (!hydrated) {
    return <div className="text-[var(--muted)]">読み込み中…</div>;
  }

  const groups = groupsOf(state);

  function handleSubmit(input: TaskInput, editingId: string | null) {
    if (editingId) updateTask(editingId, input);
    else addTask(input);
    setModal(null);
  }

  return (
    <>
      <div className="mb-5 flex gap-2.5 rounded-xl border border-[var(--blue-line)] bg-[var(--blue-soft)] px-4 py-3 text-[13px] text-[#1e3a8a]">
        💡
        <div>
          <b className="text-[var(--blue-d)]">気づきポイント</b>
          ：上部の<b>「業務名」で切り替える</b>
          と、その業務名の作業だけが棚卸し〜効果検証まで表示されます。業務名ごとに保存・管理できます。
        </div>
      </div>

      {state.currentGroup ? (
        <SingleGroupView
          group={state.currentGroup}
          onBack={() => setCurrentGroup(null)}
          onAdd={() => setModal({ mode: "new" })}
          onEdit={(task) => setModal({ mode: "edit", task })}
        />
      ) : (
        <AllGroupsView
          groups={groups}
          onEnter={setCurrentGroup}
          onAdd={() => setModal({ mode: "new" })}
          onEdit={(task) => setModal({ mode: "edit", task })}
        />
      )}

      <div className="mt-6 flex justify-end border-t border-[var(--line)] pt-[18px]">
        <Link href="/a/dashboard" className={btnPri}>
          改善をはじめる →
        </Link>
      </div>

      {modal && (
        <TaskModal
          editingTask={modal.mode === "edit" ? modal.task : null}
          defaultGroup={state.currentGroup}
          existingGroups={groups}
          processCategories={state.masters.processCategories}
          onClose={() => setModal(null)}
          onSubmit={handleSubmit}
        />
      )}
    </>
  );
}

// ===== 全体一覧モード =====
function AllGroupsView({
  groups,
  onEnter,
  onAdd,
  onEdit,
}: {
  groups: string[];
  onEnter: (g: string) => void;
  onAdd: () => void;
  onEdit: (task: Task) => void;
}) {
  const { state, deleteTask } = useStore();

  return (
    <>
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <h2 className="text-[20px] font-extrabold">
          業務一覧（{groups.length}業務名・作業{state.tasks.length}件）
        </h2>
        <div className="flex flex-wrap items-center gap-2.5">
          <button
            type="button"
            disabled
            title="ステップ7で実装予定"
            className={`${btnAi} disabled:opacity-40`}
          >
            🤖 AIアシスト
          </button>
          <button
            type="button"
            disabled
            title="Phase2で実装予定"
            className={`${btnOutSm} disabled:opacity-40`}
          >
            CSVインポート
          </button>
          <button type="button" onClick={onAdd} className={btnPri}>
            ＋ 業務を追加
          </button>
        </div>
      </div>

      {state.tasks.length === 0 ? (
        <Empty>「＋ 業務を追加」または上部の「＋ 業務名」で登録しましょう。</Empty>
      ) : (
        <div className="flex flex-col gap-4">
          {groups.map((g) => {
            const items = state.tasks.filter((t) => t.group === g);
            const totMin = items.reduce((s, t) => s + monthly(t), 0);
            const gp = groupPlans(state, g);
            const done = gp.filter((p) => p.status === "完了").length;
            const prog = gp.length ? `改善 ${done}/${gp.length}` : "改善 未着手";
            return (
              <div
                key={g}
                className="overflow-hidden rounded-xl border border-[var(--line)] bg-white shadow-[0_1px_3px_rgba(0,0,0,.04)]"
              >
                <div className="flex flex-wrap items-center justify-between gap-2.5 border-b border-[var(--blue-line)] bg-[var(--blue-soft)] px-4 py-[11px]">
                  <div>
                    <b className="text-[15px] text-[var(--blue-d)]">📁 {g}</b>{" "}
                    <span className="text-[12.5px] text-[var(--muted)]">
                      （作業{items.length}件・{totMin}分/月 ・ {prog}）
                    </span>
                  </div>
                  <div className="flex gap-2">
                    <button
                      type="button"
                      onClick={() => onEnter(g)}
                      className={btnOutSm}
                    >
                      この業務名で作業 →
                    </button>
                    <button
                      type="button"
                      onClick={() => exportGroup(state, g)}
                      className={btnOutSm}
                    >
                      💾 保存
                    </button>
                  </div>
                </div>
                <TaskTable
                  items={items}
                  onEdit={onEdit}
                  onDelete={(id) => {
                    if (window.confirm("この作業を削除しますか？")) deleteTask(id);
                  }}
                />
              </div>
            );
          })}
        </div>
      )}
    </>
  );
}

// ===== 単一業務名モード =====
function SingleGroupView({
  group,
  onBack,
  onAdd,
  onEdit,
}: {
  group: string;
  onBack: () => void;
  onAdd: () => void;
  onEdit: (task: Task) => void;
}) {
  const { state, deleteTask, renameGroup, deleteGroup } = useStore();
  const items = curTasks(state);
  const totMin = items.reduce((s, t) => s + monthly(t), 0);

  function handleRename() {
    const next = window.prompt("業務名を変更", group);
    if (next && next.trim()) renameGroup(group, next.trim());
  }

  function handleDelete() {
    const cnt = state.tasks.filter((t) => t.group === group).length;
    if (
      window.confirm(
        `業務名「${group}」と、その作業${cnt}件・関連する改善計画をすべて削除します。よろしいですか？`,
      )
    ) {
      deleteGroup(group);
    }
  }

  return (
    <>
      <button type="button" onClick={onBack} className={`mb-3 ${btnOutSm}`}>
        ← すべての業務名（全体一覧）に戻る
      </button>
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <h2 className="text-[20px] font-extrabold">
          📁 {group}（作業{items.length}件・{totMin}分/月）
        </h2>
        <div className="flex flex-wrap gap-2">
          <button type="button" onClick={handleRename} className={btnOutSm}>
            ✏ 名称変更
          </button>
          <button type="button" onClick={handleDelete} className={btnOutSm}>
            🗑 削除
          </button>
          <button
            type="button"
            onClick={() => exportGroup(state, group)}
            className={btnOutSm}
          >
            💾 保存
          </button>
          <button type="button" onClick={onAdd} className={btnPri}>
            ＋ 業務を追加
          </button>
        </div>
      </div>

      {items.length ? (
        <div className="overflow-hidden rounded-xl border border-[var(--line)] bg-white shadow-[0_1px_3px_rgba(0,0,0,.05)]">
          <TaskTable
            items={items}
            onEdit={onEdit}
            onDelete={(id) => {
              if (window.confirm("この作業を削除しますか？")) deleteTask(id);
            }}
          />
        </div>
      ) : (
        <Empty>
          この業務名にはまだ作業がありません。「＋ 業務を追加」で登録しましょう。
        </Empty>
      )}
    </>
  );
}

// ===== 作業テーブル =====
function TaskTable({
  items,
  onEdit,
  onDelete,
}: {
  items: Task[];
  onEdit: (task: Task) => void;
  onDelete: (id: string) => void;
}) {
  return (
    <table className="w-full border-collapse text-[13px]">
      <thead>
        <tr>
          <th className={th}>作業</th>
          <th className={`${th} w-[86px]`}>種別</th>
          <th className={`${th} w-[110px]`}>ムダ等</th>
          <th className={`${th} w-[110px]`} />
        </tr>
      </thead>
      <tbody>
        {items.map((t) => (
          <tr key={t.id}>
            <td className={td}>
              <b>{t.content}</b>
              <div className="text-[12px] text-[var(--muted)]">
                {t.category} / {t.person} / {t.time}分・{t.freq}
              </div>
            </td>
            <td className={td}>
              <TypeTag type={t.workType} />
            </td>
            <td className={td}>
              <MmmTags mmm={t.mmm} />
            </td>
            <td className={td}>
              <div className="flex gap-1">
                <button
                  type="button"
                  onClick={() => onEdit(t)}
                  className={btnOutSm}
                >
                  編集
                </button>
                <button
                  type="button"
                  onClick={() => onDelete(t.id)}
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
  );
}

function Empty({ children }: { children: React.ReactNode }) {
  return (
    <div className="rounded-xl border border-dashed border-[var(--line)] bg-white p-[34px] text-center text-[var(--muted)]">
      {children}
    </div>
  );
}

// 共通クラス
const btnPri =
  "rounded-[9px] bg-[var(--blue)] px-[18px] py-2.5 text-[13.5px] font-bold text-white hover:bg-[var(--blue-d)]";
const btnOutSm =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-3 py-1.5 text-[12px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
const btnAi =
  "rounded-[9px] border border-[#ddd6fe] bg-[#ede9fe] px-3 py-1.5 text-[12px] font-bold text-[#6d28d9]";
const th =
  "border-b border-[var(--line)] bg-[#f1f5f9] px-3 py-2.5 text-left text-[12px] font-bold text-[#475569]";
const td =
  "border-b border-[#f1f5f9] px-3 py-2.5 align-middle";
