"use client";

import { monthly } from "@/lib/calc";
import { curTasks, groupsOf, useStore } from "@/lib/store";

// S-C1 業務の棚卸し（共通）— ステップ1〜3 の基盤版。
// ここでは「業務名切替の絞り込みが効くこと」を確認する。
// 追加/編集/削除モーダル・保存/読込・名称変更/削除はステップ4で実装する。

export default function StockPage() {
  const { state, hydrated, setCurrentGroup } = useStore();

  if (!hydrated) {
    return <div className="text-[var(--muted)]">読み込み中…</div>;
  }

  const groups = groupsOf(state);

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
        />
      ) : (
        <AllGroupsView groups={groups} onEnter={setCurrentGroup} />
      )}

      <div className="mt-3 rounded-lg border border-dashed border-[var(--line)] bg-white px-4 py-3 text-[12px] text-[var(--muted)]">
        ※ 基盤版（ステップ1〜3）です。業務の追加／編集／削除・保存／読込・業務名の名称変更／削除はステップ4で実装します。
      </div>
    </>
  );
}

function AllGroupsView({
  groups,
  onEnter,
}: {
  groups: string[];
  onEnter: (g: string) => void;
}) {
  const { state } = useStore();

  if (state.tasks.length === 0) {
    return (
      <div className="rounded-xl border border-dashed border-[var(--line)] bg-white p-[34px] text-center text-[var(--muted)]">
        まだ業務がありません。「＋ 業務名」から始めましょう。
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-4">
      <h2 className="text-[20px] font-extrabold">
        業務一覧（{groups.length}業務名・作業{state.tasks.length}件）
      </h2>
      {groups.map((g) => {
        const items = state.tasks.filter((t) => t.group === g);
        const totMin = items.reduce((s, t) => s + monthly(t), 0);
        return (
          <div
            key={g}
            className="overflow-hidden rounded-xl border border-[var(--line)] bg-white shadow-[0_1px_3px_rgba(0,0,0,.04)]"
          >
            <div className="flex flex-wrap items-center justify-between gap-2.5 border-b border-[var(--blue-line)] bg-[var(--blue-soft)] px-4 py-[11px]">
              <div>
                <b className="text-[15px] text-[var(--blue-d)]">📁 {g}</b>{" "}
                <span className="text-[12.5px] text-[var(--muted)]">
                  （作業{items.length}件・{totMin}分/月）
                </span>
              </div>
              <button
                type="button"
                onClick={() => onEnter(g)}
                className="rounded-lg border-[1.5px] border-[var(--line)] bg-white px-3 py-1.5 text-[12px] font-bold text-[var(--navy)] hover:border-[#94a3b8]"
              >
                この業務名で作業 →
              </button>
            </div>
            <ul className="divide-y divide-[#f1f5f9]">
              {items.map((t) => (
                <li
                  key={t.id}
                  className="flex items-center justify-between px-4 py-2.5 text-[13px]"
                >
                  <span className="font-bold">{t.content}</span>
                  <span className="text-[12px] text-[var(--muted)]">
                    {t.category} / {t.person} / {monthly(t)}分/月 ・ {t.workType}
                  </span>
                </li>
              ))}
            </ul>
          </div>
        );
      })}
    </div>
  );
}

function SingleGroupView({
  group,
  onBack,
}: {
  group: string;
  onBack: () => void;
}) {
  const { state } = useStore();
  const items = curTasks(state);
  const totMin = items.reduce((s, t) => s + monthly(t), 0);

  return (
    <div className="flex flex-col gap-3">
      <button
        type="button"
        onClick={onBack}
        className="self-start rounded-lg border-[1.5px] border-[var(--line)] bg-white px-3 py-1.5 text-[12px] font-bold text-[var(--navy)] hover:border-[#94a3b8]"
      >
        ← すべての業務名（全体一覧）に戻る
      </button>
      <h2 className="text-[20px] font-extrabold">
        📁 {group}（作業{items.length}件・{totMin}分/月）
      </h2>
      {items.length ? (
        <ul className="divide-y divide-[#f1f5f9] overflow-hidden rounded-xl border border-[var(--line)] bg-white">
          {items.map((t) => (
            <li
              key={t.id}
              className="flex items-center justify-between px-4 py-2.5 text-[13px]"
            >
              <span className="font-bold">{t.content}</span>
              <span className="text-[12px] text-[var(--muted)]">
                {t.category} / {t.person} / {monthly(t)}分/月 ・ {t.workType}
              </span>
            </li>
          ))}
        </ul>
      ) : (
        <div className="rounded-xl border border-dashed border-[var(--line)] bg-white p-[34px] text-center text-[var(--muted)]">
          この業務名にはまだ作業がありません。（追加はステップ4で実装）
        </div>
      )}
    </div>
  );
}
