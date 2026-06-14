"use client";

import Link from "next/link";
import { useState } from "react";
import { monthly, totalScore } from "@/lib/calc";
import { curTasks, useStore } from "@/lib/store";
import type { Task } from "@/lib/types";

// A-2 問題の見える化（影響度×頻度の2軸スコア・即時更新、問題カテゴリ別パレート）
export default function AnalyzePage() {
  const { state, hydrated, setScore, toggleProblem } = useStore();
  const [showDetail, setShowDetail] = useState(false);
  if (!hydrated) return <div className="text-[var(--muted)]">読み込み中…</div>;

  const ts = curTasks(state);
  const problemCats = state.masters.problemCategories.map((p) => p.name);

  // パレート（問題カテゴリ別の件数・多い順）
  const counts: Record<string, number> = {};
  ts.forEach((t) => t.problems.forEach((p) => (counts[p] = (counts[p] || 0) + 1)));
  const pareto = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  const max = pareto.length ? pareto[0][1] : 1;

  return (
    <>
      <div className="mb-5 flex gap-2.5 rounded-xl border border-[var(--blue-line)] bg-[var(--blue-soft)] px-4 py-3 text-[13px] text-[#1e3a8a]">
        🔍
        <div>
          <b className="text-[var(--blue-d)]">気づきポイント</b>
          ：影響度と頻度で採点し、問題を分類します。<b>既定は2軸</b>
          （影響度×頻度）で十分。細かい分析は折りたたみに格納しています。
        </div>
      </div>

      {/* パレート */}
      <div className="mb-[18px] rounded-xl border border-[var(--line)] bg-white p-[18px] shadow-[0_1px_3px_rgba(0,0,0,.04)]">
        <h3 className="text-[15px] font-extrabold">問題カテゴリ別パレート（多い順）</h3>
        <div className="mt-2.5">
          {pareto.length ? (
            pareto.map(([p, c]) => (
              <div
                key={p}
                className="mb-[7px] flex items-center gap-2.5 text-[12.5px]"
              >
                <span className="w-[130px] flex-shrink-0 text-right">{p}</span>
                <span className="flex-1 rounded bg-[#f1f5f9]">
                  <span
                    className="block h-[18px] rounded bg-[var(--blue)]"
                    style={{ width: `${(c / max) * 100}%` }}
                  />
                </span>
                <b>{c}</b>
              </div>
            ))
          ) : (
            <div className="text-[var(--muted)]">問題カテゴリ未選択</div>
          )}
        </div>
        <div className="mt-3.5">
          <button
            type="button"
            onClick={() => setShowDetail((v) => !v)}
            className="py-1 text-[12.5px] font-bold text-[var(--blue)]"
          >
            {showDetail ? "－" : "＋"} 詳しく見る（時間分析・ヒートマップ・バブル）
          </button>
          {showDetail && (
            <div className="mt-2.5 rounded-[10px] border border-dashed border-[var(--line)] bg-[#f8fafc] p-3.5 text-[12.5px] text-[var(--muted)]">
              【上級者向け】付加価値/非付加価値の時間分析、プロセス×問題のヒートマップ、効果×難易度バブルチャートをここに格納します（Phase2）。最初の画面では見せず、必要な人だけ開く設計です。
            </div>
          )}
        </div>
      </div>

      <h2 className="mb-3 text-[20px] font-extrabold">業務ごとのスコアリング</h2>
      {ts.length ? (
        ts.map((t) => (
          <ScoreCard
            key={t.id}
            task={t}
            problemCats={problemCats}
            onScore={setScore}
            onToggleProblem={toggleProblem}
          />
        ))
      ) : (
        <div className="rounded-xl border border-dashed border-[var(--line)] bg-white p-[34px] text-center text-[var(--muted)]">
          業務がありません。棚卸しで登録してください。
        </div>
      )}

      <div className="mt-6 flex justify-between border-t border-[var(--line)] pt-[18px]">
        <Link href="/a/dashboard" className={btnOut}>
          ← ダッシュボード
        </Link>
        <Link href="/a/ecrs" className={btnPri}>
          ECRSで改善案 →
        </Link>
      </div>
    </>
  );
}

function ScoreCard({
  task,
  problemCats,
  onScore,
  onToggleProblem,
}: {
  task: Task;
  problemCats: string[];
  onScore: (id: string, key: "impact" | "freq", value: number) => void;
  onToggleProblem: (id: string, problem: string) => void;
}) {
  return (
    <div className="mb-3.5 rounded-xl border border-[var(--line)] bg-white p-[18px] shadow-[0_1px_3px_rgba(0,0,0,.04)]">
      <div className="mb-2 flex items-start justify-between gap-3">
        <div>
          <h3 className="text-[15px] font-extrabold">{task.content}</h3>
          <div className="text-[12px] text-[var(--muted)]">
            {task.category} / {task.person} / {monthly(task)}分/月
          </div>
        </div>
        <div className="text-right">
          <div className="text-[11px] text-[var(--muted)]">優先度</div>
          <div className="text-[22px] font-extrabold text-[var(--blue)]">
            {totalScore(task.scores).toFixed(1)}
          </div>
        </div>
      </div>

      <Slider
        label="影響度"
        value={task.scores.impact}
        onChange={(v) => onScore(task.id, "impact", v)}
      />
      <Slider
        label="頻度"
        value={task.scores.freq}
        onChange={(v) => onScore(task.id, "freq", v)}
      />

      <div className="mt-2">
        <div className="mb-1 text-[11.5px] font-bold text-[var(--muted)]">
          問題カテゴリ
        </div>
        {problemCats.map((p) => {
          const on = task.problems.includes(p);
          return (
            <button
              key={p}
              type="button"
              onClick={() => onToggleProblem(task.id, p)}
              className={`mr-1 mt-[3px] inline-block rounded-full border-[1.5px] px-[11px] py-[3px] text-[11.5px] ${
                on
                  ? "border-[var(--blue)] bg-[var(--blue)] text-white"
                  : "border-[var(--line)] text-[var(--ink)]"
              }`}
            >
              {p}
            </button>
          );
        })}
      </div>
    </div>
  );
}

function Slider({
  label,
  value,
  onChange,
}: {
  label: string;
  value: number;
  onChange: (v: number) => void;
}) {
  return (
    <div className="my-1.5 flex items-center gap-3">
      <label className="w-20 text-[12.5px] font-bold text-[#475569]">
        {label}
      </label>
      <input
        type="range"
        min={1}
        max={5}
        value={value}
        onChange={(e) => onChange(Number(e.target.value))}
        className="flex-1"
      />
      <span className="w-6 text-center font-extrabold text-[var(--blue)]">
        {value}
      </span>
    </div>
  );
}

const btnPri =
  "rounded-[9px] bg-[var(--blue)] px-[18px] py-2.5 text-[13.5px] font-bold text-white hover:bg-[var(--blue-d)]";
const btnOut =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-[18px] py-2.5 text-[13.5px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
