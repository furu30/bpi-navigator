"use client";

import Link from "next/link";
import { MmmTags, TypeTag } from "@/components/Tags";
import { monthly } from "@/lib/calc";
import { curTasks, useStore } from "@/lib/store";

// A-1 改善ダッシュボード（curTasksを月間時間降順で一覧・統計）
export default function DashboardPage() {
  const { state, hydrated } = useStore();
  if (!hydrated) return <div className="text-[var(--muted)]">読み込み中…</div>;

  const ts = curTasks(state);
  const sorted = [...ts].sort((a, b) => monthly(b) - monthly(a));
  const withProb = ts.filter(
    (t) => t.problems.length || t.mmm.muri || t.mmm.muda || t.mmm.mura,
  ).length;
  const totMin = ts.reduce((s, t) => s + monthly(t), 0);

  return (
    <>
      <div className="mb-5 flex gap-2.5 rounded-xl border border-[var(--blue-line)] bg-[var(--blue-soft)] px-4 py-3 text-[13px] text-[#1e3a8a]">
        📊
        <div>
          <b className="text-[var(--blue-d)]">着手の見当をつける</b>
          ：月あたりの時間が大きく、ムダ・問題のある業務から手をつけると効果が出やすいです。
        </div>
      </div>

      <div className="mb-[18px] grid grid-cols-2 gap-3 sm:grid-cols-4">
        <Stat value={ts.length} label="登録業務" />
        <Stat value={withProb} label="問題あり" />
        <Stat value={totMin.toLocaleString()} label="合計(分/月)" />
        <Stat value={Math.round(totMin / 60)} label="合計(時間/月)" />
      </div>

      <h2 className="mb-3 text-[20px] font-extrabold">時間の大きい業務 TOP</h2>
      {sorted.length ? (
        <table className="w-full border-collapse overflow-hidden rounded-xl text-[13px] shadow-[0_1px_3px_rgba(0,0,0,.05)]">
          <thead>
            <tr>
              <th className={th}>業務</th>
              <th className={`${th} w-[80px]`}>種別</th>
              <th className={`${th} w-[90px]`}>時間/月</th>
              <th className={`${th} w-[120px]`}>ムダ等</th>
              <th className={`${th} w-[100px]`} />
            </tr>
          </thead>
          <tbody>
            {sorted.map((t) => (
              <tr key={t.id}>
                <td className={td}>
                  <b>{t.content}</b>
                  <div className="text-[12px] text-[var(--muted)]">
                    📁 {t.group} / {t.person}
                  </div>
                </td>
                <td className={td}>
                  <TypeTag type={t.workType} />
                </td>
                <td className={td}>
                  <b>{monthly(t)}</b>分/月
                </td>
                <td className={td}>
                  <MmmTags mmm={t.mmm} />
                </td>
                <td className={td}>
                  <Link href="/a/analyze" className={btnOutSm}>
                    分析へ →
                  </Link>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      ) : (
        <div className="rounded-xl border border-dashed border-[var(--line)] bg-white p-[34px] text-center text-[var(--muted)]">
          業務がありません。棚卸しで登録してください。
        </div>
      )}

      <div className="mt-6 flex justify-between border-t border-[var(--line)] pt-[18px]">
        <Link href="/stock" className={btnOut}>
          ← 棚卸し
        </Link>
        <Link href="/a/analyze" className={btnPri}>
          問題の見える化 →
        </Link>
      </div>
    </>
  );
}

function Stat({ value, label }: { value: React.ReactNode; label: string }) {
  return (
    <div className="rounded-xl border border-[var(--line)] bg-white p-4 text-center">
      <div className="text-[26px] font-extrabold text-[var(--blue)]">{value}</div>
      <div className="mt-0.5 text-[12px] text-[var(--muted)]">{label}</div>
    </div>
  );
}

const btnPri =
  "rounded-[9px] bg-[var(--blue)] px-[18px] py-2.5 text-[13.5px] font-bold text-white hover:bg-[var(--blue-d)]";
const btnOut =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-[18px] py-2.5 text-[13.5px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
const btnOutSm =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-3 py-1.5 text-[12px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
const th =
  "border-b border-[var(--line)] bg-[#f1f5f9] px-3 py-2.5 text-left text-[12px] font-bold text-[#475569]";
const td = "border-b border-[#f1f5f9] bg-white px-3 py-2.5 align-middle";
