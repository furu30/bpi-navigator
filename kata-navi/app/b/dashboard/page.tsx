"use client";

import Link from "next/link";
import { TypeTag } from "@/components/Tags";
import { attrition, tacitDone } from "@/lib/calc";
import { curTasks, useStore } from "@/lib/store";

// B-1 再現性ダッシュボード（作業/感覚/混在の件数・属人性割合・比率バー・止まると困る業務）
export default function BDashboardPage() {
  const { state, hydrated } = useStore();
  if (!hydrated) return <div className="text-[var(--muted)]">読み込み中…</div>;

  const ts = curTasks(state);
  const sagyo = ts.filter((t) => t.workType === "作業").length;
  const kankaku = ts.filter((t) => t.workType === "感覚").length;
  const konzai = ts.filter((t) => t.workType === "混在").length;
  const tot = sagyo + kankaku + konzai || 1;
  const zokujin = Math.round(((kankaku + konzai) / tot) * 100);
  const risky = ts
    .filter((t) => t.workType === "感覚" || t.workType === "混在")
    .sort((a, b) => attrition(b) - attrition(a));

  return (
    <>
      <div className="mb-5 flex gap-2.5 rounded-xl border border-[var(--teal-line)] bg-[var(--teal-soft)] px-4 py-3 text-[13px] text-[#134e4a]">
        📊
        <div>
          <b className="text-[var(--teal-d)]">属人性を見渡す</b>
          ：感覚業務（特定の人頼み）で頻度が高いものほど「止まると困る」業務。ここから受け継ぎましょう。
        </div>
      </div>

      <div className="mb-[18px] grid grid-cols-2 gap-3 sm:grid-cols-4">
        <Stat value={sagyo} label="作業（標準化）" />
        <Stat value={kankaku} label="感覚（再現）" />
        <Stat value={konzai} label="混在" />
        <Stat value={`${zokujin}%`} label="属人性の割合" />
      </div>

      <div className="mb-[18px] rounded-xl border border-[var(--line)] bg-white p-[18px] shadow-[0_1px_3px_rgba(0,0,0,.04)]">
        <h3 className="text-[15px] font-extrabold">作業 / 感覚の比率</h3>
        <div className="mt-2.5">
          <RatioBar label="作業（標準化）" value={sagyo} tot={tot} color="blue" />
          <RatioBar
            label="感覚＋混在（再現）"
            value={kankaku + konzai}
            tot={tot}
            color="teal"
          />
        </div>
      </div>

      <h2 className="mb-3 text-[20px] font-extrabold">
        止まると困る業務（属人性の高い順）
      </h2>
      {risky.length ? (
        <div className="overflow-x-auto rounded-xl shadow-[0_1px_3px_rgba(0,0,0,.05)]">
          <table className="w-full min-w-[560px] border-collapse text-[13px]">
            <thead>
              <tr>
                <th className={th}>業務</th>
                <th className={`${th} w-[80px]`}>種別</th>
                <th className={`${th} w-[60px]`}>頻度</th>
                <th className={`${th} w-[96px]`}>暗黙知</th>
                <th className={`${th} w-[100px]`} />
              </tr>
            </thead>
            <tbody>
              {risky.map((t) => (
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
                  <td className={td}>{t.freq}</td>
                  <td className={td}>
                    <TacitPill done={tacitDone(t)} />
                  </td>
                  <td className={td}>
                    <Link href="/b/identify" className={btnOutSm}>
                      引き出す →
                    </Link>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <div className="rounded-xl border border-dashed border-[var(--line)] bg-white p-[34px] text-center text-[var(--muted)]">
          感覚・混在の業務がありません。棚卸しで登録してください。
        </div>
      )}

      <div className="mt-6 flex justify-between border-t border-[var(--line)] pt-[18px]">
        <Link href="/stock" className={btnOut}>
          ← 棚卸し
        </Link>
        <Link href="/b/identify" className={btnPri}>
          切り分け・暗黙知 →
        </Link>
      </div>
    </>
  );
}

function Stat({ value, label }: { value: React.ReactNode; label: string }) {
  return (
    <div className="rounded-xl border border-[var(--line)] bg-white p-4 text-center">
      <div className="text-[26px] font-extrabold text-[var(--teal)]">{value}</div>
      <div className="mt-0.5 text-[12px] text-[var(--muted)]">{label}</div>
    </div>
  );
}

function RatioBar({
  label,
  value,
  tot,
  color,
}: {
  label: string;
  value: number;
  tot: number;
  color: "blue" | "teal";
}) {
  return (
    <div className="mb-[7px] flex items-center gap-2.5 text-[12.5px]">
      <span className="w-[140px] flex-shrink-0 text-right">{label}</span>
      <span className="flex-1 rounded bg-[#f1f5f9]">
        <span
          className={`block h-[18px] rounded ${
            color === "teal" ? "bg-[var(--teal)]" : "bg-[var(--blue)]"
          }`}
          style={{ width: `${(value / tot) * 100}%` }}
        />
      </span>
      <b>{value}</b>
    </div>
  );
}

function TacitPill({ done }: { done: boolean }) {
  return done ? (
    <span className="rounded-full bg-[#dcfce7] px-2 py-0.5 text-[11px] font-bold text-[#15803d]">
      言語化済
    </span>
  ) : (
    <span className="rounded-full bg-[#f1f5f9] px-2 py-0.5 text-[11px] font-bold text-[#64748b]">
      未
    </span>
  );
}

const btnPri =
  "rounded-[9px] bg-[var(--teal)] px-[18px] py-2.5 text-[13.5px] font-bold text-white hover:bg-[var(--teal-d)]";
const btnOut =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-[18px] py-2.5 text-[13.5px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
const btnOutSm =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-3 py-1.5 text-[12px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
const th =
  "border-b border-[var(--line)] bg-[#f1f5f9] px-3 py-2.5 text-left text-[12px] font-bold text-[#475569]";
const td = "border-b border-[#f1f5f9] bg-white px-3 py-2.5 align-middle";
