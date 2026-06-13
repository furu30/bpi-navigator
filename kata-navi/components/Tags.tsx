import type { WorkType } from "@/lib/types";

const TYPE_STYLE: Record<WorkType, string> = {
  作業: "bg-[#dbeafe] text-[#1d4ed8]",
  感覚: "bg-[#ccfbf1] text-[#0f766e]",
  混在: "bg-[#fef3c7] text-[#92400e]",
  未判定: "bg-[#f1f5f9] text-[#64748b]",
};

/** 業務種別タグ */
export function TypeTag({ type }: { type: WorkType }) {
  return (
    <span
      className={`inline-block rounded-full px-[9px] py-0.5 text-[11px] font-bold ${TYPE_STYLE[type]}`}
    >
      {type}
    </span>
  );
}

/** ムリ・ムダ・ムラのタグ群 */
export function MmmTags({
  mmm,
}: {
  mmm: { muri: boolean; muda: boolean; mura: boolean };
}) {
  const tags: { on: boolean; label: string; cls: string }[] = [
    { on: mmm.muri, label: "ムリ", cls: "bg-[#fee2e2] text-[#b91c1c]" },
    { on: mmm.muda, label: "ムダ", cls: "bg-[#ffedd5] text-[#c2410c]" },
    { on: mmm.mura, label: "ムラ", cls: "bg-[#fef9c3] text-[#a16207]" },
  ];
  const active = tags.filter((t) => t.on);
  if (active.length === 0) {
    return <span className="text-[var(--muted)]">-</span>;
  }
  return (
    <>
      {active.map((t) => (
        <span
          key={t.label}
          className={`mr-[3px] inline-block rounded-[5px] px-1.5 py-px text-[10.5px] font-bold ${t.cls}`}
        >
          {t.label}
        </span>
      ))}
    </>
  );
}
