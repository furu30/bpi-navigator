import Link from "next/link";

export default function HomePage() {
  return (
    <div className="px-5 py-10 text-center">
      <div className="inline-block rounded-[10px] bg-[#2563eb] px-3.5 py-[3px] text-[30px] font-extrabold text-white">
        型
      </div>
      <h2 className="my-4 text-[26px] font-extrabold">業務改善ナビ</h2>
      <p className="mb-2 text-[var(--muted)]">業務のムダを見つけて、減らす。</p>
      <p className="mb-2 text-[var(--muted)]">
        作業を ECRS（排除・結合・交換・簡素化）で見直し、効率化します。
      </p>

      <div className="mx-auto mt-[26px] grid max-w-[620px] grid-cols-1 gap-4 text-left sm:grid-cols-2">
        <Link
          href="/stock"
          className="block rounded-[14px] border-2 border-[var(--line)] bg-white p-[22px] hover:border-[var(--blue)] hover:shadow-[0_6px_18px_rgba(37,99,235,.12)]"
        >
          <h4 className="mb-1.5 text-[16px] font-extrabold text-[var(--blue-d)]">
            ① まず業務を棚卸し
          </h4>
          <div className="text-[12.5px] text-[var(--muted)]">
            全業務を洗い出し、作業/感覚に切り分けます（共通の土台）。
          </div>
        </Link>

        <div className="block cursor-not-allowed rounded-[14px] border-2 border-[var(--line)] bg-white p-[22px] opacity-50">
          <h4 className="mb-1.5 text-[16px] font-extrabold text-[#0f766e]">
            再現性ナビ（アプリB）
          </h4>
          <div className="text-[12.5px] text-[var(--muted)]">
            属人性の再現は別アプリ。※今回のスコープ外（将来）。
          </div>
        </div>
      </div>

      <div className="mx-auto mt-4 max-w-[620px] rounded-lg border border-[#fde68a] bg-[#fffbeb] px-3 py-2 text-[11.5px] text-[var(--amber)]">
        データはこのブラウザ内（localStorage）に自動保存されます。サーバーには送信しません。
      </div>
    </div>
  );
}
