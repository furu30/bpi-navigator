"use client";

import { useStore } from "@/lib/store";

export default function SettingsPage() {
  const { state, hydrated, setCompany } = useStore();

  return (
    <>
      <div className="mb-5 flex gap-2.5 rounded-xl border border-[var(--blue-line)] bg-[var(--blue-soft)] px-4 py-3 text-[13px] text-[#1e3a8a]">
        💡
        <div>
          <b className="text-[var(--blue-d)]">会社名</b>
          はここで設定します。設定すると画面右上の表示や、保存するJSONにも反映されます。
        </div>
      </div>

      <div className="rounded-xl border border-[var(--line)] bg-white p-[18px] shadow-[0_1px_3px_rgba(0,0,0,.04)]">
        <h3 className="mb-1 text-[15px] font-extrabold">会社名（クライアント名）</h3>
        <p className="mb-3 text-[12.5px] text-[var(--muted)]">
          例：(株)KK精工
        </p>
        <input
          value={hydrated ? state.company : ""}
          onChange={(e) => setCompany(e.target.value)}
          placeholder="会社名を入力"
          disabled={!hydrated}
          className="w-full max-w-[420px] rounded-lg border-[1.5px] border-[var(--line)] px-[11px] py-[9px] text-[13.5px] focus:border-[var(--blue)] focus:outline-none"
        />
        <p className="mt-2 text-[12px] text-[var(--muted)]">
          入力は自動保存されます（このブラウザのlocalStorage）。
        </p>
      </div>

      <div className="mt-4 rounded-xl border border-dashed border-[var(--line)] bg-white p-5 text-[12.5px] text-[var(--muted)]">
        プロセス区分マスタ／問題カテゴリマスタの編集は、ステップ6で実装します。
      </div>
    </>
  );
}
