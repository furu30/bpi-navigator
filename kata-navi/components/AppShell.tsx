"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { useRef, useState } from "react";
import { exportAll, parseImport } from "@/lib/io";
import { useStore } from "@/lib/store";
import { GroupSwitcher } from "./GroupSwitcher";

interface NavItem {
  href: string;
  label: string;
  title: string; // トップバー見出し
  icon?: string;
  num?: number;
}

const START_NAV: NavItem[] = [
  { href: "/", label: "ホーム", title: "ホーム", icon: "🏠" },
  {
    href: "/stock",
    label: "業務の棚卸し",
    title: "業務の棚卸し（共通）",
    icon: "🗂",
  },
];

const FLOW_NAV: NavItem[] = [
  { href: "/a/dashboard", label: "改善ダッシュボード", title: "A-1 改善ダッシュボード", num: 1 },
  { href: "/a/analyze", label: "問題の見える化", title: "A-2 問題の見える化", num: 2 },
  { href: "/a/ecrs", label: "ECRSで改善案", title: "A-3 ECRSで改善案", num: 3 },
  { href: "/a/plan", label: "改善計画", title: "A-4 改善計画", num: 4 },
  { href: "/a/result", label: "効果検証・レポート", title: "A-5 効果検証・レポート", num: 5 },
];

const MANAGE_NAV: NavItem[] = [
  { href: "/settings", label: "設定・マスタ", title: "設定・マスタ", icon: "⚙" },
];

const ALL_NAV = [...START_NAV, ...FLOW_NAV, ...MANAGE_NAV];

function titleFor(pathname: string): string {
  const hit = ALL_NAV.find((n) =>
    n.href === "/" ? pathname === "/" : pathname.startsWith(n.href),
  );
  if (hit) return hit.title;
  if (pathname.startsWith("/settings")) return "設定・マスタ";
  return "KATA Navi";
}

function isActive(pathname: string, href: string): boolean {
  return href === "/" ? pathname === "/" : pathname.startsWith(href);
}

export function AppShell({ children }: { children: React.ReactNode }) {
  const pathname = usePathname();
  const { state, resetDemo, replaceState, mergeGroupFile } = useStore();
  const isHome = pathname === "/";
  const fileInputRef = useRef<HTMLInputElement>(null);
  // モバイル用サイドバー（ドロワー）の開閉
  const [navOpen, setNavOpen] = useState(false);
  const closeNav = () => setNavOpen(false);

  function handleImport(file: File | undefined) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const result = parseImport(String(reader.result));
      if (result.kind === "error") {
        window.alert(result.message);
        return;
      }
      if (result.kind === "group") mergeGroupFile(result.data);
      else replaceState(result.data);
      window.alert("読み込みました。");
    };
    reader.readAsText(file);
  }

  return (
    <div className="flex min-h-screen">
      {/* モバイル時のオーバーレイ（ドロワー背面） */}
      {navOpen && (
        <button
          type="button"
          aria-label="メニューを閉じる"
          onClick={closeNav}
          className="fixed inset-0 z-30 bg-black/40 md:hidden"
        />
      )}

      {/* ===== サイドバー（md未満はドロワー、md以上は常設） ===== */}
      <nav
        className={`fixed inset-y-0 left-0 z-40 flex h-screen w-[230px] flex-shrink-0 flex-col bg-[var(--navy)] text-[#cbd5e1] transition-transform duration-200 md:sticky md:top-0 md:z-auto md:translate-x-0 ${
          navOpen ? "translate-x-0" : "-translate-x-full"
        }`}
      >
        <div className="flex items-start justify-between border-b border-[#334155] px-[18px] pb-[14px] pt-[18px]">
          <div>
            <div className="text-[18px] font-extrabold text-white">
              <span className="mr-1 inline-block rounded-md bg-[#2563eb] px-[7px] py-px">
                型
              </span>
              KATA Navi
            </div>
            <div className="mt-1 text-[11.5px] font-bold text-[#7dd3fc]">
              アプリA：業務改善ナビ
            </div>
          </div>
          {/* モバイルの閉じるボタン */}
          <button
            type="button"
            aria-label="メニューを閉じる"
            onClick={closeNav}
            className="-mr-1 text-[22px] leading-none text-[#cbd5e1] hover:text-white md:hidden"
          >
            ×
          </button>
        </div>

        <NavSection label="スタート" items={START_NAV} pathname={pathname} onNavigate={closeNav} />
        <NavSection label="改善の流れ" items={FLOW_NAV} pathname={pathname} onNavigate={closeNav} />
        <NavSection label="管理" items={MANAGE_NAV} pathname={pathname} onNavigate={closeNav} />

        <div className="mt-auto flex flex-wrap gap-2 border-t border-[#334155] px-4 py-[14px]">
          <FootButton
            label="💾 保存"
            title="全体（全業務名）をJSONで保存"
            onClick={() => exportAll(state)}
          />
          <FootButton
            label="📂 読込"
            title="JSONを読み込み（全体／業務名ファイル）"
            onClick={() => fileInputRef.current?.click()}
          />
          <FootButton
            label="↺ デモ初期化"
            onClick={() => {
              if (window.confirm("デモデータに戻します。よろしいですか？")) {
                resetDemo();
              }
            }}
          />
          <input
            ref={fileInputRef}
            type="file"
            accept=".json"
            className="hidden"
            onChange={(e) => {
              handleImport(e.target.files?.[0]);
              e.target.value = "";
            }}
          />
        </div>
      </nav>

      {/* ===== メイン ===== */}
      <div className="flex min-w-0 flex-1 flex-col">
        <div className="sticky top-0 z-[5] flex flex-wrap items-center justify-between gap-2 border-b border-[var(--line)] bg-white px-4 py-3 md:px-[30px] md:py-[14px]">
          <div className="flex min-w-0 items-center gap-2">
            <button
              type="button"
              aria-label="メニューを開く"
              onClick={() => setNavOpen(true)}
              className="-ml-1 rounded-md px-2 py-1 text-[20px] leading-none text-[var(--navy)] hover:bg-[var(--bg)] md:hidden"
            >
              ☰
            </button>
            <h1 className="truncate text-[16px] font-extrabold md:text-[17px]">
              {titleFor(pathname)}
            </h1>
          </div>
          <div className="flex min-w-0 items-center gap-2 md:gap-3.5">
            {!isHome && <GroupSwitcher />}
            <Link
              href="/settings"
              title="設定・マスタで会社名を変更"
              className="max-w-[160px] truncate whitespace-nowrap rounded-full bg-[var(--blue-soft)] px-[11px] py-1 text-[11.5px] text-[var(--muted)] hover:bg-[var(--blue-line)]"
            >
              🏢 {state.company?.trim() ? state.company : "会社名 未設定"}
            </Link>
          </div>
        </div>
        <div className="mx-auto w-full max-w-[980px] px-4 py-5 md:px-[30px] md:py-[26px]">
          {children}
        </div>
      </div>
    </div>
  );
}

function NavSection({
  label,
  items,
  pathname,
  onNavigate,
}: {
  label: string;
  items: NavItem[];
  pathname: string;
  onNavigate?: () => void;
}) {
  return (
    <>
      <div className="px-[18px] pb-[5px] pt-[14px] text-[10.5px] tracking-[0.1em] text-[#64748b]">
        {label}
      </div>
      <div className="flex flex-col gap-px px-2">
        {items.map((item) => {
          const active = isActive(pathname, item.href);
          return (
            <Link
              key={item.href}
              href={item.href}
              onClick={onNavigate}
              className={`flex w-full items-center gap-[9px] rounded-lg px-3 py-2.5 text-[13.5px] ${
                active
                  ? "bg-[#2563eb] font-bold text-white"
                  : "text-[#cbd5e1] hover:bg-[#334155] hover:text-white"
              }`}
            >
              {item.num != null ? (
                <span
                  className={`flex h-5 w-5 flex-shrink-0 items-center justify-center rounded-md text-[11px] ${
                    active ? "bg-white text-[#2563eb]" : "bg-[#475569] text-white"
                  }`}
                >
                  {item.num}
                </span>
              ) : (
                <span>{item.icon}</span>
              )}
              {item.label}
            </Link>
          );
        })}
      </div>
    </>
  );
}

function FootButton({
  label,
  onClick,
  disabled,
  title,
}: {
  label: string;
  onClick?: () => void;
  disabled?: boolean;
  title?: string;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      title={title}
      className="flex-1 rounded-md bg-[#334155] px-2 py-2 text-[11.5px] text-[#cbd5e1] enabled:cursor-pointer enabled:hover:bg-[#475569] enabled:hover:text-white disabled:opacity-40"
    >
      {label}
    </button>
  );
}
