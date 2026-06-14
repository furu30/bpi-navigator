"use client";

import { useEffect, useRef, useState } from "react";
import {
  type AiSettings,
  loadAiSettings,
  PROVIDERS,
  providerMeta,
  saveAiSettings,
} from "@/lib/aiSettings";
import { exportAll, exportGroup, parseImport } from "@/lib/io";
import { groupsOf, useStore } from "@/lib/store";

// 設定・マスタ管理
// - 会社名
// - プロセス区分マスタ／問題カテゴリマスタの追加・編集・削除
// - データの保存／読込（全体・業務名単位）の導線を集約
export default function SettingsPage() {
  const { hydrated } = useStore();

  if (!hydrated) return <div className="text-[var(--muted)]">読み込み中…</div>;

  return (
    <div className="flex flex-col gap-4">
      <div className="flex gap-2.5 rounded-xl border border-[var(--blue-line)] bg-[var(--blue-soft)] px-4 py-3 text-[13px] text-[#1e3a8a]">
        ⚙️
        <div>
          <b className="text-[var(--blue-d)]">設定・マスタ管理</b>
          ：会社名・各マスタの編集と、データの保存／読込をここにまとめています。マスタの名称変更・削除は既存の作業データにも反映されます。
        </div>
      </div>

      <CompanyCard />
      <AiSettingsCard />
      <ProcessCategoryCard />
      <ProblemCategoryCard />
      <DataIoCard />
    </div>
  );
}

function Card({
  title,
  desc,
  children,
}: {
  title: string;
  desc?: string;
  children: React.ReactNode;
}) {
  return (
    <div className="rounded-xl border border-[var(--line)] bg-white p-[18px] shadow-[0_1px_3px_rgba(0,0,0,.04)]">
      <h3 className="text-[15px] font-extrabold">{title}</h3>
      {desc && <p className="mb-3 mt-0.5 text-[12.5px] text-[var(--muted)]">{desc}</p>}
      {children}
    </div>
  );
}

// ===== 会社名 =====
function CompanyCard() {
  const { state, setCompany } = useStore();
  return (
    <Card title="会社名（クライアント名）" desc="例：(株)KK精工。右上の表示や保存JSONに反映されます。">
      <input
        value={state.company}
        onChange={(e) => setCompany(e.target.value)}
        placeholder="会社名を入力"
        className={`${inputCls} max-w-[420px]`}
      />
      <p className="mt-2 text-[12px] text-[var(--muted)]">
        入力は自動保存されます（このブラウザのlocalStorage）。
      </p>
    </Card>
  );
}

// ===== AI設定（BYOキー） =====
function AiSettingsCard() {
  const [s, setS] = useState<AiSettings | null>(null);
  const [show, setShow] = useState(false);

  useEffect(() => {
    // eslint-disable-next-line react-hooks/set-state-in-effect -- マウント後にlocalStorageから一度だけ復元
    setS(loadAiSettings());
  }, []);

  if (!s) return null;

  function update(patch: Partial<AiSettings>) {
    setS((prev) => {
      const base = prev ?? loadAiSettings();
      let next = { ...base, ...patch };
      // プロバイダー変更時はモデルを既定値に切り替える
      if (patch.provider && patch.provider !== base.provider) {
        next = { ...next, model: providerMeta(patch.provider).defaultModel };
      }
      saveAiSettings(next);
      return next;
    });
  }

  const meta = providerMeta(s.provider);

  return (
    <Card
      title="AI設定（あなたのAPIキーを使用）"
      desc="AIアシストは、ここで登録したあなたのキーで、ブラウザから各プロバイダへ直接アクセスします（サーバーは経由しません）。"
    >
      <div className="mb-3 rounded-lg border border-[#fde68a] bg-[#fffbeb] px-3 py-2 text-[11.5px] text-[var(--amber)]">
        🔒 APIキーはお使いのブラウザ（localStorage）にのみ保存され、外部サーバーには送信されません。AIアシストは任意で、「AIに送る」を押したときだけ送信されます。
      </div>

      <div className="grid gap-3 sm:grid-cols-2">
        <div>
          <label className={labelCls}>AIプロバイダー</label>
          <select
            value={s.provider}
            onChange={(e) =>
              update({ provider: e.target.value as AiSettings["provider"] })
            }
            className={inputCls}
          >
            {PROVIDERS.map((p) => (
              <option key={p.id} value={p.id}>
                {p.label}
              </option>
            ))}
          </select>
        </div>
        <div>
          <label className={labelCls}>モデルID</label>
          <input
            value={s.model}
            onChange={(e) => update({ model: e.target.value })}
            placeholder={meta.defaultModel}
            className={inputCls}
          />
        </div>
      </div>

      <div className="mt-3">
        <label className={labelCls}>APIキー</label>
        <div className="flex gap-2">
          <input
            type={show ? "text" : "password"}
            value={s.apiKey}
            onChange={(e) => update({ apiKey: e.target.value })}
            placeholder={meta.keyHint}
            autoComplete="off"
            className={inputCls}
          />
          <button
            type="button"
            onClick={() => setShow((v) => !v)}
            className={btnOut}
          >
            {show ? "隠す" : "表示"}
          </button>
        </div>
        <p className="mt-1.5 text-[12px] text-[var(--muted)]">
          キーの取得：
          <a
            href={meta.keyUrl}
            target="_blank"
            rel="noopener noreferrer"
            className="font-bold text-[var(--blue)]"
          >
            {meta.label} のキー発行ページ
          </a>
          （{meta.keyHint}）。{s.apiKey.trim() ? "✅ キー登録済み" : "未登録"}
        </p>
      </div>
    </Card>
  );
}

// ===== プロセス区分マスタ =====
function ProcessCategoryCard() {
  const {
    state,
    addProcessCategory,
    renameProcessCategory,
    deleteProcessCategory,
  } = useStore();
  const [name, setName] = useState("");
  const cats = state.masters.processCategories;

  function usedCount(c: string) {
    return state.tasks.filter((t) => t.category === c).length;
  }

  return (
    <Card title="プロセス区分マスタ" desc="業務追加時の「プロセス区分」の選択肢です。">
      <ul className="mb-3 flex flex-col gap-1.5">
        {cats.map((c) => {
          const used = usedCount(c);
          return (
            <li
              key={c}
              className="flex items-center justify-between rounded-lg border border-[var(--line)] px-3 py-2 text-[13px]"
            >
              <span>
                {c}
                {used > 0 && (
                  <span className="ml-2 text-[11.5px] text-[var(--muted)]">
                    （{used}件で使用中）
                  </span>
                )}
              </span>
              <span className="flex gap-1">
                <button
                  type="button"
                  className={btnOutSm}
                  onClick={() => {
                    const next = window.prompt("プロセス区分の名称を変更", c);
                    if (next && next.trim()) renameProcessCategory(c, next.trim());
                  }}
                >
                  編集
                </button>
                <button
                  type="button"
                  className={btnOutSm}
                  onClick={() => {
                    const msg =
                      used > 0
                        ? `「${c}」を削除します（${used}件の作業のプロセス区分はそのまま残ります）。よろしいですか？`
                        : `「${c}」を削除しますか？`;
                    if (window.confirm(msg)) deleteProcessCategory(c);
                  }}
                >
                  削除
                </button>
              </span>
            </li>
          );
        })}
        {cats.length === 0 && (
          <li className="text-[12.5px] text-[var(--muted)]">
            まだ登録がありません。
          </li>
        )}
      </ul>
      <div className="flex gap-2">
        <input
          value={name}
          onChange={(e) => setName(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              addProcessCategory(name);
              setName("");
            }
          }}
          placeholder="例：購買"
          className={`${inputCls} max-w-[260px]`}
        />
        <button
          type="button"
          className={btnPri}
          onClick={() => {
            addProcessCategory(name);
            setName("");
          }}
        >
          ＋ 追加
        </button>
      </div>
    </Card>
  );
}

// ===== 問題カテゴリマスタ =====
function ProblemCategoryCard() {
  const {
    state,
    addProblemCategory,
    updateProblemCategory,
    deleteProblemCategory,
  } = useStore();
  const [name, setName] = useState("");
  const [desc, setDesc] = useState("");
  const cats = state.masters.problemCategories;

  return (
    <Card
      title="問題カテゴリマスタ"
      desc="A-2「問題の見える化」で選べる問題カテゴリです。名称変更・削除は既存作業のタグにも反映されます。"
    >
      <ul className="mb-3 flex flex-col gap-1.5">
        {cats.map((p) => (
          <li
            key={p.id}
            className="flex items-center justify-between rounded-lg border border-[var(--line)] px-3 py-2 text-[13px]"
          >
            <span>
              <b>{p.name}</b>
              {p.desc && (
                <span className="ml-2 text-[11.5px] text-[var(--muted)]">
                  {p.desc}
                </span>
              )}
            </span>
            <span className="flex gap-1">
              <button
                type="button"
                className={btnOutSm}
                onClick={() => {
                  const nextName = window.prompt("問題カテゴリの名称", p.name);
                  if (!nextName || !nextName.trim()) return;
                  const nextDesc = window.prompt("説明（任意・空欄可）", p.desc ?? "");
                  updateProblemCategory(
                    p.id,
                    nextName.trim(),
                    nextDesc?.trim() || undefined,
                  );
                }}
              >
                編集
              </button>
              <button
                type="button"
                className={btnOutSm}
                onClick={() => {
                  if (window.confirm(`「${p.name}」を削除しますか？（既存作業のタグからも外れます）`))
                    deleteProblemCategory(p.id);
                }}
              >
                削除
              </button>
            </span>
          </li>
        ))}
        {cats.length === 0 && (
          <li className="text-[12.5px] text-[var(--muted)]">
            まだ登録がありません。
          </li>
        )}
      </ul>
      <div className="flex flex-wrap gap-2">
        <input
          value={name}
          onChange={(e) => setName(e.target.value)}
          placeholder="例：手戻りが多い"
          className={`${inputCls} max-w-[220px]`}
        />
        <input
          value={desc}
          onChange={(e) => setDesc(e.target.value)}
          placeholder="説明（任意）"
          className={`${inputCls} max-w-[220px]`}
        />
        <button
          type="button"
          className={btnPri}
          onClick={() => {
            addProblemCategory(name, desc);
            setName("");
            setDesc("");
          }}
        >
          ＋ 追加
        </button>
      </div>
    </Card>
  );
}

// ===== データの保存／読込 =====
function DataIoCard() {
  const { state, replaceState, mergeGroupFile, resetDemo } = useStore();
  const groups = groupsOf(state);
  const [group, setGroup] = useState(state.currentGroup ?? groups[0] ?? "");
  const fileRef = useRef<HTMLInputElement>(null);

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
    <Card
      title="データの保存／読込"
      desc="全体（全業務名）または業務名単位でJSON保存できます。読込はどちらの形式も自動判定します。"
    >
      <div className="flex flex-col gap-3">
        <div className="flex flex-wrap items-center gap-2">
          <span className="w-[120px] text-[12.5px] font-bold text-[#475569]">
            全体（全業務名）
          </span>
          <button type="button" className={btnPri} onClick={() => exportAll(state)}>
            💾 全体を保存
          </button>
          <button
            type="button"
            className={btnOut}
            onClick={() => fileRef.current?.click()}
          >
            📂 読込（全体／業務名）
          </button>
          <input
            ref={fileRef}
            type="file"
            accept=".json"
            className="hidden"
            onChange={(e) => {
              handleImport(e.target.files?.[0]);
              e.target.value = "";
            }}
          />
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <span className="w-[120px] text-[12.5px] font-bold text-[#475569]">
            業務名単位
          </span>
          <select
            value={group}
            onChange={(e) => setGroup(e.target.value)}
            disabled={groups.length === 0}
            className={`${inputCls} max-w-[220px]`}
          >
            {groups.length === 0 ? (
              <option value="">（業務名がありません）</option>
            ) : (
              groups.map((g) => <option key={g}>{g}</option>)
            )}
          </select>
          <button
            type="button"
            className={btnOut}
            disabled={!group}
            onClick={() => group && exportGroup(state, group)}
          >
            💾 この業務名を保存
          </button>
        </div>

        <div className="flex flex-wrap items-center gap-2 border-t border-[var(--line)] pt-3">
          <span className="w-[120px] text-[12.5px] font-bold text-[#475569]">
            初期化
          </span>
          <button
            type="button"
            className={btnOut}
            onClick={() => {
              if (window.confirm("デモデータに戻します。よろしいですか？"))
                resetDemo();
            }}
          >
            ↺ デモ初期化
          </button>
        </div>
      </div>
    </Card>
  );
}

const labelCls = "mb-1 block text-[12.5px] font-bold text-[#475569]";
const inputCls =
  "w-full rounded-lg border-[1.5px] border-[var(--line)] px-[11px] py-[9px] text-[13.5px] focus:border-[var(--blue)] focus:outline-none disabled:opacity-50";
const btnPri =
  "rounded-[9px] bg-[var(--blue)] px-[14px] py-2 text-[13px] font-bold text-white hover:bg-[var(--blue-d)]";
const btnOut =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-[14px] py-2 text-[13px] font-bold text-[var(--navy)] hover:border-[#94a3b8] disabled:opacity-50";
const btnOutSm =
  "rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-3 py-1 text-[12px] font-bold text-[var(--navy)] hover:border-[#94a3b8]";
