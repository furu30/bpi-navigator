"use client";

import Link from "next/link";
import { useState } from "react";
import { callAi } from "@/lib/ai";
import { loadAiSettings, providerMeta } from "@/lib/aiSettings";

interface Props {
  title: string;
  /** システムプロンプト（役割指示） */
  system: string;
  /** AIに送る基礎データ（業務一覧の要約など） */
  context: string;
  onClose: () => void;
}

// AIアシスト（BYOキー）。保存済みのユーザーキーで、押したときだけプロバイダへ直接送信する。
export function AiAssistModal({ title, system, context, onClose }: Props) {
  const settings = loadAiSettings();
  const hasKey = !!settings.apiKey.trim();
  const meta = providerMeta(settings.provider);

  const [extra, setExtra] = useState("");
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState("");
  const [error, setError] = useState("");

  async function handleSend() {
    setError("");
    setResult("");
    setLoading(true);
    try {
      const user = extra.trim()
        ? `${context}\n\n# 追加の依頼\n${extra.trim()}`
        : context;
      const text = await callAi(settings, { system, user });
      setResult(text);
    } catch (e) {
      setError(e instanceof Error ? e.message : "送信に失敗しました。");
    } finally {
      setLoading(false);
    }
  }

  return (
    <div
      className="fixed inset-0 z-50 flex items-start justify-center overflow-auto bg-[rgba(15,23,42,.5)] px-4 py-10"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="w-full max-w-[620px] rounded-2xl bg-white shadow-[0_20px_50px_rgba(0,0,0,.3)]">
        <h3 className="flex items-center justify-between border-b border-[var(--line)] px-5 py-4 text-[16px] font-extrabold">
          🤖 {title}
          <button
            type="button"
            onClick={onClose}
            className="text-[22px] leading-none text-[#94a3b8] hover:text-[var(--muted)]"
            aria-label="閉じる"
          >
            ×
          </button>
        </h3>

        <div className="px-5 py-[18px]">
          {/* プライバシー明記 */}
          <div className="mb-3 rounded-lg border border-[#fde68a] bg-[#fffbeb] px-3 py-2 text-[11.5px] text-[var(--amber)]">
            🔒 APIキーはお使いのブラウザにのみ保存され、外部サーバーには送信されません。AIへの送信は任意です（「AIに送る」を押したときだけ、選択中のプロバイダへ直接送られます）。
          </div>

          {!hasKey ? (
            <div className="rounded-lg border border-[var(--line)] bg-[#f8fafc] px-4 py-4 text-[13px]">
              AIアシストを使うには、まず
              <Link href="/settings" className="font-bold text-[var(--blue)]">
                設定・マスタの「AI設定」
              </Link>
              でプロバイダーとAPIキーを登録してください。
            </div>
          ) : (
            <>
              <div className="mb-2 text-[12px] text-[var(--muted)]">
                送信先：<b>{meta.label}</b> ／ モデル：<b>{settings.model}</b>
              </div>
              <div className="mb-1 text-[12.5px] font-bold text-[#475569]">
                送信内容（プレビュー）
              </div>
              <pre className="mb-3 max-h-[180px] overflow-auto whitespace-pre-wrap rounded-lg border border-[var(--line)] bg-[#f8fafc] p-3 text-[12px] text-[var(--ink)]">
                {context}
              </pre>

              <div className="mb-1 text-[12.5px] font-bold text-[#475569]">
                追加の依頼（任意）
              </div>
              <textarea
                value={extra}
                onChange={(e) => setExtra(e.target.value)}
                rows={2}
                placeholder="例：コストを増やさない案に絞って"
                className="mb-3 w-full rounded-lg border-[1.5px] border-[var(--line)] px-[11px] py-[9px] text-[13.5px] focus:border-[var(--blue)] focus:outline-none"
              />

              {error && (
                <p className="mb-3 rounded-lg bg-[#fee2e2] px-3 py-2 text-[12.5px] font-bold text-[var(--red)]">
                  {error}
                </p>
              )}

              {result && (
                <div className="mb-2">
                  <div className="mb-1 flex items-center justify-between">
                    <span className="text-[12.5px] font-bold text-[#475569]">
                      AIの提案
                    </span>
                    <button
                      type="button"
                      onClick={() => navigator.clipboard?.writeText(result)}
                      className="rounded-md border border-[var(--line)] px-2 py-1 text-[11.5px] font-bold text-[var(--navy)] hover:border-[#94a3b8]"
                    >
                      コピー
                    </button>
                  </div>
                  <div className="max-h-[280px] overflow-auto whitespace-pre-wrap rounded-lg border border-[var(--blue-line)] bg-[var(--blue-soft)] p-3 text-[13px] text-[var(--ink)]">
                    {result}
                  </div>
                </div>
              )}
            </>
          )}
        </div>

        <div className="flex justify-end gap-2.5 border-t border-[var(--line)] px-5 py-3.5">
          <button
            type="button"
            onClick={onClose}
            className="rounded-[9px] border-[1.5px] border-[var(--line)] bg-white px-[18px] py-2.5 text-[13.5px] font-bold text-[var(--navy)] hover:border-[#94a3b8]"
          >
            閉じる
          </button>
          {hasKey && (
            <button
              type="button"
              onClick={handleSend}
              disabled={loading}
              className="rounded-[9px] bg-[var(--blue)] px-[18px] py-2.5 text-[13.5px] font-bold text-white hover:bg-[var(--blue-d)] disabled:opacity-50"
            >
              {loading ? "送信中…" : result ? "もう一度AIに送る" : "AIに送る"}
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
