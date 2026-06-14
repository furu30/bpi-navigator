"use client";

// ブラウザから各AIプロバイダを直接呼び出すクライアント（BYOキー）。
// サーバープロキシは経由しない。キーはユーザーのブラウザ内のみで使用する。

import type { AiSettings } from "./aiSettings";

export interface AiMessage {
  system: string;
  user: string;
}

/** プロバイダを直接呼び出して本文テキストを返す。失敗時は例外を投げる。 */
export async function callAi(
  settings: AiSettings,
  msg: AiMessage,
): Promise<string> {
  if (!settings.apiKey.trim()) {
    throw new Error("APIキーが未設定です。設定・マスタの「AI設定」で入力してください。");
  }
  switch (settings.provider) {
    case "anthropic":
      return callAnthropic(settings, msg);
    case "openai":
      return callOpenAI(settings, msg);
    case "gemini":
      return callGemini(settings, msg);
    default:
      throw new Error("不明なプロバイダです。");
  }
}

async function callAnthropic(s: AiSettings, msg: AiMessage): Promise<string> {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": s.apiKey,
      "anthropic-version": "2023-06-01",
      // ブラウザからの直接アクセスを許可するヘッダ
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: s.model,
      max_tokens: 1500,
      system: msg.system,
      messages: [{ role: "user", content: msg.user }],
    }),
  });
  const data = await readJson(res);
  if (!res.ok) throw new Error(errMessage(res, data));
  const text = data?.content?.map((b: { text?: string }) => b.text ?? "").join("") ?? "";
  if (!text) throw new Error("応答が空でした。");
  return text;
}

async function callOpenAI(s: AiSettings, msg: AiMessage): Promise<string> {
  const res = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: `Bearer ${s.apiKey}`,
    },
    body: JSON.stringify({
      model: s.model,
      max_tokens: 1500,
      messages: [
        { role: "system", content: msg.system },
        { role: "user", content: msg.user },
      ],
    }),
  });
  const data = await readJson(res);
  if (!res.ok) throw new Error(errMessage(res, data));
  const text = data?.choices?.[0]?.message?.content ?? "";
  if (!text) throw new Error("応答が空でした。");
  return text;
}

async function callGemini(s: AiSettings, msg: AiMessage): Promise<string> {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(
    s.model,
  )}:generateContent?key=${encodeURIComponent(s.apiKey)}`;
  const res = await fetch(url, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      systemInstruction: { parts: [{ text: msg.system }] },
      contents: [{ role: "user", parts: [{ text: msg.user }] }],
    }),
  });
  const data = await readJson(res);
  if (!res.ok) throw new Error(errMessage(res, data));
  const text =
    data?.candidates?.[0]?.content?.parts
      ?.map((p: { text?: string }) => p.text ?? "")
      .join("") ?? "";
  if (!text) throw new Error("応答が空でした。");
  return text;
}

// 外部APIのレスポンス形状はプロバイダ依存のため緩く扱う
// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function readJson(res: Response): Promise<any> {
  try {
    return await res.json();
  } catch {
    return {};
  }
}

function errMessage(res: Response, data: { error?: { message?: string } }): string {
  const detail = data?.error?.message;
  if (res.status === 401 || res.status === 403) {
    return `認証に失敗しました（${res.status}）。APIキーをご確認ください。`;
  }
  if (res.status === 429) {
    return "レート上限に達しました（429）。少し時間をおいて再度お試しください。";
  }
  return `AI呼び出しに失敗しました（${res.status}）。${detail ?? ""}`.trim();
}
