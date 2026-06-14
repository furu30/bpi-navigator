"use client";

// AI設定（BYOキー）。アプリ業務データ(AppState)とは別のlocalStorageキーで保持し、
// JSON保存/読込（業務データのエクスポート）には含めない＝キーを外部に出さない。

export type AiProvider = "anthropic" | "openai" | "gemini";

export interface AiSettings {
  provider: AiProvider;
  apiKey: string;
  /** モデルID（プロバイダ既定値を初期表示。必要に応じてユーザーが変更） */
  model: string;
}

export const AI_LS_KEY = "kata-navi-ai-settings";

export interface ProviderMeta {
  id: AiProvider;
  label: string;
  defaultModel: string;
  /** APIキー取得ページ */
  keyUrl: string;
  keyHint: string;
}

export const PROVIDERS: ProviderMeta[] = [
  {
    id: "anthropic",
    label: "Claude（Anthropic）",
    defaultModel: "claude-sonnet-4-5",
    keyUrl: "https://console.anthropic.com/settings/keys",
    keyHint: "sk-ant-… で始まるキー",
  },
  {
    id: "openai",
    label: "OpenAI（ChatGPT）",
    defaultModel: "gpt-4o-mini",
    keyUrl: "https://platform.openai.com/api-keys",
    keyHint: "sk-… で始まるキー",
  },
  {
    id: "gemini",
    label: "Gemini（Google）",
    defaultModel: "gemini-2.0-flash",
    keyUrl: "https://aistudio.google.com/app/apikey",
    keyHint: "AIza… で始まるキー",
  },
];

export function providerMeta(id: AiProvider): ProviderMeta {
  return PROVIDERS.find((p) => p.id === id) ?? PROVIDERS[0];
}

export function defaultAiSettings(): AiSettings {
  return { provider: "anthropic", apiKey: "", model: "claude-sonnet-4-5" };
}

export function loadAiSettings(): AiSettings {
  try {
    const raw = localStorage.getItem(AI_LS_KEY);
    if (raw) {
      const d = JSON.parse(raw);
      if (d && typeof d.apiKey === "string" && typeof d.provider === "string") {
        return {
          provider: d.provider,
          apiKey: d.apiKey,
          model: typeof d.model === "string" && d.model ? d.model : providerMeta(d.provider).defaultModel,
        };
      }
    }
  } catch {
    /* 破損時は既定値 */
  }
  return defaultAiSettings();
}

export function saveAiSettings(s: AiSettings): void {
  try {
    localStorage.setItem(AI_LS_KEY, JSON.stringify(s));
  } catch {
    /* 容量超過などは無視 */
  }
}
