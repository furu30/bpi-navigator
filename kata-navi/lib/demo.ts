// デモデータ（(株)KK精工）— 試作HTMLの初期データを §4 スキーマへ移植

import type { AppState, Masters, Task } from "./types";

/** 短いランダムID（試作HTML uid() 相当） */
export function uid(): string {
  return Math.random().toString(36).slice(2, 8);
}

export const DEFAULT_PROCESS_CATEGORIES = [
  "受注",
  "生産管理",
  "製造",
  "品質管理",
  "出荷",
  "購買",
];

export const DEFAULT_PROBLEM_CATEGORIES = [
  "手作業が多い",
  "二重入力",
  "待ち時間",
  "ミスが多い",
  "属人化",
  "紙ベース",
];

export function defaultMasters(): Masters {
  return {
    processCategories: [...DEFAULT_PROCESS_CATEGORIES],
    problemCategories: DEFAULT_PROBLEM_CATEGORIES.map((name) => ({
      id: uid(),
      name,
    })),
  };
}

type DemoSeed = Omit<Task, "id">;

const SEED: DemoSeed[] = [
  {
    group: "受注業務",
    content: "受注データの入力",
    person: "事務",
    time: 40,
    freq: "日次",
    workType: "作業",
    category: "受注",
    mmm: { muri: false, muda: true, mura: false },
    problems: ["二重入力", "手作業が多い"],
    scores: { impact: 5, freq: 5 },
  },
  {
    group: "受注業務",
    content: "納期回答の作成",
    person: "営業",
    time: 25,
    freq: "日次",
    workType: "混在",
    category: "受注",
    mmm: { muri: false, muda: false, mura: true },
    problems: ["属人化"],
    scores: { impact: 4, freq: 5 },
  },
  {
    group: "製造業務",
    content: "日報の転記",
    person: "製造",
    time: 30,
    freq: "日次",
    workType: "作業",
    category: "製造",
    mmm: { muri: false, muda: true, mura: false },
    problems: ["紙ベース", "二重入力"],
    scores: { impact: 4, freq: 5 },
  },
  {
    group: "品質業務",
    content: "検査記録の清書",
    person: "品質",
    time: 35,
    freq: "日次",
    workType: "作業",
    category: "品質管理",
    mmm: { muri: false, muda: true, mura: false },
    problems: ["紙ベース"],
    scores: { impact: 3, freq: 5 },
  },
  {
    group: "生産管理業務",
    content: "在庫の目視確認",
    person: "製造",
    time: 20,
    freq: "日次",
    workType: "感覚",
    category: "生産管理",
    mmm: { muri: false, muda: false, mura: true },
    problems: ["属人化"],
    scores: { impact: 3, freq: 5 },
  },
  {
    group: "出荷業務",
    content: "出荷伝票の照合",
    person: "出荷",
    time: 15,
    freq: "日次",
    workType: "作業",
    category: "出荷",
    mmm: { muri: true, muda: false, mura: false },
    problems: ["ミスが多い"],
    scores: { impact: 3, freq: 5 },
  },
];

/** デモ初期状態を生成 */
export function demoState(): AppState {
  return {
    company: "(株)KK精工",
    tasks: SEED.map((s) => ({ id: uid(), ...s })),
    plans: [],
    reproPlans: [],
    masters: defaultMasters(),
    currentGroup: null,
  };
}
