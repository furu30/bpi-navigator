/**
 * AI Assist Module for BPI Navigator
 * Claude / OpenAI / Gemini の統一インターフェース
 */
const AIAssist = (() => {
  'use strict';

  // ========== 定数 ==========
  const STORAGE_KEY = 'bpiNavi_aiConfig';

  const PROVIDERS = {
    claude: {
      name: 'Claude (Anthropic)',
      models: [
        { id: 'claude-sonnet-4-6', name: 'Claude Sonnet 4.6' },
        { id: 'claude-opus-4-6', name: 'Claude Opus 4.6' },
        { id: 'claude-haiku-4-5', name: 'Claude Haiku 4.5' }
      ],
      defaultModel: 'claude-sonnet-4-6'
    },
    openai: {
      name: 'OpenAI',
      models: [
        { id: 'gpt-5-mini', name: 'GPT-5 mini' },
        { id: 'gpt-5', name: 'GPT-5' },
        { id: 'gpt-4.1', name: 'GPT-4.1' },
        { id: 'gpt-4o-mini', name: 'GPT-4o mini' }
      ],
      defaultModel: 'gpt-5-mini'
    },
    gemini: {
      name: 'Gemini (Google)',
      models: [
        { id: 'gemini-2.5-flash', name: 'Gemini 2.5 Flash' },
        { id: 'gemini-2.5-flash-lite', name: 'Gemini 2.5 Flash Lite' },
        { id: 'gemini-2.5-pro', name: 'Gemini 2.5 Pro' }
      ],
      defaultModel: 'gemini-2.5-flash'
    }
  };

  // ========== 設定の読み書き ==========
  function loadConfig() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) return JSON.parse(raw);
    } catch (e) { /* ignore */ }
    return { provider: 'claude', apiKey: '', model: '' };
  }

  function saveConfig(cfg) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
    } catch (e) { /* ignore */ }
  }

  function getEffectiveModel(cfg) {
    if (cfg.model) return cfg.model;
    const prov = PROVIDERS[cfg.provider];
    return prov ? prov.defaultModel : '';
  }

  // ========== API 呼び出し ==========
  async function callAI(systemPrompt, userMessage, extraInstruction) {
    const cfg = loadConfig();
    if (!cfg.apiKey) throw new Error('APIキーが設定されていません。設定画面で入力してください。');

    const model = getEffectiveModel(cfg);
    const finalUser = extraInstruction
      ? `${userMessage}\n\n【追加指示】${extraInstruction}`
      : userMessage;

    switch (cfg.provider) {
      case 'claude': return callClaude(cfg.apiKey, model, systemPrompt, finalUser);
      case 'openai': return callOpenAI(cfg.apiKey, model, systemPrompt, finalUser);
      case 'gemini': return callGemini(cfg.apiKey, model, systemPrompt, finalUser);
      default: throw new Error('不明なプロバイダー: ' + cfg.provider);
    }
  }

  async function callClaude(apiKey, model, system, user) {
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true'
      },
      body: JSON.stringify({
        model,
        max_tokens: 4096,
        system,
        messages: [{ role: 'user', content: user }]
      })
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(`Claude API エラー (${res.status}): ${err.error?.message || res.statusText}`);
    }
    const data = await res.json();
    return data.content?.[0]?.text || '';
  }

  async function callOpenAI(apiKey, model, system, user) {
    const res = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model,
        messages: [
          { role: 'system', content: system },
          { role: 'user', content: user }
        ],
        temperature: 0.7,
        max_tokens: 4096
      })
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(`OpenAI API エラー (${res.status}): ${err.error?.message || res.statusText}`);
    }
    const data = await res.json();
    return data.choices?.[0]?.message?.content || '';
  }

  async function callGemini(apiKey, model, system, user) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        system_instruction: { parts: [{ text: system }] },
        contents: [{ parts: [{ text: user }] }],
        generationConfig: { temperature: 0.7, maxOutputTokens: 4096 }
      })
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(`Gemini API エラー (${res.status}): ${err.error?.message || res.statusText}`);
    }
    const data = await res.json();
    return data.candidates?.[0]?.content?.parts?.[0]?.text || '';
  }

  // ========== 接続テスト ==========
  async function testConnection() {
    return await callAI(
      'あなたはテスト用のアシスタントです。',
      '接続テストです。「接続成功」と返してください。',
      ''
    );
  }

  // ========== レスポンス解析 ==========
  function parseResponse(text) {
    // ```json ... ``` ブロックを探す
    const jsonBlock = text.match(/```(?:json)?\s*([\s\S]*?)```/);
    const target = jsonBlock ? jsonBlock[1].trim() : text.trim();
    try {
      return { ok: true, data: JSON.parse(target) };
    } catch (e) {
      // JSONの配列/オブジェクト部分を探す
      const arrMatch = target.match(/\[[\s\S]*\]/);
      const objMatch = target.match(/\{[\s\S]*\}/);
      const found = arrMatch ? arrMatch[0] : (objMatch ? objMatch[0] : null);
      if (found) {
        try {
          return { ok: true, data: JSON.parse(found) };
        } catch (e2) { /* fall through */ }
      }
      return { ok: false, data: null, raw: text };
    }
  }

  // ========== 共通システムプロンプト ==========
  const BASE_SYSTEM = `あなたは製造業の業務プロセス改善に精通したコンサルタントです。
日本語で回答してください。
データはJSON形式で出力してください（アプリに取り込むため）。
JSONは必ず \`\`\`json ... \`\`\` ブロックで囲んでください。`;

  // ========== Step別プロンプト生成 ==========

  function buildStep1Prompt(proj) {
    const categories = (proj.processes || []).join('、') || 'すべて';
    const existingCount = (proj.tasks || []).length;
    return {
      system: BASE_SYSTEM + `\n\n## 出力フォーマット
JSON配列で、各要素は以下のフィールドを持つオブジェクト:
- content: 作業内容（具体的に）
- person: 担当者（役職名）
- target: 対象（何を処理するか）
- method: 方法（どうやるか）
- timeRequired: 所要時間（数値、文字列）
- timeUnit: 単位（"分" or "時間"）
- freqType: 頻度タイプ（"日次" or "週次" or "月次"）
- freqCount: 頻度回数（数値、文字列）
- tools: 使用ツール
- notes: 備考・気になる点

15〜20件程度の業務を洗い出してください。`,

      user: `以下の企業の「${categories}」プロセスにおける業務を洗い出してください。

【企業名】${proj.company || '（未設定）'}
【企業概要】${proj.companyDescription || '（未設定）'}
【プロジェクト名】${proj.name || ''}
【プロジェクトメモ】${proj.memo || '（未設定）'}
${existingCount > 0 ? `\n【既に登録済みの業務数】${existingCount}件（重複しない新規業務を提案してください）` : ''}`
    };
  }

  function buildStep2Prompt(proj) {
    const tasks = (proj.tasks || []).map(t => ({
      code: t.code, content: t.content, person: t.person,
      method: t.method, tools: t.tools, notes: t.notes
    }));
    return {
      system: BASE_SYSTEM + `\n\n## 問題カテゴリ（以下から選択）
- duplicate: 転記ミス・二重入力
- waiting: 待ち時間・遅延
- nostandard: 標準化されていない
- paper: 紙・アナログ作業
- personal: 属人化
- search: 検索・照合に手間
- communication: コミュニケーションロス
- system: システム未連携

## スコア（各1〜5の整数）
- timeImpact: 時間への影響度
- qualityImpact: 品質への影響度
- frequency: 発生頻度
- difficulty: 改善の難易度

## 出力フォーマット
JSON配列で、各要素:
- code: タスクコード
- problems: 問題カテゴリIDの配列（1〜3個）
- scores: { timeImpact, qualityImpact, frequency, difficulty }
- reason: 判断根拠（短文）`,

      user: `以下の業務一覧について、各業務の問題点を分析し、問題カテゴリの分類とスコアリングを提案してください。

【企業名】${proj.company || ''}
【企業概要】${proj.companyDescription || '（未設定）'}
【業務一覧】
${JSON.stringify(tasks, null, 2)}`
    };
  }

  function buildStep3Prompt(proj) {
    const tasks = (proj.tasks || [])
      .filter(t => t.scores && t.problems?.length > 0)
      .map(t => ({
        code: t.code, content: t.content, person: t.person,
        method: t.method, problems: t.problems,
        scores: t.scores, totalScore: calcScore(t.scores)
      }))
      .sort((a, b) => b.totalScore - a.totalScore);

    return {
      system: BASE_SYSTEM + `\n\n## ECRS分析
E(Eliminate/排除): その業務をなくせないか
C(Combine/結合): 他の業務と統合できないか
R(Rearrange/交換): 順序や担当を変えられないか
S(Simplify/簡素化): もっと簡単にできないか

各項目の判定は "yes"（該当）, "maybe"（検討余地あり）, "no"（該当しない）

## 出力フォーマット
JSON配列で、各要素:
- code: タスクコード
- ecrs: { E: "yes"|"maybe"|"no", C: ..., R: ..., S: ... }
- reason: ECRS判定の根拠（各項目の簡潔な理由）`,

      user: `以下の業務について、ECRS分析を行い改善候補を特定してください。
優先度スコアの高い順に並んでいます。

【企業名】${proj.company || ''}
【企業概要】${proj.companyDescription || '（未設定）'}
【分析対象業務】
${JSON.stringify(tasks, null, 2)}`
    };
  }

  function buildStep4Prompt(proj) {
    const candidates = (proj.tasks || [])
      .filter(t => t.ecrs)
      .map(t => {
        const keys = [];
        if (t.ecrs) {
          ['E', 'C', 'R', 'S'].forEach(k => {
            if (t.ecrs[k] === 'yes' || t.ecrs[k] === 'maybe') keys.push(k);
          });
        }
        return { code: t.code, content: t.content, person: t.person, method: t.method, ecrsKeys: keys, problems: t.problems };
      })
      .filter(t => t.ecrsKeys.length > 0);

    return {
      system: BASE_SYSTEM + `\n\n## 出力フォーマット
JSON配列で、各要素:
- taskCode: 対象タスクコード
- ecrsKey: 適用するECRSキー（"E", "C", "R", "S"のいずれか）
- content: 改善策の内容
- method: 具体的な実施方法
- person: 推進担当者
- resources: 必要リソース・費用
- effectQuantity: 定量的な期待効果
- effectQuality: 定性的な期待効果

現実的で段階的に実施可能な改善策を提案してください。`,

      user: `以下のECRS分析結果に基づき、具体的な改善策を提案してください。

【企業名】${proj.company || ''}
【企業概要】${proj.companyDescription || '（未設定）'}
【ECRS分析結果（改善候補）】
${JSON.stringify(candidates, null, 2)}`
    };
  }

  function buildStep5Prompt(proj) {
    const improvements = (proj.improvements || []).map(imp => {
      const task = (proj.tasks || []).find(t => t.id === imp.taskId);
      return {
        taskCode: task?.code || '?',
        taskContent: task?.content || '?',
        ecrsKey: imp.ecrsKey,
        content: imp.content,
        method: imp.method,
        status: imp.status,
        effectQuantity: imp.effectQuantity,
        effectQuality: imp.effectQuality
      };
    });

    return {
      system: BASE_SYSTEM + `\n\n## 出力フォーマット
JSON配列で、各要素:
- taskCode: 対象タスクコード
- ecrsKey: ECRSキー
- suggestedKPI: 測定すべきKPI指標
- measurementMethod: 測定方法の提案
- expectedValue: 期待される目標値
- checkpoints: 確認すべきポイント（文字列）

改善の効果を客観的に測定できる指標を提案してください。`,

      user: `以下の改善計画について、効果測定の方法とKPIを提案してください。

【企業名】${proj.company || ''}
【企業概要】${proj.companyDescription || '（未設定）'}
【改善計画一覧】
${JSON.stringify(improvements, null, 2)}`
    };
  }

  // スコア計算ヘルパー
  function calcScore(scores) {
    if (!scores) return 0;
    return (scores.timeImpact || 0) * 0.3
      + (scores.qualityImpact || 0) * 0.25
      + (scores.frequency || 0) * 0.25
      + (scores.difficulty || 0) * 0.2;
  }

  // ========== 公開API ==========
  return {
    PROVIDERS,
    loadConfig,
    saveConfig,
    getEffectiveModel,
    callAI,
    testConnection,
    parseResponse,
    buildStep1Prompt,
    buildStep2Prompt,
    buildStep3Prompt,
    buildStep4Prompt,
    buildStep5Prompt
  };
})();
