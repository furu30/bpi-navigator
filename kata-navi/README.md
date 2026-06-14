# KATA Navi｜業務改善 × 再現性ナビ

中小製造業向けの「型づくり」アプリ。**仕事を"型"にして、改善し、受け継ぐ。**
業務を「作業／感覚」に切り分け、共通の業務棚卸しを土台に2つのアプリを回す。

- **アプリA（業務改善ナビ）**：作業を ECRS（排除・結合・交換・簡素化）で効率化し、ムダを減らす（業務を型にする）。
- **アプリB（再現性ナビ）**：感覚（暗黙知）を「5つの問い」で言語化し、標準化／形式知化／教育で受け継ぐ（技を型にする）。

タスクは両アプリで共有し、同じ業務を「改善」と「再現性」の両面から扱える。

- **技術**：Next.js（App Router）+ TypeScript + Tailwind CSS v4
- **データ**：ブラウザの localStorage に保存（DB・ログインなし）。サーバーに業務データは保存しない。
- **AIアシスト（BYOキー）**：各ユーザーが自分のAPIキー（Claude / OpenAI / Gemini）を
  「設定・マスタ > AI設定」で登録。キーはそのブラウザの localStorage にのみ保存し、
  ブラウザから各プロバイダへ**直接**アクセスする（サーバープロキシ・サーバー環境変数は不要）。

## 開発

```bash
npm install
npm run dev      # http://localhost:3000
npm run build    # 本番ビルド（全ページ静的プリレンダリング）
npm run lint
```

`.env` / APIキー等のサーバー秘密情報は不要。設定しない。

## セキュリティ

- AIキーはサーバーに置かない（BYOキー）。ソース・環境変数・JSONエクスポートのいずれにも含まれない。
  AI設定は業務データとは別の localStorage キー（`kata-navi-ai-settings`）で保持。
- 画面の入力はすべて React 経由で描画（自動エスケープ）。`dangerouslySetInnerHTML` 不使用。
- `next.config.ts` でセキュリティヘッダ（CSP / HSTS / X-Frame-Options / X-Content-Type-Options /
  Referrer-Policy / Permissions-Policy）を全ルートに付与。
  CSP の `connect-src` で各AIプロバイダのエンドポイントのみ許可。

## Vercel へのデプロイ

1. このリポジトリ（`kata-navi/`）を Vercel にインポート（Framework: Next.js 自動検出）。
2. 環境変数の設定は不要（サーバー秘密情報を持たない）。
3. デプロイ後、HTTPS で配信される。AIアシストを使うユーザーは各自のキーを画面で登録する。

全ページが静的プリレンダリングされるため、配信は実質的に静的。`next.config.ts` の
`headers()` によりセキュリティヘッダが付与される。
