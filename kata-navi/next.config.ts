import type { NextConfig } from "next";

const isDev = process.env.NODE_ENV !== "production";

// AIアシストはブラウザから各プロバイダへ直接アクセスする（BYOキー）。
// そのため connect-src に各APIエンドポイントを明示的に許可する。
const AI_ENDPOINTS = [
  "https://api.anthropic.com",
  "https://api.openai.com",
  "https://generativelanguage.googleapis.com",
];

// Content-Security-Policy
// - 本アプリは localStorage 中心のクライアントアプリ。サーバーに業務データ/キーは置かない。
// - Next.js のブートストラップ用インラインscript/styleのため 'unsafe-inline' を許可
//   （静的配信前提・nonceミドルウェアを使わない構成のための割り切り）。
// - 開発時は HMR のため 'unsafe-eval' と ws: を追加で許可。
const csp = [
  `default-src 'self'`,
  `base-uri 'self'`,
  `object-src 'none'`,
  `frame-ancestors 'none'`,
  `form-action 'self'`,
  `img-src 'self' data: blob:`,
  `font-src 'self'`,
  `style-src 'self' 'unsafe-inline'`,
  `script-src 'self' 'unsafe-inline'${isDev ? " 'unsafe-eval'" : ""}`,
  `connect-src 'self' ${AI_ENDPOINTS.join(" ")}${isDev ? " ws: wss:" : ""}`,
  `upgrade-insecure-requests`,
].join("; ");

const securityHeaders = [
  { key: "Content-Security-Policy", value: csp },
  {
    key: "Strict-Transport-Security",
    value: "max-age=63072000; includeSubDomains; preload",
  },
  { key: "X-Frame-Options", value: "DENY" },
  { key: "X-Content-Type-Options", value: "nosniff" },
  { key: "Referrer-Policy", value: "strict-origin-when-cross-origin" },
  { key: "X-DNS-Prefetch-Control", value: "off" },
  {
    key: "Permissions-Policy",
    value: "camera=(), microphone=(), geolocation=(), browsing-topics=()",
  },
];

const nextConfig: NextConfig = {
  // 開発時の潜在バグ検出
  reactStrictMode: true,
  // X-Powered-By を出さない
  poweredByHeader: false,
  async headers() {
    return [
      {
        // すべてのルートに適用
        source: "/:path*",
        headers: securityHeaders,
      },
    ];
  },
};

export default nextConfig;
