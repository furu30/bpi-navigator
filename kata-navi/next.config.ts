import type { NextConfig } from "next";

// Netlify への静的サイト配信（static export）。
// - `output: "export"` で out/ に静的HTMLを書き出す（サーバーランタイム不要）。
// - 本アプリは localStorage 中心のクライアントアプリで、API Route / middleware /
//   server actions / 動的ルートは使用していないため static export 可能。
// - セキュリティヘッダ（CSP 等）は static export では next.config の headers() が
//   無効になるため、ホスティング側（netlify.toml の [[headers]]）で付与する。
// - next/image は未使用のため images.unoptimized は不要。
const nextConfig: NextConfig = {
  output: "export",
  reactStrictMode: true,
};

export default nextConfig;
