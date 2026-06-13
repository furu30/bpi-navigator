import type { Metadata } from "next";
import "./globals.css";
import { AppShell } from "@/components/AppShell";
import { StoreProvider } from "@/components/StoreProvider";

export const metadata: Metadata = {
  title: "KATA Navi｜業務改善ナビ（アプリA）",
  description:
    "中小製造業の現場が、業務のムダを見つけて減らすまでを一本道で回せる業務改善アプリ。",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ja">
      <body>
        <StoreProvider>
          <AppShell>{children}</AppShell>
        </StoreProvider>
      </body>
    </html>
  );
}
