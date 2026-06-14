import type { Metadata } from "next";
import "./globals.css";
import { AppShell } from "@/components/AppShell";
import { StoreProvider } from "@/components/StoreProvider";

export const metadata: Metadata = {
  title: "KATA Navi｜業務改善 × 再現性ナビ",
  description:
    "中小製造業の現場が、業務のムダを減らし（業務改善ナビ）、属人的な技を受け継ぐ（再現性ナビ）までを一本道で回せるアプリ。仕事を「型」にして、改善し、受け継ぐ。",
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
