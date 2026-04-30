import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "轉介競賽計分小工具",
  description: "個人使用的轉介藥局統計 MVP",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-Hant">
      <body>{children}</body>
    </html>
  );
}
