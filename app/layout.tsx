import type { Metadata } from "next";
// 這行最重要！負責把全域樣式注入到你的網頁中
import "./globals.css";

export const metadata: Metadata = {
  title: "崁頂庫存系統",
  description: "Live Inventory System",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="zh-TW">
      <body className="bg-[#F5F5F7] text-gray-900">{children}</body>
    </html>
  );
}