import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Data Cleaner",
  description: "Tool làm sạch & pivot dữ liệu CSV / Excel online",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="vi">
      <body className="bg-gray-50 min-h-screen">{children}</body>
    </html>
  );
}
