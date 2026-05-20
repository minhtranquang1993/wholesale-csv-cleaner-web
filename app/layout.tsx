import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Wholesale CSV Cleaner",
  description: "Tool làm sạch dữ liệu Wholesale CSV online",
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
