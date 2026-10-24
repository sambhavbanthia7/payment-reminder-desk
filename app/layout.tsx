import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "JAL Receivables Desk",
  description: "Review debtor ageing, record collections, and send payment reminders from JAL's Excel tracker.",
  icons: {
    icon: "/favicon.svg",
    shortcut: "/favicon.svg",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body className="antialiased">{children}</body>
    </html>
  );
}
