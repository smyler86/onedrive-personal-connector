import type { Metadata } from "next";
import "./globals.css";
import "./drive-ui.css";

export const metadata: Metadata = {
  title: "Personal OneDrive",
  description: "Connect your personal Microsoft account and transfer files between ChatGPT and OneDrive.",
  referrer: "no-referrer",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="en"><body>{children}</body></html>;
}
