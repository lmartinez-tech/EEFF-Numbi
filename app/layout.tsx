import type { Metadata } from "next";
import { headers } from "next/headers";
import "./globals.css";

const title = "Numbi EEFF | Estados financieros desde tu balance";
const description = "Carga balances de prueba por tercero de Siigo y conviértelos en estados financieros, controles y análisis listos para decidir.";

export async function generateMetadata(): Promise<Metadata> {
  const incoming = await headers();
  const host = incoming.get("x-forwarded-host") || incoming.get("host") || "localhost:5173";
  const protocol = incoming.get("x-forwarded-proto") || (host.startsWith("localhost") ? "http" : "https");
  const image = `${protocol}://${host}/og-numbi-eeff.png`;
  return {
    title,
    description,
    icons: { icon: "/numbi-assistant.png", shortcut: "/numbi-assistant.png" },
    themeColor: "#233037",
    openGraph: { title, description, images: [{ url: image, width: 1792, height: 896, alt: "Numbi EEFF - Del balance a la decisión" }] },
    twitter: { card: "summary_large_image", title, description, images: [image] },
  };
}

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="es">
      <body>{children}</body>
    </html>
  );
}
