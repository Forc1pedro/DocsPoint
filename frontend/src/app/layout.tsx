import type { Metadata } from 'next';
import { Providers } from './providers';

export const metadata: Metadata = {
  title: 'SharePoint — Documentos',
  description: 'Microsoft SharePoint Modern Clone construído com Fluent UI v9 e Next.js',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="pt-BR">
      <body
        style={{
          margin: 0,
          padding: 0,
          fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif",
          backgroundColor: '#f5f5f5',
          height: '100vh',
          overflow: 'hidden',
        }}
      >
        <Providers>{children}</Providers>
      </body>
    </html>
  );
}
