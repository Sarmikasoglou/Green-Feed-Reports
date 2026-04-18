export const metadata = {
  title: 'GreenFeed QC Dashboard',
  description: 'GreenFeed QC + MVH + Treatments app for Vercel'
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body style={{ margin: 0, fontFamily: 'Arial, sans-serif', background: '#f5f7fb' }}>
        {children}
      </body>
    </html>
  );
}
