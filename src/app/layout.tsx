
import type { Metadata } from 'next';
import { Geist_Mono } from 'next/font/google';
import { Inter } from 'next/font/google';
import Link from 'next/link'; // Import Link
import './globals.css';
import { cn } from '@/lib/utils';
import { Toaster } from '@/components/ui/toaster';
import { Button } from '@/components/ui/button'; // Import Button
import { Home, Settings } from 'lucide-react'; // Import icons

const fontSans = Inter({
  subsets: ['latin'],
  variable: '--font-sans',
});

const fontMono = Geist_Mono({
  variable: '--font-mono',
  subsets: ['latin'],
});

export const metadata: Metadata = {
  title: 'PowerShell GUI',
  description: 'Execute Azure and Office 365 PowerShell commands via GUI',
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" className="dark">
      <body
        className={cn(
          'min-h-screen bg-background font-sans antialiased flex flex-col', // Use flex-col for layout
          fontSans.variable,
          fontMono.variable
        )}
      >
        {/* Basic Navigation Header */}
        <header className="sticky top-0 z-50 w-full border-b border-border/40 bg-background/95 backdrop-blur supports-[backdrop-filter]:bg-background/60">
          <div className="container flex h-14 max-w-screen-2xl items-center">
             <Link href="/" className="mr-6 flex items-center space-x-2">
                {/* Optional Logo/Icon here */}
               <span className="font-bold sm:inline-block">PowerShell GUI</span>
             </Link>
            <nav className="flex items-center gap-4 text-sm ml-auto">
              <Button variant="ghost" size="sm" asChild>
                <Link href="/">
                  <Home className="h-4 w-4 mr-1" /> Home
                </Link>
              </Button>
              <Button variant="ghost" size="sm" asChild>
                <Link href="/settings">
                  <Settings className="h-4 w-4 mr-1" /> Settings
                </Link>
              </Button>
            </nav>
          </div>
        </header>

         {/* Main Content Area - Make it grow */}
         <div className="flex-grow container mx-auto"> {/* Wrap children in a growing div */}
           {children}
         </div>

        <Toaster />
      </body>
    </html>
  );
}
