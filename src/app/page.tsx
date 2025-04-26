
'use client';

import { Reporting } from '@/components/reporting';
import { FileText } from 'lucide-react';

export default function Home() {
  return (
    // Use flex-col and ensure height fills available space within layout
    <main className="flex flex-col items-center p-4 md:p-6 lg:p-8 bg-background flex-grow overflow-hidden">
      {/* Use max-width for content centering and flex-grow to fill height */}
      <div className="w-full max-w-screen-2xl flex flex-col flex-grow overflow-hidden">
        <h1 className="text-3xl font-bold mb-6 text-center text-primary shrink-0">
          <FileText className="inline-block mr-2 h-8 w-8 align-text-bottom" />
          Azure &amp; Office 365 Reporting
        </h1>
        {/* Reporting component now takes up the entire main content area */}
        <div className="flex flex-col flex-grow overflow-hidden">
          {/* Reporting component manages its sidebar + display area */}
          <Reporting />
        </div>
      </div>
    </main>
  );
}
