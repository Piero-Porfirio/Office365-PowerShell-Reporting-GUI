
'use client';

import React, { useState, useCallback } from 'react';
import { SidebarProvider } from '@/components/ui/sidebar';
import { ReportingSidebar } from './reporting-sidebar';
import { ReportDisplay } from './report-display';
import type { ReportCategory } from '@/types/reporting';

export function Reporting() {
  const [selectedReport, setSelectedReport] = useState<ReportCategory | null>(
    'Shared Mailbox Traffic Stats' // Default to a report
  );
  const [isSidebarOpen, setIsSidebarOpen] = useState(true); // Default sidebar state

  const handleSelectReport = useCallback((report: ReportCategory | null) => {
    setSelectedReport(report);
    // Potentially close mobile sidebar on selection if needed
  }, []);


  return (
    <SidebarProvider
      defaultOpen={isSidebarOpen}
      onOpenChange={setIsSidebarOpen}
    >
      <ReportingSidebar
        selectedReport={selectedReport}
        onSelectReport={handleSelectReport}
      />
      <ReportDisplay selectedReport={selectedReport} />
    </SidebarProvider>
  );
}
