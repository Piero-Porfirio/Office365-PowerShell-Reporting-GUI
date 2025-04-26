
'use client';

import React, { useState, useCallback } from 'react';
import { SidebarProvider } from '@/components/ui/sidebar';
import { ReportingSidebar } from './reporting-sidebar';
import { ReportDisplay } from './report-display';
import type { ReportCategory } from '@/types/reporting';

export function Reporting() {
  // Start with no report selected to encourage user interaction
  const [selectedReport, setSelectedReport] = useState<ReportCategory | null>(null);
  const [isSidebarOpen, setIsSidebarOpen] = useState(true); // Default sidebar state

  const handleSelectReport = useCallback((report: ReportCategory | null) => {
    setSelectedReport(report);
    // Potentially close mobile sidebar on selection if needed
  }, []);


  return (
    // Ensure SidebarProvider wraps both Sidebar and Display area correctly
    // Use flex layout to position sidebar and content side-by-side
    <div className="flex h-full w-full">
      <SidebarProvider
        defaultOpen={isSidebarOpen}
        onOpenChange={setIsSidebarOpen}
      >
        <ReportingSidebar
          selectedReport={selectedReport}
          onSelectReport={handleSelectReport}
        />
        {/* ReportDisplay should be outside SidebarProvider if it's the main content area */}
        {/* Or ensure SidebarProvider wraps the *entire* layout structure */}
         <ReportDisplay selectedReport={selectedReport} />
      </SidebarProvider>
       {/* If SidebarProvider should only manage the sidebar itself: */}
       {/* <ReportDisplay selectedReport={selectedReport} /> */}
    </div>
  );
}
