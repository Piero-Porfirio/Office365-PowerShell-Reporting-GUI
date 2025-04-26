
'use client';

import React, { useState, useEffect } from 'react';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { ReportTable } from './report-table';
import type { ReportCategory, SharedMailboxTrafficStats } from '@/types/reporting';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { Loader2 } from 'lucide-react';

interface ReportDisplayProps {
  selectedReport: ReportCategory | null;
}

// --- Mock Data ---
const mockSharedMailboxStats: SharedMailboxTrafficStats[] = [
  { month: 'April 2025', sharedMailboxName: '24x7support', sharedMailboxUPN: '24x7support@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 1, externalMailsSent: null, externalMailsReceived: null },
  { month: 'April 2025', sharedMailboxName: 'Amy Tucker', sharedMailboxUPN: 'amy.tucker@o365droid.onmicrosoft.com', totalMailsSent: 1, totalMailsReceived: 8, externalMailsSent: 5, externalMailsReceived: 1 },
  { month: 'April 2025', sharedMailboxName: 'Finance', sharedMailboxUPN: 'finance@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 19, externalMailsSent: 10, externalMailsReceived: null },
  { month: 'April 2025', sharedMailboxName: 'HR', sharedMailboxUPN: 'hr@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 15, externalMailsSent: 2, externalMailsReceived: null },
  { month: 'April 2025', sharedMailboxName: 'Production', sharedMailboxUPN: 'production@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 4, externalMailsSent: 4, externalMailsReceived: null },
  { month: 'April 2025', sharedMailboxName: 'TL', sharedMailboxUPN: 'tl@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 14, externalMailsSent: 3, externalMailsReceived: null },
  { month: 'March 2025', sharedMailboxName: 'Amy Tucker', sharedMailboxUPN: 'amy.tucker@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 18, externalMailsSent: 18, externalMailsReceived: null },
  { month: 'March 2025', sharedMailboxName: 'Finance', sharedMailboxUPN: 'finance@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 45, externalMailsSent: 19, externalMailsReceived: null },
  { month: 'March 2025', sharedMailboxName: 'HR', sharedMailboxUPN: 'hr@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 29, externalMailsSent: 4, externalMailsReceived: null },
];
// --- End Mock Data ---


// --- Mock API Fetch Function ---
async function fetchReportData(report: ReportCategory): Promise<any> {
  console.log(`Fetching data for: ${report}`);
  // Simulate network delay
  await new Promise(resolve => setTimeout(resolve, 1000));

  // Return mock data based on the report type
  switch (report) {
    case 'Shared Mailbox Traffic Stats':
      return mockSharedMailboxStats;
    // Add cases for other reports here, returning appropriate mock data or empty arrays
    case 'Email Activities':
    case 'Domain-wise Summary':
    case 'Group Email Activities':
    case 'Org Email Traffic Stats':
    case 'User Email Traffic Stats':
    case 'Email Traffic by 30 min':
    case 'Hourly Email Traffic':
    case 'Daily Email Traffic':
    case 'Monthly Email Traffic':
    case 'Groups Mail Traffic Stats':
    case 'Total Mails By Hour/Day':
    case "Organizations' Total Mails":
    case 'User Total Mails':
    case 'Shared Mailbox Total Mails':
    case 'Groups Total Mails':
    case 'Peak Period Analysis':
       return []; // Return empty array for now
    default:
      return []; // Default empty array
  }
}
// --- End Mock API Fetch Function ---


export function ReportDisplay({ selectedReport }: ReportDisplayProps) {
  const [reportData, setReportData] = useState<any>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (selectedReport) {
      setIsLoading(true);
      setError(null);
      setReportData(null); // Clear previous data

      fetchReportData(selectedReport)
        .then(data => {
          setReportData(data);
        })
        .catch(err => {
          console.error("Error fetching report data:", err);
          setError(`Failed to load data for ${selectedReport}.`);
        })
        .finally(() => {
          setIsLoading(false);
        });
    } else {
      setReportData(null);
      setError(null);
      setIsLoading(false);
    }
  }, [selectedReport]);


  const renderReportContent = () => {
    if (isLoading) {
      return (
        <div className="flex items-center justify-center h-64">
          <Loader2 className="h-8 w-8 animate-spin text-primary" />
          <span className="ml-2">Loading report data...</span>
        </div>
      );
    }

    if (error) {
      return (
        <Alert variant="destructive">
          <AlertTitle>Error</AlertTitle>
          <AlertDescription>{error}</AlertDescription>
        </Alert>
      );
    }

    if (!selectedReport) {
       return (
         <Alert>
           <AlertTitle>No Report Selected</AlertTitle>
           <AlertDescription>Please select a report category from the sidebar to view data.</AlertDescription>
         </Alert>
       );
    }

    if (!reportData || (Array.isArray(reportData) && reportData.length === 0)) {
       return (
         <Alert>
           <AlertTitle>No Data Available</AlertTitle>
           <AlertDescription>There is no data available for the selected report: {selectedReport}.</AlertDescription>
         </Alert>
       );
    }

    // Render specific report components based on selectedReport
    switch (selectedReport) {
      case 'Shared Mailbox Traffic Stats':
        return <ReportTable data={reportData as SharedMailboxTrafficStats[]} />;
      // Add cases for other report types here
      // Example:
      // case 'User Email Traffic Stats':
      //   return <UserTrafficChart data={reportData} />;
      default:
        return (
           <Alert>
             <AlertTitle>Report View Not Implemented</AlertTitle>
             <AlertDescription>The display for '{selectedReport}' is not yet available.</AlertDescription>
           </Alert>
         );
    }
  };

  return (
    // Use SidebarInset for proper layouting with the sidebar
    <main className="flex-1 p-4 md:p-6 lg:p-8 overflow-auto">
      <Card className="w-full shadow-lg h-full flex flex-col">
        <CardHeader>
          <CardTitle className="text-xl">{selectedReport || 'Reporting'}</CardTitle>
          <CardDescription>
            {selectedReport ? `Details for ${selectedReport}` : 'Select a report from the sidebar.'}
          </CardDescription>
        </CardHeader>
        <CardContent className="flex-1 overflow-auto">
           {renderReportContent()}
        </CardContent>
      </Card>
     </main>
  );
}
