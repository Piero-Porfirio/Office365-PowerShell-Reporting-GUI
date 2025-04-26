
'use client';

import React, { useState, useEffect } from 'react';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { ReportTable } from './report-table';
import type { ReportCategory, SharedMailboxTrafficStats } from '@/types/reporting';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { Loader2, Info } from 'lucide-react';

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
    // Add cases for other reports here, returning appropriate mock data or empty arrays for now
    // Azure AD Reports
    case 'License Utilization':
    case 'License Expiry':
    case 'Users & Groups':
    case 'Unlicensed User':
    case 'User Managers & Direct Reports':
    case 'Group Members':
    case 'Group Owners':
    case 'Group-based Licensing':
    case 'Login Activities': // Azure AD
    case 'Password Changes':
    case 'MFA Disabled Users':
    case 'Device Registrations':
    // Security Reports
    case 'MFA Enforced Users':
    case 'MFA Non-Activated Users':
    case 'Password Expiry':
    case 'Users with Weak Passwords':
    case 'External Users':
    case 'Guest Users':
    case 'Risky Login Attempts':
    case 'External User License Assignments':
    case 'Secure Score':
    case 'eDiscovery':
    case 'Non-Owner Mailbox Access':
    case 'Data Loss Prevention (DLP)':
    case 'Advanced Threat Protection':
    // Sign-in Analysis Reports
    case 'User Sign-in Location':
    case 'Last Log-on Summary':
    case 'External User Sign-ins':
    case 'Guest Sign-ins':
    case 'Non-Compliant Device Sign-ins':
    case 'Unmanaged Device Sign-ins':
    case 'MFA failed Sign-ins':
    case '2FA Authentication Methods':
    case 'Conditional Access failures':
    case 'Conditional Access Policies':
    case 'Sign-ins Risk Level':
    case 'Unlikely Travel Risky Sign-ins':
    case 'Anonymous IP Sign-ins':
    case 'Compromised Risky Sign-ins':
    // Email Reports
    case 'Send as Emails':
    case 'Send on Behalf Emails':
    case 'Undelivered Mails':
    case 'Mails Sent by Delegates':
    case 'Email Traffic':
    case 'External Email Forwarding':
    case 'Spam Detections':
    case 'Top Phish Receivers':
    case 'External Spoof Mails':
    case 'Top Malwares':
    case 'Internal Mail flow':
    case 'External Mail flow':
    case "Users' Active Hours":
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
    // Teams Reports
    case 'Public Teams':
    case 'Private Teams':
    case 'Teams Membership':
    case 'Private & Shared Channels':
    // case 'Login Activities': // Teams - Covered by Azure AD? Need specific API/command
    case 'Private Channels':
    case 'Team Setting Changes':
    case 'Inactive Users': // Teams
    case 'Teams Device Usage':
    case 'External File Sharing':
    case 'Teams Add-ons':
    case 'Private Channel Membership Changes':
    case 'Ownership Promotions and Demotions':
    // Exchange Online Reports
    case 'Mailbox Usage':
    case 'Inactive Mailboxes':
    case 'Active Out of Office Settings':
    case 'Mailbox Permissions':
    case 'Audit Disabled Mailboxes':
    case 'Mailbox Hold':
    case 'Mailbox Non-owner Access':
    case 'Mobile Device Configurations':
    case 'Public Folders':
    case 'Role Assignments':
    case 'Mail Flow': // Exchange
    case 'Exchange Contacts':
       return []; // Return empty array for all new reports for now
    default:
      // Ensure exhaustive check or handle unknown cases
      const _exhaustiveCheck: never = report;
      console.warn(`Unhandled report type: ${report}`);
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
          setError(`Failed to load data for ${selectedReport}. Please try again later.`);
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
          <span className="ml-2 text-muted-foreground">Loading report data...</span>
        </div>
      );
    }

    if (error) {
      return (
        <Alert variant="destructive" className="mt-4">
          <AlertTriangle className="h-4 w-4" />
          <AlertTitle>Error Loading Report</AlertTitle>
          <AlertDescription>{error}</AlertDescription>
        </Alert>
      );
    }

    if (!selectedReport) {
       return (
         <Alert className="mt-4 border-primary/30">
           <Info className="h-4 w-4" />
           <AlertTitle>Select a Report</AlertTitle>
           <AlertDescription>Please choose a report category from the sidebar to view the corresponding data.</AlertDescription>
         </Alert>
       );
    }

    // --- Render Specific Reports ---
    // Only 'Shared Mailbox Traffic Stats' is implemented with a table currently.
    // All others will show the 'Not Implemented' or 'No Data' message.
    switch (selectedReport) {
      case 'Shared Mailbox Traffic Stats':
         if (!reportData || (Array.isArray(reportData) && reportData.length === 0)) {
           return (
             <Alert className="mt-4">
                <Info className="h-4 w-4" />
               <AlertTitle>No Data Available</AlertTitle>
               <AlertDescription>There is no data available for the selected report: {selectedReport}.</AlertDescription>
             </Alert>
           );
         }
        return <ReportTable data={reportData as SharedMailboxTrafficStats[]} caption={selectedReport} />;

      // --- Placeholder for all other reports ---
      // Add specific components here as they are built (e.g., charts, different tables)
      // Example:
      // case 'User Email Traffic Stats':
      //   return <UserTrafficChart data={reportData} />;
      // case 'Secure Score':
      //    return <SecureScoreDisplay score={reportData?.score} recommendations={reportData?.recommendations} />

      default:
         // Check if data exists but the component isn't implemented yet
         if (reportData && Array.isArray(reportData) && reportData.length > 0) {
             return (
               <Alert className="mt-4">
                 <Info className="h-4 w-4" />
                 <AlertTitle>Report View Not Implemented</AlertTitle>
                 <AlertDescription>Data loaded for '{selectedReport}', but the specific display component is not yet available.</AlertDescription>
                 {/* Optionally display raw data for debugging */}
                 {/* <pre className="mt-2 text-xs bg-muted p-2 rounded overflow-auto max-h-60">
                   {JSON.stringify(reportData, null, 2)}
                 </pre> */}
               </Alert>
             );
         }
         // If no data was loaded (empty array from mock fetch)
         return (
           <Alert className="mt-4">
             <Info className="h-4 w-4" />
             <AlertTitle>No Data / Not Implemented</AlertTitle>
             <AlertDescription>
               Either there is no data available for '{selectedReport}' or the report functionality is not yet implemented.
             </AlertDescription>
           </Alert>
         );
    }
  };

  return (
    // Use flex-1 and overflow-auto for proper layouting within the tab content
    <main className="flex-1 p-4 md:p-6 overflow-auto">
      <Card className="w-full shadow-lg h-full flex flex-col border-border/50">
        <CardHeader className="border-b border-border/50">
          <CardTitle className="text-xl text-primary">{selectedReport || 'Select a Report'}</CardTitle>
          <CardDescription>
            {selectedReport ? `Viewing details for ${selectedReport}` : 'Choose a report from the sidebar.'}
          </CardDescription>
        </CardHeader>
        <CardContent className="flex-1 overflow-auto pt-4">
           {renderReportContent()}
        </CardContent>
      </Card>
     </main>
  );
}

// --- Placeholder for other report components ---
// Example:
// function UserTrafficChart({ data }: { data: any }) {
//   return <div>User Traffic Chart Component (Not Implemented)</div>;
// }
// function SecureScoreDisplay({ score, recommendations }: { score: number | null, recommendations: any[] | null }) {
//   return <div>Secure Score Component (Not Implemented)</div>;
// }
