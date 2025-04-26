
'use client';

import React, { useState, useEffect, useCallback } from 'react';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { ReportTable } from './report-table';
import type { ReportCategory, SharedMailboxTrafficStats, LicenseUtilizationData, AzureADUserGroupInfo } from '@/types/reporting';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { Loader2, Info, AlertTriangle } from 'lucide-react';
import { executeAzurePowerShellCommand } from '@/services/azure';
import { executeOffice365PowerShellCommand } from '@/services/office365';
import { PowerShellResult } from '@/services/powershell'; // Generic result type

interface ReportDisplayProps {
  selectedReport: ReportCategory | null;
}

// --- Report Configuration Mapping ---
// Maps report categories to their corresponding PowerShell commands,
// the service to call, a parser for the expected result, and the component to render it.
// NOTE: The API endpoints are currently stubbed. This configuration assumes the stubs
//       will return JSON matching the expected structure based on the command sent.
//       Actual PowerShell execution and output parsing would happen on the backend.
const reportConfig: Record<ReportCategory, {
    command: string;
    service: (command: string) => Promise<PowerShellResult>; // Using generic PowerShellResult as base
    parser: (output: string | null) => any; // Parses the string output (simulated JSON)
    component: React.ComponentType<{ data: any; caption?: string }>; // Component to render data
    simulatedData?: any; // For stubbed backend simulation (optional)
    apiTarget: 'azure' | 'office365'; // Specify which API to hit
}> = {
    // --- Azure AD Reports ---
    "License Utilization": {
        command: "Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber, ConsumedUnits, @{N='TotalUnits'; E={$_.PrepaidUnits.Suspended + $_.PrepaidUnits.Enabled + $_.PrepaidUnits.Warning}}",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) as LicenseUtilizationData[] : []; } catch (e) { console.error("Error parsing License Utilization:", e); return []; } },
        component: ReportTable,
         simulatedData: [ // Sample structure
           { SkuId: 'guid1', SkuPartNumber: 'ENTERPRISEPREMIUM', ConsumedUnits: 150, TotalUnits: 200 },
           { SkuId: 'guid2', SkuPartNumber: 'VISIOCLIENT', ConsumedUnits: 25, TotalUnits: 50 },
         ]
    },
     "Users & Groups": {
       command: "Get-MgUser -Select Id,DisplayName,UserPrincipalName,AssignedLicenses -Filter 'AccountEnabled eq true' | ConvertTo-Json -Depth 3; Get-MgGroup -Select Id,DisplayName,Description,GroupTypes | ConvertTo-Json -Depth 3", // Example combined - Needs better handling
       service: executeAzurePowerShellCommand,
       apiTarget: 'azure',
       parser: (output) => { try { return output ? JSON.parse(output) as AzureADUserGroupInfo[] : []; } catch (e) { console.error("Error parsing Users & Groups:", e); return []; } }, // Needs specific type
       component: ReportTable, // Might need a custom component later
        simulatedData: [ // Sample structure (needs refinement for combined data)
          { type: 'User', id: 'user1', displayName: 'Alice Smith', userPrincipalName: 'alice@example.com', license: 'E3' },
          { type: 'Group', id: 'group1', displayName: 'Sales Team', description: 'Global Sales', groupType: 'Unified' },
        ]
     },
    // --- Email Reports ---
    "Shared Mailbox Traffic Stats": {
        command: "# Placeholder: Needs complex Exchange Online command (e.g., Get-MailTrafficReport, Get-MailboxStatistics combined with message trace)",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) as SharedMailboxTrafficStats[] : []; } catch (e) { console.error("Error parsing Shared Mailbox Stats:", e); return []; } },
        component: ReportTable,
        simulatedData: [ // Retaining previous mock data for simulation
          { month: 'April 2025', sharedMailboxName: '24x7support', sharedMailboxUPN: '24x7support@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 1, externalMailsSent: null, externalMailsReceived: null },
          { month: 'April 2025', sharedMailboxName: 'Amy Tucker', sharedMailboxUPN: 'amy.tucker@o365droid.onmicrosoft.com', totalMailsSent: 1, totalMailsReceived: 8, externalMailsSent: 5, externalMailsReceived: 1 },
          { month: 'March 2025', sharedMailboxName: 'Finance', sharedMailboxUPN: 'finance@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 45, externalMailsSent: 19, externalMailsReceived: null },
        ]
    },
     // --- Add other reports below, following the pattern ---
     // Use ReportTable for now, replace with specific charts/components later
     // Define appropriate parsers and simulated data structures
     // --- Default/Placeholder for unimplemented reports ---
     // (Using a helper function later might be cleaner)
     "License Expiry": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Unlicensed User": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "User Managers & Direct Reports": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Group Members": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Group Owners": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Group-based Licensing": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Login Activities": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Password Changes": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "MFA Disabled Users": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Device Registrations": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "MFA Enforced Users": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "MFA Non-Activated Users": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Password Expiry": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Users with Weak Passwords": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "External Users": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Guest Users": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Risky Login Attempts": { command: "# Not Implemented - Needs Identity Protection API", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "External User License Assignments": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Secure Score": { command: "# Not Implemented - Needs MgGraph Secure Score API", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Likely needs chart
     "eDiscovery": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Non-Owner Mailbox Access": { command: "# Not Implemented - Needs Audit Logs", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Data Loss Prevention (DLP)": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Advanced Threat Protection": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "User Sign-in Location": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Last Log-on Summary": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "External User Sign-ins": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Guest Sign-ins": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Non-Compliant Device Sign-ins": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Unmanaged Device Sign-ins": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "MFA failed Sign-ins": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "2FA Authentication Methods": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Conditional Access failures": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Conditional Access Policies": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Sign-ins Risk Level": { command: "# Not Implemented - Needs Identity Protection API", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Unlikely Travel Risky Sign-ins": { command: "# Not Implemented - Needs Identity Protection API", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Anonymous IP Sign-ins": { command: "# Not Implemented - Needs Identity Protection API", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Compromised Risky Sign-ins": { command: "# Not Implemented - Needs Identity Protection API", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable },
     "Send as Emails": { command: "# Not Implemented - Needs Audit Logs", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Send on Behalf Emails": { command: "# Not Implemented - Needs Audit Logs", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Undelivered Mails": { command: "# Not Implemented - Needs Message Trace", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Mails Sent by Delegates": { command: "# Not Implemented - Needs Audit Logs", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Email Traffic": { command: "# Not Implemented - Needs Reporting API", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Generic - might need chart
     "External Email Forwarding": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Spam Detections": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Top Phish Receivers": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "External Spoof Mails": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Top Malwares": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Internal Mail flow": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "External Mail flow": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Users' Active Hours": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Needs specific API
     "Email Activities": { command: "# Not Implemented - Needs Reporting API", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Domain-wise Summary": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Group Email Activities": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Org Email Traffic Stats": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "User Email Traffic Stats": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Email Traffic by 30 min": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Needs Reporting API, chart likely
     "Hourly Email Traffic": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Needs Reporting API, chart likely
     "Daily Email Traffic": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Needs Reporting API, chart likely
     "Monthly Email Traffic": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Needs Reporting API, chart likely
     "Groups Mail Traffic Stats": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Total Mails By Hour/Day": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Chart likely
     "Organizations' Total Mails": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "User Total Mails": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Shared Mailbox Total Mails": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Groups Total Mails": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable },
     "Peak Period Analysis": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Chart likely
     "Public Teams": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Teams Module
     "Private Teams": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Teams Module
     "Teams Membership": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Teams Module
     "Private & Shared Channels": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Teams Module
     // "Login Activities": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, parser: () => [], component: ReportTable }, // Covered by Azure AD
     "Private Channels": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Teams Module
     "Team Setting Changes": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Teams Audit Logs
     "Inactive Users": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Teams - Needs criteria
     "Teams Device Usage": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Teams Reporting API
     "External File Sharing": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Teams/SharePoint
     "Teams Add-ons": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Teams Module
     "Private Channel Membership Changes": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Teams Audit Logs
     "Ownership Promotions and Demotions": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Teams Audit Logs
     "Mailbox Usage": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Exchange Module
     "Inactive Mailboxes": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Exchange Module
     "Active Out of Office Settings": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Exchange Module
     "Mailbox Permissions": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Exchange Module
     "Audit Disabled Mailboxes": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Exchange Module
     "Mailbox Hold": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Exchange Module
     "Mailbox Non-owner Access": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Exchange Audit Logs
     "Mobile Device Configurations": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Exchange/Intune
     "Public Folders": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Exchange Module
     "Role Assignments": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Exchange Module
     "Mail Flow": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Exchange Message Trace
     "Exchange Contacts": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Exchange Module
};


export function ReportDisplay({ selectedReport }: ReportDisplayProps) {
  const [reportData, setReportData] = useState<any>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [reportCommand, setReportCommand] = useState<string | null>(null);

  const fetchReportData = useCallback(async (report: ReportCategory) => {
    const config = reportConfig[report];
    if (!config) {
      console.error(`No configuration found for report: ${report}`);
      setError(`Configuration missing for report: ${report}.`);
      return;
    }

    setIsLoading(true);
    setError(null);
    setReportData(null);
    setReportCommand(config.command); // Store the command being executed

    try {
       // Use simulated data if available in config for stubbed backend behavior
       // Otherwise, call the actual service (which hits the stubbed API)
      let resultOutput: string | null = null;
      if (config.simulatedData) {
          console.log(`Using simulated data for: ${report}`);
          resultOutput = JSON.stringify(config.simulatedData); // Simulate JSON string output
      } else if (config.command !== "# Not Implemented") {
          console.log(`Executing command for ${report}: ${config.command}`);
          const result = await config.service(config.command);
          if (result.error) {
              throw new Error(result.error);
          }
          resultOutput = result.output; // Output from API is expected to be JSON string
      } else {
           // Handle "Not Implemented" case gracefully
           console.log(`Report not implemented: ${report}`);
           resultOutput = "[]"; // Simulate empty array for unimplemented reports
      }


      const parsedData = config.parser(resultOutput);
      setReportData(parsedData);

    } catch (err) {
      console.error(`Error fetching data for ${report}:`, err);
      const message = err instanceof Error ? err.message : 'An unknown error occurred';
      setError(`Failed to load data for ${report}. Error: ${message}`);
      setReportData([]); // Ensure data is empty on error
    } finally {
      setIsLoading(false);
    }
  }, []); // Empty dependency array as fetchReportData relies on the constant `reportConfig`

  useEffect(() => {
    if (selectedReport) {
      fetchReportData(selectedReport);
    } else {
      // Clear state when no report is selected
      setReportData(null);
      setError(null);
      setIsLoading(false);
      setReportCommand(null);
    }
  }, [selectedReport, fetchReportData]);


  const renderReportContent = () => {
    if (isLoading) {
      return (
        <div className="flex flex-col items-center justify-center h-64">
          <Loader2 className="h-8 w-8 animate-spin text-primary" />
          <span className="mt-2 text-muted-foreground">Loading report data...</span>
          {reportCommand && reportCommand !== "# Not Implemented" && (
            <code className="mt-4 text-xs p-2 bg-muted rounded max-w-full overflow-x-auto">
              Executing: {reportCommand}
            </code>
          )}
        </div>
      );
    }

    if (error) {
      return (
        <Alert variant="destructive" className="mt-4">
          <AlertTriangle className="h-4 w-4" />
          <AlertTitle>Error Loading Report</AlertTitle>
          <AlertDescription>{error}</AlertDescription>
           {reportCommand && (
             <code className="mt-2 text-xs p-1 bg-destructive/20 rounded block max-w-full overflow-x-auto">
               Command: {reportCommand}
             </code>
           )}
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

     const config = reportConfig[selectedReport];
     const ReportComponent = config.component;

     if (config.command === "# Not Implemented") {
         return (
           <Alert className="mt-4">
             <Info className="h-4 w-4" />
             <AlertTitle>Report Not Implemented</AlertTitle>
             <AlertDescription>
               The reporting functionality for '{selectedReport}' is not yet implemented.
             </AlertDescription>
           </Alert>
         );
     }


     if (!reportData || (Array.isArray(reportData) && reportData.length === 0)) {
       return (
         <Alert className="mt-4">
            <Info className="h-4 w-4" />
           <AlertTitle>No Data Available</AlertTitle>
           <AlertDescription>There is no data available for the selected report: {selectedReport}. This might be because the command returned no results or is not fully implemented yet.</AlertDescription>
            {reportCommand && (
             <code className="mt-2 text-xs p-1 bg-muted/50 rounded block max-w-full overflow-x-auto">
               Command: {reportCommand}
             </code>
           )}
         </Alert>
       );
     }

     // Render the specific component for the report
     return <ReportComponent data={reportData} caption={selectedReport} />;
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
