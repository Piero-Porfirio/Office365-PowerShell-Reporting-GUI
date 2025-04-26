
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
import { LicenseUtilizationChart } from './license-utilization-chart'; // Import new chart component
import { SharedMailboxTrafficChart } from './shared-mailbox-traffic-chart'; // Import new chart component

interface ReportDisplayProps {
  selectedReport: ReportCategory | null;
}

// Base interface for component props (data and optional caption)
interface ReportComponentProps {
    data: any;
    caption?: string;
}

// --- Report Configuration Mapping ---
// Maps report categories to their commands, services, parsers, and rendering components.
const reportConfig: Record<ReportCategory, {
    command: string;
    service: (command: string) => Promise<PowerShellResult>;
    parser: (output: string | null) => any;
    component: React.ComponentType<ReportComponentProps>; // Use base interface
    simulatedData?: any;
    apiTarget: 'azure' | 'office365';
}> = {
    // --- Azure AD Reports ---
    "License Utilization": {
        command: "Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber, ConsumedUnits, @{N='TotalUnits'; E={$_.PrepaidUnits.Suspended + $_.PrepaidUnits.Enabled + $_.PrepaidUnits.Warning}}",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { const data = output ? JSON.parse(output) as LicenseUtilizationData[] : []; return data.map(d => ({...d, AvailableUnits: d.TotalUnits - d.ConsumedUnits})); } catch (e) { console.error("Error parsing License Utilization:", e); return []; } },
        component: LicenseUtilizationChart, // Use Pie Chart
        simulatedData: [
           { SkuId: 'c7df2760-2c81-4ef7-b578-5b5392b571df', SkuPartNumber: 'ENTERPRISEPREMIUM', ConsumedUnits: 150, TotalUnits: 200 },
           { SkuId: '6f23d6d4-3be2-4f0d-b04d-a59e13b9e9a7', SkuPartNumber: 'VISIOCLIENT', ConsumedUnits: 25, TotalUnits: 50 },
           { SkuId: 'cfc14177-ef78-4702-864d-3f6d63e56f5a', SkuPartNumber: 'PROJECTPROFESSIONAL', ConsumedUnits: 10, TotalUnits: 15 },
         ]
    },
     "Users & Groups": {
       command: "Get-MgUser -Select Id,DisplayName,UserPrincipalName,AssignedLicenses -Filter 'AccountEnabled eq true' | ConvertTo-Json -Depth 3; Get-MgGroup -Select Id,DisplayName,Description,GroupTypes | ConvertTo-Json -Depth 3", // Example combined - Needs better handling
       service: executeAzurePowerShellCommand,
       apiTarget: 'azure',
       parser: (output) => { try { return output ? JSON.parse(output) as AzureADUserGroupInfo[] : []; } catch (e) { console.error("Error parsing Users & Groups:", e); return []; } }, // Needs specific type
       component: ReportTable, // Use Table
        simulatedData: [ // Sample structure (needs refinement for combined data)
          { type: 'User', id: 'user1', displayName: 'Alice Smith', userPrincipalName: 'alice@contoso.onmicrosoft.com', license: 'ENTERPRISEPREMIUM' },
          { type: 'Group', id: 'group1', displayName: 'Sales Team', description: 'Global Sales', groupType: 'Unified' },
          { type: 'User', id: 'user2', displayName: 'Bob Johnson', userPrincipalName: 'bob@contoso.onmicrosoft.com', license: 'ENTERPRISEPREMIUM' },
          { type: 'Group', id: 'group2', displayName: 'Marketing Team', description: 'Product Marketing', groupType: 'Security' },
          { type: 'User', id: 'user3', displayName: 'Charlie Brown', userPrincipalName: 'charlie@contoso.onmicrosoft.com', license: 'VISIOCLIENT' },
        ]
     },
    // --- Email Reports ---
    "Shared Mailbox Traffic Stats": {
        command: "# Placeholder: Needs complex Exchange Online command (e.g., Get-MailTrafficReport, Get-MailboxStatistics combined with message trace)",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) as SharedMailboxTrafficStats[] : []; } catch (e) { console.error("Error parsing Shared Mailbox Stats:", e); return []; } },
        component: SharedMailboxTrafficChart, // Use Bar Chart
        simulatedData: [ // Retaining previous mock data for simulation
          { month: 'April 2025', sharedMailboxName: '24x7support', sharedMailboxUPN: '24x7support@o365droid.onmicrosoft.com', totalMailsSent: 5, totalMailsReceived: 10, externalMailsSent: 2, externalMailsReceived: 3 },
          { month: 'April 2025', sharedMailboxName: 'Amy Tucker', sharedMailboxUPN: 'amy.tucker@o365droid.onmicrosoft.com', totalMailsSent: 15, totalMailsReceived: 80, externalMailsSent: 5, externalMailsReceived: 10 },
          { month: 'April 2025', sharedMailboxName: 'Finance', sharedMailboxUPN: 'finance@o365droid.onmicrosoft.com', totalMailsSent: 2, totalMailsReceived: 190, externalMailsSent: 10, externalMailsReceived: 50 },
          { month: 'April 2025', sharedMailboxName: 'HR', sharedMailboxUPN: 'hr@o365droid.onmicrosoft.com', totalMailsSent: 1, totalMailsReceived: 150, externalMailsSent: 2, externalMailsReceived: 20 },
          { month: 'April 2025', sharedMailboxName: 'Production', sharedMailboxUPN: 'production@o365droid.onmicrosoft.com', totalMailsSent: 8, totalMailsReceived: 40, externalMailsSent: 4, externalMailsReceived: 15 },
          { month: 'April 2025', sharedMailboxName: 'TL', sharedMailboxUPN: 'tl@o365droid.onmicrosoft.com', totalMailsSent: 3, totalMailsReceived: 140, externalMailsSent: 3, externalMailsReceived: 30 },
          { month: 'March 2025', sharedMailboxName: 'Amy Tucker', sharedMailboxUPN: 'amy.tucker@o365droid.onmicrosoft.com', totalMailsSent: 10, totalMailsReceived: 180, externalMailsSent: 18, externalMailsReceived: 40 },
          { month: 'March 2025', sharedMailboxName: 'Finance', sharedMailboxUPN: 'finance@o365droid.onmicrosoft.com', totalMailsSent: 5, totalMailsReceived: 450, externalMailsSent: 19, externalMailsReceived: 100 },
          { month: 'March 2025', sharedMailboxName: 'HR', sharedMailboxUPN: 'hr@o365droid.onmicrosoft.com', totalMailsSent: 2, totalMailsReceived: 290, externalMailsSent: 4, externalMailsReceived: 80 },
        ]
    },
     // --- Add other reports below, following the pattern ---
     // Use ReportTable for now, replace with specific charts/components later
     // Define appropriate parsers and simulated data structures
     // --- Default/Placeholder for unimplemented reports ---
     "License Expiry": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Unlicensed User": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "User Managers & Direct Reports": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Group Members": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Group Owners": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Group-based Licensing": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Login Activities": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Password Changes": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "MFA Disabled Users": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table (Maybe Pie Chart later)
     "Device Registrations": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "MFA Enforced Users": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table (Maybe Pie Chart later)
     "MFA Non-Activated Users": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table (Maybe Pie Chart later)
     "Password Expiry": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Users with Weak Passwords": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "External Users": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Guest Users": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Risky Login Attempts": { command: "# Not Implemented - Needs Identity Protection API", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "External User License Assignments": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Secure Score": { command: "# Not Implemented - Needs MgGraph Secure Score API", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table (Likely needs chart)
     "eDiscovery": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Non-Owner Mailbox Access": { command: "# Not Implemented - Needs Audit Logs", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Data Loss Prevention (DLP)": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Advanced Threat Protection": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "User Sign-in Location": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table (Maybe Map later)
     "Last Log-on Summary": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "External User Sign-ins": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Guest Sign-ins": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Non-Compliant Device Sign-ins": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Unmanaged Device Sign-ins": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "MFA failed Sign-ins": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "2FA Authentication Methods": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Conditional Access failures": { command: "# Not Implemented - Needs MgGraph Audit Logs", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Conditional Access Policies": { command: "# Not Implemented", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Sign-ins Risk Level": { command: "# Not Implemented - Needs Identity Protection API", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table (Maybe Chart)
     "Unlikely Travel Risky Sign-ins": { command: "# Not Implemented - Needs Identity Protection API", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Anonymous IP Sign-ins": { command: "# Not Implemented - Needs Identity Protection API", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Compromised Risky Sign-ins": { command: "# Not Implemented - Needs Identity Protection API", service: executeAzurePowerShellCommand, apiTarget: 'azure', parser: () => [], component: ReportTable }, // Use Table
     "Send as Emails": { command: "# Not Implemented - Needs Audit Logs", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Send on Behalf Emails": { command: "# Not Implemented - Needs Audit Logs", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Undelivered Mails": { command: "# Not Implemented - Needs Message Trace", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Mails Sent by Delegates": { command: "# Not Implemented - Needs Audit Logs", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Email Traffic": { command: "# Not Implemented - Needs Reporting API", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Generic - might need chart)
     "External Email Forwarding": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Spam Detections": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Maybe Chart)
     "Top Phish Receivers": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "External Spoof Mails": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Top Malwares": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Maybe Chart)
     "Internal Mail flow": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "External Mail flow": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Users' Active Hours": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Needs specific API)
     "Email Activities": { command: "# Not Implemented - Needs Reporting API", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Domain-wise Summary": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Group Email Activities": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Org Email Traffic Stats": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Maybe Chart)
     "User Email Traffic Stats": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Email Traffic by 30 min": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Needs Reporting API, chart likely)
     "Hourly Email Traffic": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Needs Reporting API, chart likely)
     "Daily Email Traffic": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Needs Reporting API, chart likely)
     "Monthly Email Traffic": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Needs Reporting API, chart likely)
     "Groups Mail Traffic Stats": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Total Mails By Hour/Day": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Chart likely)
     "Organizations' Total Mails": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "User Total Mails": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Shared Mailbox Total Mails": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Groups Total Mails": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table
     "Peak Period Analysis": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Chart likely)
     "Public Teams": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Teams Module)
     "Private Teams": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Teams Module)
     "Teams Membership": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Teams Module)
     "Private & Shared Channels": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Teams Module)
     // "Login Activities": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, parser: () => [], component: ReportTable }, // Covered by Azure AD
     "Private Channels": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Teams Module)
     "Team Setting Changes": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Teams Audit Logs)
     "Inactive Users": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Teams - Needs criteria)
     "Teams Device Usage": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Teams Reporting API)
     "External File Sharing": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Teams/SharePoint)
     "Teams Add-ons": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Teams Module)
     "Private Channel Membership Changes": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Teams Audit Logs)
     "Ownership Promotions and Demotions": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Teams Audit Logs)
     "Mailbox Usage": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Exchange Module)
     "Inactive Mailboxes": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Exchange Module)
     "Active Out of Office Settings": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Exchange Module)
     "Mailbox Permissions": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Exchange Module)
     "Audit Disabled Mailboxes": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Exchange Module)
     "Mailbox Hold": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Exchange Module)
     "Mailbox Non-owner Access": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Exchange Audit Logs)
     "Mobile Device Configurations": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Exchange/Intune)
     "Public Folders": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Exchange Module)
     "Role Assignments": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Exchange Module)
     "Mail Flow": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Exchange Message Trace)
     "Exchange Contacts": { command: "# Not Implemented", service: executeOffice365PowerShellCommand, apiTarget: 'office365', parser: () => [], component: ReportTable }, // Use Table (Exchange Module)
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
      setIsLoading(false); // Ensure loading stops if config is missing
      setReportData(null);
      setReportCommand(null);
      return;
    }

    setIsLoading(true);
    setError(null);
    setReportData(null);
    setReportCommand(config.command); // Store the command being executed

    // Introduce a small delay for simulation to show loading state
    await new Promise(resolve => setTimeout(resolve, 300));

    try {
       // Handle "Not Implemented" case first
       if (config.command === "# Not Implemented") {
          setError(`Report not implemented: ${report}`);
          setReportData([]); // Ensure empty data
          setIsLoading(false);
          return; // Stop further execution for this report
       }

       // Use simulated data if available in config for stubbed backend behavior
       // Otherwise, call the actual service (which hits the stubbed API)
      let resultOutput: string | null = null;

      if (config.simulatedData) {
          console.log(`Using simulated data for: ${report}`);
          resultOutput = JSON.stringify(config.simulatedData); // Simulate JSON string output
      } else {
          console.log(`Executing command for ${report}: ${config.command}`);
          const result = await config.service(config.command);
          if (result.error) {
              // If the API returns a simulation error, treat it as an error
              if (result.error.includes('simulation not implemented')) {
                  setError(`Report simulation not implemented: ${report}`); // Set error instead of throwing
                  setReportData([]);
                  setIsLoading(false);
                  return;
              }
              throw new Error(result.error); // Throw other errors
          }
          resultOutput = result.output; // Output from API is expected to be JSON string
      }


      const parsedData = config.parser(resultOutput);
      // Check if data is genuinely empty or just an empty array from parser error handling
      if (!parsedData || (Array.isArray(parsedData) && parsedData.length === 0 && resultOutput && resultOutput !== '[]')) {
          // This might indicate a parsing error or the command legitimately returned nothing
          // We'll let the rendering logic handle the 'No Data' message, but log if needed
          console.warn(`Parser for "${report}" returned empty data. Original output:`, resultOutput);
      }
      setReportData(parsedData);

    } catch (err) {
      console.error(`Error fetching data for ${report}:`, err);
      const message = err instanceof Error ? err.message : 'An unknown error occurred';
      // Set error state instead of letting the error bubble up unhandled
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
      // Special handling for "Not Implemented" errors based on the error message set in fetchReportData
      if (error.startsWith('Report not implemented') || error.startsWith('Report simulation not implemented')) {
        return (
          <Alert className="mt-4">
            <Info className="h-4 w-4" />
            <AlertTitle>Report Not Implemented</AlertTitle>
            <AlertDescription>
              {selectedReport ? `The reporting functionality for '${selectedReport}' is not yet implemented or simulated.` : 'Report functionality not implemented.'}
            </AlertDescription>
          </Alert>
        );
      }
      // General error display
      return (
        <Alert variant="destructive" className="mt-4">
          <AlertTriangle className="h-4 w-4" />
          <AlertTitle>Error Loading Report</AlertTitle>
          <AlertDescription>{error}</AlertDescription>
           {reportCommand && reportCommand !== "# Not Implemented" && ( // Show command only if it's not the placeholder
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
     // Check config again in case selectedReport changed during async operation
     if (!config) return null;

     const ReportComponent = config.component;


     if (!reportData || (Array.isArray(reportData) && reportData.length === 0)) {
       return (
         <Alert className="mt-4">
            <Info className="h-4 w-4" />
           <AlertTitle>No Data Available</AlertTitle>
           <AlertDescription>There is no data available for the selected report: {selectedReport}. This might be because the command returned no results or the data is empty.</AlertDescription>
            {reportCommand && reportCommand !== "# Not Implemented" && (
             <code className="mt-2 text-xs p-1 bg-muted/50 rounded block max-w-full overflow-x-auto">
               Command: {reportCommand}
             </code>
           )}
         </Alert>
       );
     }

     // Render the specific component for the report
     // Pass the caption which might be used by the component (e.g., Table caption)
     return <ReportComponent data={reportData} caption={selectedReport} />;
  };

  return (
    // Use flex-1 and overflow-auto for proper layouting within the tab content
    <main className="flex-1 p-4 md:p-6 overflow-auto bg-background"> {/* Ensure background color */}
      <Card className="w-full shadow-lg h-full flex flex-col border-border/50 bg-card"> {/* Ensure card background */}
        <CardHeader className="border-b border-border/50">
          <CardTitle className="text-xl text-primary">{selectedReport || 'Select a Report'}</CardTitle>
          <CardDescription className="text-muted-foreground">
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
