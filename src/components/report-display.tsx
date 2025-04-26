
'use client';

import React, { useState, useEffect, useCallback } from 'react';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { ReportTable } from './report-table';
import type { ReportCategory, SharedMailboxTrafficStats, LicenseUtilizationData, AzureADUserGroupInfo, VisualizationType } from '@/types/reporting';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { Loader2, Info, AlertTriangle, BarChart, PieChartIcon, TableIcon } from 'lucide-react'; // Added icons
import { executeAzurePowerShellCommand } from '@/services/azure';
import { executeOffice365PowerShellCommand } from '@/services/office365';
import { PowerShellResult } from '@/services/powershell'; // Generic result type
import { LicenseUtilizationChart } from './license-utilization-chart'; // Import new chart component
import { SharedMailboxTrafficChart } from './shared-mailbox-traffic-chart'; // Import new chart component
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select"; // Import Select components
import { Label } from '@/components/ui/label'; // Import Label

interface ReportDisplayProps {
  selectedReport: ReportCategory | null;
}

// Base interface for component props (data and optional caption)
interface ReportComponentProps {
    data: any;
    caption?: string;
}

// Map VisualizationType to Component and Icon
const visualizationComponents: Record<VisualizationType, { component: React.ComponentType<ReportComponentProps>, icon: React.ElementType }> = {
  Table: { component: ReportTable, icon: TableIcon },
  PieChart: { component: LicenseUtilizationChart, icon: PieChartIcon },
  BarChart: { component: SharedMailboxTrafficChart, icon: BarChart },
  // Add other chart types here as needed
  // LineChart: { component: LineChartComponent, icon: LineChartIcon },
};

interface ReportConfigEntry {
    command: string;
    service: (command: string) => Promise<PowerShellResult>;
    parser: (output: string | null) => any;
    visualizations: VisualizationType[]; // List available visualization types
    defaultVisualization: VisualizationType; // Default view for this report
    simulatedData?: any;
    apiTarget: 'azure' | 'office365';
}


// --- Report Configuration Mapping ---
// Maps report categories to their commands, services, parsers, and rendering components.
const reportConfig: Record<ReportCategory, ReportConfigEntry> = {
    // --- Azure AD Reports ---
    "License Utilization": {
        command: "Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber, ConsumedUnits, @{N='TotalUnits'; E={$_.PrepaidUnits.Enabled + $_.PrepaidUnits.Suspended + $_.PrepaidUnits.Warning}} | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { const data = output ? JSON.parse(output) as LicenseUtilizationData[] : []; return data.map(d => ({...d, AvailableUnits: (d.TotalUnits || 0) - (d.ConsumedUnits || 0)})); } catch (e) { console.error("Error parsing License Utilization:", e); return []; } },
        visualizations: ['Table', 'PieChart'], // Supports Table and Pie Chart
        defaultVisualization: 'PieChart',
        simulatedData: [
           { SkuId: 'c7df2760-2c81-4ef7-b578-5b5392b571df', SkuPartNumber: 'ENTERPRISEPREMIUM', ConsumedUnits: 150, TotalUnits: 200 },
           { SkuId: '6f23d6d4-3be2-4f0d-b04d-a59e13b9e9a7', SkuPartNumber: 'VISIOCLIENT', ConsumedUnits: 25, TotalUnits: 50 },
           { SkuId: 'cfc14177-ef78-4702-864d-3f6d63e56f5a', SkuPartNumber: 'PROJECTPROFESSIONAL', ConsumedUnits: 10, TotalUnits: 15 },
         ]
    },
    "License Expiry": {
        command: "Get-MgSubscription | Select-Object SkuId, SkuPartNumber, Id, ExpirationDateTime | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing License Expiry:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { SkuId: 'c7df2760-2c81-4ef7-b578-5b5392b571df', SkuPartNumber: 'ENTERPRISEPREMIUM', Id: 'sub1', ExpirationDateTime: '2025-12-31T23:59:59Z' } ]
    },
     "Users & Groups": {
       command: "(Get-MgUser -Select Id,DisplayName,UserPrincipalName,AssignedLicenses -Filter 'AccountEnabled eq true' -Top 999 | Select *, @{N='Type'; E={'User'}}) + (Get-MgGroup -Select Id,DisplayName,Description,GroupTypes -Top 999 | Select *, @{N='Type'; E={'Group'}}) | ConvertTo-Json -Depth 3",
       service: executeAzurePowerShellCommand,
       apiTarget: 'azure',
       parser: (output) => { try { return output ? JSON.parse(output) as AzureADUserGroupInfo[] : []; } catch (e) { console.error("Error parsing Users & Groups:", e); return []; } },
       visualizations: ['Table'],
       defaultVisualization: 'Table',
        simulatedData: [
          { type: 'User', id: 'user1', displayName: 'Alice Smith', userPrincipalName: 'alice@contoso.onmicrosoft.com', license: 'ENTERPRISEPREMIUM' },
          { type: 'Group', id: 'group1', displayName: 'Sales Team', description: 'Global Sales', groupType: 'Unified' },
          { type: 'User', id: 'user2', displayName: 'Bob Johnson', userPrincipalName: 'bob@contoso.onmicrosoft.com', license: 'ENTERPRISEPREMIUM' },
          { type: 'Group', id: 'group2', displayName: 'Marketing Team', description: 'Product Marketing', groupType: 'Security' },
          { type: 'User', id: 'user3', displayName: 'Charlie Brown', userPrincipalName: 'charlie@contoso.onmicrosoft.com', license: 'VISIOCLIENT' },
        ]
     },
     "Unlicensed User": {
        command: "Get-MgUser -Filter 'assignedLicenses/$count eq 0 and userType eq 'Member'' -Select Id,DisplayName,UserPrincipalName,CreatedDateTime | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Unlicensed User:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { id: 'user4', displayName: 'Unlicensed Dave', userPrincipalName: 'dave@contoso.onmicrosoft.com', CreatedDateTime: '2024-01-15T10:00:00Z' } ]
    },
    "User Managers & Direct Reports": {
        command: "Get-MgUser -UserId 'USER_ID_PLACEHOLDER' -ExpandProperty Manager,DirectReports | Select Id,DisplayName,@{N='Manager';E={$_.Manager.DisplayName}},@{N='DirectReports';E={$_.DirectReports.DisplayName -join ', '}} | ConvertTo-Json", // Needs user ID
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Managers/Direct Reports:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { id: 'user1', displayName: 'Alice Smith', Manager: 'Manager Mary', DirectReports: 'Bob Johnson, Charlie Brown' } ]
    },
    "Group Members": {
        command: "Get-MgGroupMember -GroupId 'GROUP_ID_PLACEHOLDER' -All | Select-Object Id, DisplayName, @{N='ObjectType';E={$_.'@odata.type'.Split('.')[-1]}} | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Group Members:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { id: 'user1', displayName: 'Alice Smith', ObjectType: 'User'}, { id: 'user2', displayName: 'Bob Johnson', ObjectType: 'User' } ]
    },
    "Group Owners": {
        command: "Get-MgGroupOwner -GroupId 'GROUP_ID_PLACEHOLDER' -All | Select-Object Id, DisplayName, UserPrincipalName | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Group Owners:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { id: 'owner1', displayName: 'Owner Olivia', userPrincipalName: 'olivia@contoso.onmicrosoft.com' } ]
    },
    "Group-based Licensing": {
        command: "Get-MgGroup -Filter 'assignedLicenses/$count gt 0' -Select Id,DisplayName,AssignedLicenses | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Group Licenses:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { id: 'group1', displayName: 'Sales Team', AssignedLicenses: [{SkuId: 'c7df2760-2c81-4ef7-b578-5b5392b571df'}] } ]
    },
    "Login Activities": {
        command: "Get-MgAuditLogSignIn -Top 100 -Filter \"createdDateTime ge $(Get-Date).AddDays(-1).ToString('yyyy-MM-ddTHH:mm:ssZ')\" | Select CreatedDateTime,UserDisplayName,UserPrincipalName,AppDisplayName,IpAddress,Location,Status | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Login Activities:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { CreatedDateTime: '2024-05-20T10:00:00Z', UserDisplayName: 'Alice Smith', UserPrincipalName: 'alice@contoso.com', AppDisplayName: 'Office 365 Portal', IpAddress: '1.2.3.4', Location: 'City, Country', Status: 'Success' } ]
    },
    "Password Changes": {
        command: "Get-MgAuditLogDirectoryAudit -Top 100 -Filter \"activityDisplayName eq 'Change password' and loggedByService eq 'Self-service Password Reset'\" | Select ActivityDateTime,InitiatedBy,TargetResources | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Password Changes:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { ActivityDateTime: '2024-05-19T09:30:00Z', InitiatedBy: 'Bob Johnson', TargetResources: [{ UserPrincipalName: 'bob@contoso.com' }] } ]
    },
    "MFA Disabled Users": {
        command: "Get-MgUser -Filter 'userType eq 'Member'' -Select Id,DisplayName,UserPrincipalName,RefreshTokensValidFromDateTime | ConvertTo-Json # NOTE: This command doesn't directly show MFA status. Needs Identity/Auth Methods API.",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing MFA Disabled:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { id: 'user5', displayName: 'MFA Maybe Disabled', userPrincipalName: 'maybe@contoso.com' } ]
    },
    "Device Registrations": {
        command: "Get-MgDevice -Top 100 | Select DeviceId,DisplayName,OperatingSystem,TrustType,IsManaged,RegistrationDateTime | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Device Registrations:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { DeviceId: 'dev1', DisplayName: 'Laptop-Alice', OperatingSystem: 'Windows 11', TrustType: 'AzureAD', IsManaged: true, RegistrationDateTime: '2024-03-10T11:00:00Z' } ]
    },

    // --- Security Reports ---
    "MFA Enforced Users": {
        command: "Get-MgPolicyAuthenticationStrengthPolicy -ExpandProperty CombinationConfigurations | ConvertTo-Json # NOTE: Requires parsing policy, doesn't list users directly.",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing MFA Enforced:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { policyName: 'Strong Auth Policy', appliesTo: 'All Users' } ]
    },
    "MFA Non-Activated Users": {
        command: "Get-MgUser -Filter 'userType eq 'Member'' | ForEach-Object { Get-MgUserAuthenticationMethod -UserId $_.Id } | Where-Object { $_ -eq $null } | Select Id # NOTE: Very inefficient, needs better API call.",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing MFA Non-Activated:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { id: 'user6', displayName: 'MFA Not Activated Yet', userPrincipalName: 'no-mfa@contoso.com'} ]
    },
    "Password Expiry": {
        command: "Get-MgUser -Filter 'userType eq 'Member'' -Select Id,DisplayName,PasswordPolicies,PasswordLastSetDateTime | ConvertTo-Json # NOTE: Requires policy check.",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Password Expiry:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { id: 'user1', displayName: 'Alice Smith', PasswordPolicies: 'DisablePasswordExpiration', PasswordLastSetDateTime: '2024-01-01T00:00:00Z' } ]
    },
     "Users with Weak Passwords": {
        command: "# Requires Identity Protection API/Reports - Get-MgRiskDetection?",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: () => [], // Placeholder
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [] // No easy way to simulate without specific API structure
     },
     "External Users": {
        command: "Get-MgUser -Filter \"userType eq 'Guest'\" -Select Id,DisplayName,UserPrincipalName,CreatedDateTime,SignInActivity | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing External Users:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { id: 'guest1', displayName: 'Guest User One', userPrincipalName: 'guest_example.com#EXT#@contoso.onmicrosoft.com', CreatedDateTime: '2024-04-01T12:00:00Z' } ]
    },
    "Guest Users": {
        command: "Get-MgUser -Filter \"userType eq 'Guest'\" -Select Id,DisplayName,UserPrincipalName,CreatedDateTime,SignInActivity | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Guest Users:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { id: 'guest1', displayName: 'Guest User One', userPrincipalName: 'guest_example.com#EXT#@contoso.onmicrosoft.com', CreatedDateTime: '2024-04-01T12:00:00Z' } ]
    },
     "Risky Login Attempts": {
        command: "Get-MgRiskDetection -Top 100 | Where-Object {$_.RiskState -ne 'remediated'} | Select UserDisplayName,RiskLevel,RiskState,DetectedDateTime,IpAddress,Location | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Risky Logins:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { UserDisplayName: 'Bob Johnson', RiskLevel: 'medium', RiskState: 'atRisk', DetectedDateTime: '2024-05-20T08:00:00Z', IpAddress: '10.0.0.5', Location: 'Unknown' } ]
     },
     "External User License Assignments": {
        command: "Get-MgUser -Filter \"userType eq 'Guest'\" -Select Id,DisplayName,AssignedLicenses | Where-Object {$_.AssignedLicenses.Count -gt 0} | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Ext User Licenses:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { id: 'guest2', displayName: 'Licensed Guest', AssignedLicenses: [{ SkuId: 'some-license-sku' }] } ]
     },
     "Secure Score": {
        command: "Get-MgSecureScore | Select CurrentScore, MaxScore, LastModifiedDateTime | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Secure Score:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { CurrentScore: 75, MaxScore: 100, LastModifiedDateTime: '2024-05-20T00:00:00Z' } ]
     },
     "eDiscovery": {
        command: "# Requires Compliance module/API - Get-ComplianceSearch?",
        service: executeOffice365PowerShellCommand, // Might be Office 365 side
        apiTarget: 'office365',
        parser: () => [],
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: []
     },
     "Non-Owner Mailbox Access": {
        command: "Search-UnifiedAuditLog -RecordType ExchangeAdmin -Operations 'Add-MailboxPermission','Remove-MailboxPermission','Set-Mailbox' -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | ConvertTo-Json # Needs parsing",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Non-Owner Access:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Timestamp: '2024-05-18T14:00:00Z', User: 'admin@contoso.com', Operation: 'Add-MailboxPermission', Target: 'finance@contoso.com', Details: 'User admin added FullAccess for support@contoso.com' } ]
     },
     "Data Loss Prevention (DLP)": {
        command: "Get-DlpComplianceRule -Policy 'POLICY_NAME_PLACEHOLDER' | ConvertTo-Json # Requires policy name",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing DLP:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { RuleName: 'Block SSN', Policy: 'Global DLP Policy', State: 'Enabled' } ]
     },
     "Advanced Threat Protection": {
        command: "Get-ATPPolicyForO365 | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing ATP:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Name: 'Default ATP Policy', SafeLinksEnabled: true, SafeAttachmentsEnabled: true } ]
     },

    // --- Sign-in Analysis Reports ---
    "User Sign-in Location": {
        command: "Get-MgAuditLogSignIn -Top 100 -Filter \"createdDateTime ge $(Get-Date).AddDays(-1).ToString('yyyy-MM-ddTHH:mm:ssZ')\" | Select CreatedDateTime,UserDisplayName,IpAddress,Location | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Sign-in Location:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { CreatedDateTime: '2024-05-20T10:00:00Z', UserDisplayName: 'Alice Smith', IpAddress: '1.2.3.4', Location: 'City, Country' } ]
    },
     "Last Log-on Summary": {
        command: "Get-MgUser -Filter 'userType eq 'Member'' -Select Id,DisplayName,UserPrincipalName,SignInActivity | Select Id,DisplayName,UserPrincipalName,@{N='LastSignInDateTime';E={$_.SignInActivity.LastSignInDateTime}} | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Last Logon:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { id: 'user1', displayName: 'Alice Smith', userPrincipalName: 'alice@contoso.com', LastSignInDateTime: '2024-05-20T10:00:00Z' } ]
     },
     "External User Sign-ins": {
        command: "Get-MgAuditLogSignIn -Top 100 -Filter \"createdDateTime ge $(Get-Date).AddDays(-1).ToString('yyyy-MM-ddTHH:mm:ssZ') and userType eq 'Guest'\" | Select CreatedDateTime,UserDisplayName,UserPrincipalName,AppDisplayName,IpAddress,Location | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing External Sign-ins:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { CreatedDateTime: '2024-05-19T11:00:00Z', UserDisplayName: 'Guest User One', UserPrincipalName: 'guest#EXT#@contoso.com', AppDisplayName: 'Shared App', IpAddress: '5.6.7.8', Location: 'Another City, Country' } ]
    },
    "Guest Sign-ins": {
        command: "Get-MgAuditLogSignIn -Top 100 -Filter \"createdDateTime ge $(Get-Date).AddDays(-1).ToString('yyyy-MM-ddTHH:mm:ssZ') and userType eq 'Guest'\" | Select CreatedDateTime,UserDisplayName,UserPrincipalName,AppDisplayName,IpAddress,Location | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Guest Sign-ins:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
         simulatedData: [ { CreatedDateTime: '2024-05-19T11:00:00Z', UserDisplayName: 'Guest User One', UserPrincipalName: 'guest#EXT#@contoso.com', AppDisplayName: 'Shared App', IpAddress: '5.6.7.8', Location: 'Another City, Country' } ]
    },
    "Non-Compliant Device Sign-ins": {
        command: "Get-MgAuditLogSignIn -Top 100 -Filter \"createdDateTime ge $(Get-Date).AddDays(-7).ToString('yyyy-MM-ddTHH:mm:ssZ') and deviceDetail.isCompliant eq false\" | Select CreatedDateTime,UserDisplayName,AppDisplayName,DeviceDetail | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Non-Compliant Sign-ins:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { CreatedDateTime: '2024-05-18T10:00:00Z', UserDisplayName: 'Bob Johnson', AppDisplayName: 'Risky App', DeviceDetail: { deviceId: 'dev2', isCompliant: false } } ]
    },
    "Unmanaged Device Sign-ins": {
        command: "Get-MgAuditLogSignIn -Top 100 -Filter \"createdDateTime ge $(Get-Date).AddDays(-7).ToString('yyyy-MM-ddTHH:mm:ssZ') and deviceDetail.isManaged eq false\" | Select CreatedDateTime,UserDisplayName,AppDisplayName,DeviceDetail | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Unmanaged Sign-ins:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { CreatedDateTime: '2024-05-17T10:00:00Z', UserDisplayName: 'Charlie Brown', AppDisplayName: 'Webmail', DeviceDetail: { deviceId: 'dev3', isManaged: false } } ]
    },
    "MFA failed Sign-ins": {
        command: "Get-MgAuditLogSignIn -Top 100 -Filter \"createdDateTime ge $(Get-Date).AddDays(-7).ToString('yyyy-MM-ddTHH:mm:ssZ') and status.errorCode eq 50074\" | Select CreatedDateTime,UserDisplayName,AppDisplayName,Status | ConvertTo-Json # Error code 50074 is MFA required",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing MFA Failed Sign-ins:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { CreatedDateTime: '2024-05-16T15:00:00Z', UserDisplayName: 'Alice Smith', AppDisplayName: 'Office Portal', Status: { errorCode: 50074, failureReason: 'MFA required but not provided.' } } ]
    },
    "2FA Authentication Methods": {
        command: "Get-MgUser -Filter 'userType eq 'Member'' | ForEach-Object { Get-MgUserAuthenticationMethod -UserId $_.Id | Select-Object *, @{N='UserPrincipalName';E={$_.UserPrincipalName}} } | ConvertTo-Json # Needs user iteration",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing 2FA Methods:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { UserPrincipalName: 'alice@contoso.com', MethodType: 'Microsoft Authenticator', Detail: 'Registered' } ]
    },
    "Conditional Access failures": {
        command: "Get-MgAuditLogSignIn -Top 100 -Filter \"createdDateTime ge $(Get-Date).AddDays(-7).ToString('yyyy-MM-ddTHH:mm:ssZ') and conditionalAccessStatus eq 'failure'\" | Select CreatedDateTime,UserDisplayName,AppDisplayName,ConditionalAccessStatus,Status | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing CA Failures:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { CreatedDateTime: '2024-05-19T12:00:00Z', UserDisplayName: 'Bob Johnson', AppDisplayName: 'Legacy App', ConditionalAccessStatus: 'failure', Status: { failureReason: 'Device not compliant' } } ]
    },
    "Conditional Access Policies": {
        command: "Get-MgIdentityConditionalAccessPolicy | Select DisplayName,State,Conditions,GrantControls | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing CA Policies:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { DisplayName: 'Block Legacy Auth', State: 'enabled', Conditions: '...', GrantControls: '...' } ]
    },
     "Sign-ins Risk Level": {
        command: "Get-MgAuditLogSignIn -Top 100 -Filter \"createdDateTime ge $(Get-Date).AddDays(-7).ToString('yyyy-MM-ddTHH:mm:ssZ') and riskLevelDuringSignIn ne null\" | Select CreatedDateTime,UserDisplayName,RiskLevelDuringSignIn,RiskState,IpAddress,Location | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Sign-in Risk:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { CreatedDateTime: '2024-05-20T08:00:00Z', UserDisplayName: 'Bob Johnson', RiskLevelDuringSignIn: 'medium', RiskState: 'atRisk', IpAddress: '10.0.0.5', Location: 'Unknown' } ]
     },
     "Unlikely Travel Risky Sign-ins": {
        command: "Get-MgRiskDetection -Filter \"riskType eq 'unlikelyTravel'\" -Top 100 | Select UserDisplayName,RiskLevel,RiskState,DetectedDateTime,IpAddress,Location | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Unlikely Travel:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { UserDisplayName: 'Charlie Brown', RiskLevel: 'high', RiskState: 'confirmedCompromised', DetectedDateTime: '2024-05-15T10:00:00Z', IpAddress: '9.8.7.6', Location: 'Far Away Country' } ]
     },
     "Anonymous IP Sign-ins": {
        command: "Get-MgRiskDetection -Filter \"riskType eq 'anonymousIPAddress'\" -Top 100 | Select UserDisplayName,RiskLevel,RiskState,DetectedDateTime,IpAddress,Location | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Anonymous IP:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { UserDisplayName: 'Alice Smith', RiskLevel: 'low', RiskState: 'atRisk', DetectedDateTime: '2024-05-14T11:00:00Z', IpAddress: 'Tor Exit Node', Location: 'Unknown' } ]
     },
     "Compromised Risky Sign-ins": {
        command: "Get-MgRiskDetection -Filter \"riskState eq 'confirmedCompromised'\" -Top 100 | Select UserDisplayName,RiskType,RiskLevel,DetectedDateTime,IpAddress,Location | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Compromised Sign-ins:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { UserDisplayName: 'Charlie Brown', RiskType: 'unlikelyTravel', RiskLevel: 'high', DetectedDateTime: '2024-05-15T10:00:00Z', IpAddress: '9.8.7.6', Location: 'Far Away Country' } ]
     },

    // --- Email Reports (Office 365 / Exchange Online) ---
    "Send as Emails": {
        command: "Search-UnifiedAuditLog -RecordType ExchangeMailbox -Operations SendAs -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | ConvertTo-Json # Needs parsing",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Send As:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Timestamp: '2024-05-18T10:00:00Z', User: 'delegate@contoso.com', Operation: 'SendAs', Target: 'boss@contoso.com', Details: 'Sent email as boss' } ]
    },
    "Send on Behalf Emails": {
        command: "Search-UnifiedAuditLog -RecordType ExchangeMailbox -Operations SendOnBehalf -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | ConvertTo-Json # Needs parsing",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Send On Behalf:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Timestamp: '2024-05-17T11:00:00Z', User: 'assistant@contoso.com', Operation: 'SendOnBehalf', Target: 'manager@contoso.com', Details: 'Sent email on behalf of manager' } ]
    },
    "Undelivered Mails": {
        command: "Get-MessageTrace -Status Failed -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) | Select Received,SenderAddress,RecipientAddress,Subject,Status,Error | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Undelivered Mails:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Received: '2024-05-20T09:00:00Z', SenderAddress: 'user@contoso.com', RecipientAddress: 'nonexistent@external.com', Subject: 'Important Update', Status: 'Failed', Error: 'Recipient not found' } ]
    },
    "Mails Sent by Delegates": {
        command: "Search-UnifiedAuditLog -RecordType ExchangeMailbox -Operations SendAs,SendOnBehalf -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | ConvertTo-Json # Needs parsing",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Delegate Sent Mails:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ /* Combine SendAs and SendOnBehalf data */ ]
    },
    "Email Traffic": {
        command: "Get-MailTrafficReport -StartDate (Get-Date).AddDays(-7).ToString('yyyy-MM-dd') -EndDate (Get-Date).ToString('yyyy-MM-dd') | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Email Traffic:", e); return []; } },
        visualizations: ['Table', 'BarChart'], // Example: Could show daily traffic in a bar chart
        defaultVisualization: 'Table',
        simulatedData: [ { Date: '2024-05-20', ReceivedCount: 5000, SentCount: 3000 } ]
    },
    "External Email Forwarding": {
        command: "Get-Mailbox -ResultSize Unlimited | Where-Object {$_.ForwardingSmtpAddress -ne $null -or $_.ForwardingAddress -ne $null} | Select DisplayName,PrimarySmtpAddress,ForwardingSmtpAddress,ForwardingAddress,DeliverToMailboxAndForward | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing External Forwarding:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { DisplayName: 'Forwarder Fred', PrimarySmtpAddress: 'fred@contoso.com', ForwardingSmtpAddress: 'fred@personal.com', DeliverToMailboxAndForward: true } ]
    },
    "Spam Detections": {
        command: "Get-MailTrafficProtectionReport -ReportType Spam -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Spam Detections:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Date: '2024-05-20', Direction: 'Inbound', MessageCount: 100, SpamCount: 20 } ]
    },
    "Top Phish Receivers": {
        command: "Get-MessageTrace -Status Delivered -Filter 'Direction eq 'Inbound' and MessageEvents.EventType eq 'Phish'' -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | Group-Object RecipientAddress | Sort-Object Count -Descending | Select Name,Count | ConvertTo-Json # Needs refinement",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Top Phish Receivers:", e); return []; } },
        visualizations: ['Table', 'BarChart'], // Could show top receivers in a bar chart
        defaultVisualization: 'Table',
        simulatedData: [ { Name: 'target@contoso.com', Count: 15 }, { Name: 'victim@contoso.com', Count: 10 } ]
    },
    "External Spoof Mails": {
        command: "Get-MessageTrace -Status Delivered -Filter 'Direction eq 'Inbound' and SenderAddress -notlike '*@contoso.com' and MessageEvents.EventType eq 'SpoofExternal'' -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | ConvertTo-Json # Needs refinement",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing External Spoof:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Received: '2024-05-19T08:00:00Z', SenderAddress: 'ceo@external-imposter.com', RecipientAddress: 'finance@contoso.com', Subject: 'Urgent Wire Transfer' } ]
    },
     "Top Malwares": {
        command: "Get-MailTrafficProtectionReport -ReportType Malware -Category MalwareName -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | Sort-Object MessageCount -Descending | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Top Malwares:", e); return []; } },
        visualizations: ['Table', 'BarChart'], // Could show top malware types in a bar chart
        defaultVisualization: 'Table',
        simulatedData: [ { MalwareName: 'Trojan.GenericKD', MessageCount: 50 }, { MalwareName: 'Worm.SomeWorm', MessageCount: 30 } ]
     },
     "Internal Mail flow": {
        command: "Get-MessageTrace -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) -Filter \"SenderAddress -like '*@contoso.com' and RecipientAddress -like '*@contoso.com'\" | Select Received,SenderAddress,RecipientAddress,Subject,Status | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Internal Mail Flow:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Received: '2024-05-20T11:00:00Z', SenderAddress: 'alice@contoso.com', RecipientAddress: 'bob@contoso.com', Subject: 'Project Update', Status: 'Delivered' } ]
     },
     "External Mail flow": {
        command: "Get-MessageTrace -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) -Filter \"(SenderAddress -like '*@contoso.com' and RecipientAddress -notlike '*@contoso.com') or (SenderAddress -notlike '*@contoso.com' and RecipientAddress -like '*@contoso.com')\" | Select Received,SenderAddress,RecipientAddress,Subject,Status | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing External Mail Flow:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Received: '2024-05-20T11:05:00Z', SenderAddress: 'alice@contoso.com', RecipientAddress: 'partner@external.com', Subject: 'Collaboration', Status: 'Delivered' } ]
     },
    "Users' Active Hours": {
        command: "Get-MailboxUsageReport -ReportType MailboxUsage -Date (Get-Date).AddDays(-1).ToString('yyyy-MM-dd') | ConvertTo-Json # Needs specific data points",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Active Hours:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { UserPrincipalName: 'alice@contoso.com', ActiveHoursEstimate: '9am-5pm' } ]
    },
    "Email Activities": {
        command: "Get-MailboxActivityReport -ReportType MailboxActivity -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Email Activities:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { UserPrincipalName: 'alice@contoso.com', SendCount: 50, ReceiveCount: 200, ReadCount: 180 } ]
    },
    "Domain-wise Summary": {
        command: "Get-MessageTrace -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) | Select SenderAddress,RecipientAddress | # Complex aggregation needed here... | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Domain Summary:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Domain: 'contoso.com', SentInternal: 1000, ReceivedInternal: 1200, SentExternal: 500, ReceivedExternal: 800 } ]
    },
    "Group Email Activities": {
        command: "Get-MailboxGroupActivityReport -ReportType GroupActivity -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Group Email Activities:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { GroupName: 'Sales Team', MessageCountReceived: 500, MessageCountSent: 50 } ]
    },
    "Org Email Traffic Stats": {
        command: "Get-MailTrafficReport -StartDate (Get-Date).AddDays(-7).ToString('yyyy-MM-dd') -EndDate (Get-Date).ToString('yyyy-MM-dd') | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Org Traffic:", e); return []; } },
        visualizations: ['Table', 'BarChart'],
        defaultVisualization: 'Table',
        simulatedData: [ { Date: '2024-05-20', ReceivedCount: 5000, SentCount: 3000 } ]
    },
    "User Email Traffic Stats": {
        command: "Get-MailboxActivityReport -ReportType MailboxActivity -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing User Traffic:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { UserPrincipalName: 'alice@contoso.com', SendCount: 50, ReceiveCount: 200 } ]
    },
    "Shared Mailbox Traffic Stats": {
        command: "# Get-MailboxStatistics for size, message trace for activity. Complex query.",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) as SharedMailboxTrafficStats[] : []; } catch (e) { console.error("Error parsing Shared Mailbox Stats:", e); return []; } },
        visualizations: ['Table', 'BarChart'], // Supports Table and Bar Chart
        defaultVisualization: 'BarChart',
        simulatedData: [
          { month: 'April 2025', sharedMailboxName: '24x7support', sharedMailboxUPN: '24x7support@o365droid.onmicrosoft.com', totalMailsSent: 5, totalMailsReceived: 10, externalMailsSent: 2, externalMailsReceived: 3 },
          { month: 'April 2025', sharedMailboxName: 'Amy Tucker', sharedMailboxUPN: 'amy.tucker@o365droid.onmicrosoft.com', totalMailsSent: 15, totalMailsReceived: 80, externalMailsSent: 5, externalMailsReceived: 10 },
          { month: 'April 2025', sharedMailboxName: 'Finance', sharedMailboxUPN: 'finance@o365droid.onmicrosoft.com', totalMailsSent: 2, totalMailsReceived: 190, externalMailsSent: 10, externalMailsReceived: 50 },
          { month: 'March 2025', sharedMailboxName: 'Finance', sharedMailboxUPN: 'finance@o365droid.onmicrosoft.com', totalMailsSent: 5, totalMailsReceived: 450, externalMailsSent: 19, externalMailsReceived: 100 },
        ]
    },
    "Email Traffic by 30 min": {
        command: "Get-MailTrafficReport -Interval HalfHour -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) | ConvertTo-Json # Needs correct cmdlet/params",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing 30min Traffic:", e); return []; } },
        visualizations: ['Table'], // Could add LineChart if implemented
        defaultVisualization: 'Table',
        simulatedData: [ { TimeSlot: '2024-05-20T10:00:00', ReceivedCount: 50, SentCount: 20 } ]
    },
    "Hourly Email Traffic": {
        command: "Get-MailTrafficReport -Interval Hourly -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) | ConvertTo-Json # Needs correct cmdlet/params",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Hourly Traffic:", e); return []; } },
        visualizations: ['Table'], // Could add LineChart if implemented
        defaultVisualization: 'Table',
        simulatedData: [ { Hour: '2024-05-20T10:00:00', ReceivedCount: 100, SentCount: 40 } ]
    },
    "Daily Email Traffic": {
        command: "Get-MailTrafficReport -StartDate (Get-Date).AddDays(-7).ToString('yyyy-MM-dd') -EndDate (Get-Date).ToString('yyyy-MM-dd') | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Daily Traffic:", e); return []; } },
        visualizations: ['Table', 'BarChart'], // Example: Bar chart showing daily counts
        defaultVisualization: 'Table',
        simulatedData: [ { Date: '2024-05-20', ReceivedCount: 5000, SentCount: 3000 } ]
    },
    "Monthly Email Traffic": {
        command: "Get-MailTrafficReport -Interval Monthly -StartDate (Get-Date).AddMonths(-6) -EndDate (Get-Date) | ConvertTo-Json # Needs correct cmdlet/params",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Monthly Traffic:", e); return []; } },
        visualizations: ['Table', 'BarChart'], // Example: Bar chart showing monthly counts
        defaultVisualization: 'Table',
        simulatedData: [ { Month: '2024-05', ReceivedCount: 100000, SentCount: 60000 } ]
    },
     "Groups Mail Traffic Stats": {
        command: "Get-MailboxGroupActivityReport -ReportType GroupActivity -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Group Traffic:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { GroupName: 'Sales Team', MessageCountReceived: 500, MessageCountSent: 50 } ]
    },
     "Total Mails By Hour/Day": {
        command: "# Aggregate from Get-MailTrafficReport -Interval Hourly/Daily",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: () => [],
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: []
    },
     "Organizations' Total Mails": {
        command: "Get-MailTrafficReport -StartDate (Get-Date).AddDays(-7).ToString('yyyy-MM-dd') -EndDate (Get-Date).ToString('yyyy-MM-dd') | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Org Total Mails:", e); return []; } },
        visualizations: ['Table'], // Could be a simple display component
        defaultVisualization: 'Table',
        simulatedData: [ { TotalReceived: 35000, TotalSent: 21000 } ] // Aggregated
    },
     "User Total Mails": {
        command: "Get-MailboxActivityReport -ReportType MailboxActivity -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing User Total Mails:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { UserPrincipalName: 'alice@contoso.com', TotalSent: 50, TotalReceived: 200 } ]
    },
     "Shared Mailbox Total Mails": {
        command: "Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Get-MailboxStatistics | Select DisplayName, TotalItemCount, TotalDeletedItemSize | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Shared Mailbox Total:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { DisplayName: 'Finance', TotalItemCount: 10000 } ]
    },
     "Groups Total Mails": {
        command: "Get-MailboxGroupActivityReport -ReportType GroupActivity -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Group Total Mails:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { GroupName: 'Sales Team', TotalMessages: 550 } ] // Aggregated
    },
     "Peak Period Analysis": {
        command: "# Analyze Get-MailTrafficReport -Interval Hourly/HalfHour",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: () => [],
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { PeakHour: '10:00 AM', PeakVolume: 150 } ]
    },

    // --- Teams Reports ---
    "Public Teams": {
        command: "Get-MgTeam -Filter \"visibility eq 'Public'\" -Select Id,DisplayName,Description,CreatedDateTime | ConvertTo-Json",
        service: executeAzurePowerShellCommand, // Teams commands are often in MgGraph now
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Public Teams:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { id: 'team1', displayName: 'Public Forum', description: 'Open discussion', CreatedDateTime: '2024-02-01T10:00:00Z' } ]
    },
    "Private Teams": {
        command: "Get-MgTeam -Filter \"visibility eq 'Private'\" -Select Id,DisplayName,Description,CreatedDateTime | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Private Teams:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { id: 'team2', displayName: 'Project X', description: 'Confidential project', CreatedDateTime: '2024-03-01T11:00:00Z' } ]
    },
    "Teams Membership": {
        command: "Get-MgTeamMember -TeamId 'TEAM_ID_PLACEHOLDER' -All | Select-Object DisplayName, Roles, UserId | ConvertTo-Json",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Teams Membership:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { DisplayName: 'Alice Smith', Roles: ['owner'], UserId: 'user1' } ]
    },
    "Private & Shared Channels": {
        command: "Get-MgTeam -All | ForEach-Object { Get-MgTeamChannel -TeamId $_.Id -Filter \"membershipType ne 'standard'\" } | Select DisplayName,MembershipType,@{N='Team';E={$_.TeamId}} | ConvertTo-Json # Inefficient",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Channels:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { DisplayName: 'Secret Subproject', MembershipType: 'private', Team: 'team2' } ]
    },
    "Private Channels": {
        command: "Get-MgTeam -All | ForEach-Object { Get-MgTeamChannel -TeamId $_.Id -Filter \"membershipType eq 'private'\" } | Select DisplayName,@{N='Team';E={$_.TeamId}} | ConvertTo-Json # Inefficient",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Private Channels:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { DisplayName: 'Secret Subproject', Team: 'team2' } ]
    },
    "Team Setting Changes": {
        command: "Search-UnifiedAuditLog -RecordType MicrosoftTeams -Operations 'Update team','Add channel','Remove channel' -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | ConvertTo-Json # Needs parsing",
        service: executeOffice365PowerShellCommand, // Audit log might be EXO side
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Team Setting Changes:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Timestamp: '2024-05-19T15:00:00Z', User: 'admin@contoso.com', Operation: 'Update team', Target: 'Project X', Details: 'Changed team description' } ]
    },
    "Inactive Users": {
        command: "Get-MSTeamsUserActivityReport -Date (Get-Date).AddDays(-30).ToString('yyyy-MM-dd') | Where-Object {$_.LastActivityDate -lt (Get-Date).AddDays(-30)} | ConvertTo-Json # Needs specific report cmdlet",
        service: executeOffice365PowerShellCommand, // Teams reports might be EXO side
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Inactive Teams Users:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { UserPrincipalName: 'inactive@contoso.com', LastActivityDate: '2024-03-15T00:00:00Z' } ]
    },
    "Teams Device Usage": {
        command: "Get-MSTeamsDeviceUsageReport -Period D7 | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Teams Device Usage:", e); return []; } },
        visualizations: ['Table', 'PieChart'], // Example: Pie chart showing device breakdown
        defaultVisualization: 'Table',
        simulatedData: [ { ReportRefreshDate: '2024-05-20', Windows: 100, Mac: 20, Web: 50, Mobile: 80 } ]
    },
    "External File Sharing": {
        command: "Search-UnifiedAuditLog -RecordType SharePointFileOperation -Operations FileAccessed,FileShared -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | Where-Object {$_.UserId -like '*#EXT#*' -or $_.SiteUrl -like '*sharing*'} | ConvertTo-Json # Very simplified",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing External File Sharing:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Timestamp: '2024-05-18T16:00:00Z', User: 'guest#EXT#@contoso.com', Operation: 'FileAccessed', Target: 'Shared Document.docx', Site: 'Team Site' } ]
    },
    "Teams Add-ons": {
        command: "Get-MgTeam -All | ForEach-Object { Get-MgTeamInstalledApp -TeamId $_.Id } | Select @{N='TeamName';E={(Get-MgTeam -TeamId $_.TeamId).DisplayName}},@{N='AppName';E={$_.TeamsAppDefinition.DisplayName}} | ConvertTo-Json # Inefficient",
        service: executeAzurePowerShellCommand,
        apiTarget: 'azure',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Teams Add-ons:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { TeamName: 'Project X', AppName: 'Planner' }, { TeamName: 'Public Forum', AppName: 'Forms' } ]
    },
    "Private Channel Membership Changes": {
        command: "Search-UnifiedAuditLog -RecordType MicrosoftTeams -Operations 'Member added','Member removed' -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | Where-Object {$_.ObjectId -like '*-channel' -and $_.ObjectId -notlike 'General'} | ConvertTo-Json # Needs parsing",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Private Channel Changes:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Timestamp: '2024-05-17T09:00:00Z', User: 'admin@contoso.com', Operation: 'Member added', Target: 'Secret Subproject Channel', Details: 'Added Bob Johnson' } ]
    },
    "Ownership Promotions and Demotions": {
        command: "Search-UnifiedAuditLog -RecordType MicrosoftTeams -Operations 'Member role updated' -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | ConvertTo-Json # Needs parsing",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Ownership Changes:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Timestamp: '2024-05-16T14:00:00Z', User: 'alice@contoso.com', Operation: 'Member role updated', Target: 'Project X', Details: 'Promoted Bob Johnson to Owner' } ]
    },

    // --- Exchange Online Reports ---
    "Mailbox Usage": {
        command: "Get-MailboxStatistics -ResultSize Unlimited | Select DisplayName,TotalItemCount,ItemCount,TotalItemSize,StorageLimitStatus | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Mailbox Usage:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { DisplayName: 'Alice Smith', TotalItemCount: 5000, TotalItemSize: '1.5 GB', StorageLimitStatus: 'BelowLimit' } ]
    },
    "Inactive Mailboxes": {
        command: "Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics | Where-Object {$_.LastLogonTime -lt (Get-Date).AddDays(-90)} | Select DisplayName,LastLogonTime | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Inactive Mailboxes:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { DisplayName: 'Old User', LastLogonTime: '2023-12-01T10:00:00Z' } ]
    },
    "Active Out of Office Settings": {
        command: "Get-Mailbox -ResultSize Unlimited | Get-MailboxAutoReplyConfiguration | Where-Object {$_.AutoReplyState -eq 'Scheduled' -or $_.AutoReplyState -eq 'Enabled'} | Select Identity,AutoReplyState,StartTime,EndTime,InternalMessage,ExternalMessage | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing OOF Settings:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Identity: 'bob@contoso.com', AutoReplyState: 'Enabled', InternalMessage: 'Out of office until Monday.' } ]
    },
    "Mailbox Permissions": {
        command: "Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | Where-Object {$_.User -ne 'NT AUTHORITY\\SELF' -and $_.IsInherited -eq $false} | Select Identity,User,AccessRights | ConvertTo-Json # Inefficient",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Mailbox Permissions:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Identity: 'finance@contoso.com', User: 'admin@contoso.com', AccessRights: 'FullAccess' } ]
    },
    "Audit Disabled Mailboxes": {
        command: "Get-Mailbox -ResultSize Unlimited | Where-Object {$_.AuditEnabled -eq $false} | Select DisplayName,PrimarySmtpAddress,AuditEnabled | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Audit Disabled:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { DisplayName: 'Legacy Mailbox', PrimarySmtpAddress: 'legacy@contoso.com', AuditEnabled: false } ]
    },
    "Mailbox Hold": {
        command: "Get-Mailbox -ResultSize Unlimited | Where-Object {$_.LitigationHoldEnabled -eq $true -or $_.InPlaceHolds -ne $null} | Select DisplayName,LitigationHoldEnabled,InPlaceHolds | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Mailbox Hold:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { DisplayName: 'Legal Hold User', LitigationHoldEnabled: true, InPlaceHolds: null } ]
    },
    "Mailbox Non-owner Access": {
        command: "Search-UnifiedAuditLog -RecordType ExchangeAdmin,ExchangeMailbox -Operations 'Add-MailboxPermission','Remove-MailboxPermission','Set-Mailbox','New-MailboxAccess' -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) | ConvertTo-Json # Needs detailed parsing",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Mailbox Non-Owner Access:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Timestamp: '2024-05-18T14:00:00Z', User: 'admin@contoso.com', Operation: 'Add-MailboxPermission', Target: 'finance@contoso.com', Details: 'User admin added FullAccess for support@contoso.com' } ]
    },
    "Mobile Device Configurations": {
        command: "Get-MobileDevice -ResultSize Unlimited | Get-MobileDeviceStatistics | Select DeviceModel,DeviceOS,LastSuccessSync | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Mobile Devices:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { DeviceModel: 'iPhone 15', DeviceOS: 'iOS 17', LastSuccessSync: '2024-05-20T12:00:00Z' } ]
    },
    "Public Folders": {
        command: "Get-PublicFolder -Recurse -ResultSize Unlimited | Get-PublicFolderStatistics | Select Name,ItemCount,TotalItemSize | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Public Folders:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Name: 'Company Announcements', ItemCount: 50, TotalItemSize: '100 MB' } ]
    },
    "Role Assignments": {
        command: "Get-ManagementRoleAssignment -RoleAssigneeType User -ResultSize Unlimited | Select Role,User,RoleAssigneeName | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Role Assignments:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { Role: 'Organization Management', User: 'admin@contoso.com', RoleAssigneeName: 'Admin User' } ]
    },
    "Mail Flow": {
        command: "Get-MessageTrace -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) | Select Received,SenderAddress,RecipientAddress,Subject,Status | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Mail Flow:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ /* Combined internal/external flow data */ ]
    },
    "Exchange Contacts": {
        command: "Get-MailContact -ResultSize Unlimited | Select DisplayName,ExternalEmailAddress,PrimarySmtpAddress | ConvertTo-Json",
        service: executeOffice365PowerShellCommand,
        apiTarget: 'office365',
        parser: (output) => { try { return output ? JSON.parse(output) : []; } catch (e) { console.error("Error parsing Exchange Contacts:", e); return []; } },
        visualizations: ['Table'],
        defaultVisualization: 'Table',
        simulatedData: [ { DisplayName: 'External Partner', ExternalEmailAddress: 'partner@example.com' } ]
    },
};


export function ReportDisplay({ selectedReport }: ReportDisplayProps) {
  const [reportData, setReportData] = useState<any>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [reportCommand, setReportCommand] = useState<string | null>(null);
  const [selectedVisualization, setSelectedVisualization] = useState<VisualizationType | null>(null);

  // Reset visualization when the report changes
  useEffect(() => {
    if (selectedReport) {
      const config = reportConfig[selectedReport];
      setSelectedVisualization(config?.defaultVisualization ?? null);
    } else {
      setSelectedVisualization(null);
    }
  }, [selectedReport]);


  const fetchReportData = useCallback(async (report: ReportCategory) => {
    const config = reportConfig[report];
    if (!config) {
      console.error(`No configuration found for report: ${report}`);
      setError(`Configuration missing for report: ${report}.`);
      setIsLoading(false);
      setReportData(null);
      setReportCommand(null);
      return;
    }

    setIsLoading(true);
    setError(null);
    setReportData(null);
    setReportCommand(config.command); // Store the command being executed
    // Set the default visualization for this report initially
    setSelectedVisualization(config.defaultVisualization);


    // Introduce a small delay for simulation to show loading state
    await new Promise(resolve => setTimeout(resolve, 300));

    try {
       // Handle placeholder commands starting with '#' - Treat as not implemented
       if (config.command.startsWith("#")) {
          // Use specific error message for placeholder commands
          setError(`Report not implemented or requires complex setup: ${report}`);
          setReportData([]); // Ensure empty data
          setIsLoading(false);
          return; // Stop further execution for this report
       }

       // Use simulated data if available in config for stubbed backend behavior
       let resultOutput: string | null = null;
       let apiError: string | undefined;

      // Always use simulated data if present, simulating the backend
      if (config.simulatedData) {
          console.log(`Using simulated data for: ${report}`);
          resultOutput = JSON.stringify(config.simulatedData);
      } else {
          // If no simulated data, call the API (which might have its own simulation)
          console.log(`Executing command for ${report}: ${config.command}`);
          const result = await config.service(config.command);
          apiError = result.error;
          resultOutput = result.output;
      }


       // Check for API errors FIRST (e.g., simulation not implemented in the API itself)
      if (apiError) {
          // If the API returns a simulation error, treat it as an error
          if (apiError.includes('simulation not implemented')) {
              setError(`Report simulation not implemented: ${report}. Error: ${apiError}`);
              setReportData([]);
              setIsLoading(false);
              return;
          }
          // For other API errors, set the error state
          setError(`API Error for ${report}: ${apiError}`);
          setReportData([]);
          setIsLoading(false);
          return;
      }


      const parsedData = config.parser(resultOutput);

      if (!parsedData || (Array.isArray(parsedData) && parsedData.length === 0 && resultOutput && resultOutput !== '[]')) {
          console.warn(`Parser for "${report}" returned empty data. Original output:`, resultOutput);
          setReportData([]);
      } else {
           setReportData(parsedData);
      }


    } catch (err) {
      // Catch errors specifically thrown (like the 'Report not implemented' error)
      console.error(`Error processing report ${report}:`, err);
       const message = err instanceof Error ? err.message : 'An unknown error occurred during processing';
       setError(`Failed to load or parse data for ${report}. Error: ${message}`);
       setReportData([]); // Ensure data is empty on error
    } finally {
      setIsLoading(false);
    }
  }, []); // Empty dependency array

  useEffect(() => {
    if (selectedReport) {
      fetchReportData(selectedReport);
    } else {
      // Clear state when no report is selected
      setReportData(null);
      setError(null);
      setIsLoading(false);
      setReportCommand(null);
      setSelectedVisualization(null);
    }
  }, [selectedReport, fetchReportData]);


  const renderReportContent = () => {
    if (isLoading) {
      return (
        <div className="flex flex-col items-center justify-center h-64">
          <Loader2 className="h-8 w-8 animate-spin text-primary" />
          <span className="mt-2 text-muted-foreground">Loading report data...</span>
          {reportCommand && !reportCommand.startsWith("#") && (
            <code className="mt-4 text-xs p-2 bg-muted rounded max-w-full overflow-x-auto">
              Executing: {reportCommand}
            </code>
          )}
        </div>
      );
    }

    if (error) {
      // Special handling for "Not Implemented" or simulation errors based on the error message
      if (error.includes('not implemented') || error.includes('requires complex setup')) {
        return (
          <Alert className="mt-4">
            <Info className="h-4 w-4" />
            <AlertTitle>Report Not Implemented or Requires Setup</AlertTitle>
            <AlertDescription>
              {selectedReport ? `The reporting functionality for '${selectedReport}' is either not yet fully implemented, requires specific configuration (like IDs or date ranges), or relies on simulated data.` : 'Report functionality not implemented.'}
              <br/> <span className="text-muted-foreground text-xs">({error})</span>
            </AlertDescription>
             {reportCommand && ( // Show command even if placeholder
               <code className="mt-2 text-xs p-1 bg-muted/50 rounded block max-w-full overflow-x-auto">
                 Command: {reportCommand}
               </code>
             )}
          </Alert>
        );
      }
      // General error display
      return (
        <Alert variant="destructive" className="mt-4">
          <AlertTriangle className="h-4 w-4" />
          <AlertTitle>Error Loading Report</AlertTitle>
          <AlertDescription>{error}</AlertDescription>
           {reportCommand && !reportCommand.startsWith("#") && ( // Show command only if it's not the placeholder
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
     if (!config) return null; // Should not happen if selectedReport is valid

     // Get the component based on the selected visualization type
     const currentVisualizationType = selectedVisualization ?? config.defaultVisualization;
     const visualizationInfo = visualizationComponents[currentVisualizationType];

     if (!visualizationInfo) {
        return <Alert variant="destructive"><AlertTitle>Error</AlertTitle><AlertDescription>Invalid visualization type selected.</AlertDescription></Alert>;
     }

     const ReportComponent = visualizationInfo.component;


     if (!reportData || (Array.isArray(reportData) && reportData.length === 0)) {
       return (
         <Alert className="mt-4">
            <Info className="h-4 w-4" />
           <AlertTitle>No Data Available</AlertTitle>
           <AlertDescription>There is no data available for the selected report: {selectedReport}. This might be because the command returned no results, the data is genuinely empty, or there was an issue parsing the response.</AlertDescription>
            {reportCommand && !reportCommand.startsWith("#") && (
             <code className="mt-2 text-xs p-1 bg-muted/50 rounded block max-w-full overflow-x-auto">
               Command: {reportCommand}
             </code>
           )}
         </Alert>
       );
     }

     return <ReportComponent data={reportData} caption={selectedReport} />;
  };

  // Determine available visualizations for the current report
  const availableVisualizations = selectedReport ? reportConfig[selectedReport]?.visualizations : [];
  const currentVisualizationType = selectedVisualization ?? (selectedReport ? reportConfig[selectedReport]?.defaultVisualization : null);

  return (
    <main className="flex-1 p-4 md:p-6 overflow-auto bg-background">
      <Card className="w-full shadow-lg h-full flex flex-col border-border/50 bg-card">
        <CardHeader className="border-b border-border/50 flex flex-row items-center justify-between">
          <div>
            <CardTitle className="text-xl text-primary">{selectedReport || 'Select a Report'}</CardTitle>
            <CardDescription className="text-muted-foreground">
              {selectedReport ? `Viewing details for ${selectedReport}` : 'Choose a report from the sidebar.'}
            </CardDescription>
          </div>
          {/* Visualization Dropdown */}
          {selectedReport && availableVisualizations && availableVisualizations.length > 1 && !isLoading && !error && reportData && (
            <div className="flex items-center gap-2">
              <Label htmlFor="visualization-select" className="text-sm font-medium text-muted-foreground">View As:</Label>
              <Select
                value={currentVisualizationType ?? undefined}
                onValueChange={(value) => setSelectedVisualization(value as VisualizationType)}
              >
                <SelectTrigger id="visualization-select" className="w-[180px]">
                  <SelectValue placeholder="Select visualization" />
                </SelectTrigger>
                <SelectContent>
                  {availableVisualizations.map((visType) => {
                    const visInfo = visualizationComponents[visType];
                    if (!visInfo) return null;
                    const IconComponent = visInfo.icon;
                    return (
                      <SelectItem key={visType} value={visType}>
                        <div className="flex items-center gap-2">
                           <IconComponent className="h-4 w-4 text-muted-foreground" />
                           {visType}
                         </div>
                      </SelectItem>
                    );
                  })}
                </SelectContent>
              </Select>
            </div>
          )}
        </CardHeader>
        <CardContent className="flex-1 overflow-auto pt-4">
           {renderReportContent()}
        </CardContent>
      </Card>
     </main>
  );
}
    
