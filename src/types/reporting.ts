

/**
 * Defines the categories of reports available in the reporting section.
 * These should match the labels/identifiers used in the sidebar and display logic.
 */
export type ReportCategory =
  // Azure AD Reports
  | "License Utilization"
  | "License Expiry"
  | "Users & Groups"
  | "Unlicensed User"
  | "User Managers & Direct Reports"
  | "Group Members"
  | "Group Owners"
  | "Group-based Licensing"
  | "Login Activities" // Azure AD Login Activities
  | "Password Changes"
  | "MFA Disabled Users"
  | "Device Registrations"

  // Security Reports
  | "MFA Enforced Users"
  | "MFA Non-Activated Users"
  | "Password Expiry"
  | "Users with Weak Passwords"
  | "External Users"
  | "Guest Users"
  | "Risky Login Attempts"
  | "External User License Assignments"
  | "Secure Score"
  | "eDiscovery"
  | "Non-Owner Mailbox Access"
  | "Data Loss Prevention (DLP)"
  | "Advanced Threat Protection"

  // Sign-in Analysis Reports
  | "User Sign-in Location"
  | "Last Log-on Summary"
  | "External User Sign-ins"
  | "Guest Sign-ins"
  | "Non-Compliant Device Sign-ins"
  | "Unmanaged Device Sign-ins"
  | "MFA failed Sign-ins"
  | "2FA Authentication Methods"
  | "Conditional Access failures"
  | "Conditional Access Policies"
  | "Sign-ins Risk Level"
  | "Unlikely Travel Risky Sign-ins"
  | "Anonymous IP Sign-ins"
  | "Compromised Risky Sign-ins"

  // Email Reports (Considered part of Exchange Online or separate?)
  | "Send as Emails"
  | "Send on Behalf Emails"
  | "Undelivered Mails"
  | "Mails Sent by Delegates"
  | "Email Traffic"
  | "External Email Forwarding"
  | "Spam Detections"
  | "Top Phish Receivers"
  | "External Spoof Mails"
  | "Top Malwares"
  | "Internal Mail flow"
  | "External Mail flow"
  | "Users' Active Hours"
  // Existing Email Reports (Can be merged or kept separate)
  | "Email Activities" // Generic, maybe remove or scope?
  | "Domain-wise Summary"
  | "Group Email Activities" // Generic, maybe remove or scope?
  | "Org Email Traffic Stats"
  | "User Email Traffic Stats"
  | "Shared Mailbox Traffic Stats"
  | "Email Traffic by 30 min"
  | "Hourly Email Traffic"
  | "Daily Email Traffic"
  | "Monthly Email Traffic"
  | "Groups Mail Traffic Stats"
  | "Total Mails By Hour/Day"
  | "Organizations' Total Mails"
  | "User Total Mails"
  | "Shared Mailbox Total Mails"
  | "Groups Total Mails"
  | "Peak Period Analysis"


  // Teams Reports
  | "Public Teams"
  | "Private Teams"
  | "Teams Membership"
  | "Private & Shared Channels"
  | "Login Activities" // Teams Login Activities (potentially same as Azure AD)
  | "Private Channels"
  | "Team Setting Changes"
  | "Inactive Users" // Teams Inactive Users
  | "Teams Device Usage"
  | "External File Sharing"
  | "Teams Add-ons"
  | "Private Channel Membership Changes"
  | "Ownership Promotions and Demotions"

  // Exchange Online Reports
  | "Mailbox Usage"
  | "Inactive Mailboxes"
  | "Active Out of Office Settings"
  | "Mailbox Permissions"
  | "Audit Disabled Mailboxes"
  | "Mailbox Hold"
  // | "Non-owner Access" // Duplicate in Security? Clarify/Choose one. Renaming to Mailbox Non-owner Access
  | "Mailbox Non-owner Access"
  | "Mobile Device Configurations"
  | "Public Folders"
  | "Role Assignments"
  | "Mail Flow" // Exchange Mail Flow (potentially distinct from general email flow)
  | "Exchange Contacts";


/**
 * Represents the structure for Shared Mailbox Traffic Statistics data.
 * Used for the 'Shared Mailbox Traffic Stats' report table.
 */
export interface SharedMailboxTrafficStats {
  month: string;
  sharedMailboxName: string;
  sharedMailboxUPN: string;
  totalMailsSent: number | null;
  totalMailsReceived: number | null;
  externalMailsSent: number | null;
  externalMailsReceived: number | null;
}

/**
 * Represents the structure for License Utilization data from `Get-MgSubscribedSku`.
 */
export interface LicenseUtilizationData {
  SkuId: string;
  SkuPartNumber: string;
  ConsumedUnits: number;
  TotalUnits: number; // Combined from PrepaidUnits sub-properties
}

/**
 * Represents combined information for Azure AD Users and Groups.
 * Needs refinement based on how the PowerShell command combines/returns data.
 */
export interface AzureADUserGroupInfo {
  type: 'User' | 'Group'; // To distinguish between users and groups in a combined list
  id: string;
  displayName: string;
  // User specific
  userPrincipalName?: string;
  license?: string | string[]; // Might be an array or specific license string
  // Group specific
  description?: string | null;
  groupType?: string | string[];
}


// --- Report Category Grouping ---
// Optional: Define groups for easier sidebar management if needed elsewhere
export const ReportGroups = {
  AZURE_AD: [
    "License Utilization", "License Expiry", "Users & Groups", "Unlicensed User",
    "User Managers & Direct Reports", "Group Members", "Group Owners",
    "Group-based Licensing", "Login Activities", "Password Changes",
    "MFA Disabled Users", "Device Registrations"
  ],
  SECURITY: [
    "MFA Enforced Users", "MFA Non-Activated Users", "Password Expiry",
    "Users with Weak Passwords", "External Users", "Guest Users",
    "Risky Login Attempts", "External User License Assignments", "Secure Score",
    "eDiscovery", "Non-Owner Mailbox Access", "Data Loss Prevention (DLP)",
    "Advanced Threat Protection"
  ],
  SIGN_IN_ANALYSIS: [
    "User Sign-in Location", "Last Log-on Summary", "External User Sign-ins",
    "Guest Sign-ins", "Non-Compliant Device Sign-ins", "Unmanaged Device Sign-ins",
    "MFA failed Sign-ins", "2FA Authentication Methods", "Conditional Access failures",
    "Conditional Access Policies", "Sign-ins Risk Level", "Unlikely Travel Risky Sign-ins",
    "Anonymous IP Sign-ins", "Compromised Risky Sign-ins"
  ],
  EMAIL: [
    "Send as Emails", "Send on Behalf Emails", "Undelivered Mails",
    "Mails Sent by Delegates", "Email Traffic", "External Email Forwarding",
    "Spam Detections", "Top Phish Receivers", "External Spoof Mails", "Top Malwares",
    "Internal Mail flow", "External Mail flow", "Users' Active Hours",
    // Including previous email reports under this category for consolidation
    "Email Activities", "Domain-wise Summary", "Group Email Activities",
    "Org Email Traffic Stats", "User Email Traffic Stats", "Shared Mailbox Traffic Stats",
    "Email Traffic by 30 min", "Hourly Email Traffic", "Daily Email Traffic",
    "Monthly Email Traffic", "Groups Mail Traffic Stats", "Total Mails By Hour/Day",
    "Organizations' Total Mails", "User Total Mails", "Shared Mailbox Total Mails",
    "Groups Total Mails", "Peak Period Analysis"
  ],
  TEAMS: [
    "Public Teams", "Private Teams", "Teams Membership", "Private & Shared Channels",
    "Login Activities", // Specify context if needed: "Teams Login Activities"
    "Private Channels", "Team Setting Changes", "Inactive Users", // Specify context: "Teams Inactive Users"
    "Teams Device Usage", "External File Sharing", "Teams Add-ons",
    "Private Channel Membership Changes", "Ownership Promotions and Demotions"
  ],
  EXCHANGE_ONLINE: [
    "Mailbox Usage", "Inactive Mailboxes", "Active Out of Office Settings",
    "Mailbox Permissions", "Audit Disabled Mailboxes", "Mailbox Hold",
    "Mailbox Non-owner Access", "Mobile Device Configurations", "Public Folders",
    "Role Assignments", "Mail Flow", // Specify context: "Exchange Mail Flow"
    "Exchange Contacts"
  ],
};

// Add interfaces for other report data structures here as commands are implemented
// e.g.,
// export interface AzureADGroupMember { ... }
// export interface MfaStatusReport { ... }
// export interface TeamsActivityLog { ... }

// Generic type or union type for fetched data if needed for advanced components
// export type ReportData =
//    SharedMailboxTrafficStats |
//    LicenseUtilizationData |
//    AzureADUserGroupInfo |
//    // ... other report types
//    null;
