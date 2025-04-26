
/**
 * Defines the categories of reports available in the reporting section.
 * These should match the labels/identifiers used in the sidebar and display logic.
 */
export type ReportCategory =
  | "Email Activities"
  | "Domain-wise Summary"
  | "Group Email Activities"
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
  | "Peak Period Analysis";
  // Add other specific report names here as needed


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

// Add interfaces for other report data structures here
// e.g.,
// export interface UserEmailTrafficStats { ... }
// export interface DomainSummary { ... }

// You might also want a generic type or union type if fetching logic is centralized
// export type ReportData = SharedMailboxTrafficStats | UserEmailTrafficStats | DomainSummary | ...;
