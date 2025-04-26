
'use client';

import React from 'react';
import {
  Sidebar,
  SidebarContent,
  SidebarHeader,
  SidebarFooter,
  SidebarTrigger,
  SidebarMenu,
  SidebarMenuItem,
  SidebarMenuButton,
  SidebarMenuSub,
  SidebarMenuSubItem,
  SidebarMenuSubButton,
  useSidebar,
} from '@/components/ui/sidebar';
import {
  Mail,
  LineChart,
  Users,
  Inbox,
  Settings,
  AreaChart,
  BarChartHorizontal,
  CalendarClock,
  Clock,
  CalendarDays,
  Calendar,
  Group,
  User,
  Box, // Placeholder, consider Package
  Building2,
  ShieldCheck, // For Security
  Fingerprint, // For Sign-in Analysis
  Database, // For Exchange Online / Azure AD
  Activity, // For Login Activities
  Lock, // Password related
  Users2, // Teams icon
  FileText, // Reports / Generic
  Key, // Authentication Methods / Password
  AlertTriangle, // Risky / Failures
  Globe, // Domain / Location
  Network, // Group-based Licensing / Connections
  ClipboardList, // Role Assignments / Policies
  Folder, // Public Folders
  Smartphone, // Mobile Devices
  Briefcase, // License related
  ShieldAlert, // DLP / ATP
  Search, // eDiscovery
  UserCheck, // MFA Enforced
  UserX, // MFA Disabled/Non-Activated
  MailWarning, // Spam/Phish/Spoof
  Bug, // Malware
  Plane, // Travel Sign-ins
  Waypoints, // Mail Flow
  Contact, // Exchange Contacts
  Presentation, // Promotions/Demotions
  Share2, // External Sharing
  Plug, // Add-ons
  Tv, // Device Sign-ins / Registrations
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import type { ReportCategory, ReportGroups } from '@/types/reporting';
import { cn } from '@/lib/utils';

interface ReportingSidebarProps {
  selectedReport: ReportCategory | null;
  onSelectReport: (report: ReportCategory | null) => void;
}

// Helper to create menu items/buttons
const MenuItem = ({
  report,
  icon: Icon,
  label,
  selectedReport,
  onSelectReport,
  isSubItem = false,
}: {
  report: ReportCategory;
  icon?: React.ElementType;
  label: string;
  selectedReport: ReportCategory | null;
  onSelectReport: (report: ReportCategory) => void;
  isSubItem?: boolean;
}) => {
  const Comp = isSubItem ? SidebarMenuSubButton : SidebarMenuButton;
  const { setOpenMobile } = useSidebar(); // To close mobile sidebar on selection

  return (
    <Comp
      onClick={() => {
        onSelectReport(report);
        setOpenMobile(false); // Close mobile sidebar on selection
      }}
      isActive={selectedReport === report}
      className={cn(
        'justify-start text-left w-full', // Ensure text wraps if needed
        isSubItem && 'h-auto py-1.5 text-xs' // Adjust sub-item styling
      )}
      asChild={isSubItem} // Use anchor tag for sub-buttons if needed for routing later
    >
      {isSubItem ? (
        <a href="#"> {/* Placeholder href */}
          {Icon && <Icon className="size-3.5 mr-1.5 shrink-0" />}
          <span className="whitespace-normal">{label}</span>
        </a>
      ) : (
        <>
          {Icon && <Icon />}
          <span className="whitespace-normal">{label}</span>
        </>
      )}
    </Comp>
  );
};


export function ReportingSidebar({ selectedReport, onSelectReport }: ReportingSidebarProps) {
   const { isMobile } = useSidebar();

   const renderMenuItems = (reports: ReportCategory[], categoryIcon?: React.ElementType) => {
    return reports.map((report) => (
       <SidebarMenuSubItem key={report}>
        {/* Add specific icons for each report if desired, otherwise pass category icon or null */}
        <MenuItem
          report={report}
          // icon={categoryIcon} // Example: Use category icon for all sub-items
          label={report}
          selectedReport={selectedReport}
          onSelectReport={onSelectReport}
          isSubItem
        />
       </SidebarMenuSubItem>
    ));
   };

  return (
    <Sidebar collapsible="icon">
       <SidebarHeader className="items-center gap-2">
        <SidebarTrigger className={cn('md:hidden', isMobile && 'block')} />
        <span className="text-lg font-semibold group-data-[collapsible=icon]:hidden">
          Reports
        </span>
      </SidebarHeader>
      <SidebarContent className="p-2">
        <SidebarMenu>

          {/* Azure AD Section */}
          <SidebarMenuItem>
            <SidebarMenuButton icon={Database}>Azure AD</SidebarMenuButton>
            <SidebarMenuSub>
              {renderMenuItems([
                "License Utilization", "License Expiry", "Users & Groups", "Unlicensed User",
                "User Managers & Direct Reports", "Group Members", "Group Owners",
                "Group-based Licensing", "Login Activities", "Password Changes",
                "MFA Disabled Users", "Device Registrations"
              ], Database)}
            </SidebarMenuSub>
          </SidebarMenuItem>

          {/* Security Section */}
           <SidebarMenuItem>
            <SidebarMenuButton icon={ShieldCheck}>Security</SidebarMenuButton>
            <SidebarMenuSub>
              {renderMenuItems([
                "MFA Enforced Users", "MFA Non-Activated Users", "Password Expiry",
                "Users with Weak Passwords", "External Users", "Guest Users",
                "Risky Login Attempts", "External User License Assignments", "Secure Score",
                "eDiscovery", "Non-Owner Mailbox Access", "Data Loss Prevention (DLP)",
                "Advanced Threat Protection"
              ], ShieldCheck)}
            </SidebarMenuSub>
          </SidebarMenuItem>

          {/* Sign-in Analysis Section */}
           <SidebarMenuItem>
            <SidebarMenuButton icon={Fingerprint}>Sign-in Analysis</SidebarMenuButton>
            <SidebarMenuSub>
               {renderMenuItems([
                "User Sign-in Location", "Last Log-on Summary", "External User Sign-ins",
                "Guest Sign-ins", "Non-Compliant Device Sign-ins", "Unmanaged Device Sign-ins",
                "MFA failed Sign-ins", "2FA Authentication Methods", "Conditional Access failures",
                "Conditional Access Policies", "Sign-ins Risk Level", "Unlikely Travel Risky Sign-ins",
                "Anonymous IP Sign-ins", "Compromised Risky Sign-ins"
               ], Fingerprint)}
            </SidebarMenuSub>
          </SidebarMenuItem>

          {/* Email Section */}
          <SidebarMenuItem>
            <SidebarMenuButton icon={Mail}>Email</SidebarMenuButton>
             <SidebarMenuSub>
               {renderMenuItems([
                  "Send as Emails", "Send on Behalf Emails", "Undelivered Mails",
                  "Mails Sent by Delegates", "Email Traffic", "External Email Forwarding",
                  "Spam Detections", "Top Phish Receivers", "External Spoof Mails", "Top Malwares",
                  "Internal Mail flow", "External Mail flow", "Users' Active Hours",
                  // Including previous email reports
                  "Email Activities", "Domain-wise Summary", "Group Email Activities",
                  "Org Email Traffic Stats", "User Email Traffic Stats", "Shared Mailbox Traffic Stats",
                  "Email Traffic by 30 min", "Hourly Email Traffic", "Daily Email Traffic",
                  "Monthly Email Traffic", "Groups Mail Traffic Stats", "Total Mails By Hour/Day",
                  "Organizations' Total Mails", "User Total Mails", "Shared Mailbox Total Mails",
                  "Groups Total Mails", "Peak Period Analysis"
               ], Mail)}
             </SidebarMenuSub>
          </SidebarMenuItem>

          {/* Teams Section */}
           <SidebarMenuItem>
            <SidebarMenuButton icon={Users2}>Teams</SidebarMenuButton>
             <SidebarMenuSub>
               {renderMenuItems([
                "Public Teams", "Private Teams", "Teams Membership", "Private & Shared Channels",
                "Login Activities", // Consider differentiating if needed
                "Private Channels", "Team Setting Changes", "Inactive Users", // Consider differentiating
                "Teams Device Usage", "External File Sharing", "Teams Add-ons",
                "Private Channel Membership Changes", "Ownership Promotions and Demotions"
               ], Users2)}
             </SidebarMenuSub>
          </SidebarMenuItem>

           {/* Exchange Online Section */}
           <SidebarMenuItem>
            <SidebarMenuButton icon={Box}>Exchange Online</SidebarMenuButton>
            <SidebarMenuSub>
               {renderMenuItems([
                  "Mailbox Usage", "Inactive Mailboxes", "Active Out of Office Settings",
                  "Mailbox Permissions", "Audit Disabled Mailboxes", "Mailbox Hold",
                  "Mailbox Non-owner Access", "Mobile Device Configurations", "Public Folders",
                  "Role Assignments", "Mail Flow", // Consider differentiating
                  "Exchange Contacts"
               ], Box)}
            </SidebarMenuSub>
          </SidebarMenuItem>


        </SidebarMenu>
      </SidebarContent>
       {/* <SidebarFooter className="p-2">
        <SidebarMenu>
          <SidebarMenuItem>
            <SidebarMenuButton icon={Settings}>Settings</SidebarMenuButton>
          </SidebarMenuItem>
        </SidebarMenu>
      </SidebarFooter> */}
    </Sidebar>
  );
}
