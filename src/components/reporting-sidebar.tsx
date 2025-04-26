
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
  Calendar, // Changed from CalendarMonth
  Group,
  User,
  Box,
  Building2,
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import type { ReportCategory } from '@/types/reporting';
import { cn } from '@/lib/utils';

interface ReportingSidebarProps {
  selectedReport: ReportCategory | null;
  onSelectReport: (report: ReportCategory | null) => void;
}

// Helper to create menu items/buttons
const MenuItem = ({
  report,
  icon,
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
  const Icon = icon;
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
        'justify-start',
        isSubItem && 'h-8 text-xs' // Adjust sub-item styling if needed
      )}
      asChild={isSubItem} // Use anchor tag for sub-buttons if needed for routing later
    >
      {isSubItem ? (
        <a href="#"> {/* Placeholder href */}
          {Icon && <Icon className="size-4 mr-2 shrink-0" />}
          <span>{label}</span>
        </a>
      ) : (
        <>
          {Icon && <Icon />}
          <span>{label}</span>
        </>
      )}
    </Comp>
  );
};

export function ReportingSidebar({ selectedReport, onSelectReport }: ReportingSidebarProps) {
   const { isMobile } = useSidebar();

  return (
    <Sidebar collapsible="icon">
       <SidebarHeader className="items-center gap-2">
        <SidebarTrigger className={cn('md:hidden', isMobile && 'block')} />
        <span className="text-lg font-semibold group-data-[collapsible=icon]:hidden">
          Analytics
        </span>
      </SidebarHeader>
      <SidebarContent className="p-2">
        <SidebarMenu>
          {/* Email Activities Section (Example with Submenu) */}
          <SidebarMenuItem>
            <SidebarMenuButton icon={Mail}>Email Activities</SidebarMenuButton>
             {/* Add sub-items if needed, e.g.,
            <SidebarMenuSub>
               <SidebarMenuSubItem>
                <MenuItem report="Sub Activity 1" label="Sub Activity 1" selectedReport={selectedReport} onSelectReport={onSelectReport} isSubItem />
              </SidebarMenuSubItem>
            </SidebarMenuSub>
            */}
          </SidebarMenuItem>

          <SidebarMenuItem>
            <MenuItem
              report="Domain-wise Summary"
              icon={AreaChart}
              label="Domain-wise Summary"
              selectedReport={selectedReport}
              onSelectReport={onSelectReport}
            />
          </SidebarMenuItem>

          <SidebarMenuItem>
            <MenuItem
              report="Group Email Activities"
              icon={Users}
              label="Group Email Activities"
              selectedReport={selectedReport}
              onSelectReport={onSelectReport}
            />
          </SidebarMenuItem>

          {/* Email Traffic Summary Section */}
          <SidebarMenuItem>
            <SidebarMenuButton icon={BarChartHorizontal}>
              Email Traffic Summary
            </SidebarMenuButton>
            <SidebarMenuSub>
              <SidebarMenuSubItem>
                <MenuItem report="Org Email Traffic Stats" icon={Building2} label="Org Traffic Stats" selectedReport={selectedReport} onSelectReport={onSelectReport} isSubItem />
              </SidebarMenuSubItem>
               <SidebarMenuSubItem>
                <MenuItem report="User Email Traffic Stats" icon={User} label="User Traffic Stats" selectedReport={selectedReport} onSelectReport={onSelectReport} isSubItem />
              </SidebarMenuSubItem>
              <SidebarMenuSubItem>
                <MenuItem report="Shared Mailbox Traffic Stats" icon={Inbox} label="Shared Mailbox Stats" selectedReport={selectedReport} onSelectReport={onSelectReport} isSubItem />
              </SidebarMenuSubItem>
               <SidebarMenuSubItem>
                <MenuItem report="Email Traffic by 30 min" icon={Clock} label="Traffic by 30 min" selectedReport={selectedReport} onSelectReport={onSelectReport} isSubItem />
              </SidebarMenuSubItem>
              <SidebarMenuSubItem>
                <MenuItem report="Hourly Email Traffic" icon={CalendarClock} label="Hourly Traffic" selectedReport={selectedReport} onSelectReport={onSelectReport} isSubItem />
              </SidebarMenuSubItem>
               <SidebarMenuSubItem>
                <MenuItem report="Daily Email Traffic" icon={CalendarDays} label="Daily Traffic" selectedReport={selectedReport} onSelectReport={onSelectReport} isSubItem />
              </SidebarMenuSubItem>
               <SidebarMenuSubItem>
                <MenuItem report="Monthly Email Traffic" icon={Calendar} label="Monthly Traffic" selectedReport={selectedReport} onSelectReport={onSelectReport} isSubItem />
              </SidebarMenuSubItem>
            </SidebarMenuSub>
          </SidebarMenuItem>

           {/* Groups Mail Traffic Stats */}
          <SidebarMenuItem>
            <MenuItem
              report="Groups Mail Traffic Stats"
              icon={Group}
              label="Groups Mail Traffic"
              selectedReport={selectedReport}
              onSelectReport={onSelectReport}
            />
          </SidebarMenuItem>

           {/* Total Mails By Hour/Day */}
           <SidebarMenuItem>
            <MenuItem
              report="Total Mails By Hour/Day"
              icon={Clock} // Using Clock, adjust if better icon exists
              label="Total Mails By Hour/Day"
              selectedReport={selectedReport}
              onSelectReport={onSelectReport}
            />
          </SidebarMenuItem>

           {/* Organizations' Total Mails */}
           <SidebarMenuItem>
            <MenuItem
              report="Organizations' Total Mails"
              icon={Building2}
              label="Organization Total Mails"
              selectedReport={selectedReport}
              onSelectReport={onSelectReport}
            />
          </SidebarMenuItem>

           {/* User Total Mails */}
           <SidebarMenuItem>
            <MenuItem
              report="User Total Mails"
              icon={User}
              label="User Total Mails"
              selectedReport={selectedReport}
              onSelectReport={onSelectReport}
            />
          </SidebarMenuItem>

           {/* Shared Mailbox Total Mails */}
            <SidebarMenuItem>
            <MenuItem
              report="Shared Mailbox Total Mails"
              icon={Inbox}
              label="Shared Mailbox Total Mails"
              selectedReport={selectedReport}
              onSelectReport={onSelectReport}
            />
          </SidebarMenuItem>

           {/* Groups Total Mails */}
           <SidebarMenuItem>
            <MenuItem
              report="Groups Total Mails"
              icon={Group}
              label="Groups Total Mails"
              selectedReport={selectedReport}
              onSelectReport={onSelectReport}
            />
          </SidebarMenuItem>

           {/* Peak Period Analysis */}
           <SidebarMenuItem>
            <MenuItem
              report="Peak Period Analysis"
              icon={LineChart}
              label="Peak Period Analysis"
              selectedReport={selectedReport}
              onSelectReport={onSelectReport}
            />
          </SidebarMenuItem>

        </SidebarMenu>
      </SidebarContent>
      <SidebarFooter className="p-2">
        <SidebarMenu>
          <SidebarMenuItem>
            <SidebarMenuButton icon={Settings}>Settings</SidebarMenuButton>
          </SidebarMenuItem>
        </SidebarMenu>
      </SidebarFooter>
    </Sidebar>
  );
}
