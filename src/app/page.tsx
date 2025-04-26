
'use client';

import {
  Tabs,
  TabsContent,
  TabsList,
  TabsTrigger,
} from '@/components/ui/tabs';
import { CommandExecutor } from '@/components/command-executor';
import { Reporting } from '@/components/reporting';
import { executeAzurePowerShellCommand } from '@/services/azure';
import { executeOffice365PowerShellCommand } from '@/services/office365';
import { executePowerShellCommand } from '@/services/powershell';
import { Container, Database, FileText, TerminalSquare } from 'lucide-react';

export default function Home() {
  return (
    // Use flex-col and h-screen to make the layout fill the viewport height
    <main className="flex h-screen flex-col items-center p-4 md:p-6 lg:p-8 bg-background">
       {/* Use max-width for content centering and flex-grow to fill height */}
       <div className="w-full max-w-7xl flex flex-col flex-grow overflow-hidden">
        <h1 className="text-3xl font-bold mb-6 text-center text-primary shrink-0">
          <TerminalSquare className="inline-block mr-2 h-8 w-8 align-text-bottom" />
          PowerShell GUI
        </h1>
        {/* Make Tabs component grow to fill remaining space */}
        <Tabs defaultValue="reporting" className="w-full flex flex-col flex-grow overflow-hidden">
           <TabsList className="grid w-full grid-cols-4 mb-4 shrink-0">
            <TabsTrigger value="azure">
              <Database className="mr-2 h-4 w-4" /> Azure
            </TabsTrigger>
            <TabsTrigger value="office365">
              <Container className="mr-2 h-4 w-4" /> Office 365
            </TabsTrigger>
             <TabsTrigger value="powershell">
              <TerminalSquare className="mr-2 h-4 w-4" /> PowerShell
            </TabsTrigger>
            <TabsTrigger value="reporting">
              <FileText className="mr-2 h-4 w-4" /> Reporting
            </TabsTrigger>
          </TabsList>

          {/* Command Executor Tabs - Allow internal scroll */}
          <TabsContent value="azure" className="flex-grow overflow-auto p-1">
            <CommandExecutor
              title="Azure PowerShell"
              commandType="Azure"
              executeCommand={executeAzurePowerShellCommand}
              placeholder="Enter Azure PowerShell command (e.g., Get-AzVM)"
            />
          </TabsContent>
          <TabsContent value="office365" className="flex-grow overflow-auto p-1">
            <CommandExecutor
              title="Office 365 PowerShell"
              commandType="Office365"
              executeCommand={executeOffice365PowerShellCommand}
              placeholder="Enter Office 365 PowerShell command (e.g., Get-Mailbox)"
            />
          </TabsContent>
           <TabsContent value="powershell" className="flex-grow overflow-auto p-1">
            <CommandExecutor
              title="General PowerShell"
              commandType="PowerShell"
              executeCommand={executePowerShellCommand}
              placeholder="Enter general PowerShell command (e.g., Get-Process)"
            />
          </TabsContent>

           {/* Reporting Tab - Ensure it fills space and Reporting component handles internal layout */}
           <TabsContent
             value="reporting"
             className="flex flex-col flex-grow overflow-hidden data-[state=active]:flex" // Ensure display:flex when active
           >
             {/* Reporting component now manages its sidebar + display area */}
             <Reporting />
           </TabsContent>
        </Tabs>
      </div>
    </main>
  );
}
