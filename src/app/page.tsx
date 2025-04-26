
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
    <main className="flex h-screen flex-col items-center p-4 md:p-8 lg:p-12 bg-background">
       <div className="w-full max-w-6xl flex flex-col flex-grow"> {/* Use max-w-6xl and flex-grow */}
        <h1 className="text-3xl font-bold mb-6 text-center text-primary shrink-0"> {/* Make header shrink */}
          <TerminalSquare className="inline-block mr-2 h-8 w-8 align-text-bottom" />
          PowerShell GUI
        </h1>
        {/* Make Tabs component grow to fill space */}
        <Tabs defaultValue="reporting" className="w-full flex flex-col flex-grow">
           <TabsList className="grid w-full grid-cols-4 mb-4 shrink-0"> {/* Make TabsList shrink */}
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
          {/* Command Executor Tabs */}
          <TabsContent value="azure" className="flex-grow overflow-auto">
            <CommandExecutor
              title="Azure PowerShell"
              commandType="Azure"
              executeCommand={executeAzurePowerShellCommand}
              placeholder="Enter Azure PowerShell command (e.g., Get-AzVM)"
            />
          </TabsContent>
          <TabsContent value="office365" className="flex-grow overflow-auto">
            <CommandExecutor
              title="Office 365 PowerShell"
              commandType="Office365"
              executeCommand={executeOffice365PowerShellCommand}
              placeholder="Enter Office 365 PowerShell command (e.g., Get-Mailbox)"
            />
          </TabsContent>
           <TabsContent value="powershell" className="flex-grow overflow-auto">
            <CommandExecutor
              title="General PowerShell"
              commandType="PowerShell"
              executeCommand={executePowerShellCommand}
              placeholder="Enter general PowerShell command (e.g., Get-Process)"
            />
          </TabsContent>
           {/* Reporting Tab - Make it take remaining height and allow internal scrolling */}
           <TabsContent
             value="reporting"
             className="flex flex-col flex-grow overflow-hidden" // Use overflow-hidden here
           >
             <Reporting /> {/* Reporting component now handles its own layout */}
           </TabsContent>
        </Tabs>
      </div>
    </main>
  );
}
