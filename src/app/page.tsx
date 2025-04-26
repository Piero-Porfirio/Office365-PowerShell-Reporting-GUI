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
    <main className="flex min-h-screen flex-col items-center p-4 md:p-8 lg:p-12">
      <div className="w-full max-w-4xl">
        <h1 className="text-3xl font-bold mb-6 text-center text-primary">
          <TerminalSquare className="inline-block mr-2 h-8 w-8 align-text-bottom" />
          PowerShell GUI
        </h1>
        <Tabs defaultValue="azure" className="w-full">
          <TabsList className="grid w-full grid-cols-4 mb-4">
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
          <TabsContent value="azure">
            <CommandExecutor
              title="Azure PowerShell"
              commandType="Azure"
              executeCommand={executeAzurePowerShellCommand}
              placeholder="Enter Azure PowerShell command (e.g., Get-AzVM)"
            />
          </TabsContent>
          <TabsContent value="office365">
            <CommandExecutor
              title="Office 365 PowerShell"
              commandType="Office365"
              executeCommand={executeOffice365PowerShellCommand}
              placeholder="Enter Office 365 PowerShell command (e.g., Get-Mailbox)"
            />
          </TabsContent>
          <TabsContent value="powershell">
            <CommandExecutor
              title="General PowerShell"
              commandType="PowerShell"
              executeCommand={executePowerShellCommand}
              placeholder="Enter general PowerShell command (e.g., Get-Process)"
            />
          </TabsContent>
          <TabsContent value="reporting">
            <Reporting />
          </TabsContent>
        </Tabs>
      </div>
    </main>
  );
}
