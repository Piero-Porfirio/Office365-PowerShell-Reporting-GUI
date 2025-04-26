'use client';

import type { PowerShellResult } from '@/services/powershell';
import type { AzurePowerShellResult } from '@/services/azure';
import type { Office365PowerShellResult } from '@/services/office365';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { ScrollArea } from '@/components/ui/scroll-area';
import { Terminal } from 'lucide-react';

type ResultsDisplayProps = {
  result: PowerShellResult | AzurePowerShellResult | Office365PowerShellResult;
};

export function ResultsDisplay({ result }: ResultsDisplayProps) {
  return (
    <Card className="mt-4 border-primary/30 shadow-md">
      <CardHeader>
        <CardTitle className="text-lg flex items-center">
          <Terminal className="mr-2 h-5 w-5" />
          Command Results
        </CardTitle>
      </CardHeader>
      <CardContent>
        {result.error ? (
          <Alert variant="destructive">
            <AlertTitle>Error</AlertTitle>
            <AlertDescription>
              <pre className="whitespace-pre-wrap font-mono text-sm">
                {result.error}
              </pre>
            </AlertDescription>
          </Alert>
        ) : (
          <ScrollArea className="h-[300px] w-full rounded-md border border-input p-4 bg-muted/30">
            <pre className="whitespace-pre-wrap font-mono text-sm text-foreground">
              {result.output || 'No output received.'}
            </pre>
          </ScrollArea>
        )}
      </CardContent>
    </Card>
  );
}
