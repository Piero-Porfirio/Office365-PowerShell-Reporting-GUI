'use client';

import type { PowerShellResult } from '@/services/powershell';
import type { AzurePowerShellResult } from '@/services/azure';
import type { Office365PowerShellResult } from '@/services/office365';
import React, { useState } from 'react';
import { Button } from '@/components/ui/button';
import { Textarea } from '@/components/ui/textarea';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Loader2, Play } from 'lucide-react';
import { ResultsDisplay } from './results-display';
import { useToast } from '@/hooks/use-toast';

type CommandExecutorProps = {
  title: string;
  commandType: 'Azure' | 'Office365' | 'PowerShell';
  executeCommand: (
    command: string
  ) => Promise<PowerShellResult | AzurePowerShellResult | Office365PowerShellResult>;
  placeholder?: string;
};

type CommandResult = PowerShellResult | AzurePowerShellResult | Office365PowerShellResult | null;

export function CommandExecutor({
  title,
  commandType,
  executeCommand,
  placeholder = 'Enter command...',
}: CommandExecutorProps) {
  const [command, setCommand] = useState('');
  const [result, setResult] = useState<CommandResult>(null);
  const [isLoading, setIsLoading] = useState(false);
  const { toast } = useToast();

  const handleExecute = async () => {
    if (!command.trim()) {
      toast({
        title: 'Input Error',
        description: 'Please enter a command to execute.',
        variant: 'destructive',
      });
      return;
    }

    setIsLoading(true);
    setResult(null); // Clear previous results

    try {
      const commandResult = await executeCommand(command);
      setResult(commandResult);
      if (commandResult.error) {
         toast({
           title: `${commandType} Command Error`,
           description: commandResult.error || 'An unknown error occurred.',
           variant: 'destructive',
         });
      } else {
          toast({
            title: `${commandType} Command Executed`,
            description: 'Command executed successfully.',
          });
      }
    } catch (error) {
      console.error('Error executing command:', error);
      const errorMessage = error instanceof Error ? error.message : String(error);
      setResult({ output: '', error: `Failed to execute command: ${errorMessage}` });
       toast({
         title: 'Execution Failed',
         description: `Failed to execute command: ${errorMessage}`,
         variant: 'destructive',
       });
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <Card className="w-full shadow-lg">
      <CardHeader>
        <CardTitle className="text-xl">{title}</CardTitle>
      </CardHeader>
      <CardContent>
        <div className="space-y-4">
          <Textarea
            placeholder={placeholder}
            value={command}
            onChange={(e) => setCommand(e.target.value)}
            className="min-h-[100px] font-mono text-sm bg-muted/50 border-input focus:ring-primary"
            aria-label={`${commandType} command input`}
            disabled={isLoading}
          />
          <Button
            onClick={handleExecute}
            disabled={isLoading}
            className="w-full transition-colors duration-200"
            aria-label={`Execute ${commandType} command`}
          >
            {isLoading ? (
              <>
                <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                Executing...
              </>
            ) : (
              <>
                <Play className="mr-2 h-4 w-4" />
                Execute Command
              </>
            )}
          </Button>
          {result && <ResultsDisplay result={result} />}
        </div>
      </CardContent>
    </Card>
  );
}
