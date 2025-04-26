'use client';

import type { PowerShellResult } from '@/services/powershell';
import type { AzurePowerShellResult } from '@/services/azure';
import type { Office365PowerShellResult } from '@/services/office365';
import type { SuggestPowerShellOutput } from '@/ai/flows/suggest-powershell-flow';
import React, { useState, useCallback, useRef, useEffect } from 'react';
import { Button } from '@/components/ui/button';
import { Textarea } from '@/components/ui/textarea';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Loader2, Play, Bot } from 'lucide-react';
import { ResultsDisplay } from './results-display';
import { useToast } from '@/hooks/use-toast';
import { suggestPowerShellCommand } from '@/ai/flows/suggest-powershell-flow';
import { Popover, PopoverContent, PopoverTrigger } from '@/components/ui/popover';
import { ScrollArea } from '@/components/ui/scroll-area';

type CommandExecutorProps = {
  title: string;
  commandType: 'Azure' | 'Office365' | 'PowerShell';
  executeCommand: (
    command: string
  ) => Promise<PowerShellResult | AzurePowerShellResult | Office365PowerShellResult>;
  placeholder?: string;
};

type CommandResult = PowerShellResult | AzurePowerShellResult | Office365PowerShellResult | null;

// Debounce hook
function useDebounce<T>(value: T, delay: number): T {
  const [debouncedValue, setDebouncedValue] = useState<T>(value);

  useEffect(() => {
    const timer = setTimeout(() => setDebouncedValue(value), delay);
    return () => clearTimeout(timer);
  }, [value, delay]);

  return debouncedValue;
}


export function CommandExecutor({
  title,
  commandType,
  executeCommand,
  placeholder = 'Enter command...',
}: CommandExecutorProps) {
  const [command, setCommand] = useState('');
  const [result, setResult] = useState<CommandResult>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [suggestions, setSuggestions] = useState<SuggestPowerShellOutput['suggestions']>([]);
  const [isSuggestionsLoading, setIsSuggestionsLoading] = useState(false);
  const [isSuggestionPopoverOpen, setIsSuggestionPopoverOpen] = useState(false);
  const { toast } = useToast();
  const debounceTimeoutRef = useRef<NodeJS.Timeout | null>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const fetchSuggestions = useCallback(async (currentInput: string) => {
    if (currentInput.trim().length < 2) {
      setSuggestions([]);
      setIsSuggestionPopoverOpen(false);
      return;
    }
    setIsSuggestionsLoading(true);
    try {
      const output = await suggestPowerShellCommand({ currentInput, commandType });
      // Only update if there are suggestions to show
      if (output.suggestions && output.suggestions.length > 0) {
        setSuggestions(output.suggestions);
        setIsSuggestionPopoverOpen(true); // Open popover when suggestions arrive
      } else {
        setSuggestions([]);
        setIsSuggestionPopoverOpen(false); // Close if no suggestions
      }
    } catch (error) {
      console.error('Error fetching suggestions:', error);
      setSuggestions([]);
      setIsSuggestionPopoverOpen(false);
      // Optionally show a toast for suggestion errors, but might be noisy
      // toast({ title: "Suggestion Error", description: "Could not fetch command suggestions.", variant: "destructive" });
    } finally {
      setIsSuggestionsLoading(false);
    }
  }, [commandType]);


  const handleInputChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    const newValue = e.target.value;
    setCommand(newValue);

    if (debounceTimeoutRef.current) {
      clearTimeout(debounceTimeoutRef.current);
    }

    debounceTimeoutRef.current = setTimeout(() => {
       if (newValue.trim().length >= 2) { // Only fetch if input is meaningful
         fetchSuggestions(newValue);
       } else {
         setSuggestions([]); // Clear suggestions for short input
         setIsSuggestionPopoverOpen(false);
       }
    }, 500); // 500ms debounce delay
  };

  const handleSuggestionClick = (suggestionCmd: string) => {
    setCommand(suggestionCmd);
    setSuggestions([]);
    setIsSuggestionPopoverOpen(false);
    textareaRef.current?.focus(); // Refocus textarea after selection
  };

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
    setIsSuggestionPopoverOpen(false); // Close suggestions on execute

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

   // Close popover if textarea loses focus, unless the focus moved to the popover itself
  const handleBlur = (e: React.FocusEvent<HTMLTextAreaElement>) => {
    // Delay check slightly to allow focus to shift to popover
    setTimeout(() => {
      if (
        document.activeElement !== textareaRef.current &&
        !document.activeElement?.closest('[data-radix-popper-content-wrapper]') // Check if focus is inside any popover
      ) {
        setIsSuggestionPopoverOpen(false);
      }
    }, 100);
  };


  return (
    <Card className="w-full shadow-lg">
      <CardHeader>
        <CardTitle className="text-xl">{title}</CardTitle>
      </CardHeader>
      <CardContent>
        <div className="space-y-4">
          {/* Wrap Textarea with Popover for suggestions */}
          <Popover open={isSuggestionPopoverOpen} onOpenChange={setIsSuggestionPopoverOpen}>
             <PopoverTrigger asChild>
               {/* Anchor the popover to the textarea */}
               {/* Using a div wrapper allows PopoverTrigger to work correctly with Textarea */}
               <div className="relative">
                <Textarea
                  ref={textareaRef}
                  placeholder={placeholder}
                  value={command}
                  onChange={handleInputChange}
                  onFocus={() => { if (suggestions.length > 0) setIsSuggestionPopoverOpen(true); } } // Open on focus if suggestions exist
                  onBlur={handleBlur} // Close on blur
                  className="min-h-[100px] font-mono text-sm bg-muted/50 border-input focus:ring-primary"
                  aria-label={`${commandType} command input`}
                  disabled={isLoading}
                  aria-autocomplete="list"
                  aria-controls="suggestions-list"
                 />
                 {isSuggestionsLoading && (
                    <Loader2 className="absolute right-2 top-2 h-4 w-4 animate-spin text-muted-foreground" />
                  )}
               </div>
             </PopoverTrigger>
             <PopoverContent
                id="suggestions-list"
                className="w-[--radix-popover-trigger-width] p-0" // Match width of trigger
                align="start"
                side="bottom"
                sideOffset={5}
                onOpenAutoFocus={(e) => e.preventDefault()} // Prevent stealing focus
             >
                <ScrollArea className="max-h-[200px]">
                  <ul className="text-sm">
                    {suggestions.map((suggestion, index) => (
                      <li key={index}
                          className="p-2 hover:bg-accent hover:text-accent-foreground cursor-pointer"
                          onMouseDown={(e) => { // Use onMouseDown to prevent blur before click registers
                             e.preventDefault(); // Prevent blur
                             handleSuggestionClick(suggestion.command);
                          }}
                          role="option"
                          aria-selected={false} // Can add focus management later
                      >
                         <div className="font-mono font-semibold">{suggestion.command}</div>
                         <div className="text-xs text-muted-foreground">{suggestion.description}</div>
                      </li>
                    ))}
                  </ul>
                </ScrollArea>
                <div className="p-2 text-xs text-muted-foreground border-t flex items-center gap-1">
                    <Bot size={14} /> AI Suggestions
                </div>
             </PopoverContent>
          </Popover>

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
