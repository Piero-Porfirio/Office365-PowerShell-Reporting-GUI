'use server';
/**
 * @fileOverview Provides PowerShell command suggestions using an AI model.
 *
 * - suggestPowerShellCommand - A function that suggests PowerShell commands.
 * - SuggestPowerShellInput - The input type for the suggestPowerShellCommand function.
 * - SuggestPowerShellOutput - The return type for the suggestPowerShellCommand function.
 */

import {ai} from '@/ai/ai-instance';
import {z} from 'genkit';

const SuggestPowerShellInputSchema = z.object({
  currentInput: z.string().describe('The current text typed by the user in the command input field.'),
  commandType: z.enum(['Azure', 'Office365', 'PowerShell']).describe('The context or type of PowerShell environment (Azure, Office 365, or general PowerShell).'),
});
export type SuggestPowerShellInput = z.infer<typeof SuggestPowerShellInputSchema>;

const SuggestionSchema = z.object({
    command: z.string().describe('The suggested PowerShell command or cmdlet.'),
    description: z.string().describe('A brief description of what the command does or its common usage.'),
});

const SuggestPowerShellOutputSchema = z.object({
  suggestions: z.array(SuggestionSchema).describe('A list of relevant PowerShell command suggestions.'),
});
export type SuggestPowerShellOutput = z.infer<typeof SuggestPowerShellOutputSchema>;

export async function suggestPowerShellCommand(input: SuggestPowerShellInput): Promise<SuggestPowerShellOutput> {
  // Basic check to avoid calling the flow with very short/empty input
  if (input.currentInput.trim().length < 2) {
    return { suggestions: [] };
  }
  return suggestPowerShellFlow(input);
}

const prompt = ai.definePrompt({
  name: 'suggestPowerShellPrompt',
  input: {
    schema: SuggestPowerShellInputSchema,
  },
  output: {
    schema: SuggestPowerShellOutputSchema,
  },
  prompt: `You are an expert assistant specializing in PowerShell scripting for IT administration, particularly for {{commandType}}.
The user is typing a command in a PowerShell input field. Based on their current input, suggest relevant and valid PowerShell cmdlets or command fragments they might be trying to type.

Provide a list of up to 5 suggestions. For each suggestion, include the command/cmdlet and a brief description. Focus on cmdlets available in the {{commandType}} context.

Current user input: {{{currentInput}}}

Generate suggestions relevant to the input and the {{commandType}} context. Prioritize cmdlets that start with or closely match the user's input. If the input is generic, suggest common starting cmdlets for the context.
`,
});


const suggestPowerShellFlow = ai.defineFlow<
  typeof SuggestPowerShellInputSchema,
  typeof SuggestPowerShellOutputSchema
>(
  {
    name: 'suggestPowerShellFlow',
    inputSchema: SuggestPowerShellInputSchema,
    outputSchema: SuggestPowerShellOutputSchema,
  },
  async (input) => {
    try {
        const { output } = await prompt(input);
        // Ensure output is not null and suggestions is an array
        return output ?? { suggestions: [] };
    } catch (error) {
        console.error("Error in suggestPowerShellFlow:", error);
        // Return empty suggestions in case of error
        return { suggestions: [] };
    }
  }
);
