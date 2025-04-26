/**
 * Represents the result of a PowerShell command execution.
 */
export interface PowerShellResult {
  /**
   * The output of the PowerShell command.
   */
  output: string;
  /**
   * Any error messages generated during the command execution.
   */
  error?: string;
}

/**
 * Executes a PowerShell command via the API endpoint.
 *
 * @param command The PowerShell command to execute.
 * @returns A promise that resolves to a PowerShellResult object.
 */
export async function executePowerShellCommand(command: string): Promise<PowerShellResult> {
  try {
    const response = await fetch('/api/powershell', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ command }),
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({ error: 'Failed to parse error response' }));
      throw new Error(errorData?.error || `HTTP error! status: ${response.status}`);
    }

    const result: PowerShellResult = await response.json();
    return result;
  } catch (error) {
    console.error('Error calling PowerShell API:', error);
    const errorMessage = error instanceof Error ? error.message : 'An unknown network error occurred';
    return { output: '', error: `API call failed: ${errorMessage}` };
  }
}
