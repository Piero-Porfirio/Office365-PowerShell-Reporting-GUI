/**
 * Represents the result of an Office365 PowerShell command execution.
 */
export interface Office365PowerShellResult {
  /**
   * The output of the Office365 PowerShell command.
   */
  output: string;
  /**
   * Any error messages generated during the command execution.
   */
  error?: string;
}

/**
 * Executes an Office365 PowerShell command via the API endpoint.
 *
 * @param command The Office365 PowerShell command to execute.
 * @returns A promise that resolves to a Office365PowerShellResult object.
 */
export async function executeOffice365PowerShellCommand(command: string): Promise<Office365PowerShellResult> {
   try {
    const response = await fetch('/api/office365', {
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

    const result: Office365PowerShellResult = await response.json();
    return result;
  } catch (error) {
     console.error('Error calling Office365 API:', error);
     const errorMessage = error instanceof Error ? error.message : 'An unknown network error occurred';
     return { output: '', error: `API call failed: ${errorMessage}` };
  }
}
