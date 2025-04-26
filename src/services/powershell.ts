
/**
 * Represents the result of a generic PowerShell command execution.
 * NOTE: In the current setup, specific services (Azure/Office365) are used,
 * but this generic interface remains for potential future use or abstraction.
 */
export interface PowerShellResult {
  /**
   * The string output of the PowerShell command, expected to be JSON for parsing.
   */
  output: string;
  /**
   * Any error messages generated during the command execution or API call.
   */
  error?: string;
}

/**
 * Executes a generic PowerShell command via a hypothetical API endpoint.
 * WARNING: This function is currently **NOT USED** as specific endpoints
 * (/api/azure, /api/office365) handle the simulated requests.
 * Executing arbitrary commands is a security risk.
 *
 * @param command The PowerShell command to execute.
 * @returns A promise that resolves to a PowerShellResult object.
 * @deprecated Use executeAzurePowerShellCommand or executeOffice365PowerShellCommand instead.
 */
export async function executePowerShellCommand(command: string): Promise<PowerShellResult> {
   console.warn("Attempted to call deprecated executePowerShellCommand. Use specific service functions.");
   // Returning an error state as this endpoint is removed/unused.
   return Promise.resolve({ output: '', error: 'Generic PowerShell endpoint is deprecated/removed.' });

  /* // Original implementation (if the endpoint existed):
   try {
    const response = await fetch('/api/powershell', { // This route is deleted
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
  */
}
