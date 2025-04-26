
/**
 * Represents the result of an Azure PowerShell command execution.
 */
export interface AzurePowerShellResult {
  /**
   * The output of the Azure PowerShell command.
   */
  output: string;
  /**
   * Any error messages generated during the command execution.
   */
  error?: string;
}

/**
 * Executes an Azure PowerShell command via the API endpoint.
 *
 * @param command The Azure PowerShell command to execute.
 * @returns A promise that resolves to a AzurePowerShellResult object.
 */
export async function executeAzurePowerShellCommand(command: string): Promise<AzurePowerShellResult> {
  console.log(`Executing Azure command via API: ${command}`); // Log the command
  try {
    const response = await fetch('/api/azure', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ command }),
    });

    if (!response.ok) {
       const errorData = await response.json().catch(() => ({ error: 'Failed to parse error response' }));
       console.error(`Azure API HTTP Error ${response.status}:`, errorData?.error || 'Unknown error'); // Log HTTP error details
       throw new Error(errorData?.error || `HTTP error! status: ${response.status}`);
    }

    const result: AzurePowerShellResult = await response.json();
    // Optionally log the received result for debugging
    // console.log('Azure API Result:', result);
    return result;
  } catch (error) {
     console.error('Error calling Azure API:', error);
     const errorMessage = error instanceof Error ? error.message : 'An unknown network error occurred';
     return { output: '', error: `API call failed: ${errorMessage}` };
  }
}

    