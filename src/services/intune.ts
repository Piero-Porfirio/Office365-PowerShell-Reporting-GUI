
/**
 * Represents the result of an Intune (Microsoft Graph) PowerShell command execution.
 */
export interface IntunePowerShellResult {
  /**
   * The output of the Intune/Graph PowerShell command. Expected to be JSON.
   */
  output: string;
  /**
   * Any error messages generated during the command execution or API call.
   */
  error?: string;
}

/**
 * Executes an Intune PowerShell command via the API endpoint.
 * Note: Currently, this uses the `/api/intune` endpoint which simulates results.
 *
 * @param command The Intune/Graph PowerShell command to execute.
 * @returns A promise that resolves to a IntunePowerShellResult object.
 */
export async function executeIntunePowerShellCommand(command: string): Promise<IntunePowerShellResult> {
  console.log(`Executing Intune command via API: ${command}`); // Log the command
  try {
    const response = await fetch('/api/intune', { // Target the new Intune API route
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ command }),
    });

    if (!response.ok) {
       const errorData = await response.json().catch(() => ({ error: 'Failed to parse error response' }));
       console.error(`Intune API HTTP Error ${response.status}:`, errorData?.error || 'Unknown error'); // Log HTTP error details
       throw new Error(errorData?.error || `HTTP error! status: ${response.status}`);
    }

    const result: IntunePowerShellResult = await response.json();
    // Optionally log the received result for debugging
    // console.log('Intune API Result:', result);
    return result;
  } catch (error) {
     console.error('Error calling Intune API:', error);
     const errorMessage = error instanceof Error ? error.message : 'An unknown network error occurred';
     return { output: '', error: `API call failed: ${errorMessage}` };
  }
}
