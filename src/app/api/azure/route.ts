
import { NextResponse } from 'next/server';
// Remove direct PowerShell execution import
// import { executeAzurePowerShellCommand } from '@/services/azure'; // Assuming this was backend logic before
import type { LicenseUtilizationData, AzureADUserGroupInfo } from '@/types/reporting'; // Import types for simulated data

// --- Simulated Data ---
const simulatedLicenseData: LicenseUtilizationData[] = [
  { SkuId: 'c7df2760-2c81-4ef7-b578-5b5392b571df', SkuPartNumber: 'ENTERPRISEPREMIUM', ConsumedUnits: 150, TotalUnits: 200 },
  { SkuId: '6f23d6d4-3be2-4f0d-b04d-a59e13b9e9a7', SkuPartNumber: 'VISIOCLIENT', ConsumedUnits: 25, TotalUnits: 50 },
  { SkuId: 'cfc14177-ef78-4702-864d-3f6d63e56f5a', SkuPartNumber: 'PROJECTPROFESSIONAL', ConsumedUnits: 10, TotalUnits: 15 },
];

const simulatedUserGroupData: AzureADUserGroupInfo[] = [
  { type: 'User', id: 'user1', displayName: 'Alice Smith', userPrincipalName: 'alice@contoso.onmicrosoft.com', license: 'ENTERPRISEPREMIUM' },
  { type: 'User', id: 'user2', displayName: 'Bob Johnson', userPrincipalName: 'bob@contoso.onmicrosoft.com', license: 'ENTERPRISEPREMIUM' },
  { type: 'Group', id: 'group1', displayName: 'Sales Team', description: 'Global Sales Department', groupType: 'Unified' },
  { type: 'Group', id: 'group2', displayName: 'Marketing Team', description: 'Product Marketing', groupType: 'Security' },
  { type: 'User', id: 'user3', displayName: 'Charlie Brown', userPrincipalName: 'charlie@contoso.onmicrosoft.com', license: 'VISIOCLIENT' },
];

// --- Helper to get simulated data based on command ---
function getSimulatedData(command: string): { output: string; error?: string } {
    console.log(`Azure API received command (simulated): ${command}`);

    // Match command prefixes to return appropriate simulated data
    if (command.startsWith('Get-MgSubscribedSku')) {
        return { output: JSON.stringify(simulatedLicenseData) };
    }
    if (command.startsWith('Get-MgUser') || command.startsWith('Get-MgGroup')) {
         // Simple simulation, a real backend would parse the command precisely
         return { output: JSON.stringify(simulatedUserGroupData) };
    }

    // Default for unimplemented commands or errors
    console.warn(`Azure API: No simulation defined for command: ${command}`);
    return { output: JSON.stringify([]), error: `Command simulation not implemented: ${command}` };
}


export async function POST(request: Request) {
  try {
    const { command } = await request.json();

    if (typeof command !== 'string' || !command.trim()) {
      return NextResponse.json({ output: '', error: 'Invalid command provided' }, { status: 400 });
    }

    // --- Simulate PowerShell execution ---
    // Instead of actually running PowerShell, return predefined data based on the command.
    const result = getSimulatedData(command);
    // --- End Simulation ---

    // Optional: Simulate a delay
    // await new Promise(resolve => setTimeout(resolve, 500));

    // Return the simulated result
    if (result.error) {
        // You might want to return a different status code for simulation errors
        // For now, returning 200 but including the error in the body
         return NextResponse.json(result);
        // Or: return NextResponse.json(result, { status: 404 }); // If simulation not found is an error
    }

    return NextResponse.json(result);

  } catch (error) {
    console.error('Error in Azure API simulation:', error);
    const errorMessage = error instanceof Error ? error.message : 'An unknown error occurred';
    // Return a generic error structure consistent with the PowerShellResult interface
    return NextResponse.json({ output: '', error: `Internal Server Error in Simulation: ${errorMessage}` }, { status: 500 });
  }
}
