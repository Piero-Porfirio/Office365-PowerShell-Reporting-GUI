
import { NextResponse } from 'next/server';
// Remove direct PowerShell execution import
// import { executeOffice365PowerShellCommand } from '@/services/office365'; // Assuming this was backend logic
import type { SharedMailboxTrafficStats } from '@/types/reporting'; // Import types for simulated data

// --- Simulated Data ---
const simulatedSharedMailboxStats: SharedMailboxTrafficStats[] = [
  { month: 'April 2025', sharedMailboxName: '24x7support', sharedMailboxUPN: '24x7support@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 1, externalMailsSent: null, externalMailsReceived: null },
  { month: 'April 2025', sharedMailboxName: 'Amy Tucker', sharedMailboxUPN: 'amy.tucker@o365droid.onmicrosoft.com', totalMailsSent: 1, totalMailsReceived: 8, externalMailsSent: 5, externalMailsReceived: 1 },
  { month: 'April 2025', sharedMailboxName: 'Finance', sharedMailboxUPN: 'finance@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 19, externalMailsSent: 10, externalMailsReceived: null },
  { month: 'April 2025', sharedMailboxName: 'HR', sharedMailboxUPN: 'hr@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 15, externalMailsSent: 2, externalMailsReceived: null },
  { month: 'April 2025', sharedMailboxName: 'Production', sharedMailboxUPN: 'production@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 4, externalMailsSent: 4, externalMailsReceived: null },
  { month: 'April 2025', sharedMailboxName: 'TL', sharedMailboxUPN: 'tl@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 14, externalMailsSent: 3, externalMailsReceived: null },
  { month: 'March 2025', sharedMailboxName: 'Amy Tucker', sharedMailboxUPN: 'amy.tucker@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 18, externalMailsSent: 18, externalMailsReceived: null },
  { month: 'March 2025', sharedMailboxName: 'Finance', sharedMailboxUPN: 'finance@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 45, externalMailsSent: 19, externalMailsReceived: null },
  { month: 'March 2025', sharedMailboxName: 'HR', sharedMailboxUPN: 'hr@o365droid.onmicrosoft.com', totalMailsSent: null, totalMailsReceived: 29, externalMailsSent: 4, externalMailsReceived: null },
];


// --- Helper to get simulated data based on command ---
function getSimulatedData(command: string): { output: string; error?: string } {
    console.log(`Office 365 API received command (simulated): ${command}`);

    // Match command prefixes/content to return appropriate simulated data
    // For now, only simulating Shared Mailbox Stats based on placeholder command comment
    if (command.includes('Shared Mailbox Traffic Stats') || command.includes('Get-MailTrafficReport') || command.includes('Get-MailboxStatistics')) {
         // This is a very basic check; a real backend would need robust parsing
        return { output: JSON.stringify(simulatedSharedMailboxStats) };
    }

    // Default for unimplemented commands or errors
    console.warn(`Office 365 API: No simulation defined for command: ${command}`);
    // Return empty array structure for table component compatibility
    return { output: JSON.stringify([]), error: `Command simulation not implemented: ${command}` };
}


export async function POST(request: Request) {
  try {
    const { command } = await request.json();

     if (typeof command !== 'string' || !command.trim()) {
      return NextResponse.json({ output: '', error: 'Invalid command provided' }, { status: 400 });
    }

    // --- Simulate PowerShell execution ---
    const result = getSimulatedData(command);
    // --- End Simulation ---

    // Optional: Simulate delay
    // await new Promise(resolve => setTimeout(resolve, 700));

     if (result.error) {
        // Return 200 OK but include the simulation error message
        return NextResponse.json(result);
        // Or: return NextResponse.json(result, { status: 404 });
     }

    return NextResponse.json(result);

  } catch (error) {
    console.error('Error in Office 365 API simulation:', error);
    const errorMessage = error instanceof Error ? error.message : 'An unknown error occurred';
     // Return a generic error structure
    return NextResponse.json({ output: '', error: `Internal Server Error in Simulation: ${errorMessage}` }, { status: 500 });
  }
}
