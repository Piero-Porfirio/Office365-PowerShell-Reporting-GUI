
import { NextResponse } from 'next/server';
import type { IntuneDeviceCompliance, IntuneWindowsUpdateStatus, IntuneEnrolledDevice, IntuneAppInfo } from '@/types/reporting'; // Import Intune types

// --- Simulated Intune Data ---
const simulatedDeviceCompliance: IntuneDeviceCompliance[] = [
  { id: 'dev1', displayName: 'Laptop-Alice', operatingSystem: 'Windows 11', complianceState: 'Compliant', lastSyncDateTime: '2024-05-21T10:00:00Z' },
  { id: 'dev2', displayName: 'Surface-Bob', operatingSystem: 'Windows 10', complianceState: 'NonCompliant', lastSyncDateTime: '2024-05-20T09:00:00Z' },
  { id: 'dev3', displayName: 'MacBook-Charlie', operatingSystem: 'macOS', complianceState: 'Compliant', lastSyncDateTime: '2024-05-21T11:00:00Z' },
  { id: 'dev4', displayName: 'Android-Eve', operatingSystem: 'Android', complianceState: 'InGracePeriod', lastSyncDateTime: '2024-05-21T09:30:00Z' },
];

const simulatedUpdateStatus: IntuneWindowsUpdateStatus[] = [
  { deviceId: 'dev1', deviceName: 'Laptop-Alice', osVersion: '10.0.22631', status: 'Up-to-date', lastScanTime: '2024-05-21T08:00:00Z', lastUpdateTime: '2024-05-20T01:00:00Z' },
  { deviceId: 'dev2', deviceName: 'Surface-Bob', osVersion: '10.0.19045', status: 'Pending Updates', lastScanTime: '2024-05-20T07:30:00Z', lastUpdateTime: '2024-05-15T02:00:00Z' },
];

const simulatedEnrolledDevices: IntuneEnrolledDevice[] = [
  { deviceName: 'Laptop-Alice', operatingSystem: 'Windows 11', enrollmentType: 'UserEnrollment', managementAgent: 'MDM', lastSyncDateTime: '2024-05-21T10:00:00Z' },
  { deviceName: 'iPhone-Dave', operatingSystem: 'iOS', enrollmentType: 'DeviceEnrollment', managementAgent: 'MDM', lastSyncDateTime: '2024-05-21T10:30:00Z' },
  { deviceName: 'MacBook-Charlie', operatingSystem: 'macOS', enrollmentType: 'UserEnrollment', managementAgent: 'MDM', lastSyncDateTime: '2024-05-21T11:00:00Z' },
];

const simulatedAppInventory: IntuneAppInfo[] = [
  { displayName: 'Company Portal', publisher: 'Microsoft Corporation', version: '5.2403.0' },
  { displayName: 'Internal HR App', publisher: 'Contoso IT', version: '2.1.5' },
  { displayName: 'Adobe Acrobat Reader DC', publisher: 'Adobe Systems, Incorporated', version: '23.008.20470' },
];

// --- Helper to get simulated data based on command ---
function getSimulatedData(command: string): { output: string; error?: string } {
    console.log(`Intune API received command (simulated): ${command}`);

    // Match command content/keywords to return appropriate simulated data
    // This requires understanding the expected output of the PowerShell commands
    if (command.includes('Get-MgDeviceManagementManagedDevice') && command.includes('ComplianceState')) {
        return { output: JSON.stringify(simulatedDeviceCompliance) };
    }
    if (command.includes('Get-MgDeviceManagementManagedDeviceWindowsUpdateState')) {
        return { output: JSON.stringify(simulatedUpdateStatus) };
    }
     if (command.includes('Get-MgDeviceManagementManagedDevice') && command.includes('EnrollmentType')) {
         // Simple check, might need refinement if other Get-MgDeviceManagementManagedDevice calls are used
         return { output: JSON.stringify(simulatedEnrolledDevices) };
     }
    if (command.includes('Get-MgDeviceAppManagementMobileApp')) {
        return { output: JSON.stringify(simulatedAppInventory) };
    }

    // Default for unimplemented commands or errors
    console.warn(`Intune API: No simulation defined for command: ${command}`);
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

    // Optional: Simulate a delay
    // await new Promise(resolve => setTimeout(resolve, 600));

    // Return the simulated result
    if (result.error) {
         // Return 200 OK but include the simulation error message
        return NextResponse.json(result);
        // Or: return NextResponse.json(result, { status: 404 }); // If simulation not found is an error
    }

    return NextResponse.json(result);

  } catch (error) {
    console.error('Error in Intune API simulation:', error);
    const errorMessage = error instanceof Error ? error.message : 'An unknown error occurred';
    // Return a generic error structure consistent with the PowerShellResult interface
    return NextResponse.json({ output: '', error: `Internal Server Error in Simulation: ${errorMessage}` }, { status: 500 });
  }
}
