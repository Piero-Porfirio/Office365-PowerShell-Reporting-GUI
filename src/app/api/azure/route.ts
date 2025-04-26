import { NextResponse } from 'next/server';
import { executeAzurePowerShellCommand } from '@/services/azure';

export async function POST(request: Request) {
  try {
    const { command } = await request.json();

    if (typeof command !== 'string' || !command.trim()) {
      return NextResponse.json({ error: 'Invalid command provided' }, { status: 400 });
    }

    // Stub implementation - replace with actual Azure command execution logic
    // Ensure secure handling and authentication with Azure.
    const result = await executeAzurePowerShellCommand(command);

    return NextResponse.json(result);
  } catch (error) {
    console.error('Error executing Azure PowerShell command:', error);
    const errorMessage = error instanceof Error ? error.message : 'An unknown error occurred';
    return NextResponse.json({ error: `Internal Server Error: ${errorMessage}` }, { status: 500 });
  }
}
