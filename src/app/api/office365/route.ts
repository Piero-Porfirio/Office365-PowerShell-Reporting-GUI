import { NextResponse } from 'next/server';
import { executeOffice365PowerShellCommand } from '@/services/office365';

export async function POST(request: Request) {
  try {
    const { command } = await request.json();

     if (typeof command !== 'string' || !command.trim()) {
      return NextResponse.json({ error: 'Invalid command provided' }, { status: 400 });
    }

    // Stub implementation - replace with actual Office 365 command execution logic
    // Ensure secure handling and authentication with Office 365.
    const result = await executeOffice365PowerShellCommand(command);

    return NextResponse.json(result);
  } catch (error) {
    console.error('Error executing Office 365 PowerShell command:', error);
    const errorMessage = error instanceof Error ? error.message : 'An unknown error occurred';
    return NextResponse.json({ error: `Internal Server Error: ${errorMessage}` }, { status: 500 });
  }
}
