import { NextResponse } from 'next/server';
import { executePowerShellCommand } from '@/services/powershell';

export async function POST(request: Request) {
  try {
    const { command } = await request.json();

    if (typeof command !== 'string' || !command.trim()) {
      return NextResponse.json({ error: 'Invalid command provided' }, { status: 400 });
    }

    // In a real application, you would implement secure execution here.
    // This is a stub implementation.
    // WARNING: Directly executing user input commands is dangerous.
    // Proper sanitization, validation, and potentially sandboxing are required.
    const result = await executePowerShellCommand(command);

    return NextResponse.json(result);
  } catch (error) {
    console.error('Error executing PowerShell command:', error);
     const errorMessage = error instanceof Error ? error.message : 'An unknown error occurred';
    return NextResponse.json({ error: `Internal Server Error: ${errorMessage}` }, { status: 500 });
  }
}
