import { spawn } from 'node:child_process';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

export interface ConversionRequest {
  input_path: string;
  output_path: string;
  overwrite?: boolean;
  include_markup?: boolean;
  timeout_seconds?: number;
  queue_timeout_seconds?: number;
}

export async function invokeNative(
  action: 'status' | 'convert',
  request: Partial<ConversionRequest> = {},
): Promise<Record<string, unknown>> {
  if (process.platform !== 'win32') {
    if (action === 'status') return { platform: process.platform, ready: false, reason: 'Windows required' };
    throw new Error('windows_required: Native Word conversion requires Windows.');
  }
  const folder = await mkdtemp(join(tmpdir(), 'cirya-wordpdf-request-'));
  try {
    const requestPath = join(folder, 'request.json');
    await writeFile(requestPath, JSON.stringify({
      ...request,
      action,
      parent_pid: process.pid,
      node_path: process.execPath,
      validator_path: fileURLToPath(new URL('./verify-pdf.js', import.meta.url)),
    }), 'utf8');
    const script = fileURLToPath(new URL('../native/supervisor.ps1', import.meta.url));
    const executable = join(process.env.SystemRoot ?? 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
    // No shell and no document paths embedded in PowerShell source or command text.
    const child = spawn(executable, [
      '-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-ExecutionPolicy', 'Bypass',
      '-File', script, '-RequestPath', requestPath,
    ], { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
    const response = await new Promise<string>((resolve, reject) => {
      let output = '';
      let errors = '';
      let timedOut = false;
      const limit = action === 'status' ? 30 :
        (request.queue_timeout_seconds ?? 600) + (request.timeout_seconds ?? 180) + 90;
      const timer = setTimeout(() => {
        timedOut = true;
        child.kill(); // Only this supervisor; recovery journal protects the next conversion.
      }, limit * 1000);
      child.stdout.setEncoding('utf8').on('data', (data: string) => { output += data; });
      child.stderr.setEncoding('utf8').on('data', (data: string) => { errors = (errors + data).slice(-16000); });
      child.once('error', (error) => { clearTimeout(timer); reject(error); });
      child.once('close', (code) => {
        clearTimeout(timer);
        if (timedOut) reject(new Error('recovery_required: Native supervisor deadline exceeded. Inspect word_pdf_status before retrying.'));
        else if (!output.trim()) reject(new Error(`native_failed: PowerShell exited ${code}. ${errors}`));
        else resolve(output);
      });
    });
    const result = JSON.parse(response.replace(/^\uFEFF/, '')) as Record<string, unknown>;
    if (result.status === 'error') throw new Error(`${result.code ?? 'word_failed'}: ${result.message}`);
    return result;
  } finally {
    await rm(folder, { recursive: true, force: true });
  }
}
