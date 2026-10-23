import { readFileSync } from 'node:fs';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { invokeNative } from './service.js';

const { version } = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8')) as { version: string };
const server = new McpServer({ name: 'cirya-word2pdf-mcp', version });
function response(result: Record<string, unknown>) {
  return { content: [{ type: 'text' as const, text: JSON.stringify(result) }], structuredContent: result };
}
function failure(error: unknown) {
  return { content: [{ type: 'text' as const, text: error instanceof Error ? error.message : String(error) }], isError: true };
}

server.registerTool('word_pdf_status', {
  description: 'Inspect local Windows/session/recovery status without opening Word.',
  inputSchema: {},
  annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false },
}, async () => {
  try { return response(await invokeNative('status')); } catch (error) { return failure(error); }
});
server.registerTool('word_to_pdf', {
  description: 'Convert a saved local .doc/.docx/.docm using native desktop Microsoft Word COM. Source bytes stay unchanged. Calls serialize within the Windows session. Output folder must exist. No alternate renderer, font repair or OLE refresh. Default excludes revision/comment markup. Visual acceptance is still required for formulas and objects.',
  inputSchema: {
    input_path: z.string().min(1),
    output_path: z.string().min(1),
    overwrite: z.boolean().default(false),
    include_markup: z.boolean().default(false),
    timeout_seconds: z.number().int().min(10).max(1800).default(180),
    queue_timeout_seconds: z.number().int().min(1).max(3600).default(600),
  },
  annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false },
}, async (request) => {
  // An MCP cancellation must not abandon process cleanup or release the mutex early.
  try { return response(await invokeNative('convert', request)); } catch (error) { return failure(error); }
});

await server.connect(new StdioServerTransport());
