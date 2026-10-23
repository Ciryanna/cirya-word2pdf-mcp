import { test } from 'node:test';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';

test('npm entry performs real stdio handshake and exposes exactly two tools', async () => {
  const client = new Client({ name: 'integration-test', version: '1.0.0' });
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [fileURLToPath(new URL('../bin/cli.mjs', import.meta.url))],
    stderr: 'pipe',
    env: { CIRYA_WORDPDF_TRACE: process.env.CIRYA_WORDPDF_TRACE ?? '0' },
  });
  try {
    await client.connect(transport);
    const { tools } = await client.listTools();
    assert.deepEqual(tools.map((tool) => tool.name).sort(), ['word_pdf_status', 'word_to_pdf']);
    const status = await client.callTool({ name: 'word_pdf_status', arguments: {} });
    assert.ok(!status.isError, JSON.stringify(status));
    assert.equal(status.structuredContent.platform, process.platform);
    const failure = await client.callTool({ name: 'word_to_pdf', arguments: { input_path: 'missing.docx', output_path: 'missing.pdf' } });
    assert.equal(failure.isError, true);
    assert.match(failure.content[0].text, process.platform === 'win32' ? /invalid_input/ : /windows_required/);
    const invalid = await client.callTool({ name: 'word_to_pdf', arguments: { input_path: 'missing.docx', output_path: 'missing.pdf', timeout_seconds: 0 } });
    assert.equal(invalid.isError, true);
  } finally { await client.close(); }
});
