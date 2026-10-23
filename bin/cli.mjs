#!/usr/bin/env node
import { readFileSync } from 'node:fs';

const { version } = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'));
if (process.argv.includes('--version')) {
  console.log(version);
} else if (process.argv.includes('--help')) {
  console.log('cirya-word2pdf-mcp: local stdio MCP for Windows desktop Word.\nTools: word_to_pdf, word_pdf_status\nRequires Node.js 22+, Windows PowerShell 5.1 and desktop Microsoft Word.');
} else {
  await import('../lib/server.js');
}
