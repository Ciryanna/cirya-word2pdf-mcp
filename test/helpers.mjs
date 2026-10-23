import { spawn } from 'node:child_process';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync, strToU8 } from 'fflate';

export const common = fileURLToPath(new URL('../native/common.ps1', import.meta.url));
export const psQuote = (value) => "'" + value.replaceAll("'", "''") + "'";
export async function runPowerShell(script, args = []) {
  const folder = await mkdtemp(join(tmpdir(), 'cirya-native-test-'));
  try {
    const path = join(folder, 'test.ps1');
    await writeFile(path, `. ${psQuote(common)}\n${script}`, 'utf8');
    return await new Promise((resolve, reject) => {
      const child = spawn(join(process.env.SystemRoot ?? 'C:\\Windows', 'System32/WindowsPowerShell/v1.0/powershell.exe'),
        ['-NoProfile', '-NonInteractive', '-STA', '-ExecutionPolicy', 'Bypass', '-File', path, ...args],
        { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
      let output = '', errors = '';
      const timer = setTimeout(() => child.kill(), 40000);
      child.stdout.setEncoding('utf8').on('data', (data) => { output += data; });
      child.stderr.setEncoding('utf8').on('data', (data) => { errors += data; });
      child.once('error', (error) => { clearTimeout(timer); reject(error); });
      child.once('close', (code) => {
        clearTimeout(timer);
        if (code !== 0) reject(new Error(`PowerShell ${code}: ${errors}`)); else resolve(output.trim());
      });
    });
  } finally { await rm(folder, { recursive: true, force: true }); }
}
export async function makeSample(path) {
  const entries = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/document.xml': '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p><w:r><w:t>Native Word COM / 中文公式测试</w:t></w:r></w:p><w:p><m:oMath><m:f><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f><m:sSup><m:e><m:r><m:t>v</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></w:p><w:p><w:r><w:br w:type="page"/><w:t>Second page.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>',
  };
  await writeFile(path, zipSync(Object.fromEntries(Object.entries(entries).map(([name, value]) => [name, strToU8(value)]))));
}
