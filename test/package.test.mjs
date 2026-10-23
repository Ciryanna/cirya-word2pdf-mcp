import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

test('npm tarball includes runtime files and excludes Python and user artifacts', () => {
  const npm = process.env.npm_execpath ?? join(dirname(process.execPath), 'node_modules/npm/bin/npm-cli.js');
  const output = execFileSync(process.execPath, [npm, 'pack', '--dry-run', '--json', '--ignore-scripts'], {
    cwd: fileURLToPath(new URL('../', import.meta.url)), encoding: 'utf8', windowsHide: true,
  });
  const [{ files }] = JSON.parse(output);
  const names = files.map(({ path }) => path);
  for (const required of ['bin/cli.mjs', 'lib/server.js', 'lib/service.js', 'lib/verify-pdf.js',
    'native/worker.ps1', 'native/supervisor.ps1', 'native/common.ps1', 'native/WindowsNative.cs']) {
    assert.ok(names.includes(required), `Missing ${required}`);
  }
  assert.ok(!names.some((name) => /\.(py|pyc|pdf|docx?|docm|whl|env)$/i.test(name)));
  assert.ok(!names.some((name) => /^(src|test|tests|node_modules|\.venv|dist)\//.test(name)));
});
