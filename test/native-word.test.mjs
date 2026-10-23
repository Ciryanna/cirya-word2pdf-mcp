import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PDFDocument } from 'pdf-lib';
import { invokeNative } from '../lib/service.js';
import { makeSample, runPowerShell } from './helpers.mjs';

const enabled = process.platform === 'win32' &&
  (process.env.WORD_PDF_RUN_NATIVE === '1' || Boolean(process.env.WORD_PDF_TEST_DOCUMENT));
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
test('real Word COM export preserves source and leaves no extra Word processes', { skip: !enabled }, async () => {
  const folder = await mkdtemp(join(tmpdir(), 'cirya native #&-'));
  const source = process.env.WORD_PDF_TEST_DOCUMENT ?? join(folder, '中文 公式.docx');
  try {
    if (!process.env.WORD_PDF_TEST_DOCUMENT) await makeSample(source);
    const original = hash(await readFile(source));
    const snapshot = async () => JSON.parse(await runPowerShell("ConvertTo-Json -Compress -InputObject @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { Get-Identity $_.Id })"));
    const before = await snapshot();
    const output = join(folder, '中文 output.pdf');
    const result = await invokeNative('convert', { input_path: source, output_path: output, timeout_seconds: 90 });
    assert.equal(result.status, 'success', JSON.stringify(result));
    assert.match(result.api, /^ExportAsFixedFormat2?$/);
    assert.ok(result.pages > 0);
    if (!process.env.WORD_PDF_TEST_DOCUMENT) assert.equal(result.pages, 2);
    assert.equal(result.input_sha256, original);
    assert.equal(hash(await readFile(source)), original);
    assert.equal(result.visual_fidelity_verified, false);
    const pdf = await PDFDocument.load(await readFile(output), { updateMetadata: false });
    assert.match(pdf.getProducer(), /Word/i);
    await assert.rejects(invokeNative('convert', { input_path: source, output_path: output }), /output_exists/);
    assert.equal(hash(await readFile(output)), result.output_sha256);
    const sharedOutput = join(folder, 'concurrent.pdf');
    const competing = await Promise.allSettled([0, 1].map(() => invokeNative('convert', {
      input_path: source, output_path: sharedOutput, timeout_seconds: 90,
    })));
    assert.equal(competing.filter((value) => value.status === 'fulfilled').length, 1);
    const rejected = competing.find((value) => value.status === 'rejected');
    assert.match(String(rejected.reason), /output_exists/);
    assert.equal(hash(await readFile(source)), original);
    assert.deepEqual(await snapshot(), before);
    assert.equal((await invokeNative('status')).pending_record, null);
    console.log(JSON.stringify({ api: result.api, pages: result.pages, word_build: result.word_build, producer: pdf.getProducer() }));
  } finally { await rm(folder, { recursive: true, force: true }); }
});
