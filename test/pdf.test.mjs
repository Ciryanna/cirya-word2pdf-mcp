import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PDFDocument } from 'pdf-lib';
import { validatePdf } from '../lib/verify-pdf.js';

test('PDF validator rejects malformed files and page mismatch', async () => {
  const folder = await mkdtemp(join(tmpdir(), 'wordpdf-validate-'));
  try {
    const path = join(folder, 'stage.pdf');
    await writeFile(path, 'broken PDF');
    await assert.rejects(validatePdf(path, 1), /header/);
    const pdf = await PDFDocument.create();
    pdf.addPage([595, 842]);
    await writeFile(path, await pdf.save());
    await assert.rejects(validatePdf(path, 2), /Word reports 2/);
    assert.equal((await validatePdf(path, 1)).pages, 1);
  } finally { await rm(folder, { recursive: true, force: true }); }
});
