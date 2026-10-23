import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { PDFDocument } from 'pdf-lib';

export async function validatePdf(path: string, expectedPages: number) {
  const bytes = await readFile(path);
  if (bytes.length < 8 || bytes.subarray(0, 5).toString() !== '%PDF-') throw new Error('Invalid PDF header');
  const pdf = await PDFDocument.load(bytes, { updateMetadata: false });
  const pages = pdf.getPages();
  if (!pages.length) throw new Error('PDF has no pages');
  if (pages.length !== expectedPages) throw new Error(`Word reports ${expectedPages} pages; PDF contains ${pages.length}.`);
  return { pages: pages.length, page_sizes_pt: pages.map((page) => [page.getWidth(), page.getHeight()]), bytes: bytes.length };
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  try {
    console.log(JSON.stringify(await validatePdf(process.argv[2]!, Number(process.argv[3]))));
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  }
}
