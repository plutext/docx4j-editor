// SPDX-License-Identifier: GPL-3.0-or-later
//
// A save as Flat OPC (ED-005 section 12.122): a file named `.xml` is written as Word's "Word XML Document", by Save As, by the
// palette's Save as Flat OPC XML, and by every Save after it. The File System Access pickers are stubbed, as in disk.spec.ts,
// with handles that keep what is written to them.
import { test, expect, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { fixture, caretAtBlockEnd } from './helpers.js';

/** What Save As's picker is opened with. */
interface SaveOptions { suggestedName?: string; types?: { accept: Record<string, string[]> }[] }

declare global {
  interface Window {
    /** What each handle was last written with, by name. */
    __written: Record<string, number[]>;
    /** The name Save As's picker answers with, as the person types it; and the options it was opened with. */
    __saveName: string;
    __flatSaveOptions: SaveOptions | undefined;
  }
}

const DOCUMENT = '.ProseMirror[aria-label="Document"]';

/** The pickers stubbed before the page loads, over `name`; with `fsa` false the browser has none, as Firefox. */
async function stubPickers(page: Page, name: string, bytes: Uint8Array, fsa = true): Promise<void> {
  await page.addInitScript(({ name, b64, fsa }) => {
    try { localStorage.setItem('docx4j-editor.saveInPlace', 'true'); } catch { /* ignore */ }
    const w = window as unknown as Record<string, unknown>;
    if (!fsa) { delete w['showOpenFilePicker']; delete w['showSaveFilePicker']; return; }
    window.__written = {};
    window.__saveName = 'report.xml';
    const handleOf = (file: { name: string; bytes: Uint8Array; lastModified: number }): unknown => ({
      kind: 'file', name: file.name,
      getFile: async () => new File([file.bytes as BlobPart], file.name, { lastModified: file.lastModified }),
      isSameEntry: async () => true,
      queryPermission: async () => 'granted',
      requestPermission: async () => 'granted',
      createWritable: async () => {
        const chunks: Uint8Array[] = [];
        return {
          write: async (data: BlobPart) => { chunks.push(new Uint8Array(await new Blob([data]).arrayBuffer())); },
          close: async () => { file.bytes = new Uint8Array(chunks.flatMap((c) => [...c])); file.lastModified += 1_000; window.__written[file.name] = [...file.bytes]; },
        };
      },
    });
    const opened = { name, bytes: Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)), lastModified: 1_000_000 };
    w['showOpenFilePicker'] = async () => [handleOf(opened)];
    w['showSaveFilePicker'] = async (options: SaveOptions) => { window.__flatSaveOptions = options; return handleOf({ name: window.__saveName, bytes: new Uint8Array(0), lastModified: 5_000_000 }); };
  }, { name, b64: Buffer.from(bytes).toString('base64'), fsa });
}

const written = async (page: Page, name: string): Promise<string> => Buffer.from(await page.evaluate((n) => window.__written[n] ?? [], name)).toString('utf8');

/** A Flat OPC document's text through the engine, as the editor reads a document. */
async function textOfFlatOpc(flat: string): Promise<string> {
  const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
  const { EditorDocument } = await import('@docx4j/editor-model');
  return (await EditorDocument.open(await WordprocessingMLPackage.load(flat))).doc.textContent;
}

const FLAT_START = '<?xml version="1.0" standalone="yes"?><?mso-application progid="Word.Document"?><pkg:package ';

async function runPalette(page: Page, id: string): Promise<void> {
  await page.getByTestId('palette-button').click();
  await page.getByTestId('palette').locator(`li[data-action="${id}"]`).click();
}

test('Save As to a name ending .xml writes Flat OPC; Save keeps to it, with the edit; the download is Flat OPC too', async ({ page }) => {
  await stubPickers(page, 'HelloWordOnline.docx', await fixture('HelloWordOnline.docx'));
  await page.goto('/');
  await page.getByRole('button', { name: 'Open…' }).click();
  await page.getByTestId('editor').locator(DOCUMENT).waitFor();
  const status = page.getByTestId('status');

  await page.getByTestId('save-as').click();
  await expect(status).toContainText('Saved');
  // Save As's picker offers Flat OPC after the document types.
  const types = await page.evaluate(() => window.__flatSaveOptions?.types?.map((t) => Object.values(t.accept).flat().join(' ')));
  expect(types).toEqual(['.docx .docm', '.dotx .dotm', '.xml']);
  const first = await written(page, 'report.xml');
  expect(first.startsWith(FLAT_START)).toBe(true);

  await caretAtBlockEnd(page, 0);
  await page.keyboard.type(' Then saved in place.');
  await page.getByTestId('save').click();
  await expect(status).toContainText('in place');
  const second = await written(page, 'report.xml');
  expect(second.startsWith(FLAT_START)).toBe(true);
  expect(await textOfFlatOpc(second)).toContain('Then saved in place.');

  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByTestId('download').click()]);
  expect(dl.suggestedFilename()).toBe('report.xml');
  expect((await readFile((await dl.path())!, 'utf8')).startsWith(FLAT_START)).toBe(true);
});

test('the palette\'s Save as Flat OPC XML suggests the name as .xml with its type first; a name changed back to .docx is a zip', async ({ page }) => {
  await stubPickers(page, 'HelloWordOnline.docx', await fixture('HelloWordOnline.docx'));
  await page.goto('/');
  await page.getByRole('button', { name: 'Open…' }).click();
  await page.getByTestId('editor').locator(DOCUMENT).waitFor();

  await page.evaluate(() => { window.__saveName = 'kept.docx'; });
  await runPalette(page, 'save-flat-opc');
  await expect(page.getByTestId('status')).toContainText('Saved');
  const options = await page.evaluate(() => window.__flatSaveOptions);
  expect(options?.suggestedName).toBe('HelloWordOnline.xml');
  expect(Object.values(options!.types![0]!.accept).flat()).toEqual(['.xml']);
  const bytes = await page.evaluate(() => window.__written['kept.docx']!);
  expect(String.fromCharCode(bytes[0]!, bytes[1]!)).toBe('PK');
});

test('without a picker, Save as Flat OPC XML downloads the document as name.xml', async ({ page }) => {
  await stubPickers(page, 'HelloWordOnline.docx', new Uint8Array(0), false);
  await page.goto('/');
  await page.getByTestId('open').setInputFiles({ name: 'HelloWordOnline.docx', mimeType: 'application/octet-stream', buffer: Buffer.from(await fixture('HelloWordOnline.docx')) });
  await page.getByTestId('editor').locator(DOCUMENT).waitFor();
  const [dl] = await Promise.all([page.waitForEvent('download'), runPalette(page, 'save-flat-opc')]);
  expect(dl.suggestedFilename()).toBe('HelloWordOnline.xml');
  const flat = await readFile((await dl.path())!, 'utf8');
  expect(flat.startsWith(FLAT_START)).toBe(true);
  expect(await textOfFlatOpc(flat)).toContain('Hello');
});
