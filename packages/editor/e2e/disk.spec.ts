// SPDX-License-Identifier: GPL-3.0-or-later
//
// A file changed on disk while open (ED-011 section 3, steps S1 to S3). Playwright cannot drop a file with a
// handle, so the suite stubs the File System Access picker with a handle whose `getFile()` answers from a
// file the test rewrites (`window.__disk`), and the watcher's focus check stands in for its polling.
import { test, expect, type Page } from '@playwright/test';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { fixture, caretInBlock } from './helpers.js';

interface DiskFile { name: string; bytes: Uint8Array; lastModified: number }

declare global {
  interface Window {
    __disk: DiskFile;
    /** The file Save As's picker answers with, one of its own, as a person's choice of a new name gives. */
    __saved: DiskFile;
    /** The name of each file written through a handle. */
    __writes: string[];
    __saveOptions: unknown;
  }
}

const DOCUMENT = '.ProseMirror[aria-label="Document"]';

/**
 * The picker stubbed before the page loads: one handle over a mutable file. `keepWrites`: a write through the handle becomes
 * the file's bytes, its time unchanged (otherwise a write leaves the file as it was).
 */
async function stubPicker(page: Page, name: string, bytes: Uint8Array, options: { keepWrites?: boolean } = {}): Promise<void> {
  await page.addInitScript(({ name, b64, keepWrites }) => {
    const disk = { name, bytes: Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)), lastModified: 1_000_000 };
    const saved = { name: name.replace(/\.md$/, '.docx'), bytes: new Uint8Array(0), lastModified: 5_000_000 };
    window.__disk = disk;
    window.__saved = saved;
    window.__writes = [];
    const handleOf = (file: typeof disk): unknown => ({
      kind: 'file', name: file.name,
      getFile: async () => new File([file.bytes as BlobPart], file.name, { lastModified: file.lastModified }),
      isSameEntry: async () => true,
      queryPermission: async () => 'granted',
      requestPermission: async () => 'granted',
      createWritable: async () => ({
        write: async (data: BlobPart) => {
          window.__writes.push(file.name);
          if (keepWrites) file.bytes = new Uint8Array(await new Blob([data]).arrayBuffer());
        },
        close: async () => undefined,
      }),
    });
    (window as unknown as { showOpenFilePicker: unknown }).showOpenFilePicker = async () => [handleOf(disk)];
    (window as unknown as { showSaveFilePicker: unknown }).showSaveFilePicker = async (options: unknown) => { window.__saveOptions = options; return handleOf(saved); };
  }, { name, b64: Buffer.from(bytes).toString('base64'), keepWrites: options.keepWrites ?? false });
}

/** A package's parts zipped again uncompressed: other bytes, the same document. */
function rezipped(bytes: Uint8Array): Uint8Array {
  return zipSync(unzipSync(bytes), { level: 0 });
}

/** Open through the stubbed picker and wait for the editor. */
async function openPicked(page: Page): Promise<void> {
  await page.goto('/');
  await page.getByRole('button', { name: 'Open…' }).click();
  await page.getByTestId('editor').locator(DOCUMENT).waitFor();
  await page.getByTestId('toolbar').waitFor();
}

/** Another program writes the file: new bytes (or the same), a later time, and the window takes focus. */
async function writeOnDisk(page: Page, bytes?: Uint8Array): Promise<void> {
  await page.evaluate(({ b64 }) => {
    if (b64) window.__disk.bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
    window.__disk.lastModified += 1_000;
    window.dispatchEvent(new Event('focus'));
  }, { b64: bytes ? Buffer.from(bytes).toString('base64') : '' });
}

test('a file written on disk is offered for reload; Keep closes the offer until the next write; Reload opens the new bytes', async ({ page }) => {
  await stubPicker(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  await openPicked(page);
  await expect(page.getByTestId('editor').locator(DOCUMENT)).toContainText('Hello World from Word Online');
  const offer = page.getByTestId('disk-offer');
  await expect(offer).toBeHidden();

  await writeOnDisk(page, await fixture('tables.docx'));
  await expect(offer).toBeVisible();
  await expect(offer).toContainText('hello.docx changed on disk');
  // The counts (ED-011 section 3.4, S2): the one paragraph against tables.docx's 200.
  await expect(page.getByTestId('disk-counts')).toHaveText('On disk, 1 paragraph is changed, 199 added and 0 removed.');
  await page.getByTestId('disk-keep').click();
  await expect(offer).toBeHidden();
  // The same write is not offered twice; a later one is.
  await page.evaluate(() => window.dispatchEvent(new Event('focus')));
  await page.waitForTimeout(200);
  await expect(offer).toBeHidden();
  await writeOnDisk(page);
  await expect(offer).toBeVisible();

  await page.getByTestId('disk-reload').click();
  await expect(page.getByTestId('editor').locator(DOCUMENT)).toContainText('Shadings, alignment, row height');
  await expect(offer).toBeHidden();
  await expect(page.getByTestId('status')).toContainText('Reloaded hello.docx from disk');
  // Reloaded, the watcher starts from the new time: a focus reports nothing, a later write does.
  await page.evaluate(() => window.dispatchEvent(new Event('focus')));
  await page.waitForTimeout(200);
  await expect(offer).toBeHidden();
  // A time that moved over the same bytes is no change (ED-013 section 9): nothing offered.
  await writeOnDisk(page);
  await page.waitForTimeout(200);
  await expect(offer).toBeHidden();
  // The same document in other bytes is a write: the text is the same on both sides.
  await writeOnDisk(page, rezipped(await fixture('tables.docx')));
  await expect(offer).toBeVisible();
  await expect(page.getByTestId('disk-counts')).toHaveText('The text of every paragraph is the same; the change is in formatting, properties or another part.');
});

test("a file's time moved over the editor's own save is no change: nothing offered, and the next Save in place goes ahead (ED-013 section 9)", async ({ page }) => {
  await page.addInitScript(() => { try { localStorage.setItem('docx4j-editor.saveInPlace', 'true'); } catch { /* ignore */ } });
  await stubPicker(page, 'hello.docx', await fixture('HelloWordOnline.docx'), { keepWrites: true });
  await openPicked(page);
  const offer = page.getByTestId('disk-offer');
  await caretInBlock(page, 0);
  await page.keyboard.type('x');
  await page.keyboard.press('Control+s');
  await expect(page.getByTestId('status')).toHaveText('Saved with changes (in place)');
  // Android's provider, measured on the phone: the file's time moves after the editor's write, its bytes the editor's own.
  await writeOnDisk(page);
  await page.waitForTimeout(200);
  await expect(offer).toBeHidden();
  await caretInBlock(page, 0);
  await page.keyboard.type('y');
  await page.keyboard.press('Control+s');
  await expect(page.getByTestId('status')).toHaveText('Saved with changes (in place)');
  expect(await page.evaluate(() => window.__writes)).toEqual(['hello.docx', 'hello.docx']);
  // Another program's write is still one.
  await writeOnDisk(page, await fixture('tables.docx'));
  await expect(offer).toBeVisible();
  await caretInBlock(page, 0);
  await page.keyboard.type('z');
  await page.keyboard.press('Control+s');
  await expect(page.getByTestId('status')).toHaveText('Not saved: the file changed on disk since it was opened. Use Save As.');
});

test('with unsaved changes the offer says so and has no Reload; the disk version opens in a new tab, the edited one untouched', async ({ page }) => {
  await stubPicker(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  await openPicked(page);
  await caretInBlock(page, 0);
  await page.keyboard.type('x');
  await writeOnDisk(page, await fixture('tables.docx'));
  const offer = page.getByTestId('disk-offer');
  await expect(offer).toBeVisible();
  await expect(offer).toContainText('unsaved changes');
  await expect(page.getByTestId('disk-reload')).toHaveCount(0);
  await expect(page.getByTestId('disk-counts')).toHaveText('On disk, 1 paragraph is changed, 199 added and 0 removed.');
  await page.getByTestId('disk-keep').click();
  await expect(offer).toBeHidden();
  await expect(page.getByTestId('editor').locator(DOCUMENT)).toContainText('Hello World from Word Online');

  // The new tab (ED-011 section 3.6, S3): the hand-over record taken by the `#disk` hash. The stub's handle is a
  // plain object of functions, which IndexedDB cannot clone, so the record carries the bytes alone and the new
  // tab opens them without a handle: the path a refused handle takes. A real handle's hand-over is the by-hand check.
  await writeOnDisk(page);
  await expect(offer).toBeVisible();
  const [tab] = await Promise.all([page.context().waitForEvent('page'), page.getByTestId('disk-open-tab').click()]);
  await expect(offer).toBeHidden();
  await tab.getByTestId('editor').locator(DOCUMENT).waitFor();
  await expect(tab.getByTestId('editor').locator(DOCUMENT)).toContainText('Shadings, alignment, row height');
  await expect(tab.getByTestId('name')).toHaveText('hello (on disk).docx');
  await expect(tab.getByTestId('status')).toContainText('Opened the disk version of hello.docx in this tab');
  expect(new URL(tab.url()).hash).toBe('');
  await expect(tab.getByTestId('disk-offer')).toBeHidden();
  // The suffix is the display's alone: a save of any kind drops it, and the download carries the file's own name.
  const [dl] = await Promise.all([tab.waitForEvent('download'), tab.getByTestId('download').click()]);
  expect(dl.suggestedFilename()).toBe('hello.docx');
  await expect(tab.getByTestId('name')).toHaveText('hello.docx');
  // The edited tab keeps its document and its edit.
  await expect(page.getByTestId('editor').locator(DOCUMENT)).toContainText('xHello World from Word Online');
  await expect(page.getByTestId('name')).toHaveText('hello.docx •');
  // A fresh load on a hash whose record is gone (taken once) opens nothing; the hash is cleared.
  const origin = new URL(tab.url()).origin;
  await tab.goto('about:blank');
  await tab.goto(`${origin}/#disk=nothing-here`);
  await tab.getByRole('button', { name: 'Open…' }).waitFor();
  await expect(tab.getByTestId('editor')).toHaveCount(0);
  expect(new URL(tab.url()).hash).toBe('');
});

test('in a room nothing is offered: the room is the document', async ({ page }) => {
  await page.addInitScript(() => { try { localStorage.setItem('docx4j-editor.relay', '/relay'); } catch { /* ignore */ } });
  await stubPicker(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  await openPicked(page);
  await page.getByTestId('share').click();
  await page.getByTestId('share-link').waitFor();
  await page.keyboard.press('Escape');
  await writeOnDisk(page, await fixture('tables.docx'));
  await page.waitForTimeout(500);
  await expect(page.getByTestId('disk-offer')).toBeHidden();
});

// A Markdown source (ED-011 section 3.8, P10). The rows are long enough that the fit changes the widths under any font
// (markdown.spec.ts's WIDE_TABLE, shorter): the fit's writes must not count as an edit, or Reload would never be offered.
const LONG = 'A long description that wraps over several lines of the cell, as a description does when the column is narrow.';
const TABLE = '| ID | Description |\n|---|---|\n' + Array.from({ length: 6 }, (_, i) => `| ${'abcdef'[i]} | ${`${LONG} `.repeat(i + 2).trim()} |`).join('\n') + '\n';
const markdown = (paragraph: string, table = ''): Uint8Array => Buffer.from(`# Notes\n\n${paragraph}\n\n${table}`, 'utf8');
const FIT = "The table's columns were fit to its content.";

test('a Markdown file is watched as the document\'s source: Reload converts it again and nothing is written to it', async ({ page }) => {
  await page.addInitScript(() => { try { localStorage.setItem('docx4j-editor.saveInPlace', 'true'); } catch { /* ignore */ } });
  await stubPicker(page, 'notes.md', markdown('First paragraph.', TABLE));
  await openPicked(page);
  const body = page.getByTestId('editor').locator(DOCUMENT);
  await expect(body).toContainText('First paragraph.');
  await expect(page.getByTestId('status')).toContainText(FIT, { timeout: 20_000 });
  // The fit is the import's work: no local edits, and no Save in place, with the option on, since the document has no file of its own.
  await expect(page.getByTestId('name')).toHaveText('notes.docx');
  await expect(page.getByTestId('save')).toHaveCount(0);

  const offer = page.getByTestId('disk-offer');
  await expect(offer).toBeHidden();
  await writeOnDisk(page, markdown('Second paragraph.', TABLE));
  await expect(offer).toContainText('notes.md changed on disk');
  await expect(page.getByTestId('disk-counts')).toHaveText('On disk, 1 paragraph is changed, 0 added and 0 removed.');
  await page.getByTestId('disk-reload').click();
  await expect(body).toContainText('Second paragraph.');
  await expect(offer).toBeHidden();
  await expect(page.getByTestId('status')).toContainText('Reloaded notes.md from disk.');
  // Converted as an open is: the table fit again, and still no local edits.
  await expect(page.getByTestId('status')).toContainText(FIT, { timeout: 20_000 });
  await expect(page.getByTestId('name')).toHaveText('notes.docx');

  // The watcher starts from the new time: a focus reports nothing, a later write does.
  await page.evaluate(() => window.dispatchEvent(new Event('focus')));
  await page.waitForTimeout(200);
  await expect(offer).toBeHidden();
  await writeOnDisk(page);
  await expect(offer).toBeVisible();
  await expect(page.getByTestId('disk-counts')).toHaveText('The text of every paragraph is the same; the change is in formatting, properties or another part.');

  // Ctrl+S has no file to write: it downloads, and the Markdown file is untouched.
  const [dl] = await Promise.all([page.waitForEvent('download'), page.keyboard.press('Control+s')]);
  expect(dl.suggestedFilename()).toBe('notes.docx');
  expect(await page.evaluate(() => window.__writes)).toEqual([]);
});

test('with local edits a Markdown file\'s disk version is converted in a new tab, which says it does not follow the file', async ({ page }) => {
  await stubPicker(page, 'notes.md', markdown('First paragraph.'));
  await openPicked(page);
  await caretInBlock(page, 1);
  await page.keyboard.type('x');
  await writeOnDisk(page, markdown('Second paragraph.'));
  const offer = page.getByTestId('disk-offer');
  await expect(offer).toContainText('notes.md changed on disk');
  await expect(offer).toContainText('unsaved changes');
  await expect(page.getByTestId('disk-reload')).toHaveCount(0);

  // The stub's handle cannot be cloned, so the record carries the bytes alone: the path of a handle this tab cannot read.
  const [tab] = await Promise.all([page.context().waitForEvent('page'), page.getByTestId('disk-open-tab').click()]);
  await expect(offer).toBeHidden();
  await tab.getByTestId('editor').locator(DOCUMENT).waitFor();
  await expect(tab.getByTestId('editor').locator(DOCUMENT)).toContainText('Second paragraph.');
  await expect(tab.getByTestId('name')).toHaveText('notes (on disk).docx');
  await expect(tab.getByTestId('status')).toContainText('Opened the disk version of notes.md in this tab; this tab will not offer later changes to the file.');
  // The edited tab keeps its document and its edit.
  await expect(page.getByTestId('editor').locator(DOCUMENT)).toContainText('xFirst paragraph.');
  await expect(page.getByTestId('name')).toHaveText('notes.docx •');
});

/** The open session's record as the page keeps it in IndexedDB: its id, its dirty flag and what first made it dirty. */
const sessionRecord = (page: Page, id?: string) => page.evaluate(async (wanted) => {
  const key = wanted ?? localStorage.getItem('docx4j-editor.session');
  if (!key) return undefined;
  const db = await new Promise<IDBDatabase>((resolve, reject) => { const r = indexedDB.open('docx4j-editor'); r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error); });
  const rec = await new Promise<{ id: string; dirty: boolean; firstEdit?: Record<string, unknown> & { at: string; stack?: string[]; input?: { key?: string } } } | undefined>((resolve) => {
    const r = db.transaction('sessions').objectStore('sessions').get(key);
    r.onsuccess = () => resolve(r.result as never);
  });
  db.close();
  return rec ? { id: rec.id, dirty: rec.dirty, firstEdit: rec.firstEdit } : undefined;
}, id);

/** The log of first edits that outlives a session (ED-011 section 9.9). */
const causeLog = (page: Page) => page.evaluate(() => JSON.parse(localStorage.getItem('docx4j-editor.editCauses') ?? '[]') as Array<Record<string, unknown>>);

// Unsaved changes and a disk version wanted in their place (ED-011 section 9.9; Jason, 2026-11-09: "If the message had an
// overwrite option, i would've taken that in the first place and not opened a new tab").
test('with unsaved changes the disk version can replace them, after a confirmation that says they are lost', async ({ page }) => {
  await stubPicker(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  await openPicked(page);
  await caretInBlock(page, 0);
  await page.keyboard.type('x');
  await writeOnDisk(page, await fixture('tables.docx'));
  const offer = page.getByTestId('disk-offer');
  const body = page.getByTestId('editor').locator(DOCUMENT);
  await expect(offer).toContainText('unsaved changes');
  await expect(page.getByTestId('disk-reload')).toHaveCount(0);
  await expect(page.getByTestId('disk-discard')).toHaveText('Reload from disk, discarding your changes');

  // Cancel, and Escape, change nothing.
  const dialog = page.getByRole('dialog', { name: 'Discard your unsaved changes?' });
  await page.getByTestId('disk-discard').click();
  await expect(dialog).toContainText('Reloading hello.docx from disk replaces this document with the file as it is now. Your unsaved changes are lost: Undo cannot bring them back, and recovery will not offer them.');
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(dialog).toBeHidden();
  await page.getByTestId('disk-discard').click();
  await expect(dialog).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(offer).toBeVisible();
  await expect(body).toContainText('xHello World from Word Online');
  await expect(page.getByTestId('name')).toHaveText('hello.docx •');

  const edited = (await sessionRecord(page))!.id;
  await page.getByTestId('disk-discard').click();
  await dialog.getByRole('button', { name: 'Discard and reload' }).click();
  await expect(body).toContainText('Shadings, alignment, row height');
  await expect(body).not.toContainText('xHello');
  await expect(page.getByTestId('name')).toHaveText('hello.docx');
  await expect(offer).toBeHidden();
  await expect(page.getByTestId('status')).toContainText('Reloaded hello.docx from disk; your unsaved changes were discarded.');
  // The edited session went with the reload, so nothing offers it again; the open one is the disk version, unedited.
  expect(await sessionRecord(page, edited)).toBeUndefined();
  expect(await sessionRecord(page)).toMatchObject({ dirty: false });
  // The watcher follows the new time, and a later write offers a plain Reload.
  await writeOnDisk(page, await fixture('HelloWordOnline.docx'));
  await expect(page.getByTestId('disk-reload')).toBeVisible();
  await expect(page.getByTestId('disk-discard')).toHaveCount(0);
});

test('a Markdown source\'s disk version replaces unsaved changes the same way, converted again', async ({ page }) => {
  await stubPicker(page, 'notes.md', markdown('First paragraph.'));
  await openPicked(page);
  await caretInBlock(page, 1);
  await page.keyboard.type('x');
  await writeOnDisk(page, markdown('Second paragraph.'));
  await expect(page.getByTestId('disk-offer')).toContainText('unsaved changes');
  await page.getByTestId('disk-discard').click();
  const dialog = page.getByRole('dialog', { name: 'Discard your unsaved changes?' });
  await expect(dialog).toContainText('Reloading notes.md from disk');
  await dialog.getByRole('button', { name: 'Discard and reload' }).click();
  const body = page.getByTestId('editor').locator(DOCUMENT);
  await expect(body).toContainText('Second paragraph.');
  await expect(body).not.toContainText('xFirst');
  await expect(page.getByTestId('name')).toHaveText('notes.docx');
  await expect(page.getByTestId('status')).toContainText('Reloaded notes.md from disk; your unsaved changes were discarded.');
});

// What first counted as an edit (ED-011 section 9.9): a tab said a document had unsaved changes with nothing edited there,
// and its session was gone before anyone looked, so the cause is kept in the record and in a log that outlives it.
test('what first marked a document changed is kept in its record and in a log; the import\'s table fit is not an edit', async ({ page }) => {
  await stubPicker(page, 'notes.md', markdown('First paragraph.', TABLE));
  await openPicked(page);
  await expect(page.getByTestId('status')).toContainText(FIT, { timeout: 20_000 });
  expect((await sessionRecord(page))?.firstEdit).toBeUndefined();
  expect(await causeLog(page)).toEqual([]);

  await caretInBlock(page, 1);
  await page.keyboard.type('x');
  await expect.poll(async () => (await sessionRecord(page))?.firstEdit?.origin).toBe('y-sync$');
  const { id, firstEdit } = (await sessionRecord(page))!;
  expect(firstEdit).toMatchObject({ local: true, changed: ['body'], blocks: [1], focus: 'editor:Document', visible: true, input: { kind: 'keydown', key: 'character', target: 'editor:Document' } });
  // Deep enough to pass Yjs and ProseMirror to what dispatched, where V8 would stop at ten.
  expect(firstEdit!.stack!.length).toBeGreaterThan(10);
  // The first alone: a later keystroke keeps it.
  await page.keyboard.type('y');
  await page.waitForTimeout(400);
  expect((await sessionRecord(page))!.firstEdit!.at).toBe(firstEdit!.at);
  const log = await causeLog(page);
  expect(log).toHaveLength(1);
  expect(log[0]).toMatchObject({ session: id, name: 'notes.docx', origin: 'y-sync$', blocks: [1], at: firstEdit!.at });
  // Nothing of the document's text is kept.
  expect(JSON.stringify(log)).not.toContain('First paragraph');

  // A save clears it, and the next edit is a first again.
  await Promise.all([page.waitForEvent('download'), page.getByTestId('download').click()]);
  await expect.poll(async () => (await sessionRecord(page))?.dirty).toBe(false);
  expect((await sessionRecord(page))?.firstEdit).toBeUndefined();
  await caretInBlock(page, 0);
  await page.keyboard.press('Enter');
  await expect.poll(async () => (await sessionRecord(page))?.firstEdit?.input?.key).toBe('Enter');
  expect(await causeLog(page)).toHaveLength(2);
});

test('Save As ends a Markdown source: the .docx it wrote is watched from then on, the .md is not, and an offer about the .md closes', async ({ page }) => {
  await stubPicker(page, 'notes.md', markdown('First paragraph.'));
  await openPicked(page);
  const offer = page.getByTestId('disk-offer');
  await writeOnDisk(page, markdown('Second paragraph.'));
  await expect(offer).toContainText('notes.md changed on disk');

  await page.getByTestId('save-as').click();
  await expect(page.getByTestId('status')).toContainText('Saved');
  await expect(offer).toBeHidden();
  expect(await page.evaluate(() => window.__writes)).toEqual(['notes.docx']);
  // Save As writes a docx's bytes, so Markdown is no type it offers.
  expect((await page.evaluate(() => JSON.stringify(window.__saveOptions))).toLowerCase()).not.toContain('markdown');

  await writeOnDisk(page, markdown('Third paragraph.'));
  await page.waitForTimeout(200);
  await expect(offer).toBeHidden();
  await page.evaluate(() => { window.__saved.lastModified += 1_000; window.dispatchEvent(new Event('focus')); });
  await expect(offer).toContainText('notes.docx changed on disk');
});

// The caret and the view kept across a reload (ED-011 section 3.5). A document of plain numbered clauses, long enough to
// scroll, built here so that a disk version can differ from it by what each test needs.
const clause = (n: number): string => `Clause ${n} of the rules, which says something of its own about number ${n * 7 + 3}.`;
const clauses = (from: number, to: number): string[] => Array.from({ length: to - from + 1 }, (_, i) => clause(from + i));
const PAPER = clauses(1, 240);

function docx(paragraphs: readonly string[]): Uint8Array {
  const body = paragraphs.map((p) => `<w:p><w:r><w:t xml:space="preserve">${p}</w:t></w:r></w:p>`).join('');
  return zipSync({
    '[Content_Types].xml': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    '_rels/.rels': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/document.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`),
  });
}

/** Where the caret is: its paragraph's text, its offset there, and its height in the page area. */
const caretReport = (page: Page) => page.evaluate(() => {
  const { view } = (window as unknown as { docx4jEditor: { view: { state: { selection: { $head: { parent: { textContent: string }; parentOffset: number }; head: number } }; coordsAtPos(pos: number): { top: number } } } }).docx4jEditor;
  const { $head, head } = view.state.selection;
  const area = document.querySelector('.page-area')!;
  return { text: $head.parent.textContent, offset: $head.parentOffset, y: view.coordsAtPos(head).top - area.getBoundingClientRect().top };
});

/** What is at the top of the page area, by the paragraph drawn there, and how far down it stands. */
const topOfView = (page: Page) => page.evaluate((selector) => {
  const area = document.querySelector('.page-area')!;
  const box = area.getBoundingClientRect();
  const first = [...document.querySelectorAll(`${selector} p`)].find((p) => p.getBoundingClientRect().bottom > box.top + 1)!;
  return { text: first.textContent ?? '', y: first.getBoundingClientRect().top - box.top, scrollTop: area.scrollTop };
}, DOCUMENT);

/** Scroll a clause to a height in the page area, its text found in the view. */
async function scrollClauseTo(page: Page, text: string, y: number): Promise<void> {
  await page.evaluate(([selector, wanted, height]) => {
    const area = document.querySelector('.page-area')!;
    const p = [...document.querySelectorAll(`${selector as string} p`)].find((e) => e.textContent === wanted)!;
    area.scrollTop += p.getBoundingClientRect().top - area.getBoundingClientRect().top - (height as number);
  }, [DOCUMENT, text, y] as const);
}

/** Put the caret in a clause, by its text, at an offset. */
async function caretInClause(page: Page, text: string, offset: number): Promise<void> {
  await page.evaluate(([wanted, off]) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { descendants(f: (n: { type: { name: string }; textContent: string }, pos: number) => boolean): void } }; caretAt(pos: number): void } }).docx4jEditor;
    let at = -1;
    v.state.doc.descendants((n, pos) => { if (n.type.name === 'paragraph' && n.textContent === wanted) at = pos; return at < 0; });
    v.caretAt(at + 1 + (off as number));
  }, [text, offset] as const);
}

async function openPaper(page: Page): Promise<void> {
  await stubPicker(page, 'rules.docx', docx(PAPER));
  await openPicked(page);
  await expect(page.getByTestId('editor').locator(DOCUMENT)).toContainText(clause(240));
}

/** Write a version on disk and reload it, and wait for the frame the place is put back in. */
async function reloadWith(page: Page, paragraphs: readonly string[]): Promise<void> {
  await writeOnDisk(page, docx(paragraphs));
  await page.getByTestId('disk-reload').click();
  await expect(page.getByTestId('status')).toContainText('Reloaded rules.docx from disk');
  await page.evaluate(() => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done))));
}

test('a reload keeps the caret in its text and on screen at its height, with paragraphs added above it', async ({ page }) => {
  await openPaper(page);
  await caretInClause(page, clause(120), 11);
  await scrollClauseTo(page, clause(120), 180);
  const before = await caretReport(page);
  await reloadWith(page, ['A new opening.', 'And another.', ...PAPER]);
  const after = await caretReport(page);
  expect(after).toMatchObject({ text: clause(120), offset: 11 });
  expect(Math.abs(after.y - before.y)).toBeLessThan(2);
  await expect(page.getByTestId('editor').locator(DOCUMENT)).toContainText('A new opening.');
  // Nothing is said, and a caret put back is no edit: the next write is offered a Reload again.
  await expect(page.getByTestId('status')).toHaveText('Reloaded rules.docx from disk.');
  await expect(page.getByTestId('name')).toHaveText('rules.docx');
});

test('a caret on screen keeps its height when what stands above it on screen changes, the top line giving way', async ({ page }) => {
  await openPaper(page);
  await caretInClause(page, clause(120), 11);
  await scrollClauseTo(page, clause(120), 180);
  const top = await topOfView(page);
  const paper = PAPER.slice();
  paper.splice(119, 0, 'Added just above.', 'And again.', 'And a third time.');
  await reloadWith(page, paper);
  const after = await caretReport(page);
  expect(after).toMatchObject({ text: clause(120), offset: 11 });
  expect(Math.abs(after.y - 180)).toBeLessThan(2);
  expect((await topOfView(page)).text).not.toBe(top.text);
});

test('a caret off screen comes back where it was, and the view stays on what was being read', async ({ page }) => {
  await openPaper(page);
  await caretInClause(page, clause(5), 3);
  await scrollClauseTo(page, clause(150), 0);
  const before = await topOfView(page);
  expect(before.text).toBe(clause(150));
  await reloadWith(page, ['A new opening.', ...PAPER]);
  expect(await caretReport(page)).toMatchObject({ text: clause(5), offset: 3 });
  const after = await topOfView(page);
  expect(after.text).toBe(clause(150));
  expect(Math.abs(after.y - before.y)).toBeLessThan(2);
  expect(after.scrollTop).toBeGreaterThan(before.scrollTop);
});

test('a document read without a click keeps its view: the caret at the start is not where the person is', async ({ page }) => {
  await openPaper(page);
  await scrollClauseTo(page, clause(90), 0);
  await reloadWith(page, PAPER.filter((p) => p !== clause(2)));
  expect((await topOfView(page)).text).toBe(clause(90));
  // The first paragraph is the caret's and it stands: the caret keeps it.
  expect(await caretReport(page)).toMatchObject({ text: clause(1), offset: 0 });
});

test('the caret\'s paragraph removed: the caret is where it stood, and the status line says what went', async ({ page }) => {
  await openPaper(page);
  await caretInClause(page, clause(120), 11);
  await scrollClauseTo(page, clause(120), 180);
  const before = await caretReport(page);
  await reloadWith(page, PAPER.filter((p) => p !== clause(120)));
  const after = await caretReport(page);
  expect(after).toMatchObject({ text: clause(121), offset: 0 });
  expect(Math.abs(after.y - before.y)).toBeLessThan(2);
  await expect(page.getByTestId('status')).toHaveText(/^Reloaded rules\.docx from disk\. The paragraph the caret was in, beginning .Clause 120 of the rules, which says som.*is gone; the caret is where it was\.$/);
});

test('everything from the caret to the end removed: the caret is at the end of what is left', async ({ page }) => {
  await openPaper(page);
  await caretInClause(page, clause(200), 4);
  await scrollClauseTo(page, clause(200), 100);
  await reloadWith(page, PAPER.slice(0, 150));
  expect(await caretReport(page)).toMatchObject({ text: clause(150), offset: clause(150).length });
  await expect(page.getByTestId('status')).toContainText('is gone; the caret is where it was.');
});

test('the caret\'s paragraph moved: the caret follows it, and nothing is said', async ({ page }) => {
  await openPaper(page);
  await caretInClause(page, clause(60), 9);
  await scrollClauseTo(page, clause(60), 150);
  const moved = PAPER.filter((p) => p !== clause(60));
  moved.splice(199, 0, clause(60));
  await reloadWith(page, moved);
  const after = await caretReport(page);
  expect(after).toMatchObject({ text: clause(60), offset: 9 });
  expect(Math.abs(after.y - 150)).toBeLessThan(2);
  await expect(page.getByTestId('status')).toHaveText('Reloaded rules.docx from disk.');
});

test('a file with nothing in common: the caret and the view at the start, and the status line says why', async ({ page }) => {
  await openPaper(page);
  await caretInClause(page, clause(120), 11);
  await scrollClauseTo(page, clause(120), 180);
  await reloadWith(page, Array.from({ length: 200 }, (_, i) => `Another document altogether, line ${i}.`));
  expect(await caretReport(page)).toMatchObject({ text: 'Another document altogether, line 0.', offset: 0 });
  expect((await topOfView(page)).scrollTop).toBe(0);
  await expect(page.getByTestId('status')).toContainText('This version has nothing in common with the one that was open, so the caret is at the start.');
});

test('the place is held while the page settles, until the person moves', async ({ page }) => {
  await openPaper(page);
  await caretInClause(page, clause(120), 11);
  await scrollClauseTo(page, clause(120), 180);
  await reloadWith(page, ['A new opening.', ...PAPER]);
  // A late reflow on screen above the caret, as an embedded font's arrival or a Markdown table's fit is. The browser's own
  // scroll anchoring holds the first line in view, and so only what is above the view: this one is the hold's.
  const growAboveCaret = async (height: number): Promise<void> => {
    await page.evaluate(([selector, wanted, px]) => {
      const p = [...document.querySelectorAll(`${selector as string} p`)].find((e) => e.textContent === wanted)!;
      const style = document.createElement('style');
      style.textContent = `${selector as string} > :nth-child(${[...p.parentElement!.children].indexOf(p) + 1}) { padding-block-start: ${px as number}px !important; }`;
      document.head.append(style);
    }, [DOCUMENT, clause(119), height] as const);
    await page.evaluate(() => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done))));
  };
  const scrolled = (await topOfView(page)).scrollTop;
  await growAboveCaret(300);
  expect(Math.abs((await caretReport(page)).y - 180)).toBeLessThan(2);
  expect((await topOfView(page)).scrollTop - scrolled).toBeGreaterThan(250);
  // The person takes over: the next reflow is the browser's to place, and nothing pulls the page back.
  await page.getByTestId('editor').hover();
  await page.mouse.wheel(0, 120);
  await expect.poll(async () => (await caretReport(page)).y).toBeLessThan(100);
  const moved = (await caretReport(page)).y;
  await growAboveCaret(700);
  expect((await caretReport(page)).y - moved).toBeGreaterThan(300);
});

test('a reload at another zoom keeps the place too', async ({ page }) => {
  await page.addInitScript(() => { try { localStorage.setItem('docx4j-editor.zoom', '1.5'); } catch { /* ignore */ } });
  await openPaper(page);
  await caretInClause(page, clause(120), 11);
  await scrollClauseTo(page, clause(120), 180);
  const before = await caretReport(page);
  await reloadWith(page, ['A new opening.', ...PAPER]);
  const after = await caretReport(page);
  expect(after).toMatchObject({ text: clause(120), offset: 11 });
  expect(Math.abs(after.y - before.y)).toBeLessThan(2);
});

test('a reload keeps which view of the tracked changes was on', async ({ page }) => {
  await openPaper(page);
  await page.evaluate(() => { (window as unknown as { docx4jEditor: { markupView: string } }).docx4jEditor.markupView = 'accepted'; });
  await reloadWith(page, ['A new opening.', ...PAPER]);
  expect(await page.evaluate(() => (window as unknown as { docx4jEditor: { markupView: string } }).docx4jEditor.markupView)).toBe('accepted');
});

test('a Markdown source reloaded keeps the caret in its text, through the conversion and the table fit', async ({ page }) => {
  const words = (n: number): string => clauses(1, n).join('\n\n');
  await stubPicker(page, 'notes.md', markdown(words(120), TABLE));
  await openPicked(page);
  await expect(page.getByTestId('status')).toContainText(FIT, { timeout: 20_000 });
  await caretInClause(page, clause(100), 6);
  await scrollClauseTo(page, clause(100), 200);
  await writeOnDisk(page, Buffer.from(`# Notes\n\nA new opening.\n\n${words(120)}\n\n${TABLE}`, 'utf8'));
  await page.getByTestId('disk-reload').click();
  // In this order: the first open's fit is still in the status line until the reload replaces it.
  await expect(page.getByTestId('status')).toContainText('Reloaded notes.md from disk.');
  await expect(page.getByTestId('status')).toContainText(FIT, { timeout: 20_000 });
  const after = await caretReport(page);
  expect(after).toMatchObject({ text: clause(100), offset: 6 });
  expect(Math.abs(after.y - 200)).toBeLessThan(2);
  await expect(page.getByTestId('name')).toHaveText('notes.docx');
});

test('a late message of the open adds to what the open said: the embedded fonts\' line leaves the note standing', async ({ page }) => {
  await openPaper(page);
  await caretInClause(page, clause(120), 11);
  await writeOnDisk(page, await fixture('fonts-embedded.docx'));
  await page.getByTestId('disk-reload').click();
  await expect(page.getByTestId('status')).toContainText('Embedded fonts loaded');
  await expect(page.getByTestId('status')).toContainText('nothing in common');
  await expect(page.getByTestId('status')).toContainText('Reloaded rules.docx from disk.');
});

test('a reload that fails carries nothing to the next document opened', async ({ page }) => {
  await openPaper(page);
  await caretInClause(page, clause(120), 11);
  await scrollClauseTo(page, clause(120), 180);
  await writeOnDisk(page, new Uint8Array([1, 2, 3, 4]));
  await page.getByTestId('disk-reload').click();
  await expect(page.getByTestId('status')).toContainText('rules.docx');
  await expect(page.getByTestId('editor')).toHaveCount(0);
  await page.evaluate(({ b64 }) => { window.__disk.bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)); }, { b64: Buffer.from(docx(PAPER)).toString('base64') });
  await page.getByRole('button', { name: 'Open…' }).click();
  await page.getByTestId('editor').locator(DOCUMENT).waitFor();
  await page.evaluate(() => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done))));
  expect(await caretReport(page)).toMatchObject({ text: clause(1), offset: 0 });
  expect((await topOfView(page)).scrollTop).toBe(0);
});

test('a row selected comes back as a caret in its first cell, where a selection\'s far end would be past it', async ({ page }) => {
  await stubPicker(page, 'tables.docx', await fixture('tables.docx'));
  await openPicked(page);
  const row = page.getByTestId('editor').locator(`${DOCUMENT} table`).first().locator('tr').nth(1);
  const texts = await row.locator('td, th').allTextContents();
  expect(new Set(texts).size).toBeGreaterThan(1);
  const box = (await row.boundingBox())!;
  await page.mouse.click(box.x - 12, box.y + box.height / 2);
  await expect.poll(() => page.evaluate(() => (window as unknown as { docx4jEditor: { state: { selection: { constructor: { name: string } } } } }).docx4jEditor.state.selection.constructor.name)).toContain('CellSelection');
  // The same document in other bytes: a time moved over the same bytes is no change (ED-013 section 9).
  await writeOnDisk(page, rezipped(await fixture('tables.docx')));
  await page.getByTestId('disk-reload').click();
  await expect(page.getByTestId('status')).toContainText('Reloaded tables.docx from disk');
  await page.evaluate(() => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done))));
  expect(await caretReport(page)).toMatchObject({ text: texts[0], offset: 0 });
});

// The same place carried into the disk version's tab (ED-011 sections 3.5 and 3.6): the edited tab's caret and view,
// handed over with the bytes, so the two versions can be compared by eye from the same line.
test('the disk version in a new tab opens at the edited tab\'s caret and view', async ({ page }) => {
  await openPaper(page);
  await caretInClause(page, clause(120), 11);
  await page.keyboard.type('x');
  await scrollClauseTo(page, `${clause(120).slice(0, 11)}x${clause(120).slice(11)}`, 180);
  await writeOnDisk(page, docx(['A new opening.', 'And another.', ...PAPER]));
  await expect(page.getByTestId('disk-offer')).toContainText('unsaved changes');
  const [tab] = await Promise.all([page.context().waitForEvent('page'), page.getByTestId('disk-open-tab').click()]);
  await tab.getByTestId('editor').locator(DOCUMENT).waitFor();
  await expect(tab.getByTestId('status')).toContainText('Opened the disk version of rules.docx in this tab');
  await tab.evaluate(() => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done))));
  // The edited paragraph is the disk's with one letter more: the same paragraph, the caret at the edit.
  const after = await caretReport(tab);
  expect(after).toMatchObject({ text: clause(120), offset: 11 });
  expect(Math.abs(after.y - 180)).toBeLessThan(2);
  await expect(tab.getByTestId('status')).not.toContainText('is gone');
  await expect(tab.getByTestId('name')).toHaveText('rules (on disk).docx');
  // The tab that handed it over is as it was.
  await expect(page.getByTestId('name')).toHaveText('rules.docx •');
});

test('in the new tab a paragraph that is only in the edited version is said to be gone, with the caret where it stood', async ({ page }) => {
  await openPaper(page);
  await caretInClause(page, clause(120), clause(120).length);
  await page.keyboard.press('Enter');
  await page.keyboard.type('A paragraph written here and saved nowhere yet.');
  // The opened text in other bytes (two `docx(PAPER)` within fflate's two-second clock are the same bytes, no change).
  await writeOnDisk(page, rezipped(docx(PAPER)));
  await expect(page.getByTestId('disk-offer')).toContainText('unsaved changes');
  const [tab] = await Promise.all([page.context().waitForEvent('page'), page.getByTestId('disk-open-tab').click()]);
  await tab.getByTestId('editor').locator(DOCUMENT).waitFor();
  await expect(tab.getByTestId('status')).toContainText('A paragraph written here and saved nowhe…');
  await expect(tab.getByTestId('status')).toContainText('is gone; the caret is where it was.');
  expect(await caretReport(tab)).toMatchObject({ text: clause(121), offset: 0 });
});

test('a Markdown source\'s disk version in a new tab keeps the place too', async ({ page }) => {
  const words = clauses(1, 120).join('\n\n');
  await stubPicker(page, 'notes.md', markdown(words));
  await openPicked(page);
  await caretInClause(page, clause(100), 6);
  await page.keyboard.type('x');
  await writeOnDisk(page, Buffer.from(`# Notes\n\nA new opening.\n\n${words}\n`, 'utf8'));
  await expect(page.getByTestId('disk-offer')).toContainText('unsaved changes');
  const [tab] = await Promise.all([page.context().waitForEvent('page'), page.getByTestId('disk-open-tab').click()]);
  await tab.getByTestId('editor').locator(DOCUMENT).waitFor();
  await tab.evaluate(() => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done))));
  expect(await caretReport(tab)).toMatchObject({ text: clause(100), offset: 6 });
});

test('a disk version that will not open carries nothing to the next document opened in its tab', async ({ page }) => {
  await openPaper(page);
  await caretInClause(page, clause(120), 11);
  await page.keyboard.type('x');
  await writeOnDisk(page, new Uint8Array([1, 2, 3, 4]));
  await expect(page.getByTestId('disk-offer')).toContainText('unsaved changes');
  const [tab] = await Promise.all([page.context().waitForEvent('page'), page.getByTestId('disk-open-tab').click()]);
  await expect(tab.getByTestId('status')).not.toHaveText('');
  await expect(tab.getByTestId('editor')).toHaveCount(0);
  await tab.locator('input[type=file]').first().setInputFiles({ name: 'rules.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(docx(PAPER)) });
  await tab.getByTestId('editor').locator(DOCUMENT).waitFor();
  await tab.evaluate(() => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done))));
  expect(await caretReport(tab)).toMatchObject({ text: clause(1), offset: 0 });
  await expect(tab.getByTestId('status')).not.toContainText('is gone');
});
