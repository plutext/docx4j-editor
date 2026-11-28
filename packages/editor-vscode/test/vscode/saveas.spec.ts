// SPDX-License-Identifier: GPL-3.0-or-later
//
// New Word Document and Save As (ED-014 section 4.3, P17, P25) in a real VS Code: a new document saved through VS Code's
// own dialog and opened afresh with the caret where it was; Save As to a new name, with an edit during the export carried
// to the new editor; and the refusals: a file open in another editor, a name of no Word type. Then `.xml`, Flat OPC, in
// VS Code's text editor.
import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { VsCode, bodyText, caretAtEndOf, exists, insertInPage, sleep } from './harness.js';
import { mainPart } from './zip.js';

test.describe.configure({ mode: 'serial' });

let vs: VsCode;

test.beforeAll(async () => {
  vs = await VsCode.launch({
    profile: 'saveas',
    files: { 'source.docx': 'HelloWordOnline.docx', 'open-one.docx': 'HelloWordOnline.docx', 'open-two.docx': 'fonts-hebrew-no-cs.docx', 'flat.docx': 'HelloWordOnline.docx', 'large.docx': 'large/anon-250.docx' },
  });
});

test.afterAll(async () => { await vs?.close(); });

test.afterEach(async () => {
  expect(vs.problems, 'content security policy refusals').toEqual([]);
  await vs.closeAll();
});

/**
 * A Save As settled: the old editor gone and the target's tab there, VS Code's prompt about the old one answered where it
 * asks (before 1.147, ED-014 section 9.10). Whether it asked is the case's annotation.
 */
async function settled(old: string, target: string, timeout = 60_000): Promise<void> {
  const prompted = await vs.afterSaveAs(old, async () => (await vs.tab(old)) === undefined && (await vs.tab(target)) !== undefined, timeout);
  test.info().annotations.push({ type: 'old editor prompt', description: `${prompted} (VS Code ${vs.version})` });
}

const head = (frame: import('@playwright/test').Frame): Promise<number> => frame.evaluate(() => (window as unknown as { docx4jEditor: { state: { selection: { head: number } } } }).docx4jEditor.state.selection.head);

test('New Word Document: the page size asked, typing, Save through VS Code\'s dialog; the file written and opened afresh, the caret kept', async () => {
  await vs.exec('docx4j.newDocument');
  // The untitled editor's page asks the page size, as New does in a browser.
  let page: import('@playwright/test').Frame | undefined;
  await expect.poll(async () => {
    for (const f of await vs.frames()) if (await f.locator('[role="dialog"]').count() > 0) { page = f; return true; }
    return false;
  }).toBe(true);
  await page!.locator('[role="dialog"] button[type="submit"]').click();
  const frame = await vs.frameFor('Document1.docx');
  expect((await vs.tab('Document1.docx'))?.uri).toMatch(/^untitled:/);
  await frame.evaluate(() => (window as unknown as { docx4jEditor: { focus(): void } }).docx4jEditor.focus());
  await vs.page.keyboard.type('A new document', { delay: 20 });
  await expect.poll(async () => (await vs.tab('Document1.docx'))?.isDirty).toBe(true);
  const caret = await head(frame);
  // An untitled document's Save is VS Code's Save As (section 4.3).
  void vs.exec('workbench.action.files.save').catch(() => undefined);
  const input = vs.page.locator('.quick-input-widget input:not([type="checkbox"])');
  await input.waitFor();
  await input.fill(vs.path('made.docx'));
  await vs.page.keyboard.press('Enter');
  await settled('Document1.docx', 'made.docx');
  const saved = await vs.frameFor('made.docx');
  expect(mainPart(vs.path('made.docx'))).toContain('A new document');
  const tab = await vs.tab('made.docx');
  expect(tab).toMatchObject({ viewType: 'docx4j.editor', isDirty: false });
  expect(tab!.uri).toMatch(/^file:/);
  expect(await vs.tab('Document1.docx')).toBeUndefined();
  await expect.poll(() => head(saved)).toBe(caret);
});

test('Save As to a new .docx: the editor replaced, the caret kept', async () => {
  const frame = await vs.open('source.docx');
  await frame.evaluate(() => { const v = (window as unknown as { docx4jEditor: { caretAt(p: number): void; focus(): void } }).docx4jEditor; v.caretAt(6); v.focus(); });
  await vs.page.keyboard.type('S', { delay: 20 });
  const caret = await head(frame);
  await vs.saveAs(vs.path('source-as.docx'));
  await settled('source.docx', 'source-as.docx');
  const again = await vs.frameFor('source-as.docx');
  await expect.poll(() => head(again)).toBe(caret);
  expect(await vs.tab('source.docx')).toBeUndefined();
  expect((await vs.tab('source-as.docx'))?.isDirty).toBe(false);
  expect(mainPart(vs.path('source-as.docx'))).toContain('S');
});

test('Save As with an edit during the export, on the 255-page fixture: the new editor dirty with it, named as the target (P25)', async () => {
  const frame = await vs.open('large.docx');
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.type('E', { delay: 20 });
  await vs.saveAs(vs.path('large-as.docx'));
  await sleep(500);
  // An edit the page makes while it exports: it runs once the export is done, before the hand-over is asked.
  void insertInPage(frame, 'LATE');
  await settled('large.docx', 'large-as.docx', 240_000);
  const again = await vs.frameFor('large-as.docx', 240_000);
  await expect.poll(() => bodyText(again), { timeout: 60_000 }).toContain('ELATE');
  await expect.poll(async () => (await vs.tab('large-as.docx'))?.isDirty).toBe(true);
  await expect(again.getByTestId('name')).toHaveText('large-as.docx •');
  // The file holds what was exported; the late edit is the new editor's, unsaved.
  const saved = mainPart(vs.path('large-as.docx'));
  expect(saved).toContain('E');
  expect(saved).not.toContain('ELATE');
});

test('Save As onto a file open, and unsaved, in another editor of this extension: refused, nothing written', async () => {
  // A target open with unsaved edits stays open, and the extension refuses (section 9.8 finding 6). One open and clean VS Code
  // closes itself, and fails the save with its own "No custom document found" (section 9.9): nothing is written either way.
  const two = await vs.open('open-two.docx');
  await caretAtEndOf(two, 0);
  await vs.page.keyboard.type('2', { delay: 20 });
  await expect.poll(async () => (await vs.tab('open-two.docx'))?.isDirty).toBe(true);
  const before = readFileSync(vs.path('open-two.docx'));
  const frame = await vs.open('open-one.docx');
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.type('1', { delay: 20 });
  // VS Code asks about replacing the file first, in its file dialog.
  await vs.saveAs(vs.path('open-two.docx'), { replace: true });
  await expect.poll(() => vs.notifications()).toContainEqual(expect.stringContaining('open-two.docx is open in another editor'));
  expect(readFileSync(vs.path('open-two.docx'))).toEqual(before);
  expect((await vs.tab('open-one.docx'))?.isDirty).toBe(true);
  expect((await vs.tab('open-two.docx'))?.isDirty).toBe(true);
});

test('Save As to notes.md: refused, nothing written', async () => {
  const frame = await vs.open('source.docx');
  await caretAtEndOf(frame, 0);
  await vs.saveAs(vs.path('notes.md'));
  await expect.poll(() => vs.notifications()).toContainEqual(expect.stringContaining('notes.md is not a Word document\'s name'));
  expect(exists(vs.path('notes.md'))).toBe(false);
});

test('Save As to .xml: Flat OPC written, and the file open in VS Code\'s text editor', async () => {
  const frame = await vs.open('flat.docx');
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.type(' flat', { delay: 20 });
  await vs.saveAs(vs.path('flat.xml'));
  await settled('flat.docx', 'flat.xml');
  expect((await vs.tab('flat.xml'))?.kind).toBe('text');
  const xml = readFileSync(vs.path('flat.xml'), 'utf8');
  expect(xml).toContain('pkg:package');
  expect(xml).toContain(' flat');
  expect(await vs.tab('flat.docx')).toBeUndefined();
});
