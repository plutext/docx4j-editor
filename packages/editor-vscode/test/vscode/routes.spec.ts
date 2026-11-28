// SPDX-License-Identifier: GPL-3.0-or-later
//
// The rest of section 4.3's suite in a real VS Code: auto save; the options shared between editors and across a restart
// (P19); Restricted Mode (P27); a link's address (P28); the console's TypeScript diagnostic from a worker started from the
// resource origin's script. Each describe has a VS Code of its own.
import { test, expect } from '@playwright/test';
import { VsCode, caretAtEndOf, sleep } from './harness.js';
import { mainPart, withLink } from './zip.js';
import { join } from 'node:path';
import { FIXTURES } from './harness.js';

test.describe.configure({ mode: 'serial' });

test.describe('auto save at afterDelay', () => {
  let vs: VsCode;
  test.beforeAll(async () => {
    vs = await VsCode.launch({ profile: 'autosave', files: { 'auto.docx': 'HelloWordOnline.docx', 'large.docx': 'large/anon-250.docx' }, settings: { 'files.autoSave': 'afterDelay', 'files.autoSaveDelay': 1000 } });
  });
  test.afterAll(async () => { await vs?.close(); });

  test('a Word document saved after a pause; an untitled one not', async () => {
    const frame = await vs.open('auto.docx');
    await caretAtEndOf(frame, 0);
    await vs.page.keyboard.type(' saved by itself', { delay: 20 });
    await expect.poll(async () => (await vs.tab('auto.docx'))?.isDirty, { timeout: 20_000 }).toBe(false);
    expect(mainPart(vs.path('auto.docx'))).toContain('saved by itself');
    await vs.exec('docx4j.newDocument');
    let page: import('@playwright/test').Frame | undefined;
    await expect.poll(async () => {
      for (const f of await vs.frames()) if (await f.locator('[role="dialog"]').count() > 0) { page = f; return true; }
      return false;
    }).toBe(true);
    await page!.locator('[role="dialog"] button[type="submit"]').click();
    const untitled = await vs.frameFor('Document1.docx');
    await untitled.evaluate(() => (window as unknown as { docx4jEditor: { focus(): void } }).docx4jEditor.focus());
    await vs.page.keyboard.type('never saved by itself', { delay: 20 });
    await sleep(4000);
    expect((await vs.tab('Document1.docx'))?.isDirty).toBe(true);
    expect(vs.problems).toEqual([]);
    await vs.closeAll();
  });

  test('on the 255-page fixture, the stall a key meets while auto save exports: measured and reported, not gated', async () => {
    const frame = await vs.open('large.docx');
    await caretAtEndOf(frame, 0);
    await vs.page.keyboard.type('a', { delay: 20 });
    // The pause that starts the auto save, then a key: how long the key waits for the page.
    await sleep(1500);
    const t0 = Date.now();
    await vs.page.keyboard.type('b');
    const stallMs = Date.now() - t0;
    test.info().annotations.push({ type: 'auto save stall (ms)', description: String(stallMs) });
    console.log(`auto save stall on the 255-page fixture: ${stallMs} ms`);
    await expect.poll(async () => (await vs.tab('large.docx'))?.isDirty, { timeout: 200_000 }).toBe(false);
    expect(vs.problems).toEqual([]);
  });
});

test.describe('the options (P19)', () => {
  let vs: VsCode;
  test.beforeAll(async () => {
    vs = await VsCode.launch({ profile: 'options', files: { 'one.docx': 'HelloWordOnline.docx', 'two.docx': 'fonts-hebrew-no-cs.docx' } });
  });
  test.afterAll(async () => { await vs?.close(); });

  const storedName = (frame: import('@playwright/test').Frame): Promise<string | null> => frame.evaluate(() => localStorage.getItem('docx4j-editor.name'));

  test('the author\'s name set in one editor\'s Options: found in a second editor\'s storage, and after a restart', async () => {
    const one = await vs.open('one.docx');
    await one.locator('.options-menu > summary').click();
    await one.getByTestId('option-name').fill('Suite Author');
    await one.getByTestId('option-name').press('Tab');
    await one.locator('.options-menu > summary').click();
    await expect.poll(() => storedName(one)).toBe('Suite Author');
    // A second editor open already hears it, and one opened afterwards has it at its start.
    const two = await vs.open('two.docx');
    await expect.poll(() => storedName(two)).toBe('Suite Author');
    vs = await vs.restart();
    await vs.exec('workbench.action.closeAllEditors');
    const again = await vs.open('one.docx');
    await expect.poll(() => storedName(again)).toBe('Suite Author');
    expect(vs.problems).toEqual([]);
  });
});

test.describe('Restricted Mode (P27)', () => {
  let vs: VsCode;
  test.beforeAll(async () => {
    vs = await VsCode.launch({ profile: 'restricted', trust: true, files: { 'a.docx': 'HelloWordOnline.docx' } });
  });
  test.afterAll(async () => { await vs?.close(); });

  test('a document in an untrusted folder offers no document macro', async () => {
    expect((await vs.ctl<{ trusted: boolean }>({ op: 'ping' })).trusted).toBe(false);
    const frame = await vs.open('a.docx');
    await frame.getByTestId('macros-menu').click();
    await expect(frame.getByTestId('macros-disabled')).toContainText('Restricted Mode');
    expect(vs.problems).toEqual([]);
  });
});

test.describe('links (P28)', () => {
  let vs: VsCode;
  test.beforeAll(async () => {
    vs = await VsCode.launch({ profile: 'links', files: { 'target.docx': 'fonts-hebrew-no-cs.docx' } });
    const plain = join(FIXTURES, 'HelloWordOnline.docx');
    withLink(plain, vs.path('web.docx'), 'https://example.com/docx4j', 'a web link');
    withLink(plain, vs.path('relative.docx'), 'target.docx', 'a relative link');
    withLink(plain, vs.path('command.docx'), 'command:workbench.action.reloadWindow', 'a command link');
  });
  test.afterAll(async () => { await vs?.close(); });
  test.afterEach(async () => {
    expect(vs.problems).toEqual([]);
    await vs.closeAll();
  });

  test('an https link reaches VS Code\'s prompt for an outside site, which the test dismisses', async () => {
    const frame = await vs.open('web.docx');
    await frame.locator('.ProseMirror a.link', { hasText: 'a web link' }).click({ modifiers: ['Control'] });
    await expect.poll(() => vs.dialog()).toContain('https://example.com/docx4j');
    await vs.clickDialogButton('Cancel');
  });

  test('a relative .docx link opens that file in this editor', async () => {
    const frame = await vs.open('relative.docx');
    await frame.locator('.ProseMirror a.link', { hasText: 'a relative link' }).click({ modifiers: ['Control'] });
    await vs.frameFor('target.docx');
    expect((await vs.tab('target.docx'))?.viewType).toBe('docx4j.editor');
  });

  test('a command: link is refused, and VS Code says so', async () => {
    const frame = await vs.open('command.docx');
    await frame.locator('.ProseMirror a.link', { hasText: 'a command link' }).click({ modifiers: ['Control'] });
    await expect.poll(() => vs.notifications()).toContainEqual(expect.stringContaining('The link was not followed'));
    expect(await vs.frames()).toHaveLength(1);
  });
});

test.describe('the console', () => {
  let vs: VsCode;
  test.beforeAll(async () => {
    vs = await VsCode.launch({ profile: 'console', files: { 'a.docx': 'HelloWordOnline.docx' } });
  });
  test.afterAll(async () => { await vs?.close(); });

  test('a TypeScript diagnostic, from a worker whose script came from the resource origin, and a Run under the policy', async () => {
    const frame = await vs.open('a.docx');
    await frame.getByTestId('perspective-developer').click();
    await expect(frame.getByTestId('console-types')).toContainText('types loaded', { timeout: 60_000 });
    await frame.waitForFunction(() => (window as unknown as { docx4jConsole?: unknown }).docx4jConsole !== undefined);
    const setScript = (s: string): Promise<void> => frame.evaluate((v) => { (window as unknown as { docx4jConsole: { setValue(v: string): void } }).docx4jConsole.setValue(v); }, s);
    await setScript(`await Word.run(async (context) => {\n  const body = context.document.body;\n  body.insertFileFromBase64('AAAA', 'End');\n});`);
    await frame.getByTestId('console-check').click();
    const output = frame.getByTestId('console-output');
    await expect(output).toContainText('not in the subset');
    // The Run imports a Blob, allowed by the nonce alone: the policy has no blob: in script-src (section 9.9).
    await setScript(`const p = body.insertParagraph('Run in VS Code', 'End');\nreturn p.text;`);
    await frame.getByTestId('console-run').click();
    await expect(output).toContainText('Run in VS Code');
    expect(await frame.evaluate(() => document.querySelector('meta[http-equiv="Content-Security-Policy"]')?.getAttribute('content'))).not.toMatch(/script-src[^;]*blob:/);
    expect(vs.problems).toEqual([]);
  });
});
