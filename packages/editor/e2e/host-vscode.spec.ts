// SPDX-License-Identifier: GPL-3.0-or-later
//
// The application as VS Code's page (ED-014 section 4.1, S1.4), in the production build: an init script defines
// `acquireVsCodeApi` and records what the page posts, and each test sends the extension's messages. There is no
// extension yet (S3); this is the page half of the protocol, held to section 4.1's list. The whole application under
// StrictMode, `ready` once, is `test/host-app.test.tsx` (the production build never replays an effect). The fake API
// is `vscode-page.ts`'s, shared with `webview.spec.ts`, which loads the page the extension's writer makes.
import { test, expect, type Page } from '@playwright/test';
import { fixture, caretInBlock, selectBlock } from './helpers.js';
import { hostBootMeta } from '../src/app/webview-page.js';
import { VSCODE_API_INIT, b64, bytesOf, ofType, send, type Posted } from './vscode-page.js';
// The browser suite's provider double, whose answers the extension's would carry.
import { answerFor } from '../../editor-model/test/fixtures/build/ask-double.mjs';

const body = '.docx-editor .ProseMirror[aria-label="Document"]';

/** A page as VS Code's webview would load it: the fake API, and the boot data in the page the extension writes. */
async function vscodePage(page: Page, boot?: Record<string, string>): Promise<void> {
  if (boot) {
    await page.route((url) => url.pathname === '/' || url.pathname === '/index.html', async (route) => {
      const response = await route.fetch();
      // The boot meta as the extension's page writer writes it (`app/webview-page.ts`, ED-014 P24), escaped the same way.
      const html = (await response.text()).replace('<head>', `<head>${hostBootMeta({ options: boot })}`);
      await route.fulfill({ response, body: html });
    });
  }
  await page.addInitScript(VSCODE_API_INIT);
  await page.goto('/');
  await page.getByTestId('opening').waitFor();
}

/** The extension's `open`, and the waits `openFixture` has. */
async function open(page: Page, name: string, bytes: Uint8Array, resume?: { update: Uint8Array; meta: unknown }): Promise<void> {
  await send(page, { type: 'open', name, bytes: b64(bytes), ...(resume ? { resume: { update: b64(resume.update), meta: resume.meta } } : {}) });
  await page.locator(body).waitFor();
  await page.getByTestId('toolbar').waitFor();
}

const nameShown = (page: Page) => page.getByTestId('name');

/**
 * The boot data put in the page by an init script, before the application reads it: `page.route` would put the page off
 * local network access, and its WebSocket to the preview server's relay with it.
 */
async function bootByScript(page: Page, boot: Record<string, unknown>): Promise<void> {
  await page.addInitScript((content) => {
    const add = (): boolean => {
      if (!document.head) return false;
      const meta = document.createElement('meta');
      meta.name = 'docx4j-host';
      meta.content = content;
      document.head.prepend(meta);
      return true;
    };
    if (!add()) {
      const watch = new MutationObserver(() => { if (add()) watch.disconnect(); });
      watch.observe(document, { childList: true, subtree: true });
    }
  }, JSON.stringify({ options: {}, ...boot }));
}

const PREVIEW_RELAY = 'ws://localhost:4173/relay';

/** The extension's answer to the page's next `companionStart`: the preview server's relay as the companion's. */
async function answerStart(page: Page, url = 'http://127.0.0.1:24402', relay = PREVIEW_RELAY, count = 1): Promise<void> {
  await expect.poll(async () => (await ofType(page, 'companionStart')).length, { timeout: 20_000 }).toBe(count);
  const start = (await ofType(page, 'companionStart')).at(-1)!;
  await send(page, { type: 'companion', id: start['id'], url, relay });
}

test('an open: ready posted once, the page waiting until then, VS Code\'s chrome gone, typing posting changed', async ({ page }) => {
  await vscodePage(page);
  expect(await ofType(page, 'ready')).toEqual([{ type: 'ready', origin: 'http://localhost:4173' }]);
  await expect(page.getByTestId('opening')).toHaveText('Opening…');
  await expect(page.getByTestId('open-sample')).toHaveCount(0);
  await open(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  await expect(nameShown(page)).toHaveText('hello.docx');
  for (const id of ['new', 'open', 'save', 'save-as', 'download', 'print', 'join', 'close']) await expect(page.getByTestId(id), id).toHaveCount(0);
  await page.getByLabel('Options').click();
  await expect(page.getByTestId('option-layout')).toHaveCount(0);
  await expect(page.getByTestId('option-save-in-place')).toHaveCount(0);
  await page.getByLabel('Options').click();
  expect(await ofType(page, 'changed')).toEqual([]);
  await caretInBlock(page, 0);
  await page.keyboard.type('Typed ');
  await expect.poll(async () => (await ofType(page, 'changed')).length).toBeGreaterThan(0);
  // At most one per 250 ms: six keys typed at once are one or two.
  expect((await ofType(page, 'changed')).length).toBeLessThanOrEqual(2);
  await expect(nameShown(page)).toHaveText('hello.docx •');
  expect(await ofType(page, 'ready')).toHaveLength(1);
});

test('Ctrl+S, Ctrl+P and Ctrl+Shift+P do nothing in the page and reach the window for VS Code; the host\'s palette message opens the palette', async ({ page }) => {
  await vscodePage(page);
  await open(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  await caretInBlock(page, 0);
  await page.evaluate(() => { (window as unknown as { __keys: KeyboardEvent[] }).__keys.length = 0; });
  let downloads = 0;
  page.on('download', () => { downloads++; });
  await page.keyboard.press('Control+s');
  await page.keyboard.press('Control+p');
  await page.keyboard.press('Control+Shift+P');
  // Outside an editing surface too: App's own listener's branch.
  await page.evaluate(() => { (document.activeElement as HTMLElement | null)?.blur(); });
  await page.keyboard.press('Control+Shift+P');
  await page.waitForTimeout(300);
  const keys = await page.evaluate(() => (window as unknown as { __keys: KeyboardEvent[] }).__keys
    .filter((e) => e.ctrlKey && e.key !== 'Control' && e.key !== 'Shift').map((e) => ({ key: e.key.toLowerCase(), shift: e.shiftKey, prevented: e.defaultPrevented })));
  expect(keys).toEqual([
    { key: 's', shift: false, prevented: false },
    { key: 'p', shift: false, prevented: false },
    { key: 'p', shift: true, prevented: false },
    { key: 'p', shift: true, prevented: false },
  ]);
  expect(downloads).toBe(0);
  await expect(page.getByTestId('palette')).toHaveCount(0);
  await expect(page.getByTestId('status')).not.toContainText('print');
  expect(await ofType(page, 'saved')).toEqual([]);
  // Ctrl+F stays the application's (S0 item 5).
  await caretInBlock(page, 0);
  await page.keyboard.press('Control+f');
  await expect(page.locator('.find-bar')).toBeVisible();
  await send(page, { type: 'palette' });
  await expect(page.getByTestId('palette')).toBeVisible();
});

test('a save in two phases: the bytes reopen with the edit; written clears the mark, an edit between saved and written keeps it; a .xml name saves Flat OPC', async ({ page, context }) => {
  await vscodePage(page);
  await open(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  await caretInBlock(page, 0);
  await page.keyboard.type('Saved edit ');
  await send(page, { type: 'save', id: 1, name: 'hello.docx' });
  await expect.poll(async () => (await ofType(page, 'saved')).length, { timeout: 30_000 }).toBe(1);
  const [first] = await ofType(page, 'saved');
  expect(first).toMatchObject({ id: 1 });
  expect(first!['refused']).toBeUndefined();
  // Not saved until the host has written it.
  await expect(nameShown(page)).toHaveText('hello.docx •');
  await send(page, { type: 'written', id: 1, name: 'hello.docx' });
  await expect(nameShown(page)).toHaveText('hello.docx');

  // An edit between the export and the write: still unsaved, and VS Code told so again after the write.
  await page.keyboard.type('More ');
  await send(page, { type: 'save', id: 2, name: 'hello.docx' });
  await expect.poll(async () => (await ofType(page, 'saved')).length, { timeout: 30_000 }).toBe(2);
  await page.keyboard.type('Between ');
  await page.waitForTimeout(300);
  const before = (await ofType(page, 'changed')).length;
  await send(page, { type: 'written', id: 2, name: 'hello.docx' });
  await expect(nameShown(page)).toHaveText('hello.docx •');
  await expect.poll(async () => (await ofType(page, 'changed')).length).toBe(before + 1);
  await expect(page.getByTestId('status')).toContainText('The changes made during the save are not saved yet.');

  // Save As to a .xml name: Flat OPC.
  await send(page, { type: 'save', id: 3, name: 'hello.xml' });
  await expect.poll(async () => (await ofType(page, 'saved')).length, { timeout: 30_000 }).toBe(3);
  const flat = new TextDecoder().decode(bytesOf((await ofType(page, 'saved'))[2]!['bytes']));
  expect(flat.startsWith('<?xml')).toBe(true);
  expect(flat).toContain('pkg:package');
  await send(page, { type: 'written', id: 3, name: 'hello.xml' });
  await expect(nameShown(page)).toHaveText(/^hello\.xml/);

  // A write that fails: the document stays unsaved, and the page says why.
  await caretInBlock(page, 0);
  await page.keyboard.type('Unwritten ');
  await send(page, { type: 'save', id: 4, name: 'hello.xml' });
  await expect.poll(async () => (await ofType(page, 'saved')).length, { timeout: 30_000 }).toBe(4);
  await send(page, { type: 'writeFailed', id: 4, error: 'disk full' });
  await expect(page.getByTestId('status')).toHaveText('Not saved: disk full');
  await expect(nameShown(page)).toHaveText('hello.xml •');

  // The first save's bytes, opened in another page, hold its edit.
  const other = await context.newPage();
  await vscodePage(other);
  await open(other, 'hello.docx', bytesOf(first!['bytes']));
  await expect(other.locator(`${body} > p`).first()).toContainText('Saved edit');
  await expect(nameShown(other)).toHaveText('hello.docx');
});

test('a backup, resumed in a new page, shows the edit and is dirty', async ({ page, context }) => {
  const bytes = await fixture('HelloWordOnline.docx');
  await vscodePage(page);
  await open(page, 'hello.docx', bytes);
  await caretInBlock(page, 0);
  await page.keyboard.type('Backed up ');
  await send(page, { type: 'backup', id: 1 });
  await expect.poll(async () => (await ofType(page, 'backup')).length).toBe(1);
  const [backup] = await ofType(page, 'backup');
  expect(backup!['meta']).toMatchObject({ name: 'hello.docx' });
  const other = await context.newPage();
  await vscodePage(other);
  await open(other, 'hello.docx', bytes, { update: bytesOf(backup!['update']), meta: backup!['meta'] });
  await expect(other.locator(`${body} > p`).first()).toContainText('Backed up');
  await expect(nameShown(other)).toHaveText('hello.docx •');
  await expect(other.getByTestId('status')).toContainText('Recovered unsaved changes');
});

test('a backup taken in a room is resumed in the room when its relay answers', async ({ page, browser }) => {
  const bytes = await fixture('HelloWordOnline.docx');
  // Shared into the companion's relay, which the double answers with the preview server's (ED-014 P41: VS Code's page has no
  // relay option). The boot by an init script, since `page.route` would block the page's WebSocket to localhost.
  await bootByScript(page, { agent: 'Agent' });
  await vscodePage(page);
  await open(page, 'hello.docx', bytes);
  await page.getByTestId('share').click();
  await answerStart(page);
  await page.getByTestId('share-link').waitFor();
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 20_000 });
  await caretInBlock(page, 0);
  await page.keyboard.type('In the room ');
  await send(page, { type: 'backup', id: 1 });
  await expect.poll(async () => (await ofType(page, 'backup')).length).toBe(1);
  const [backup] = await ofType(page, 'backup');
  expect((backup!['meta'] as { room?: { relay: string } }).room?.relay).toBe('ws://localhost:4173/relay');
  const context = await browser.newContext();
  const other = await context.newPage();
  await vscodePage(other);
  await open(other, 'hello.docx', bytes, { update: bytesOf(backup!['update']), meta: backup!['meta'] });
  await expect(other.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 20_000 });
  await expect(other.locator(`${body} > p`).first()).toContainText('In the room');
  await expect(nameShown(other)).toHaveText('hello.docx •');
  await expect(other.getByTestId('status')).not.toContainText('could not be reached');
  await context.close();
});

test('a room the host asks to join that the relay refuses: the waiting page says the open failed, and the host hears it', async ({ page }) => {
  await vscodePage(page);
  await send(page, { type: 'open', name: '', join: { room: 'B'.repeat(22), token: 'C'.repeat(43), relay: 'ws://localhost:4173/relay' } });
  await expect(page.getByTestId('opening')).toHaveText('The document could not be opened.', { timeout: 30_000 });
  await expect(page.getByTestId('status')).toContainText('holds no such room');
  expect(await ofType(page, 'status')).toContainEqual({ type: 'status', text: 'The document could not be opened.' });
});

test('revert opens the bytes again in place, the edit gone and the document saved; with unsaved edits, it says they were discarded', async ({ page }) => {
  const bytes = await fixture('HelloWordOnline.docx');
  await vscodePage(page);
  await open(page, 'hello.docx', bytes);
  await caretInBlock(page, 0);
  await page.keyboard.type('Reverted away ');
  await expect(nameShown(page)).toHaveText('hello.docx •');
  await send(page, { type: 'revert', bytes: b64(bytes) });
  await expect(nameShown(page)).toHaveText('hello.docx');
  await expect(page.locator(body)).not.toContainText('Reverted away');
  // Whoever asked (the banner's discard, or VS Code's Revert File), a revert of a dirty document discards its edits.
  await expect(page.getByTestId('status')).toContainText('Reloaded hello.docx from disk; your unsaved changes were discarded.');
  await send(page, { type: 'revert', bytes: b64(bytes) });
  await expect(page.getByTestId('status')).toHaveText('Reloaded hello.docx from disk.');
});

test('disk shows the banner with the counts, whose Reload posts revertRequest; dirty, Discard and reload does after its confirmation', async ({ page }) => {
  await vscodePage(page);
  await open(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  const other = await fixture('tracked-changes.docx');
  await send(page, { type: 'disk', bytes: b64(other), lastModified: Date.now() });
  const offer = page.getByTestId('disk-offer');
  await expect(offer).toContainText('hello.docx changed on disk since it was opened.');
  await expect(page.getByTestId('disk-counts')).toContainText('On disk');
  await page.getByTestId('disk-reload').click();
  await expect.poll(async () => (await ofType(page, 'revertRequest')).length).toBe(1);
  await expect(offer).toHaveCount(0);
  await caretInBlock(page, 0);
  await page.keyboard.type('Local ');
  await send(page, { type: 'disk', bytes: b64(other), lastModified: Date.now() });
  await expect(offer).toBeVisible();
  // The disk version in a new tab is the browser's alone (P9).
  await expect(page.getByTestId('disk-open-tab')).toHaveCount(0);
  await page.getByTestId('disk-discard').click();
  await page.getByRole('button', { name: 'Discard and reload' }).click();
  await expect.poll(async () => (await ofType(page, 'revertRequest')).length).toBe(2);
});

test('Ctrl+click on a link posts openExternal, and opens no tab', async ({ page, context }) => {
  await vscodePage(page);
  await open(page, 'tracked-changes.docx', await fixture('tracked-changes.docx'));
  await selectBlock(page, 4);
  await page.getByTestId('link-button').click();
  await page.getByTestId('link-address').fill('https://docx4j.org/');
  await page.getByRole('button', { name: 'OK' }).click();
  const link = page.locator(`${body} > p`).nth(4).locator('a.link');
  await expect(link).toHaveText('Some text.');
  await link.click({ modifiers: ['Control'] });
  await expect.poll(async () => ofType(page, 'openExternal')).toEqual([{ type: 'openExternal', url: 'https://docx4j.org/' }]);
  expect(context.pages()).toHaveLength(1);
});

test('New from a template posts untitled with the new document, and leaves the template open here (P17)', async ({ page }) => {
  const { WordprocessingMLPackage, ContentTypes } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(await fixture('HelloWordOnline.docx'));
  const main = pkg.getMainDocumentPart();
  main.contentType = ContentTypes.WORDPROCESSINGML_TEMPLATE;
  pkg.contentTypeManager.addOverrideContentType(main.partName, ContentTypes.WORDPROCESSINGML_TEMPLATE);
  await vscodePage(page);
  await open(page, 'letterhead.dotx', await pkg.save());
  await page.getByTestId('template-new').click();
  await expect.poll(async () => (await ofType(page, 'untitled')).length).toBe(1);
  const [untitled] = await ofType(page, 'untitled');
  expect(String(untitled!['name'])).toMatch(/\.docx$/);
  expect(bytesOf(untitled!['bytes']).length).toBeGreaterThan(1000);
  await expect(nameShown(page)).toHaveText('letterhead.dotx');
});

test('an open with empty bytes asks the page size, makes the document under the host\'s name, and posts its base', async ({ page }) => {
  await vscodePage(page);
  await send(page, { type: 'open', name: 'Untitled-1.docx', bytes: b64(new Uint8Array()) });
  await page.getByTestId('page-size').selectOption('LETTER');
  await page.getByRole('button', { name: 'OK' }).click();
  await page.locator(body).waitFor();
  await expect(nameShown(page)).toHaveText('Untitled-1.docx');
  await expect.poll(async () => (await ofType(page, 'base')).length).toBe(1);
  const [base] = await ofType(page, 'base');
  expect(bytesOf(base!['bytes']).length).toBeGreaterThan(1000);
  // The choice is an option, posted for the other editors (P19).
  await expect.poll(async () => (await ofType(page, 'options')).some((m) => (m['set'] as Record<string, string> | undefined)?.['docx4j-editor.pageSize'] === 'LETTER')).toBe(true);
});

test('a Markdown open: its conversion posted as the base, its table fit posting no changed; disk with the .md\'s bytes shows the banner', async ({ page }) => {
  const rows = Array.from({ length: 6 }, (_, i) => `| ${'abcdef'[i]} | ${'A description long enough to wrap over several lines of a narrow column, which the fit widens. '.repeat(2 + i)} |`);
  const md = `# Notes\n\n| ID | Description |\n|---|---|\n${rows.join('\n')}\n`;
  await vscodePage(page);
  await send(page, { type: 'open', name: 'notes.md', bytes: b64(new TextEncoder().encode(md)) });
  await page.locator(body).waitFor();
  await expect(nameShown(page)).toHaveText('notes.docx');
  await expect(page.getByTestId('status')).toContainText('fit to its content', { timeout: 20_000 });
  expect((await ofType(page, 'base')).length).toBe(1);
  expect(await ofType(page, 'changed')).toEqual([]);
  await expect(nameShown(page)).toHaveText('notes.docx');
  await send(page, { type: 'disk', bytes: b64(new TextEncoder().encode(md.replace('# Notes', '# Notes, changed'))), lastModified: Date.now() });
  await expect(page.getByTestId('disk-offer')).toContainText('notes.md changed on disk since it was opened.');
  await expect(page.getByTestId('disk-counts')).toContainText('1 paragraph is changed');
});

test('a file dropped on the page opens nothing', async ({ page }) => {
  await vscodePage(page);
  await open(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  const other = Buffer.from(await fixture('tracked-changes.docx')).toString('base64');
  await page.evaluate((data) => {
    const dt = new DataTransfer();
    dt.items.add(new File([Uint8Array.from(atob(data), (c) => c.charCodeAt(0))], 'other.docx', { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
    const target = document.querySelector('.app')!;
    target.dispatchEvent(new DragEvent('dragover', { dataTransfer: dt, bubbles: true, cancelable: true }));
    target.dispatchEvent(new DragEvent('drop', { dataTransfer: dt, bubbles: true, cancelable: true }));
  }, other);
  await page.waitForTimeout(1000);
  await expect(nameShown(page)).toHaveText('hello.docx');
});

test('the options: the host\'s boot data read before the language loads, a change posted, another editor\'s applied unposted (P19)', async ({ page }) => {
  await vscodePage(page, { 'docx4j-editor.language': 'fr', 'docx4j-editor.relay': 'wss://relay.example.org' });
  await expect(page.getByTestId('opening')).toHaveText('Ouverture…');
  await open(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  expect(await page.evaluate(() => localStorage.getItem('docx4j-editor.relay'))).toBe('wss://relay.example.org');
  await page.getByTestId('zoom-in').click();
  await expect.poll(async () => (await ofType(page, 'options')).some((m) => (m['set'] as Record<string, string> | undefined)?.['docx4j-editor.zoom'] !== undefined)).toBe(true);
  const count = (await ofType(page, 'options')).length;
  await send(page, { type: 'options', set: { 'docx4j-editor.name': 'Ann' } });
  expect(await page.evaluate(() => localStorage.getItem('docx4j-editor.name'))).toBe('Ann');
  await page.waitForTimeout(100);
  expect((await ofType(page, 'options')).length).toBe(count);
});

// ---- S4: the keyboard and the clipboard (ED-014 section 4.4). The init script's listener on `window`, added before the
// application's code, stands for VS Code's host page, which forwards every key that reaches it to the workbench.

/** The keydowns that reached the window since the last call, as `Ctrl+Shift+k`, modifiers' own keys left out. */
async function keysSeen(page: Page): Promise<string[]> {
  return page.evaluate(() => {
    const keys = (window as unknown as { __keys: KeyboardEvent[] }).__keys;
    const out = keys.filter((e) => !['Control', 'Shift', 'Alt', 'Meta'].includes(e.key))
      .map((e) => `${e.ctrlKey ? 'Ctrl+' : ''}${e.altKey ? 'Alt+' : ''}${e.shiftKey ? 'Shift+' : ''}${e.key.length === 1 ? e.key.toLowerCase() : e.key}`);
    keys.length = 0;
    return out;
  });
}

const footer = '.docx-editor .page-boundary .boundary-footer .ProseMirror';

test('S4: the document\'s keys stay in the page; VS Code\'s go on, from the body and from a footer (P32)', async ({ page }) => {
  await vscodePage(page);
  await open(page, 'hf.docx', await fixture('odd-even-first-page.docx'));
  await caretInBlock(page, 0);
  await keysSeen(page);
  // The page's: Word's keys, its shortcut Ctrl+K, and Ctrl+Q, which VS Code would take for Quit.
  await page.keyboard.press('Control+b');
  await page.keyboard.press('Control+b');
  await page.keyboard.press('Control+q');
  await page.keyboard.press('Control+e');
  await page.keyboard.press('Control+j');
  await page.keyboard.press('Control+Equal');
  await page.keyboard.press('Control+Equal');
  await page.keyboard.press('Control+k');
  const dialog = page.getByTestId('link-address');
  await expect(dialog).toBeVisible();
  // Typed into the dialog's field: VS Code sees these keys, but no chord was started, so they end none.
  await dialog.pressSequentially('www.');
  await page.keyboard.press('Escape');
  await expect(dialog).toHaveCount(0);
  expect(await keysSeen(page)).toEqual(['w', 'w', 'w', '.', 'Escape']);
  // VS Code's: nothing in the page acts on them.
  await caretInBlock(page, 0);
  await page.keyboard.press('Control+s');
  await page.keyboard.press('Control+p');
  await page.keyboard.press('Control+Shift+P');
  await page.keyboard.press('F1');
  await page.keyboard.press('Control+w');
  expect(await keysSeen(page)).toEqual(['Ctrl+s', 'Ctrl+p', 'Ctrl+Shift+p', 'F1', 'Ctrl+w']);
  // The same from a footer, whose keys a page-boundary widget had kept from the page and VS Code alike.
  await page.locator(footer).first().click();
  await expect(page.locator('.page-column')).toHaveClass(/editing-header/);
  await keysSeen(page);
  await page.keyboard.press('Control+s');
  await page.keyboard.press('Control+p');
  await page.keyboard.press('F1');
  await page.keyboard.press('Control+b');
  expect(await keysSeen(page)).toEqual(['Ctrl+s', 'Ctrl+p', 'F1']);
  // Ctrl+K in the footer opens the link dialog, over the footer.
  await page.keyboard.press('Control+k');
  await expect(dialog).toBeVisible();
  await page.keyboard.press('Escape');
  expect(await keysSeen(page)).toEqual(['Escape']);
  await expect(page.locator('.page-column')).toHaveClass(/editing-header/);
});

test('S4: Ctrl+Shift+E from a panel is kept; Alt+PageDown and Escape go on though the page acts on them (P32)', async ({ page }) => {
  await vscodePage(page);
  await open(page, 'tracked.docx', await fixture('tracked-changes.docx'));
  await page.getByTestId('perspective-review').click();
  // A panel's button has the focus: the panel listener, now in the capture phase, acts and prevents.
  const tracking = page.getByTestId('tracking-switch');
  await expect(tracking).toHaveAttribute('aria-pressed', 'false');
  await page.getByTestId('perspective-review').focus();
  await keysSeen(page);
  await page.keyboard.press('Control+Shift+E');
  // With no name set, a name is asked for first.
  const name = page.getByTestId('tracking-name');
  await expect(name).toBeVisible();
  expect(await keysSeen(page)).toEqual([]);
  await name.fill('Ann');
  await name.press('Enter');
  await expect(tracking).toHaveAttribute('aria-pressed', 'true');
  // The dialog's Enter, which nothing in the page prevents, goes on.
  expect(await keysSeen(page)).toEqual(['Enter']);
  // The next change, from the document: the caret moves, and VS Code sees the key, so its menu bar never reads a lone Alt.
  await caretInBlock(page, 0);
  const at = (): Promise<number> => page.evaluate(() => (window as unknown as { docx4jEditor: { state: { selection: { from: number } } } }).docx4jEditor.state.selection.from);
  const from = await at();
  await page.keyboard.press('Alt+PageDown');
  await expect.poll(at).not.toBe(from);
  await page.keyboard.press('Escape');
  expect(await keysSeen(page)).toEqual(['Alt+PageDown', 'Escape']);
});

test('S4: Ctrl+Z in a field undoes the field\'s typing, and is kept from VS Code (P32)', async ({ page }) => {
  await vscodePage(page);
  await open(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  await caretInBlock(page, 0);
  await page.keyboard.press('Control+f');
  const query = page.getByTestId('find-query');
  await query.pressSequentially('abc');
  await keysSeen(page);
  await page.keyboard.press('Control+z');
  await expect(query).toHaveValue('');
  expect(await keysSeen(page)).toEqual([]);
});

test('S4: Ctrl+Shift+V pastes the clipboard\'s text alone, in the body, in a field and with nothing focused; pasteText does the same (P34)', async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await vscodePage(page);
  await open(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  const first = page.locator(`${body} > p`).first();
  const clip = (text: string, html?: string): Promise<void> => page.evaluate(async ([t, h]) => {
    const items: Record<string, Blob> = { 'text/plain': new Blob([t!], { type: 'text/plain' }) };
    if (h) items['text/html'] = new Blob([h], { type: 'text/html' });
    await navigator.clipboard.write([new ClipboardItem(items)]);
  }, [text, html] as const);
  // The body: HTML on the clipboard, and only its text comes in.
  await clip('Plain words', '<b>Plain words</b>');
  await caretInBlock(page, 0);
  await keysSeen(page);
  await page.keyboard.press('Control+Shift+V');
  await expect(first).toContainText('Plain words');
  await expect(first.locator('strong, b, [style*="bold"]')).toHaveCount(0);
  expect(await keysSeen(page)).toEqual([]);
  // One web address is a link, as a plain paste makes one in a browser.
  await clip('https://docx4j.org/');
  await caretInBlock(page, 0);
  await page.keyboard.press('Control+Shift+V');
  await expect(first.locator('a.link')).toHaveAttribute('data-target', 'https://docx4j.org/');
  // A field: the find bar's.
  await page.keyboard.press('Control+f');
  const query = page.getByTestId('find-query');
  await query.fill('');
  await clip('in the field');
  await page.keyboard.press('Control+Shift+V');
  await expect(query).toHaveValue('in the field');
  await page.keyboard.press('Escape');
  // Nothing focused: the active surface.
  await caretInBlock(page, 0);
  await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur());
  expect(await page.evaluate(() => document.activeElement === document.body)).toBe(true);
  await clip('Unfocused ');
  await page.keyboard.press('Control+Shift+V');
  await expect(first).toContainText('Unfocused');
  // "Paste as Plain Text" from VS Code's menus: the extension's message, the text pasted where the focus is.
  await caretInBlock(page, 0);
  await send(page, { type: 'pasteText', text: 'From the menu ' });
  await expect(first).toContainText('From the menu');
  await page.keyboard.press('Control+f');
  await query.fill('');
  await send(page, { type: 'pasteText', text: 'menu in field' });
  await expect(query).toHaveValue('menu in field');
  // Where nothing can take a paste (a button), the key is left alone: it goes on to VS Code, and nothing is pasted.
  await page.keyboard.press('Escape');
  const text = await page.locator(body).innerText();
  await page.getByTestId('zoom-in').focus();
  await keysSeen(page);
  await clip('Not pasted');
  await page.keyboard.press('Control+Shift+V');
  await page.waitForTimeout(300);
  expect(await keysSeen(page)).toEqual(['Ctrl+Shift+v']);
  expect(await page.locator(body).innerText()).toBe(text);
});

test('S4: Options has no spelling option in VS Code (P35)', async ({ page }) => {
  await vscodePage(page);
  await open(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  await page.getByLabel('Options').click();
  await expect(page.getByTestId('option-name')).toBeVisible();
  await expect(page.getByTestId('option-spellcheck')).toHaveCount(0);
});

test('S5: Invite labelled from the boot asks for the companion, shares into its relay, and invites the agent into the room (P38)', async ({ page }) => {
  await bootByScript(page, { agent: 'Agent' });
  await vscodePage(page);
  await open(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  await expect(page.getByTestId('invite-agent')).toHaveText('Invite Agent');
  // No probe: the policy names no HTTP port of the companion's.
  const probes: string[] = [];
  page.on('request', (r) => { if (r.url().includes('/companion')) probes.push(r.url()); });
  await page.getByTestId('invite-agent').click();
  await answerStart(page);
  await expect.poll(async () => (await ofType(page, 'invite')).length, { timeout: 20_000 }).toBe(1);
  const [invite] = await ofType(page, 'invite');
  await expect(page.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 20_000 });
  expect(await ofType(page, 'room')).toEqual([{ type: 'room', relay: PREVIEW_RELAY }]);
  await send(page, { type: 'invited', id: invite!['id'] });
  await expect(page.getByTestId('status')).toContainText('Agent joined');
  // The backup names the room, so a hand-over and a restart rejoin it.
  await send(page, { type: 'backup', id: 1 });
  await expect.poll(async () => (await ofType(page, 'backup')).length).toBe(1);
  expect((await ofType(page, 'backup'))[0]!['meta']).toMatchObject({ room: { room: invite!['room'], token: invite!['token'], relay: PREVIEW_RELAY } });
  expect(probes).toEqual([]);
  // A refusal is said.
  await page.getByTestId('invite-agent').click();
  await answerStart(page, undefined, undefined, 2);
  await expect.poll(async () => (await ofType(page, 'invite')).length).toBe(2);
  await send(page, { type: 'invited', id: (await ofType(page, 'invite'))[1]!['id'], error: 'the relay stopped' });
  await expect(page.getByTestId('status')).toContainText('Agent could not join: the relay stopped');
});

test('S5: Share asks for the companion and shares into its relay; its link is on the companion\'s address, on this computer only; Options has no relay or companion field (P41)', async ({ page }) => {
  await bootByScript(page, { agent: 'Agent' });
  await vscodePage(page);
  await open(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  await page.getByTestId('share').click();
  await answerStart(page);
  await expect(page.getByTestId('share-link')).toHaveValue(/^http:\/\/127\.0\.0\.1:24402\/#room=[^&]+&token=[^&]+&relay=ws%3A%2F%2Flocalhost%3A4173%2Frelay$/, { timeout: 20_000 });
  await expect(page.getByTestId('share-local')).toBeVisible();
  await page.keyboard.press('Escape');
  await page.getByLabel('Options').click();
  await expect(page.getByTestId('option-relay')).toHaveCount(0);
  await expect(page.getByTestId('option-companion')).toHaveCount(0);
  await expect(page.getByTestId('option-name')).toBeVisible();
});

test('S5: with no agent in the boot (Restricted Mode), no Invite and no Share; a start refused is said', async ({ page }) => {
  await page.addInitScript(() => { try { localStorage.setItem('docx4j-editor.relay', '/relay'); localStorage.setItem('docx4j-editor.companion', 'http://localhost:4173/companion'); } catch { /* ignore */ } });
  await vscodePage(page);
  await open(page, 'hello.docx', await fixture('HelloWordOnline.docx'));
  // Options kept from a browser are no way round: VS Code's page offers neither, and probes nothing.
  await expect(page.getByTestId('invite-agent')).toHaveCount(0);
  await expect(page.getByTestId('share')).toHaveCount(0);
  await expect(page.getByTestId('repaginate')).toHaveCount(0);
  const other = await page.context().newPage();
  await bootByScript(other, { agent: 'Agent' });
  await vscodePage(other);
  await open(other, 'hello.docx', await fixture('HelloWordOnline.docx'));
  await other.getByTestId('share').click();
  await expect.poll(async () => (await ofType(other, 'companionStart')).length).toBe(1);
  await send(other, { type: 'companion', id: (await ofType(other, 'companionStart'))[0]!['id'], error: 'The agent\'s companion is not started in Restricted Mode.' });
  await expect(other.getByTestId('status')).toContainText('Restricted Mode');
  await expect(other.getByTestId('room-status')).toHaveCount(0);
});

test('S5: open { join } joins a room on the companion\'s relay and posts its base; a hand-over\'s resume with clean opens saved, in its room (P42)', async ({ page, browser }) => {
  const bytes = await fixture('HelloWordOnline.docx');
  await bootByScript(page, { agent: 'Agent' });
  await vscodePage(page);
  await open(page, 'hello.docx', bytes);
  await page.getByTestId('share').click();
  await answerStart(page);
  const link = await page.getByTestId('share-link').inputValue({ timeout: 20_000 });
  await page.keyboard.press('Escape');
  const params = new URLSearchParams(link.slice(link.indexOf('#') + 1));
  const room = { room: params.get('room')!, token: params.get('token')!, relay: params.get('relay')! };

  // The extension's Join a Room…: an untitled editor whose page joins.
  const context = await browser.newContext();
  const joined = await context.newPage();
  await bootByScript(joined, { agent: 'Agent' });
  await vscodePage(joined);
  await send(joined, { type: 'open', name: 'Room.docx', join: room });
  await expect(joined.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 20_000 });
  await expect.poll(async () => (await ofType(joined, 'base')).length).toBe(1);
  expect(await ofType(joined, 'room')).toEqual([{ type: 'room', relay: PREVIEW_RELAY }]);

  // A Save As in the room with nothing after it: the target's page resumes the store saved, and rejoins.
  await caretInBlock(page, 0);
  await page.keyboard.type('Kept in the room ');
  await send(page, { type: 'save', id: 7, name: 'b.docx' });
  await expect.poll(async () => (await ofType(page, 'saved')).length, { timeout: 30_000 }).toBe(1);
  await send(page, { type: 'handoff', id: 7 });
  await expect.poll(async () => (await ofType(page, 'handoff')).length).toBe(1);
  const [handoff] = await ofType(page, 'handoff');
  expect(handoff).toMatchObject({ late: false, resume: { meta: { room } } });
  const resume = handoff!['resume'] as { update: unknown; meta: unknown };
  const target = await context.newPage();
  await bootByScript(target, { agent: 'Agent' });
  await vscodePage(target);
  // The extension names the resume as the target (`initOf`).
  await send(target, { type: 'open', name: 'b.docx', bytes: b64(bytes), resume: { update: resume.update, meta: { ...(resume.meta as object), name: 'b.docx' } }, clean: true });
  await target.locator(body).waitFor();
  await expect(target.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 20_000 });
  await expect(target.locator(`${body} > p`).first()).toContainText('Kept in the room');
  await expect(nameShown(target)).toHaveText('b.docx');
  // No agent has worked in this room, so nothing is said of one.
  await target.waitForTimeout(2000);
  await expect(target.getByTestId('status')).not.toContainText('not in the room');
  await context.close();
});

// ---- S6, the Ask box (ED-014 section 4.6): the extension's double answers askModels and ask.

const TEST_CLAUDE = { vendor: 'docx4j-test', id: 'claude-test', version: '1', name: 'Test Claude', family: 'claude-test', maxInputTokens: 100_000 };
const COPILOT_GPT = { vendor: 'copilot', id: 'gpt-x', version: '1', name: 'GPT X', family: 'gpt-x', maxInputTokens: 64_000 };

/** The page's next unanswered `askModels`, answered with these models. */
async function answerModels(page: Page, answered: Set<unknown>, models: unknown[]): Promise<void> {
  let asked: Posted | undefined;
  await expect.poll(async () => { asked = (await ofType(page, 'askModels')).find((m) => !answered.has(m['id'])); return asked !== undefined; }, { timeout: 20_000 }).toBe(true);
  answered.add(asked!['id']);
  await send(page, { type: 'askModels', id: asked!['id'], models });
}

/** The page's `ask` number `n` (from 1), once posted. */
async function nthAsk(page: Page, n: number): Promise<Posted> {
  await expect.poll(async () => (await ofType(page, 'ask')).length, { timeout: 20_000 }).toBeGreaterThanOrEqual(n);
  return (await ofType(page, 'ask'))[n - 1]!;
}

/** A VS Code model's answer as the extension posts it: the double's text over the question, the model's name, no stop reason. */
async function answerAsk(page: Page, ask: Posted, name = 'Test Claude'): Promise<void> {
  const request = ask['request'] as { prompt: string };
  await send(page, { type: 'asked', id: ask['id'], answer: { text: answerFor(request.prompt), model: name, stopReason: null } });
}

/** The key's command as VS Code's palette names it in French (ED-014 section 4.7, P49). */
const SET_KEY_FR = 'docx4j editor: Définir la clé API Anthropic…';

async function askBoxIn(page: Page) {
  await page.getByTestId('perspective-developer').click();
  const box = page.getByTestId('ask-box');
  await expect(box).toBeVisible();
  return box;
}

test('S6: VS Code\'s models and the key\'s in two groups, the first Claude the default; one message shown; Send posts its key; the answer inserted, a script unrun (P44, P46)', async ({ page }) => {
  await bootByScript(page, { askKey: true });
  await vscodePage(page);
  await open(page, 'tracked.docx', await fixture('tracked-changes.docx'));
  // Options shows whether a key is kept, with VS Code's buttons, and no key or address field; opening it lists no model.
  await page.getByLabel('Options').click();
  await expect(page.getByTestId('option-ask-key-state')).toContainText('VS Code keeps an Anthropic API key for you');
  await expect(page.getByTestId('option-ask-key')).toHaveCount(0);
  await expect(page.getByTestId('option-ask-base-url')).toHaveCount(0);
  await page.getByTestId('option-ask-key-set').click();
  await page.getByTestId('option-ask-key-settings').click();
  await page.getByLabel('Options').click();
  expect(await ofType(page, 'askKey')).toEqual([{ type: 'askKey', action: 'set' }, { type: 'askKey', action: 'settings' }]);
  expect(await ofType(page, 'askModels')).toEqual([]);
  const answered = new Set<unknown>();
  const box = await askBoxIn(page);
  await answerModels(page, answered, [COPILOT_GPT, TEST_CLAUDE]);
  const select = box.getByTestId('ask-model');
  await expect(select.locator('optgroup')).toHaveCount(2);
  await expect(select.locator('optgroup').nth(0)).toHaveAttribute('label', 'VS Code');
  await expect(select.locator('optgroup').nth(1)).toHaveAttribute('label', 'Anthropic API (your key)');
  await expect(select.locator('optgroup').nth(0).locator('option')).toHaveText(['GPT X (copilot)', 'Test Claude (docx4j-test)']);
  await expect(select.locator('optgroup').nth(1).locator('option')).toHaveText(['Claude Sonnet 5', 'Claude Opus 5', 'Claude Fable 5.1']);
  // The first Claude, though listed second.
  await expect(select).toHaveValue('vscode:docx4j-test/claude-test');
  await caretInBlock(page, 4);
  await box.getByTestId('ask-prompt').fill('Add a heading about the double');
  await box.getByTestId('ask-ask').click();
  const request = box.getByTestId('ask-request');
  await expect(request).toHaveAttribute('data-route', 'vscode');
  await expect(box.getByTestId('ask-one-message')).toContainText('Test Claude');
  await expect(box.getByTestId('ask-one-message')).toContainText('one message');
  await expect(request).toContainText('Add a heading about the double');
  await box.getByTestId('ask-send').click();
  const ask = await nthAsk(page, 1);
  expect(ask['route']).toBe('vscode');
  expect((ask['request'] as { model: string }).model).toBe('docx4j-test/claude-test');
  expect((ask['request'] as { maxTokens: number }).maxTokens).toBe(4096);
  await answerAsk(page, ask);
  await expect(box.getByTestId('ask-preview')).toContainText('The double answers: Add a heading about the double');
  await expect(page.getByTestId('status')).toContainText('Answered by Test Claude');
  const paragraphs = page.locator(`${body} > p`);
  await box.getByTestId('ask-insert').click();
  await expect(paragraphs.nth(5)).toHaveText('The double answers: Add a heading about the double');
  // An Anthropic model goes by the key's route, its bare id; the model kept as chosen.
  await select.selectOption('claude-opus-5');
  await caretInBlock(page, 4);
  await box.getByTestId('ask-send-at-once').check();
  await box.getByTestId('ask-prompt').fill('Write a script that italicises the selection');
  await box.getByTestId('ask-ask').click();
  const second = await nthAsk(page, 2);
  expect(second['route']).toBe('key');
  expect((second['request'] as { model: string }).model).toBe('claude-opus-5');
  await answerAsk(page, second, 'claude-opus-5');
  await expect(box.getByTestId('ask-answer')).toHaveAttribute('data-form', 'script');
  await expect(box.getByTestId('ask-script-note')).toContainText('has not run');
  await expect(page.getByTestId('panel-console').locator('.console-editor .view-lines')).toContainText('p.font.italic = true');
  await expect(page.getByRole('button', { name: 'Italic' })).toHaveAttribute('aria-pressed', 'false');
  // No key or address went into the page's storage, and the page asked for nothing outside.
  expect(await page.evaluate(() => Object.keys(localStorage).filter((k) => k.includes('ask.key') || k.includes('ask.baseUrl')))).toEqual([]);
});

test('S6: the failures in VS Code\'s words; Cancel posts askCancel and drops the late answer; a gone model lists again; with no model and no key, the chat route alone, naming the boot\'s command (S7, P49)', async ({ page }) => {
  // VS Code in French names the key's command in French, whatever the page's language (ED-014 section 4.7).
  await bootByScript(page, { commands: { 'docx4j.ask.setKey': SET_KEY_FR } });
  await vscodePage(page);
  await open(page, 'tracked.docx', await fixture('tracked-changes.docx'));
  await page.getByLabel('Options').click();
  await expect(page.getByTestId('option-ask-key-state')).toContainText('no Anthropic API key is set');
  await expect(page.getByTestId('option-ask-key-remove')).toBeDisabled();
  await page.getByLabel('Options').click();
  const answered = new Set<unknown>();
  const box = await askBoxIn(page);
  await answerModels(page, answered, [TEST_CLAUDE]);
  // No key: VS Code's group alone.
  await expect(box.getByTestId('ask-model').locator('optgroup')).toHaveCount(1);
  await caretInBlock(page, 4);
  await box.getByTestId('ask-send-at-once').check();
  const ask = async (n: number): Promise<Posted> => {
    await box.getByTestId('ask-prompt').fill(`Question ${n}`);
    await box.getByTestId('ask-ask').click();
    return nthAsk(page, n);
  };
  let a = await ask(1);
  await send(page, { type: 'asked', id: a['id'], failure: { kind: 'noPermissions', message: 'refused' } });
  await expect(box.getByTestId('ask-error')).toHaveText('VS Code was not allowed to use this model. Send asks again.');
  a = await ask(2);
  await send(page, { type: 'asked', id: a['id'], failure: { kind: 'tooLong', message: '12000 > 8000', tokens: 12_000, limit: 8_000 } });
  await expect(box.getByTestId('ask-error')).toHaveText('The request is too long for this model: 12,000 tokens, where it takes 8,000');
  a = await ask(3);
  await send(page, { type: 'asked', id: a['id'], failure: { kind: 'blocked', message: 'the quota is spent' } });
  await expect(box.getByTestId('ask-error')).toContainText('the quota is spent');
  // Cancel: askCancel posted, the late answer dropped.
  a = await ask(4);
  await expect(box.getByTestId('ask-sending')).toContainText('Test Claude');
  await box.getByRole('button', { name: 'Cancel' }).click();
  await expect.poll(async () => ofType(page, 'askCancel')).toEqual([{ type: 'askCancel', id: a['id'] }]);
  await expect(page.getByTestId('status')).toContainText('The request was cancelled');
  await answerAsk(page, a);
  await page.waitForTimeout(300);
  await expect(box.getByTestId('ask-answer')).toHaveCount(0);
  // A model that is gone: said, and the list asked for again.
  const before = (await ofType(page, 'askModels')).length;
  a = await ask(5);
  await send(page, { type: 'asked', id: a['id'], failure: { kind: 'notFound', message: 'docx4j-test/claude-test' } });
  await expect(box.getByTestId('ask-error')).toHaveText('This model is no longer offered: choose another');
  await expect.poll(async () => (await ofType(page, 'askModels')).length).toBeGreaterThan(before);
  await answerModels(page, answered, []);
  // Neither a model nor a key: the chat route alone, in VS Code's words.
  await expect(box.getByTestId('ask-model')).toHaveCount(0);
  await expect(box.getByTestId('ask-send-at-once')).toHaveCount(0);
  await box.getByTestId('ask-prompt').fill('Question for a chat');
  await box.getByTestId('ask-ask').click();
  await expect(box.getByTestId('ask-no-key')).toContainText('VS Code offers none');
  await expect(box.getByTestId('ask-no-key')).toContainText(`("${SET_KEY_FR}")`);
  await expect(box.getByTestId('ask-copy-prompt')).toBeVisible();
  await expect(box.getByTestId('ask-send')).toHaveCount(0);
  // A key set by the command comes back as a push: the Anthropic group appears.
  await send(page, { type: 'askKeyState', stored: true });
  await expect(box.getByTestId('ask-model').locator('optgroup')).toHaveCount(1);
  await expect(box.getByTestId('ask-model').locator('optgroup')).toHaveAttribute('label', 'Anthropic API (your key)');
  expect((await ofType(page, 'ask')).length).toBe(5);
});
