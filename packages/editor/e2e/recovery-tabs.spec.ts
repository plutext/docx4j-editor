// SPDX-License-Identifier: GPL-3.0-or-later
//
// Recovery with more than one tab (ED-011 section 3.6): a document open in another tab is not lost, so it is not
// offered, and an offer made before another tab took it touches nothing. Tabs of one browser share the session store,
// so these are pages of one context, where the room tests' peers are contexts of their own.
import { test, expect, type BrowserContext, type Page } from '@playwright/test';
import { openFixture, caretInBlock } from './helpers.js';

const NAME = 'tracked-changes.docx';
const DOCUMENT = '.ProseMirror[aria-label="Document"]';

/** Type, and wait for the session record to say so: it is written 300 ms after the last update. */
async function typeAndKeep(page: Page, text: string): Promise<void> {
  await caretInBlock(page, 4, 10);
  await page.keyboard.type(text);
  await page.waitForTimeout(600);
}

/** A new tab on the editor with nothing open, given longer than the check waits for a tab that is going away. */
async function newTab(context: BrowserContext): Promise<Page> {
  const tab = await context.newPage();
  await tab.goto('/');
  await tab.waitForTimeout(1500);
  return tab;
}

const blockText = (page: Page, i: number): Promise<string> =>
  page.evaluate((index) => (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc.child(index).textContent, i);

/** The session locks this browser holds, by name. */
const heldLocks = (page: Page): Promise<string[]> =>
  page.evaluate(async () => ((await navigator.locks.query()).held ?? []).map((l) => l.name ?? '').filter((n) => n.includes('session')));

test('a document open in another tab is not offered; once that tab has gone, it is', async ({ page, context }) => {
  await openFixture(page, NAME);
  await typeAndKeep(page, ' (unsaved)');
  const second = await newTab(context);
  await expect(second.getByTestId('recovery')).toHaveCount(0);
  // The first tab is left as it was, and the second can open what it likes.
  await expect(page.getByTestId('name')).toHaveText(`${NAME} •`);
  await page.close();   // no chance to say goodbye: as a crash leaves it
  const third = await newTab(context);
  await expect(third.getByTestId('recovery')).toContainText(NAME);
  await third.getByTestId('recover').click();
  await third.getByTestId('editor').locator(DOCUMENT).waitFor();
  expect(await blockText(third, 4)).toBe('Some text. (unsaved)');
});

test('a Markdown file open in another tab is not offered (Jason, 2026-11-07)', async ({ page, context }) => {
  await page.goto('/');
  await page.getByTestId('open').setInputFiles({ name: 'notes.md', mimeType: 'text/markdown', buffer: Buffer.from('# Notes\n\nOne paragraph.\n', 'utf8') });
  await page.getByTestId('editor').locator(DOCUMENT).waitFor();
  await caretInBlock(page, 1, 3);
  await page.keyboard.type('X');
  await page.waitForTimeout(600);
  await expect(page.getByTestId('name')).toHaveText('notes.docx •');
  const second = await newTab(context);
  await expect(second.getByTestId('recovery')).toHaveCount(0);
});

test('a reload still offers its own document, with another tab open beside it', async ({ page, context }) => {
  await openFixture(page, NAME);
  await typeAndKeep(page, ' (unsaved)');
  const second = await newTab(context);
  await page.reload();
  await expect(page.getByTestId('recovery')).toContainText(NAME);
  await expect(second.getByTestId('recovery')).toHaveCount(0);
});

test('an offer another tab has overtaken touches nothing: Recover says so, Discard deletes nothing', async ({ page, context }) => {
  await openFixture(page, NAME);
  await typeAndKeep(page, ' (unsaved)');
  await page.close();
  // Three tabs, as a browser restores them after a crash: each offers the document, none having taken it.
  const [first, second, third] = [await newTab(context), await newTab(context), await newTab(context)];
  for (const tab of [first, second, third]) await expect(tab.getByTestId('recovery')).toContainText(NAME);
  await first.getByTestId('recover').click();
  await first.getByTestId('editor').locator(DOCUMENT).waitFor();

  await second.getByTestId('recover').click();
  await expect(second.getByTestId('status')).toContainText(`${NAME} is open in another tab`);
  await expect(second.getByTestId('recovery')).toHaveCount(0);
  await expect(second.getByTestId('editor').locator(DOCUMENT)).toHaveCount(0);

  await third.getByTestId('discard').click();
  await expect(third.getByTestId('status')).toContainText(`${NAME} is open in another tab`);
  await expect(third.getByTestId('recovery')).toHaveCount(0);

  // What the first tab recovered is still kept: a crash now is still recovered, with what was typed since.
  await caretInBlock(first, 4, 20);
  await first.keyboard.type(' again');
  await first.waitForTimeout(600);
  await first.close();
  const last = await newTab(context);
  await expect(last.getByTestId('recovery')).toContainText(NAME);
  await last.getByTestId('recover').click();
  await last.getByTestId('editor').locator(DOCUMENT).waitFor();
  expect(await blockText(last, 4)).toBe('Some text. (unsaved) again');
});

test('a tab holds the session it has open and no other', async ({ page }) => {
  await openFixture(page, NAME);
  const before = await heldLocks(page);
  expect(before).toHaveLength(1);
  await page.getByTestId('open').setInputFiles({ name: 'other.md', mimeType: 'text/markdown', buffer: Buffer.from('# Other\n', 'utf8') });
  await expect(page.getByTestId('name')).toHaveText('other.docx');
  const after = await heldLocks(page);
  expect(after).toHaveLength(1);
  expect(after).not.toEqual(before);
  await page.getByTestId('close').click();
  await expect.poll(() => heldLocks(page)).toEqual([]);
});

test('two tabs of one browser join one room, and a third offers neither until both have gone', async ({ browser }) => {
  const peer = await browser.newContext();
  const context = await browser.newContext();
  for (const c of [peer, context]) await c.addInitScript(() => { try { localStorage.setItem('docx4j-editor.relay', '/relay'); } catch { /* ignore */ } });
  const sharer = await peer.newPage();
  await openFixture(sharer, NAME);
  await sharer.getByTestId('share').click();
  const link = await sharer.getByTestId('share-link').inputValue();
  await sharer.keyboard.press('Escape');
  await expect(sharer.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 20_000 });

  const join = async (): Promise<Page> => {
    const tab = await context.newPage();
    await tab.goto(link);
    await tab.getByTestId('join-accept').click();
    await tab.getByTestId('editor').locator(DOCUMENT).waitFor();
    await expect(tab.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 20_000 });
    return tab;
  };
  const one = await join();
  const two = await join();
  await typeAndKeep(one, ' (one)');
  await expect.poll(() => blockText(two, 4)).toContain('(one)');
  await expect.poll(() => blockText(sharer, 4)).toContain('(one)');

  await one.close();
  const third = await newTab(context);
  await expect(third.getByTestId('recovery')).toHaveCount(0);   // the second tab has it still
  await two.close();
  const fourth = await newTab(context);
  await expect(fourth.getByTestId('recovery')).toBeVisible();
  await peer.close();
  await context.close();
});

test('a recovery refused lets go of the session, so the banner can still discard it', async ({ page }) => {
  await openFixture(page, NAME);
  await typeAndKeep(page, ' (unsaved)');
  // The kept bytes made unreadable, as a session from another version is refused.
  await page.evaluate(async () => {
    const id = localStorage.getItem('docx4j-editor.session')!;
    const db = await new Promise<IDBDatabase>((resolve, reject) => { const r = indexedDB.open('docx4j-editor'); r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error); });
    const store = db.transaction('sessions', 'readwrite').objectStore('sessions');
    await new Promise<void>((resolve, reject) => {
      const get = store.get(id);
      get.onsuccess = () => { const put = store.put({ ...get.result, bytes: new Uint8Array([1, 2, 3]) }); put.onsuccess = () => resolve(); put.onerror = () => reject(put.error); };
    });
    db.close();
  });
  await page.reload();
  await page.getByTestId('recover').click();
  await expect(page.getByTestId('status')).toContainText('Not recovered');
  await expect(page.getByTestId('recovery')).toBeVisible();
  await expect.poll(() => heldLocks(page)).toEqual([]);
  await page.getByTestId('discard').click();
  await expect(page.getByTestId('recovery')).toHaveCount(0);
  await expect(page.getByTestId('status')).not.toContainText('another tab');
  await page.reload();
  await page.waitForTimeout(1500);
  await expect(page.getByTestId('recovery')).toHaveCount(0);
});

test('a room that cannot be joined lets go of its session', async ({ browser }) => {
  const peer = await browser.newContext();
  await peer.addInitScript(() => { try { localStorage.setItem('docx4j-editor.relay', '/relay'); } catch { /* ignore */ } });
  const sharer = await peer.newPage();
  await openFixture(sharer, NAME);
  await sharer.getByTestId('share').click();
  const link = await sharer.getByTestId('share-link').inputValue();
  await sharer.keyboard.press('Escape');
  const context = await browser.newContext();
  const tab = await context.newPage();
  await tab.goto(link.replace(/#room=[^&]+/, '#room=AAAAAAAAAAAAAAAAAAAAAA'));
  await tab.getByTestId('join-accept').click();
  await expect(tab.getByTestId('status')).toContainText(/room|Room/, { timeout: 25_000 });
  await expect.poll(() => heldLocks(tab), { timeout: 25_000 }).toEqual([]);
  await peer.close();
  await context.close();
});
