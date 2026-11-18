// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-004 section 4.6's browser test: two pages in two browser contexts against the preview
// server's `/relay`. Share in one, the link opened in the other (asked about first), typing on
// both sides seen on the other, one page taken offline and edited, back online and converged,
// both downloads byte-identical part by part, and the joiner, reloaded with unsaved edits,
// recovered back into the room (section 4.5).
import { test, expect, type Page, type BrowserContext } from '@playwright/test';
import { openFixture, caretInBlock, download, comparePackages } from './helpers.js';

/** A context whose pages have the relay option set to the preview server's mount. */
async function withRelay(context: BrowserContext): Promise<void> {
  await context.addInitScript(() => { try { localStorage.setItem('docx4j-editor.relay', '/relay'); } catch { /* ignore */ } });
}

const blockText = (page: Page, i: number): Promise<string> =>
  page.evaluate((index) => (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc.child(index).textContent, i);

async function roomStatus(page: Page, status: string): Promise<void> {
  await expect(page.getByTestId('room-status')).toHaveAttribute('data-status', status, { timeout: 20_000 });
}

test('share in one page, join from the link in another, type on both, go offline and back, download the same bytes', async ({ browser }) => {
  const contextA = await browser.newContext();
  const contextB = await browser.newContext();
  await withRelay(contextA);
  await withRelay(contextB);
  const a = await contextA.newPage();
  const b = await contextB.newPage();

  // A shares.
  await openFixture(a, 'tracked-changes.docx');
  await a.getByTestId('share').click();
  const link = await a.getByTestId('share-link').inputValue();
  expect(link).toMatch(/#room=[A-Za-z0-9_-]{22}&token=[A-Za-z0-9_-]{43}&relay=ws/);
  await expect(a.getByTestId('share-relay')).toContainText('localhost:4173');
  await a.keyboard.press('Escape');
  await roomStatus(a, 'connected');

  // B opens the link: asked before anything connects, then joined.
  await b.goto(link);
  await expect(b.getByTestId('join-offer')).toContainText('localhost:4173');
  await expect(b.getByTestId('room-status')).toHaveCount(0);
  await b.getByTestId('join-accept').click();
  await b.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await roomStatus(b, 'connected');
  expect(new URL(b.url()).hash).toBe('');
  expect(await blockText(b, 4)).toBe(await blockText(a, 4));

  // Typing on both sides reaches the other.
  await caretInBlock(a, 1);
  await a.keyboard.type('Alpha ');
  await caretInBlock(b, 3);
  await b.keyboard.type('Beta ');
  await expect.poll(() => blockText(b, 1)).toMatch(/^Alpha /);
  await expect.poll(() => blockText(a, 3)).toMatch(/^Beta /);

  // B offline: edits stay local; back online, both converge.
  await contextB.setOffline(true);
  await roomStatus(b, 'offline');
  await caretInBlock(b, 4);
  await b.keyboard.type('Offline ');
  await caretInBlock(a, 5);
  await a.keyboard.type('Meanwhile ');
  await contextB.setOffline(false);
  await roomStatus(b, 'connected');
  await expect.poll(() => blockText(a, 4)).toMatch(/^Offline /);
  await expect.poll(() => blockText(b, 5)).toMatch(/^Meanwhile /);

  // Both downloads the same bytes, part by part.
  const [x, y] = [await download(a), await download(b)];
  expect(comparePackages(x, y)).toEqual([]);

  // B edits, reloads before saving, and recovers: back in the room, its edit kept, A's later ones arriving.
  await caretInBlock(b, 6);
  await b.keyboard.type('Unsaved ');
  await expect.poll(() => blockText(a, 6)).toMatch(/^Unsaved /);
  await b.waitForTimeout(500);   // the session record's debounce (300 ms)
  await b.reload();
  await b.getByTestId('recover').click();
  await b.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await roomStatus(b, 'connected');
  expect(await blockText(b, 6)).toMatch(/^Unsaved /);
  await caretInBlock(a, 7);
  await a.keyboard.type('After ');
  await expect.poll(() => blockText(b, 7)).toMatch(/^After /);

  await contextA.close();
  await contextB.close();
});

test('a link declined connects nothing; a pasted link that is not one is said so', async ({ browser }) => {
  const context = await browser.newContext();
  const page = await context.newPage();
  await page.goto('/#room=AAAAAAAAAAAAAAAAAAAAAA&token=x&relay=ws%3A%2F%2Frelay.example.org');
  await expect(page.getByTestId('join-offer')).toContainText('relay.example.org');
  await page.getByTestId('join-decline').click();
  await expect(page.getByTestId('join-offer')).toHaveCount(0);
  expect(new URL(page.url()).hash).toBe('');
  await page.getByTestId('join').click();
  await page.getByTestId('join-link').fill('https://example.org/not-a-room');
  await page.getByRole('button', { name: 'OK' }).click();
  await expect(page.getByTestId('join-error')).toBeVisible();
  await context.close();
});

test('a link to a room the relay does not hold is said as not found, opens nothing, and makes no room', async ({ browser }) => {
  const context = await browser.newContext();
  const page = await context.newPage();
  await page.goto('/#room=BBBBBBBBBBBBBBBBBBBBBB&token=CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC&relay=ws%3A%2F%2Flocalhost%3A4173%2Frelay');
  await page.getByTestId('join-accept').click();
  await expect(page.getByTestId('status')).toContainText('no such room');
  await expect(page.getByTestId('editor')).toHaveCount(0);
  await context.close();
});

/** Put the caret in a story surface through the registry (as notes.spec.ts does). */
async function caretInStory(page: Page, id: string, pos: number): Promise<void> {
  await page.evaluate(([sid, at]) => {
    const s = (window as unknown as { docx4jSurfaces: Map<string, { caretAt(a: number): void }> }).docx4jSurfaces.get(sid as string);
    if (!s) throw new Error(`no surface ${sid}`);
    s.caretAt(at as number);
  }, [id, pos] as const);
}

test('presence: each sees the other in the people list and their caret where they put it, in the body and in a footnote', async ({ browser }) => {
  const named = async (name: string): Promise<{ context: BrowserContext; page: Page }> => {
    const context = await browser.newContext();
    await withRelay(context);
    await context.addInitScript((n) => { try { localStorage.setItem('docx4j-editor.name', n); } catch { /* ignore */ } }, name);
    return { context, page: await context.newPage() };
  };
  const { context: ca, page: a } = await named('Ada Lovelace');
  const { context: cb, page: b } = await named('Grace Hopper');
  await openFixture(a, 'footnote.docx');
  await a.getByTestId('share').click();
  const link = await a.getByTestId('share-link').inputValue();
  await a.keyboard.press('Escape');
  await b.goto(link);
  await b.getByTestId('join-accept').click();
  await roomStatus(b, 'connected');

  // The people list: each sees the other's initials; the arrival is announced.
  await expect(a.getByTestId('person')).toHaveText(['GH']);
  await expect(b.getByTestId('person')).toHaveText(['AL']);
  await expect(a.getByTestId('people-announcement')).toContainText('Grace Hopper joined');

  // Ada's caret in the body's second block shows in Grace's body there, with Ada's name.
  await caretInBlock(a, 1, 2);
  const inBody = b.locator('.ProseMirror[aria-label="Document"] .peer-caret');
  await expect(inBody).toHaveCount(1);
  await expect(inBody).toContainText('Ada Lovelace');
  // The block the caret is in, from the view (the page has widgets among the blocks, so not the DOM index).
  const block = await inBody.evaluate((el) => {
    const v = (window as unknown as { docx4jEditor: { view: { posAtDOM(n: Node, o: number): number }; state: { doc: { resolve(p: number): { index(d: number): number } } } } }).docx4jEditor;
    return v.state.doc.resolve(v.view.posAtDOM(el, 0)).index(0);
  });
  expect(block).toBe(1);
  await expect(b.getByTestId('person')).toHaveAttribute('title', 'Ada Lovelace, in the body');

  // Into the first footnote: the caret leaves the body and shows in that note, named by its number.
  const note = b.getByTestId('footnotes-panel').getByTestId('footnotes').locator('li.note').nth(0);
  const id = await note.getAttribute('data-note');
  await caretInStory(a, id!, 2);
  await expect(inBody).toHaveCount(0);
  await expect(note.locator('.peer-caret')).toHaveCount(1);
  await expect(b.getByTestId('person')).toHaveAttribute('title', /Ada Lovelace, in footnote 1/);

  // Grace's caret shows in Ada's body; a click on Ada's entry for Grace goes to it.
  await caretInBlock(b, 0, 1);
  await expect(a.locator('.ProseMirror[aria-label="Document"] .peer-caret')).toContainText('Grace Hopper');
  await a.getByTestId('person').click();
  await expect(a.locator('.peer-caret.flash')).toHaveCount(1);

  // Grace leaves: Ada's list empties and says so.
  await cb.close();
  await expect(a.getByTestId('person')).toHaveCount(0, { timeout: 40_000 });
  await expect(a.getByTestId('people-announcement')).toContainText('Grace Hopper left');
  await ca.close();
});

test('undo across peers (ED-004 section 6.2): A types, B types, A presses Ctrl+Z and only A\'s typing goes, on both pages', async ({ browser }) => {
  const contextA = await browser.newContext();
  const contextB = await browser.newContext();
  await withRelay(contextA);
  await withRelay(contextB);
  const a = await contextA.newPage();
  const b = await contextB.newPage();
  await openFixture(a, 'tracked-changes.docx');
  await a.getByTestId('share').click();
  const link = await a.getByTestId('share-link').inputValue();
  await a.keyboard.press('Escape');
  await b.goto(link);
  await b.getByTestId('join-accept').click();
  await roomStatus(b, 'connected');
  const before1 = await blockText(a, 1);
  await caretInBlock(a, 1);
  await a.keyboard.type('Mine ');
  await expect.poll(() => blockText(b, 1)).toBe(`Mine ${before1}`);
  await caretInBlock(b, 3);
  await b.keyboard.type('Theirs ');
  await expect.poll(() => blockText(a, 3)).toMatch(/^Theirs /);
  await caretInBlock(a, 1);
  await a.keyboard.press('Control+z');
  await expect.poll(() => blockText(a, 1)).toBe(before1);
  await expect.poll(() => blockText(b, 1)).toBe(before1);
  expect(await blockText(a, 3)).toMatch(/^Theirs /);
  expect(await blockText(b, 3)).toMatch(/^Theirs /);
  await contextA.close();
  await contextB.close();
});
