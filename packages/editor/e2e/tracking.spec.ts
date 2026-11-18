// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 5.4, step C3: the tracker on screen. Typing with tracking on is an insertion,
// Backspace keeps what it deletes (struck through, the browser's default for `del` until E4.d),
// this user's own insertion is deleted for real; an IME composition, driven through Chrome's
// DevTools protocol over the Japanese text of the CJK fixture, is one insertion once it ends;
// an untracked command says so in the status line.
import { test, expect, type Page } from '@playwright/test';
import { openFixture, caretInBlock } from './helpers.js';

const body = '.docx-editor .ProseMirror[aria-label="Document"]';

/** Turn tracking on through the command, as Author C. */
async function trackAs(page: Page, author: string): Promise<void> {
  await page.evaluate(async (name) => {
    const v = (window as unknown as { docx4jEditor: { session: { trackingAuthor: string | undefined }; runAsync(n: string, a: unknown): Promise<boolean> } }).docx4jEditor;
    v.session.trackingAuthor = name;
    await v.runAsync('document.setChangeTrackingMode', { value: 'TrackAll' });
  }, author);
}

/** The index of the first top-level block whose text includes `text`. */
async function blockWith(page: Page, text: string): Promise<number> {
  return page.evaluate((t) => {
    const doc = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc;
    for (let i = 0; i < doc.childCount; i++) if (doc.child(i).textContent.includes(t)) return i;
    return -1;
  }, text);
}

test('with tracking on, typing is an insertion; Backspace deletes this user\'s own typing, and keeps other text struck through', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await trackAs(page, 'Author C');
  const index = await blockWith(page, 'Some text.');
  const paragraph = page.locator(`${body} > p`).nth(index);
  await caretInBlock(page, index, 5);   // "Some |text."
  await page.keyboard.type('new ');
  await expect(paragraph.locator('ins.insertion')).toHaveText('new ');
  await expect(paragraph.locator('ins.insertion')).toHaveAttribute('data-author', 'Author C');
  await page.keyboard.press('Backspace');
  await expect(paragraph.locator('ins.insertion')).toHaveText('new');
  await expect(paragraph).toHaveText('Some newtext.');
  // Before "new": Backspace over the file's text keeps it, deleted, and the caret goes before it.
  await caretInBlock(page, index, 5);
  await page.keyboard.press('Backspace');
  await expect(paragraph.locator('del.deletion')).toHaveText(' ');
  await page.keyboard.press('Backspace');
  await expect(paragraph.locator('del.deletion')).toHaveText('e ');
  await expect(paragraph).toHaveText('Some newtext.');
});

test('an IME composition with tracking on is one insertion once it ends, and the composition is not broken', async ({ page, browserName }) => {
  // The composition is driven through Chromium's DevTools protocol, which the other browsers lack (ED-005 section 12.41).
  test.skip(browserName !== 'chromium', 'the IME is simulated over CDP, Chromium only');
  await openFixture(page, 'fonts-cjk-linebox.docx');
  await trackAs(page, 'Author C');
  const index = await blockWith(page, '日本語のテキスト');
  expect(index).toBeGreaterThanOrEqual(0);
  const offset = await page.evaluate((i) => {
    const doc = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc;
    return doc.child(i).textContent.indexOf('テキスト') + 'テキスト'.length;
  }, index);
  await caretInBlock(page, index, offset);
  const client = await page.context().newCDPSession(page);
  for (const text of ['か', 'かん', 'かんじ']) await client.send('Input.imeSetComposition', { text, selectionStart: text.length, selectionEnd: text.length });
  await client.send('Input.insertText', { text: '漢字' });
  const paragraph = page.locator(`${body} > p`).nth(index);
  await expect(paragraph.locator('ins.insertion')).toHaveText('漢字');
  const text = await page.evaluate((i) => (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc.child(i).textContent, index);
  expect(text).toContain('日本語のテキスト漢字');
  expect(text).not.toContain('かんじ');
});

test('an untracked command with tracking on says so in the status line', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await trackAs(page, 'Author C');
  const table = await page.evaluate(() => {
    const doc = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { type: { name: string } } } } } }).docx4jEditor.state.doc;
    for (let i = 0; i < doc.childCount; i++) if (doc.child(i).type.name === 'table') return i;
    return -1;
  });
  await page.evaluate((t) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { nodeSize: number } } }; caretAt(p: number): void; run(n: string, a?: unknown): boolean } }).docx4jEditor;
    let pos = 0;
    for (let k = 0; k < t; k++) pos += v.state.doc.child(k).nodeSize;
    v.caretAt(pos + 4);
    v.run('table.addColumns');
  }, table);
  await expect(page.getByTestId('status')).toContainText('is not tracked');
});
