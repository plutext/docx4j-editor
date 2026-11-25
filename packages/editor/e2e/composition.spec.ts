// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-013 section 9, S5 defect 2: a composition inside a marked run keeps its caret. And S5's second part: Enter as
// Chrome on Android delivers it is the editor's Enter. From a composition's second
// update, ProseMirror reported the whole text node as changed and mapped the caret to its end, past a final period
// or a trailing space; Samsung Keyboard then ended the composition there. The composition is driven through
// Chromium's DevTools protocol, which the other browsers lack, as in tracking.spec.ts.
import { test, expect, devices, type Page, type CDPSession } from '@playwright/test';
import { openFixture, caretInBlock } from './helpers.js';

async function trackAs(page: Page, author: string): Promise<void> {
  await page.evaluate(async (name) => {
    const v = (window as unknown as { docx4jEditor: { session: { trackingAuthor: string | undefined }; runAsync(n: string, a: unknown): Promise<boolean> } }).docx4jEditor;
    v.session.trackingAuthor = name;
    await v.runAsync('document.setChangeTrackingMode', { value: 'TrackAll' });
  }, author);
}

async function blockWith(page: Page, text: string): Promise<number> {
  return page.evaluate((t) => {
    const doc = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc;
    for (let i = 0; i < doc.childCount; i++) if (doc.child(i).textContent.includes(t)) return i;
    return -1;
  }, text);
}

/** The block's text and the caret's offset in it. */
async function textAndCaret(page: Page, index: number): Promise<{ text: string; caret: number }> {
  return page.evaluate((i) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { textContent: string; nodeSize: number } }; selection: { head: number } } } }).docx4jEditor;
    let start = 1;
    for (let k = 0; k < i; k++) start += v.state.doc.child(k).nodeSize;
    return { text: v.state.doc.child(i).textContent, caret: v.state.selection.head - start };
  }, index);
}

async function compose(client: CDPSession, texts: string[]): Promise<void> {
  for (const text of texts) await client.send('Input.imeSetComposition', { text, selectionStart: text.length, selectionEnd: text.length });
}

test.beforeEach(({ browserName }) => {
  test.skip(browserName !== 'chromium', 'the IME is simulated over CDP, Chromium only');
});

test("a composition before a run's final period keeps its caret before the period (tracking off)", async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const index = await blockWith(page, 'Some text.');
  const offset = 'Some text'.length;
  await caretInBlock(page, index, offset);
  const client = await page.context().newCDPSession(page);
  await compose(client, ['a', 'ab', 'abc']);
  await expect.poll(() => textAndCaret(page, index)).toEqual({ text: 'Some textabc.', caret: offset + 3 });
  await client.send('Input.insertText', { text: 'abc' });
  await client.send('Input.insertText', { text: 'd' });
  await expect.poll(() => textAndCaret(page, index)).toEqual({ text: 'Some textabcd.', caret: offset + 4 });
});

test("a composition inside this author's insertion, a space after it, keeps its caret, and is one insertion (tracking on)", async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await trackAs(page, 'Author C');
  const index = await blockWith(page, 'Some text.');
  await caretInBlock(page, index, 'Some '.length);
  await page.keyboard.type('xyz ');
  const offset = 'Some xyz'.length;
  await caretInBlock(page, index, offset);
  const client = await page.context().newCDPSession(page);
  await compose(client, ['a', 'ab', 'abc']);
  await expect.poll(() => textAndCaret(page, index)).toEqual({ text: 'Some xyzabc text.', caret: offset + 3 });
  await client.send('Input.insertText', { text: 'abc' });
  await client.send('Input.insertText', { text: 'd' });
  await expect.poll(() => textAndCaret(page, index)).toEqual({ text: 'Some xyzabcd text.', caret: offset + 4 });
  await expect(page.locator('.ProseMirror[aria-label="Document"] > p').nth(index).locator('ins.insertion')).toHaveText('xyzabcd ');
});

/**
 * Enter as Chrome on Android delivers it (ED-013 section 9, S5's second part): ProseMirror leaves the key to the browser,
 * whose split it read back as a replacement of the whole paragraph, its runs lost and, with tracking on, the paragraph
 * recorded as deleted and inserted again. Under a Pixel 7's profile (ProseMirror's `android` and `chrome` paths) in the
 * full layout; the keys go through CDP, so the browser's default actions run.
 */
test.describe('Enter as Chrome on Android delivers it', () => {
  const { defaultBrowserType: _b, ...pixel } = devices['Pixel 7'];
  void _b;
  test.use(pixel);
  test.beforeEach(async ({ page }) => {
    await page.addInitScript(() => { try { localStorage.setItem('docx4j-editor.layout', 'full'); } catch { /* ignore */ } });
  });

  async function key(client: CDPSession, name: 'Backspace' | 'Enter'): Promise<void> {
    const code = name === 'Enter' ? 13 : 8;
    await client.send('Input.dispatchKeyEvent', { type: 'rawKeyDown', key: name, code: name, windowsVirtualKeyCode: code, nativeVirtualKeyCode: code });
    if (name === 'Enter') await client.send('Input.dispatchKeyEvent', { type: 'char', key: 'Enter', text: '\r', unmodifiedText: '\r', windowsVirtualKeyCode: 13 });
    await client.send('Input.dispatchKeyEvent', { type: 'keyUp', key: name, code: name, windowsVirtualKeyCode: code, nativeVirtualKeyCode: code });
  }

  /** Blocks `index` and the next: each text node with its marks, and the paragraphs' source and copy attributes. */
  async function split(page: Page, index: number): Promise<{ texts: string[][]; sourceId: (string | null)[]; copiedFrom: (string | null)[] }> {
    return page.evaluate((i) => {
      const doc = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { attrs: { sourceId: string | null; copiedFrom: string | null }; forEach(f: (n: { text?: string; marks: { type: { name: string } }[] }) => void): void } } } } }).docx4jEditor.state.doc;
      const blocks = [doc.child(i), doc.child(i + 1)];
      return {
        texts: blocks.map((b) => { const out: string[] = []; b.forEach((n) => out.push(`${n.text ?? ''}[${n.marks.map((m) => m.type.name).join(',')}]`)); return out; }),
        sourceId: blocks.map((b) => b.attrs.sourceId),
        copiedFrom: blocks.map((b) => b.attrs.copiedFrom),
      };
    }, index);
  }

  for (const tracking of [false, true]) {
    test(`Enter splits the paragraph as the editor's Enter: its runs kept, the new paragraph a copy of it${tracking ? ', nothing deleted with tracking on' : ''}`, async ({ page }) => {
      await openFixture(page, 'tracked-changes.docx');
      if (tracking) await trackAs(page, 'Phone');
      const index = await blockWith(page, 'Some text.');
      await caretInBlock(page, index, 'Some '.length);
      const client = await page.context().newCDPSession(page);
      await key(client, 'Enter');
      await expect.poll(async () => (await split(page, index)).texts).toEqual([['Some [runSource]'], ['text.[runSource]']]);
      const s = await split(page, index);
      // Enter makes the first half the new paragraph, a copy of the one it split.
      expect(s.copiedFrom[0]).toBe(s.sourceId[1]);
      expect(s.sourceId[1]).not.toBeNull();
      await expect.poll(() => textAndCaret(page, index + 1)).toEqual({ text: 'text.', caret: 0 });
    });
  }

  test('a Backspace then Enter, tracking on (S0, seen once in Gboard): the deletion stays, the split before it', async ({ page }) => {
    await openFixture(page, 'tracked-changes.docx');
    await trackAs(page, 'Phone');
    const index = await blockWith(page, 'Some text.');
    await caretInBlock(page, index, 'Some te'.length);
    const client = await page.context().newCDPSession(page);
    await key(client, 'Backspace');
    await expect.poll(async () => (await split(page, index)).texts[0]).toEqual(['Some t[runSource]', 'e[runSource,deletion]', 'xt.[runSource]']);
    await key(client, 'Enter');
    await expect.poll(async () => (await split(page, index)).texts).toEqual([['Some t[runSource]'], ['e[runSource,deletion]', 'xt.[runSource]']]);
  });
});
