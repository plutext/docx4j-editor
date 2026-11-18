// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-010: copy and paste within the editor in the browser, over the Word 365 probe of section
// 9.1. The copy is the browser's own (Ctrl+C through ProseMirror's handler, the HTML it wrote
// read back by a listener), the paste a clipboard event carrying that HTML, as paste.spec.ts
// pastes Word's captures: the system clipboard is not reached.
import { test, expect, type Page } from '@playwright/test';
import { openFixture, caretAtBlockEnd } from './helpers.js';

type Attrs = Record<string, unknown>;
type Doc = { childCount: number; child(i: number): { nodeSize: number; textContent: string; attrs: Attrs } };
type Editor = { state: { doc: Doc }; caretAt(pos: number, to?: number): void };

/** Select block `i` whole, its paragraph mark included, through the view. */
async function selectWhole(page: Page, i: number): Promise<void> {
  await page.evaluate((index) => {
    const v = (window as unknown as { docx4jEditor: Editor }).docx4jEditor;
    let pos = 0;
    for (let k = 0; k < index; k++) pos += v.state.doc.child(k).nodeSize;
    v.caretAt(pos, pos + v.state.doc.child(index).nodeSize);
  }, i);
}

/** Ctrl+C on the editor, and the HTML its copy handler put on the clipboard. */
async function copy(page: Page): Promise<string> {
  await page.evaluate(() => {
    const w = window as unknown as { __copied?: string };
    w.__copied = undefined;
    document.addEventListener('copy', (e) => { w.__copied = e.clipboardData?.getData('text/html') ?? ''; }, { once: true });
  });
  await page.keyboard.press('Control+c');
  const html = await page.evaluate(() => (window as unknown as { __copied?: string }).__copied);
  expect(html).toBeTruthy();
  return html!;
}

/** A paste event carrying `html` on the focused editor. */
async function pasteHtml(page: Page, html: string): Promise<void> {
  await page.evaluate((h) => {
    const dt = new DataTransfer();
    dt.setData('text/html', h);
    dt.setData('text/plain', 'plain');
    const target = document.activeElement?.closest('.ProseMirror') ?? document.querySelector('.ProseMirror[aria-label="Document"]')!;
    target.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  }, html);
}

const block = (page: Page, i: number): Promise<{ text: string; attrs: Attrs; count: number }> => page.evaluate((index) => {
  const v = (window as unknown as { docx4jEditor: Editor }).docx4jEditor;
  const n = v.state.doc.child(index);
  return { text: n.textContent, attrs: n.attrs, count: v.state.doc.childCount };
}, i);

// The probe's blocks: 8 is the Heading 2 "Case 9", 7 the numbered "Case 8", 18 the target T9, 17 T8.
const HEADING = 8;
const NUMBERED = 7;
const T8 = 17;
const T9 = 18;

test('a heading copied whole and pasted keeps its style; undo takes it back (ED-010)', async ({ page }) => {
  await openFixture(page, 'clipboard/clipboard-probe.docx');
  const before = (await block(page, 0)).count;
  await selectWhole(page, HEADING);
  const html = await copy(page);
  expect(html).toContain('data-docx4j-slice');
  await caretAtBlockEnd(page, T9);
  await pasteHtml(page, html);
  await expect.poll(async () => (await block(page, 0)).count).toBe(before + 1);
  const pasted = await block(page, T9 + 1);
  expect(pasted.text).toBe('Case 9: HEADING TWO TEXT (copy the words alone, not the mark)');
  expect(pasted.attrs['pStyle']).toBe('Heading2');
  expect(pasted.attrs['sourceId']).toBeNull();
  expect(pasted.attrs['copiedFrom']).toBeTruthy();
  // the target kept its own paragraph (P8), and it is drawn as a heading: the paragraph's style class
  const target = await block(page, T9);
  expect(target.attrs['pStyle']).toBeNull();
  const paragraphs = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p');
  await expect(paragraphs.nth(T9 + 1)).toHaveText(/HEADING TWO TEXT/);
  await expect(paragraphs.nth(T9 + 1)).toHaveClass(/s-Heading2/);
  await page.keyboard.press('Control+z');
  await expect.poll(async () => (await block(page, 0)).count).toBe(before);
});

test('a heading selected as the selection bar selects it (to the next paragraph\'s start) pastes whole with its style (Jason, 2026-10-29)', async ({ page }) => {
  await openFixture(page, 'clipboard/clipboard-probe.docx');
  const before = (await block(page, 0)).count;
  // The selection bar's and Shift+Down's shape: from the content's start to the next paragraph's content start.
  await page.evaluate((index) => {
    const v = (window as unknown as { docx4jEditor: Editor }).docx4jEditor;
    let pos = 0;
    for (let k = 0; k < index; k++) pos += v.state.doc.child(k).nodeSize;
    v.caretAt(pos + 1, pos + v.state.doc.child(index).nodeSize + 1);
  }, HEADING);
  const html = await copy(page);
  await caretAtBlockEnd(page, T9);
  await pasteHtml(page, html);
  await expect.poll(async () => (await block(page, 0)).count).toBe(before + 1);
  const pasted = await block(page, T9 + 1);
  expect(pasted.text).toBe('Case 9: HEADING TWO TEXT (copy the words alone, not the mark)');
  expect(pasted.attrs['pStyle']).toBe('Heading2');
  expect(pasted.attrs['copiedFrom']).toBeTruthy();
  const target = await block(page, T9);
  expect(target.attrs['pStyle']).toBeNull();
  expect(target.text).not.toContain('HEADING TWO');
});

test('a numbered item pasted keeps its numbering and is drawn with a label; a clipboard from another session falls through to the HTML', async ({ page }) => {
  await openFixture(page, 'clipboard/clipboard-probe.docx');
  const before = (await block(page, 0)).count;
  await selectWhole(page, NUMBERED);
  const html = await copy(page);
  await caretAtBlockEnd(page, T8);
  await pasteHtml(page, html);
  await expect.poll(async () => (await block(page, 0)).count).toBe(before + 1);
  const pasted = await block(page, T8 + 1);
  expect(pasted.text).toBe('Case 8: NUMBERED ITEM (copy the whole paragraph, its mark included)');
  expect(pasted.attrs['numPr']).toEqual({ numId: '1', ilvl: 0 });
  const paragraphs = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p');
  await expect(paragraphs.nth(T8 + 1).locator('.list-label')).toHaveText(/2\./);

  // another session's stamp: ProseMirror's reading of the HTML, text and paragraph breaks alone
  const foreign = html.replace(/data-docx4j-session="[^"]*"/, 'data-docx4j-session="another"');
  await caretAtBlockEnd(page, T9 + 1);
  await pasteHtml(page, foreign);
  await expect.poll(async () => (await block(page, 0)).count).toBe(before + 2);
  const plain = await block(page, T9 + 2);
  expect(plain.text).toBe('Case 8: NUMBERED ITEM (copy the whole paragraph, its mark included)');
  expect(plain.attrs['numPr']).toBeNull();
});
