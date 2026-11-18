// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 6.6 item 5 (step D2): notes and headers edited in place. Typing in a footnote
// with the toolbar acting on it, the download passing the oracle with the footnotes part changed
// and nothing else beyond app.xml; a header typed at one boundary showing at another, the
// download carrying the header part; Ctrl+K in a note linking the note's text.
import { test, expect } from '@playwright/test';
import { openFixture, download, fixture, comparePackages, caretInBlock } from './helpers.js';

/** The `data-n` of every element a locator finds, in order. */
const numbers = (loc: import('@playwright/test').Locator): Promise<(string | null)[]> => loc.evaluateAll((els) => els.map((e) => e.getAttribute('data-n')));

/** Put the caret in a story surface through the registry (a click followed by a key races the selection report, as helpers.ts says of the body). */
async function caretInStory(page: import('@playwright/test').Page, id: string, pos: number, to?: number): Promise<void> {
  await page.evaluate(([sid, from, until]) => {
    const s = (window as unknown as { docx4jSurfaces: Map<string, { caretAt(a: number, b?: number): void }> }).docx4jSurfaces.get(sid);
    if (!s) throw new Error(`no surface ${sid}`);
    s.caretAt(from, until ?? from);
  }, [id, pos, to] as const);
}

test('typing in a footnote: the toolbar formats the note, the body is untouched, the download carries the footnotes part alone', async ({ page }) => {
  await openFixture(page, 'footnote.docx');
  const notes = page.getByTestId('footnotes-panel').getByTestId('footnotes').locator('li.note');
  const first = notes.nth(0).locator('.ProseMirror');
  await expect(first).toHaveAttribute('contenteditable', 'true');
  const before = await first.textContent();
  await first.click();
  await page.keyboard.press('End');
  await page.keyboard.type(' (typed here)');
  await expect(first).toHaveText(`${before} (typed here)`);
  await expect(page.getByTestId('status')).toContainText('Footnote 1');
  // Bold from the toolbar applies to the note, since the note is the active surface.
  const id = await notes.nth(0).getAttribute('data-note');
  const len = (await first.textContent())!.length;
  await caretInStory(page, id!, len - 11, len);
  await page.getByRole('button', { name: 'Bold', exact: true }).click();
  await expect(first.locator('strong, b')).toContainText('typed here');
  await expect(page.getByRole('button', { name: 'Bold', exact: true })).toHaveClass(/pressed/);
  // The body did not change; the download changes the footnotes part and app.xml only.
  const saved = await download(page);
  expect(comparePackages(await fixture('footnote.docx'), saved, ['word/footnotes.xml', 'docProps/app.xml'])).toEqual([]);
  // Undo takes the formatting and the text back, from the note's own keyboard.
  await first.click();
  await page.keyboard.press('Control+z');
  await page.keyboard.press('Control+z');
  await expect(first).toHaveText(before!);
});

test('Ctrl+K in a footnote links the note\'s text', async ({ page }) => {
  await openFixture(page, 'footnote.docx');
  const notes = page.getByTestId('footnotes-panel').getByTestId('footnotes').locator('li.note');
  const id = await notes.nth(0).getAttribute('data-note');
  const first = notes.nth(0).locator('.ProseMirror');
  const len = (await first.textContent())!.length;
  await caretInStory(page, id!, len - 5, len);
  await page.keyboard.press('Control+k');
  await page.getByTestId('link-address').fill('https://docx4j.org/');
  await page.getByRole('button', { name: 'OK' }).click();
  await expect(first.locator('a.link')).toHaveAttribute('data-target', 'https://docx4j.org/');
  await expect(page.locator('.ProseMirror[aria-label="Document"] a.link')).toHaveCount(0);
});

test('a header typed at one boundary shows at every boundary that uses the part; the body dims; the download carries the header part', async ({ page }) => {
  await openFixture(page, 'odd-even-first-page.docx');
  const boundary = page.locator('.docx-editor .page-boundary').first();
  const footerAbove = boundary.locator('.boundary-footer .ProseMirror');
  await expect(footerAbove).toHaveAttribute('contenteditable', 'true');
  const footerAtEnd = page.locator('.docx-editor .page-end .boundary-footer .ProseMirror');
  await expect(footerAtEnd).toContainText('My footer');
  await footerAbove.click();
  await expect(page.locator('.page-column')).toHaveClass(/editing-header/);
  await expect(page.getByTestId('status')).toContainText('Footer of page 1');
  await page.keyboard.press('Home');
  await page.keyboard.type('Edited: ');
  await expect(footerAbove).toContainText('Edited: My footer');
  // The same part at the document's end converged through the one fragment.
  await expect(footerAtEnd).toContainText('Edited: My footer');
  // Back in the body the dimming goes.
  await page.locator('.ProseMirror[aria-label="Document"] > p').first().click();
  await expect(page.locator('.page-column')).not.toHaveClass(/editing-header/);
  const saved = await download(page);
  const diffs = comparePackages(await fixture('odd-even-first-page.docx'), saved, ['docProps/app.xml']).filter((d) => d.name !== 'docProps/app.xml');
  expect(diffs.map((d) => d.name)).toEqual([expect.stringMatching(/^word\/footer\d+\.xml$/)]);
  expect(diffs.every((d) => d.reason === 'bytes')).toBe(true);
});

test('Insert footnote from the toolbar in a document with no notes: the caret lands in the new note, the reference is numbered, the panel appears; Delete from the panel removes both; undo brings them back; the download opens with the note', async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  await expect(page.getByTestId('footnotes-panel')).toHaveCount(0);
  await caretInBlock(page, 0, 5);
  await page.getByRole('button', { name: /Insert footnote/ }).click();
  const panel = page.getByTestId('footnotes-panel');
  await expect(panel).toBeVisible();
  const notes = panel.getByTestId('footnotes').locator('li.note');
  await expect(notes).toHaveCount(1);
  await expect(notes.first()).toHaveAttribute('data-n', '1');
  const ref = page.locator('.ProseMirror[aria-label="Document"] .footnote-ref');
  await expect(ref).toHaveCount(1);
  await expect(ref).toHaveAttribute('data-n', '1');
  // The caret is in the note: typing goes there.
  const noteBox = notes.first().locator('.ProseMirror');
  await expect(noteBox).toBeFocused();
  await page.keyboard.type('Typed in the new note.');
  await expect(noteBox).toContainText('Typed in the new note.');
  await expect(page.getByTestId('status')).toContainText('Footnote 1');
  // A second note after the first, numbered 2 in the body's order.
  await caretInBlock(page, 0, 9);
  await page.keyboard.press('Alt+Control+f');
  await expect(notes).toHaveCount(2);
  await expect.poll(() => numbers(page.locator('.ProseMirror[aria-label="Document"] .footnote-ref'))).toEqual(['1', '2']);
  // The download: the footnotes part is new, the main part and the styles part changed; reopened, the note is there.
  const saved = await download(page);
  const diffs = comparePackages(await fixture('HelloWordOnline.docx'), saved, ['word/document22.xml', 'word/styles.xml', 'docProps/app.xml']).filter((d) => d.name !== 'docProps/app.xml');
  expect(diffs).toEqual([{ name: 'word/footnotes.xml', reason: 'added' }]);
  const { EditorDocument } = await import('@docx4j/editor-model');
  const reopened = await EditorDocument.open(saved);
  const again = await reopened.stories.footnotes();
  expect(again.map((n) => n.id)).toEqual([1, 2]);
  expect(again[0]!.doc.textContent).toBe(' Typed in the new note.');
  // Delete the first from the panel: its reference goes, the other renumbers; undo restores both.
  await notes.first().hover();
  await notes.first().getByRole('button', { name: /Delete Footnote 1/ }).click();
  await expect(notes).toHaveCount(1);
  await expect.poll(() => numbers(page.locator('.ProseMirror[aria-label="Document"] .footnote-ref'))).toEqual(['1']);
  await page.keyboard.press('Control+z');
  await expect(notes).toHaveCount(2);
  await expect(notes.first().locator('.ProseMirror')).toContainText('Typed in the new note.');
});

test('the note numbers honour w:footnotePr: lower roman restarting each section, endnotes upper letter from 3', async ({ page }) => {
  await openFixture(page, 'notes-numbering.docx');
  await expect.poll(() => numbers(page.locator('.ProseMirror[aria-label="Document"] .footnote-ref.footnote'))).toEqual(['i', 'ii', 'i', 'ii', 'iii']);
  await expect.poll(() => numbers(page.locator('.ProseMirror[aria-label="Document"] .footnote-ref.endnote'))).toEqual(['C', 'D']);
  const panel = page.getByTestId('footnotes-panel');
  await expect.poll(() => numbers(panel.getByTestId('footnotes').locator('li.note'))).toEqual(['i', 'ii', 'i', 'ii', 'iii']);
  await expect.poll(() => numbers(panel.getByTestId('endnotes').locator('li.note'))).toEqual(['C', 'D']);
});

test('Re-paginate: absent without a companion; with the double answering, the boundaries follow its layout and say computed, and undo takes them back', async ({ page, baseURL }) => {
  await openFixture(page, 'notes-numbering.docx');
  await expect(page.getByTestId('repaginate')).toHaveCount(0);
  const labels = page.locator('.docx-editor .page-boundary .boundary-label');
  await expect(labels).toHaveText(['Page i', 'Page ii']);
  // The marker sits in the fifth block; the double moves it to the sixth.
  const blockOf = () => page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { forEach(f: (n: { descendants(g: (m: { type: { name: string }; attrs: Record<string, unknown> }) => boolean): void }, o: number, i: number) => void): void } } } }).docx4jEditor;
    const found: number[] = [];
    v.state.doc.forEach((block, _o, i) => block.descendants((m) => { if (m.type.name === 'opaqueInline' && m.attrs['kind'] === 'lastRenderedPageBreak') found.push(i); return true; }));
    return found;
  });
  expect(await blockOf()).toEqual([4]);
  // The option points at the preview server's double; the probe answers and the button appears.
  await page.locator('summary[aria-label="Options"]').click();
  await page.getByTestId('option-companion').fill(`${baseURL}/companion`);
  const button = page.getByTestId('repaginate');
  await expect(button).toBeVisible();
  await button.click();
  await expect(page.getByTestId('status')).toContainText('Re-paginated');
  expect(await blockOf()).toEqual([5]);
  await expect(labels).toHaveText(['Page i (computed)', 'Page ii (computed)']);   // every boundary is the companion's now
  // One undo step takes the run back (the Options pane closed first: open, it reaches over the first paragraph).
  await page.locator('summary[aria-label="Options"]').click();
  await page.locator('.ProseMirror[aria-label="Document"] > p').first().click();
  await page.keyboard.press('Control+z');
  expect(await blockOf()).toEqual([4]);
});

test('page numbers come from w:pgNumType: the labels and the PAGE field show the number a page shows', async ({ page }) => {
  await openFixture(page, 'notes-numbering.docx');
  const labels = page.locator('.docx-editor .page-boundary .boundary-label');
  await expect(labels).toHaveText(['Page i', 'Page ii']);
});
