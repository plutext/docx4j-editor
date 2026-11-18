// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 7.6: the other surfaces (read-only in E1.e; notes and headers editable since E2.d, ED-003 section 6.6, which notes.spec.ts covers).
import { test, expect } from '@playwright/test';
import { openFixture, download, fixture, comparePackages } from './helpers.js';

test('the footnote panel lists the notes with the numbers the references carry; navigation both ways; the round trip holds', async ({ page }) => {
  await openFixture(page, 'footnote.docx');
  const panel = page.getByTestId('footnotes-panel');
  await expect(panel).toBeVisible();
  const notes = panel.getByTestId('footnotes').locator('li.note');
  await expect(notes).toHaveCount(4);   // three referenced, one orphan; the endnotes list holds one orphan endnote
  await expect(panel.getByTestId('endnotes').locator('li.note')).toHaveCount(1);
  const refs = page.locator('.docx-editor .ProseMirror .footnote-ref');
  await expect(refs).toHaveCount(3);
  // The numbers match the decoration's, in reference order; the orphan is last, numberless.
  const numbers = await refs.evaluateAll((els) => els.map((e) => e.getAttribute('data-n')));
  expect(numbers).toEqual(['1', '2', '3']);
  await expect(notes.nth(0)).toHaveAttribute('data-n', '1');
  await expect(notes.nth(2)).toHaveAttribute('data-n', '3');
  await expect(notes.nth(3)).toHaveClass(/orphan/);
  await expect(notes.nth(0).locator('.story-view p').first()).toContainText('THIRD');
  // The notes are editable (E2.d): a textbox of their own over the note's fragment.
  await expect(notes.nth(0).locator('.ProseMirror')).toHaveAttribute('contenteditable', 'true');
  await expect(notes.nth(0).locator('.ProseMirror')).toHaveAttribute('aria-label', 'Footnote 1');
  // A click on a reference in the body flashes its note.
  await refs.nth(1).click();
  await expect(notes.nth(1)).toHaveClass(/flash/);
  // A note's number selects the reference in the body.
  await notes.nth(2).locator('.note-number').click();
  await expect(page.locator('.docx-editor .ProseMirror .footnote-ref.ProseMirror-selectednode')).toHaveAttribute('data-n', '3');
  // A caret in a note names the note in the status line.
  await notes.nth(0).locator('.ProseMirror p').first().click();
  await expect(page.getByTestId('status')).toContainText('Footnote 1');
  const saved = await download(page);
  expect(comparePackages(await fixture('footnote.docx'), saved)).toEqual([]);
});

test('a document without notes shows no panel; the diagnostics count them', async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  await expect(page.getByTestId('footnotes-panel')).toHaveCount(0);
  await page.getByTestId('perspective-developer').click();
  await page.getByTestId('tab-diagnostics').click();
  await expect(page.getByTestId('diag-footnotes')).toHaveText('0');
});

test('comments sit in the margin beside their anchors, coloured by author; a click selects the anchored text; the round trip holds', async ({ page }) => {
  await openFixture(page, 'comments-two.docx');
  const bubbles = page.getByTestId('comment-bubble');
  await expect(bubbles).toHaveCount(2);
  await expect(bubbles.first().locator('.author')).toHaveText('jharrop2');
  await expect(bubbles.first().locator('.comment-body p').first()).toContainText('One comment');
  await expect(bubbles.nth(1).locator('.comment-body p').first()).toContainText('blagh');
  // Each bubble is level with its anchor, or pushed below the previous bubble.
  const anchor = page.locator('.docx-editor .ProseMirror .comment[data-comment="0"]').first();
  const a = (await anchor.boundingBox())!;
  const b0 = (await bubbles.first().boundingBox())!;
  expect(Math.abs(b0.y - a.y)).toBeLessThan(12);
  const b1 = (await bubbles.nth(1).boundingBox())!;
  expect(b1.y).toBeGreaterThanOrEqual(b0.y + b0.height);
  // The anchored text and the bubble share the author's colour.
  const anchorColor = await anchor.evaluate((el) => getComputedStyle(el).getPropertyValue('--comment-color').trim());
  const bubbleColor = await bubbles.first().evaluate((el) => getComputedStyle(el).getPropertyValue('--comment-color').trim());
  expect(anchorColor).not.toBe('');
  expect(anchorColor).toBe(bubbleColor);
  // A click on a bubble's head selects the commented text.
  await bubbles.nth(1).locator('.comment-head').first().click();
  const selected = await page.evaluate(() => {
    const s = (window as unknown as { docx4jEditor: { state: { selection: { from: number; to: number }; doc: { textBetween(a: number, b: number): string } } } }).docx4jEditor.state;
    return s.doc.textBetween(s.selection.from, s.selection.to);
  });
  expect(selected).toContain('blsgh');
  await expect(page.getByTestId('review-offer')).toBeVisible();
  const saved = await download(page);
  expect(comparePackages(await fixture('comments-two.docx'), saved)).toEqual([]);
});

test('a document without comments has no margin', async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  await expect(page.getByTestId('comments-margin')).toHaveCount(0);
});

test('headers and footers at the page boundaries: first, even and default variants, the footer above and the header below; the round trip holds', async ({ page }) => {
  await openFixture(page, 'odd-even-first-page.docx');
  const start = page.locator('.docx-editor .page-start');
  await expect(start.locator('.boundary-header .story-view')).toContainText('First page');
  const boundary = page.locator('.docx-editor .page-boundary').first();
  await expect(boundary.locator('.boundary-footer .story-view')).toContainText('My footer');
  await expect(boundary.locator('.boundary-header .story-view')).toContainText('Even header');
  await expect(boundary.locator('.boundary-header .boundary-story')).toHaveClass(/variant-even/);
  await expect(boundary.locator('.boundary-header .ProseMirror')).toHaveAttribute('contenteditable', 'true');   // editable in place (E2.d)
  await expect(page.locator('.docx-editor .page-end .boundary-footer .story-view')).toContainText('My footer');
  const saved = await download(page);
  expect(comparePackages(await fixture('odd-even-first-page.docx'), saved)).toEqual([]);
});

test('a next-page section break is a page boundary with the new section\'s header; a section without one inherits', async ({ page }) => {
  await openFixture(page, 'headers.docx');   // five sections, six header parts
  const boundaries = page.locator('.docx-editor .page-boundary');
  await expect(boundaries).toHaveCount(4);   // one marker, three next-page breaks
  await expect(boundaries.first().locator('.boundary-header .story-view')).toBeVisible();
  const saved = await download(page);
  expect(comparePackages(await fixture('headers.docx'), saved)).toEqual([]);
});

test('a simple field in the body shows its cached result', async ({ page }) => {
  await openFixture(page, 'mergefield.docx');   // eight w:fldSimple MERGEFIELDs
  const fields = page.locator('.docx-editor .ProseMirror[aria-label="Document"] .opaque-inline.field');
  await expect(fields).toHaveCount(8);
  await expect(fields.first()).not.toHaveText('⬚');
  await expect(fields.first()).toHaveAttribute('title', /MERGEFIELD/);
});

test('a page boundary whose marker sits in a table cell is a row between the rows, with its footer and header, not inside the cell (ED-003 section 11.14)', async ({ page }) => {
  await openFixture(page, 'large/anon-250.docx');
  const rows = page.locator('.docx-editor tr.page-boundary');
  await expect(rows.first()).toBeAttached({ timeout: 60_000 });
  await expect(page.locator('.docx-editor td > .page-boundary')).toHaveCount(0);
  await expect(rows.first().locator('.boundary-label')).toHaveText(/^Page \d+/);
  await expect(rows.first().locator('.boundary-footer .story-view')).toBeAttached();
  // The band is as wide as the table, not one cell: its box spans the table's box.
  const band = (await rows.first().locator('td').boundingBox())!;
  const table = (await rows.first().locator('xpath=ancestor::table[1]').boundingBox())!;
  expect(Math.abs(band.width - table.width)).toBeLessThan(4);
});

test("the document's end footer spans the text column, and a PAGE field with no cached result shows the page's number (ED-005 section 12.98)", async ({ page }) => {
  await openFixture(page, 'numbering-clauses.docx');   // a generated footer: "Updated 6/2024", a tab to a right stop at the text's edge, "Page " and a complex PAGE field with no result
  const footer = page.locator('.docx-editor .page-end .boundary-footer .story-view');
  await expect(footer).toHaveCount(1);
  await expect(footer.locator('.field')).toHaveText('1');
  const geometry = await footer.evaluate((el) => {
    const column = el.closest('.docx-editor.body')!.querySelector(':scope > .ProseMirror')!.getBoundingClientRect();
    const story = el.getBoundingClientRect();
    const p = el.querySelector('p')!;
    return { storyLeft: story.left - column.left, storyWidth: story.width - column.width, lines: Math.round(p.getBoundingClientRect().height / parseFloat(getComputedStyle(p).lineHeight)) };
  });
  // Inside the margins the footer had sat a margin further in, too narrow for its tab stop, and "Page" wrapped.
  expect(Math.abs(geometry.storyLeft)).toBeLessThan(1.5);
  expect(Math.abs(geometry.storyWidth)).toBeLessThan(1.5);
  expect(geometry.lines).toBe(1);
});
