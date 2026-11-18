// SPDX-License-Identifier: GPL-3.0-or-later
//
// Markdown (ED-005 section 12.66) in the browser: a plain paste that reads as Markdown stays plain
// and the status line offers Format as Markdown, one undo step; a `.md` file opens as a new document,
// and its download reloads.
import { test, expect, type Page } from '@playwright/test';
import { openFixture, download, caretAtBlockEnd, caretInBlock } from './helpers.js';

/** Dispatch a paste event with `text/plain` alone, as a terminal's copy (Claude Code's `/copy`) hands it over. */
async function pastePlain(page: Page, text: string): Promise<void> {
  await page.evaluate((t) => {
    const dt = new DataTransfer();
    dt.setData('text/plain', t);
    const target = document.activeElement?.closest('.ProseMirror') ?? document.querySelector('.ProseMirror[aria-label="Document"]')!;
    target.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  }, text);
}

const MARKDOWN = '## Steps\n\n1. Run `npm ci`.\n2. Read [the guide](https://example.com/guide).\n\n```\nnpm test\n```\n';

test('a plain paste that reads as Markdown is offered Format as Markdown: one undo step back to the plain text', async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  await caretAtBlockEnd(page, 0);
  await page.keyboard.press('Enter');
  await pastePlain(page, MARKDOWN);
  const body = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  await expect(body.locator('> p', { hasText: '## Steps' })).toHaveCount(1);
  const format = page.getByTestId('format-markdown');
  await expect(format).toBeVisible();
  await expect(page.getByTestId('status')).toHaveText('The pasted text looks like Markdown.');
  await format.click();
  await expect(format).toHaveCount(0);
  await expect(body.locator('> p.s-Heading2', { hasText: 'Steps' })).toHaveCount(1);
  await expect(body.locator('> p.s-HTMLPreformatted', { hasText: 'npm test' })).toHaveCount(1);
  await expect(body.locator('a.link', { hasText: 'the guide' })).toHaveCount(1);
  await expect(page.getByTestId('status')).toContainText('Pasted');
  await expect(page.getByTestId('status')).toContainText('1 list');
  // Undo: the plain text again, not the document before the paste.
  await page.keyboard.press('Control+z');
  await expect(body.locator('> p', { hasText: '## Steps' })).toHaveCount(1);
  await expect(body.locator('> p.s-Heading2')).toHaveCount(0);
});

/** The body's `k`th table: its column widths as drawn, in pixels. */
async function tableWidths(page: Page, k: number): Promise<number[]> {
  return page.locator('.docx-editor .ProseMirror[aria-label="Document"] > .tableWrapper').nth(k).evaluate((el) =>
    Array.from(el.querySelector('table')!.querySelectorAll('colgroup > col'), (c) => parseFloat((c as HTMLElement).style.width) || 0));
}

// Descriptions long enough that the second column's width decides their line count: the reader's grid gives the first
// column twelve characters' worth (a sixth of the text width), and the fit gives it its longest word instead.
const LONG = 'A long description that wraps over several lines of the cell, as a description does when the column is narrow, and goes on for long enough that a wider column saves a line of it, which is what the fit measures and what the reader cannot know from the characters alone.';
const TABLE = `| ID | Description |\n|---|---|\n| a | ${LONG} |\n| b | ${LONG} |\n| c | ${LONG} |\n`;
// Ten rows of 150 to 600 characters, so that some row loses a line when the first column gives up its sixth whatever
// the browser's font: TABLE's three rows keep their line count under CI's Liberation Sans (the run of 2026-10-20 failed
// on it), where Carlito's narrower glyphs lose one.
const WIDE_ROWS = Array.from({ length: 10 }, (_, i) => `${LONG} ${LONG} ${LONG}`.slice(0, 150 + 50 * i).replace(/\s+\S*$/, ''));
const WIDE_TABLE = '| ID | Description |\n|---|---|\n' + WIDE_ROWS.map((r, i) => `| ${'abcdefghij'[i]} | ${r} |`).join('\n') + '\n';

test('a pasted Markdown table is fit to its content after Format as Markdown (ED-007 section 8.1 item 14)', async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  await caretAtBlockEnd(page, 0);
  await page.keyboard.press('Enter');
  await pastePlain(page, TABLE);
  await page.getByTestId('format-markdown').click();
  const body = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  await expect(body.locator('> .tableWrapper table')).toHaveCount(1);
  await expect(page.getByTestId('status')).toContainText('The table\'s columns were fit to its content.');
  // The reader's grid gave the short column twelve characters' worth of the text width (a sixth of it, about 100 px);
  // the fit gives it its longest word and the long column the rest.
  const widths = await tableWidths(page, 0);
  expect(widths).toHaveLength(2);
  expect(widths[0]!).toBeLessThan(80);
  expect(widths[1]!).toBeGreaterThan(widths[0]! * 5);
  // Undo takes the format and its fit back together: the plain text again.
  await page.keyboard.press('Control+z');
  await expect(body.locator('> .tableWrapper table')).toHaveCount(0);
  await expect(body.locator('> p', { hasText: '| ID | Description |' })).toHaveCount(1);
});

test('a .md file\'s table is fit at open, and Undo takes nothing back', async ({ page }) => {
  await page.goto('/');
  await page.getByTestId('open').setInputFiles({ name: 'table.md', mimeType: 'text/markdown', buffer: Buffer.from('# Table\n\n' + WIDE_TABLE, 'utf8') });
  const body = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  await expect(body.locator('> p.s-Heading1', { hasText: 'Table' })).toHaveCount(1);
  // Ten rows measured over many frames: longer than the default wait.
  await expect(page.getByTestId('status')).toHaveText('Opened table.md as a new document The table\'s columns were fit to its content.', { timeout: 20_000 });
  await expect(body.locator('> .tableWrapper table')).toHaveCount(1);
  const widths = await tableWidths(page, 0);
  expect(widths[0]!).toBeLessThan(80);
  await caretInBlock(page, 0);
  await page.keyboard.press('Control+z');
  await expect(body.locator('> p.s-Heading1', { hasText: 'Table' })).toHaveCount(1);
  expect(await tableWidths(page, 0)).toEqual(widths);
});

test('with Narrow margins chosen in Options, a .md file opens with them and its table is fit to the wider text width (ED-007 section 8.1 item 15)', async ({ page }) => {
  await page.goto('/');
  await page.getByLabel('Options').click();
  await page.getByTestId('option-margins').selectOption('narrow');
  await page.keyboard.press('Escape');
  await page.getByTestId('open').setInputFiles({ name: 'table.md', mimeType: 'text/markdown', buffer: Buffer.from('# Table\n\n' + WIDE_TABLE, 'utf8') });
  const body = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  await expect(page.getByTestId('status')).toHaveText('Opened table.md as a new document The table\'s columns were fit to its content.', { timeout: 20_000 });
  const pgMar = await page.evaluate(() => (window as unknown as { docx4jEditor: { state: { doc: { attrs: { finalSectPr: { pgMar: Record<string, number> } } } } } }).docx4jEditor.state.doc.attrs.finalSectPr.pgMar);
  expect(pgMar).toMatchObject({ top: 720, right: 720, bottom: 720, left: 720 });
  // A4's text width at 1.27 cm sides is 10,466 twips: 523 pt, 698 px; at Normal's it is 9,026 twips, 602 px. The fit
  // spreads the long column over the width the narrow margins leave, past what Normal's would allow.
  await expect(body.locator('> .tableWrapper table')).toHaveCount(1);
  const widths = await tableWidths(page, 0);
  expect(widths[0]!).toBeLessThan(80);
  expect(widths.reduce((a, b) => a + b, 0)).toBeGreaterThan(650);
  expect(widths.reduce((a, b) => a + b, 0)).toBeLessThanOrEqual(699);
});

test('the offer goes when the person types', async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  await caretAtBlockEnd(page, 0);
  await pastePlain(page, '- one\n- two');
  await expect(page.getByTestId('format-markdown')).toBeVisible();
  await page.keyboard.type('x');
  await expect(page.getByTestId('format-markdown')).toHaveCount(0);
});

test('a .md file opens as a new document, saved as .docx', async ({ page }) => {
  await page.goto('/');
  const text = '# Notes\n\nSome *emphasis* and `code`.\n\n> A quote.\n\n![a picture](p.png)\n';
  await page.getByTestId('open').setInputFiles({ name: 'notes.md', mimeType: 'text/markdown', buffer: Buffer.from(text, 'utf8') });
  const body = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  await expect(body.locator('> p.s-Heading1', { hasText: 'Notes' })).toHaveCount(1);
  await expect(body.locator('> p.s-Quote', { hasText: 'A quote.' })).toHaveCount(1);
  await expect(page.getByTestId('status')).toHaveText('Opened notes.md as a new document; dropped 1 image');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByTestId('download').click()]);
  expect(dl.suggestedFilename()).toBe('notes.docx');
  // Ctrl+S saves too (the download, with no file handle), and the browser's Save Page dialog never opens.
  await caretInBlock(page, 0);
  const [byKey] = await Promise.all([page.waitForEvent('download'), page.keyboard.press('Control+s')]);
  expect(byKey.suggestedFilename()).toBe('notes.docx');
  // The download opens again as a document.
  const saved = await download(page);
  await page.getByTestId('open').setInputFiles({ name: 'reopened.docx', mimeType: 'application/octet-stream', buffer: Buffer.from(saved) });
  await expect(body.locator('> p.s-Heading1', { hasText: 'Notes' })).toHaveCount(1);
});
