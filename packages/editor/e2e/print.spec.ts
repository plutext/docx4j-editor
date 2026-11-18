// SPDX-License-Identifier: GPL-3.0-or-later
//
// Printing (ED-006, level 1) in the browser: what print media shows and the PDF Chromium makes of it.
import { test, expect, type Page } from '@playwright/test';
import { openFixture, pdfPages, caretInBlock } from './helpers.js';

/** A computed style of the first element matching, or of its pseudo-element. */
async function styleOf(page: Page, selector: string, property: string, pseudo?: string): Promise<string> {
  return page.locator(selector).first().evaluate((el, [p, ps]) => getComputedStyle(el, ps ?? null).getPropertyValue(p!), [property, pseudo] as const);
}

test('in print only the body shows, without the page box or what the view draws to help editing; the screen is as it was (P1)', async ({ page }) => {
  await openFixture(page, 'loadAndSave.docx');
  const chrome = ['.app-bar', '[data-testid="toolbar"]', '.status-bar', '.comments-margin'];
  for (const selector of [...chrome, '.docx-editor .page-boundary']) await expect(page.locator(selector).first(), selector).toBeVisible();
  await page.emulateMedia({ media: 'print' });
  for (const selector of [...chrome, '.docx-editor .page-boundary', '.docx-editor .section-break']) await expect(page.locator(selector).first(), selector).toBeHidden();
  const body = page.locator('.docx-editor.body');
  await expect(body).toBeVisible();
  expect(await styleOf(page, '.docx-editor.body', 'padding-top')).toBe('0px');
  expect(await styleOf(page, '.docx-editor.body', 'box-shadow')).toBe('none');
  expect(await styleOf(page, '.docx-editor.body', 'print-color-adjust')).toBe('exact');
  // A content control prints without its outline or tags; a comment without its highlight.
  const control = '.docx-editor .sdt-block, .docx-editor .sdt-inline';
  expect(await styleOf(page, control, 'outline-style')).toBe('none');
  expect(await styleOf(page, control, 'content', '::before')).toBe('none');
  expect(await styleOf(page, '.docx-editor .ProseMirror .comment', 'background-color')).toBe('rgba(0, 0, 0, 0)');
  await page.emulateMedia({ media: 'screen' });
  for (const selector of chrome) await expect(page.locator(selector).first(), selector).toBeVisible();
});

test('a long document prints at its full length, not the one screen the page area shows (P1)', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', "page.pdf is Chromium's");
  test.setTimeout(90_000);
  await openFixture(page, 'headers-footers.docx');
  await page.emulateMedia({ media: 'print' });
  const pages = pdfPages(await page.pdf({ preferCSSPageSize: true }));
  // Word laid it out on 38 pages; the browser's own pages differ (18 since tables stopped scrolling inside the page,
  // ED-005 section 12.78: a scrollbar had added its height to every wide table), but a print clipped to the screen is one or two.
  expect(pages.length).toBeGreaterThan(10);
});

test('each section prints on its own page: A4, then A4 landscape shared by a continuous section, then US Letter (P2)', async ({ page, browserName }) => {
  test.setTimeout(90_000);
  await openFixture(page, 'print-pages.docx');
  // The section breaks carry their index for the rules.
  await expect(page.locator('.docx-editor .section-break')).toHaveCount(3);
  expect(await page.locator('.docx-editor .section-break').evaluateAll((els) => els.map((e) => (e as HTMLElement).dataset['section']))).toEqual(['0', '1', '2']);
  if (browserName !== 'chromium') return;   // page.pdf is Chromium's (Firefox's pages: ED-006 section 9.1)
  // Whether Chromium's printing to PDF fires `beforeprint`, which writes the rules.
  await page.evaluate(() => { (window as unknown as { beforePrints: number }).beforePrints = 0; window.addEventListener('beforeprint', () => { (window as unknown as { beforePrints: number }).beforePrints++; }); });
  await page.emulateMedia({ media: 'print' });
  const pages = pdfPages(await page.pdf({ preferCSSPageSize: true }));
  const fired = await page.evaluate(() => (window as unknown as { beforePrints: number }).beforePrints);
  console.log(`print [chromium] page.pdf fired beforeprint ${fired} time(s); ${pages.length} pages ${pages.map((p) => `${p.width}x${p.height}`).join(' ')}`);
  const runs: { size: string; count: number }[] = [];
  for (const p of pages) {
    const size = `${p.width}x${p.height}`;
    if (runs.at(-1)?.size === size) runs.at(-1)!.count++; else runs.push({ size, count: 1 });
  }
  expect(runs.map((r) => r.size)).toEqual(['595x842', '842x595', '612x792']);
  // Sections two and three share one landscape page: the break between them is continuous.
  expect(runs[1]!.count).toBe(1);
});

test("Word's breaks and keeps print: a forced break in a paragraph, page break before, keep with next on and off, a row that cannot split, repeated header rows (P3)", async ({ page, browserName }) => {
  test.setTimeout(90_000);
  await openFixture(page, 'print-pages.docx');
  const para = (text: string) => `.docx-editor.body p:has-text("${text}")`;
  await page.emulateMedia({ media: 'print' });
  expect(await styleOf(page, '.docx-editor .page-break', 'break-before')).toBe('page');
  expect(await styleOf(page, para('This paragraph starts a page'), 'break-before')).toBe('page');
  expect(await styleOf(page, para('This paragraph keeps with the next'), 'break-after')).toBe('avoid');
  // Heading 1 keeps with the next by its style; this one turns it off.
  expect(await styleOf(page, para('Section one: A4 portrait'), 'break-after')).toBe('avoid');
  expect(await styleOf(page, para('A heading that does not keep'), 'break-after')).toBe('auto');
  // Firefox has no orphans or widows (ED-006 section 9.1).
  if (browserName === 'chromium') expect(await styleOf(page, para('A paragraph with widow control off'), 'orphans')).toBe('1');
  expect(await styleOf(page, '.docx-editor tr[data-cant-split="1"]', 'break-inside')).toBe('avoid');
  // At beforeprint the header rows are copied into a print-only thead, the originals hidden.
  await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
  expect(await styleOf(page, '.docx-editor table > thead.print-header', 'display')).toBe('table-header-group');
  expect(await page.locator('.docx-editor table > tbody > tr[data-header="1"]').evaluateAll((rows) => rows.map((r) => getComputedStyle(r).display))).toEqual(['none', 'none']);
  await page.evaluate(() => window.dispatchEvent(new Event('afterprint')));
  await expect(page.locator('.docx-editor thead.print-header')).toHaveCount(0);
  if (browserName !== 'chromium') return;
  // The two forced breaks add pages before the table; the sections' runs are as they were.
  const pages = pdfPages(await page.pdf({ preferCSSPageSize: true }));
  const sizes = pages.map((p) => `${p.width}x${p.height}`);
  expect(sizes.filter((s) => s === '595x842').length).toBeGreaterThanOrEqual(3);
  expect([...new Set(sizes)]).toEqual(['595x842', '842x595', '612x792']);
  // page.pdf fired beforeprint and afterprint: the copy is gone again.
  await expect(page.locator('.docx-editor thead.print-header')).toHaveCount(0);
});

/** `window.print` replaced, so no dialog opens: it counts its calls, notes what was ready, and closes as a dialog would. */
async function stubPrint(page: Page): Promise<void> {
  await page.evaluate(() => {
    const w = window as unknown as { prints: { rules: number; fonts: string }[] };
    w.prints = [];
    window.print = () => {
      w.prints.push({ rules: document.querySelector('style[data-docx4j-editor="print-pages"]')?.textContent?.length ?? 0, fonts: document.fonts.status });
      window.dispatchEvent(new Event('afterprint'));
    };
  });
}
const prints = (page: Page) => page.evaluate(() => (window as unknown as { prints: { rules: number; fonts: string }[] }).prints);

test('Ctrl+P, the Print button and the palette print the document, readied: fonts loaded, page rules written; the status line says what the print is (P4)', async ({ page }) => {
  await openFixture(page, 'print-pages.docx');
  await stubPrint(page);
  await caretInBlock(page, 2);
  await page.keyboard.press('Control+p');
  await expect.poll(() => prints(page).then((p) => p.length)).toBe(1);
  const [first] = await prints(page);
  expect(first!.rules).toBeGreaterThan(0);
  expect(first!.fonts).toBe('loaded');
  await expect(page.getByTestId('status')).toHaveText("Printed by the browser: its pages are not Word's, and headers, footers and comments are not printed");
  await page.getByTestId('print').click();
  await expect.poll(() => prints(page).then((p) => p.length)).toBe(2);
  await page.keyboard.press('Control+Shift+p');
  await page.getByTestId('palette').locator('input').first().fill('Print the document');
  await page.keyboard.press('Enter');
  await expect.poll(() => prints(page).then((p) => p.length)).toBe(3);
});

test('the footnotes and endnotes print after the body, labelled as in the text, prepared by the Print command and gone after it; their headings in French (P5)', async ({ page }) => {
  await openFixture(page, 'print-pages.docx');
  /** The notes as print shows them when `print` is called (print media, so the note's own mark is hidden as printed). */
  const capture = () => page.evaluate(() => {
    const w = window as unknown as { notesAtPrint?: string[]; display?: string; notesOffset?: number };
    delete w.notesAtPrint;
    window.print = () => {
      const notes = document.querySelector('.print-notes');
      w.display = notes ? getComputedStyle(notes).display : 'absent';
      // The notes stand where the body's text does, the page's margins having given way to padding on both.
      w.notesOffset = (document.querySelector('.print-notes .print-note')?.getBoundingClientRect().left ?? NaN) - document.querySelector('.docx-editor.body p')!.getBoundingClientRect().left;
      w.notesAtPrint = [...document.querySelectorAll('.print-notes .print-notes-heading, .print-notes .print-note')].map((e) => (e as HTMLElement).innerText.replace(/\s+/g, ' ').trim());
      window.dispatchEvent(new Event('afterprint'));
    };
  });
  const atPrint = () => page.evaluate(() => (window as unknown as { notesAtPrint?: string[] }).notesAtPrint);
  await expect(page.locator('.print-notes')).toHaveCount(0);
  await page.emulateMedia({ media: 'print' });
  await capture();
  // The app bar is not shown in print media: the key.
  await page.keyboard.press('Control+p');
  // Endnotes are i, ii, iii where the document states no format, as Word numbers them.
  await expect.poll(atPrint).toEqual(['Footnotes', '1 The footnote, printed after the body.', 'Endnotes', 'i The endnote, printed after the footnotes.']);
  expect(await page.evaluate(() => (window as unknown as { display?: string }).display)).toBe('block');
  expect(await page.evaluate(() => (window as unknown as { notesOffset?: number }).notesOffset)).toBeCloseTo(0, 0);
  await expect(page.locator('.print-notes')).toHaveCount(0);
  await page.emulateMedia({ media: 'screen' });
  // French: the headings are the application's words.
  const options = page.locator('details.options-menu');
  await options.locator('summary').click();
  await page.getByTestId('option-language').selectOption('fr');
  await options.locator('summary').click();
  await page.emulateMedia({ media: 'print' });
  await capture();
  // The app bar is not shown in print media: the key.
  await page.keyboard.press('Control+p');
  await expect.poll(async () => (await atPrint())?.filter((line) => !/^[0-9i]+ /.test(line))).toEqual(['Notes de bas de page', 'Notes de fin']);
});

test('a table that stands in the page margin prints whole: inside the page, the text where the section says (Jason, 2026-11-07)', async ({ page }) => {
  // Below compatibility mode 15 a table stands its first cell's margin left of the text, as Word puts it (ED-005 section
  // 12.132 item 5): table-indent's T1, 5.4pt out, in a section of one-inch margins.
  await openFixture(page, 'table-indent/ind14-mode14.docx');
  await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
  await page.emulateMedia({ media: 'print' });
  const at = await page.evaluate(() => {
    const left = (selector: string): number => document.querySelector(selector)!.getBoundingClientRect().left;
    const pageRule = [...document.querySelectorAll('style[data-docx4j-editor="print-pages"]')].map((s) => s.textContent ?? '').join('');
    return { area: left('.app'), text: left('.docx-editor.body p'), table: left('.docx-editor.body table'), pageMargin: Number(/@page docx-s0 \{[^}]*margin: \S+ \S+ \S+ (\S+)pt;/.exec(pageRule)![1]) };
  });
  // What stands left of the page's area is cut, by Chromium whatever the stylesheets say: this was the table's left border.
  expect(at.table).toBeGreaterThanOrEqual(at.area);
  expect(at.text - at.table).toBeCloseTo(5.4 * 96 / 72, 0);
  // The page's margin and the padding together are the section's margin, an inch.
  expect(at.pageMargin * 96 / 72 + at.text - at.area).toBeCloseTo(96, 0);
});

test('a new document\'s table, from Markdown, starts at the margin, as Word 365 puts a table in compatibility mode 15 (ED-005 section 12.132 item 5)', async ({ page }) => {
  await page.goto('/');
  await page.getByTestId('open').setInputFiles({ name: 'notes.md', mimeType: 'text/markdown', buffer: Buffer.from('Some text.\n\n| a | b |\n|---|---|\n| 1 | 2 |\n', 'utf8') });
  await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
  await page.emulateMedia({ media: 'print' });
  const at = await page.evaluate(() => {
    const left = (selector: string): number => document.querySelector(selector)!.getBoundingClientRect().left;
    const pageRule = [...document.querySelectorAll('style[data-docx4j-editor="print-pages"]')].map((s) => s.textContent ?? '').join('');
    return { area: left('.app'), text: left('.docx-editor.body p'), table: left('.docx-editor.body table'), pageMargin: Number(/@page docx-s0 \{[^}]*margin: \S+ \S+ \S+ (\S+)pt;/.exec(pageRule)![1]) };
  });
  expect(at.table).toBeGreaterThanOrEqual(at.area);
  expect(Math.abs(at.text - at.table)).toBeLessThan(1.5);
  // A new document's margin is an inch unless the option says otherwise.
  expect(at.pageMargin * 96 / 72 + at.text - at.area).toBeCloseTo(96, 0);
});
