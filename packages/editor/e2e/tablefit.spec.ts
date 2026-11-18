// SPDX-License-Identifier: GPL-3.0-or-later
//
// Fit columns (ED-007): each fixture table measured and pressed is shorter, the download holds the widths, undo
// restores, the status line reports, and a second press writes nothing.
import { mkdir, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { test, expect, type Page } from '@playwright/test';
import { openFixture, openBytes, download, fixture, comparePackages, caretInCell, caretInBlock } from './helpers.js';

/** Where the fit files are kept for ED-007 section 8.2's Word check (Playwright clears it at each run). */
const outDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'test-results', 'table-fit');

const body = '.docx-editor .ProseMirror[aria-label="Document"]';

/** The height of the body's `k`th table and its column widths as drawn (the widgets between the blocks are not counted). */
async function tableGeometry(page: Page, k: number): Promise<{ height: number; widths: number[] }> {
  return page.locator(`${body} > .tableWrapper`).nth(k).evaluate((el) => {
    const table = el.querySelector('table')!;
    const widths = Array.from(table.querySelectorAll('colgroup > col'), (c) => parseFloat((c as HTMLElement).style.width) || 0);
    return { height: table.getBoundingClientRect().height, widths };
  });
}

async function fit(page: Page, block: number, r = 0, c = 0): Promise<void> {
  await caretInCell(page, block, r, c);
  await page.getByTestId('table-fit').click();
  await expect(page.locator('.status-bar')).toContainText(/shorter|longest words|No other column widths|too large|cancelled/);
}

for (const name of ['table-fit-fixed.docx', 'table-fit-auto.docx']) {
  test(`${name}: each table is shorter after Fit columns, the second press writes nothing, undo restores, the download holds the widths`, async ({ page }) => {
    await openFixture(page, `tables/${name}`);
    const tables = [1, 4, 7, 10];   // the document's blocks: a heading, the table, an empty paragraph, four times
    const before = await Promise.all(tables.map((_, k) => tableGeometry(page, k)));
    const after: { height: number; widths: number[] }[] = [];
    for (const [k, i] of tables.entries()) {
      await fit(page, i);
      await expect(page.locator('.status-bar')).toContainText(/shorter/);
      // Table C's exact and minimum row heights are removed with the widths (section 3.5), and said.
      if (i === 7) await expect(page.locator('.status-bar')).toContainText(/heights of 2 rows were removed/);
      else await expect(page.locator('.status-bar')).not.toContainText(/removed/);
      const now = await tableGeometry(page, k);
      expect(now.height, `table ${k}`).toBeLessThan(before[k]!.height);
      // The widths sum to what they did (the table keeps its width), the long column the widest.
      const sum = (ws: number[]): number => ws.reduce((a, b) => a + b, 0);
      expect(Math.abs(sum(now.widths) - sum(before[k]!.widths)) < 2 || name.includes('auto')).toBe(true);
      after.push(now);
    }
    // Table A: the long middle column took the width; table D: its third.
    expect(after[0]!.widths[1]).toBeGreaterThan(after[0]!.widths[0]! * 2);
    expect(after[3]!.widths[2]).toBeGreaterThan(after[3]!.widths[0]! * 2);
    // Table C's exact row stays its height, and its long identifier's column is at least its word's width.
    expect(after[2]!.widths[0]).toBeGreaterThan(100);
    // A second press writes nothing: the status line says so and the geometry holds.
    await fit(page, 1);
    await expect(page.locator('.status-bar')).toContainText(/No other column widths/);
    expect((await tableGeometry(page, 0)).widths).toEqual(after[0]!.widths);
    // The report stays while the caret moves within its table, and goes when the caret leaves it.
    await caretInCell(page, 1, 1, 1);
    await expect(page.locator('.status-bar')).toContainText(/No other column widths/);
    await caretInBlock(page, 2);
    await expect(page.locator('.status-bar')).not.toContainText(/No other column widths/);
    await caretInCell(page, 1, 0, 0);
    // Undo, once per table and the no-op press, restores the first table's widths.
    for (let k = 0; k < tables.length; k++) await page.keyboard.press('Control+z');
    expect((await tableGeometry(page, 0)).widths.map(Math.round)).toEqual(before[0]!.widths.map(Math.round));
    for (let k = 0; k < tables.length; k++) await page.keyboard.press('Control+y');
    expect((await tableGeometry(page, 0)).widths).toEqual(after[0]!.widths);
    // The download carries the grid and the cells as fit, and is kept for the Word check (section 8.2).
    const saved = await download(page);
    const diffs = comparePackages(await fixture(`tables/${name}`), saved, ['word/document.xml', 'docProps/app.xml']);
    expect(diffs).toEqual([]);
    await mkdir(outDir, { recursive: true });
    await writeFile(join(outDir, name.replace(/\.docx$/, '-fit.docx')), saved);
    // Opened again, the download draws every table as it was fit (section 8.1 item 11: table B's merges had lost it).
    await openBytes(page, name, saved);
    for (const [k] of tables.entries()) {
      const again = await tableGeometry(page, k);
      expect(again.widths.length, `table ${k}`).toBe(after[k]!.widths.length);
      again.widths.forEach((w, i) => expect(w, `table ${k} column ${i}`).toBeCloseTo(after[k]!.widths[i]!, 0));
    }
  });
}

test('the Table properties dialog has the fit too, and the palette lists it in a table', async ({ page }) => {
  await openFixture(page, 'tables/table-fit-fixed.docx');
  const before = await tableGeometry(page, 0);
  await caretInCell(page, 1, 1, 1);
  await page.getByTestId('table-properties').click();
  await page.getByTestId('table-fit-dialog').click();
  await expect(page.locator('.status-bar')).toContainText(/shorter/);
  expect((await tableGeometry(page, 0)).height).toBeLessThan(before.height);
  await page.keyboard.press('Control+Shift+p');
  await page.getByTestId('palette').locator('input').first().fill("Fit the table's columns");
  await expect(page.getByTestId('palette')).toContainText("Fit the table's columns to its content");
  await page.keyboard.press('Escape');
});

test('a column narrower than its longest word (a file fit under other fonts) is widened to it, and the fit says shorter', async ({ page }) => {
  await openFixture(page, 'tables/table-fit-fixed.docx');
  await caretInCell(page, 7, 1, 0);
  // Table C's identifier column squeezed under its word: the browser breaks the word at the cell's edge.
  await page.evaluate(() => (window as unknown as { docx4jEditor: { run(name: string, args: unknown): boolean } }).docx4jEditor.run('table.columnWidths', { widths: [80, 480, 40] }));
  // The lines the identifier takes: the distinct tops of its text's rectangles (a line may give several).
  const lines = () => page.locator(`${body} > .tableWrapper`).nth(2).locator('td').nth(3).evaluate((td) => {
    const range = document.createRange();
    range.selectNodeContents(td.querySelector('p')!);
    return new Set(Array.from(range.getClientRects()).filter((r) => r.width > 1).map((r) => Math.round(r.top))).size;
  });
  expect(await lines()).toBeGreaterThan(1);
  await fit(page, 7, 1, 0);
  // Shorter for the released heights here; the widened column is the point.
  await expect(page.locator('.status-bar')).toContainText(/shorter|longest words/);
  expect(await lines()).toBe(1);
});

test('a table whose longest words together are wider than it: the narrow columns are widened where that costs less height, and the status says some words still break', async ({ page }) => {
  await openFixture(page, 'tables/table-fit-fixed.docx');
  // Table E (block 13): ten columns, the fourth, seventh and eighth 37 px wide, where "price" breaks over two lines.
  const wordLines = (td: number, word: string) => page.locator(`${body} > .tableWrapper`).nth(4).locator('tr').nth(1).locator('td').nth(td).evaluate((cell, w) => {
    const walker = document.createTreeWalker(cell, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const at = (n.nodeValue ?? '').indexOf(w);
      if (at < 0) continue;
      const range = document.createRange();
      range.setStart(n, at);
      range.setEnd(n, at + w.length);
      return new Set(Array.from(range.getClientRects()).filter((r) => r.width > 1).map((r) => Math.round(r.top))).size;
    }
    return 0;
  }, word);
  expect(await wordLines(3, 'price')).toBeGreaterThan(1);
  expect(await wordLines(6, 'price')).toBeGreaterThan(1);
  const before = await tableGeometry(page, 4);
  await fit(page, 13, 1, 3);
  await expect(page.locator('.status-bar')).toContainText(/shorter/);
  await expect(page.locator('.status-bar')).toContainText(/The longest words together are wider than the table, so \d+ words? still breaks?\./);
  expect((await tableGeometry(page, 4)).height).toBeLessThan(before.height);
  for (const td of [3, 6, 7]) expect(await wordLines(td, 'price'), `column ${td}`).toBe(1);
  // The offer (section 8.1 item 13): the words still broken that have a period or a capital inside take an optional
  // break there, the table is fit again, and fewer words break; the download carries the character.
  const brokenCount = async (): Promise<number> => Number(/so (\d+) words? still break/.exec((await page.locator('.status-bar').textContent()) ?? '')?.[1] ?? 0);
  const brokenBefore = await brokenCount();
  expect(brokenBefore).toBeGreaterThan(0);
  const offer = page.getByTestId('break-words');
  await expect(offer).toHaveText(/Break \d+ words? at/);
  await offer.click();
  await expect(page.locator('.status-bar')).toContainText(/shorter|No other column widths/);
  await expect(offer).toHaveCount(0);
  expect(await brokenCount()).toBeLessThan(brokenBefore);
  const heading = await page.locator(`${body} > .tableWrapper`).nth(4).locator('tr').nth(0).textContent();
  expect(heading).toContain('Gamma.\u200BDen');
  expect(heading).toContain('Eta\u200BEastern');
});

