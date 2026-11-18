// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 3.5 in the browser: tables render as tables, a cell is edited, rows come and
// go through the toolbar, Tab walks the cells, a table is inserted, and every download passes
// the oracle with only the main part changed.
import { test, expect } from '@playwright/test';
import { openFixture, download, fixture, comparePackages, caretInCell, caretAtBlockEnd, caretInBlock } from './helpers.js';

const body = '.docx-editor .ProseMirror[aria-label="Document"]';

test('tables.docx renders its merges as rowspan and colspan; the untouched round trip holds', async ({ page }) => {
  await openFixture(page, 'tables.docx');
  const tables = page.locator(`${body} > .tableWrapper > table`);
  expect(await tables.count()).toBeGreaterThan(10);
  await expect(tables.nth(1).locator('td[rowspan="2"]')).toHaveCount(1);
  await expect(tables.nth(3).locator('td[colspan="2"]')).toHaveCount(1);
  // The table style's borders reach the cells.
  const border = await tables.nth(0).locator('td').first().evaluate((el) => getComputedStyle(el).borderTopStyle);
  expect(border).toBe('solid');
  const saved = await download(page);
  expect(comparePackages(await fixture('tables.docx'), saved)).toEqual([]);
});

test('type in a cell, add a row below a merge, delete a column, download: the saved file carries the edits', async ({ page }) => {
  await openFixture(page, 'tables.docx');
  const table = page.locator(`${body} > .tableWrapper > table`).nth(1);   // A | B(merged) / C
  await caretInCell(page, 3, 1, 0);   // "C"
  await page.keyboard.press('End');
  await page.keyboard.type('X');
  await expect(table.locator('td').nth(2)).toContainText('CX');
  await expect(page.getByTestId('table-tools')).toBeVisible();
  await page.getByRole('button', { name: 'Add row below' }).click();
  await expect(table.locator('tr')).toHaveCount(3);
  await expect(table.locator('td[rowspan="2"]')).toHaveCount(1);   // the merge does not extend below its last row
  await page.getByRole('button', { name: 'Delete column' }).click();   // the column holding A and CX goes; B's merge stays
  await expect(table.locator('tr').first().locator('td')).toHaveCount(1);
  const saved = await download(page);
  const { WordprocessingMLPackage, rowsOf, cellsOf, textOf } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(saved);
  const bodyEl = (await pkg.getMainDocumentPart().readContents()).body!;
  const tbl = bodyEl.content![3]!.value as { tblGrid: { gridCol?: unknown[] } };
  expect(tbl.tblGrid.gridCol).toHaveLength(1);
  const rows = rowsOf(tbl);
  expect(rows).toHaveLength(3);
  expect(textOf(cellsOf(rows[0]!.element.value)[0]!.element)).toBe('B');
  const continuation = cellsOf(rows[1]!.element.value)[0]!.element.value as { tcPr?: { vMerge?: { val?: string } } };
  expect(continuation.tcPr?.vMerge).toBeDefined();
  expect(continuation.tcPr?.vMerge?.val).not.toBe('restart');
  expect(cellsOf(rows[2]!.element.value)).toHaveLength(1);
  expect(comparePackages(await fixture('tables.docx'), saved, ['word/document.xml', 'docProps/app.xml'])).toEqual([]);
});

test('Tab walks the cells and adds a row at the end; the caret can leave the table', async ({ page }) => {
  await openFixture(page, 'tables.docx');
  const table = page.locator(`${body} > .tableWrapper > table`).nth(0);   // A B / C D
  await caretInCell(page, 1, 0, 0);
  await page.keyboard.press('Tab');   // selects the next cell's content, as Word does
  await page.keyboard.press('End');
  await page.keyboard.type('!');
  await expect(table.locator('td').nth(1)).toContainText('B!');
  await page.keyboard.press('Tab');
  await page.keyboard.press('Tab');
  await page.keyboard.press('Tab');
  await expect(table.locator('tr')).toHaveCount(3);
  await page.keyboard.type('new');
  await expect(table.locator('tr').nth(2).locator('td').first()).toContainText('new');
});

test('insert a table from the toolbar in an empty paragraph, fill it, save, reopen', async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  await caretAtBlockEnd(page, 0);
  await page.keyboard.press('Enter');
  await page.getByRole('button', { name: 'Insert table' }).click();
  // Word's dialog (ED-005 section 12.83): two columns, two rows, fixed column width Auto.
  await page.getByTestId('insert-table-columns').fill('2');
  await page.getByTestId('insert-table-rows').fill('2');
  await page.getByRole('button', { name: 'OK' }).click();
  const table = page.locator(`${body} > .tableWrapper > table`);
  await expect(table).toHaveCount(1);
  await expect(table.locator('td')).toHaveCount(4);
  await page.keyboard.type('one');
  await page.keyboard.press('Tab');
  await page.keyboard.type('two');
  await page.getByRole('button', { name: 'Header row' }).click();
  await expect(page.getByRole('button', { name: 'Header row' })).toHaveAttribute('aria-pressed', 'true');
  const saved = await download(page);
  await page.getByTestId('open').setInputFiles({ name: 'again.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(saved) });
  await expect(page.getByTestId('name')).toContainText('again.docx');
  await page.locator(body).waitFor();
  await expect(page.locator(`${body} > .tableWrapper > table td`).first()).toContainText('one');
  await page.getByTestId('perspective-developer').click();
  await page.getByTestId('tab-styles').click();
  await caretInCell(page, 1, 0, 0);
  await expect(page.getByTestId('styles-table')).toContainText('default');   // the document has no TableGrid style: the default table style
});

test('row-level content controls show as a gutter on the rows they wrap', async ({ page }) => {
  await openFixture(page, 'invoice.docx');
  const rows = page.locator(`${body} tr[data-sdt]`);
  await expect(rows).toHaveCount(1);
  await expect(rows.first()).toHaveClass(/sdt-first/);
  await expect(rows.first()).toHaveClass(/sdt-last/);
  const saved = await download(page);
  expect(comparePackages(await fixture('invoice.docx'), saved)).toEqual([]);
});

test('a nested table renders inside the outer cell, and Insert table from the toolbar in a cell nests (ED-003 section 11.13)', async ({ page }) => {
  await openFixture(page, 'nested-table.docx');
  const outer = page.locator(`${body} > .tableWrapper > table`);
  await expect(outer).toHaveCount(1);
  const nested = outer.locator('td .tableWrapper > table');
  await expect(nested).toHaveCount(1);
  await expect(nested.locator('td').first()).toHaveText('inner a1');
  // The caret in the outer table's last cell: the table group shows, its insert button nests a table there.
  await caretInCell(page, 1, 1, 1);
  await page.getByTestId('table-tools').getByRole('button', { name: 'Insert table in this cell' }).click();
  await page.getByTestId('insert-table-columns').fill('2');
  await page.getByTestId('insert-table-rows').fill('2');
  await page.getByRole('button', { name: 'OK' }).click();
  await expect(outer.locator('td .tableWrapper > table')).toHaveCount(2);
  const saved = await download(page);
  expect(comparePackages(await fixture('nested-table.docx'), saved, ['word/document.xml', 'docProps/app.xml'])).toEqual([]);
});

const undo = (page: import('@playwright/test').Page): Promise<void> => page.evaluate(() => { (window as unknown as { docx4jEditor: { run(n: string): boolean } }).docx4jEditor.run('history.undo'); });

// ED-005 section 12.73: a body that ends with a table (Word adds a paragraph on open; the editor keeps the file as it is)
// gets a gap cursor after the table, from Down or Right at the last cell's end, where typing or Enter makes the paragraph.
test('a trailing table: Down or Right from the last cell gives a gap cursor after it, typing or Enter makes a paragraph there, and the file ends table, paragraph, sectPr', async ({ page }) => {
  await openFixture(page, 'trailing-table.docx');
  const shape = (): Promise<{ last: string; text: string; count: number; sel: string; gap: number }> => page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { lastChild: { type: { name: string }; textContent: string }; childCount: number; content: { size: number } }; selection: { constructor: { name: string } } } } }).docx4jEditor;
    const d = v.state.doc;
    return { last: d.lastChild.type.name, text: d.lastChild.textContent, count: d.childCount, sel: v.state.selection.constructor.name, gap: document.querySelectorAll('.ProseMirror-gapcursor').length };
  });
  expect((await shape()).last).toBe('table');
  const lastCellEnd = (): Promise<void> => page.evaluate(() => { const v = (window as unknown as { docx4jEditor: { caretAt(p: number): void; state: { doc: { content: { size: number } } } } }).docx4jEditor; v.caretAt(v.state.doc.content.size - 4); });
  await lastCellEnd();
  await page.keyboard.press('ArrowDown');
  await expect.poll(() => shape().then((s) => s.sel)).toBe('GapCursor');
  expect((await shape()).gap).toBe(1);
  await page.keyboard.type('after');
  await expect.poll(() => shape().then((s) => `${s.last}:${s.text}`)).toBe('paragraph:after');
  await undo(page);
  await expect.poll(() => shape().then((s) => s.last)).toBe('table');
  await lastCellEnd();
  await page.keyboard.press('ArrowRight');
  await expect.poll(() => shape().then((s) => s.sel)).toBe('GapCursor');
  await page.keyboard.press('Enter');
  await expect.poll(() => shape().then((s) => `${s.last}:${s.text}:${s.count}`)).toBe('paragraph::6');   // the five blocks and the new paragraph
  await page.keyboard.type('after the table');
  // Saved, the body ends with the table, the paragraph, then the section properties.
  const saved = await download(page);
  const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(saved);
  const xml = await pkg.getMainDocumentPart().getXml();
  const tail = xml.slice(xml.lastIndexOf('</w:tbl>'));
  expect(tail).toMatch(/^<\/w:tbl>\s*<w:p[ >][\s\S]*after the table[\s\S]*<\/w:p>\s*<w:sectPr/);
});

// ED-005 section 12.78: a column added keeps the layout mode and never scrolls inside the page; Table properties switches the mode.
test('a column added: no scrollbar inside the page; fixed widths grow the table into the margin, AutoFit to window keeps it at the text width; Table properties switches', async ({ page }) => {
  await openFixture(page, 'tables.docx');
  await caretInCell(page, 1, 0, 0);
  const measure = (): Promise<{ table: number; content: number; overflowX: string; cols: number; width: string }> => page.evaluate(() => {
    const t = document.querySelector('.docx-editor .ProseMirror[aria-label="Document"] table') as HTMLElement;
    const pm = document.querySelector('.docx-editor .ProseMirror[aria-label="Document"]') as HTMLElement;
    return { table: Math.round(t.getBoundingClientRect().width), content: pm.clientWidth, overflowX: getComputedStyle(t.parentElement!).overflowX, cols: t.querySelectorAll('col').length, width: t.style.width };
  });
  const before = await measure();
  expect(before.overflowX).toBe('visible');
  // Fixed: the table grows by the column copied.
  await page.getByTestId('table-properties').click();
  await page.getByTestId('table-autofit').selectOption('fixed');
  await page.getByRole('dialog').getByRole('button', { name: 'OK' }).click();
  await page.getByTestId('table-tools').getByRole('button', { name: 'Add column after' }).click();
  await expect.poll(() => measure().then((m) => m.cols)).toBe(3);
  const fixed = await measure();
  expect(fixed.table).toBeGreaterThan(before.table + before.table / 4);
  expect(fixed.overflowX).toBe('visible');
  // AutoFit to window: the table the text width, the fourth column scaled in.
  await caretInCell(page, 1, 0, 0);
  await page.getByTestId('table-properties').click();
  await page.getByTestId('table-autofit').selectOption('window');
  await page.getByRole('dialog').getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => measure().then((m) => m.width)).toBe('100%');
  await page.getByTestId('table-tools').getByRole('button', { name: 'Add column after' }).click();
  await expect.poll(() => measure().then((m) => m.cols)).toBe(4);
  const windowed = await measure();
  // The table's wrapper is the text column widened by the cell margin it is shifted by (tables.ts), so 100% is that much more.
  expect(Math.abs(windowed.table - windowed.content)).toBeLessThanOrEqual(10);
});

// ED-005 section 12.83: the Insert Table dialog's defaults and its AutoFit choices. Each table goes after the first
// paragraph (which has text, so none replaces it), so the last made comes first.
test('the Insert Table dialog: five by two fixed by default; a column width; AutoFit to window and to contents', async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  const tableAt = (i: number): Promise<unknown> => page.evaluate((i) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { type: { name: string }; attrs: Record<string, unknown>; childCount: number; child(j: number): { childCount: number } } } } } }).docx4jEditor;
    const t = v.state.doc.child(i);
    return { type: t.type.name, rows: t.childCount, columns: t.child(0).childCount, tblPr: t.attrs['tblPr'], grid: t.attrs['tblGrid'] };
  }, i);
  const open = async (): Promise<void> => { await caretInBlock(page, 0, 1); await page.getByRole('button', { name: 'Insert table' }).click(); };
  await open();
  await expect(page.getByTestId('insert-table-columns')).toHaveValue('5');
  await expect(page.getByTestId('insert-table-rows')).toHaveValue('2');
  await expect(page.getByTestId('insert-table-fixed')).toBeChecked();
  await page.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => tableAt(1)).toMatchObject({ type: 'table', rows: 2, columns: 5, tblPr: { tblW: { type: 'auto' } } });   // Word's default table (check 42): no layout element
  // A column width: 72 pt each, three columns.
  await open();
  await page.getByTestId('insert-table-columns').fill('3');
  await page.getByTestId('insert-table-rows').fill('1');
  await page.getByTestId('insert-table-width').fill('72');
  await page.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => tableAt(1)).toMatchObject({ rows: 1, columns: 3, tblPr: { tblW: { type: 'auto' } }, grid: [1440, 1440, 1440] });
  // AutoFit to window.
  await open();
  await page.getByTestId('insert-table-window').check();
  await expect(page.getByTestId('insert-table-width')).toBeDisabled();
  await page.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => tableAt(1)).toMatchObject({ tblPr: { tblW: { w: 5000, type: 'pct' } } });
  // AutoFit to contents.
  await open();
  await page.getByTestId('insert-table-contents').check();
  await page.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => tableAt(1)).toMatchObject({ tblPr: { tblW: { w: 0, type: 'auto' } } });
  await expect.poll(() => tableAt(4)).toMatchObject({ columns: 5 });
});

test('Row shading (ED-005 section 12.128): the menu shades every cell of the caret\'s row, a merge whole, then closes and gives the text back; No colour and Undo', async ({ page }) => {
  await openFixture(page, 'tables.docx');
  const table = page.locator(`${body} > .tableWrapper > table`).nth(1);   // A | B (merged down) / C
  const background = (i: number): Promise<string> => table.locator('td').nth(i).evaluate((el) => getComputedStyle(el).backgroundColor);
  await caretInCell(page, 3, 0, 0);   // "A"
  const menu = page.getByTestId('row-shading');
  await menu.locator('summary').click();
  await expect(menu.locator('.menu-pane')).toBeVisible();
  await menu.getByRole('button', { name: 'Yellow' }).click();
  await expect(menu.locator('.menu-pane')).toBeHidden();
  const yellow = 'rgb(255, 255, 0)';
  await expect.poll(() => Promise.all([background(0), background(1)])).toEqual([yellow, yellow]);
  expect(await background(2)).not.toBe(yellow);   // C, in the row below
  await page.keyboard.type('!');   // the focus is back in the text, at the caret (the cell's start)
  await expect(table.locator('td').nth(0)).toContainText('!A');
  const saved = await download(page);
  const { WordprocessingMLPackage, rowsOf, cellsOf } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(saved);
  const bodyEl = (await pkg.getMainDocumentPart().readContents()).body!;
  const tbl = bodyEl.content![3]!.value as Parameters<typeof rowsOf>[0];
  const fills = rowsOf(tbl).map((r) => cellsOf(r.element.value).map((c) => (c.element.value as { tcPr?: { shd?: { fill?: string } } }).tcPr?.shd?.fill ?? null));
  expect(fills).toEqual([['FFFF00', 'FFFF00'], [null, 'FFFF00']]);   // the merge's continuation follows it
  // The bar under the button shows the rows' colour; No colour takes it away; Undo brings it back.
  await expect(menu.locator('summary .bar')).toHaveCSS('background-color', yellow);
  await menu.locator('summary').click();
  await page.getByTestId('row-shading-none').click();
  await expect.poll(() => background(0)).not.toBe(yellow);
  await page.keyboard.press('Control+z');
  await expect.poll(() => background(0)).toBe(yellow);
  // Escape closes the menu without a change.
  await menu.locator('summary').click();
  await page.keyboard.press('Escape');
  await expect(menu.locator('.menu-pane')).toBeHidden();
  expect(await background(0)).toBe(yellow);
});

test('a cell\'s own shading is drawn as Word 365 drew it (ED-005 section 12.128, the row-shading round): fill auto hides a style\'s colour, val nil lets it show; a merge from its restart cell', async ({ page }) => {
  const backgrounds = async (t: number): Promise<string[]> => page.locator(`${body} > .tableWrapper > table`).nth(t).locator('tr')
    .evaluateAll((rows) => rows.map((r) => getComputedStyle(r.querySelector('td')!).backgroundColor));
  const orange = 'rgb(255, 192, 0)', white = 'rgb(255, 255, 255)', yellow = 'rgb(255, 255, 0)', none = 'rgba(0, 0, 0, 0)';
  await openFixture(page, 'row-shading/shd1-style-forms.docx');
  // Word's PDF (shd1-style-forms-word365.pdf): A1 orange, A2 none, A3 orange, A4 white, A5 yellow; B1, B5 orange, the rest none.
  await expect.poll(() => backgrounds(0)).toEqual([orange, none, orange, white, yellow]);
  expect(await backgrounds(1)).toEqual([orange, none, none, none, orange, none]);
  await openFixture(page, 'row-shading/shd2-merged.docx');
  // Word's PDF: the merged cell yellow where its restart w:tc is (C, E), none where only the continuation is (D).
  const merged = (t: number): Promise<string> => page.locator(`${body} > .tableWrapper > table`).nth(t).locator('td[rowspan="2"]').evaluate((el) => getComputedStyle(el).backgroundColor);
  expect([await merged(0), await merged(1), await merged(2)]).toEqual([yellow, none, yellow]);
});

test('a left-aligned table stands where Word 365 puts it: at its indent in compatibility mode 15, below it out by the first row\'s first cell\'s left margin (ED-005 section 12.132 item 5)', async ({ page }) => {
  // Each table's first cell text against the body text, in points, from Word's PDFs (fixtures/table-indent/README.md, Run 1).
  const word: Record<string, Record<string, number>> = {
    'ind15-mode15': { 'T1 first cell': 5.8, 'T2 first cell': 20.2, 'T3 first cell': 0.5, 'T4 first cell': 20.2, 'T5 first cell': 41.8, 'T6 first cell': 56.2, 'T7 row 1': 20.2, 'T7 row 2': 5.8 },
    'ind14-mode14': { 'T1 first cell': 0, 'T2 first cell': 0, 'T3 first cell': 0, 'T4 first cell': 0, 'T5 first cell': 36, 'T6 first cell': 36, 'T7 row 1': 0, 'T7 row 2': -14.6 },
  };
  word['ind00-no-mode'] = word['ind14-mode14']!;
  for (const [name, expected] of Object.entries(word)) {
    await openFixture(page, `table-indent/${name}.docx`);
    await expect(page.locator(body).getByText('T8 after the tables.')).toBeVisible();
    const offsets = await page.locator(body).evaluate((root) => {
      const out: Record<string, number> = {};
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        const key = /^T0 body|^T\d first cell|^T7 row \d/.exec(n.textContent ?? '')?.[0];
        if (!key) continue;
        const range = document.createRange();
        range.setStart(n, 0);
        range.setEnd(n, 1);
        out[key] = range.getBoundingClientRect().left;
      }
      return out;
    });
    for (const [key, pt] of Object.entries(expected)) {
      // CSS pixels to points; the cell border puts the editor's text up to a point to the right of Word's.
      const drawn = (offsets[key]! - offsets['T0 body']!) * 0.75;
      expect(Math.abs(drawn - pt), `${name} ${key}: drawn at ${drawn.toFixed(1)}pt, Word ${pt}pt`).toBeLessThan(1);
    }
  }
});
