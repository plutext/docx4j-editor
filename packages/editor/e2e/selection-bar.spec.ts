// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.74: Word's selection bar. A click in the page's left margin selects the line beside
// it; a paragraph on one line is selected whole with its mark, so typing over it joins the next paragraph.
import { test, expect, type Page } from '@playwright/test';
import { openFixture } from './helpers.js';

const body = '.docx-editor .ProseMirror[aria-label="Document"]';

const selection = (page: Page): Promise<{ from: number; to: number; start: number; end: number; nextStart: number; block: number }> => page.evaluate(() => {
  const v = (window as unknown as { docx4jEditor: { state: { doc: { resolve(p: number): { start(): number; end(): number; after(): number; index(d: number): number }; childCount: number }; selection: { from: number; to: number } } } }).docx4jEditor;
  const { from, to } = v.state.selection;
  const $from = v.state.doc.resolve(from);
  return { from, to, start: $from.start(), end: $from.end(), nextStart: $from.after() + 1, block: $from.index(0) };
});

async function clickMargin(page: Page, el: import('@playwright/test').Locator, line = 0): Promise<void> {
  await el.scrollIntoViewIfNeeded();   // the mouse reaches the viewport alone
  const box = (await el.boundingBox())!;
  const lineHeight = await el.evaluate((e) => parseFloat(getComputedStyle(e).lineHeight));
  const pageBox = (await page.locator('.docx-editor').first().boundingBox())!;
  const editorBox = (await page.locator(body).boundingBox())!;
  const x = (pageBox.x + editorBox.x) / 2;   // the middle of the left margin
  await page.mouse.click(x, box.y + lineHeight * (line + 0.5));
}

test('a click in the left margin selects the line: a one-line paragraph whole with its mark, a line of a longer one alone', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const paragraphs = page.locator(`${body} > p`);
  // Block 4, "Some text.", one line: selected to the next paragraph's start.
  await clickMargin(page, paragraphs.nth(4));
  await expect.poll(() => selection(page)).toMatchObject({ block: 4 });
  let s = await selection(page);
  expect(s.from).toBe(s.start);
  expect(s.to).toBe(s.nextStart);
  // A paragraph of several lines: its second line alone, inside the paragraph.
  const long = paragraphs.filter({ hasText: /indented indented indented/ }).first();
  const lines = await long.evaluate((e) => Math.round(e.getBoundingClientRect().height / parseFloat(getComputedStyle(e).lineHeight)));
  expect(lines).toBeGreaterThan(1);
  await clickMargin(page, long, 0);
  await expect.poll(() => selection(page).then((x) => x.from === x.start && x.to < x.end)).toBe(true);   // the first line alone
  if (lines > 2) {
    await clickMargin(page, long, 1);
    await expect.poll(() => selection(page).then((x) => x.from > x.start && x.to < x.end)).toBe(true);   // a middle line alone
  }
  // Its last line takes the mark.
  await clickMargin(page, long, lines - 1);
  await expect.poll(() => selection(page).then((x) => x.from > x.start && x.to === x.nextStart)).toBe(true);
  // Typing over the one-line paragraph's selection removes its mark: the next paragraph's text joins the typed text.
  const count = (): Promise<number> => page.evaluate(() => (window as unknown as { docx4jEditor: { state: { doc: { childCount: number } } } }).docx4jEditor.state.doc.childCount);
  const before = await count();
  const next = await paragraphs.nth(5).textContent();
  await clickMargin(page, paragraphs.nth(4));
  await expect.poll(() => selection(page)).toMatchObject({ block: 4 });
  await page.keyboard.type('x');
  await expect(paragraphs.nth(4)).toHaveText(`x${next}`);
  expect(await count()).toBe(before - 1);
});

// ED-005 section 12.81: the mark is drawn selected, a line ending in a control selects the control with it, and
// a click beside a table row selects the row.
test('the selected mark is marked on its paragraph; a paragraph ending in a control is selected whole; a click beside a row selects the row', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const paragraphs = page.locator(`${body} > p`);
  await clickMargin(page, paragraphs.nth(4));
  await expect.poll(() => selection(page)).toMatchObject({ block: 4 });
  await expect(paragraphs.nth(4)).toHaveClass(/mark-selected/);
  await expect(paragraphs.nth(5)).not.toHaveClass(/mark-selected/);
  // A caret inside the paragraph marks nothing.
  await page.evaluate(() => (window as unknown as { docx4jEditor: { caretAt(a: number): void } }).docx4jEditor.caretAt(1));
  await expect(paragraphs.nth(4)).not.toHaveClass(/mark-selected/);
  // The paragraph ends in an inline control: the margin click still selects it whole, with the control.
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { nodeSize: number } } }; caretAt(a: number): void; run(name: string, args: unknown): boolean } }).docx4jEditor;
    let pos = 0;
    for (let i = 0; i < 4; i++) pos += v.state.doc.child(i).nodeSize;
    v.caretAt(pos + v.state.doc.child(4).nodeSize - 1);
    v.run('contentControl.insert', { kind: 'PlainText' });
  });
  await expect(paragraphs.nth(4).locator('.sdt-inline')).toBeAttached();
  await clickMargin(page, paragraphs.nth(4));
  await expect.poll(() => selection(page)).toMatchObject({ block: 4 });
  const s = await selection(page);
  expect(s.from).toBe(s.start);
  expect(s.to).toBe(s.nextStart);
  await expect(paragraphs.nth(4)).toHaveClass(/mark-selected/);
  // A table row.
  await openFixture(page, 'tables.docx');
  const row = page.locator(`${body} table`).first().locator('tr').first();
  const cells = await row.locator('td, th').count();
  await clickMargin(page, row);
  await expect.poll(() => page.evaluate(() => (window as unknown as { docx4jEditor: { state: { selection: { constructor: { name: string } } } } }).docx4jEditor.state.selection.constructor.name)).toBe('CellSelection');
  await expect(page.locator(`${body} .selectedCell`)).toHaveCount(cells);
});

// ED-005 section 12.81 item 5: the selection the bar makes runs to the next paragraph's start, which a block control
// and a paragraph command do not take as that paragraph.
test('a rich text control over the bar\'s selection wraps the one paragraph, and alignment touches the one', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const paragraphs = page.locator(`${body} > p`);
  await clickMargin(page, paragraphs.nth(4));
  await expect.poll(() => selection(page)).toMatchObject({ block: 4 });
  const next = await paragraphs.nth(5).textContent();
  await page.evaluate(() => (window as unknown as { docx4jEditor: { run(name: string, args: unknown): boolean } }).docx4jEditor.run('paragraph.alignment', { value: 'Centered' }));
  await expect.poll(() => page.evaluate(() => { const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { attrs: Record<string, unknown> } } } } }).docx4jEditor; return [v.state.doc.child(4).attrs['jc'], v.state.doc.child(5).attrs['jc']]; })).toEqual(['center', null]);
  await page.evaluate(() => (window as unknown as { docx4jEditor: { run(name: string, args: unknown): boolean } }).docx4jEditor.run('contentControl.insert', { kind: 'RichText' }));
  await expect.poll(() => page.evaluate(() => { const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { type: { name: string }; childCount: number; textContent: string } } } } }).docx4jEditor; const b = v.state.doc.child(4); return [b.type.name, b.childCount, v.state.doc.child(5).type.name, v.state.doc.child(5).textContent]; })).toEqual(['sdtBlock', 1, 'paragraph', next]);
});


// ED-005 section 12.107: Word's column selection. A click on a table's top border, above a column, selects the
// column; a drag along the border takes in the columns crossed.
test('a click on a table\'s top border selects the column, and a drag along it the columns crossed', async ({ page }) => {
  await openFixture(page, 'tables.docx');
  // The first table with three columns or more (the fixture's first has two).
  const tables = page.locator(`${body} table`);
  let index = 0;
  while ((await tables.nth(index).locator('colgroup col').count()) < 3) index++;
  const table = tables.nth(index);
  await table.scrollIntoViewIfNeeded();
  const box = (await table.boundingBox())!;
  const cols = await table.locator('colgroup col').evaluateAll((els) => els.map((e) => { const b = e.getBoundingClientRect(); return { left: b.left, right: b.right }; }));
  const middle = (c: { left: number; right: number }): number => (c.left + c.right) / 2;
  /** The cells whose box holds `x`, as the count of cells a column selection at `x` must take. */
  const cellsAt = (xs: number[]): Promise<number> => table.locator('td, th').evaluateAll((els, xs) => els.filter((e) => { const b = e.getBoundingClientRect(); return xs.some((x) => x >= b.left && x <= b.right); }).length, xs);
  const kind = (): Promise<{ name: string; column: boolean }> => page.evaluate(() => {
    const s = (window as unknown as { docx4jEditor: { state: { selection: { constructor: { name: string }; isColSelection?: () => boolean } } } }).docx4jEditor.state.selection;
    return { name: s.constructor.name, column: s.isColSelection?.() ?? false };
  });
  // The pointer on the border shows the arrow; a click selects the second column.
  await page.mouse.move(middle(cols[1]!), box.y + 1);
  await expect(table).toHaveClass(/column-arrow/);
  await page.mouse.down();
  await page.mouse.up();
  await expect.poll(kind).toEqual({ name: 'CellSelection', column: true });
  await expect(page.locator(`${body} .selectedCell`)).toHaveCount(await cellsAt([middle(cols[1]!)]));
  // A drag from the first column to the third: three columns.
  await page.mouse.move(middle(cols[0]!), box.y + 1);
  await page.mouse.down();
  await page.mouse.move(middle(cols[1]!), box.y + 1);
  await page.mouse.move(middle(cols[2]!), box.y + 1);
  await page.mouse.up();
  await expect.poll(kind).toEqual({ name: 'CellSelection', column: true });
  await expect(page.locator(`${body} .selectedCell`)).toHaveCount(await cellsAt([middle(cols[0]!), middle(cols[1]!), middle(cols[2]!)]));
  // Off the border, the arrow goes and a click puts a caret in the cell.
  await page.mouse.move(middle(cols[1]!), box.y + 20);
  await expect(table).not.toHaveClass(/column-arrow/);
  await page.mouse.click(middle(cols[1]!), box.y + 20);
  await expect.poll(kind).toMatchObject({ name: 'TextSelection' });
});
