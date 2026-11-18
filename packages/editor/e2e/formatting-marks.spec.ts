// SPDX-License-Identifier: GPL-3.0-or-later
// Formatting marks (ED-005 section 12.80): the Developer and Template perspectives draw every paragraph's pilcrow
// (a cell's last paragraph its cell mark, a tab its arrow); Edit and Review draw them under an option; the pilcrow
// takes the paragraph mark's own size and colour; a tracked mark keeps its own pilcrow; printing hides them.
import { test, expect, type Page } from '@playwright/test';
import { openFixture, caretInBlock } from './helpers';

const pseudo = (page: Page, selector: string, which: '::after' | '::before', prop = 'content'): Promise<string> =>
  page.evaluate(([s, w, p]) => { const el = document.querySelector(s); return el ? getComputedStyle(el, w).getPropertyValue(p) : 'missing'; }, [selector, which, prop] as const);

test('the pilcrow per perspective, the option, the mark\'s own size and colour, a cell mark and a tab', async ({ page }) => {
  await openFixture(page, 'tables.docx');
  const first = '.docx-editor.body .ProseMirror > p';
  // Edit: no marks until the option is on.
  await expect.poll(() => pseudo(page, first, '::after')).toBe('none');
  await page.getByTestId('perspective-developer').click();
  await expect.poll(() => pseudo(page, first, '::after')).toBe('"¶"');
  await expect.poll(() => pseudo(page, '.docx-editor.body td > p:last-child', '::after')).toBe('"¤"');
  await page.getByTestId('perspective-template').click();
  await expect.poll(() => pseudo(page, first, '::after')).toBe('"¶"');
  await page.getByTestId('perspective-edit').click();
  await expect.poll(() => pseudo(page, first, '::after')).toBe('none');
  // The option: Edit draws them too.
  await page.locator('summary[aria-label="Options"]').click();
  await page.getByTestId('option-marks').check();
  await expect.poll(() => pseudo(page, first, '::after')).toBe('"¶"');
  await page.keyboard.press('Escape');
  // The mark's own run formatting: a paragraph whose mark states a size and a colour draws its pilcrow in them.
  const at = await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { type: { name: string }; attrs: Record<string, unknown> } } }; run(name: string, args: unknown): boolean; caretAt(a: number): void } }).docx4jEditor;
    return v.state.doc.child(0).type.name === 'paragraph' ? 0 : -1;
  });
  expect(at).toBe(0);
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { view: { state: { doc: { child(i: number): { attrs: Record<string, unknown> } }; tr: { setNodeMarkup(pos: number, type: null, attrs: Record<string, unknown>): unknown } }; dispatch(tr: unknown): void } } }).docx4jEditor.view;
    const p = v.state.doc.child(0);
    v.dispatch(v.state.tr.setNodeMarkup(0, null, { ...p.attrs, markRPr: { sz: { val: 48 }, color: { val: 'FF0000' } } }));
  });
  await expect.poll(() => pseudo(page, first, '::after', 'font-size')).toBe('32px');
  await expect.poll(() => pseudo(page, first, '::after', 'color')).toBe('rgb(255, 0, 0)');
  // A paragraph ending in an inline control (ProseMirror's trailing break after it): the pilcrow on the control, not a new line.
  const heightBefore = await page.evaluate((s) => document.querySelector(s)!.getBoundingClientRect().height, first);
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { nodeSize: number } } }; caretAt(a: number): void; run(name: string, args: unknown): boolean } }).docx4jEditor;
    v.caretAt(v.state.doc.child(0).nodeSize - 1);
    v.run('contentControl.insert', { kind: 'PlainText' });
  });
  await expect(page.locator(`${first} > .sdt-inline + br.ProseMirror-trailingBreak`)).toBeAttached();
  await expect.poll(() => pseudo(page, first, '::after')).toBe('none');
  await expect.poll(() => pseudo(page, `${first} > .sdt-inline`, '::after')).toBe('"¶"');
  await expect.poll(() => page.evaluate((s) => document.querySelector(s)!.getBoundingClientRect().height, first)).toBe(heightBefore);
  // A tab's arrow.
  await page.evaluate(() => (window as unknown as { docx4jEditor: { caretAt(a: number): void } }).docx4jEditor.caretAt(1));
  await page.keyboard.press('Tab');
  await expect.poll(() => pseudo(page, '.docx-editor.body .tab', '::before')).toBe('"→"');
  // A list paragraph's hanging indent does not pull the pilcrow back.
  await page.evaluate(() => (window as unknown as { docx4jEditor: { caretAt(a: number): void; run(name: string, args: unknown): boolean } }).docx4jEditor.caretAt(1));
  await page.evaluate(() => (window as unknown as { docx4jEditor: { run(name: string, args: unknown): boolean } }).docx4jEditor.run('list.toggle', { bullet: true }));
  await expect(page.locator(`${first} .list-label`)).toBeAttached();
  await expect.poll(() => page.evaluate((s) => getComputedStyle(document.querySelector(s)!).textIndent, first)).toMatch(/^-/);
  await expect.poll(() => pseudo(page, `${first} > .sdt-inline`, '::after', 'text-indent')).toBe('0px');
  // Printing hides them.
  await page.emulateMedia({ media: 'print' });
  await expect.poll(() => pseudo(page, first, '::after')).toBe('none');
  await expect.poll(() => pseudo(page, '.docx-editor.body .tab', '::before')).toBe('none');
  await page.emulateMedia({ media: 'screen' });
});

test('a tracked paragraph mark keeps its own pilcrow in Review, and the plain ones follow the option', async ({ page }) => {
  await openFixture(page, 'revisions/revisions-word2010.docx');
  await page.getByTestId('perspective-review').click();
  const tracked = '.docx-editor.body p.rev-mark-ins';
  await expect(page.locator(tracked).first()).toBeAttached();
  await expect.poll(() => pseudo(page, tracked, '::after')).toBe('"¶"');
  await expect.poll(() => pseudo(page, tracked, '::after', 'text-decoration-line')).toBe('underline');
  await expect.poll(() => pseudo(page, '.docx-editor.body .ProseMirror > p:not(.rev-mark-ins):not(.rev-mark-del)', '::after')).toBe('none');
  // An empty inserted paragraph (Enter with tracking on): its pilcrow is its list marker, inline where text would go,
  // not a box at the paragraph's edge beside the change bar (ED-005 section 12.80 item 7).
  await page.evaluate(async () => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { nodeSize: number } } }; caretAt(a: number): void; runAsync(name: string, args: unknown): Promise<unknown> } }).docx4jEditor;
    await v.runAsync('document.setChangeTrackingMode', { value: 'TrackAll' });
    v.caretAt(v.state.doc.child(0).nodeSize - 1);
  });
  // Enter puts the new mark at the caret (core-ts check 29): the first Enter marks the title's mark inserted and the
  // new empty paragraph carries the old one; the second, in that empty paragraph, marks its mark inserted.
  await page.keyboard.press('Enter');
  await page.keyboard.press('Enter');
  const empty = '.docx-editor.body .ProseMirror > p.rev-mark-ins:has(> br.ProseMirror-trailingBreak:only-child)';
  await expect(page.locator(empty)).toBeAttached();
  await expect.poll(() => pseudo(page, empty, '::after')).toBe('"¶"');
  await expect.poll(() => page.evaluate((s) => getComputedStyle(document.querySelector(s)!.querySelector('br')!).display, empty)).toBe('none');
  const oneLine = await page.evaluate(() => (document.querySelector('.docx-editor.body .ProseMirror > p') as HTMLElement).getBoundingClientRect().height);
  await expect.poll(() => page.evaluate((s) => document.querySelector(s)!.getBoundingClientRect().height, empty)).toBeLessThanOrEqual(oneLine);
  // The caret at its start (set there: the second Enter left it in the paragraph below) takes typed text.
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { nodeSize: number; textContent: string; attrs: Record<string, unknown> } } }; caretAt(a: number): void } }).docx4jEditor;
    let pos = 0;
    for (let i = 0; i < v.state.doc.childCount; i++) {
      const b = v.state.doc.child(i);
      if (b.textContent === '' && (b.attrs['markRevision'] as { insertion?: unknown } | null)?.insertion) { v.caretAt(pos + 1); return; }
      pos += b.nodeSize;
    }
  });
  await page.keyboard.type('q');
  await expect(page.locator('.docx-editor.body .ProseMirror > p.rev-mark-ins', { hasText: /^q$/ })).toBeAttached();
  await page.keyboard.press('Backspace');
  // The same for a plain empty paragraph under the formatting marks.
  await page.getByTestId('perspective-developer').click();
  await page.evaluate(async () => {
    const v = (window as unknown as { docx4jEditor: { runAsync(name: string, args: unknown): Promise<unknown> } }).docx4jEditor;
    await v.runAsync('document.setChangeTrackingMode', { value: 'Off' });
  });
  await page.keyboard.press('Enter');
  const plain = '.docx-editor.body .ProseMirror > p:not(.rev-mark-ins):not(.rev-mark-del):has(> br.ProseMirror-trailingBreak:only-child)';
  await expect(page.locator(plain).first()).toBeAttached();
  await expect.poll(() => pseudo(page, plain, '::after')).toBe('"¶"');
  await expect.poll(() => page.evaluate((s) => getComputedStyle(document.querySelector(s)!.querySelector('br')!).display, plain)).toBe('none');
  // As accepted: the tracked pilcrow goes, the plain one stands in for it, and the break stays hidden only there.
  await page.getByTestId('perspective-edit').click();
  await page.locator('summary[aria-label="Options"]').click();
  await page.getByTestId('option-marks').uncheck();
  await page.keyboard.press('Escape');
  await expect.poll(() => page.evaluate((s) => getComputedStyle(document.querySelector(s)!.querySelector('br')!).display, plain)).not.toBe('none');
  await expect.poll(() => page.evaluate((s) => getComputedStyle(document.querySelector(s)!.querySelector('br')!).display, empty)).toBe('none');
});

// ED-005 section 12.80 item 8: Down and Up land in an empty paragraph, before its pilcrow, with the marks shown.
test('Down and Up stop in an empty paragraph with the marks shown', async ({ page }) => {
  await openFixture(page, 'tables.docx');
  await page.getByTestId('perspective-developer').click();
  const block = (): Promise<number> => page.evaluate(() => { const v = (window as unknown as { docx4jEditor: { state: { selection: { $from: { index(d: number): number; parentOffset: number } } } } }).docx4jEditor; return v.state.selection.$from.index(0); });
  // Two empty paragraphs after the first (Enter at its end, twice), the caret back at the first's end.
  await page.evaluate(() => { const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { nodeSize: number } } }; caretAt(a: number): void } }).docx4jEditor; v.caretAt(v.state.doc.child(0).nodeSize - 1); });
  await page.keyboard.press('Enter');
  await page.keyboard.press('Enter');
  await page.evaluate(() => { const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { nodeSize: number } } }; caretAt(a: number): void } }).docx4jEditor; v.caretAt(v.state.doc.child(0).nodeSize - 1); });
  await expect.poll(block).toBe(0);
  await page.keyboard.press('ArrowDown');
  await expect.poll(block).toBe(1);
  await page.keyboard.press('ArrowDown');
  await expect.poll(block).toBe(2);
  await page.keyboard.press('ArrowUp');
  await expect.poll(block).toBe(1);
  await page.keyboard.press('ArrowUp');
  await expect.poll(block).toBe(0);
});

test("an optional break's mark (ED-007 section 8.1 item 13): drawn with the marks, a click selects the character, Delete removes it", async ({ page }) => {
  await openFixture(page, 'tables/table-fit-fixed.docx');
  await page.getByTestId('perspective-developer').click();
  // The heading of table E, "E. Ten narrow columns...": an optional break after "E. ".
  await caretInBlock(page, 12, 3);
  await page.evaluate(() => (window as unknown as { docx4jEditor: { run(name: string): boolean } }).docx4jEditor.run('range.insertOptionalBreak'));
  const mark = page.locator('.docx-editor .optional-break');
  await expect(mark).toHaveCount(1);
  expect(await mark.evaluate((el) => getComputedStyle(el, '::after').content)).toBe('"▯"');
  // The block's text from the state: the body's DOM holds widgets among the blocks.
  const headingText = () => page.evaluate(() => (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc.child(12).textContent);
  expect(await headingText()).toContain('E. \u200BTen');
  // A click on the mark selects the character, drawn highlighted; Delete takes it out, and the mark with it.
  await mark.click();
  await expect(page.locator('.docx-editor .optional-break.selected')).toHaveCount(1);
  expect(await page.evaluate(() => { const s = (window as unknown as { docx4jEditor: { state: { selection: { from: number; to: number } } } }).docx4jEditor.state.selection; return s.to - s.from; })).toBe(1);
  await page.keyboard.press('Delete');
  await expect(mark).toHaveCount(0);
  expect(await headingText()).not.toContain('\u200B');
  expect(await headingText()).toContain('E. Ten');
  // Without the marks shown the character draws nothing.
  await page.getByTestId('perspective-edit').click();
  await caretInBlock(page, 12, 3);
  await page.evaluate(() => (window as unknown as { docx4jEditor: { run(name: string): boolean } }).docx4jEditor.run('range.insertOptionalBreak'));
  await expect(mark).toHaveCount(1);
  expect(await mark.evaluate((el) => getComputedStyle(el, '::after').content)).toBe('none');
});

test('a paragraph ending in invisible atoms (a bookmark\'s end) draws its pilcrow on its last visible child, not on the clipped atom', async ({ page }) => {
  // numbering-clauses.docx: the clause headings end with a bookmark end and empty runs, drawn as nothing.
  await openFixture(page, 'numbering-clauses.docx');
  await page.getByTestId('perspective-developer').click();
  const p = '.docx-editor.body .ProseMirror > p:nth-of-type(2)';
  await expect.poll(() => page.evaluate((sel) => { const el = document.querySelector(sel)!; return { br: !!el.querySelector(':scope > br.ProseMirror-trailingBreak:last-child'), lastInvisible: el.lastElementChild?.previousElementSibling?.previousElementSibling?.classList.contains('invisible') }; }, p)).toEqual({ br: true, lastInvisible: true });
  // The paragraph's own pilcrow is off; the last visible child carries it; the clipped atoms carry none.
  await expect.poll(() => pseudo(page, p, '::after')).toBe('none');
  await expect.poll(() => pseudo(page, `${p} > :nth-last-child(1 of :not(.invisible, br.ProseMirror-trailingBreak, img.ProseMirror-separator))`, '::after')).toBe('"¶"');
  await expect.poll(() => pseudo(page, `${p} > .invisible:last-of-type`, '::after')).toBe('none');
  // And it is drawn where it can be seen: the carrier is not clipped.
  const carrier = await page.evaluate((sel) => { const el = document.querySelector(`${sel} > :nth-last-child(1 of :not(.invisible, br.ProseMirror-trailingBreak, img.ProseMirror-separator))`)!; return { invisible: el.classList.contains('invisible'), width: el.getBoundingClientRect().width }; }, p);
  expect(carrier.invisible).toBe(false);
  expect(carrier.width).toBeGreaterThan(0);
});

test('the selected marks\' highlight goes with the editor\'s focus: after Ctrl+A, a click beside the page leaves no pilcrow highlighted', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await page.getByTestId('perspective-developer').click();
  const first = '.docx-editor.body > .ProseMirror > p';
  await expect.poll(() => pseudo(page, first, '::after')).toBe('"¶"');
  await caretInBlock(page, 4);
  await page.keyboard.press('Control+a');
  await expect(page.locator(first).first()).toHaveClass(/mark-selected/);
  const highlighted = await pseudo(page, first, '::after', 'background-color');
  expect(highlighted).not.toBe('rgba(0, 0, 0, 0)');
  // A click in the page area beside the page (Jason, 2026-11-02): the browser stops painting the text's selection,
  // so the marks' highlight must go too, though the editor's selection is still the whole document.
  const area = (await page.locator('.page-area').boundingBox())!;
  await page.mouse.click(area.x + area.width - 6, area.y + area.height / 2);
  await expect(page.locator('.docx-editor.body > .ProseMirror')).not.toHaveClass(/ProseMirror-focused/);
  await expect.poll(() => pseudo(page, first, '::after', 'background-color')).toBe('rgba(0, 0, 0, 0)');
  // Focus back in the editor, the selection as it was: highlighted again.
  await page.evaluate(() => (window as unknown as { docx4jEditor: { focus(): void } }).docx4jEditor.focus());
  await expect.poll(() => pseudo(page, first, '::after', 'background-color')).toBe(highlighted);
});
