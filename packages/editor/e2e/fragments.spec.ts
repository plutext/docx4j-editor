// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 proposal 46 (Word fragments), the editor's parts that need nothing new from the engine, over check 46c's template
// (fixtures/fragments/): the Data tab says what a Flat OPC node holds, read-only, and binds it as Word's mapped rich text
// around an empty paragraph (decision (e)); the Control tab's facts for both forms, and its switch to OpenDoPE's form.
import { test, expect } from '@playwright/test';
import { openFixture, download, comparePackages, fixture } from './helpers.js';

const body = '.docx-editor .ProseMirror[aria-label="Document"]';
const FRAGMENT = '/data[1]/fragment[1]';

async function template(page: import('@playwright/test').Page): Promise<void> {
  await openFixture(page, 'fragments/46c-template.docx');
  await page.getByTestId('perspective-template').click();
  await page.getByTestId('panel-control').waitFor();
}

/** The caret inside the control whose title is `label`. */
async function caretInControl(page: import('@playwright/test').Page, label: string): Promise<void> {
  await page.evaluate((label) => {
    const v = (window as unknown as { docx4jEditor: { view: { posAtDOM(n: Node, o: number): number }; caretAt(a: number): void } }).docx4jEditor;
    const el = document.querySelector(`.docx-editor.body [data-label="${label}"]`);
    if (!el) throw new Error(`no control ${label}`);
    const text = document.createTreeWalker(el, NodeFilter.SHOW_TEXT).nextNode();
    v.caretAt(v.view.posAtDOM(text ?? el, 0) + (text ? 1 : 0));
  }, label);
}

/** The caret at the end of the paragraph whose text is `text`. */
async function caretAtEndOf(page: import('@playwright/test').Page, text: string): Promise<void> {
  await page.evaluate((text) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { descendants(f: (n: { type: { name: string }; textContent: string; content: { size: number } }, pos: number) => boolean): void } }; caretAt(a: number): void } }).docx4jEditor;
    let at = -1;
    v.state.doc.descendants((n, pos) => { if (at < 0 && n.type.name === 'paragraph' && n.textContent === text) at = pos + 1 + n.content.size; return at < 0; });
    v.caretAt(at);
  }, text);
}

test('the Data tab says what a Flat OPC node holds, read-only, and binds it as Word\'s mapped rich text around an empty paragraph', async ({ page }) => {
  await template(page);
  await page.getByTestId('tab-data').click();
  const row = page.locator(`[data-xpath="${FRAGMENT}"] > .tree-row`);
  await expect(row.getByTestId('node-fragment')).toHaveText(/^Word document, \d+ KB: paragraphs 7, tables 1$/);
  // Read-only: no value box to type into, where the raw XML had been.
  await expect(row.locator('.node-value')).toHaveCount(0);
  // In a paragraph with text the bind is refused, saying why.
  await caretAtEndOf(page, 'After the control.');
  await row.hover();
  const bind = page.getByRole('button', { name: `Bind to ${FRAGMENT}: the control at the caret, or a new rich text control mapped to it` });
  await bind.click();
  await expect(page.getByTestId('status')).toHaveText('A Word document needs a paragraph of its own: put the caret in an empty paragraph');
  // In an empty paragraph: a block-level rich text control mapped by w15:dataBinding around it.
  await page.evaluate(() => (window as unknown as { docx4jEditor: { run(name: string): boolean } }).docx4jEditor.run('insertParagraph'));
  const before = await page.locator(`${body} > .sdt-block`).count();
  await row.hover();
  await bind.click();
  await expect(page.getByTestId('status')).toHaveText(`Inserted a rich text control mapped to ${FRAGMENT}: Word fills it with the Word document when it opens the file`);
  await expect(page.locator(`${body} > .sdt-block`)).toHaveCount(before + 1);
  // Filled at once with its Word document (part 3), as Word fills it on open.
  await expect(page.locator(`${body} > .sdt-block`).last()).toContainText('Fragment heading');
  // The control at the caret is Word's form: the Control tab's facts say so.
  await page.getByRole('tab', { name: 'Control', exact: true }).click();
  await expect(page.getByTestId('cc-fragment-kind')).toHaveText("Word's mapped rich text (w15:dataBinding): Word fills it from the data when it opens the file");
});

test('the Control tab: a Word fragment\'s facts in both forms, and the switch from Word\'s mapping to OpenDoPE\'s od:progid; untouched, the file saves as it came', async ({ page }) => {
  await template(page);
  await caretInControl(page, 'Rich text mapped');
  await page.getByRole('tab', { name: 'Control', exact: true }).click();
  await expect(page.getByTestId('control-fragment')).toBeVisible();
  await expect(page.getByTestId('cc-fragment-kind')).toHaveText("Word's mapped rich text (w15:dataBinding): Word fills it from the data when it opens the file");
  await expect(page.getByTestId('cc-fragment-data')).toHaveText(/^Word document, \d+ KB: paragraphs 7, tables 1$/);
  await expect(page.getByTestId('cc-fragment-kept')).toHaveText('pictures 1, notes 1, controls 1');
  await expect(page.getByTestId('cc-fragment-dropped')).toHaveCount(0);
  // A Word document is no XHTML: that switch is not offered for it.
  await expect(page.getByTestId('cc-xhtml')).toHaveCount(0);
  // OpenDoPE's own control: the same facts, its kind the processor's.
  await caretInControl(page, 'Word fragment');
  await expect(page.getByTestId('cc-fragment-kind')).toHaveText("OpenDoPE's Word fragment (od:progid=Word.Document), bound by its od:xpath entry: a processor resolves it");
  await expect(page.getByTestId('cc-fragment')).toBeChecked();
  // The plain text control mapped to the same node is no fragment.
  await caretInControl(page, 'Plain text mapped');
  await expect(page.getByTestId('control-fragment')).toHaveCount(0);
  // Untouched by any of this, the file saves as it came.
  const saved = await download(page);
  expect(comparePackages(await fixture('fragments/46c-template.docx'), saved)).toEqual([]);
  // The switch: Word's mapping becomes OpenDoPE's form, through the XPaths entry.
  await caretInControl(page, 'Rich text mapped');
  await expect(page.getByTestId('cc-fragment')).not.toBeChecked();
  await page.getByTestId('cc-fragment').check();
  await expect(page.getByTestId('cc-tag')).toHaveValue(/od:progid=Word\.Document/);
  await expect(page.getByTestId('cc-fragment-kind')).toHaveText("OpenDoPE's Word fragment (od:progid=Word.Document), bound by its od:xpath entry: a processor resolves it");
  await page.getByTestId('cc-fragment').uncheck();
  await expect(page.getByTestId('cc-fragment-kind')).toHaveText("Word's mapped rich text (w15:dataBinding): Word fills it from the data when it opens the file");
});

// Part 3 (section 12.119): the engine's pass fills a fragment control, out of the undo history, its styles and lists drawn.
test('the refresh: the stale banner offers Word\'s w15 control at open and fills it as Word does; the od:progid control is left to the processor; the Preview shows both filled', async ({ page }) => {
  await openFixture(page, 'fragments/46c-template.docx');
  await expect(page.getByTestId('stale-offer')).toContainText('One bound control shows content older than its data.');
  const mapped = page.locator(`${body} [data-label="Rich text mapped"]`);
  const progid = page.locator(`${body} [data-label="Word fragment"]`);
  await expect(mapped).toContainText('(3) rich text control placeholder');
  await page.getByTestId('stale-refresh').click();
  await expect(mapped).toContainText('Fragment heading (the fragment\'s Heading 1: red, 20 pt)');
  await expect(mapped).not.toContainText('(3) rich text control placeholder');
  await expect(progid).toContainText('(1) Word fragment placeholder');
  // As Word reconciles it: the document's Heading 1 (blue, 14 pt), the fragment's style (green italic), its own list (I., II.).
  const heading = mapped.locator('p.s-Heading1');
  expect(await heading.evaluate((el) => getComputedStyle(el).color)).toBe('rgb(31, 78, 121)');
  const note = mapped.locator('p.s-FragmentNote');
  expect(await note.evaluate((el) => [getComputedStyle(el).fontStyle, getComputedStyle(el).color])).toEqual(['italic', 'rgb(0, 176, 80)']);
  await expect(mapped.locator('p', { hasText: /Fragment list, item (one|two)/ }).locator('.list-label')).toHaveText(['I.', 'II.']);
  await expect(page.locator(`${body} > p`, { hasText: 'Document list, item three' }).locator('.list-label')).toHaveText('3.');
  await expect(mapped.locator('img').first()).toBeAttached();
  // Out of the undo history: Undo does not take the fill back.
  await page.evaluate(() => (window as unknown as { docx4jEditor: { focus(): void } }).docx4jEditor.focus());
  await page.keyboard.press('Control+z');
  await expect(mapped).toContainText('Fragment heading');
  // The Preview: the engine's bind step fills both forms.
  await page.getByTestId('perspective-template').click();
  await page.getByTestId('tab-preview').click();
  const story = page.getByTestId('preview-story').locator('.ProseMirror');
  await expect(story.locator('[data-label="Word fragment"]')).toContainText('Fragment heading');
  await expect(story.locator('[data-label="Rich text mapped"]')).toContainText('Fragment heading');
});

test('the Control tab\'s switch fills the control at once, in OpenDoPE\'s form too', async ({ page }) => {
  await template(page);
  await page.getByTestId('stale-keep').click();
  await caretInControl(page, 'Word fragment');
  await page.getByRole('tab', { name: 'Control', exact: true }).click();
  const progid = page.locator(`${body} [data-label="Word fragment"]`);
  await expect(progid).toContainText('(1) Word fragment placeholder');
  // Off and on again: Word's mapping, then OpenDoPE's form, each filled from the node.
  await page.getByTestId('cc-fragment').uncheck();
  await expect(progid).toContainText('Fragment heading');
  await page.getByTestId('cc-fragment').check();
  await expect(page.getByTestId('cc-tag')).toHaveValue(/od:progid=Word\.Document/);
  await expect(progid).toContainText('Fragment heading');
  await expect(progid).not.toContainText('(1) Word fragment placeholder');
});

// Section 12.124 (Jason, 2026-11-12): a bind of a control already there shows the Word document, as a bind of XHTML shows its
// markup; that control alone is filled, and Word's w15 control left at open stays as the stale banner's Keep left it.
test('the Data tab\'s bind of a plain text control alone in its paragraph: Word\'s mapped rich text, filled at once, and no other control', async ({ page }) => {
  await template(page);
  await page.getByTestId('stale-keep').click();
  const mapped = page.locator(`${body} [data-label="Rich text mapped"]`);
  await expect(mapped).toContainText('(3) rich text control placeholder');
  // A plain text control in an empty paragraph of its own, the caret inside it.
  await caretAtEndOf(page, 'After the control.');
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { run(name: string, args?: unknown): boolean } }).docx4jEditor;
    v.run('insertParagraph');
    v.run('contentControl.insert', { kind: 'PlainText', title: 'Was plain' });
  });
  await expect(page.locator(`${body} [data-label="Was plain"]`)).toHaveCount(1);
  await page.getByTestId('tab-data').click();
  const row = page.locator(`[data-xpath="${FRAGMENT}"] > .tree-row`);
  await row.hover();
  await page.getByRole('button', { name: `Bind to ${FRAGMENT}: the control at the caret, or a new rich text control mapped to it` }).click();
  await expect(page.getByTestId('status')).toContainText(`Bound to ${FRAGMENT}`);
  const control = page.locator(`${body} > .sdt-block[data-label="Was plain"]`);
  await expect(control).toContainText('Fragment heading');
  await expect(mapped).toContainText('(3) rich text control placeholder');
  await page.getByRole('tab', { name: 'Control', exact: true }).click();
  await expect(page.getByTestId('cc-fragment-kind')).toHaveText("Word's mapped rich text (w15:dataBinding): Word fills it from the data when it opens the file");
});

test('a Flat OPC node dragged onto such a control is bound the same way and filled at once', async ({ page }) => {
  await template(page);
  await page.getByTestId('stale-keep').click();
  await caretAtEndOf(page, 'After the control.');
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { run(name: string, args?: unknown): boolean } }).docx4jEditor;
    v.run('insertParagraph');
    v.run('contentControl.insert', { kind: 'PlainText', title: 'Was plain' });
  });
  await page.getByTestId('tab-data').click();
  await page.locator(`[data-xpath="${FRAGMENT}"] .node-name`).first().dragTo(page.locator(`${body} [data-label="Was plain"]`));
  await expect(page.getByTestId('status')).toContainText(`Bound to ${FRAGMENT}`);
  await expect(page.locator(`${body} > .sdt-block[data-label="Was plain"]`)).toContainText('Fragment heading');
  await expect(page.locator(`${body} [data-label="Rich text mapped"]`)).toContainText('(3) rich text control placeholder');
});
