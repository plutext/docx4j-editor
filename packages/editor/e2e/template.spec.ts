// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 4.10 item 5, the first half (step B3): invoice.docx suggests the Template
// perspective; its tabs; the controls drawn with their tags; the repeat row selected shows the
// repeat and its count against the data; bind by drag from a data node; the round trip of an
// untouched template holds after every panel was opened.
import { test, expect } from '@playwright/test';
import { openFixture, download, fixture, comparePackages, selectBlock, selectBlocks, caretInBlock, xmlEditorValue, setXmlEditor } from './helpers.js';

const body = '.docx-editor .ProseMirror[aria-label="Document"]';

async function template(page: import('@playwright/test').Page): Promise<void> {
  await openFixture(page, 'invoice.docx');
  await page.getByTestId('offer-template').click();
  await expect(page.getByTestId('perspective-template')).toHaveAttribute('aria-selected', 'true');
  await page.getByTestId('panel-control').waitFor();
}

test('invoice.docx suggests Template; the perspective has its tabs, the control group and the tags; the untouched round trip holds after every tab', async ({ page }) => {
  // Wide enough for the page, the comments margin and the side panel: a narrower viewport shrinks the page, and a fixed-width table then scrolls inside its wrapper.
  await page.setViewportSize({ width: 1800, height: 900 });
  await openFixture(page, 'invoice.docx');
  const offer = page.getByTestId('perspective-offer');
  await expect(offer).toContainText('template');
  await page.getByTestId('offer-template').click();
  await expect(offer).toHaveCount(0);
  await expect(page.getByTestId('perspective-template')).toHaveAttribute('aria-selected', 'true');
  const tabs = page.getByRole('tablist', { name: 'Panel tabs' });
  await expect(tabs.getByRole('tab')).toHaveText(['Control', 'Controls', 'Data', 'OpenDoPE', 'Preview']);
  await expect(page.getByTestId('control-tools')).toBeVisible();
  await expect(page.getByTestId('font-name')).toHaveCount(0);   // the basic toolbar
  // The controls are drawn with their tags: the customer name's label in the margin, the repeat row's.
  const block = page.locator(`${body} .sdt-block[data-label="Customer name"]`);
  await expect(block).toHaveCount(1);
  expect(await block.evaluate((el) => getComputedStyle(el, '::before').content)).toContain('Customer name');
  await expect(page.locator(`${body} tr[data-sdt-role="repeat"]`)).toHaveCount(1);
  // The repeat row's label is a badge in its first cell (with the cell control's own), never a box that lays out as a cell:
  // the table keeps two cells per row and its wrapper does not scroll sideways.
  const firstCell = page.locator(`${body} tr[data-sdt-role="repeat"] > td`).first();
  expect(await firstCell.evaluate((el) => getComputedStyle(el, '::before').content)).toContain('Repeat');
  expect(await firstCell.evaluate((el) => getComputedStyle(el, '::before').content)).toContain('Description');
  expect(await page.locator(`${body} .tableWrapper`).first().evaluate((el) => el.scrollWidth <= el.clientWidth)).toBe(true);
  const cells = await page.locator(`${body} tr[data-sdt-role="repeat"] > td`).evaluateAll((tds) => tds.map((td) => (td as HTMLElement).getBoundingClientRect().left));
  const headerCells = await page.locator(`${body} tr[data-sdt-role="repeat"]`).evaluate((tr) => Array.from(tr.previousElementSibling!.children).map((td) => td.getBoundingClientRect().left));
  expect(cells.map(Math.round)).toEqual(headerCells.map(Math.round));   // the cells line up under the header's
  // Every tab opened and read.
  await page.getByTestId('tab-controls').click();
  await expect(page.getByTestId('controls-list').locator('li')).toHaveCount(6);
  await page.getByTestId('tab-data').click();
  await expect(page.getByTestId('data-part')).toContainText('rule 1');
  await expect(page.getByTestId('od-xpaths')).toContainText('x2');
  await page.getByTestId('xpath-expression').fill('/invoice/items/item');
  await page.getByTestId('xpath-evaluate').click();
  await expect(page.getByTestId('xpath-result')).toContainText('3 nodes');
  await page.getByTestId('tab-opendope').click();
  await expect(page.getByTestId('od-entry-x2')).toContainText('the row control');
  await expect(page.getByTestId('od-condition-c6-verdict')).toHaveText('false (dropped)');
  await page.getByTestId('tab-control').click();
  const saved = await download(page);
  expect(comparePackages(await fixture('invoice.docx'), saved)).toEqual([]);
  // The choice persists.
  await page.reload();
  await page.getByTestId('open').setInputFiles(new URL('../../editor-model/test/fixtures/invoice.docx', import.meta.url).pathname);
  await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await expect(page.getByTestId('perspective-template')).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByTestId('perspective-offer')).toHaveCount(0);
});

test('the repeat row selected from the tree shows the repeat and its count; a property edit through the panel is one undo step', async ({ page }) => {
  await template(page);
  await page.getByTestId('tab-controls').click();
  const rows = page.getByTestId('controls-list').locator('li');
  await expect(rows.nth(1)).toContainText('repeat x2');
  await expect(rows.nth(1)).toContainText('3 instances');
  await rows.nth(1).getByRole('button').click();
  await expect(rows.nth(1)).toHaveAttribute('aria-current', 'true');
  await page.getByTestId('tab-control').click();
  await expect(page.getByTestId('control-summary')).toContainText('row · Rich text');
  await expect(page.getByTestId('cc-role')).toHaveText('repeat x2');
  await expect(page.getByTestId('cc-repeat-count')).toContainText('3 against the data part');
  await page.getByTestId('cc-title').fill('Items');
  await page.getByTestId('cc-title').press('Enter');
  await expect(page.locator(`${body} tr[data-sdt="Items"]`)).toHaveCount(1);
  await expect(page.getByTestId('name')).toContainText('•');
  await page.locator(body).focus();
  await page.keyboard.press('Control+z');
  await expect(page.locator(`${body} tr[data-sdt="Items"]`)).toHaveCount(0);
  await expect(page.locator(`${body} tr[data-sdt="Repeat"]`)).toHaveCount(1);   // the alias the fixture gives the row
  await expect(page.getByTestId('cc-title')).toHaveValue('Repeat');
});

test('bind by drag from a data node onto a control; the download carries the binding and the XPaths entry', async ({ page }) => {
  await template(page);
  await page.getByTestId('tab-data').click();
  const source = page.locator('[data-xpath="/invoice[1]/items[1]/item[1]/price[1]"] .node-name');
  await expect(source).toBeVisible();
  const target = page.locator(`${body} .sdt-block[data-label="Customer name"]`);
  await source.dragTo(target);
  await expect(page.getByTestId('status')).toContainText('Bound to /invoice[1]/items[1]/item[1]/price[1]');
  await page.getByTestId('tab-control').click();
  await expect(page.getByTestId('control-summary')).toContainText('Plain text');
  await expect(page.getByTestId('cc-binding-selects')).toContainText('$20');
  await expect(page.getByTestId('cc-entry')).toContainText('x1');
  await expect(page.getByTestId('cc-verdict')).toContainText('$20');
  const saved = await download(page);
  expect(comparePackages(await fixture('invoice.docx'), saved, ['word/document.xml', 'docProps/app.xml', 'customXml/item4.xml'])).toEqual([]);
  const { EditorDocument } = await import('@docx4j/editor-model');
  const session = await EditorDocument.open(saved);
  await session.template.load({ engine: false });
  expect(session.template.xpaths!.entries[0]).toMatchObject({ id: 'x1', xpath: '/invoice[1]/items[1]/item[1]/price[1]' });
  expect(session.template.xpaths!.authoringTool).toBe('docx4j-editor');
});

// ED-005 section 12.54 (Jason's report): binding showed only in the Verdict. As Word's XML mapping shows the node's
// value in the control at once, the document now does: the Customer name control, bound to the first item's price.
test('bind from the Data tab replaces the control\'s text in the document with the bound value; one undo takes both back', async ({ page }) => {
  await template(page);
  await caretInBlock(page, 3, 2);   // the "Customer name" control, a plain text control showing "Joe Bloggs"
  const controlText = (): Promise<string> => page.evaluate(() => (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc.child(3).textContent);
  expect(await controlText()).toBe('Joe Bloggs');
  await page.getByRole('tab', { name: 'Data' }).click();
  // The node's bind action shows on its row, under the pointer, as for a person.
  await page.locator('[data-xpath="/invoice[1]/items[1]/item[1]/price[1]"] > .tree-row').hover();
  await page.getByRole('button', { name: 'Bind to /invoice[1]/items[1]/item[1]/price[1]: the control at the caret, or a new plain text control' }).click();
  await expect(page.getByTestId('status')).toHaveText('Bound to /invoice[1]/items[1]/item[1]/price[1]');
  await expect.poll(controlText).toBe('$20');
  await expect(page.locator(body).locator('.sdt-block').first()).toContainText('$20');
  await expect(page.locator(body)).not.toContainText('Joe Bloggs');
  // The caret is still in the control: the Control tab shows it, bound to the price.
  await page.getByRole('tab', { name: 'Control', exact: true }).click();
  await expect(page.getByTestId('cc-binding-selects')).toContainText('$20');
  await page.evaluate(() => (window as unknown as { docx4jEditor: { run(name: string): boolean } }).docx4jEditor.run('history.undo'));
  await expect.poll(controlText).toBe('Joe Bloggs');
});

// ED-005 section 12.55: Word's mapping both ways. Typing in the bound Customer name control writes its data node, which the
// Data tab shows; a value edited in the Data tab shows in the control.
test('two-way binding: typing in a bound control writes its node, and a value edited in the Data tab shows in the control', async ({ page }) => {
  await template(page);
  const controlText = (): Promise<string> => page.evaluate(() => (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc.child(3).textContent);
  expect(await controlText()).toBe('Joe Bloggs');
  // The caret after "Joe Bloggs", inside the control's paragraph, then typing.
  await caretInBlock(page, 3, 11);
  await page.keyboard.type(' Jr');
  await expect.poll(controlText).toBe('Joe Bloggs Jr');
  await page.getByTestId('tab-data').click();
  const value = page.locator('[data-xpath="/invoice[1]/customer[1]/name[1]"] .node-value');
  await expect(value).toContainText('Joe Bloggs Jr');
  // The other way: the value edited in the Data tab, and the control follows.
  await value.click();
  const box = page.locator('[data-xpath="/invoice[1]/customer[1]/name[1]"] input');
  await box.fill('Jane Doe');
  await box.press('Enter');
  await expect.poll(controlText).toBe('Jane Doe');
  await expect(page.locator(body).locator('.sdt-block').first()).toHaveText('Jane Doe');
});

// ED-005 section 12.56: with no control at the caret, the Data tab's bind makes a plain text control around the
// selection (or at the caret) and binds it in one step; a selection across paragraphs is refused, saying why.
test('bind with no control at the caret: a plain text control made around the selection and bound in one step; across paragraphs, refused with the reason', async ({ page }) => {
  await template(page);
  const view = (): Promise<{ block: number; from: number; text: string }> => page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { textContent: string; nodeSize: number } } } } }).docx4jEditor;
    let pos = 0;
    for (let i = 0; i < v.state.doc.childCount; i++) {
      const b = v.state.doc.child(i);
      if (b.textContent === 'INVOICE') return { block: i, from: pos, text: b.textContent };
      pos += b.nodeSize;
    }
    return { block: -1, from: -1, text: '' };
  });
  const before = await view();
  expect(before.block).toBeGreaterThan(0);
  // Select "INVOICE", a paragraph with no control around it.
  await page.evaluate(([from, to]) => (window as unknown as { docx4jEditor: { caretAt(a: number, b: number): void } }).docx4jEditor.caretAt(from, to), [before.from + 1, before.from + 1 + 7] as const);
  await page.getByTestId('tab-data').click();
  const price = '/invoice[1]/items[1]/item[1]/price[1]';
  await page.locator(`[data-xpath="${price}"] > .tree-row`).hover();
  await page.getByRole('button', { name: `Bind to ${price}: the control at the caret, or a new plain text control` }).click();
  await expect(page.getByTestId('status')).toHaveText(`Inserted a plain text control bound to ${price}`);
  const after = { block: before.block, text: await page.evaluate((i) => (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc.child(i).textContent, before.block) };
  expect(after.text).toBe('$20');
  // The new control is the caret's: the Control tab shows it, bound to the price.
  await page.getByRole('tab', { name: 'Control', exact: true }).click();
  await expect(page.getByTestId('cc-binding-selects')).toContainText('$20');
  // A selection across two paragraphs cannot become a plain text control: refused, saying why.
  await selectBlocks(page, after.block - 1, after.block);
  await page.getByTestId('tab-data').click();
  await page.locator(`[data-xpath="${price}"] > .tree-row`).hover();
  await page.getByRole('button', { name: `Bind to ${price}: the control at the caret, or a new plain text control` }).click();
  await expect(page.getByTestId('status')).toHaveText('A plain text control holds text within one paragraph: select less, or put the caret where the control should go');
});

// ED-005 section 12.75: the Control tab's "XHTML content" switch binds the control through an od:xpath entry
// and converts the node's markup into it; a Data tab edit of the node converts again; nothing is written back.
test('XHTML content: the switch converts the bound node\'s markup into the control, a Data tab edit follows, the node is not written back', async ({ page }) => {
  await template(page);
  const view = (): Promise<{ block: number; from: number }> => page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { textContent: string; nodeSize: number } } } } }).docx4jEditor;
    let pos = 0;
    for (let i = 0; i < v.state.doc.childCount; i++) { const b = v.state.doc.child(i); if (b.textContent === 'INVOICE') return { block: i, from: pos }; pos += b.nodeSize; }
    return { block: -1, from: -1 };
  });
  const at = await view();
  // A plain text control over "INVOICE", bound to the customer's name through the Data tab.
  await page.evaluate(([from, to]) => (window as unknown as { docx4jEditor: { caretAt(a: number, b: number): void } }).docx4jEditor.caretAt(from, to), [at.from + 1, at.from + 1 + 7] as const);
  await page.getByTestId('tab-data').click();
  const name = '/invoice[1]/customer[1]/name[1]';
  await page.locator(`[data-xpath="${name}"] > .tree-row`).hover();
  await page.getByRole('button', { name: `Bind to ${name}: the control at the caret, or a new plain text control` }).click();
  await expect(page.getByTestId('status')).toHaveText(`Inserted a plain text control bound to ${name}`);
  const blockText = (): Promise<string> => page.evaluate((i) => (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc.child(i).textContent, at.block);
  expect(await blockText()).toBe('Joe Bloggs');
  // The switch: the markup the node now holds is converted into the control.
  await page.getByRole('tab', { name: 'Control', exact: true }).click();
  await page.getByTestId('cc-xhtml').check();
  await expect(page.getByTestId('cc-xhtml')).toBeChecked();
  await expect(page.getByTestId('cc-entry')).toBeVisible();
  // The node's value edited in the Data tab to markup: the control shows it converted (a run-level control: the first paragraph's inlines).
  await page.getByTestId('tab-data').click();
  await page.locator(`[data-xpath="${name}"] .node-value`).click();
  const box = page.locator(`[data-xpath="${name}"] input`);
  await box.fill('<p>Jane <b>Doe</b></p><p>second</p>');
  await box.press('Enter');
  await expect.poll(blockText).toBe('Jane Doe');
  const boldText = await page.evaluate((i) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { descendants(f: (n: { isText: boolean; text?: string; marks: { type: { name: string } }[] }) => void): void } } } } }).docx4jEditor;
    let out = '';
    v.state.doc.child(i).descendants((n) => { if (n.isText && n.marks.some((m) => m.type.name === 'bold')) out += n.text; });
    return out;
  }, at.block);
  expect(boldText).toBe('Doe');
  // Typing inside writes nothing back: the node keeps its markup.
  await page.evaluate(([from]) => (window as unknown as { docx4jEditor: { caretAt(a: number): void } }).docx4jEditor.caretAt(from + 2), [at.from] as const);
  await page.keyboard.type('X');
  await page.waitForTimeout(200);
  await expect(page.locator(`[data-xpath="${name}"] .node-value`)).toContainText('<p>Jane <b>Doe</b></p>');
});

// ED-005 section 12.86: the Control tab's bind form on a node holding markup converts it into the control at once.
test('the Control tab\'s bind of a node holding XHTML converts the markup into the control', async ({ page }) => {
  await template(page);
  const view = (): Promise<{ block: number; from: number }> => page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { textContent: string; nodeSize: number } } } } }).docx4jEditor;
    let pos = 0;
    for (let i = 0; i < v.state.doc.childCount; i++) { const b = v.state.doc.child(i); if (b.textContent === 'INVOICE') return { block: i, from: pos }; pos += b.nodeSize; }
    return { block: -1, from: -1 };
  });
  const at = await view();
  await page.getByTestId('tab-data').click();
  const name = '/invoice[1]/customer[1]/name[1]';
  await page.locator(`[data-xpath="${name}"] .node-value`).click();
  const box = page.locator(`[data-xpath="${name}"] input`);
  await box.fill('<p>Jane <b>Doe</b></p>');
  await box.press('Enter');
  await expect(page.locator(`[data-xpath="${name}"] .node-value`)).toContainText('<b>Doe</b>');
  // A rich text control over INVOICE, then the Control tab's form binds it to the node.
  await page.evaluate(([from, to]) => { const v = (window as unknown as { docx4jEditor: { caretAt(a: number, b: number): void; run(name: string, args: unknown): boolean } }).docx4jEditor; v.caretAt(from, to); v.run('contentControl.insert', { kind: 'RichText' }); }, [at.from + 1, at.from + 1 + 7] as const);
  await page.getByRole('tab', { name: 'Control', exact: true }).click();
  await page.getByTestId('bind-xpath').fill(name);
  await page.getByTestId('bind-apply').click();
  await expect(page.getByTestId('cc-xhtml')).toBeChecked();
  const blockText = (): Promise<string> => page.evaluate((i) => (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc.child(i).textContent, at.block);
  await expect.poll(blockText).toBe('Jane Doe');
});

// ED-005 section 12.79: the Data tab's bind of a node holding XHTML makes a rich text control bound through the
// entry, block-level markup around a blank paragraph (refused inside one with content), inline markup at the caret.
test('Data tab bind of an XHTML node: refused inside a paragraph with content, a block rich text control around a blank one, converted', async ({ page }) => {
  await template(page);
  const view = (): Promise<{ block: number; from: number }> => page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { textContent: string; nodeSize: number } } } } }).docx4jEditor;
    let pos = 0;
    for (let i = 0; i < v.state.doc.childCount; i++) { const b = v.state.doc.child(i); if (b.textContent === 'INVOICE') return { block: i, from: pos }; pos += b.nodeSize; }
    return { block: -1, from: -1 };
  });
  const at = await view();
  // The customer's name set to block-level markup in the Data tab.
  await page.getByTestId('tab-data').click();
  const name = '/invoice[1]/customer[1]/name[1]';
  await page.locator(`[data-xpath="${name}"] .node-value`).click();
  const box = page.locator(`[data-xpath="${name}"] input`);
  await box.fill('<p>  Sample <ul><li>Sample List 1</li><li>Sample List 2</li></ul></p>');
  await box.press('Enter');
  await expect(page.locator(`[data-xpath="${name}"] .node-value`)).toContainText('<ul>');
  const bind = async (): Promise<void> => {
    await page.locator(`[data-xpath="${name}"] > .tree-row`).hover();
    await page.getByRole('button', { name: `Bind to ${name}: the control at the caret, or a new plain text control` }).click();
  };
  // Inside "INVOICE": refused, nothing made.
  await page.evaluate(([from]) => (window as unknown as { docx4jEditor: { caretAt(a: number): void } }).docx4jEditor.caretAt(from), [at.from + 3] as const);
  await bind();
  await expect(page.getByTestId('status')).toHaveText("Can't add block level XHTML inside a paragraph.");
  const blockName = (i: number): Promise<string> => page.evaluate((i) => (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { type: { name: string } } } } } }).docx4jEditor.state.doc.child(i).type.name, i);
  expect(await blockName(at.block)).toBe('paragraph');
  // Enter at the end of INVOICE makes a blank paragraph; the bind there wraps it in a rich text control, converted.
  await page.evaluate(([from]) => (window as unknown as { docx4jEditor: { caretAt(a: number): void } }).docx4jEditor.caretAt(from), [at.from + 1 + 7] as const);
  await page.keyboard.press('Enter');
  await expect.poll(() => blockName(at.block + 1)).toBe('paragraph');
  await bind();
  await expect(page.getByTestId('status')).toHaveText(`Inserted a rich text control bound to ${name}, its XHTML converted`);
  await expect.poll(() => blockName(at.block + 1)).toBe('sdtBlock');
  const inside = (): Promise<string[]> => page.evaluate((i) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { childCount: number; child(j: number): { textContent: string } } } } } }).docx4jEditor;
    const b = v.state.doc.child(i); const out: string[] = [];
    for (let j = 0; j < b.childCount; j++) out.push(b.child(j).textContent);
    return out;
  }, at.block + 1);
  // The markup's trailing </p> after the list (invalid HTML) parses as an empty paragraph, as a browser reads it.
  await expect.poll(inside).toEqual(['Sample', 'Sample List 1', 'Sample List 2', '']);
  expect(await blockName(at.block)).toBe('paragraph');
});

test('the OpenDoPE tab: an entry and a condition made in the editors, the condition put on a control from the Control tab, the download carries the parts', async ({ page }) => {
  await template(page);
  await page.getByTestId('tab-opendope').click();
  // A new entry from an expression, checked by evaluation.
  await page.getByTestId('od-entry-new').click();
  await page.getByTestId('od-entry-xpath').fill('/invoice/customer/name');
  await page.getByTestId('od-entry-name').fill('Customer');
  await page.getByTestId('od-entry-apply').click();
  await expect(page.getByTestId('od-entry-x7')).toContainText('Joe Bloggs');
  // The boolean column is empty for an entry used as a repeat or a binding, and shown for one that might be a condition's (ED-005 section 12.84).
  await expect(page.locator('[data-testid="od-xpaths"] tr.use-repeat [data-testid$="-boolean"]').first()).toHaveText('');
  await expect(page.locator('[data-testid="od-xpaths"] tr.use-bound [data-testid$="-boolean"]').first()).toHaveText('');
  await expect(page.getByTestId('od-entry-x7-boolean')).not.toHaveText('');
  // A condition: and(x5, not(x6)) built in the tree (x5 is a boolean in the data; x7's string would be false in the java mode).
  await page.getByTestId('od-condition-new').click();
  await page.getByTestId('od-condition-name').fill('named and not spam');
  await page.getByTestId('operand-root-op').selectOption('and');
  await page.getByTestId('operand-root-0-id').selectOption('x5');
  await page.getByTestId('operand-root-add').click();
  await page.getByTestId('operand-root-1-op').selectOption('not');
  await page.getByTestId('operand-root-1-0-id').selectOption('x6');
  await expect(page.getByTestId('od-condition-form')).toContainText('and(x5');
  await page.getByTestId('od-condition-apply').click();
  await expect(page.getByTestId('od-condition-c1-verdict')).toHaveText('true (kept)');
  // The paragraph "Please remit funds…" wrapped in a control, then given the condition in the Control tab.
  await selectBlock(page, 2);
  await page.getByTestId('insert-control').selectOption('RichText');
  await caretInBlock(page, 2, 1);
  await page.getByTestId('tab-control').click();
  await expect(page.getByTestId('control-summary')).toContainText('block · Rich text');
  await page.getByTestId('cc-role-kind').selectOption('condition');
  await page.getByTestId('cc-role-existing').selectOption('c1');
  await page.getByTestId('cc-role-apply').click();
  await expect(page.getByTestId('cc-role')).toHaveText('condition c1');
  await expect(page.getByTestId('cc-role-verdict')).toContainText('kept');
  await expect(page.locator(`${body} .sdt-block[data-role="condition"][data-label="od:condition=c1"]`)).toHaveCount(1);
  const saved = await download(page);
  expect(comparePackages(await fixture('invoice.docx'), saved, ['word/document.xml', 'docProps/app.xml', 'customXml/item1.xml', 'customXml/item4.xml'])).toEqual([]);
  const { EditorDocument, listControls } = await import('@docx4j/editor-model');
  const session = await EditorDocument.open(saved);
  await session.template.load({ engine: false });
  expect(session.template.xpaths!.entries.find((e) => e.id === 'x7')).toMatchObject({ xpath: '/invoice/customer/name', attrs: { name: 'Customer' } });
  expect(session.template.conditions!.conditions.find((c) => c.id === 'c1')).toEqual({ id: 'c1', attrs: { name: 'named and not spam' }, operand: { op: 'and', operands: [{ op: 'xpathref', id: 'x5' }, { op: 'not', operand: { op: 'xpathref', id: 'x6' } }] } });
  expect(listControls(session.doc, session.controlEntry).map((c) => c.tag.roleId)).toContain('c1');
  expect(session.template.problems.filter((p) => p.requirement !== 'REQ-009')).toEqual([]);
});

test('an XPath 2 template: the xpath2 mode evaluates in the browser through the engine the Template chunk hands over (core-ts 0.1.4)', async ({ page }) => {
  await openFixture(page, 'opendope/invoice_Saxon_XPath2.docx');
  await page.getByTestId('offer-template').click();
  await page.getByTestId('tab-controls').click();
  const row = page.getByTestId('controls-list').locator('li', { hasText: 'condition dateGtCondition' });
  await expect(row).toContainText('kept');
  // The template declares no mode (2.3-era): the expression's XPath 2 syntax is what needs the engine, and it answers.
  await page.getByTestId('tab-opendope').click();
  await expect(page.getByTestId('od-version')).toContainText('none declared');
  await expect(page.getByTestId('od-condition-dateGtCondition-verdict')).toHaveText('true (kept)');
});

test('the Preview tab shows the customer name in place with the verdict badges; an edited data value shows on the next preview; the page carries the badges', async ({ page }) => {
  await template(page);
  // On the page: the dropped paragraph and the repeat row's count.
  await expect(page.locator(`${body} .sdt-block[data-verdict="dropped"]`)).toHaveCount(1);
  await expect(page.locator(`${body} td[data-sdt-verdict="3 instances"]`)).toHaveCount(1);
  await page.getByTestId('tab-preview').click();
  const story = page.getByTestId('preview-story').locator('.ProseMirror');
  await expect(story.locator('.sdt-block').first()).toHaveText('Joe Bloggs');
  await expect(story.locator('.sdt-block[data-verdict="dropped"]')).toHaveCount(1);
  await expect(story.locator('td[data-sdt-verdict="3 instances"]')).toHaveCount(1);
  await expect(page.getByTestId('preview-summary')).toContainText('3 bound controls');
  // The customer's name edited in the Data tab, then the preview again.
  await page.getByTestId('tab-data').click();
  await page.locator('[data-xpath="/invoice[1]/customer[1]/name[1]"] .node-value').click();
  const box = page.locator('[data-xpath="/invoice[1]/customer[1]/name[1]"] input');
  await box.fill('Jane Doe');
  await box.press('Enter');
  await page.getByTestId('tab-preview').click();
  await expect(story.locator('.sdt-block').first()).toHaveText('Jane Doe');
  // The document's own Customer name control follows the data, as Word's mapping does (ED-005 section 12.55): the data
  // part and the main part changed, nothing else.
  await expect(page.locator(body).locator('.sdt-block').first()).toHaveText('Jane Doe');
  const saved = await download(page);
  expect(comparePackages(await fixture('invoice.docx'), saved, ['customXml/item3.xml', 'word/document.xml', 'docProps/app.xml'])).toEqual([]);
});

test('the Data tab edits the data part: an item duplicated raises the repeat count on the page and in the tree; a node dragged before another; an attribute added and removed', async ({ page }) => {
  await template(page);
  await page.getByTestId('tab-data').click();
  const item = page.locator('[data-xpath="/invoice[1]/items[1]/item[1]"] > .tree-row');
  await item.locator('.node-name').click();
  await item.getByRole('button', { name: 'Duplicate /invoice[1]/items[1]/item[1]' }).click();
  await expect(page.locator('[data-xpath="/invoice[1]/items[1]/item[4]"]')).toHaveCount(1);
  await expect(page.locator(`${body} td[data-sdt-verdict="4 instances"]`)).toHaveCount(1);
  await page.getByTestId('tab-controls').click();
  await expect(page.getByTestId('controls-list').locator('li').nth(1)).toContainText('4 instances');
  await page.getByTestId('tab-data').click();
  // Drag to move: the third item dropped on the upper edge of the first goes before it (its name tells them apart).
  const third = page.locator('[data-xpath="/invoice[1]/items[1]/item[3]"] > .tree-row');
  const thirdName = (await page.locator('[data-xpath="/invoice[1]/items[1]/item[3]/name[1]"] .node-value').textContent())!;
  const first = page.locator('[data-xpath="/invoice[1]/items[1]/item[1]"] > .tree-row');
  await third.locator('.node-name').dragTo(first, { targetPosition: { x: 40, y: 2 } });
  await expect(page.getByTestId('status')).toContainText('moved before /invoice[1]/items[1]/item[1]');
  await expect(page.locator('[data-xpath="/invoice[1]/items[1]/item[1]/name[1]"] .node-value')).toHaveText(thirdName);
  const customer = page.locator('[data-xpath="/invoice[1]/customer[1]"] > .tree-row');
  await customer.locator('.node-name').click();
  await customer.getByRole('button', { name: 'Add an attribute to /invoice[1]/customer[1]' }).click();
  await page.getByTestId('node-form-name').fill('id');
  await page.getByTestId('node-form-value').fill('42');
  await page.getByTestId('node-form-apply').click();
  await expect(customer).toContainText('@id="42"');
  await customer.locator('.node-name').click();
  await customer.getByRole('button', { name: 'Remove @id from /invoice[1]/customer[1]' }).click();
  await expect(customer).not.toContainText('@id');
  const saved = await download(page);
  expect(comparePackages(await fixture('invoice.docx'), saved, ['customXml/item3.xml', 'docProps/app.xml'])).toEqual([]);
  const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(saved);
  await pkg.customXmlParts.load();
  const items = pkg.customXmlParts.getItem('{8B049945-9DFE-4726-9DE9-CF5691E53858}')!.selectNodes('/invoice/items/item');
  expect(items).toHaveLength(4);
  expect(items[0]!.selectSingleNode('name')!.text).toBe(thirdName);
});

test('a document with no data part gets one from the Data tab (section 11.2 item 31): the root named, a child added, the download carries the part', async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  await page.getByTestId('perspective-template').click();
  await page.getByTestId('tab-data').click();
  await expect(page.getByTestId('data-part-undetermined')).toContainText('no custom XML part could hold data');
  await page.getByTestId('new-data-root').fill('invoice');
  await page.getByTestId('new-data-namespace').fill('urn:acme:invoice');
  await page.getByTestId('new-data-create').click();
  await expect(page.getByTestId('status')).toContainText('<invoice> added');
  await expect(page.getByTestId('data-part-undetermined')).toHaveCount(0);
  // The root is in a namespace, so its path carries the prefix the part gives it; the tree is addressed by role.
  const tree = page.getByTestId('data-part');
  const root = tree.getByRole('treeitem').first().locator('> .tree-row');
  await expect(root).toContainText('invoice');
  await root.locator('.node-name').click();
  await root.getByRole('button', { name: /^Add a child to / }).click();
  await page.getByTestId('node-form-name').fill('customer');
  await page.getByTestId('node-form-value').fill('Ann');
  await page.getByTestId('node-form-apply').click();
  await expect(tree.locator('.node-value')).toHaveText('Ann');
  const saved = await download(page);
  const diffs = comparePackages(await fixture('HelloWordOnline.docx'), saved, ['docProps/app.xml']);
  // The part, its properties part, and the relationship between them (a .rels part, regenerated by design).
  expect(diffs.map((d) => `${d.reason} ${d.name}`).sort()).toEqual(['added customXml/_rels/item1.xml.rels', 'added customXml/item1.xml', 'added customXml/itemProps1.xml']);
  const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(saved);
  await pkg.customXmlParts.load();
  expect(pkg.customXmlParts.items[0]!.namespaceUri).toBe('urn:acme:invoice');
  expect(pkg.customXmlParts.items[0]!.selectSingleNode('/ns0:invoice/ns0:customer', "xmlns:ns0='urn:acme:invoice'")!.text).toBe('Ann');
});

// ED-005 section 12.64 (Jason): Copy XPath in the tree and in the XML view; the XML view edits the data part in Monaco,
// and its context menu binds the element at the cursor.
test("the Data tab: Copy XPath from the tree; the XML view's Apply refreshes bound controls, its context menu binds and copies", async ({ page }) => {
  await template(page);
  // The clipboard as the page sees it: what the panel writes is kept to read back.
  await page.evaluate(() => {
    const copied: string[] = [];
    (window as unknown as { copied: string[] }).copied = copied;
    Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: async (text: string) => { copied.push(text); } } });
  });
  const copied = (): Promise<string[]> => page.evaluate(() => (window as unknown as { copied: string[] }).copied);
  await page.getByTestId('tab-data').click();
  const name = '/invoice[1]/customer[1]/name[1]';
  // The header's Copy XPath copies the node selected in the tree.
  await expect(page.getByTestId('data-part-copy-xpath')).toBeDisabled();
  await page.locator(`[data-xpath="${name}"] > .tree-row > .node-name`).click();
  await page.getByRole('button', { name: `Copy the XPath ${name}` }).click();
  await expect(page.getByTestId('status')).toHaveText(`Copied ${name}`);
  expect(await copied()).toEqual([name]);
  // Ctrl+C on a node's name copies it too.
  await page.locator(`[data-xpath="/invoice[1]/customer[1]"] > .tree-row > .node-name`).focus();
  await page.keyboard.press('Control+c');
  await expect.poll(copied).toEqual([name, '/invoice[1]/customer[1]']);

  // The XML view: the part in Monaco.
  await page.getByTestId('data-part-view-xml').click();
  const text = await xmlEditorValue(page, 'data-part-xml');
  expect(text).toContain('<name>Joe Bloggs</name>');
  // Its context menu offers Bind and Copy XPath.
  await page.getByTestId('data-part-xml').locator('.monaco-editor .view-lines').click({ button: 'right' });
  await expect(page.locator('.monaco-menu')).toContainText('Bind to the control at the document');
  await expect(page.locator('.monaco-menu')).toContainText('Copy XPath');
  await page.keyboard.press('Escape');
  /** The XML editor's cursor put on a needle's first character plus `plus`, then one of the context menu's actions run. */
  const runAt = (needle: string, plus: number, action: string): Promise<void> => page.evaluate(([n, p, a]) => {
    const ed = (window as unknown as { docx4jXmlEditors: Record<string, { getModel(): { getValue(): string; getPositionAt(o: number): unknown }; setPosition(p: unknown): void; getAction(id: string): { run(): Promise<void> } }> }).docx4jXmlEditors['data-part-xml']!;
    const m = ed.getModel();
    ed.setPosition(m.getPositionAt(m.getValue().indexOf(n as string) + (p as number)));
    return ed.getAction(`docx4j.${a as string}`).run();
  }, [needle, plus, action] as const);
  await runAt('<price>', 2, 'copyXPath');
  await expect(page.getByTestId('status')).toHaveText('Copied /invoice[1]/items[1]/item[1]/price[1]');
  // Bind: "INVOICE" selected in the document, the first price at the cursor: a plain text control made and bound.
  const invoice = await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { textContent: string; nodeSize: number } } }; caretAt(a: number, b: number): void } }).docx4jEditor;
    let pos = 0;
    for (let i = 0; i < v.state.doc.childCount; i++) {
      const b = v.state.doc.child(i);
      if (b.textContent === 'INVOICE') { v.caretAt(pos + 1, pos + 8); return i; }
      pos += b.nodeSize;
    }
    return -1;
  });
  expect(invoice).toBeGreaterThan(0);
  await runAt('<price>', 2, 'bind');
  await expect(page.getByTestId('status')).toHaveText('Inserted a plain text control bound to /invoice[1]/items[1]/item[1]/price[1]');
  expect(await page.evaluate((i) => (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc.child(i).textContent, invoice)).toBe('$20');
  // An edit applied: the customer's name changes, and the control bound to it shows it.
  await setXmlEditor(page, (await xmlEditorValue(page, 'data-part-xml')).replace('Joe Bloggs', 'Jane Doe'), 'data-part-xml');
  // Not applied yet: the context menu works on the part as applied, and says so.
  await runAt('<price>', 2, 'copyXPath');
  await expect(page.getByTestId('status')).toHaveText('Apply the XML first: Bind and Copy XPath work on the part as applied');
  await page.getByTestId('part-apply').click();
  await expect(page.getByTestId('part-status')).toHaveText('Applied.');
  await expect(page.locator('.docx-editor .ProseMirror[aria-label="Document"] .sdt-block[data-label="Customer name"]')).toHaveText('Jane Doe');
  // The tree again shows the new value; the view chosen is remembered.
  await page.getByTestId('data-part-view-tree').click();
  await expect(page.locator(`[data-xpath="${name}"]`)).toHaveAttribute('aria-label', 'name: Jane Doe');
});

// ED-005 section 12.117: the Preview tab converts an XHTML-bound control's markup, as the processor does, though the
// document still shows its placeholder (the editor converts nothing at open). check37a-template.docx is core-ts check 37's
// template: a block-level control and a run-level one, bound through od:xpath entries, never filled. The heading is drawn
// in its style (which the body does not use, so the body's stylesheet has no rule for it) and the list with its bullets
// (from the preview's own numbering, where the conversion made the list).
test('the Preview tab shows XHTML-bound controls converted: the heading in its style, bold, a link, a bulleted list, a table; the run control\'s first paragraph', async ({ page }) => {
  await openFixture(page, 'xhtml/check37a-template.docx');
  await expect(page.locator(body)).toContainText('XHTML placeholder (block-level)');
  await page.getByTestId('perspective-template').click();
  await page.getByTestId('tab-preview').click();
  const story = page.getByTestId('preview-story').locator('.ProseMirror');
  const heading = story.locator('p.s-Heading1');
  await expect(heading).toHaveText('Delivery terms');
  // Heading 1 in the fixture's styles: bold, 14pt.
  const look = await heading.evaluate((el) => { const cs = getComputedStyle(el); return { weight: cs.fontWeight, size: parseFloat(cs.fontSize) }; });
  expect(look.weight).toBe('700');
  expect(look.size).toBeCloseTo((14 * 96) / 72, 2);
  await expect(story).toContainText('Goods are delivered within 14 days of the order, see the terms.');
  expect(await story.getByText('within 14 days', { exact: true }).evaluate((el) => getComputedStyle(el).fontWeight)).toBe('700');
  await expect(story.locator('a', { hasText: 'the terms' })).toHaveCount(1);
  const items = story.locator('p', { hasText: /^(•\s*)?(Standard delivery is free|Express delivery costs extra)\.$/ });
  await expect(items).toHaveCount(2);
  await expect(items.locator('.list-label')).toHaveText(['•', '•']);
  expect(await items.first().evaluate((el) => parseFloat(getComputedStyle(el).paddingInlineStart) + parseFloat(getComputedStyle(el).marginInlineStart))).toBeGreaterThan(0);
  await expect(story.locator('table td').filter({ hasText: 'Express' })).toHaveCount(1);
  // The run-level control takes the first paragraph's inlines, its italics kept; the second paragraph is dropped.
  await expect(story).toContainText('[First paragraph in italics');
  expect(await story.getByText('in italics', { exact: true }).evaluate((el) => getComputedStyle(el).fontStyle)).toBe('italic');
  await expect(story).not.toContainText('A second paragraph');
  await expect(story).not.toContainText('XHTML placeholder');
  // The document is untouched: its controls keep their placeholders, and the body draws no heading.
  await expect(page.locator(body)).toContainText('XHTML placeholder (block-level)');
  await expect(page.locator(body).locator('p.s-Heading1')).toHaveCount(0);
});
