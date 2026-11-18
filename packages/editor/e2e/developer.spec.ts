// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 6.5: the Developer perspective's panels.
import { test, expect } from '@playwright/test';
import { openFixture, download, fixture, comparePackages, caretInBlock, caretInCell, selectBlocks, xmlEditorValue, setXmlEditor } from './helpers.js';

async function developer(page: import('@playwright/test').Page, name: string): Promise<void> {
  await openFixture(page, name);
  await page.getByTestId('perspective-developer').click();
  await page.getByTestId('panel-xml').waitFor();
}

test('reveal codes: breadcrumbs, the run under the caret, an edit that applies, a bad edit refused, the API script', async ({ page }) => {
  await developer(page, 'invoice.docx');
  // Block 3 is a content control holding "Joe Bloggs": the chain shows the sdt.
  await caretInBlock(page, 3, 2);
  const crumbs = page.getByTestId('panel-xml').getByRole('navigation', { name: 'Ancestors' });
  await expect(crumbs).toContainText('body > sdt');
  await expect(crumbs).toContainText('> p');
  // The XML is shown in Monaco, read-only (ED-005 section 12.99): a test reads it through the live instance.
  await expect.poll(() => xmlEditorValue(page)).toContain('<w:sdt');
  // The run crumb narrows the XML to the run.
  await page.getByTestId('crumb-run').click();
  await expect.poll(() => xmlEditorValue(page)).toContain('<w:r');
  await expect.poll(() => xmlEditorValue(page)).not.toContain('<w:sdt');
  // Back to the block; edit it: give the first paragraph a style.
  await developer(page, 'tracked-changes.docx');
  await caretInBlock(page, 4);
  await expect.poll(() => xmlEditorValue(page)).toContain('Some text.');
  await page.getByTestId('xml-edit').click();
  const xml = await xmlEditorValue(page);
  // Monaco's own find opens inside the pane (ED-005 section 3.8), not the document's bar.
  await page.getByTestId('xml-editor').locator('.monaco-editor textarea').focus();
  await page.keyboard.press('Control+f');
  await expect(page.getByTestId('xml-editor').locator('.find-widget.visible')).toHaveCount(1);
  await expect(page.getByTestId('find-bar')).toHaveCount(0);
  await page.keyboard.press('Escape');
  await setXmlEditor(page, xml.replace('<w:r>', '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r>'));
  await page.getByTestId('xml-apply').click();
  await expect(page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p').nth(4)).toHaveClass(/s-Heading1/);
  await expect(page.getByTestId('style-picker')).toHaveValue('Heading1');
  // Malformed XML: the parser's message names the line, the line is marked, nothing changes.
  await page.getByTestId('xml-edit').click();
  await setXmlEditor(page, '<w:p>\n  <w:r>\n    <w:t>Run properties</w:t\n  </w:r>\n</w:p>');
  await expect(page.getByTestId('xml-error')).toContainText('line 4');
  await expect(page.getByTestId('xml-editor')).toHaveAttribute('aria-invalid', 'true');
  await expect(page.getByTestId('xml-editor')).toHaveAttribute('data-error-line', '4');
  await page.getByTestId('xml-apply').click();
  await expect(page.getByTestId('xml-error')).not.toContainText('parsererror');
  await expect(page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p').nth(4)).toHaveText('Some text.');
  // Well-formed but wrong: the unmarshaller's error, no line.
  await setXmlEditor(page, '<w:p><w:nonsense/></w:p>');
  await expect(page.getByTestId('xml-editor')).not.toHaveAttribute('data-error-line');
  await page.getByTestId('xml-apply').click();
  await expect(page.getByTestId('xml-error')).toBeVisible();
  await page.getByRole('button', { name: 'Cancel' }).click();
  // The API-script pane.
  await page.getByTestId('panel-xml').getByLabel('API script').check();
  await expect(page.getByTestId('api-script')).toContainText('insertParagraph');
  // Undo takes the XML edit back as one step (the caret back in the editor first).
  await caretInBlock(page, 4);
  await page.keyboard.press('Control+z');
  await expect(page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p').nth(4)).not.toHaveClass(/s-Heading1/);
});

test('console: a script edits the tree, shows on the page, is one undo step, and the download carries it', async ({ page }) => {
  await developer(page, 'tracked-changes.docx');
  const console = page.getByTestId('panel-console');
  await console.waitFor();
  await expect(console.locator('.console-editor .monaco-editor')).toBeVisible();
  // Replace the editor's text through Monaco's model.
  await page.evaluate(() => {
    const m = (window as unknown as { monaco?: { editor: { getModels(): { setValue(v: string): void }[] } } }).monaco;
    void m;
  });
  const script = `const p = body.paragraphs[4];
p.insertText(' from the console', 'End');
const h = p.insertParagraph('Console heading', 'After');
h.style = 'Heading 1';
selection.paragraphs[0].font.bold = true;
console.log('paragraphs', body.paragraphs.length);
return h.text;`;
  await caretInBlock(page, 1);
  await console.locator('.console-editor textarea').first().focus();
  await page.keyboard.press('Control+a');
  await page.keyboard.insertText(script);
  await page.getByTestId('console-run').click();
  const output = page.getByTestId('console-output');
  await expect(output).toContainText('paragraphs');
  await expect(output).toContainText('Console heading');
  const paragraphs = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p');
  await expect(paragraphs.nth(4)).toHaveText('Some text. from the console');
  await expect(paragraphs.nth(5)).toHaveText('Console heading');
  await expect(paragraphs.nth(5)).toHaveClass(/s-Heading1/);
  await expect(page.getByRole('button', { name: 'Bold' })).toHaveAttribute('aria-pressed', 'true');
  const saved = await download(page);
  const { EditorDocument } = await import('@docx4j/editor-model');
  const session = await EditorDocument.open(saved);
  expect(session.doc.child(5).textContent).toBe('Console heading');
  expect(comparePackages(await fixture('tracked-changes.docx'), saved, ['word/document.xml', 'docProps/app.xml']).filter((d) => d.name !== 'docProps/app.xml')).toEqual([]);
  // One undo step.
  await caretInBlock(page, 4);
  await page.keyboard.press('Control+z');
  await expect(paragraphs.nth(4)).toHaveText('Some text.');
  await expect(paragraphs.nth(5)).not.toHaveText('Console heading');
});

test('load all types adds the other module declarations to the console; the factory-source pane says what it waits for', async ({ page }) => {
  await developer(page, 'tracked-changes.docx');
  const console = page.getByTestId('panel-console');
  await expect(console.locator('.console-editor .monaco-editor')).toBeVisible();
  await expect(page.getByTestId('console-types')).toContainText('types loaded');
  await page.getByTestId('console-load-all-types').click();
  // The count is the object model's (103 modules besides wml in objects 0.2.0, 104 in 0.3.0), not the editor's.
  await expect(page.getByTestId('console-types')).toContainText(/all \d+ modules/, { timeout: 60_000 });
  await page.getByTestId('panel-xml').getByLabel('Factory source').check();
  await expect(page.getByTestId('factory-source-note')).toContainText('toSource');
});

test('console: a throwing script prints its line', async ({ page }) => {
  await developer(page, 'tracked-changes.docx');
  const console = page.getByTestId('panel-console');
  await expect(console.locator('.console-editor .monaco-editor')).toBeVisible();
  await console.locator('.console-editor textarea').first().focus();
  await page.keyboard.press('Control+a');
  await page.keyboard.insertText("console.log('before');\nthrow new Error('deliberate');");
  await page.getByTestId('console-run').click();
  const output = page.getByTestId('console-output');
  await expect(output).toContainText('before');
  await expect(output).toContainText('deliberate');
  await expect(output).toContainText('line 2');
});

test('parts and outline', async ({ page }) => {
  await developer(page, 'tracked-changes.docx');
  await page.getByTestId('tab-parts').click();
  const parts = page.getByTestId('panel-parts');
  const tree = page.getByTestId('parts-tree');
  // The main part hangs off the package relationships; the styles part off the main part's.
  const main = tree.getByRole('treeitem').filter({ has: page.getByRole('button', { name: '/word/document.xml', exact: true }) }).first();
  await expect(main).toContainText('officeDocument');
  await expect(main.getByRole('button', { name: '/word/styles.xml', exact: true })).toBeVisible();
  // Collapsing the main part hides its children; expanding brings them back.
  await main.getByRole('button', { name: 'Collapse' }).first().click();
  await expect(tree.getByRole('button', { name: '/word/styles.xml', exact: true })).toHaveCount(0);
  await main.getByRole('button', { name: 'Expand' }).first().click();
  // The main part shows its relationships part between the tree and its content.
  await tree.getByRole('button', { name: '/word/document.xml', exact: true }).click();
  await expect(page.getByTestId('part-rels')).toContainText('/word/_rels/document.xml.rels');
  await expect(page.getByTestId('part-rels')).toContainText('Relationship ');
  await expect(page.getByTestId('part-content')).toContainText('w:document');
  // A part without relationships shows its content alone.
  await tree.getByRole('button', { name: '/word/styles.xml', exact: true }).click();
  await expect(page.getByTestId('part-rels')).toHaveCount(0);
  await expect(page.getByTestId('part-content')).toContainText('w:styles');
  await expect(page.getByTestId('part-content')).toContainText('as in the file');
  await page.getByTestId('tab-outline').click();
  await expect(page.getByTestId('panel-parts')).toBeHidden();   // the parts tab stays mounted, out of sight
  const outline = page.getByTestId('panel-outline');
  await expect(outline).toContainText('Heading1');
  await expect(outline).toContainText('body/2');
  await outline.getByRole('button', { name: 'body/2', exact: true }).click();
  await expect(page.getByTestId('style-picker')).toHaveValue('Heading1');
});

test('the self-check dialog shows both sides as XML and offers the diagnosis download', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { session: { selfCheck: unknown } } }).docx4jEditor;
    v.session.selfCheck = async () => ({ equal: false, difference: { index: 2, live: '{"type":"paragraph"}', reloaded: '{"type":"paragraph","x":1}' } });
  });
  await page.getByTestId('download').click();
  const dialog = page.getByTestId('self-check-dialog');
  await expect(dialog).toBeVisible();
  await expect(dialog).toContainText('Block 2 differs');
  await expect(dialog.locator('pre').first()).toContainText('<w:p');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByTestId('self-check-download').click()]);
  expect(dl.suggestedFilename()).toBe('tracked-changes-selfcheck-failed.docx');
  await expect(page.getByTestId('status')).toContainText('Not saved');
});

test("reveal codes in a table marks the caret's cell and folds the other rows and cells (ED-005 section 12.99)", async ({ page }) => {
  await developer(page, 'tables/table-fit-fixed.docx');
  // Table B (block 4): a heading over two columns, then a vertical merge in the first column; the caret in the
  // third row's first cell, which starts at grid column 1 behind the continuation.
  await caretInCell(page, 4, 2, 0);
  const crumbs = page.getByTestId('panel-xml').getByRole('navigation', { name: 'Ancestors' });
  await expect(crumbs).toContainText('body > tbl > tr > tc > p');
  await expect.poll(() => xmlEditorValue(page)).toContain('Where the premises are damaged');
  const text = await xmlEditorValue(page);
  const lines = text.split('\n');
  const rows = lines.flatMap((l, i) => (/^\s*<w:tr[ >]/.test(l) ? [i + 1] : []));
  expect(rows).toHaveLength(3);
  const host = page.getByTestId('xml-editor');
  await expect(host).toHaveAttribute('data-focus', /^\d+-\d+$/);
  const [from, to] = (await host.getAttribute('data-focus'))!.split('-').map(Number) as [number, number];
  expect(from).toBeGreaterThan(rows[2]!);
  expect(lines[from - 1]).toMatch(/<w:tc>/);
  expect(lines.slice(from - 1, to).join('\n')).toContain('Where the premises are damaged');
  // The folds, read from Monaco's folding model: the first two rows and, in the third, the two other cells.
  const collapsed = () => page.evaluate(async () => {
    type Regions = { length: number; isCollapsed(i: number): boolean; getStartLineNumber(i: number): number };
    const ed = (window as unknown as { docx4jXmlEditors: Record<string, { getContribution(id: string): { getFoldingModel(): Promise<{ regions: Regions } | null> } }> }).docx4jXmlEditors['xml-editor']!;
    const model = await ed.getContribution('editor.contrib.folding').getFoldingModel();
    const out: number[] = [];
    if (model) for (let i = 0; i < model.regions.length; i++) if (model.regions.isCollapsed(i)) out.push(model.regions.getStartLineNumber(i));
    return out;
  });
  await expect.poll(collapsed).toContain(rows[0]);
  const folded = await collapsed();
  expect(folded).toContain(rows[1]);
  expect(folded).not.toContain(rows[2]);
  const cells = lines.flatMap((l, i) => (/^\s*<w:tc>/.test(l) && i + 1 > rows[2]! ? [i + 1] : []));
  expect(cells).toHaveLength(3);
  expect(folded.filter((l) => cells.includes(l))).toEqual(cells.filter((l) => l !== from));
  expect(folded).toHaveLength(4);
  // The cell's lines are marked and in view.
  await expect(host.locator('.view-overlays .xml-focus').first()).toBeVisible();
  // The caret in the heading row's second cell moves the focus: two rows and the merged cell folded.
  await caretInCell(page, 4, 0, 1);
  await expect(host).not.toHaveAttribute('data-focus', `${from}-${to}`);
  await expect.poll(async () => (await collapsed()).length).toBe(3);
  // A paragraph outside the table has no focus; Edit keeps what is shown.
  await caretInBlock(page, 3);
  await expect(host).not.toHaveAttribute('data-focus', /./);
  await expect.poll(() => xmlEditorValue(page)).not.toContain('<w:tbl');
  await caretInCell(page, 4, 2, 0);
  await expect(host).toHaveAttribute('data-focus', `${from}-${to}`);
  await page.getByTestId('xml-edit').click();
  await expect(host).not.toHaveAttribute('data-readonly', 'true');
  expect(await xmlEditorValue(page)).toBe(text);
  await page.getByTestId('panel-xml').getByRole('button', { name: 'Cancel' }).click();
  // The caret's readout under the editor (Jason, 2026-10-15): its line and column, and the code point after it,
  // the glyph beside it when it draws one; an optional break put into the cell's text is otherwise invisible.
  await page.evaluate(() => (window as unknown as { docx4jEditor: { run(name: string): boolean } }).docx4jEditor.run('range.insertOptionalBreak'));
  await expect.poll(() => xmlEditorValue(page)).toContain('\u200B');
  const shown = await xmlEditorValue(page);
  const lineIndex = shown.split('\n').findIndex((l) => l.includes('\u200B'));
  const column = shown.split('\n')[lineIndex]!.indexOf('\u200B') + 1;
  const setPosition = (lineNumber: number, col: number) => page.evaluate(([l, c]) => { (window as unknown as { docx4jXmlEditors: Record<string, { setPosition(p: { lineNumber: number; column: number }): void }> }).docx4jXmlEditors['xml-editor']!.setPosition({ lineNumber: l, column: c }); }, [lineNumber, col] as const);
  await setPosition(lineIndex + 1, column);
  const readout = page.getByTestId('xml-editor-caret');
  await expect(readout).toContainText(`Ln ${lineIndex + 1}, Col ${column}`);
  await expect(readout).toContainText('U+200B');
  await expect(readout).not.toContainText('\u201C');
  await setPosition(lineIndex + 1, column + 1);
  await expect(readout).toContainText(/U\+00[0-9A-F]{2} \u201C\w\u201D/);
});

test('reveal codes shows every selected block, and an edit replaces them as a group', async ({ page }) => {
  await developer(page, 'tracked-changes.docx');
  await selectBlocks(page, 2, 4);   // the heading, the empty paragraph, "Some text."
  const pane = page.getByTestId('panel-xml');
  await expect(pane).toContainText('3 blocks selected');
  await expect.poll(() => xmlEditorValue(page)).toContain('This is style Heading 1');
  const shown = await xmlEditorValue(page);
  expect(shown).toContain('Some text.');
  expect(shown.split('<w:p ').length - 1 + shown.split('<w:p>').length - 1).toBe(3);
  await expect(page.getByTestId('crumb-run')).toHaveCount(0);
  await page.getByTestId('xml-edit').click();
  const text = await xmlEditorValue(page);
  // Drop the empty paragraph between the two: three blocks become two. Top-level paragraphs
  // start at column 0 in the pretty-printed text; nested ones are indented.
  const blocks = text.split(/\n(?=<w:p[ >/])/);
  expect(blocks.length).toBe(3);
  await setXmlEditor(page, [blocks[0], blocks[2]].join('\n'));
  await page.getByTestId('xml-apply').click();
  const paragraphs = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p');
  await expect(paragraphs.nth(2)).toHaveText('This is style Heading 1');
  await expect(paragraphs.nth(3)).toHaveText('Some text.');
  await caretInBlock(page, 3);
  await page.keyboard.press('Control+z');
  await expect(paragraphs.nth(3)).toHaveText('');
  await expect(paragraphs.nth(4)).toHaveText('Some text.');
});

test('the side panel is tabbed and resizable: the tab and the width persist across a reload', async ({ page }) => {
  await developer(page, 'tracked-changes.docx');
  const tabs = page.getByRole('tablist', { name: 'Panel tabs' });
  // Six tabs in E1 to E3, the XPath tab the seventh since E4.a (ED-005 section 3.5).
  await expect(tabs.getByRole('tab')).toHaveCount(7);
  await expect(page.getByTestId('tab-codes')).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByTestId('panel-console')).toBeVisible();
  // Arrow keys walk the tabs.
  await page.getByTestId('tab-codes').focus();
  await page.keyboard.press('ArrowRight');
  await expect(page.getByTestId('tab-format')).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByTestId('panel-xml')).toBeHidden();
  await page.getByTestId('tab-parts').click();
  // Drag the separator 200 px to the left: the panel grows by 200.
  const panels = page.getByTestId('panels');
  const before = (await panels.boundingBox())!.width;
  const splitter = page.getByTestId('splitter');
  const box = (await splitter.boundingBox())!;
  await page.mouse.move(box.x + box.width / 2, box.y + 100);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width / 2 - 100, box.y + 100);
  await page.mouse.move(box.x + box.width / 2 - 200, box.y + 100);
  await page.mouse.up();
  const after = (await panels.boundingBox())!.width;
  expect(Math.round(after - before)).toBe(200);
  // Both survive a reload.
  await page.reload();
  await page.getByTestId('open').setInputFiles(new URL('../../editor-model/test/fixtures/tracked-changes.docx', import.meta.url).pathname);
  await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await expect(page.getByTestId('tab-parts')).toHaveAttribute('aria-selected', 'true');
  expect(Math.round((await panels.boundingBox())!.width)).toBe(Math.round(after));
  // The Panels menu closes the side panel.
  await page.locator('.panels-menu > summary[aria-label="Panels"]').click();
  await page.getByTestId('option-panel').uncheck();
  await expect(panels).toHaveCount(0);
});

test('the diagnostics are grouped by kind and filtered to the selection (ED-003 section 8.6)', async ({ page }) => {
  await developer(page, 'MergeWithNext.docx');
  await page.getByTestId('tab-diagnostics').click();
  const list = page.getByTestId('diagnostic-list');
  const rows = list.locator('details.diagnostic-kind');
  await expect(rows).toHaveCount(3);
  await expect(rows.nth(0).locator('> summary')).toHaveText('fldChar 6');
  await expect(rows.nth(1).locator('> summary')).toHaveText('fldSimple 4');
  await expect(rows.nth(2).locator('> summary')).toHaveText('instrText 2');
  // Closed until clicked; open, the addresses are buttons that select the element.
  await expect(rows.nth(1).locator('li').first()).toBeHidden();
  await rows.nth(1).locator('> summary').click();
  await expect(rows.nth(1).locator('li')).toHaveCount(4);
  await rows.nth(1).getByRole('button', { name: /^fldSimple at / }).first().click();
  const selected = await page.evaluate(() => { const s = (window as unknown as { docx4jEditor: { state: { selection: { node?: { attrs: { kind: string } } } } } }).docx4jEditor.state.selection; return s.node?.attrs.kind; });
  expect(selected).toBe('fldSimple');
  // The filter keeps what the selection touches: the selected field alone.
  await page.getByTestId('diagnostic-list-in-selection').check();
  await expect(rows).toHaveCount(1);
  await expect(rows.first().locator('> summary')).toHaveText('fldSimple 1');
  await page.getByTestId('diagnostic-list-in-selection').uncheck();
  await expect(rows).toHaveCount(3);
});

test('the panel sections fold, and stay folded across a reload (ED-003 section 8.6)', async ({ page }) => {
  await developer(page, 'tracked-changes.docx');
  const consolePanel = page.getByTestId('panel-console');
  await expect(consolePanel.locator('> details')).toHaveAttribute('open', '');
  await expect(page.getByTestId('console-run')).toBeVisible();
  // The console folded under reveal codes: its body goes, its heading stays; reveal codes is untouched.
  await consolePanel.locator('> details > summary').click();
  await expect(consolePanel.locator('> details')).not.toHaveAttribute('open', '');
  await expect(page.getByTestId('console-run')).toBeHidden();
  await expect(consolePanel.getByRole('heading', { name: /Console/ })).toBeVisible();
  await expect(page.getByTestId('xml-editor')).toBeVisible();
  // The state is this browser's: a reload brings the console back folded, and a click opens it.
  await page.reload();
  await page.getByTestId('open').setInputFiles(new URL('../../editor-model/test/fixtures/tracked-changes.docx', import.meta.url).pathname);
  await page.getByTestId('panel-console').waitFor();
  await expect(page.getByTestId('panel-console').locator('> details')).not.toHaveAttribute('open', '');
  await expect(page.getByTestId('console-run')).toBeHidden();
  await page.getByTestId('panel-console').locator('> details > summary').click();
  await expect(page.getByTestId('console-run')).toBeVisible();
  await expect(page.getByTestId('panel-console').locator('.console-editor .monaco-editor')).toBeVisible();
});

test('the Format tab shows the effective run, paragraph and section properties at the caret', async ({ page }) => {
  await developer(page, 'tracked-changes.docx');
  await page.getByTestId('tab-format').click();
  await caretInBlock(page, 2);   // "This is style Heading 1"
  const font = page.getByTestId('format-rPr');
  await expect(font).toContainText('w:b');
  await expect(font).toContainText('w:sz');
  await expect(font.locator('.attr-value').first()).toBeVisible();
  const paragraph = page.getByTestId('format-pPr');
  await expect(paragraph).not.toContainText('nothing set');
  const section = page.getByTestId('format-sectPr');
  await expect(section).toContainText('w:pgSz');
  await expect(section).toContainText('w:pgMar');
  // The plain paragraph is not bold.
  await caretInBlock(page, 4);   // "Some text."
  await expect(font.locator('.element', { hasText: /^w:sz$/ })).toHaveCount(1);
  await expect(font.locator('.element', { hasText: /^w:b$/ })).toHaveCount(0);
});

test('the Styles tab climbs the hierarchy at the caret; a style edit applies to the page and saves', async ({ page }) => {
  await developer(page, 'tracked-changes.docx');
  await page.getByTestId('tab-styles').click();
  await caretInBlock(page, 2);
  const paragraph = page.getByTestId('styles-paragraph');
  const items = paragraph.locator('ol.styles li:not(.layer)');
  await expect(items.first()).toContainText('Heading1');
  await expect(items.nth(1)).toContainText('Normal');
  // The stack has the direct formatting on top and the document defaults at the bottom.
  await expect(paragraph.locator('ol.styles li').first()).toContainText('Direct paragraph formatting');
  await expect(paragraph.locator('ol.styles li').last()).toContainText('Document defaults, paragraphs');
  // A run with no character style carries the default one.
  const character = page.getByTestId('styles-character');
  await expect(character).toContainText('Default Paragraph Font');
  await expect(character).toContainText('the run names none');
  await expect(character.locator('ol.styles li').first()).toContainText('direct run formatting');   // none on this run
  await expect(page.getByTestId('styles-table')).toContainText('not in a table');
  // The default paragraph style when the paragraph names none.
  await caretInBlock(page, 4);
  await expect(items.first()).toContainText('Normal');
  await expect(paragraph).toContainText('default');
  // The document defaults' XML.
  await page.getByTestId('defaults-rPrDefault').click();
  await expect(page.getByTestId('style-xml-text')).toContainText('<w:rPrDefault');
  const heading = page.locator('.docx-editor .ProseMirror[aria-label="Document"] > p.s-Heading1').first();
  // A style's XML, edited: Heading 1 becomes italic on the page.
  await caretInBlock(page, 2);
  await expect(items.first()).toContainText('Heading1');
  await items.first().getByRole('button').click();
  await expect(page.getByTestId('style-xml-text')).toContainText('w:styleId="Heading1"');
  expect(await heading.evaluate((el) => getComputedStyle(el).fontStyle)).toBe('normal');
  await page.getByTestId('style-edit').click();
  const text = await xmlEditorValue(page);
  expect(text).toContain('<w:rPr>');
  await setXmlEditor(page, text.replace('<w:rPr>', '<w:rPr><w:i/>'));
  await page.getByTestId('style-apply').click();
  await expect(page.getByTestId('style-xml-text')).toContainText('<w:i/>');
  expect(await heading.evaluate((el) => getComputedStyle(el).fontStyle)).toBe('italic');
  await expect(page.getByTestId('name')).toContainText('•');
  // A bad edit is refused with its line.
  await page.getByTestId('style-edit').click();
  await setXmlEditor(page, '<w:style w:styleId="Heading1">\n<w:name w:val="x"\n</w:style>');
  await page.getByTestId('style-apply').click();
  await expect(page.getByTestId('style-error')).toContainText('line');
  await page.getByRole('button', { name: 'Cancel' }).click();
  // The save carries the style; the main part is untouched.
  const saved = await download(page);
  expect(comparePackages(await fixture('tracked-changes.docx'), saved, ['word/styles.xml', 'docProps/app.xml'])).toEqual([]);
  const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(saved);
  const styles = await pkg.getMainDocumentPart().styleDefinitionsPart!.readContents();
  expect(styles.style!.find((s) => s.styleId === 'Heading1')!.rPr!.i).toBeDefined();
  // Direct paragraph formatting edited through the tree (which makes the main part live, hence
  // after the download): the heading centres; undo takes it back.
  await caretInBlock(page, 2);
  await expect(items.first()).toContainText('Heading1');
  await page.getByTestId('direct-pPr').click();
  await expect(page.getByTestId('style-xml-text')).toContainText('w:pStyle');
  await page.getByTestId('style-edit').click();
  await setXmlEditor(page, (await xmlEditorValue(page)).replace('</w:pPr>', '<w:jc w:val="center"/></w:pPr>'));
  await page.getByTestId('style-apply').click();
  await expect(heading).toHaveAttribute('style', /text-align: center/);
  await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').focus();
  await page.keyboard.press('Control+z');
  await expect(heading).not.toHaveAttribute('style', /text-align: center/);
});

test('Select by XPath (ED-005 section 3.5): hits listed on Enter, a click selects where the hit projects, an edit is seen, a bad expression says why', async ({ page }) => {
  await developer(page, 'tracked-changes.docx');
  await page.getByTestId('tab-xpath').click();
  const panel = page.getByTestId('panel-xpath');
  await expect(panel.getByTestId('xpath-run')).toBeEnabled();
  await panel.getByTestId('xpath-expression').fill("//w:p[w:pPr/w:pStyle/@w:val = 'Heading1']");
  await panel.getByTestId('xpath-expression').press('Enter');
  await expect(panel.getByTestId('xpath-count')).toContainText('hit');
  const count = await panel.getByTestId('xpath-hit').count();
  expect(count).toBeGreaterThan(0);
  // A click selects the paragraph's text in the body.
  await panel.getByTestId('xpath-hit').first().click();
  const selected = await page.evaluate(() => { const v = (window as unknown as { docx4jEditor: { state: { doc: { textBetween(a: number, b: number): string }; selection: { from: number; to: number } } } }).docx4jEditor; return v.state.doc.textBetween(v.state.selection.from, v.state.selection.to); });
  expect(selected.length).toBeGreaterThan(0);
  // Typing is seen by the next select: the edited block is where the hit is placed.
  await caretInBlock(page, 4);
  await page.keyboard.type('Zebra ');
  await panel.getByTestId('xpath-expression').fill("//w:t[contains(., 'Zebra')]");
  await panel.getByTestId('xpath-run').click();
  await expect(panel.getByTestId('xpath-count')).toContainText('1 hit');
  await expect(panel.getByTestId('xpath-hit')).toHaveText('body/4');
  // XPath 3.1 (FontoXPath): a function XPath 1.0 lacks, refused by the browser's engine, answered once the option is on.
  await panel.getByTestId('xpath-expression').fill("//w:t[matches(., '^Z')]");
  await panel.getByTestId('xpath-run').click();
  await expect(panel.getByTestId('xpath-error')).toBeVisible();
  await panel.getByTestId('xpath-31').check();
  await panel.getByTestId('xpath-run').click();
  await expect(panel.getByTestId('xpath-count')).toContainText('1 hit', { timeout: 20_000 });
  await panel.getByTestId('xpath-31').uncheck();
  await panel.getByTestId('xpath-expression').fill('//w:p[');
  await panel.getByTestId('xpath-run').click();
  await expect(panel.getByTestId('xpath-error')).toBeVisible();
  // Open in console: the select as a script.
  await panel.getByTestId('xpath-console').click();
  await expect(page.getByTestId('tab-codes')).toHaveAttribute('aria-selected', 'true');
  await expect.poll(() => page.evaluate(() => (window as unknown as { docx4jConsole?: { getValue(): string } }).docx4jConsole?.getValue() ?? '')).toContain("await body.select(\"//w:p[\")");
});

test('the normalisation report: made on the button, by part and kind; it leaves an unedited file saving byte for byte and the document clean (ED-008 section 8.8)', async ({ page }) => {
  await developer(page, 'numbering-clauses.docx');
  await page.getByTestId('tab-diagnostics').click();
  await expect(page.getByTestId('normalisation-report')).toHaveCount(0);
  await page.getByTestId('normalisation-run').click();
  const report = page.getByTestId('normalisation-report');
  await expect(page.getByTestId('normalisation-summary')).toHaveText('Values not in normal form: 1,053');
  await expect(report.locator('[data-part="/word/styles.xml"]')).toContainText('1053');
  const fill = report.locator('details[data-kind="w:shd/@w:fill"]');
  await fill.locator('summary').click();
  await expect(fill.locator('li').first()).toContainText(/[0-9a-f]{6} would be written as [0-9A-F]{6}/);

  await expect(page.getByTestId('name')).toHaveText('numbering-clauses.docx');
  expect(comparePackages(await fixture('numbering-clauses.docx'), await download(page))).toEqual([]);
});

test('the normalisation report reads the document as it is now, in every part: an edit in a header takes that part\'s value out; a file in normal form says so (ED-008 section 8.8)', async ({ page }) => {
  await developer(page, 'normal-form-values.docx');
  await page.getByTestId('tab-diagnostics').click();
  await page.getByTestId('normalisation-run').click();
  const report = page.getByTestId('normalisation-report');
  await expect(report.locator('[data-part="/word/header3.xml"]')).toBeVisible();
  await expect(report.locator('details[data-kind="w:highlight"] li')).toContainText('purple would be left out');

  // The first page's header is header3: its one coloured run emptied, the value is gone from the file the next save would write.
  const header = page.locator('.docx-editor .page-start .boundary-header .ProseMirror');
  await header.click();
  await page.keyboard.press('Control+A');
  await page.keyboard.press('Delete');
  await expect(header).not.toContainText('First page');
  await page.getByTestId('normalisation-run').click();
  await expect(report.locator('[data-part="/word/header1.xml"]')).toBeVisible();
  await expect(report.locator('[data-part="/word/header3.xml"]')).toHaveCount(0);

  await openFixture(page, 'tracked-changes.docx');
  await page.getByTestId('normalisation-run').click();
  await expect(page.getByTestId('normalisation-summary')).toHaveText('Every value is in its normal form.');
});
