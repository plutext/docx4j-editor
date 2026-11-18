// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 4.10 item 4: the Template perspective's tabs over a fixture in jsdom. The
// Control tab shows and edits a control's properties, the binding and the role; the Controls
// tab lists every control and a click selects it; the Data tab shows the data part first,
// evaluates an XPath, and edits a value in place.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { TextSelection } from 'prosemirror-state';
import { CellSelection } from 'prosemirror-tables';
import { ZipPartStore } from '@docx4j/core-ts';
import { EditorDocument, readControl, readControlEntries, controlAt, type SdtGroup } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import ControlPanel from '../src/app/panels/template/ControlPanel.js';
import ControlsTree from '../src/app/panels/template/ControlsTree.js';
import DataPanel from '../src/app/panels/template/DataPanel.js';
const MOVE_MIME = 'application/x-docx4j-move';
import PreviewPanel from '../src/app/panels/template/PreviewPanel.js';
import type { TabContext } from '../src/app/perspectives.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

interface Mounted {
  session: EditorDocument;
  editor: DocxEditorView;
  root: Root;
  host: HTMLElement;
  status: string[];
  tabs: string[];
}

async function mount(name: string, Panel: (ctx: TabContext) => React.ReactNode): Promise<Mounted> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));
  const container = document.createElement('div');
  document.body.appendChild(container);
  const editor = new DocxEditorView(container, session);
  const host = document.createElement('div');
  document.body.appendChild(host);
  const root = createRoot(host);
  const status: string[] = [];
  const tabs: string[] = [];
  const ctx: TabContext = {
    editor, active: editor, session, perspective: { id: 'template', panel: { open: true, tab: 'control' } }, consoleScript: null,
    onConsoleScript: () => {}, onSaveMacro: () => {}, onStatus: (s) => status.push(s), onTab: (t) => tabs.push(t),
  };
  await act(async () => { root.render(<Panel {...ctx} />); });
  await until(() => host.querySelector('[data-testid="panel-control"], [data-testid="panel-controls"], [data-testid="panel-data"]') !== null);
  return { session, editor, root, host, status, tabs };
}

async function until(check: () => boolean, ms = 5000): Promise<void> {
  const start = Date.now();
  while (!check()) {
    if (Date.now() - start > ms) throw new Error('timed out');
    await act(async () => { await new Promise((r) => setTimeout(r, 20)); });
  }
}

function blockStart(editor: DocxEditorView, i: number): number {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += editor.state.doc.child(k).nodeSize;
  return pos;
}

async function caret(editor: DocxEditorView, pos: number): Promise<void> {
  await act(async () => { editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos))); });
}

function input(host: HTMLElement, testId: string): HTMLInputElement {
  return host.querySelector<HTMLInputElement>(`[data-testid="${testId}"]`)!;
}

/** Set a text box's value without committing it (the form's first field; Enter applies the form). */
async function fill(el: HTMLInputElement, value: string): Promise<void> {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
  await act(async () => {
    setter.call(el, value);
    el.dispatchEvent(new Event('input', { bubbles: true }));
  });
}

async function type(el: HTMLInputElement, value: string): Promise<void> {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
  await act(async () => {
    setter.call(el, value);
    el.dispatchEvent(new Event('input', { bubbles: true }));
  });
  await act(async () => { el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })); });
}

describe('the Control tab', () => {
  test('shows the block control at the caret with its binding and verdict; a title edit goes through the command and undoes', async () => {
    const { session, editor, host, root } = await mount('invoice.docx', ControlPanel);
    // Block 3 is the "Customer name" control (a plain-text control bound to the customer's name).
    const control = editor.state.doc.child(3);
    expect(control.type.name).toBe('sdtBlock');
    await caret(editor, blockStart(editor, 3) + 2);
    await until(() => host.querySelector('[data-testid="control-binding"]') !== null);
    expect(host.querySelector('[data-testid="control-summary"]')!.textContent).toContain('block · Plain text · id 1418037941');
    expect(input(host, 'cc-title').value).toBe('Customer name');
    expect(host.querySelector('[data-testid="cc-binding-selects"]')!.textContent).toContain('/invoice[1]/customer[1]/name[1]');
    expect(host.querySelector('[data-testid="cc-binding-selects"]')!.textContent).toContain('Joe Bloggs');
    expect(host.querySelector('[data-testid="cc-verdict"]')!.textContent).toContain('Joe Bloggs');
    expect(host.querySelector('[data-testid="cc-role"]')!.textContent).toBe('none');
    // A bound control takes no condition or repeat (REQ-008): the role form says so.
    expect(host.querySelector('[data-testid="cc-role-cannot"]')!.textContent).toContain('REQ-008');
    await type(input(host, 'cc-title'), 'Customer');
    expect(readControl(editor.state.doc.child(3).attrs['sdtPr']).title).toBe('Customer');
    expect(session.exportTree().changed).toBe(true);
    await act(async () => { editor.run('history.undo'); });
    expect(readControl(editor.state.doc.child(3).attrs['sdtPr']).title).toBe('Customer name');
    expect(input(host, 'cc-title').value).toBe('Customer name');
    root.unmount();
    editor.destroy();
  });

  test('a whole repeat row selected shows the repeat with its count; a tag with two roles is refused (REQ-007); a cell control edits the map entry in one undo step', async () => {
    const { session, editor, host, root, status } = await mount('invoice.docx', ControlPanel);
    // Select the repeat row (block 5, row 1) as a cell selection: the row group, not its first cell's control.
    await act(async () => { editor.run('contentControl.select', { groupId: (editor.state.doc.child(5).child(1).attrs['sdt'] as SdtGroup[])[0]!.groupId }); });
    expect(editor.state.selection).toBeInstanceOf(CellSelection);
    expect(controlAt(editor.state)?.form).toBe('row');
    await until(() => host.querySelector('[data-testid="cc-role"]')?.textContent === 'repeat x2');
    expect(host.querySelector('[data-testid="control-summary"]')!.textContent).toContain('row · Rich text');
    expect(host.querySelector('[data-testid="cc-repeat-count"]')!.textContent).toContain('3 against the data part');
    // Two roles in one tag are refused with the reason, and nothing is written.
    await type(input(host, 'cc-tag'), 'od:repeat=x2&od:condition=c5');
    expect(host.querySelector('[data-testid="cc-tag-error"]')!.textContent).toContain('REQ-007');
    const groupId = controlAt(editor.state)!.form === 'row' ? (controlAt(editor.state) as { groupId: string }).groupId : '';
    expect(readControl(session.controlEntry(groupId)!.sdtPr).tag).toBe('od:repeat=x2&od:finish=t_r');
    // The caret in the first cell: the cell control (a plain-text control bound to the item's name); its title is written to the map entry.
    const table = editor.state.doc.child(5);
    const rowPos = blockStart(editor, 5) + 1 + table.child(0).nodeSize;
    await caret(editor, rowPos + 3);
    await until(() => host.querySelector('[data-testid="control-summary"]')?.textContent?.includes('cell') === true);
    expect(host.querySelector('[data-testid="cc-verdict"]')!.textContent).toContain('apples');
    const before = editor.state.doc;
    await type(input(host, 'cc-title'), 'Item name');
    const cellGroup = (table.child(1).child(0).attrs['sdt'] as SdtGroup[])[0]!.groupId;
    expect(readControl(session.controlEntry(cellGroup)!.sdtPr).title).toBe('Item name');
    expect(editor.state.doc).toBe(before);   // a map write alone: the document is not rebuilt
    await act(async () => { editor.run('history.undo'); });
    expect(readControl(session.controlEntry(cellGroup)!.sdtPr).title).toBe('Description');
    await until(() => input(host, 'cc-title').value === 'Description');
    expect(status).toEqual([]);
    root.unmount();
    editor.destroy();
  });

  test('a condition made in the panel: the parts gain the definition, the tag names it, the verdict evaluates it', async () => {
    const { session, editor, host, root } = await mount('invoice.docx', ControlPanel);
    // Wrap paragraph 2 in a rich-text control, then give it a new condition from an expression.
    const start = blockStart(editor, 2) + 1;
    await act(async () => { editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, start, start + editor.state.doc.child(2).content.size))); });
    await act(async () => { editor.run('contentControl.insert', { title: 'Spam' }); });
    await caret(editor, start + 2);
    await until(() => host.querySelector('[data-testid="cc-role-kind"]') !== null);
    const kind = host.querySelector<HTMLSelectElement>('[data-testid="cc-role-kind"]')!;
    const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value')!.set!;
    await act(async () => { setter.call(kind, 'condition'); kind.dispatchEvent(new Event('change', { bubbles: true })); });
    await type(input(host, 'cc-role-expression'), "/invoice/customer/name = 'Joe Bloggs'");
    await act(async () => { host.querySelector<HTMLButtonElement>('[data-testid="cc-role-apply"]')!.click(); });
    await until(() => host.querySelector('[data-testid="cc-role"]')?.textContent?.startsWith('condition') === true);
    const sdtPr = editor.state.doc.child(2).attrs['sdtPr'];
    const tag = readControl(sdtPr).tag;
    expect(tag).toMatch(/^od:condition=c\d+$/);
    expect(session.template.conditions!.conditions).toHaveLength(3);
    expect(session.template.xpaths!.entries).toHaveLength(7);
    expect(session.template.verdictFor(sdtPr)).toMatchObject({ role: 'condition', kept: true });
    expect(host.querySelector('[data-testid="cc-role-verdict"]')!.textContent).toContain('kept');
    expect([...session.touchedParts].filter((p) => p.startsWith('/customXml/'))).toHaveLength(2);
    // Clearing the role leaves the definition in its part.
    await act(async () => { host.querySelector<HTMLButtonElement>('[data-testid="cc-role-clear"]')!.click(); });
    await until(() => host.querySelector('[data-testid="cc-role"]')?.textContent === 'none');
    expect(session.template.conditions!.conditions).toHaveLength(3);
    root.unmount();
    editor.destroy();
  });
});

describe('the Controls tab', () => {
  test('lists every control in order with its role and verdict; a click selects it', async () => {
    const { editor, host, root } = await mount('invoice.docx', ControlsTree);
    await until(() => (host.querySelector('[data-testid="controls-list"]')?.textContent ?? '').includes('Joe Bloggs'));
    const rows = host.querySelectorAll('[data-testid="controls-list"] > li');
    expect(rows.length).toBe(6);
    expect(rows[0]!.textContent).toContain('Customer name');
    expect(rows[0]!.textContent).toContain('bound');
    expect(rows[1]!.textContent).toContain('repeat x2');
    expect(rows[1]!.textContent).toContain('3 instances');
    expect(rows[2]!.getAttribute('style')).toContain('1.1rem');   // nested in the repeat
    expect(rows[4]!.textContent).toContain('condition c5');
    expect(rows[5]!.textContent).toContain('dropped');
    await act(async () => { rows[1]!.querySelector<HTMLButtonElement>('button.link')!.click(); });
    expect(editor.state.selection).toBeInstanceOf(CellSelection);
    expect(controlAt(editor.state)?.form).toBe('row');
    await until(() => host.querySelectorAll('[data-testid="controls-list"] > li')[1]!.getAttribute('aria-current') === 'true');
    root.unmount();
    editor.destroy();
  });
});

describe('the Data tab', () => {
  test('the data part comes first with its rule; the XPath box evaluates and highlights; a value edited in place reaches the part, the verdict and the save', async () => {
    const { session, editor, host, root, status } = await mount('invoice.docx', DataPanel);
    await until(() => host.querySelector('[data-testid="data-part"]') !== null);
    expect(host.querySelector('[data-testid="data-part"]')!.textContent).toContain('invoice · /customXml/item3.xml');
    expect(host.querySelector('[data-testid="data-part"]')!.textContent).toContain('rule 1');
    expect(host.querySelector('[data-testid="od-xpaths"]')!.textContent).toContain('x2');
    expect(host.querySelector('[data-testid="od-xpaths"]')!.textContent).toContain('the row control');
    expect(host.querySelector('[data-testid="od-conditions"]')!.textContent).toContain('c5');
    await type(input(host, 'xpath-expression'), '/invoice/items/item');
    await until(() => host.querySelector('[data-testid="xpath-result"]') !== null);
    expect(host.querySelector('[data-testid="xpath-result"]')!.textContent).toContain('3 nodes');
    expect(host.querySelectorAll('[data-testid="data-part"] li.hit').length).toBe(3);
    // The customer's name edited in place.
    const name = host.querySelector<HTMLButtonElement>('[data-xpath="/invoice[1]/customer[1]/name[1]"] .node-value')!;
    expect(name.textContent).toBe('Joe Bloggs');
    await act(async () => { name.click(); });
    const box = host.querySelector<HTMLInputElement>('[data-xpath="/invoice[1]/customer[1]/name[1]"] input')!;
    await type(box, 'Jane Doe');
    expect(session.template.dataPart.part!.selectSingleNode('/invoice/customer/name')!.text).toBe('Jane Doe');
    expect(session.template.verdictFor(editor.state.doc.child(3).attrs['sdtPr'])).toMatchObject({ value: 'Jane Doe' });
    expect([...session.touchedParts]).toContain('/customXml/item3.xml');
    expect(status[0]).toContain('no undo');
    // The save carries the part. (No self-check here: jsdom's XML serializer writes the regenerated
    // relationships parts with a duplicated xmlns, which its own parser then rejects; Chromium's does
    // not, and the Playwright tests reload the saves.)
    const saved = await session.save({ selfCheck: false });
    expect(saved.changed).toBe(true);
    const zip = new ZipPartStore(saved.bytes);
    expect(new TextDecoder().decode(zip.loadSync('customXml/item3.xml'))).toContain('<name>Jane Doe</name>');
    expect(readControlEntries(session.store).size).toBe(3);
    root.unmount();
    editor.destroy();
  });
});

describe('verdict badges and the Preview tab', () => {
  test('the page carries the verdicts once the template is loaded; the preview shows the bound values with the badges', async () => {
    const { session, editor, root } = await mount('invoice.docx', ControlPanel);
    await until(() => editor.container.querySelector('.sdt-block[data-verdict="dropped"]') !== null);
    const blocks = editor.container.querySelectorAll<HTMLElement>('.sdt-block');
    expect(blocks[1]!.dataset['verdict']).toBe('kept');
    expect(blocks[2]!.dataset['verdict']).toBe('dropped');
    expect(blocks[2]!.dataset['verdictKept']).toBe('false');
    expect(editor.container.querySelector('td[data-sdt-verdict="3 instances"]')).not.toBeNull();
    // The data part edited: the badges follow the template view's notification.
    session.template.dataPart.part!.selectSingleNode('/invoice/paid')?.delete();
    session.template.invalidate();
    await until(() => editor.container.querySelectorAll('.sdt-block[data-verdict]').length >= 2);
    root.unmount();
    // The Preview tab.
    const previewHost = document.createElement('div');
    document.body.appendChild(previewHost);
    const previewRoot = createRoot(previewHost);
    const ctx: TabContext = { editor, active: editor, session, perspective: { id: 'template', panel: { open: true, tab: 'preview' } }, consoleScript: null, onConsoleScript: () => {}, onSaveMacro: () => {}, onStatus: () => {}, onTab: () => {} };
    await act(async () => { previewRoot.render(<PreviewPanel {...ctx} />); });
    // The preview clones the package by a save and a reload, which jsdom cannot do (its serializer duplicates
    // xmlns on the regenerated relationships parts, ED-003 section 11.2 item 18): here the tab reports that; the
    // story itself is the Playwright test's, and the model test covers the clone in Node.
    await until(() => previewHost.querySelector('[data-testid="preview-story"] .ProseMirror, [data-testid="preview-error"]') !== null, 8000);
    const story = previewHost.querySelector('[data-testid="preview-story"] .ProseMirror');
    if (story) {
      expect(story.getAttribute('contenteditable')).toBe('false');
      expect(story.querySelectorAll('.sdt-block')[0]!.textContent).toBe('Joe Bloggs');
      expect(story.querySelector('td[data-sdt-verdict="3 instances"]')).not.toBeNull();
      expect(story.querySelector('.sdt-block[data-verdict="dropped"]')).not.toBeNull();
      expect(previewHost.querySelector('[data-testid="preview-summary"]')!.textContent).toContain('3 bound controls');
    } else {
      expect(previewHost.querySelector('[data-testid="preview-error"]')!.textContent).toMatch(/parsererror|unmarshalled|Malformed XML/);
    }
    expect(session.exportTree().changed).toBe(false);
    previewRoot.unmount();
    editor.destroy();
  });
});

describe('the Data tab from the keyboard (ED-003 section 8.6\'s checklist)', () => {
  test('the arrow keys walk the tree, expanding and collapsing; the bind button binds the control at the caret', async () => {
    const { session, editor, host, root, status } = await mount('invoice.docx', DataPanel);
    await until(() => host.querySelector('[data-testid="data-part"]') !== null);
    const tree = host.querySelector<HTMLElement>('[data-testid="data-part"] [role="tree"]')!;
    expect(tree.getAttribute('aria-label')).toContain('invoice');
    const name = (xpath: string): HTMLElement => host.querySelector<HTMLElement>(`[data-xpath="${xpath}"] > .tree-row > .node-name`)!;
    const item = (xpath: string): HTMLElement => host.querySelector<HTMLElement>(`[data-xpath="${xpath}"]`)!;
    const key = (k: string): Promise<void> => act(async () => { (document.activeElement as HTMLElement).dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true })); });
    // Only the root is a tab stop; a leaf is named with its value.
    expect(name('/invoice[1]').tabIndex).toBe(0);
    expect(name('/invoice[1]/customer[1]').tabIndex).toBe(-1);
    expect(item('/invoice[1]/customer[1]/name[1]').getAttribute('aria-label')).toBe('name: Joe Bloggs');
    await act(async () => { name('/invoice[1]').focus(); });
    await key('ArrowDown');
    expect(document.activeElement).toBe(name('/invoice[1]/customer[1]'));
    expect(item('/invoice[1]/customer[1]').getAttribute('aria-selected')).toBe('true');
    await key('ArrowLeft');   // collapses the expanded customer
    expect(item('/invoice[1]/customer[1]').getAttribute('aria-expanded')).toBe('false');
    expect(host.querySelector('[data-xpath="/invoice[1]/customer[1]/name[1]"]')).toBeNull();
    await key('ArrowRight');   // expands it again
    expect(item('/invoice[1]/customer[1]').getAttribute('aria-expanded')).toBe('true');
    await key('ArrowRight');   // steps into it
    expect(document.activeElement).toBe(name('/invoice[1]/customer[1]/name[1]'));
    await key('ArrowLeft');    // a leaf: out to the parent
    expect(document.activeElement).toBe(name('/invoice[1]/customer[1]'));
    await key('End');
    expect(document.activeElement).toBe(name('/invoice[1]/misc[1]/wantspam[1]'));
    await key('Home');
    expect(document.activeElement).toBe(name('/invoice[1]'));
    // Bind from the keyboard: with no control at the caret a plain text control is made there and bound (ED-005 section
    // 12.56); with one, the node is bound as a drop would bind it.
    const bind = host.querySelector<HTMLButtonElement>('[aria-label="Bind to /invoice[1]/items[1]/item[1]/price[1]: the control at the caret, or a new plain text control"]')!;
    await act(async () => { bind.click(); });
    expect(status.at(-1)).toBe('Inserted a plain text control bound to /invoice[1]/items[1]/item[1]/price[1]');
    await caret(editor, blockStart(editor, 3) + 2);   // the "Customer name" control
    await act(async () => { bind.click(); });
    expect(status.at(-1)).toBe('Bound to /invoice[1]/items[1]/item[1]/price[1]');
    expect(readControl(editor.state.doc.child(3).attrs['sdtPr']).binding?.xpath).toBe('/invoice[1]/items[1]/item[1]/price[1]');
    expect(session.template.verdictFor(editor.state.doc.child(3).attrs['sdtPr'])).toMatchObject({ value: '$20' });
    // Inside a rich text control (a condition): a plain text control is made in it and bound, as Word's pane nests one
    // (ED-005 section 12.57); the condition keeps its own properties.
    let condition = -1;
    editor.state.doc.forEach((b, _o, i) => { if (condition < 0 && b.type.name === 'sdtBlock' && readControl(b.attrs['sdtPr']).kind === 'RichText') condition = i; });
    const conditionSdtPr = editor.state.doc.child(condition).attrs['sdtPr'];
    await caret(editor, blockStart(editor, condition) + 2);
    await act(async () => { bind.click(); });
    expect(status.at(-1)).toBe('Inserted a plain text control bound to /invoice[1]/items[1]/item[1]/price[1]');
    const nested: string[] = [];
    editor.state.doc.child(condition).descendants((n) => { if (n.type.name === 'sdtInline') nested.push(readControl(n.attrs['sdtPr']).binding?.xpath ?? ''); return true; });
    expect(nested).toEqual(['/invoice[1]/items[1]/item[1]/price[1]']);
    expect(editor.state.doc.child(condition).attrs['sdtPr']).toEqual(conditionSdtPr);
    root.unmount();
    editor.destroy();
  });
});

describe('the Data tab as an editor', () => {
  test('attributes added, edited and removed; an element renamed, given a sibling, duplicated and removed; the verdicts follow', async () => {
    const { session, editor, host, root, status } = await mount('invoice.docx', DataPanel);
    await until(() => host.querySelector('[data-testid="data-part"]') !== null);
    const row = (xpath: string): HTMLElement => host.querySelector<HTMLElement>(`[data-xpath="${xpath}"] > .tree-row`)!;
    const part = session.template.dataPart.part!;
    // An attribute on the customer.
    await act(async () => { row('/invoice[1]/customer[1]').querySelector<HTMLButtonElement>('[aria-label="Add an attribute to /invoice[1]/customer[1]"]')!.click(); });
    await fill(input(host, 'node-form-name'), 'id');
    await type(input(host, 'node-form-value'), '42');
    expect(part.selectSingleNode('/invoice/customer')!.attributes.map((a) => [a.baseName, a.nodeValue])).toEqual([['id', '42']]);
    expect(row('/invoice[1]/customer[1]').textContent).toContain('@id="42"');
    // Edited, then removed.
    await act(async () => { row('/invoice[1]/customer[1]').querySelector<HTMLButtonElement>('button.attr')!.click(); });
    await type(input(host, 'node-form-value'), '43');
    expect(part.selectSingleNode('/invoice/customer/@id')!.nodeValue).toBe('43');
    await act(async () => { row('/invoice[1]/customer[1]').querySelector<HTMLButtonElement>('[aria-label="Remove @id from /invoice[1]/customer[1]"]')!.click(); });
    expect(part.selectSingleNode('/invoice/customer')!.attributes).toHaveLength(0);
    // Renamed: the binding to /invoice/customer/name now selects nothing, and the verdict says so.
    await act(async () => { row('/invoice[1]/customer[1]/name[1]').querySelector<HTMLButtonElement>('[aria-label="Rename /invoice[1]/customer[1]/name[1]"]')!.click(); });
    await type(input(host, 'node-form-name'), 'fullName');
    expect(part.selectSingleNode('/invoice/customer/fullName')!.text).toBe('Joe Bloggs');
    expect(session.template.verdictFor(editor.state.doc.child(3).attrs['sdtPr'])).toMatchObject({ selectsNothing: true });
    await act(async () => { row('/invoice[1]/customer[1]/fullName[1]').querySelector<HTMLButtonElement>('[aria-label="Rename /invoice[1]/customer[1]/fullName[1]"]')!.click(); });
    await type(input(host, 'node-form-name'), 'name');
    expect(session.template.verdictFor(editor.state.doc.child(3).attrs['sdtPr'])).toMatchObject({ value: 'Joe Bloggs' });
    // A fourth item by duplication: the repeat's count follows; a sibling with a value; the last item removed.
    const repeatRow = (editor.state.doc.child(5).child(1).attrs['sdt'] as SdtGroup[])[0]!.groupId;
    expect(session.template.verdictFor(session.controlEntry(repeatRow)!.sdtPr)).toMatchObject({ count: 3 });
    await act(async () => { row('/invoice[1]/items[1]/item[1]').querySelector<HTMLButtonElement>('[aria-label="Duplicate /invoice[1]/items[1]/item[1]"]')!.click(); });
    expect(part.selectNodes('/invoice/items/item')).toHaveLength(4);
    expect(part.selectSingleNode('/invoice/items/item[2]/name')!.text).toBe('apples');
    expect(session.template.verdictFor(session.controlEntry(repeatRow)!.sdtPr)).toMatchObject({ count: 4 });
    await act(async () => { row('/invoice[1]/items[1]/item[4]').querySelector<HTMLButtonElement>('[aria-label="Add a sibling after /invoice[1]/items[1]/item[4]"]')!.click(); });
    expect(input(host, 'node-form-name').value).toBe('item');   // the sibling's name defaults to this element's
    await type(input(host, 'node-form-value'), 'x');
    expect(part.selectNodes('/invoice/items/item')).toHaveLength(5);
    expect(part.selectSingleNode('/invoice/items/item[5]')!.text).toBe('x');
    await act(async () => { row('/invoice[1]/items[1]/item[5]').querySelector<HTMLButtonElement>('[aria-label="Remove /invoice[1]/items[1]/item[5]"]')!.click(); });
    expect(part.selectNodes('/invoice/items/item')).toHaveLength(4);
    // A first child, then moved: by the buttons, then by a drop (the upper quarter of a row: before it).
    await act(async () => { row('/invoice[1]/items[1]').querySelector<HTMLButtonElement>('[aria-label="Add a child to /invoice[1]/items[1]"]')!.click(); });
    const position = host.querySelector<HTMLSelectElement>('[data-testid="node-form-position"]')!;
    await act(async () => { Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value')!.set!.call(position, 'first'); position.dispatchEvent(new Event('change', { bubbles: true })); });
    await fill(input(host, 'node-form-name'), 'note');
    await type(input(host, 'node-form-value'), 'first');
    expect(part.selectSingleNode('/invoice/items/*[1]')!.baseName).toBe('note');
    await act(async () => { row('/invoice[1]/items[1]/note[1]').querySelector<HTMLButtonElement>('[aria-label="Move /invoice[1]/items[1]/note[1] down"]')!.click(); });
    expect(part.selectSingleNode('/invoice/items/*[2]')!.baseName).toBe('note');
    expect(row('/invoice[1]/items[1]/note[1]').querySelector<HTMLButtonElement>('[aria-label="Move /invoice[1]/items[1]/note[1] up"]')!.disabled).toBe(false);
    const drop = (xpathTo: string, y: number): void => {
      const target = row(xpathTo);
      const rect = { top: 100, height: 20 };
      target.getBoundingClientRect = () => ({ ...rect, bottom: 120, left: 0, right: 100, width: 100, x: 0, y: 100, toJSON: () => ({}) });
      const data: Record<string, string> = { [MOVE_MIME]: JSON.stringify({ storeItemID: part.id, xpath: '/invoice[1]/items[1]/note[1]', prefixMappings: '' }) };
      const dataTransfer = { types: Object.keys(data), getData: (t: string) => data[t] ?? '', dropEffect: 'none' };
      const over = new Event('dragover', { bubbles: true, cancelable: true });
      Object.defineProperties(over, { dataTransfer: { value: dataTransfer }, clientY: { value: rect.top + y } });
      target.dispatchEvent(over);
      const ev = new Event('drop', { bubbles: true, cancelable: true });
      Object.defineProperties(ev, { dataTransfer: { value: dataTransfer }, clientY: { value: rect.top + y } });
      target.dispatchEvent(ev);
    };
    await act(async () => { drop('/invoice[1]/items[1]/item[4]', 18); });   // after the last item
    expect(part.selectNodes('/invoice/items/*').map((n) => n.baseName)).toEqual(['item', 'item', 'item', 'item', 'note']);
    await act(async () => { drop('/invoice[1]/items[1]/item[1]', 10); });   // into the first item, as its last child
    expect(part.selectSingleNode('/invoice/items/item[1]/note')).toBeDefined();
    expect(part.selectNodes('/invoice/items/*')).toHaveLength(4);
    expect([...session.touchedParts]).toContain('/customXml/item3.xml');
    expect(status.every((s) => s.includes('no undo'))).toBe(true);
    // The OpenDoPE parts are not editable here.
    expect(host.querySelector('[data-testid="od-xpaths"] .node-actions')).toBeNull();
    root.unmount();
    editor.destroy();
  });
});
