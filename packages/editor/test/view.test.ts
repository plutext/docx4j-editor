// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 5.6 step C1: the view mounts over a session's store in jsdom, a keystroke and a
// command reach the store, and the export sees them.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { TextSelection } from 'prosemirror-state';
import { EditorDocument, readControl } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

async function open(name: string): Promise<{ session: EditorDocument; editor: DocxEditorView; original: Uint8Array }> {
  const original = new Uint8Array(await readFile(join(fixtures, name)));
  const session = await EditorDocument.open(original);
  const container = document.createElement('div');
  document.body.appendChild(container);
  const editor = new DocxEditorView(container, session);
  return { session, editor, original };
}

describe('DocxEditorView', () => {
  test('mounts the store, and a command through the view reaches the store and the export', async () => {
    const { session, editor } = await open('tracked-changes.docx');
    expect(editor.state.doc.eq(session.doc)).toBe(true);
    expect(editor.container.querySelector('.ProseMirror')).not.toBeNull();
    // "Some text." is block 4; select "text" and bold it.
    let pos = 0;
    for (let k = 0; k < 4; k++) pos += editor.state.doc.child(k).nodeSize;
    const from = pos + 1 + 5;
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, from + 4)));
    expect(editor.run('font.bold', { value: true })).toBe(true);
    expect(editor.formatting().font.bold).toBe(true);
    // The store followed the view; the export sees the change; undo takes it back through the store.
    expect(session.doc.child(4).child(1).marks.some((m) => m.type.name === 'bold')).toBe(true);
    expect(session.exportTree().changed).toBe(true);
    expect(session.undoManager.canUndo()).toBe(true);
    expect(editor.run('history.undo')).toBe(true);
    expect(session.exportTree().changed).toBe(false);
    expect(editor.state.doc.eq(session.imported)).toBe(true);
    editor.destroy();
  });

  test('the final section is on the root node, and a section command writes the package map', async () => {
    const { session, editor } = await open('sections-multiple.docx');
    expect(editor.state.doc.attrs['finalSectPr']).not.toBeNull();
    const last = editor.state.doc.childCount - 1;
    let pos = 0;
    for (let k = 0; k < last; k++) pos += editor.state.doc.child(k).nodeSize;
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos + 1)));
    expect(editor.run('section.columns', { count: 2 })).toBe(true);
    expect((session.store.package.get('finalSectPr') as { cols?: { num?: number } }).cols?.num).toBe(2);
    expect(session.exportTree().changed).toBe(true);
    editor.run('history.undo');
    expect((session.store.package.get('finalSectPr') as { cols?: { num?: number } }).cols?.num).not.toBe(2);
    expect(editor.state.doc.attrs['finalSectPr']).toEqual(session.store.package.get('finalSectPr'));
    editor.destroy();
  });
});

describe('the stylesheet follows the styles in use', () => {
  test('a style applied for the first time gets its rule at once (a new document, then a fixture)', async () => {
    const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
    const session = await EditorDocument.open(await WordprocessingMLPackage.createPackage({ pageSize: 'A4' }));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const sheet = (): string => document.head.querySelector<HTMLStyleElement>('style[data-docx4j-editor="document"]')!.textContent ?? '';
    expect(sheet()).not.toContain('.s-Heading1');
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1)).insertText('A title'));
    expect(editor.run('paragraph.style', { value: 'Heading1' })).toBe(true);
    expect(editor.state.doc.child(0).attrs['pStyle']).toBe('Heading1');
    expect(sheet()).toContain('.s-Heading1');
    expect(sheet()).toMatch(/\.s-Heading1 \{[^}]*font-size: 14pt/);
    editor.destroy();
    // A fixture that never used Heading2.
    const { session: other, editor: e2 } = await open('tracked-changes.docx');
    expect(sheet()).not.toContain('.s-Heading2');
    e2.view.dispatch(e2.state.tr.setSelection(TextSelection.create(e2.state.doc, 1)));
    expect(e2.run('paragraph.style', { value: 'Heading2' })).toBe(true);
    expect(sheet()).toContain('.s-Heading2');
    expect(other.exportTree().changed).toBe(true);
    e2.destroy();
  });
});

describe('content controls in the view (ED-003 section 4.2)', () => {
  test('a group\'s properties go to the controls map in the sync plugin\'s transaction: one undo step, the document not rebuilt; wrapped rows and their entry likewise', async () => {
    const { session, editor } = await open('invoice.docx');
    const table = editor.state.doc.child(5);
    let pos = 0;
    for (let k = 0; k < 5; k++) pos += editor.state.doc.child(k).nodeSize;
    const rowPos = pos + 1 + table.child(0).nodeSize;
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, rowPos + 3)));   // the first cell of the repeat row
    const cellGroup = (table.child(1).child(0).attrs['sdt'] as { groupId: string }[])[0]!.groupId;
    const before = editor.state.doc;
    let notified = 0;
    const off = editor.subscribe(() => { notified++; });
    expect(editor.run('contentControl.properties', { title: 'Item name' })).toBe(true);
    expect(readControl(session.controlEntry(cellGroup)!.sdtPr).title).toBe('Item name');
    expect(editor.state.doc).toBe(before);   // no rebuild from the fragment: the observer fired inside the mutex
    expect(session.undoManager.canUndo()).toBe(true);
    expect(session.exportTree().changed).toBe(true);
    // Undo reverts the map; the view hears of it (an empty transaction) so the panels and labels refresh.
    const seen = notified;
    expect(editor.run('history.undo')).toBe(true);
    expect(readControl(session.controlEntry(cellGroup)!.sdtPr).title).toBe('Description');
    expect(notified).toBeGreaterThan(seen);
    expect(session.exportTree().changed).toBe(false);
    off();
    // The whole repeat row selected: the properties go to the row group, and the row's label follows at once.
    expect(editor.run('contentControl.select', { groupId: (table.child(1).attrs['sdt'] as { groupId: string }[])[0]!.groupId })).toBe(true);
    expect(editor.run('contentControl.properties', { title: 'Items' })).toBe(true);
    expect(editor.container.querySelector('tr[data-sdt="Items"]')).not.toBeNull();
    expect(editor.run('history.undo')).toBe(true);
    expect(editor.container.querySelector('tr[data-sdt="Repeat"]')).not.toBeNull();
    // Rows wrapped: the row attributes and the new entry in one step.
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos + 4)));   // the header row
    expect(editor.run('contentControl.wrapRows', { title: 'Header' })).toBe(true);
    const stack = editor.state.doc.child(5).child(0).attrs['sdt'] as { groupId: string }[];
    expect(stack).toHaveLength(1);
    expect(readControl(session.controlEntry(stack[0]!.groupId)!.sdtPr).title).toBe('Header');
    expect(editor.container.querySelector('tr[data-sdt="Header"]')).not.toBeNull();
    expect(editor.run('history.undo')).toBe(true);
    expect(editor.state.doc.child(5).child(0).attrs['sdt']).toBeNull();
    expect(session.controlEntry(stack[0]!.groupId)).toBeUndefined();
    expect(session.exportTree().changed).toBe(false);
    editor.destroy();
  });

  test('block and run controls carry their label, role and kind for the stylesheet', async () => {
    const { editor } = await open('invoice.docx');
    const block = editor.container.querySelector<HTMLElement>('.sdt-block')!;
    expect(block.dataset['label']).toBe('Customer name');
    expect(block.dataset['role']).toBe('bound');
    expect(block.dataset['kind']).toBe('PlainText');
    const row = editor.container.querySelector<HTMLElement>('tr[data-sdt]')!;
    expect(row.dataset['sdt']).toBe('Repeat');   // the alias; the tag when there is none
    expect(row.dataset['sdtRole']).toBe('repeat');
    expect(row.title).toContain('tag: od:repeat=x2');
    const cell = editor.container.querySelector<HTMLElement>('td[data-sdt]')!;
    expect(cell.dataset['sdtRole']).toBe('bound');
    const condition = editor.container.querySelectorAll<HTMLElement>('.sdt-block')[2]!;
    expect(condition.dataset['role']).toBe('condition');
    editor.destroy();
  });
});

describe('tables in the view (ED-003 section 3.6)', () => {
  test('a table mounts as a table with its merges, and a table command through the view exports', async () => {
    const { session, editor } = await open('tables.docx');
    const tables = editor.container.querySelectorAll('.ProseMirror[aria-label="Document"] > .tableWrapper > table');
    expect(tables.length).toBeGreaterThan(10);
    expect(tables[1]!.querySelectorAll('td[rowspan="2"]').length).toBe(1);
    expect(tables[3]!.querySelectorAll('td[colspan="2"]').length).toBe(1);
    expect(tables[0]!.getAttribute('style')).toContain('--tbl-top');
    // The caret in the first cell of the first table, then a row below.
    let pos = editor.state.doc.child(0).nodeSize;
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos + 4)));
    expect(editor.run('table.addRows', { location: 'After' })).toBe(true);
    expect(editor.state.doc.child(1).childCount).toBe(3);
    expect(session.exportTree().changed).toBe(true);
    expect(editor.run('history.undo')).toBe(true);
    expect(editor.state.doc.child(1).childCount).toBe(2);
    editor.destroy();
  });
});

describe('undo survives the body view being made again (ED-004 section 10.8 item 1; UPSTREAM-4)', () => {
  test('a second view over the same session still captures and undoes typing', async () => {
    const { session, editor } = await open('tracked-changes.docx');
    editor.destroy();
    const container = document.createElement('div');
    document.body.appendChild(container);
    const again = new DocxEditorView(container, session);
    again.view.dispatch(again.state.tr.insertText('Typed later. ', 2));
    expect(session.undoManager.canUndo()).toBe(true);
    expect(again.run('history.undo')).toBe(true);
    expect(again.state.doc.textContent).not.toContain('Typed later. ');
    again.destroy();
  });
});
