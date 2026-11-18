// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 4.10 items 1 and 2, the controls half: the `controls` map at import, the
// untouched round trip with every part the template view reads, and every command of section
// 4.3 through the headless runner with the export reloaded and compared: a wrapped block, a
// wrapped run split at its edges, rows wrapped and unwrapped, nested inside a repeat, delete
// keeping and dropping, a binding written and read back through the engine's `XmlMapping`, a
// condition and a repeat tag with their entries, undo per command.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage, Body, rowsOf, typeNameOf, textOf as textOfElement, type Element } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { TextSelection } from 'prosemirror-state';
import {
  EditorDocument, runCommand, stateFor, commandNamed, listControls, readControl, readControlEntry, readControlEntries, controlWritesOf, sdtTag, sdtTitle, sdtId, sdtBinding,
  CellSelection, TableMap, PLACEHOLDER_TEXT, PLACEHOLDER_STYLE, applyBoundRefresh, plainTextControlRefusal, controlToBind, dataBindRefusal, insideControlHoldingNoControls, type SdtGroup,
} from '../src/index.mjs';
import { withSdtItem } from '../src/sdtpr.mjs';
import { isXhtmlControl, xhtmlBindingOf, xhtmlControlsIn, xhtmlKindOf, readTag, relativeToRepeat } from '../src/index.mjs';
import type { Node as PMNode } from 'prosemirror-model';
import { fixture, comparePackages, textOf, mainXml } from './helpers.mjs';

async function reload(bytes: Uint8Array): Promise<{ pkg: WordprocessingMLPackage; body: wml.Body; api: Body }> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  const main = pkg.getMainDocumentPart();
  const body = (await main.readContents()).body!;
  return { pkg, body, api: new Body(main, body, 'body', pkg) };
}

/**
 * Save; only the main part, app.xml and the parts named may differ, and custom XML parts may be added, as may the
 * glossary (the default placeholder entry a control made here names, written where the file has none: ED-008 section
 * 8.8 item 14, `placeholders.test.mts`).
 */
async function saveAndReload(session: EditorDocument, original: Uint8Array, alsoChanged: string[] = []) {
  const saved = await session.save();
  expect(saved.changed).toBe(true);
  const diffs = comparePackages(original, saved.bytes, [session.main.partName.storeName, 'docProps/app.xml', ...alsoChanged]).filter((d) => !(d.reason === 'added' && (d.name.startsWith('customXml/') || d.name.startsWith('word/glossary/'))));
  expect(diffs).toEqual([]);
  return { bytes: saved.bytes, ...(await reload(saved.bytes)) };
}

function blockStart(session: EditorDocument, i: number): number {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += session.doc.child(k).nodeSize;
  return pos;
}

function cellStart(session: EditorDocument, block: number, r: number, c: number): number {
  const table = session.doc.child(block);
  const map = TableMap.get(table);
  return blockStart(session, block) + 1 + map.map[r * map.width + c]! + 2;
}

describe('projection', () => {
  test('invoice.docx: the row group has a controls entry with its properties; listControls gives every control in order with its depth', async () => {
    const session = await EditorDocument.open(await fixture('invoice.docx'));
    const entries = readControlEntries(session.store);
    expect(entries.size).toBe(3);   // the repeat row and its two cell controls
    const rowGroup = (session.doc.child(5).child(1).attrs['sdt'] as SdtGroup[])[0]!.groupId;
    expect(sdtTag(readControlEntry(session.store, rowGroup)!.sdtPr)).toBe('od:repeat=x2&od:finish=t_r');
    expect(readControlEntry(session.store, rowGroup)!.form).toBe('row');
    const controls = listControls(session.doc, session.controlEntry);
    expect(controls.map((c) => [c.form, c.depth, c.reading.kind, c.tag.role])).toEqual([
      ['block', 0, 'PlainText', 'none'], ['row', 0, 'RichText', 'repeat'], ['cell', 1, 'PlainText', 'none'], ['cell', 1, 'PlainText', 'none'], ['block', 0, 'RichText', 'condition'], ['block', 0, 'RichText', 'condition'],
    ]);
    expect(controls[1]).toMatchObject({ groupId: rowGroup, rows: [1, 1], tablePos: blockStart(session, 5) });
    expect(controls[0]!.reading).toMatchObject({ title: 'Customer name', id: 1418037941, appearance: 'BoundingBox', cannotDelete: false, binding: { xpath: '/invoice[1]/customer[1]/name[1]', w15: false } });
    // Word's repeating sections are listed by their kind, the item inside one a depth down.
    const w15 = await EditorDocument.open(await fixture('invoice2013.docx'));
    const sections = listControls(w15.doc, w15.controlEntry);
    const section = sections.findIndex((c) => c.reading.kind === 'RepeatingSection');
    expect(section).toBeGreaterThanOrEqual(0);
    expect(sections[section + 1]).toMatchObject({ depth: sections[section]!.depth + 1 });
    expect(sections[section + 1]!.reading.kind).toBe('RepeatingSectionItem');
  });

  test.each(['invoice.docx', 'invoice2013.docx', 'opendope/binding-simple.docx', 'opendope/conventions.docx', 'opendope/non-element-bind.docx', 'opendope/hyperlink-binding-test.docx', 'opendope/picture.docx', 'opendope/invoice_Saxon_XPath2.docx'])(
    '%s: the untouched round trip holds after the template view read every part and every verdict', async (name) => {
      const original = await fixture(name);
      const session = await EditorDocument.open(original);
      await session.template.load();
      const t = session.template;
      void t.xpaths; void t.conditions; void t.components; void t.questions; void t.dataPart; void t.problems; void t.isTemplate;
      for (const c of listControls(session.doc, session.controlEntry)) t.verdictFor(c.sdtPr);
      const saved = await session.save();
      expect(saved.changed).toBe(false);
      expect(comparePackages(original, saved.bytes)).toEqual([]);
    });
});

describe('block and run controls', () => {
  test('a run control around a selection splits the runs at its edges; delete keeping the content puts them back', async () => {
    const original = await fixture('HelloWordOnline.docx');
    const session = await EditorDocument.open(original);
    const para = session.doc.child(0);
    expect(para.type.name).toBe('paragraph');
    const text = textOf(para);
    // The paragraph opens with an invisible page-break atom: the first text node is where the characters start.
    let firstText = -1;
    para.forEach((n, offset) => { if (firstText < 0 && n.isText) firstText = offset; });
    const start = blockStart(session, 0) + 1 + firstText;
    expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText', tag: 'od:xpath=x1', title: 'Greeting' }, { from: start + 1, to: start + 4 }).ok).toBe(true);
    let inline = session.doc.child(0).child(0);
    session.doc.child(0).forEach((n) => { if (n.type.name === 'sdtInline') inline = n; });
    expect(inline.type.name).toBe('sdtInline');
    expect(textOf(inline)).toBe(text.slice(1, 4));
    expect(readControl(inline.attrs['sdtPr'])).toMatchObject({ kind: 'PlainText', tag: 'od:xpath=x1', title: 'Greeting' });
    expect(sdtId(inline.attrs['sdtPr'])).toBeGreaterThan(0);
    const { api, body } = await saveAndReload(session, original);
    const control = api.contentControls[0]!;
    expect(control.form).toBe('Run');
    expect(control.type).toBe('PlainText');
    expect(control.tag).toBe('od:xpath=x1');
    expect(control.title).toBe('Greeting');
    expect(control.text).toBe(text.slice(1, 4));
    expect(textOfElement(body.content![0]!)).toBe(text);
    // Unwrap: the text is back in ordinary runs.
    expect(runCommand(session, 'contentControl.delete', { keepContent: true }, { from: start + 2 }).ok).toBe(true);
    expect(session.doc.child(0).childCount).toBeGreaterThan(0);
    expect(session.doc.descendants((n) => n.type.name !== 'sdtInline')).toBeUndefined();
    let inlines = 0;
    session.doc.descendants((n) => { if (n.type.name === 'sdtInline') inlines++; return true; });
    expect(inlines).toBe(0);
    expect(textOf(session.doc.child(0))).toBe(text);
    const after = await saveAndReload(session, original);
    expect(after.api.contentControls).toHaveLength(0);
  });

  test('at a caret: a control showing the placeholder; delete without keeping removes it', async () => {
    const original = await fixture('HelloWordOnline.docx');
    const session = await EditorDocument.open(original);
    const text = textOf(session.doc.child(0));
    const end = blockStart(session, 0) + 1 + session.doc.child(0).content.size;
    expect(runCommand(session, 'contentControl.insert', undefined, { from: end }).ok).toBe(true);
    const inline = session.doc.child(0).lastChild!;
    expect(inline.type.name).toBe('sdtInline');
    expect(textOf(inline)).toBe(PLACEHOLDER_TEXT);
    expect(inline.firstChild!.marks.some((m) => m.type.name === 'runStyle' && m.attrs['id'] === PLACEHOLDER_STYLE)).toBe(true);
    expect(readControl(inline.attrs['sdtPr'])).toMatchObject({ kind: 'RichText', showingPlaceholder: true });
    // The placeholder's style, which the file lacked, is added to the styles part (UPSTREAM-19, ED-005 section 12.59).
    const { api, pkg } = await saveAndReload(session, original, ['word/styles.xml']);
    expect((await pkg.getMainDocumentPart().styleDefinitionsPart!.readContents()).style?.some((st) => st.styleId === PLACEHOLDER_STYLE)).toBe(true);
    expect(api.contentControls[0]!.isShowingPlaceholder).toBe(true);
    expect(api.contentControls[0]!.text).toBe('');   // a placeholder reads as no text since core-ts 0.6.0 (CR-012 section 18, Office JS's reading)
    expect(runCommand(session, 'contentControl.delete', undefined, { from: end + 2 }).ok).toBe(true);
    expect(textOf(session.doc.child(0))).toBe(text);
    expect(runCommand(session, 'history.undo').ok).toBe(true);
    expect(textOf(session.doc.child(0))).toBe(text + PLACEHOLDER_TEXT);
  });

  test('a whole paragraph selected: a block control; a repeating section wraps its item; nesting inside a block control', async () => {
    const original = await fixture('HelloWordOnline.docx');
    const session = await EditorDocument.open(original);
    const para = session.doc.child(0);
    const start = blockStart(session, 0) + 1;
    expect(runCommand(session, 'contentControl.insert', { tag: 'od:condition=c1' }, { from: start, to: start + para.content.size }).ok).toBe(true);
    const block = session.doc.child(0);
    expect(block.type.name).toBe('sdtBlock');
    expect(block.child(0).attrs['sourceId']).toBe(para.attrs['sourceId']);   // the paragraph itself is unchanged
    // Inside it, a repeating section around the same paragraph: section > item > paragraph.
    expect(runCommand(session, 'contentControl.insert', { kind: 'RepeatingSection', title: 'Items' }, { from: start + 1, to: start + 1 + para.content.size }).ok).toBe(true);
    const outer = session.doc.child(0);
    expect(outer.child(0).type.name).toBe('sdtBlock');
    expect(readControl(outer.child(0).attrs['sdtPr'])).toMatchObject({ kind: 'RepeatingSection', title: 'Items' });
    expect(readControl(outer.child(0).child(0).attrs['sdtPr']).kind).toBe('RepeatingSectionItem');
    expect(outer.child(0).child(0).child(0).type.name).toBe('paragraph');
    const { api } = await saveAndReload(session, original);
    expect(api.contentControls.map((c) => [c.form, c.type, c.tag])).toEqual([['Block', 'RichText', 'od:condition=c1'], ['Block', 'RepeatingSection', ''], ['Block', 'RepeatingSectionItem', '']]);
    expect(api.contentControls[1]!.repeatingSectionContentControl!.items).toHaveLength(1);
    // The ids are distinct and reload as written.
    const ids = api.contentControls.map((c) => c.id);
    expect(new Set(ids).size).toBe(3);
    expect(ids).toEqual(listControls(session.doc, session.controlEntry).map((c) => c.reading.id));
  });

  test('properties: tag, title, lock, appearance, colour and temporary, written in the engine\'s shapes and read back through it', async () => {
    const original = await fixture('invoice.docx');
    const session = await EditorDocument.open(original);
    const first = listControls(session.doc, session.controlEntry)[0]!;   // the "Customer name" control
    const pos = first.pos! + 2;   // inside its paragraph
    const block = session.doc.resolve(first.pos! + 1).index(0);
    expect(runCommand(session, 'contentControl.properties', { title: 'Client', tag: 'od:xpath=x1&client=1', cannotDelete: true, cannotEdit: true, appearance: 'Tags', color: '#ff0000', removeWhenEdited: true }, { from: pos }).ok).toBe(true);
    const reading = readControl(session.doc.child(block).attrs['sdtPr']);
    expect(reading).toMatchObject({ title: 'Client', tag: 'od:xpath=x1&client=1', cannotDelete: true, cannotEdit: true, appearance: 'Tags', color: '#FF0000', removeWhenEdited: true, id: 1418037941 });
    const { api } = await saveAndReload(session, original);
    const c = api.contentControls[0]!;
    expect([c.title, c.tag, c.cannotDelete, c.cannotEdit, c.appearance, c.color, c.removeWhenEdited, c.id]).toEqual(['Client', 'od:xpath=x1&client=1', true, true, 'Tags', '#FF0000', true, 1418037941]);
    expect(runCommand(session, 'contentControl.properties', { cannotDelete: false, appearance: 'BoundingBox', color: '' }, { from: pos }).ok).toBe(true);
    expect(readControl(session.doc.child(block).attrs['sdtPr'])).toMatchObject({ cannotDelete: false, cannotEdit: true, appearance: 'BoundingBox', color: '' });
    expect(runCommand(session, 'history.undo').ok).toBe(true);
    expect(readControl(session.doc.child(block).attrs['sdtPr']).cannotDelete).toBe(true);
  });
});

describe('row and cell controls', () => {
  test('rows wrapped become one w:sdt around them; properties edit the map entry; unwrap pops the group, delete removes the rows; one undo step each', async () => {
    const original = await fixture('tables.docx');
    const session = await EditorDocument.open(original);
    const doc = session.doc;
    const tablePos = blockStart(session, 1);
    const map = TableMap.get(doc.child(1));
    const a = tablePos + 1 + map.map[0]!;
    const b = tablePos + 1 + map.map[map.width]!;   // row 1, column 0
    // The command asks for the map write on its transaction; the runner puts it in the same Yjs transaction as the document.
    const state = stateFor(session, CellSelection.create(doc, a, b));
    let asked = 0;
    let selected = '';
    expect(commandNamed('contentControl.wrapRows')!.run({ tag: 'od:repeat=x1', title: 'Items' })(state, (tr) => {
      asked = controlWritesOf(tr).length;
      // The wrapped rows' cells are selected afterwards, so the highlight shows the control's extent (Jason, 2026-10-08).
      const sel = tr.selection;
      if (sel instanceof CellSelection) { const r = map.rectBetween(sel.$anchorCell.pos - tablePos - 1, sel.$headCell.pos - tablePos - 1); selected = `${r.top}-${r.bottom} ${r.left}-${r.right}`; }
    })).toBe(true);
    expect(asked).toBe(1);
    expect(selected).toBe(`0-2 0-${map.width}`);
    expect(runCommand(session, 'contentControl.wrapRows', { tag: 'od:repeat=x1', title: 'Items' }, CellSelection.create(doc, a, b)).ok).toBe(true);
    const table = session.doc.child(1);
    const g0 = table.child(0).attrs['sdt'] as SdtGroup[];
    const g1 = table.child(1).attrs['sdt'] as SdtGroup[];
    expect(g0).toHaveLength(1);
    expect(g0[0]).toMatchObject({ first: true, last: false });
    expect(g1[0]).toMatchObject({ groupId: g0[0]!.groupId, first: false, last: true });
    expect(g0[0]!.groupId.startsWith('new:')).toBe(true);
    const entry = readControlEntry(session.store, g0[0]!.groupId)!;
    expect(entry.form).toBe('row');
    expect(readControl(entry.sdtPr)).toMatchObject({ tag: 'od:repeat=x1', title: 'Items' });
    expect(listControls(session.doc, session.controlEntry)[0]).toMatchObject({ form: 'row', rows: [0, 1], depth: 0 });
    const { api, body } = await saveAndReload(session, original);
    const tbl = body.content![1]!.value as wml.Tbl;
    expect(tbl.content!.map((e) => typeNameOf(e))).toEqual(['org_docx4j_wml.CTSdtRow']);
    expect(rowsOf(tbl)).toHaveLength(2);
    const control = api.contentControls.find((c) => c.form === 'Row')!;
    expect([control.tag, control.title]).toEqual(['od:repeat=x1', 'Items']);
    expect(control.id).toBeGreaterThan(0);
    // The properties of a group go to the map entry, and export writes them.
    expect(runCommand(session, 'contentControl.properties', { title: 'Lines' }, { from: cellStart(session, 1, 0, 0) }).ok).toBe(true);
    expect(sdtTitle(readControlEntry(session.store, g0[0]!.groupId)!.sdtPr)).toBe('Lines');
    const again = await saveAndReload(session, original);
    expect(again.api.contentControls.find((c) => c.form === 'Row')!.title).toBe('Lines');
    // Undo the properties, then the wrap: two steps, the map entry going with the second.
    expect(runCommand(session, 'history.undo').ok).toBe(true);
    expect(sdtTitle(readControlEntry(session.store, g0[0]!.groupId)!.sdtPr)).toBe('Items');
    expect(runCommand(session, 'history.undo').ok).toBe(true);
    expect(session.doc.child(1).child(0).attrs['sdt']).toBeNull();
    expect(readControlEntry(session.store, g0[0]!.groupId)).toBeUndefined();
    expect(runCommand(session, 'history.redo').ok).toBe(true);
    expect(readControlEntry(session.store, g0[0]!.groupId)).toBeDefined();
    // Unwrap keeps the rows.
    expect(runCommand(session, 'contentControl.delete', { keepContent: true }, { from: cellStart(session, 1, 1, 0) }).ok).toBe(true);
    expect(session.doc.child(1).childCount).toBe(2);
    expect(session.doc.child(1).child(0).attrs['sdt']).toBeNull();
    expect(readControlEntry(session.store, g0[0]!.groupId)).toBeUndefined();
    // Wrap again and delete with the content: the rows go; the table is gone when they were all of it.
    expect(runCommand(session, 'contentControl.wrapRows', undefined, CellSelection.create(session.doc, a, b)).ok).toBe(true);
    expect(runCommand(session, 'contentControl.delete', undefined, { from: cellStart(session, 1, 0, 0) }).ok).toBe(true);
    expect(session.doc.child(1).type.name).not.toBe('table');
    const final = await saveAndReload(session, original);
    expect(typeNameOf(final.body.content![1]!)).not.toBe('org_docx4j_wml.Tbl');
  });

  test('the properties of an imported group edited with nothing else changed: the export sees the map entry (a row group, a cell group), and undo restores the untouched export', async () => {
    const original = await fixture('invoice.docx');
    const session = await EditorDocument.open(original);
    const cell = cellStart(session, 5, 1, 0);   // the first cell of the repeat row: its cell control
    expect(runCommand(session, 'contentControl.properties', { title: 'Item name' }, { from: cell }).ok).toBe(true);
    expect(session.doc.eq(session.imported)).toBe(true);   // a map write alone
    expect(session.exportTree().changed).toBe(true);
    expect(runCommand(session, 'history.undo').ok).toBe(true);
    expect(session.exportTree().changed).toBe(false);
    expect(runCommand(session, 'history.redo').ok).toBe(true);
    const first = await saveAndReload(session, original);
    expect(first.api.contentControls.filter((c) => c.form === 'Cell').map((c) => c.title)).toEqual(['Item name', 'Price']);
    // The row group, selected as a whole: the properties command targets it, not the first cell's control.
    const table = session.doc.child(5);
    const rowStart = blockStart(session, 5) + 1 + table.child(0).nodeSize;
    const sel = CellSelection.create(session.doc, rowStart + 1, rowStart + 1 + table.child(1).child(0).nodeSize);
    expect(runCommand(session, 'contentControl.properties', { title: 'Items' }, sel).ok).toBe(true);
    const groupId = (table.child(1).attrs['sdt'] as SdtGroup[])[0]!.groupId;
    expect(sdtTitle(readControlEntry(session.store, groupId)!.sdtPr)).toBe('Items');
    expect(session.exportTree().changed).toBe(true);
    const second = await saveAndReload(session, original);
    expect(second.api.contentControls.find((c) => c.form === 'Row')!.title).toBe('Items');
    expect(second.api.contentControls.find((c) => c.form === 'Row')!.tag).toBe('od:repeat=x2&od:finish=t_r');
  });

  test('a group wrapped inside an existing repeat nests inside it on export; one row of a two-row group wrapped nests too', async () => {
    const original = await fixture('invoice.docx');
    const session = await EditorDocument.open(original);
    const repeatRow = cellStart(session, 5, 1, 0);
    expect(runCommand(session, 'contentControl.wrapRows', { title: 'Inner' }, { from: repeatRow }).ok).toBe(true);
    const stack = session.doc.child(5).child(1).attrs['sdt'] as SdtGroup[];
    expect(stack).toHaveLength(2);
    expect(stack[0]!.groupId.startsWith('new:')).toBe(false);   // the repeat, outermost
    expect(stack[1]!.groupId.startsWith('new:')).toBe(true);
    const { api, body } = await saveAndReload(session, original);
    const tbl = body.content![5]!.value as wml.Tbl;
    const outer = tbl.content!.find((e) => typeNameOf(e) === 'org_docx4j_wml.CTSdtRow')!.value as wml.CTSdtRow;
    expect(outer.sdtContent!.content!.map((e) => typeNameOf(e))).toEqual(['org_docx4j_wml.CTSdtRow']);
    expect(api.contentControls.filter((c) => c.form === 'Row').map((c) => c.title)).toEqual(['Repeat', 'Inner']);
    expect(api.contentControls.filter((c) => c.form === 'Cell')).toHaveLength(2);   // the cell controls survive inside
    // A wrap crossing a group's edge is refused: the header row plus the repeat row, where the new group would neither contain nor sit inside the repeat... it contains it, so it is allowed and goes outside.
    const doc = session.doc;
    const tablePos = blockStart(session, 5);
    const map = TableMap.get(doc.child(5));
    const sel = CellSelection.create(doc, tablePos + 1 + map.map[0]!, tablePos + 1 + map.map[map.width]!);
    expect(runCommand(session, 'contentControl.wrapRows', { title: 'Outer' }, sel).ok).toBe(true);
    expect((session.doc.child(5).child(1).attrs['sdt'] as SdtGroup[]).map((g) => sdtTitle(readControlEntry(session.store, g.groupId)!.sdtPr))).toEqual(['Outer', 'Repeat', 'Inner']);
  });

  test('cells wrapped become a w:sdt around the cells of one row; delete without keeping empties them', async () => {
    const original = await fixture('tables.docx');
    const session = await EditorDocument.open(original);
    const doc = session.doc;
    const tablePos = blockStart(session, 1);
    const map = TableMap.get(doc.child(1));
    const sel = CellSelection.create(doc, tablePos + 1 + map.map[0]!, tablePos + 1 + map.map[1]!);
    expect(runCommand(session, 'contentControl.wrapCells', { tag: 'od:xpath=x3' }, sel).ok).toBe(true);
    const row = session.doc.child(1).child(0);
    const s0 = row.child(0).attrs['sdt'] as SdtGroup[];
    expect(s0[0]).toMatchObject({ first: true, last: false });
    expect((row.child(1).attrs['sdt'] as SdtGroup[])[0]).toMatchObject({ groupId: s0[0]!.groupId, first: false, last: true });
    expect(readControlEntry(session.store, s0[0]!.groupId)!.form).toBe('cell');
    const { api, body } = await saveAndReload(session, original);
    const tr = rowsOf(body.content![1]!.value as wml.Tbl)[0]!.element.value;
    expect(tr.content!.map((e) => typeNameOf(e))).toEqual(['org_docx4j_wml.CTSdtCell']);
    expect(api.contentControls.find((c) => c.form === 'Cell')!.tag).toBe('od:xpath=x3');
    expect(runCommand(session, 'contentControl.delete', undefined, { from: cellStart(session, 1, 0, 1) }).ok).toBe(true);
    expect(session.doc.child(1).child(0).childCount).toBe(2);
    expect(textOf(session.doc.child(1).child(0))).toBe('');
    expect(session.doc.child(1).child(0).child(0).attrs['sdt']).toBeNull();
    // A selection spanning two rows is not a cell control.
    const now = TableMap.get(session.doc.child(1));
    expect(runCommand(session, 'contentControl.wrapCells', undefined, CellSelection.create(session.doc, tablePos + 1 + now.map[0]!, tablePos + 1 + now.map[now.width]!)).ok).toBe(false);
  });

  test('select: a group as a cell selection over its rows, a node as its content', async () => {
    const session = await EditorDocument.open(await fixture('invoice.docx'));
    const groupId = (session.doc.child(5).child(1).attrs['sdt'] as SdtGroup[])[0]!.groupId;
    const r = runCommand(session, 'contentControl.select', { groupId });
    expect(r.ok).toBe(true);
    expect(r.state!.selection).toBeInstanceOf(CellSelection);
    expect((r.state!.selection as CellSelection).$anchorCell.node(-1).attrs['sdt']).not.toBeNull();
    const pos = listControls(session.doc, session.controlEntry)[0]!.pos!;
    const node = session.doc.nodeAt(pos)!;
    const n = runCommand(session, 'contentControl.select', { pos });
    expect(n.state!.selection).toBeInstanceOf(TextSelection);
    // The whole content: from the first text position inside to the last.
    expect(n.state!.selection.from).toBe(pos + 2);
    expect(n.state!.selection.to).toBe(pos + node.nodeSize - 2);
    expect(n.state!.selection.$from.node(1).type.name).toBe('sdtBlock');
  });
});

describe('bindings and the OpenDoPE roles', () => {
  test('bind: w:dataBinding written and read back through XmlMapping; with opendope, the entry and od:xpath; a new template gets its XPaths part and the authoring tool stamp', async () => {
    const original = await fixture('HelloWordOnline.docx');
    const session = await EditorDocument.open(original);
    const data = session.pkg.customXmlParts.add('<data><name>Ann</name><city>Perth</city></data>');
    await session.template.load();
    const start = blockStart(session, 0) + 1;
    expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: start, to: start + 3 }).ok).toBe(true);
    const inside = { from: start + 2 };
    // An XPath that selects nothing binds nothing, as Word reports.
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/nothing' }, inside).ok).toBe(false);
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/name' }, inside).ok).toBe(true);
    const bound = session.doc.child(0).child(0);
    expect(sdtBinding(bound.attrs['sdtPr'])).toEqual({ xpath: '/data/name', storeItemID: data.id, prefixMappings: '', w15: false });
    expect(session.template.verdictFor(bound.attrs['sdtPr'])).toMatchObject({ role: 'none', value: 'Ann' });
    expect(session.template.openDopeParts.xpaths).toBeUndefined();   // no XPaths part: a plain Word binding
    expect(session.template.touched).toBe(false);
    // With opendope: the XPaths part is created, the entry written, the tag set (REC-002).
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/city', opendope: true }, inside).ok).toBe(true);
    const od = session.doc.child(0).child(0);
    expect(readControl(od.attrs['sdtPr']).tag).toBe('od:xpath=x1');
    expect(session.template.xpaths!.entries).toEqual([{ id: 'x1', xpath: '/data/city', storeItemID: data.id, prefixMappings: '', attrs: {} }]);
    expect(session.template.xpaths!).toMatchObject({ version: '3.0', booleanConversion: 'xpath2' });
    expect(session.template.touched).toBe(true);
    expect([...session.touchedParts].some((p) => p.startsWith('/customXml/'))).toBe(true);
    // Binding again keeps the entry id and updates its expression.
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/name' }, inside).ok).toBe(true);
    expect(session.template.xpaths!.entries[0]).toMatchObject({ id: 'x1', xpath: '/data/name' });
    expect(readControl(session.doc.child(0).child(0).attrs['sdtPr']).tag).toBe('od:xpath=x1');
    const { pkg, api } = await saveAndReload(session, original);
    const control = api.contentControls[0]!;
    // The mapping is read against its store item, so the custom XML parts are parsed first (core-ts 0.7.0, CR-020).
    await pkg.customXmlParts.load();
    expect(control.xmlMapping.isMapped).toBe(true);
    expect(control.xmlMapping.xpath).toBe('/data/name');
    expect(control.xmlMapping.customXmlNode!.text).toBe('Ann');
    const reopened = await EditorDocument.open((await session.save()).bytes);
    await reopened.template.load();
    expect(reopened.template.xpaths).toMatchObject({ version: '3.0', booleanConversion: 'xpath2', authoringTool: 'docx4j-editor', authoringToolVersion: '00.0100', entries: [{ id: 'x1', xpath: '/data/name' }] });
    expect(reopened.template.dataPart).toMatchObject({ rule: 1 });
    expect(reopened.template.problems).toEqual([]);
    // Unbind removes both the binding and od:xpath.
    expect(runCommand(session, 'contentControl.bind', null, inside).ok).toBe(true);
    expect(sdtBinding(session.doc.child(0).child(0).attrs['sdtPr'])).toBeUndefined();
    expect(readControl(session.doc.child(0).child(0).attrs['sdtPr']).tag).toBe('');
  });

  // ED-005 section 12.54 (Jason's report): binding showed only in the panel's verdict. Word's XML mapping shows the node's
  // value in the control the moment it is mapped, and so does the editor now.
  test('bind shows the bound value in the control: its text, the placeholder for an empty node, lines in a multi-line one, a drop-down\'s display text; one undo step; saved', async () => {
    const original = await fixture('HelloWordOnline.docx');
    const session = await EditorDocument.open(original);
    session.pkg.customXmlParts.add('<data><name>Ann</name><city>Perth</city><empty/><lines>one\ntwo</lines><choice>b</choice></data>');
    await session.template.load();
    const start = blockStart(session, 0) + 1;
    expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: start, to: start + 3 }).ok).toBe(true);
    const inside = { from: start + 2 };
    const control = (): PMNode => session.doc.child(0).child(0);
    const firstText = (): PMNode => { let t: PMNode | undefined; control().descendants((n) => { if (!t && n.isText) t = n; return !t; }); return t!; };
    const kinds = (): string[] => control().content.content.map((n) => (n.isText ? 'text' : String(n.attrs['kind'] ?? n.type.name)));
    // The control's content holds the document's bookmark markers before its text.
    expect(kinds()).toEqual(['bookmarkStart', 'bookmarkEnd', 'text']);
    const marksBefore = firstText().marks.map((m) => m.type.name);
    const bound = runCommand(session, 'contentControl.bind', { xpath: '/data/name' }, inside);
    expect(bound.ok).toBe(true);
    expect(control().textContent).toBe('Ann');
    // The caret stays in the control, after the value.
    expect(bound.state!.selection.$from.parent.type.name).toBe('sdtInline');
    // The text keeps the formatting of the control's first text, and the bookmark's markers stay.
    expect(firstText().marks.map((m) => m.type.name)).toEqual(marksBefore);
    expect(kinds()).toEqual(['bookmarkStart', 'bookmarkEnd', 'text']);
    expect(readControl(control().attrs['sdtPr']).showingPlaceholder).toBe(false);
    // Bound again, the new node's value.
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/city' }, inside).ok).toBe(true);
    expect(control().textContent).toBe('Perth');
    // One undo step takes the binding and the text back together.
    session.undoManager.stopCapturing();
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/empty' }, inside).ok).toBe(true);
    // An empty node: the placeholder, as Word shows it.
    expect(control().textContent).toBe(PLACEHOLDER_TEXT);
    expect(readControl(control().attrs['sdtPr']).showingPlaceholder).toBe(true);
    session.undoManager.undo();
    expect(control().textContent).toBe('Perth');
    expect(sdtBinding(control().attrs['sdtPr'])?.xpath).toBe('/data/city');
    // A multi-line control takes a line break per line; a single-line one, a space.
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/lines' }, inside).ok).toBe(true);
    expect(control().textContent).toBe('one two');
    // The control's text came from the node: nothing is written back, and the node keeps its two lines (ED-005 section 12.55).
    expect(session.template.dataPart.part!.selectSingleNode('/data/lines')!.text).toBe('one\ntwo');
    expect(runCommand(session, 'contentControl.properties', { multiLine: true }, inside).ok).toBe(true);
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/lines' }, inside).ok).toBe(true);
    expect(control().content.content.map((n) => (n.isText ? n.text : String(n.attrs['kind'] === 'line' ? 'hardBreak' : n.attrs['kind'])))).toEqual(['bookmarkStart', 'bookmarkEnd', 'one', 'hardBreak', 'two']);
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/name' }, inside).ok).toBe(true);
    // Saved: the control holds the value in the file, as its mapping reads it.
    const { api, pkg } = await saveAndReload(session, original);
    await pkg.customXmlParts.load();
    expect(api.contentControls[0]!.text).toBe('Ann');
    expect(api.contentControls[0]!.xmlMapping.customXmlNode!.text).toBe('Ann');
  });

  test('bind: a drop-down shows its list item\'s display text; a date picker is left as it is', async () => {
    const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
    session.pkg.customXmlParts.add('<data><choice>b</choice><when>2026-10-03</when></data>');
    await session.template.load();
    const start = blockStart(session, 0) + 1;
    expect(runCommand(session, 'contentControl.insert', { kind: 'DropDownList' }, { from: start, to: start + 3 }).ok).toBe(true);
    const control = (): PMNode => session.doc.child(0).child(0);
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const items = { name: { namespaceURI: W, localPart: 'dropDownList' }, value: { TYPE_NAME: 'org_docx4j_wml.CTSdtDropDownList', listItem: [{ TYPE_NAME: 'org_docx4j_wml.CTSdtListItem', displayText: 'Alpha', value: 'a' }, { TYPE_NAME: 'org_docx4j_wml.CTSdtListItem', displayText: 'Bee', value: 'b' }] } };
    const state = stateFor(session);
    session.replace(state.apply(state.tr.setNodeMarkup(blockStart(session, 0) + 1, undefined, { ...control().attrs, sdtPr: withSdtItem(control().attrs['sdtPr'], items) })).doc);
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/choice' }, { from: start + 2 }).ok).toBe(true);
    expect(control().textContent).toBe('Bee');
    // A date picker: Word formats the date by w:dateFormat; the editor leaves the content as it is.
    const dateAt = blockStart(session, 0) + 1 + control().nodeSize;
    expect(runCommand(session, 'contentControl.insert', { kind: 'DatePicker' }, { from: dateAt, to: dateAt + 2 }).ok).toBe(true);
    const date = session.doc.child(0).child(1);
    const before = date.textContent;
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/when' }, { from: dateAt + 2 }).ok).toBe(true);
    expect(session.doc.child(0).child(1).textContent).toBe(before);
  });

  // ED-005 section 12.56 (Jason's request): the Data tab's bind with no control at the caret makes one, a plain text
  // control at the caret or around a selection within a paragraph, bound in one step; elsewhere it says why not.
  describe('a plain text control made and bound in one step (contentControl.insertBound)', () => {
    async function withData(name = 'HelloWordOnline.docx'): Promise<EditorDocument> {
      const session = await EditorDocument.open(await fixture(name));
      session.pkg.customXmlParts.add('<data><name>Ann</name><empty/></data>');
      await session.template.load();
      return session;
    }
    const controls = (doc: PMNode): PMNode[] => { const out: PMNode[] = []; doc.descendants((n) => { if (n.type.name === 'sdtInline') out.push(n); return true; }); return out; };

    test('a block control made at the end of the document keeps an empty paragraph after it, as Word does: the final mark is never inside a control', async () => {
      const session = await withData();
      const last = session.doc.childCount - 1;
      const para = session.doc.child(last);
      expect(para.type.name).toBe('paragraph');
      // 1. The whole last paragraph selected: a block control, then a paragraph, and the caret can leave the control.
      const start = blockStart(session, last) + 1;
      expect(runCommand(session, 'contentControl.insert', {}, { from: start, to: start + para.content.size }).ok).toBe(true);
      expect(session.doc.child(last).type.name).toBe('sdtBlock');
      expect(session.doc.childCount).toBe(last + 2);
      expect(session.doc.lastChild!.type.name).toBe('paragraph');
      expect(session.doc.lastChild!.textContent).toBe('');
      // 2. A bind at a caret in that empty last paragraph is a run control inside it (Jason, 2026-10-08: never a block one), so the paragraph stays last.
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/name' }, { from: blockStart(session, last + 1) + 1 }).ok).toBe(true);
      expect(session.doc.childCount).toBe(last + 2);
      expect(session.doc.lastChild!.type.name).toBe('paragraph');
      expect(session.doc.lastChild!.child(0).type.name).toBe('sdtInline');
      expect(session.doc.lastChild!.textContent).toBe('Ann');
      // 3. Inside a block control: the paragraph that ends the outer control's content wrapped in a block control gets a paragraph after it, inside the outer.
      const outer = session.doc.child(last);
      const innerStart = blockStart(session, last) + 1 + 1;
      expect(runCommand(session, 'contentControl.insert', {}, { from: innerStart, to: innerStart + outer.child(0).content.size }).ok).toBe(true);
      const outerAfter = session.doc.child(last);
      expect(outerAfter.childCount).toBe(2);
      expect(outerAfter.child(0).type.name).toBe('sdtBlock');
      expect(outerAfter.child(1).type.name).toBe('paragraph');
      expect(outerAfter.child(1).textContent).toBe('');
      expect(session.doc.childCount).toBe(last + 2);
      // 4. Not at the end, nothing is added: the first paragraph wrapped leaves the count alone.
      const count = session.doc.childCount;
      const first = session.doc.child(0);
      expect(runCommand(session, 'contentControl.insert', {}, { from: 1, to: 1 + first.content.size }).ok).toBe(true);
      expect(session.doc.childCount).toBe(count);
      // Saved (the self-check runs) and reopened, the document still ends with the paragraph.
      const reopened = await EditorDocument.open((await session.save()).bytes);
      expect(reopened.doc.lastChild!.type.name).toBe('paragraph');
      expect(reopened.doc.child(reopened.doc.childCount - 2).type.name).toBe('sdtBlock');
    });

    test('a role set adopts the bindings made before the XPaths part existed: each gets an entry and the od:xpath tag, one entry per XPath and part (ED-005 section 12.77)', async () => {
      const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
      session.pkg.customXmlParts.add('<data><name>Ann</name><city>Oslo</city></data>');
      await session.template.load();
      expect(session.template.xpaths).toBeUndefined();
      // Two controls bound while there is no XPaths part: w:dataBinding alone.
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/name' }, { from: 3 }).ok).toBe(true);
      const second = 1 + session.doc.child(0).content.size;
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/name' }, { from: second }).ok).toBe(true);
      const bound = (): PMNode[] => { const out: PMNode[] = []; session.doc.child(0).forEach((n) => { if (n.type.name === 'sdtInline') out.push(n); }); return out; };
      expect(bound()).toHaveLength(2);
      expect(bound().every((n) => readTag(sdtTag(n.attrs['sdtPr'])).xpath === undefined)).toBe(true);
      // A condition set on a block control around the second paragraph makes the part: the two bindings are entered and tagged.
      const p1 = session.doc.child(1);
      const start = session.doc.child(0).nodeSize + 1;
      expect(runCommand(session, 'contentControl.insert', {}, { from: start, to: start + p1.content.size }).ok).toBe(true);
      expect(runCommand(session, 'opendope.condition', { id: session.template.freshConditionId(), define: { op: 'xpathref', id: session.template.upsertEntry({ xpath: '/data/city' }).id } }, { from: start + 1 }).ok).toBe(true);
      const entries = session.template.xpaths!.entries;
      const tags = bound().map((n) => readTag(sdtTag(n.attrs['sdtPr'])).xpath);
      expect(tags[0]).toBeDefined();
      expect(tags[1]).toBe(tags[0]);   // the same XPath and part: one entry
      expect(entries.find((e) => e.id === tags[0])!.xpath).toBe('/data/name');
      expect(entries.filter((e) => e.xpath === '/data/name')).toHaveLength(1);
    });

    test('a binding inside a repeat is written relative to it (ED-005 section 12.77): at bind time inside an existing repeat, and when the repeat is set after the binds; docx4j contextualises by the textual prefix', async () => {
      // invoice.docx: its items rows are already docx4j's repeat x2 (/invoice[1]/items/item), so a bind there is relative at once.
      const session = await EditorDocument.open(await fixture('invoice.docx'));
      await session.template.load();
      const cell = cellStart(session, 5, 1, 0);
      expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: cell, to: cell + 1 }).ok).toBe(true);
      expect(runCommand(session, 'contentControl.bind', { xpath: '/invoice[1]/items[1]/item[1]/name[1]' }, { from: cell + 1 }).ok).toBe(true);
      const inlineIn = (doc: PMNode, block: number, row: number, cellIndex: number): PMNode => { let found: PMNode | undefined; doc.child(block).child(row).child(cellIndex).child(0).forEach((n) => { if (n.type.name === 'sdtInline') found = n; }); expect(found).toBeDefined(); return found!; };
      const control = inlineIn(session.doc, 5, 1, 0);
      expect(sdtBinding(control.attrs['sdtPr'])!.xpath).toBe('/invoice[1]/items/item[1]/name[1]');
      const tag = readTag(sdtTag(control.attrs['sdtPr']));
      expect(session.template.xpaths!.entries.find((e) => e.id === tag.xpath)!.xpath).toBe('/invoice[1]/items/item[1]/name[1]');
      // Jason's order, on a plain table: bound through the Data tab's canonical XPath first, the rows made a repeat afterwards.
      const plain = await EditorDocument.open(await fixture('tables.docx'));
      plain.pkg.customXmlParts.add('<invoice><items><item><name>a</name></item><item><name>b</name></item></items></invoice>');
      await plain.template.load();
      let tableIndex = -1;
      plain.doc.forEach((n, _o, i) => { if (tableIndex < 0 && n.type.name === 'table') tableIndex = i; });
      expect(tableIndex).toBeGreaterThanOrEqual(0);
      const at = cellStart(plain, tableIndex, 0, 0);
      expect(runCommand(plain, 'contentControl.insertBound', { xpath: '/invoice[1]/items[1]/item[1]/name[1]' }, { from: at }).ok).toBe(true);
      expect(sdtBinding(inlineIn(plain.doc, tableIndex, 0, 0).attrs['sdtPr'])!.xpath).toBe('/invoice[1]/items[1]/item[1]/name[1]');
      expect(runCommand(plain, 'contentControl.wrapRows', undefined, { from: cellStart(plain, tableIndex, 0, 0) }).ok).toBe(true);
      expect(runCommand(plain, 'opendope.repeat', { id: 'rows', define: { xpath: '/invoice[1]/items/item' } }, { from: cellStart(plain, tableIndex, 0, 0) }).ok).toBe(true);
      const adopted = inlineIn(plain.doc, tableIndex, 0, 0);
      expect(sdtBinding(adopted.attrs['sdtPr'])!.xpath).toBe('/invoice[1]/items/item[1]/name[1]');
      const adoptedTag = readTag(sdtTag(adopted.attrs['sdtPr']));
      expect(plain.template.xpaths!.entries.find((e) => e.id === adoptedTag.xpath)!.xpath).toBe('/invoice[1]/items/item[1]/name[1]');
      // Pure: the rewriting's shape, and where it does not apply.
      expect(relativeToRepeat('/invoice[1]/items[1]/item[1]/name[1]', '/invoice[1]/items/item')).toBe('/invoice[1]/items/item[1]/name[1]');
      expect(relativeToRepeat('/invoice[1]/items[1]/item[2]/name', '/invoice/items/item')).toBe('/invoice/items/item[2]/name');
      expect(relativeToRepeat('/invoice[1]/customer[1]/name[1]', '/invoice[1]/items/item')).toBeUndefined();
      expect(relativeToRepeat('/invoice[1]/items/item', '/invoice[1]/items/item')).toBeUndefined();
    });

    test("a node holding XHTML binds as a rich text control through the entry (ED-005 section 12.79): block markup around a blank paragraph, refused inside one with content; inline markup as a run-level control", async () => {
      const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
      session.pkg.customXmlParts.add('<data><block>&lt;p&gt;Hi &lt;b&gt;there&lt;/b&gt;&lt;/p&gt;&lt;ul&gt;&lt;li&gt;one&lt;/li&gt;&lt;/ul&gt;</block><inline>&lt;b&gt;bold&lt;/b&gt; words</inline><plain>Ann</plain></data>');
      await session.template.load();
      expect(xhtmlKindOf('<p>x</p>')).toBe('block');
      expect(xhtmlKindOf('<b>x</b>')).toBe('inline');
      expect(xhtmlKindOf('plain < text > here')).toBe('none');
      // Inside a paragraph with content: refused.
      expect(dataBindRefusal(stateFor(session, { from: 3 }), { xpath: '/data/block' })).toBe('blockXhtml');
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/block' }, { from: 3 }).ok).toBe(false);
      // In a blank paragraph (one made by Enter at a paragraph's end): a block-level rich text control around it.
      const end = 1 + session.doc.child(0).content.size;
      expect(runCommand(session, 'insertParagraph', undefined, { from: end }).ok).toBe(true);
      const blank = session.doc.child(1);
      expect(blank.textContent).toBe('');
      const at = session.doc.child(0).nodeSize + 1;
      expect(dataBindRefusal(stateFor(session, { from: at }), { xpath: '/data/block' })).toBeUndefined();
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/block' }, { from: at }).ok).toBe(true);
      const made = session.doc.child(1);
      expect(made.type.name).toBe('sdtBlock');
      expect(readControl(made.attrs['sdtPr']).kind).toBe('RichText');
      expect(sdtBinding(made.attrs['sdtPr'])).toBeUndefined();
      expect(isXhtmlControl(made.attrs['sdtPr'])).toBe(true);
      expect(xhtmlBindingOf(made.attrs['sdtPr'], session.template)!.xpath).toBe('/data/block');
      // Inline markup: a run-level rich text control at the caret, tagged the same way.
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/inline' }, { from: 3 }).ok).toBe(true);
      let inline: PMNode | undefined;
      session.doc.child(0).forEach((n) => { if (n.type.name === 'sdtInline') inline = n; });
      expect(inline).toBeDefined();
      expect(readControl(inline!.attrs['sdtPr']).kind).toBe('RichText');
      expect(isXhtmlControl(inline!.attrs['sdtPr'])).toBe(true);
      // A plain node still makes a plain text control with a w:dataBinding.
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/plain' }, { from: 3 }).ok).toBe(true);
      const kinds: string[] = [];
      let plain: PMNode | undefined;
      session.doc.child(0).forEach((n) => { if (n.type.name === 'sdtInline') { kinds.push(readControl(n.attrs['sdtPr']).kind); if (readControl(n.attrs['sdtPr']).kind === 'PlainText') plain = n; } });
      expect(kinds.sort()).toEqual(['PlainText', 'RichText']);
      expect(sdtBinding(plain!.attrs['sdtPr'])!.xpath).toBe('/data/plain');
    });

    test('contentControl.xhtml: on moves the binding to an XPaths entry with the content type on the tag and no w:dataBinding; off brings the w:dataBinding back (ED-005 section 12.75)', async () => {
      const session = await withData();
      const caret = { from: blockStart(session, 0) + 1 + 2 + 5 };
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/name' }, caret).ok).toBe(true);
      const at = { from: caret.from + 1 };
      expect(runCommand(session, 'contentControl.xhtml', { on: true }, at).ok).toBe(true);
      const [made] = controls(session.doc);
      expect(sdtBinding(made!.attrs['sdtPr'])).toBeUndefined();
      const tag = readTag(sdtTag(made!.attrs['sdtPr']));
      expect(tag.contentType).toBe('application/xhtml+xml');
      expect(tag.xpath).toBeDefined();
      const entry = session.template.xpaths!.entries.find((e) => e.id === tag.xpath)!;
      expect(entry.xpath).toBe('/data/name');
      expect(isXhtmlControl(made!.attrs['sdtPr'])).toBe(true);
      expect(xhtmlBindingOf(made!.attrs['sdtPr'], session.template)).toMatchObject({ xpath: '/data/name', storeItemID: entry.storeItemID });
      expect(xhtmlControlsIn(session.doc, session.template)).toHaveLength(1);
      // Off: the text binding again, from the entry.
      expect(runCommand(session, 'contentControl.xhtml', { on: false }, at).ok).toBe(true);
      const [back] = controls(session.doc);
      expect(sdtBinding(back!.attrs['sdtPr'])).toMatchObject({ xpath: '/data/name' });
      expect(readTag(sdtTag(back!.attrs['sdtPr'])).contentType).toBeUndefined();
      expect(isXhtmlControl(back!.attrs['sdtPr'])).toBe(false);
    });

    test('at a caret: a plain text control showing the node\'s value, bound, the caret in it; one undo step takes it all', async () => {
      const session = await withData();
      const text = session.doc.child(0).textContent;
      const caret = { from: blockStart(session, 0) + 1 + 2 + 5 };   // after the bookmark's two markers and "Hello"
      expect(plainTextControlRefusal(stateFor(session, caret))).toBeUndefined();
      session.undoManager.stopCapturing();
      const r = runCommand(session, 'contentControl.insertBound', { xpath: '/data/name' }, caret);
      expect(r.ok).toBe(true);
      const [made] = controls(session.doc);
      expect(made!.textContent).toBe('Ann');
      expect(readControl(made!.attrs['sdtPr']).kind).toBe('PlainText');
      expect(sdtBinding(made!.attrs['sdtPr'])?.xpath).toBe('/data/name');
      expect(session.doc.child(0).textContent).toBe(text.slice(0, 5) + 'Ann' + text.slice(5));
      expect(r.state!.selection.$from.parent.type.name).toBe('sdtInline');
      session.undoManager.undo();
      expect(controls(session.doc)).toEqual([]);
      expect(session.doc.child(0).textContent).toBe(text);
    });

    test('around a selection within a paragraph: the selection becomes the control, which shows the node\'s value; an empty node shows the placeholder', async () => {
      const session = await withData();
      const start = blockStart(session, 0) + 1 + 2;   // the text, after the bookmark's markers
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/name' }, { from: start, to: start + 5 }).ok).toBe(true);
      expect(controls(session.doc)[0]!.textContent).toBe('Ann');
      expect(session.doc.child(0).textContent.startsWith('Ann')).toBe(true);
      const after = start + controls(session.doc)[0]!.nodeSize + 1;
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/empty' }, { from: after, to: after + 2 }).ok).toBe(true);
      expect(controls(session.doc)[1]!.textContent).toBe(PLACEHOLDER_TEXT);
      expect(readControl(controls(session.doc)[1]!.attrs['sdtPr']).showingPlaceholder).toBe(true);
    });

    test('refused, saying why: across paragraphs, over a note reference, over table cells, inside a plain text control (bind that one), an XPath that selects nothing', async () => {
      const session = await withData();
      // Across two paragraphs.
      const many = await withData('tracked-changes.docx');
      const across = { from: blockStart(many, 0) + 2, to: blockStart(many, 1) + 2 };
      expect(plainTextControlRefusal(stateFor(many, across))).toBe('paragraphs');
      expect(runCommand(many, 'contentControl.insertBound', { xpath: '/data/name' }, across).ok).toBe(false);
      // Nothing selected by the XPath.
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/nothing' }, { from: blockStart(session, 0) + 3 }).ok).toBe(false);
      // Inside a plain text control: the bind of that control is the panel's, not a new control.
      const start = blockStart(session, 0) + 1 + 2;
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/name' }, { from: start }).ok).toBe(true);
      expect(controlToBind(stateFor(session, { from: start + 2 }))).toMatchObject({ pos: start });
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/name' }, { from: start + 2 }).ok).toBe(false);
      expect(controls(session.doc)).toHaveLength(1);
      // Over a note reference.
      const notes = await withData('footnote.docx');
      let ref = -1;
      notes.doc.descendants((n, p) => { if (ref < 0 && n.type.name === 'footnoteRef') ref = p; return ref < 0; });
      const $ref = notes.doc.resolve(ref);
      expect(plainTextControlRefusal(stateFor(notes, { from: $ref.start(), to: ref + 1 }))).toBe('content');
      // Over table cells.
      const tables = await withData('tables.docx');
      let table = -1;
      tables.doc.forEach((b, _o, i) => { if (table < 0 && b.type.name === 'table') table = i; });
      const map = TableMap.get(tables.doc.child(table));
      const tablePos = blockStart(tables, table);
      expect(plainTextControlRefusal(stateFor(tables, CellSelection.create(tables.doc, tablePos + 1 + map.map[0]!, tablePos + 1 + map.map[1]!)))).toBe('cells');
    });

    // ED-005 section 12.57 (Jason's word, core-ts check 34 1h): inside a control that holds controls, the bind nests a new
    // one, as Word's XML Mapping Pane does; a plain text control, or a bound rich text control, is bound in place.
    test('inside a rich text control: a new plain text control nested in it; a placeholder there gives way; inside a mapped one too, but not on its node', async () => {
      const session = await withData();
      const start = blockStart(session, 0) + 1 + 2;
      expect(runCommand(session, 'contentControl.insert', { kind: 'RichText' }, { from: start, to: start + 5 }).ok).toBe(true);
      const caret = { from: start + 1 + 2 };   // "He|llo" inside the rich text control
      expect(controlToBind(stateFor(session, caret))).toBeUndefined();
      expect(plainTextControlRefusal(stateFor(session, caret))).toBeUndefined();
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/name' }, caret).ok).toBe(true);
      const [outer, inner] = controls(session.doc);
      expect(readControl(outer!.attrs['sdtPr']).kind).toBe('RichText');
      expect(outer!.textContent).toBe('HeAnnllo');
      expect(outer!.child(1)).toBe(inner);
      expect(readControl(inner!.attrs['sdtPr']).kind).toBe('PlainText');
      expect(sdtBinding(inner!.attrs['sdtPr'])?.xpath).toBe('/data/name');
      // A rich text control showing its placeholder: the new control takes the placeholder's place, the flag cleared.
      const empty = await withData();
      const at = blockStart(empty, 0) + 1 + 2;
      expect(runCommand(empty, 'contentControl.insert', { kind: 'RichText' }, { from: at }).ok).toBe(true);
      expect(readControl(controls(empty.doc)[0]!.attrs['sdtPr']).showingPlaceholder).toBe(true);
      expect(runCommand(empty, 'contentControl.insertBound', { xpath: '/data/name' }, { from: at + 3 }).ok).toBe(true);
      const [holder, made] = controls(empty.doc);
      expect(holder!.textContent).toBe('Ann');
      expect(holder!.childCount).toBe(1);
      expect(readControl(holder!.attrs['sdtPr']).showingPlaceholder).toBe(false);
      expect(made!.child(0).marks.map((m) => m.type.name)).not.toContain('runStyle');
      // A rich text control bound in the editor is mapped as Word's pane maps one, with `w15:dataBinding` (core-ts check
      // 35, 1a, ED-005 section 12.58); Word nests inside it (1e), but refuses a plain text control on its node.
      const bound = await withData();
      const b = blockStart(bound, 0) + 1 + 2;
      expect(runCommand(bound, 'contentControl.insert', { kind: 'RichText' }, { from: b, to: b + 5 }).ok).toBe(true);
      expect(runCommand(bound, 'contentControl.bind', { xpath: '/data/name' }, { from: b + 2 }).ok).toBe(true);
      expect(sdtBinding(controls(bound.doc)[0]!.attrs['sdtPr'])).toMatchObject({ xpath: '/data/name', w15: true });
      expect(controls(bound.doc)[0]!.textContent).toBe('Ann');
      expect(controlToBind(stateFor(bound, { from: b + 2 }))).toBeUndefined();
      expect(dataBindRefusal(stateFor(bound, { from: b + 2 }), { xpath: '/data/name' })).toBe('mapped');
      expect(runCommand(bound, 'contentControl.insertBound', { xpath: '/data/name' }, { from: b + 2 }).ok).toBe(false);
      expect(dataBindRefusal(stateFor(bound, { from: b + 2 }), { xpath: '/data/empty' })).toBeUndefined();
      expect(runCommand(bound, 'contentControl.insertBound', { xpath: '/data/empty' }, { from: b + 2 }).ok).toBe(true);
      expect(controls(bound.doc).map((c) => readControl(c.attrs['sdtPr']).kind)).toEqual(['RichText', 'PlainText']);
      // Elsewhere too, the node takes no plain text control (Word's alert), and binding a plain text control to it is refused.
      const after = blockStart(bound, 0) + bound.doc.child(0).nodeSize - 1;
      expect(dataBindRefusal(stateFor(bound, { from: after }), { xpath: '/data/name' })).toBe('mapped');
      expect(runCommand(bound, 'contentControl.insertBound', { xpath: '/data/name' }, { from: after }).ok).toBe(false);
    });

    test('a rich text control bound with w:dataBinding is plain text, as Word takes it: bound in place, shown and written back', async () => {
      const session = await withData();
      const start = blockStart(session, 0) + 1 + 2;
      expect(runCommand(session, 'contentControl.insert', { kind: 'RichText' }, { from: start, to: start + 5 }).ok).toBe(true);
      // The form the engine's setMapping writes, and the editor wrote before ED-005 section 12.58.
      const part = session.template.dataPart.part!;
      const control = controls(session.doc)[0]!;
      const sdtPr = withSdtItem(control.attrs['sdtPr'], { name: { namespaceURI: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', localPart: 'dataBinding' }, value: { xpath: '/data/name', storeItemID: part.id, TYPE_NAME: 'org_docx4j_wml.CTDataBinding' } });
      const state = stateFor(session, { from: start });
      session.replace(state.apply(state.tr.setNodeMarkup(start, undefined, { ...control.attrs, sdtPr })).doc);
      expect(sdtBinding(controls(session.doc)[0]!.attrs['sdtPr'])?.w15).toBe(false);
      expect(controlToBind(stateFor(session, { from: start + 2 }))).toMatchObject({ pos: start });
      // Rebound in place, it keeps w:dataBinding; typing in it writes its node, as Word wrote 1f's.
      expect(runCommand(session, 'contentControl.bind', { xpath: '/data/name' }, { from: start + 2 }).ok).toBe(true);
      expect(sdtBinding(controls(session.doc)[0]!.attrs['sdtPr'])?.w15).toBe(false);
      expect(controls(session.doc)[0]!.textContent).toBe('Ann');
      expect(runCommand(session, 'insertText', { text: 'e' }, { from: start + 1 + 3 }).ok).toBe(true);
      expect(part.selectSingleNode('/data/name')!.text).toBe('Anne');
    });

    test('a selection to the next paragraph\'s start (the selection bar\'s) wraps one paragraph, and a paragraph command touches one (ED-005 section 12.81 item 5)', async () => {
      const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
      const first = session.doc.child(0);
      expect(first.type.name).toBe('paragraph');
      expect(session.doc.child(1).type.name).toBe('paragraph');
      const nextStart = first.nodeSize + 1;
      // Alignment over "the paragraph with its mark": the next paragraph untouched.
      expect(runCommand(session, 'paragraph.alignment', { value: 'Centered' }, { from: 1, to: nextStart }).ok).toBe(true);
      expect(session.doc.child(0).attrs['jc']).toBe('center');
      expect(session.doc.child(1).attrs['jc']).toBeNull();
      // A block-level rich text control around it alone.
      expect(runCommand(session, 'contentControl.insert', { kind: 'RichText' }, { from: 1, to: nextStart }).ok).toBe(true);
      const made = session.doc.child(0);
      expect(made.type.name).toBe('sdtBlock');
      expect(made.childCount).toBe(1);
      expect(made.child(0).textContent).toBe(first.textContent);
      expect(session.doc.child(1).type.name).toBe('paragraph');
      expect(session.doc.child(1).textContent).toBe(session.doc.child(1).textContent);
    });

    test('no control inside a plain text control, as Word refuses (ED-005 section 12.85); inside a rich text one, yes', async () => {
      const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
      expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: 3 }).ok).toBe(true);
      let plain: { pos: number; node: PMNode } | undefined;
      session.doc.child(0).forEach((n, offset) => { if (n.type.name === 'sdtInline') plain = { pos: 1 + offset, node: n }; });
      expect(plain).toBeDefined();
      expect(readControl(plain!.node.attrs['sdtPr']).kind).toBe('PlainText');
      const inside = plain!.pos + 2;
      expect(insideControlHoldingNoControls(stateFor(session, { from: inside }))).toBe(true);
      for (const kind of ['RichText', 'PlainText', 'CheckBox'] as const) {
        expect(runCommand(session, 'contentControl.insert', { kind }, { from: inside }).ok).toBe(false);
      }
      expect(commandNamed('contentControl.insert')!.isEnabled!(stateFor(session, { from: inside }))).toBe(false);
      let nested = 0;
      plain!.node.descendants((n) => { if (n.type.name === 'sdtInline' || n.type.name === 'sdtBlock') nested++; });
      expect(nested).toBe(0);
      // A rich text control holds one.
      const after = plain!.pos + plain!.node.nodeSize + 1;
      expect(runCommand(session, 'contentControl.insert', { kind: 'RichText' }, { from: after }).ok).toBe(true);
      let rich: { pos: number; node: PMNode } | undefined;
      session.doc.child(0).forEach((n, offset) => { if (n.type.name === 'sdtInline' && readControl(n.attrs['sdtPr']).kind === 'RichText') rich = { pos: 1 + offset, node: n }; });
      expect(rich).toBeDefined();
      expect(insideControlHoldingNoControls(stateFor(session, { from: rich!.pos + 2 }))).toBe(false);
      expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: rich!.pos + 2 }).ok).toBe(true);
    });

    test('inside a block-level rich text control: a run-level control in its paragraph, and in its empty paragraph too', async () => {
      const session = await withData('tracked-changes.docx');
      const para = session.doc.child(0);
      const start = blockStart(session, 0);
      expect(runCommand(session, 'contentControl.insert', { kind: 'RichText' }, { from: start + 1, to: start + 1 + para.content.size }).ok).toBe(true);
      expect(session.doc.child(0).type.name).toBe('sdtBlock');
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/name' }, { from: start + 2 + 3 }).ok).toBe(true);
      const block = session.doc.child(0);
      expect(block.childCount).toBe(1);
      expect(controls(block).map((c) => c.textContent)).toEqual(['Ann']);
      // An empty paragraph inside it (Enter at its end): a run-level control in that paragraph, never a block-level one
      // around it (Word's pane makes one: check 34, 1d; Jason's decision of 2026-10-08, docs/developer/word-divergences.md).
      const end = start + 1 + block.child(0).nodeSize - 1;
      expect(runCommand(session, 'insertParagraph', undefined, { from: end }).ok).toBe(true);
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/name' }, { from: end + 2 }).ok).toBe(true);
      const outer = session.doc.child(0);
      expect(outer.child(1).type.name).toBe('paragraph');
      expect(outer.child(1).child(0).type.name).toBe('sdtInline');
      expect(outer.child(1).textContent).toBe('Ann');
    });

    // ED-005 section 12.57 (Jason's word): with tracking on, a bind over text that was not bound is tracked as Word's
    // mapping tracks it (core-ts check 34, 1j: the old text deleted and the value inserted, inside the control); a bound
    // control's text is the data's, and rebinding it is not tracked.
    test('tracking on: the replaced text kept deleted and the value inserted inside the control; a bound control rebound untracked', async () => {
      const original = await fixture('HelloWordOnline.docx');
      const session = await withData();
      session.trackingMode = 'TrackAll';
      session.trackingAuthor = 'Reviewer';
      /** A control's text as runs of one kind: '-' deleted, '+' inserted, else plain; neighbours of a kind joined. */
      const pieces = (n: PMNode): string[] => {
        const out: string[] = [];
        let last: string | undefined;
        n.descendants((c) => {
          if (!c.isText) return true;
          const sign = c.marks.some((m) => m.type.name === 'deletion') ? '-' : c.marks.some((m) => m.type.name === 'insertion') ? '+' : '';
          if (sign === last) out[out.length - 1] += c.text!;
          else out.push(sign + c.text);
          last = sign;
          return true;
        });
        return out;
      };
      const start = blockStart(session, 0) + 1 + 2;
      session.undoManager.stopCapturing();
      const r = runCommand(session, 'contentControl.insertBound', { xpath: '/data/name' }, { from: start, to: start + 5 });
      expect(r.ok).toBe(true);
      expect(r.untracked).toBeUndefined();
      const [made] = controls(session.doc);
      expect(pieces(made!)).toEqual(['-Hello', '+Ann']);
      expect(session.doc.child(0).child(2).eq(made!)).toBe(true);   // after the bookmark's two markers
      expect(r.state!.selection.from).toBe(start + made!.nodeSize - 1);
      expect(session.template.dataPart.part!.selectSingleNode('/data/name')!.text).toBe('Ann');   // nothing written back
      // Rebinding the bound control: its text is the data's, replaced untracked.
      const rebind = runCommand(session, 'contentControl.bind', { xpath: '/data/empty' }, { from: start + 2 });
      expect(rebind.ok).toBe(true);
      expect(rebind.untracked).toBe('contentControl.bind');
      expect(pieces(controls(session.doc)[0]!)).toEqual([PLACEHOLDER_TEXT]);
      session.undoManager.undo();
      expect(pieces(controls(session.doc)[0]!)).toEqual(['-Hello', '+Ann']);
      // An unbound plain text control bound: tracked the same way.
      const plain = start + controls(session.doc)[0]!.nodeSize + 1;
      expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: plain, to: plain + 5 }).ok).toBe(true);
      const unbound = controls(session.doc)[1]!.textContent;
      expect(runCommand(session, 'contentControl.bind', { xpath: '/data/name' }, { from: plain + 2 }).ok).toBe(true);
      expect(pieces(controls(session.doc)[1]!)).toEqual([`-${unbound}`, '+Ann']);
      // At a caret: the value inserted; the placeholder for an empty node is no content, unmarked.
      const end = (): number => session.doc.child(0).nodeSize - 1;
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/name' }, { from: end() }).ok).toBe(true);
      expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/empty' }, { from: end() }).ok).toBe(true);
      expect(controls(session.doc).slice(2).map(pieces)).toEqual([['+Ann'], [PLACEHOLDER_TEXT]]);
      // Saved: the deletion and the insertion inside the control, as Word writes them; the save's self-check passes. The
      // placeholder's style is added to the styles part (ED-005 section 12.59).
      const saved = await saveAndReload(session, original, ['word/styles.xml']);
      expect(await mainXml(saved.bytes)).toMatch(/<w:sdtContent><w:del [^>]*w:author="Reviewer"[^>]*>(?:(?!<\/w:del>).)*<w:delText>Hello<\/w:delText>.*?<\/w:del><w:ins [^>]*>(?:(?!<\/w:ins>).)*<w:t>Ann<\/w:t>.*?<\/w:ins><\/w:sdtContent>/);
    });
  });

  // ED-005 section 12.55: Word's mapping both ways. Typing in a bound control writes its node, and the other controls bound
  // to that node show it; a node changed (the Data tab) is put into every control bound to it.
  describe('two-way: the control writes its node, the node its controls', () => {
    /** Two plain text controls over "Hel" and "lo", both bound to /data/name. */
    async function twoBound(): Promise<{ session: EditorDocument; node: () => string; texts: () => string[]; inFirst: number }> {
      const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
      const part = session.pkg.customXmlParts.add('<data><name>Ann</name><other>x</other></data>');
      await session.template.load();
      const start = blockStart(session, 0) + 1;
      expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: start, to: start + 3 }).ok).toBe(true);
      expect(runCommand(session, 'contentControl.bind', { xpath: '/data/name' }, { from: start + 2 }).ok).toBe(true);
      const second = start + session.doc.child(0).child(0).nodeSize;
      expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: second, to: second + 2 }).ok).toBe(true);
      expect(runCommand(session, 'contentControl.bind', { xpath: '/data/name' }, { from: second + 1 }).ok).toBe(true);
      const texts = (): string[] => { const out: string[] = []; session.doc.child(0).forEach((n) => { if (n.type.name === 'sdtInline') out.push(n.textContent); }); return out; };
      return { session, node: () => part.selectSingleNode('/data/name')!.text, texts, inFirst: start + 3 };
    }

    test('typing in one control writes the node, and the other control bound to it shows the new value, out of the undo history', async () => {
      const { session, node, texts, inFirst } = await twoBound();
      expect(texts()).toEqual(['Ann', 'Ann']);
      expect(runCommand(session, 'insertText', { text: 'e' }, { from: inFirst + 3 }).ok).toBe(true);   // "Ann|" -> "Anne"
      expect(node()).toBe('Anne');
      expect(texts()).toEqual(['Anne', 'Anne']);
      expect([...session.touchedParts].some((p) => p.startsWith('/customXml/'))).toBe(true);
      // A trailing space stays where it was typed, and the node and the other control hold it as it is.
      expect(runCommand(session, 'insertText', { text: ' ' }, { from: inFirst + 4 }).ok).toBe(true);
      expect(node()).toBe('Anne ');
      expect(texts()).toEqual(['Anne ', 'Anne ']);
      // The refresh is not an undo step: one undo takes back the typing in the first control.
      session.undoManager.undo();
      expect(texts()[0]).toBe('Ann');
    });

    test('a node changed (as the Data tab changes it) is put into every control bound to it, out of the undo history; the rest are left', async () => {
      const { session, node, texts } = await twoBound();
      const part = session.template.dataPart.part!;
      part.selectSingleNode('/data/name')!.text = 'Bob';
      session.template.invalidate();
      const steps = session.undoManager.undoStack.length;
      expect(applyBoundRefresh(session, stateFor(session))).toBe(true);
      expect(texts()).toEqual(['Bob', 'Bob']);
      expect(node()).toBe('Bob');
      expect(session.undoManager.undoStack.length).toBe(steps);
      // Nothing differs: nothing written.
      expect(applyBoundRefresh(session, stateFor(session))).toBe(false);
    });

    test('a drop-down writes its chosen item\'s value; with tracking on, deleted text is not part of the value', async () => {
      const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
      const part = session.pkg.customXmlParts.add('<data><choice>a</choice><name>Ann</name></data>');
      await session.template.load();
      const start = blockStart(session, 0) + 1;
      expect(runCommand(session, 'contentControl.insert', { kind: 'DropDownList' }, { from: start, to: start + 3 }).ok).toBe(true);
      const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const items = { name: { namespaceURI: W, localPart: 'dropDownList' }, value: { TYPE_NAME: 'org_docx4j_wml.CTSdtDropDownList', listItem: [{ TYPE_NAME: 'org_docx4j_wml.CTSdtListItem', displayText: 'Alpha', value: 'a' }, { TYPE_NAME: 'org_docx4j_wml.CTSdtListItem', displayText: 'Bee', value: 'b' }] } };
      let state = stateFor(session);
      const control = session.doc.child(0).child(0);
      session.replace(state.apply(state.tr.setNodeMarkup(start, undefined, { ...control.attrs, sdtPr: withSdtItem(control.attrs['sdtPr'], items) })).doc);
      expect(runCommand(session, 'contentControl.bind', { xpath: '/data/choice' }, { from: start + 2 }).ok).toBe(true);
      expect(session.doc.child(0).child(0).textContent).toBe('Alpha');
      // The text changed to another item's display text: the node takes that item's value.
      // The control's text follows the bookmark's two markers it holds.
      const textAt = (pos: number): number => { let at = -1; session.doc.nodeAt(pos)!.descendants((n, p) => { if (at < 0 && n.isText) at = pos + 1 + p; return at < 0; }); return at; };
      const alpha = textAt(start);
      expect(runCommand(session, 'insertText', { text: 'Bee' }, { from: alpha, to: alpha + 5 }).ok).toBe(true);
      expect(part.selectSingleNode('/data/choice')!.text).toBe('b');
      // A plain text control with tracking on: the deleted "Ann" is not part of the value, the typed "Zoe" is.
      const plainAt = start + session.doc.child(0).child(0).nodeSize;
      expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: plainAt, to: plainAt + 2 }).ok).toBe(true);
      expect(runCommand(session, 'contentControl.bind', { xpath: '/data/name' }, { from: plainAt + 1 }).ok).toBe(true);
      session.trackingMode = 'TrackAll';
      session.trackingAuthor = 'Reviewer';
      const ann = textAt(plainAt);
      expect(runCommand(session, 'insertText', { text: 'Zoe' }, { from: ann, to: ann + 3 }).ok).toBe(true);
      expect(session.doc.nodeAt(plainAt)!.textContent).toBe('AnnZoe');   // "Ann" kept, struck through, beside "Zoe"
      expect(part.selectSingleNode('/data/name')!.text).toBe('Zoe');
    });
  });

  test('a condition made in the command: the Conditions part gains it, the tag names it, the verdict evaluates it; REQ-007 and REQ-030 refuse', async () => {
    const original = await fixture('invoice.docx');
    const session = await EditorDocument.open(original);
    await session.template.load();
    const para = session.doc.child(2);
    const start = blockStart(session, 2) + 1;
    expect(runCommand(session, 'contentControl.insert', { title: 'Spam' }, { from: start, to: start + para.content.size }).ok).toBe(true);
    const inside = { from: start + 1 };
    expect(runCommand(session, 'opendope.condition', { id: 'nosuch' }, inside).ok).toBe(false);   // REQ-030
    expect(runCommand(session, 'opendope.condition', { id: 'c7', define: { op: 'not', operand: { op: 'xpathref', id: 'x6' } }, name: 'no spam' }, inside).ok).toBe(true);
    expect(session.template.conditions!.conditions[2]).toEqual({ id: 'c7', attrs: { name: 'no spam' }, operand: { op: 'not', operand: { op: 'xpathref', id: 'x6' } } });
    const sdtPr = session.doc.child(2).attrs['sdtPr'];
    expect(readControl(sdtPr).tag).toBe('od:condition=c7');
    expect(session.template.verdictFor(sdtPr)).toMatchObject({ role: 'condition', id: 'c7', kept: true });
    expect(runCommand(session, 'opendope.repeat', { id: 'x2' }, inside).ok).toBe(false);   // REQ-007: already a condition
    expect(session.template.usersOfCondition('c7')).toHaveLength(1);
    expect(() => session.template.deleteCondition('c7')).toThrow(/used by/);
    const { bytes } = await saveAndReload(session, original, ['customXml/item1.xml', 'customXml/item4.xml']);
    const reopened = await EditorDocument.open(bytes);
    await reopened.template.load();
    expect(reopened.template.conditions!.conditions.map((c) => c.id)).toEqual(['c5', 'c6', 'c7']);
    expect(reopened.template.xpaths!.authoringTool).toBe('docx4j-editor');
    expect(reopened.template.xpaths!.version).toBeUndefined();   // a 2.3 template stays one (REQ-009's note)
    // Clearing the role leaves the definition for the panel to delete.
    expect(runCommand(session, 'opendope.condition', { id: null }, inside).ok).toBe(true);
    expect(readControl(session.doc.child(2).attrs['sdtPr']).tag).toBe('');
    expect(session.template.usersOfCondition('c7')).toHaveLength(0);
    session.template.deleteCondition('c7');
    expect(session.template.conditions!.conditions).toHaveLength(2);
  });

  test('a repeat on wrapped rows with its entry defined; a position condition; REQ-008 refuses a role on a bound control', async () => {
    const original = await fixture('invoice.docx');
    const session = await EditorDocument.open(original);
    await session.template.load();
    const headerRow = cellStart(session, 5, 0, 0);
    expect(runCommand(session, 'contentControl.wrapRows', undefined, { from: headerRow }).ok).toBe(true);
    expect(runCommand(session, 'opendope.repeat', { id: 'items', define: { xpath: '/invoice[1]/items/item' } }, { from: headerRow }).ok).toBe(true);
    const groupId = (session.doc.child(5).child(0).attrs['sdt'] as SdtGroup[])[0]!.groupId;
    const entry = readControlEntry(session.store, groupId)!;
    expect(readControl(entry.sdtPr).tag).toBe('od:repeat=items');
    expect(session.template.xpaths!.entries.find((e) => e.id === 'items')).toMatchObject({ xpath: '/invoice[1]/items/item', storeItemID: session.template.dataPart.part!.id });
    expect(session.template.verdictFor(entry.sdtPr)).toMatchObject({ role: 'repeat', id: 'items', count: 3 });
    expect(session.template.usersOfEntry('items')).toEqual([`the row control ${groupId}`]);
    // A position condition inside it, on the header row's cell control made now.
    expect(runCommand(session, 'contentControl.wrapCells', undefined, { from: headerRow }).ok).toBe(true);
    expect(runCommand(session, 'opendope.positionCondition', { id: 'notLast', define: { xpath: 'position()<last()' } }, { from: headerRow }).ok).toBe(true);
    const cellGroup = (session.doc.child(5).child(0).child(0).attrs['sdt'] as SdtGroup[])[0]!.groupId;
    expect(session.template.verdictFor(readControlEntry(session.store, cellGroup)!.sdtPr)).toMatchObject({ role: 'positionCondition', positionPredicate: 'position()<last()' });
    // A bound control (the customer name) cannot take a condition (REQ-008).
    expect(runCommand(session, 'opendope.condition', { id: 'c5' }, { from: blockStart(session, 0) + 2 }).ok).toBe(false);
    const { api } = await saveAndReload(session, original, ['customXml/item4.xml']);
    expect(api.contentControls.filter((c) => c.form === 'Row').map((c) => c.tag)).toEqual(['od:repeat=items', 'od:repeat=x2&od:finish=t_r']);
  });

  test('verdicts on the fixtures: bound values, a repeat count, conditions in the java mode, the w15 repeating section\'s count, an XPath 2 expression through the FontoXPath engine', async () => {
    const invoice = await EditorDocument.open(await fixture('invoice.docx'));
    await invoice.template.load();
    const verdicts = listControls(invoice.doc, invoice.controlEntry).map((c) => invoice.template.verdictFor(c.sdtPr));
    expect(verdicts).toEqual([
      { role: 'none', warnings: [], value: 'Joe Bloggs' },
      { role: 'repeat', warnings: [], id: 'x2', count: 3 },
      { role: 'none', warnings: [], value: 'apples' },
      { role: 'none', warnings: [], value: '$20' },
      { role: 'condition', warnings: [], id: 'c5', kept: true },
      { role: 'condition', warnings: [], id: 'c6', kept: false },
    ]);
    const w15 = await EditorDocument.open(await fixture('invoice2013.docx'));
    await w15.template.load();
    const sections = listControls(w15.doc, w15.controlEntry).filter((c) => c.reading.kind === 'RepeatingSection').map((c) => w15.template.verdictFor(c.sdtPr));
    expect(sections.map((v) => v.role)).toEqual(['repeatingSection', 'repeatingSection']);
    expect(sections.every((v) => typeof v.count === 'number' && v.count > 0)).toBe(true);
    const saxon = await EditorDocument.open(await fixture('opendope/invoice_Saxon_XPath2.docx'));
    await saxon.template.load();
    // `xs:date(/invoice/date) > xs:date('2018-12-31')`: XPath 2 syntax, evaluated by the engine the template view installs (2019-01-15 in the data).
    const dateGt = listControls(saxon.doc, saxon.controlEntry).map((c) => saxon.template.verdictFor(c.sdtPr)).find((v) => v.id === 'dateGtCondition')!;
    expect(dateGt).toMatchObject({ kept: true });
    expect(saxon.pkg.xpathEngine.constructor.name).toBe('FontoXPathEngine');
    // The module handed over, as the application's Template chunk does (`new FontoXPathEngine(fontoxpath)`,
    // core-ts 0.1.4; 0.1.3 ignores the argument and loads the package itself, so this proves the wiring only from 0.1.4 on).
    const handed = await EditorDocument.open(await fixture('opendope/invoice_Saxon_XPath2.docx'));
    await handed.template.load({ fontoxpath: await import('fontoxpath') });
    expect(handed.template.engineLoaded).toBe(true);
    expect(listControls(handed.doc, handed.controlEntry).map((c) => handed.template.verdictFor(c.sdtPr)).find((v) => v.id === 'dateGtCondition')).toMatchObject({ kept: true });
  });
});

describe('through a tree run', () => {
  test('a control made in the editor becomes an element with an id of its own after runOnTree; the map follows', async () => {
    const original = await fixture('tables.docx');
    const session = await EditorDocument.open(original);
    const doc = session.doc;
    const tablePos = blockStart(session, 1);
    const map = TableMap.get(doc.child(1));
    expect(runCommand(session, 'contentControl.wrapRows', { title: 'G' }, CellSelection.create(doc, tablePos + 1 + map.map[0]!, tablePos + 1 + map.map[map.width]!)).ok).toBe(true);
    const before = (session.doc.child(1).child(0).attrs['sdt'] as SdtGroup[])[0]!.groupId;
    expect(before.startsWith('new:')).toBe(true);
    await session.runOnTree(({ body }) => { body.paragraphs[0]!.insertText('!', 'End'); });
    const after = (session.doc.child(1).child(0).attrs['sdt'] as SdtGroup[])[0]!.groupId;
    expect(after.startsWith('new:')).toBe(false);
    expect(readControlEntry(session.store, before)).toBeUndefined();
    expect(sdtTitle(readControlEntry(session.store, after)!.sdtPr)).toBe('G');
    expect(typeNameOf(session.ctx.ids.get(after) as Element)).toBe('org_docx4j_wml.CTSdtRow');
    const saved = await session.save();
    expect(comparePackages(original, saved.bytes, [session.main.partName.storeName, 'docProps/app.xml'])).toEqual([]);
  });
});
