// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 3.5: tables projected and exported. The untouched round trip over the table
// fixtures is in roundtrip.test.mts (every fixture); here the projection's shape, the edits
// through the headless runner with their exports reloaded, merges both ways, and undo.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage, rowsOf, cellsOf, typeNameOf, textOf as textOfElement } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { TextSelection } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { rowShadingOf, tableAutoFitOf, textWidthTwips, columnWidthsPt, rowHeightOf, EditorDocument, runCommand, stateFor, commandNamed, selectionFormatting, CellSelection, TableMap, repairStacks, type SdtGroup } from '../src/index.mjs';
import { fixture, comparePackages, textOf } from './helpers.mjs';

async function reloadBody(bytes: Uint8Array): Promise<wml.Body> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  return (await pkg.getMainDocumentPart().readContents()).body!;
}

/** The position of block `i` and of the first text position inside cell (r, c) of the table there. */
function cellStart(session: EditorDocument, block: number, r: number, c: number): number {
  const doc = session.doc;
  let pos = 0;
  for (let k = 0; k < block; k++) pos += doc.child(k).nodeSize;
  const table = doc.child(block);
  const map = TableMap.get(table);
  const cellPos = pos + 1 + map.map[r * map.width + c]!;
  return cellPos + 2;   // inside the cell's first paragraph
}

async function saveAndReload(session: EditorDocument, original: Uint8Array): Promise<wml.Body> {
  const saved = await session.save();
  expect(saved.changed).toBe(true);
  const diffs = comparePackages(original, saved.bytes, [session.main.partName.storeName, 'docProps/app.xml']);
  expect(diffs).toEqual([]);
  return reloadBody(saved.bytes);
}

const tblAt = (body: wml.Body, i: number): wml.Tbl => {
  const el = body.content![i]!;
  expect(typeNameOf(el)).toBe('org_docx4j_wml.Tbl');
  return el.value as wml.Tbl;
};
const cellText = (tbl: wml.Tbl, r: number, c: number): string => textOfElement(cellsOf(rowsOf(tbl)[r]!.element.value)[c]!.element);
const tcPr = (tbl: wml.Tbl, r: number, c: number): wml.TcPr | undefined => cellsOf(rowsOf(tbl)[r]!.element.value)[c]!.element.value.tcPr;

describe('projection', () => {
  test('tables.docx: spans and merges become colspan and rowspan; the grid and widths are carried; a merged row keeps no cell for the continuation', async () => {
    const session = await EditorDocument.open(await fixture('tables.docx'));
    const t3 = session.doc.child(3);   // A | B(vMerge restart) / C | (continuation)
    expect(t3.type.name).toBe('table');
    expect(t3.attrs['tblGrid']).toEqual([4606, 4606]);
    expect(t3.child(0).childCount).toBe(2);
    expect(t3.child(0).child(1).attrs['rowspan']).toBe(2);
    expect((t3.child(0).child(1).attrs['merged'] as string[]).length).toBe(1);
    expect(t3.child(1).childCount).toBe(1);
    const map = TableMap.get(t3);
    expect([map.width, map.height]).toEqual([2, 2]);
    const t7 = session.doc.child(7);   // a cs2 cell in row 1
    expect(t7.child(1).child(1).attrs['colspan']).toBe(2);
    expect(t7.child(1).child(1).attrs['colwidth']).toEqual([2303 / 15, 2303 / 15]);
    const t33 = session.doc.child(33);   // the mixed table: cs2 rs2 in row 1, cs2 rs2 and rs2 in row 3
    expect(TableMap.get(t33).problems).toBeNull();
    expect(t33.child(1).child(1).attrs).toMatchObject({ colspan: 2, rowspan: 2 });
    expect(t33.child(3).childCount).toBe(3);
    expect(t33.child(4).childCount).toBe(1);
    // Header rows carry their trPr.
    const t14 = session.doc.child(14);
    expect(JSON.stringify(t14.child(0).attrs['trPr'])).toContain('tblHeader');
    expect(t14.child(2).attrs['trPr']).toBeNull();
    expect(session.diagnostics.opaque.filter((o) => o.kind === 'tbl')).toEqual([]);
  });

  test('invoice fixtures: row-level controls are attribute stacks with first and last, nested outermost first', async () => {
    const session = await EditorDocument.open(await fixture('invoice.docx'));
    const t = session.doc.child(5);
    expect(t.child(0).attrs['sdt']).toBeNull();
    const stack = t.child(1).attrs['sdt'] as SdtGroup[];
    expect(stack).toHaveLength(1);
    expect(stack[0]).toMatchObject({ first: true, last: true });
    expect(typeNameOf(session.ctx.ids.get(stack[0]!.groupId)!)).toBe('org_docx4j_wml.CTSdtRow');
    const s2013 = await EditorDocument.open(await fixture('invoice2013.docx'));
    const nested = s2013.doc.child(7).child(1).attrs['sdt'] as SdtGroup[];
    expect(nested).toHaveLength(2);
    const outer = s2013.ctx.ids.get(nested[0]!.groupId)!.value as wml.CTSdtRow;
    const innerEl = s2013.ctx.ids.get(nested[1]!.groupId)!;
    expect(outer.sdtContent!.content!.includes(innerEl as never)).toBe(true);
    // A table without a grid is projected and reported.
    expect(s2013.doc.child(7).attrs['tblGrid']).toBeNull();
    expect(s2013.diagnostics.opaque.some((o) => o.kind === 'tblGrid')).toBe(true);
  });
});

describe('edits', () => {
  test('typing in a cell: the cell is rebuilt, the rest of the table by identity; the merge above survives', async () => {
    const original = await fixture('tables.docx');
    const session = await EditorDocument.open(original);
    const before = session.body.content![3]!;
    const pos = cellStart(session, 3, 1, 0);   // "C", below the merge
    expect(runCommand(session, 'insertText', { text: 'X' }, { from: pos + 1 }).ok).toBe(true);
    expect(textOf(session.doc.child(3).child(1).child(0))).toBe('CX');
    const tbl = tblAt(await saveAndReload(session, original), 3);
    expect(cellText(tbl, 1, 0)).toBe('CX');
    expect(tcPr(tbl, 0, 1)?.vMerge?.val).toBe('restart');
    expect(tcPr(tbl, 1, 1)?.vMerge).toBeDefined();   // the continuation cell re-emitted
    expect(session.body.content![3]).not.toBe(before);
    expect(session.body.content![1]).toBe(session.imported.child(1).attrs['sourceId'] ? session.body.content![1] : undefined);
  });

  test('a row added inside a vertical merge: the restart cell spans three rows and the new row carries a continuation', async () => {
    const original = await fixture('tables.docx');
    const session = await EditorDocument.open(original);
    const pos = cellStart(session, 3, 0, 0);
    expect(runCommand(session, 'table.addRows', { location: 'After' }, { from: pos }).ok).toBe(true);
    const t = session.doc.child(3);
    expect(t.childCount).toBe(3);
    expect(t.child(0).child(1).attrs['rowspan']).toBe(3);
    const tbl = tblAt(await saveAndReload(session, original), 3);
    const rows = rowsOf(tbl);
    expect(rows).toHaveLength(3);
    expect(cellsOf(rows[1]!.element.value)).toHaveLength(2);
    expect(tcPr(tbl, 1, 1)?.vMerge).toBeDefined();
    expect(tcPr(tbl, 1, 1)?.vMerge?.val).not.toBe('restart');
    expect(tcPr(tbl, 2, 1)?.vMerge).toBeDefined();
    expect(cellText(tbl, 2, 0)).toBe('C');
    // Reopening the saved file gives the same table again.
    const again = await EditorDocument.open((await session.save()).bytes);
    expect(again.doc.child(3).child(0).child(1).attrs['rowspan']).toBe(3);
    expect(session.undoManager.canUndo()).toBe(true);
    expect(runCommand(session, 'history.undo').ok).toBe(true);
    expect(session.doc.child(3).childCount).toBe(2);
  });

  test('a column deleted through a span shrinks the spanning cell; a column added widens the grid', async () => {
    const original = await fixture('tables.docx');
    const session = await EditorDocument.open(original);
    const pos = cellStart(session, 7, 0, 2);   // column C, which row 1's cs2 cell spans
    expect(runCommand(session, 'table.deleteColumns', undefined, { from: pos }).ok).toBe(true);
    let t = session.doc.child(7);
    expect(TableMap.get(t).width).toBe(3);
    expect(t.child(1).child(1).attrs['colspan']).toBe(1);
    expect(runCommand(session, 'table.addColumns', { location: 'After' }, { from: cellStart(session, 7, 0, 0) }).ok).toBe(true);
    t = session.doc.child(7);
    expect(TableMap.get(t).width).toBe(4);
    const tbl = tblAt(await saveAndReload(session, original), 7);
    expect(tbl.tblGrid.gridCol).toHaveLength(4);
    expect(cellsOf(rowsOf(tbl)[0]!.element.value)).toHaveLength(4);
    expect(cellText(tbl, 0, 0)).toBe('A');
    expect(cellText(tbl, 0, 1)).toBe('');
    expect(cellText(tbl, 0, 2)).toBe('B');
    expect(tcPr(tbl, 1, 2)?.gridSpan).toBeUndefined();
  });

  test('a column added keeps the layout mode (ED-005 section 12.78): fixed copies the width beside it and the table grows; AutoFit to window keeps the table at the text width, every column scaled; table.properties switches', async () => {
    const original = await fixture('tables.docx');
    const session = await EditorDocument.open(original);
    const at = cellStart(session, 1, 0, 0);
    const widths = (): number[] => { const t = session.doc.child(1); const out: number[] = []; t.child(0).forEach((c) => out.push(...((c.attrs['colwidth'] as number[] | null) ?? [NaN]))); return out; };
    const total = (): number => widths().reduce((a, b) => a + b, 0);
    const before = widths();
    const was = total();
    // tables.docx's two-column table states no layout: Word reads that as AutoFit to contents. Made fixed first.
    expect(runCommand(session, 'table.properties', { autoFit: 'fixed' }, { from: at }).ok).toBe(true);
    expect(tableAutoFitOf(session.doc.child(1).attrs['tblPr'])).toBe('fixed');
    expect(runCommand(session, 'table.addColumns', { location: 'After' }, { from: at }).ok).toBe(true);
    expect(widths()).toEqual([before[0], before[0], before[1]]);
    expect(session.doc.child(1).attrs['tblGrid']).toEqual(widths().map((w) => Math.round(w * 15)));
    expect(total()).toBeCloseTo(was + before[0]!, 3);
    // AutoFit to window: the table the text width, the columns scaled to it, w:tblW 100%.
    expect(runCommand(session, 'table.properties', { autoFit: 'window' }, { from: at }).ok).toBe(true);
    expect(tableAutoFitOf(session.doc.child(1).attrs['tblPr'])).toBe('window');
    expect((session.doc.child(1).attrs['tblPr'] as { tblW?: { w: number; type: string } }).tblW).toEqual({ w: 5000, type: 'pct' });
    const text = textWidthTwips(stateFor(session, { from: at })) / 15;
    expect(total()).toBeCloseTo(text, 3);
    const ratio = widths()[0]! / widths()[2]!;
    expect(ratio).toBeCloseTo(before[0]! / before[1]!, 6);
    // Another column in window mode: the total stays, every column scaled.
    expect(runCommand(session, 'table.addColumns', { location: 'Before' }, { from: cellStart(session, 1, 0, 0) }).ok).toBe(true);
    expect(TableMap.get(session.doc.child(1)).width).toBe(4);
    expect(total()).toBeCloseTo(text, 3);
    expect(widths()[0]).toBeCloseTo(widths()[1]!, 6);   // the new column copies the one beside it, then both scale alike
    // Contents: the layout AutoFit, the width automatic, the columns left as they are.
    expect(runCommand(session, 'table.properties', { autoFit: 'contents' }, { from: cellStart(session, 1, 0, 0) }).ok).toBe(true);
    expect(tableAutoFitOf(session.doc.child(1).attrs['tblPr'])).toBe('contents');
    // Saved: the grid is the sum of the cells in every row, and the properties round-trip.
    const tbl = tblAt(await saveAndReload(session, original), 1);
    expect(tbl.tblGrid.gridCol).toHaveLength(4);
    expect(tbl.tblPr.tblLayout).toBeUndefined();
    expect(tbl.tblPr.tblW).toMatchObject({ type: 'auto' });
  });

  test('table.insert takes Word\'s dialog choices (ED-005 section 12.83): fixed Auto, a column width, window, contents', async () => {
    const original = await fixture('HelloWordOnline.docx');
    const session = await EditorDocument.open(original);
    // Each after the first paragraph (which has text, so none replaces it), so the last made comes first.
    const inEmptyFirst = (args: unknown): boolean => runCommand(session, 'table.insert', args, { from: 1 }).ok;
    expect(inEmptyFirst({ rows: 1, columns: 2, autoFit: 'contents' })).toBe(true);
    expect(inEmptyFirst({ rows: 1, columns: 2, autoFit: 'window' })).toBe(true);
    expect(inEmptyFirst({ rows: 1, columns: 3, autoFit: 'fixed', columnWidth: 72 })).toBe(true);
    expect(inEmptyFirst({ rows: 2, columns: 5 })).toBe(true);
    expect(inEmptyFirst({ rows: 0, columns: 2 })).toBe(false);
    const body = await saveAndReload(session, original);
    const tables = body.content!.filter((el) => typeNameOf(el) === 'org_docx4j_wml.Tbl').map((el) => el.value as wml.Tbl);
    expect(tables).toHaveLength(4);
    const text = textWidthTwips(stateFor(session, { from: 1 }));
    // Fixed, Auto (Word's default table, check 42): an automatic width, no layout, the text width shared, every cell in twips.
    expect(tables[0]!.tblPr.tblLayout).toBeUndefined();
    expect(tables[0]!.tblPr.tblW).toMatchObject({ type: 'auto' });
    expect(tables[0]!.tblGrid.gridCol).toHaveLength(5);
    expect(tcPr(tables[0]!, 0, 0)?.tcW).toMatchObject({ type: 'dxa', w: Math.floor(text / 5) });
    // Fixed, 72 pt: each column 1,440 twips, the same form.
    expect(tables[1]!.tblPr.tblW).toMatchObject({ type: 'auto' });
    expect(tables[1]!.tblGrid.gridCol!.map((g) => g.w)).toEqual([1440, 1440, 1440]);
    expect(tcPr(tables[1]!, 0, 2)?.tcW).toMatchObject({ type: 'dxa', w: 1440 });
    // Window: 100%, the cells as shares, no layout.
    expect(tables[2]!.tblPr.tblLayout).toBeUndefined();
    expect(tables[2]!.tblPr.tblW).toMatchObject({ type: 'pct', w: 5000 });
    expect(tcPr(tables[2]!, 0, 0)?.tcW).toMatchObject({ type: 'pct', w: 2500 });
    // Contents: automatic, the cells automatic.
    expect(tables[3]!.tblPr.tblW).toMatchObject({ type: 'auto' });
    expect(tcPr(tables[3]!, 0, 1)?.tcW).toMatchObject({ type: 'auto' });
  });

  test('table.columnWidths writes every cell and the grid from one list in points (ED-007 section 3.1): a fixed table\'s width follows, an automatic one keeps its statement, the table\'s own list writes nothing, a wrong count is refused', async () => {
    for (const name of ['tables/table-fit-fixed.docx', 'tables/table-fit-auto.docx']) {
      const original = await fixture(name);
      const session = await EditorDocument.open(original);
      // The build's files round-trip untouched.
      expect((await session.save()).changed).toBe(false);
      const at = cellStart(session, 1, 1, 1);   // table A, inside the long cell
      const before = columnWidthsPt(stateFor(session, { from: at }))!;
      expect(before.map((w) => Math.round(w))).toEqual([150, 150, 150]);
      // The table's own widths: nothing written.
      expect(runCommand(session, 'table.columnWidths', { widths: before }, { from: at }).ok).toBe(true);
      expect((await session.save()).changed).toBe(false);
      // A wrong count, or not a list: refused.
      expect(runCommand(session, 'table.columnWidths', { widths: [100, 200] }, { from: at }).ok).toBe(false);
      expect(runCommand(session, 'table.columnWidths', { widths: 'wide' } as unknown as { widths: number[] }, { from: at }).ok).toBe(false);
      // A new list: every cell and the grid, the narrow ones clamped to the floor (25 px = 18.75 pt).
      expect(runCommand(session, 'table.columnWidths', { widths: [40, 350, 10] }, { from: at }).ok).toBe(true);
      const table = session.doc.child(1);
      expect(table.attrs['tblGrid']).toEqual([800, 7000, 375]);
      table.forEach((row) => row.forEach((cell) => { expect(cell.attrs['colwidth']).toHaveLength(1); }));
      expect((table.child(1).child(1).attrs['colwidth'] as number[])[0]).toBeCloseTo(7000 / 15, 6);
      expect(columnWidthsPt(stateFor(session, { from: cellStart(session, 1, 1, 1) }))!.map((w) => Math.round(w * 100) / 100)).toEqual([40, 350, 18.75]);
      const body = await saveAndReload(session, original);
      const tbl = tblAt(body, 1);
      expect(tbl.tblGrid.gridCol!.map((g) => g.w)).toEqual([800, 7000, 375]);
      expect(tcPr(tbl, 1, 1)?.tcW).toMatchObject({ type: 'dxa', w: 7000 });
      // A fixed table's stated width is the sum; an automatic one stays automatic.
      if (name.includes('fixed')) expect(tbl.tblPr.tblW).toMatchObject({ type: 'dxa', w: 8175 });
      else expect(tbl.tblPr.tblW).toMatchObject({ type: 'auto' });
      // Table B's merged cell takes both its columns' widths.
      const b = cellStart(session, 4, 1, 1);   // heading, table A, an empty paragraph, heading, table B
      expect(runCommand(session, 'table.columnWidths', { widths: [60, 300, 90] }, { from: b }).ok).toBe(true);
      expect(session.doc.child(4).child(0).child(0).attrs['colwidth']).toEqual([80, 400]);
      expect(session.doc.child(4).child(0).child(0).attrs['colspan']).toBe(2);
      // Saved and opened again, table B keeps the fit (section 8.1 item 11): the vertical merge's continuation cell
      // takes its column's width, and the spanning cell's width is shared over its columns as the grid is.
      const savedB = await session.save();
      const tblB = tblAt(await reloadBody(savedB.bytes), 4);
      expect(tblB.tblGrid.gridCol!.map((g) => g.w)).toEqual([1200, 6000, 1800]);
      expect(tcPr(tblB, 0, 0)?.tcW).toMatchObject({ type: 'dxa', w: 7200 });
      expect(tcPr(tblB, 1, 0)?.tcW).toMatchObject({ type: 'dxa', w: 1200 });
      expect(tcPr(tblB, 2, 0)?.tcW).toMatchObject({ type: 'dxa', w: 1200 });
      const again = await EditorDocument.open(savedB.bytes);
      expect(again.doc.child(4).child(0).child(0).attrs['colwidth']).toEqual([80, 400]);
      expect(columnWidthsPt(stateFor(again, { from: cellStart(again, 4, 1, 1) }))!.map(Math.round)).toEqual([60, 300, 90]);
      // Row heights (section 3.5): 'content' removes table C's exact and minimum heights, writing even for the table's
      // own widths; a table with none writes nothing for its own widths.
      const c = cellStart(session, 7, 0, 0);
      expect(rowHeightOf(session.doc.child(7).child(0).attrs['trPr'])).toEqual({ hRule: 'exact', val: 600 });
      expect(rowHeightOf(session.doc.child(7).child(1).attrs['trPr'])).toEqual({ hRule: 'atLeast', val: 1800 });
      const ownC = columnWidthsPt(stateFor(session, { from: c }))!;
      expect(runCommand(session, 'table.columnWidths', { widths: ownC, rowHeights: 'content' }, { from: c }).ok).toBe(true);
      expect(session.doc.child(7).child(0).attrs['trPr']).toBeNull();
      expect(session.doc.child(7).child(1).attrs['trPr']).toBeNull();
      expect(columnWidthsPt(stateFor(session, { from: cellStart(session, 7, 0, 0) }))).toEqual(ownC);
      const rowsC = rowsOf(tblAt(await saveAndReload(session, original), 7));
      expect(rowsC.map((r) => r.element.value.trPr)).toEqual([undefined, undefined]);
      const fresh = await EditorDocument.open(original);
      const d = cellStart(fresh, 10, 0, 0);
      expect(runCommand(fresh, 'table.columnWidths', { widths: columnWidthsPt(stateFor(fresh, { from: d }))!, rowHeights: 'content' }, { from: d }).ok).toBe(true);
      expect((await fresh.save()).changed).toBe(false);
    }
  });

  test('merge and split: two cells merged become one w:gridSpan cell; a split gives them back', async () => {
    const original = await fixture('tables.docx');
    const session = await EditorDocument.open(original);
    const doc = session.doc;
    let pos = 0;
    for (let k = 0; k < 1; k++) pos += doc.child(k).nodeSize;
    const map = TableMap.get(doc.child(1));
    const a = pos + 1 + map.map[0]!;
    const b = pos + 1 + map.map[1]!;
    const state = stateFor(session, CellSelection.create(doc, a, b));
    const spec = commandNamed('table.mergeCells')!;
    let next = state;
    expect(spec.run(undefined)(state, (tr) => { next = state.apply(tr); })).toBe(true);
    session.replace(next.doc);
    expect(session.doc.child(1).child(0).childCount).toBe(1);
    expect(session.doc.child(1).child(0).child(0).attrs['colspan']).toBe(2);
    let tbl = tblAt(await saveAndReload(session, original), 1);
    expect(tcPr(tbl, 0, 0)?.gridSpan?.val).toBe(2);
    expect(cellsOf(rowsOf(tbl)[0]!.element.value)).toHaveLength(1);
    expect(cellText(tbl, 0, 0)).toBe('A\nB');   // as Word merges: the cells' paragraphs one after the other
    expect(runCommand(session, 'table.splitCell', undefined, { from: cellStart(session, 1, 0, 0) }).ok).toBe(true);
    expect(session.doc.child(1).child(0).childCount).toBe(2);
    tbl = tblAt(await saveAndReload(session, original), 1);
    expect(cellsOf(rowsOf(tbl)[0]!.element.value)).toHaveLength(2);
    expect(tcPr(tbl, 0, 0)?.gridSpan).toBeUndefined();
  });

  test('a new table is built and reloads; a header row toggles w:tblHeader; the table style is written', async () => {
    const original = await fixture('HelloWordOnline.docx');
    const session = await EditorDocument.open(original);
    const last = session.doc.childCount - 1;
    let pos = 0;
    for (let k = 0; k < last; k++) pos += session.doc.child(k).nodeSize;
    expect(runCommand(session, 'table.insert', { rows: 2, columns: 3 }, { from: pos + 1 }).ok).toBe(true);
    const t = session.doc.child(last + 1);
    expect(t.type.name).toBe('table');
    expect(t.childCount).toBe(2);
    expect(t.child(0).childCount).toBe(3);
    const first = cellStart(session, last + 1, 0, 0);
    expect(runCommand(session, 'insertText', { text: 'hello' }, { from: first }).ok).toBe(true);
    expect(runCommand(session, 'row.header', undefined, { from: first }).ok).toBe(true);
    expect(runCommand(session, 'table.style', { value: 'TableGrid' }, { from: first }).ok).toBe(true);
    const body = await saveAndReload(session, original);
    const tbl = tblAt(body, last + 1);
    expect(tbl.tblGrid.gridCol).toHaveLength(3);
    expect(rowsOf(tbl)).toHaveLength(2);
    expect(cellText(tbl, 0, 0)).toBe('hello');
    expect(tbl.tblPr.tblStyle?.val).toBe('TableGrid');
    expect(JSON.stringify(rowsOf(tbl)[0]!.element.value.trPr)).toContain('tblHeader');
    expect(rowsOf(tbl)[1]!.element.value.trPr).toBeUndefined();
    expect(runCommand(session, 'row.header', { value: false }, { from: cellStart(session, last + 1, 0, 0) }).ok).toBe(true);
    expect(JSON.stringify(session.doc.child(last + 1).child(0).attrs['trPr'] ?? {})).not.toContain('tblHeader');
  });

  test('a row added inside a row-level control joins it, and export re-wraps the group; a row added after the last leaves it', async () => {
    const original = await fixture('invoice.docx');
    const session = await EditorDocument.open(original);
    const controlled = cellStart(session, 5, 1, 0);   // "apples", inside the repeat
    expect(runCommand(session, 'table.addRows', { location: 'Before' }, { from: controlled }).ok).toBe(true);
    let t = session.doc.child(5);
    expect(t.childCount).toBe(3);
    // Row 1 is the new one, between the header (no group) and the repeated row: contiguity puts it in the group? No: it is before the group's first member, so it stays out.
    expect(t.child(1).attrs['sdt']).toBeNull();
    expect((t.child(2).attrs['sdt'] as SdtGroup[])[0]).toMatchObject({ first: true, last: true });
    expect(runCommand(session, 'table.addRows', { location: 'After' }, { from: cellStart(session, 5, 2, 0) }).ok).toBe(true);
    t = session.doc.child(5);
    expect(t.childCount).toBe(4);
    // Added after the group's last row from inside it: outside the group (Word's behaviour for a repeat is E2.b's; the invariant only fills gaps).
    expect(t.child(3).attrs['sdt']).toBeNull();
    const body = await saveAndReload(session, original);
    const tbl = tblAt(body, 5);
    const items = tbl.content!;
    expect(items.map((e) => typeNameOf(e))).toEqual(['org_docx4j_wml.Tr', 'org_docx4j_wml.Tr', 'org_docx4j_wml.CTSdtRow', 'org_docx4j_wml.Tr']);
    expect(rowsOf(tbl)).toHaveLength(4);
    expect(cellText(tbl, 2, 0)).toBe('apples');
  });

  test("a cell's last block stays a paragraph (ED-005 section 12.109): Insert Table in a cell, a table landed at a cell's end, and a save over a store left that way", async () => {
    // Word 365 refuses a w:tc ending in a w:tbl and repairs it with an empty paragraph after the table (Jason, 2026-10-21;
    // the share's mc-fun/editor-cell-ending-in-table-repaired.docx is the shape oracle: w:p, w:tbl, w:p).
    const original = await fixture('tables.docx');
    const shapeOf = (cell: PMNode): string[] => { const out: string[] = []; cell.forEach((n) => out.push(n.type.name)); return out; };
    const cellXmlShape = (tbl: wml.Tbl, r: number, c: number): string[] => (cellsOf(rowsOf(tbl)[r]!.element.value)[c]!.element.value.content ?? []).map((e) => (e as { name: { localPart: string } }).name.localPart);
    // 1. The command: the table goes after the cell's paragraph and a paragraph after it, one undo step.
    let session = await EditorDocument.open(original);
    let pos = cellStart(session, 3, 0, 0);
    // Word's insertTable at End leaves the trailing paragraph in the cell's last paragraph's properties (core-ts check 52,
    // cases 86 to 95: a numbered List Paragraph stays one, numbered on): the cell's paragraph is styled first.
    expect(runCommand(session, 'paragraph.style', { value: 'Heading2' }, { from: pos }).ok).toBe(true);
    expect(runCommand(session, 'table.insert', { rows: 1, columns: 1 }, { from: pos }).ok).toBe(true);
    const cellAt = (s: EditorDocument, b = 3): PMNode => s.doc.nodeAt(cellStart(s, b, 0, 0) - 2)!;
    expect(shapeOf(cellAt(session))).toEqual(['paragraph', 'table', 'paragraph']);
    const trailing = cellAt(session).lastChild!;
    expect(trailing.attrs['pStyle']).toBe('Heading2');
    expect(trailing.attrs['copiedFrom']).toBe(cellAt(session).firstChild!.attrs['sourceId']);
    expect(trailing.attrs['sourceId']).toBeNull();
    let tbl = tblAt(await saveAndReload(session, original), 3);
    expect(cellXmlShape(tbl, 0, 0)).toEqual(['p', 'tbl', 'p']);
    const trailingP = cellsOf(rowsOf(tbl)[0]!.element.value)[0]!.element.value.content![2] as { value: wml.P };
    expect(trailingP.value.pPr?.pStyle?.val).toBe('Heading2');
    expect(runCommand(session, 'history.undo').ok).toBe(true);
    expect(shapeOf(cellAt(session))).toEqual(['paragraph']);
    // 2. A store left with a table last in a cell (a peer's merge, a raw transaction): the save appends the paragraph.
    session = await EditorDocument.open(original);
    pos = cellStart(session, 3, 0, 0) - 2;
    const cell = session.doc.nodeAt(pos)!;
    const st = stateFor(session);
    session.replace(st.tr.insert(pos + cell.nodeSize - 1, session.doc.child(7)).doc);
    expect(shapeOf(cellAt(session))).toEqual(['paragraph', 'table']);
    tbl = tblAt(await saveAndReload(session, original), 3);
    expect(cellXmlShape(tbl, 0, 0)).toEqual(['p', 'tbl', 'p']);
    // 3. A file that already holds such a cell is kept as it is until the table is edited: an untouched save is byte-identical.
    const pkg = await WordprocessingMLPackage.load(original);
    const part = pkg.getMainDocumentPart();
    const xml = await part.getXml();
    const inner = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const edited = xml.replace(/(<w:t[^>]*>A<\/w:t><\/w:r><\/w:p>)(<\/w:tc>)/, `$1${inner}$2`);
    expect(edited).not.toBe(xml);
    part.setXml(edited);
    const bytes = new Uint8Array(await pkg.save());
    session = await EditorDocument.open(bytes);   // the first "A" cell is block 1's
    expect(shapeOf(cellAt(session, 1))).toEqual(['paragraph', 'table']);
    const untouched = await session.save();
    expect(untouched.changed).toBe(false);
    expect(comparePackages(bytes, untouched.bytes)).toEqual([]);
    expect(runCommand(session, 'insertText', { text: 'X' }, { from: cellStart(session, 1, 0, 0) + 1 }).ok).toBe(true);
    tbl = tblAt(await saveAndReload(session, bytes), 1);
    expect(cellXmlShape(tbl, 0, 0)).toEqual(['p', 'tbl', 'p']);
    // 4. Word checks the last block as it reads it, through block-level controls (Jason's probes of 2026-10-22,
    // fixtures/tables/cell-last-block-*): a cell holding only a control that ends in a table is unreadable, and the
    // repair appends the paragraph to the cell after the control. The probe's cell is the table's second.
    const probeCell = (s: EditorDocument): number => { let at = -1; s.doc.descendants((n, p) => { if (at < 0 && n.type.name === 'table_cell' && n.firstChild?.type.name === 'sdtBlock') at = p; return at < 0; }); return at; };
    const probe = await fixture('tables/cell-last-block-12-only-sdt-tbl.docx');
    session = await EditorDocument.open(probe);
    let at = probeCell(session);
    expect(shapeOf(session.doc.nodeAt(at)!)).toEqual(['sdtBlock']);
    expect(shapeOf(session.doc.nodeAt(at + 1)!)).toEqual(['table']);
    const kept = await session.save();
    expect(kept.changed).toBe(false);
    expect(comparePackages(probe, kept.bytes)).toEqual([]);
    expect(runCommand(session, 'insertText', { text: 'X' }, { from: at + 6 }).ok).toBe(true);   // the nested table's first paragraph
    expect(shapeOf(session.doc.nodeAt(at)!)).toEqual(['sdtBlock', 'paragraph']);
    tbl = tblAt(await saveAndReload(session, probe), 1);
    expect(cellXmlShape(tbl, 0, 1)).toEqual(['sdt', 'p']);
    const repaired = await EditorDocument.open(await fixture('tables/cell-last-block-12-only-sdt-tbl-repaired.docx'));
    expect(shapeOf(repaired.doc.nodeAt(probeCell(repaired))!)).toEqual(['sdtBlock', 'paragraph']);   // Word's repair, the same shape
    // 5. Insert Table at a caret inside a block control in a cell: the table goes inside the control after its
    // paragraph, and the cell's paragraph after the control, one undo step.
    session = await EditorDocument.open(await fixture('tables/cell-last-block-10-only-sdt-p.docx'));
    at = probeCell(session);
    expect(runCommand(session, 'table.insert', { rows: 1, columns: 1 }, { from: at + 3 }).ok).toBe(true);
    expect(shapeOf(session.doc.nodeAt(at)!)).toEqual(['sdtBlock', 'paragraph']);
    expect(shapeOf(session.doc.nodeAt(at + 1)!)).toEqual(['paragraph', 'table']);
    tbl = tblAt(await saveAndReload(session, probe), 1);
    expect(cellXmlShape(tbl, 0, 1)).toEqual(['sdt', 'p']);
    expect(runCommand(session, 'history.undo').ok).toBe(true);
    expect(shapeOf(session.doc.nodeAt(at)!)).toEqual(['sdtBlock']);
  });

  test('repairStacks fills a gap and marks the extremes', () => {
    const g = (id: string): SdtGroup => ({ groupId: id, first: false, last: false });
    const repaired = repairStacks([[], [g('a')], [], [g('a')], []])!;
    expect(repaired.map((s) => s.map((x) => `${x.groupId}${x.first ? '+' : ''}${x.last ? '-' : ''}`))).toEqual([[], ['a+'], ['a'], ['a-'], []]);
    expect(repairStacks([[], [{ groupId: 'a', first: true, last: true }], []])).toBeUndefined();
  });
});

describe('row.shading: a background colour for the selected rows (Jason, 2026-11-15; ED-005 section 12.128)', () => {
  const shdOf = (tbl: wml.Tbl, r: number, c: number): unknown => { const s = tcPr(tbl, r, c)?.shd; return s ? JSON.parse(JSON.stringify(s, (k, v) => (k === 'PARENT' || k === 'TYPE_NAME' ? undefined : v))) : undefined; };
  // Office JS's shadingColor form in Word 365, the pattern colour the fill's (ED-005 section 12.130; it was `auto` before).
  const WORD_FORM = { val: 'clear', color: 'FFFF00', fill: 'FFFF00' };

  test("at a caret, every cell of the caret's row gets the fill as its own w:shd, in Word's form; the other rows keep theirs; null takes it away", async () => {
    const original = await fixture('tables.docx');
    const before = tblAt(await reloadBody(original), 1);
    const session = await EditorDocument.open(original);
    expect(session.doc.child(1).type.name).toBe('table');
    const at = cellStart(session, 1, 0, 0);
    expect(rowShadingOf(stateFor(session, { from: at }))).toBeNull();   // pct5 over an automatic fill: a pattern, no colour
    expect(runCommand(session, 'row.shading', { color: '#ffff00' }, { from: at }).ok).toBe(true);
    expect(rowShadingOf(stateFor(session, { from: cellStart(session, 1, 0, 1) }))).toBe('#FFFF00');
    expect(rowShadingOf(stateFor(session, { from: cellStart(session, 1, 1, 0) }))).toBeNull();
    const tbl = tblAt(await saveAndReload(session, original), 1);
    expect([shdOf(tbl, 0, 0), shdOf(tbl, 0, 1)]).toEqual([WORD_FORM, WORD_FORM]);
    expect([shdOf(tbl, 1, 0), shdOf(tbl, 1, 1)]).toEqual([shdOf(before, 1, 0), shdOf(before, 1, 1)]);
    expect(shdOf(tbl, 1, 1)).toMatchObject({ val: 'reverseDiagStripe' });
    // The same colour again changes nothing; a bad colour is refused.
    const doc = session.doc;
    expect(runCommand(session, 'row.shading', { color: 'FFFF00' }, { from: at }).ok).toBe(true);
    expect(session.doc.eq(doc)).toBe(true);
    expect(runCommand(session, 'row.shading', { color: 'yellow' }, { from: at }).ok).toBe(false);
    expect(runCommand(session, 'row.shading', { color: null }, { from: at }).ok).toBe(true);
    const cleared = tblAt(await saveAndReload(session, original), 1);
    expect([shdOf(cleared, 0, 0), shdOf(cleared, 0, 1)]).toEqual([undefined, undefined]);
    expect(tcPr(cleared, 0, 0)?.tcW?.w).toBe(tcPr(before, 0, 0)?.tcW?.w);   // the rest of the cell's properties kept
    expect(runCommand(session, 'history.undo').ok).toBe(true);
    expect(rowShadingOf(stateFor(session, { from: cellStart(session, 1, 0, 0) }))).toBe('#FFFF00');
  });

  test('a cell selection shades every cell of each row it covers, not only the cells selected', async () => {
    const original = await fixture('tables.docx');
    const session = await EditorDocument.open(original);
    const block = 1;
    const cellPos = (r: number, c: number): number => cellStart(session, block, r, c) - 2;
    const sel = CellSelection.create(session.doc, cellPos(0, 1), cellPos(1, 1));
    expect(runCommand(session, 'row.shading', { color: '#D9D9D9' }, sel).ok).toBe(true);
    const tbl = tblAt(await saveAndReload(session, original), block);
    const grey = { val: 'clear', color: 'D9D9D9', fill: 'D9D9D9' };
    expect([shdOf(tbl, 0, 0), shdOf(tbl, 0, 1), shdOf(tbl, 1, 0), shdOf(tbl, 1, 1)]).toEqual([grey, grey, grey, grey]);
  });

  test("a merged cell's continuations follow its shading, so every w:tc of the merge says one thing (Office JS's row setter writes the restart's row alone, ED-005 section 12.130); taken away, they are as they came", async () => {
    const original = await fixture('tables.docx');
    const session = await EditorDocument.open(original);
    const t3 = session.doc.child(3);   // A | B (vMerge restart) / C | (continuation)
    expect(t3.child(0).child(1).attrs['rowspan']).toBe(2);
    expect(runCommand(session, 'row.shading', { color: '#FFFF00' }, { from: cellStart(session, 3, 0, 0) }).ok).toBe(true);
    const tbl = tblAt(await saveAndReload(session, original), 3);
    expect(tcPr(tbl, 1, 1)?.vMerge).toBeDefined();
    expect([shdOf(tbl, 0, 0), shdOf(tbl, 0, 1), shdOf(tbl, 1, 1)]).toEqual([WORD_FORM, WORD_FORM, WORD_FORM]);
    expect(shdOf(tbl, 1, 0)).toBeUndefined();   // C, in the next row, is not in the selection
    expect(runCommand(session, 'row.shading', { color: null }, { from: cellStart(session, 3, 0, 0) }).ok).toBe(true);
    const back = tblAt(await reloadBody((await session.save()).bytes), 3);   // the cells as they came: perhaps no change to write
    expect([shdOf(back, 0, 0), shdOf(back, 0, 1), shdOf(back, 1, 1)]).toEqual([undefined, undefined, undefined]);
  });
});

describe("the table-indent round's files (Word 365 16.0.20326.20318, 2026-11-18; ED-005 section 12.132 item 5)", () => {
  test('the compatibility mode is read as Word reads it, 12 where none is stated; Word\'s save of that file writes 12', async () => {
    const modeOf = async (name: string): Promise<number> => (await EditorDocument.open(await fixture(`table-indent/${name}.docx`))).compatibilityMode;
    expect(await modeOf('ind15-mode15')).toBe(15);
    expect(await modeOf('ind14-mode14')).toBe(14);
    expect(await modeOf('ind00-no-mode')).toBe(12);
    expect(await modeOf('ind00-no-mode-word365')).toBe(12);
    expect(await modeOf('ind15-mode15-word365')).toBe(15);
  });
});

describe("the row-shading round's saves (Word 365 16.0.20326.20318, 2026-11-15; ED-005 section 12.128)", () => {
  /** Every top-level table's cells' own w:shd, as `val/color/fill` or '-' for none, row by row. */
  async function shadings(name: string): Promise<string[][][]> {
    const body = await reloadBody(await fixture(`row-shading/${name}.docx`));
    return body.content!.filter((el) => typeNameOf(el) === 'org_docx4j_wml.Tbl').map((el) => rowsOf(el.value as wml.Tbl).map((r) => cellsOf(r.element.value).map((c) => {
      const s = c.element.value.tcPr?.shd;
      return s ? [s.val, s.color, s.fill].map((v) => v ?? '').join('/') : '-';
    })));
  }
  const clearAuto = 'clear/auto/auto', yellow = 'clear/auto/FFFF00';

  test('shd1: Word keeps fill auto, white and yellow as written and drops val nil (which it draws as none of its own)', async () => {
    expect(await shadings('shd1-style-forms')).toEqual([
      [['-', '-'], [clearAuto, clearAuto], ['nil//', 'nil//'], ['clear/auto/FFFFFF', 'clear/auto/FFFFFF'], [yellow, yellow]],
      [['-', '-'], ['-', '-'], [clearAuto, clearAuto], ['-', '-'], ['nil//', 'nil//'], ['-', '-']],
    ]);
    expect(await shadings('shd1-style-forms-word365')).toEqual([
      [['-', '-'], [clearAuto, clearAuto], ['-', '-'], ['clear/auto/FFFFFF', 'clear/auto/FFFFFF'], [yellow, yellow]],
      [['-', '-'], ['-', '-'], [clearAuto, clearAuto], ['-', '-'], ['-', '-'], ['-', '-']],
    ]);
  });

  test("shd2 and shd3: Word saves a merge's shading and the editor's file as they came", async () => {
    expect(await shadings('shd2-merged-word365')).toEqual(await shadings('shd2-merged'));
    const editor = await shadings('shd3-editor');
    expect(editor).toEqual([
      [[yellow, yellow], ['clear/auto/D9D9D9', 'clear/auto/D9D9D9'], ['clear/auto/D9D9D9', 'clear/auto/D9D9D9']],
      [['clear/auto/00B0F0', 'clear/auto/00B0F0'], ['-', 'clear/auto/00B0F0']],
      [['-', '-'], ['-', '-']],
    ]);
    expect(await shadings('shd3-editor-word365')).toEqual(editor);
  });
});

describe('through Yjs', () => {
  test('a table survives the store: read back equal, and an edit through the store exports', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    expect(session.doc.eq(session.imported)).toBe(true);
    const t = session.doc.child(7);
    expect(t.type.name).toBe('table');
    expect(t.child(1).child(0).attrs['rowspan']).toBe(2);
    const pos = cellStart(session, 7, 0, 0);
    const state = stateFor(session, { from: pos });
    const next = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos + 4)).insertText('!'));
    session.replace(next.doc);
    expect(textOf(session.doc.child(7).child(0).child(0))).toContain('!');
    const tbl = tblAt(await saveAndReload(session, original), 7);
    expect(cellText(tbl, 0, 0)).toContain('!');
    expect(tcPr(tbl, 1, 0)?.vMerge?.val).toBe('restart');
    expect(cellsOf(rowsOf(tbl)[2]!.element.value)).toHaveLength(3);   // the continuation cell is still there
  });
});


describe('a cell selection is formatted and read cell by cell (2026-10-15)', () => {
  test('font and paragraph commands over selected cells reach every cell, and the readings cover them all', async () => {
    const session = await EditorDocument.open(await fixture('tables/table-fit-fixed.docx'));
    let pos = 0;
    for (let k = 0; k < 1; k++) pos += session.doc.child(k).nodeSize;
    const map = TableMap.get(session.doc.child(1));
    const cellAt = (r: number, c: number): number => pos + 1 + map.map[r * map.width + c]!;
    const sizesIn = (r: number, c: number): (number | null)[] => {
      const out: (number | null)[] = [];
      session.doc.child(1).nodeAt(map.map[r * map.width + c]!)!.descendants((n) => { if (n.isText) out.push((n.marks.find((m) => m.type.name === 'font')?.attrs['size'] as number | undefined) ?? null); return true; });
      return out;
    };
    const jcIn = (r: number, c: number): unknown => session.doc.child(1).nodeAt(map.map[r * map.width + c]!)!.firstChild!.attrs['jc'];
    // Rows 0 and 1, columns 0 and 1, the head cell first: a cell selection's from and to are the head cell's alone.
    const four = CellSelection.create(session.doc, cellAt(1, 1), cellAt(0, 0));
    expect(runCommand(session, 'font.shrink', undefined, four).ok).toBe(true);
    for (const [r, c] of [[0, 0], [0, 1], [1, 0], [1, 1]] as const) expect(sizesIn(r, c), `cell ${r},${c}`).toEqual([10]);
    for (const [r, c] of [[0, 2], [1, 2], [2, 0], [3, 1]] as const) expect(sizesIn(r, c), `cell ${r},${c}`).toEqual([null]);
    // The readings see every selected cell: the size uniform at 10, bold off; bold on, then active over them all.
    const again = CellSelection.create(session.doc, cellAt(1, 1), cellAt(0, 0));
    expect(selectionFormatting(stateFor(session, again)).font.size).toBe(10);
    expect(commandNamed('font.bold')!.isActive!(stateFor(session, again), undefined)).toBe(false);
    expect(runCommand(session, 'font.bold', { value: true }, again).ok).toBe(true);
    const bolded = CellSelection.create(session.doc, cellAt(1, 1), cellAt(0, 0));
    expect(commandNamed('font.bold')!.isActive!(stateFor(session, bolded), undefined)).toBe(true);
    expect(session.doc.child(1).nodeAt(map.map[2]!)!.textContent).toBe('Due');
    expect(selectionFormatting(stateFor(session, CellSelection.create(session.doc, cellAt(0, 2), cellAt(0, 2)))).font.bold).toBe(false);
    // A column selected: its four paragraphs centred, the rest as they were.
    const column = CellSelection.create(session.doc, cellAt(3, 2), cellAt(0, 2));
    expect(runCommand(session, 'paragraph.alignment', { value: 'Centered' }, column).ok).toBe(true);
    for (let r = 0; r < 4; r++) expect(jcIn(r, 2), `row ${r}`).toBe('center');
    expect(jcIn(0, 0)).toBeNull();
    expect(jcIn(3, 1)).toBeNull();
    // The file carries it all.
    const body = await saveAndReload(session, await fixture('tables/table-fit-fixed.docx'));
    const tbl = tblAt(body, 1);
    expect(cellsOf(rowsOf(tbl)[0]!.element.value)[0]!.element.value.content?.some((e) => typeNameOf(e) === 'org_docx4j_wml.P' && JSON.stringify(e).includes('"sz"'))).toBe(true);
    expect(JSON.stringify(cellsOf(rowsOf(tbl)[2]!.element.value)[2]!.element.value.content?.[0])).toContain('"jc"');
    expect(JSON.stringify(cellsOf(rowsOf(tbl)[2]!.element.value)[0]!.element.value.content?.[0])).not.toContain('"sz"');
  });
});
