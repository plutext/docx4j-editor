// SPDX-License-Identifier: GPL-3.0-or-later
//
// Tables (ED-003 section 3.7): commands in Office JS shapes over `prosemirror-tables`' commands,
// and the invariant of ED-001 section 4.1 that a row-level control's rows stay contiguous: a row
// inserted inside a group inherits the group, and a group's `first` and `last` flags follow its
// extremes. Every command is one transaction under the local origin.
import { TextSelection, Selection, type Command, type EditorState, type Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import {
  addRowAfter, addRowBefore, deleteRow, deleteColumn, mergeCells, splitCell, deleteTable,
  isInTable, selectedRect, goToNextCell, CellSelection, TableMap,
} from 'prosemirror-tables';
import type { SdtGroup } from '../import.mjs';
import { register, sessionOf, storyOf } from './spec.mjs';
import { sectionAt } from './sections.mjs';
import { TWIPS_PER_PX } from '../import.mjs';
import { markDeletedRange, UNTRACKED_META, type Tracking } from '../track.mjs';
import { plainEqual } from '../json.mjs';

/** A table node's position: the innermost table around the selection. */
export function tableAround(state: EditorState): { node: PMNode; pos: number } | undefined {
  const $from = state.selection.$from;
  for (let d = $from.depth; d >= 0; d--) {
    const n = $from.node(d);
    if (n.type.name === 'table') return { node: n, pos: d === 0 ? 0 : $from.before(d) };
  }
  return undefined;
}

/**
 * Keep every row-level and cell-level control group contiguous and its flags right (ED-003
 * section 3.6): rows between a group's first and last members join it at the depth their
 * neighbours hold it; `first` and `last` mark the extremes. Returns the transaction with the
 * repairs, or null when nothing needed one.
 */
/**
 * Word's rule for a cell (ED-005 section 12.109, measured in Word 365: a file whose `w:tc` ends with a
 * `w:tbl` is reported as unreadable and repaired by an empty paragraph after the table): a cell's
 * last block is a paragraph, the last block as Word reads it: through block-level controls (a `w:sdt`
 * ending in a `w:tbl` is repaired too, the paragraph after the control, never inside it; Jason's probes of
 * 2026-10-22). Appends one to every cell whose last child, through controls, is a table; `null` when none
 * needs it. Over the whole document when `whole` (less the top-level blocks `skipTopLevel` names by
 * index: a save leaves an imported block as it is), else over the cells the transaction's steps
 * changed or lie in (a table inserted at a cell's end is inside that cell). The view's plugin, the headless
 * runner and `repairedForSave` call it; a new dispatch path must too.
 */
export function endCellsWithParagraph(tr: Transaction, whole = false, skipTopLevel?: ReadonlySet<number>): Transaction | null {
  const doc = tr.doc;
  const ends = new Map<number, PMNode | null>();
  const consider = (node: PMNode, pos: number): void => {
    if (node.type.name !== 'table_cell') return;
    // Word checks the last block as it reads it, through block-level controls (Jason's probes, 2026-10-22):
    // a control ending in a table at a cell's end is repaired too, the paragraph after the control.
    let last = node.lastChild;
    while (last?.type.name === 'sdtBlock') last = last.lastChild;
    if (last?.type.name === 'table') ends.set(pos + node.nodeSize - 1, lastParagraphBefore(node));
  };
  if (whole) doc.forEach((block, pos, i) => { if (!skipTopLevel?.has(i)) block.descendants((node, p) => { consider(node, pos + 1 + p); return true; }); });
  else {
    for (const map of tr.mapping.maps) {
      map.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
        const from = Math.max(0, Math.min(newStart, doc.content.size));
        const to = Math.max(from, Math.min(newEnd, doc.content.size));
        doc.nodesBetween(from, to, (node, pos) => { consider(node, pos); return true; });
        const $from = doc.resolve(from);
        for (let d = $from.depth; d > 0; d--) consider($from.node(d), $from.before(d));
      });
    }
  }
  if (ends.size === 0) return null;
  const paragraph = doc.type.schema.nodes['paragraph']!;
  for (const at of [...ends.keys()].sort((a, b) => b - a)) {
    // Word's insertTable at a cell's end leaves an empty paragraph in the cell's last paragraph's properties
    // (core-ts check 52, cases 86 to 95): a copy of that paragraph's attributes, `copiedFrom` naming its element so
    // the export takes its whole `w:pPr` (as Enter's first half does), its own ids, section and revisions left out.
    // Where no paragraph precedes the table (a cell holding only a table, or only a control around one), Word's
    // open-time repair shape: a plain paragraph.
    const source = ends.get(at);
    const attrs = source ? { ...source.attrs, sourceId: null, sectPr: null, markRevision: null, pPrChange: null, copiedFrom: source.attrs['sourceId'] ?? source.attrs['copiedFrom'] } : null;
    tr.insert(at, paragraph.createAndFill(attrs)!);
  }
  return tr;
}

/** The cell's last paragraph before its trailing table, at the cell's level or inside a control there; `null` when none. */
function lastParagraphBefore(cell: PMNode): PMNode | null {
  let found: PMNode | null = null;
  const visit = (node: PMNode): void => {
    for (let i = node.childCount - 1; i >= 0 && !found; i--) {
      const child = node.child(i);
      if (child.type.name === 'paragraph') found = child;
      else if (child.type.name === 'sdtBlock') visit(child);
    }
  };
  visit(cell);
  return found;
}

export function normalizeRowGroups(tr: Transaction): Transaction | null {
  let touched = false;
  const tables: { node: PMNode; pos: number }[] = [];
  tr.doc.descendants((node, pos) => {
    if (node.type.name === 'table') tables.push({ node, pos });
    return true;
  });
  // A split copies the cell's attributes, its source id included: only one node may stand for
  // a source element, so the copies lose the id (export then builds them from their attributes).
  const seen = new Map<string, number>();
  const duplicates: number[] = [];
  tr.doc.descendants((node, pos) => {
    if (node.type.name !== 'table_cell' && node.type.name !== 'table_row') return true;
    const id = node.attrs['sourceId'] as string | null;
    if (!id) return true;
    if (seen.has(id)) duplicates.push(pos); else seen.set(id, pos);
    return true;
  });
  for (const pos of duplicates.reverse()) {
    const node = tr.doc.nodeAt(pos)!;
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, sourceId: null, merged: null });
    touched = true;
  }
  for (const { node: table, pos: tablePos } of tables) {
    const stacks: SdtGroup[][] = [];
    table.forEach((row) => stacks.push(((row.attrs['sdt'] as SdtGroup[] | null) ?? []).map((g) => ({ ...g }))));
    const repaired = repairStacks(stacks);
    if (!repaired) continue;
    touched = true;
    let rowPos = tablePos + 1;
    table.forEach((row, _offset, i) => {
      const next = repaired[i]!;
      const current = (row.attrs['sdt'] as SdtGroup[] | null) ?? [];
      if (JSON.stringify(current) !== JSON.stringify(next)) tr.setNodeMarkup(rowPos, undefined, { ...row.attrs, sdt: next.length > 0 ? next : null });
      rowPos += row.nodeSize;
    });
  }
  return touched ? tr : null;
}

/** The repaired stacks, or undefined when they were right already. */
export function repairStacks(stacks: SdtGroup[][]): SdtGroup[][] | undefined {
  const ids = new Map<string, { first: number; last: number; depth: number }>();
  stacks.forEach((stack, r) => stack.forEach((g, d) => {
    const e = ids.get(g.groupId);
    if (!e) ids.set(g.groupId, { first: r, last: r, depth: d });
    else e.last = r;
  }));
  const next = stacks.map((s) => s.map((g) => ({ groupId: g.groupId, first: false, last: false })));
  // Fill the gaps: a row between a group's members joins it at the group's depth.
  for (const [id, e] of ids) {
    for (let r = e.first; r <= e.last; r++) {
      const stack = next[r]!;
      if (stack.some((g) => g.groupId === id)) continue;
      stack.splice(Math.min(e.depth, stack.length), 0, { groupId: id, first: false, last: false });
    }
  }
  for (const [id, e] of ids) {
    for (const g of next[e.first]!) if (g.groupId === id) g.first = true;
    for (const g of next[e.last]!) if (g.groupId === id) g.last = true;
  }
  const same = stacks.every((s, r) => JSON.stringify(s.map((g) => ({ groupId: g.groupId, first: g.first, last: g.last }))) === JSON.stringify(next[r]));
  return same ? undefined : next;
}

/** A selection over one document, re-created over another of the same shape. */
function remap(sel: Selection, doc: PMNode): Selection {
  if (sel instanceof CellSelection) return CellSelection.create(doc, sel.$anchorCell.pos, sel.$headCell.pos);
  return TextSelection.create(doc, sel.anchor, sel.head);
}

/**
 * Run a prosemirror-tables command `count` times as one transaction, then repair the row groups
 * (ED-001 section 4.1's invariant), so a repeated command is one undo step. Stops early when the
 * selection leaves the table (a row deletion that emptied it).
 */
/**
 * Rows deleted with tracking on (ED-005 section 5.4): kept, each marked deleted with its cells'
 * paragraph marks and text, as Word writes them, since the table's structure changes only when
 * the deletion is accepted and a vertical merge through the rows stays whole; a row this author
 * inserted goes for real. A cell selection's rows, or `count` rows from the caret's down.
 */
function markRowsDeleted(state: EditorState, dispatch: ((tr: Transaction) => void) | undefined, count: number, tracking: Tracking): boolean {
  const rows = rowsToDelete(state, count);
  if (rows.length === 0) return false;
  if (!dispatch) return true;
  const tr = state.tr;
  for (const r of [...rows].reverse()) markDeletedRange(tr, r.from, r.to, tracking);
  dispatch(tr.setMeta(UNTRACKED_META, true).scrollIntoView());
  return true;
}

/** The rows a deletion takes, top to bottom, as document ranges. */
function rowsToDelete(state: EditorState, count: number): { from: number; to: number }[] {
  let table: PMNode;
  let tableStart: number;
  let top: number;
  let bottom: number;
  if (state.selection instanceof CellSelection) {
    const rect = selectedRect(state);
    [table, tableStart, top, bottom] = [rect.table, rect.tableStart, rect.top, rect.bottom - 1];
  } else {
    const $from = state.selection.$from;
    let depth = $from.depth;
    while (depth > 0 && $from.node(depth).type.name !== 'table_row') depth--;
    if (depth === 0) return [];
    table = $from.node(depth - 1);
    tableStart = $from.start(depth - 1);
    top = $from.index(depth - 1);
    bottom = Math.min(top + count - 1, table.childCount - 1);
  }
  const out: { from: number; to: number }[] = [];
  let pos = tableStart;
  table.forEach((row, _offset, i) => {
    if (i >= top && i <= bottom) out.push({ from: pos, to: pos + row.nodeSize });
    pos += row.nodeSize;
  });
  return out;
}

function repeated(command: Command, count = 1, stopOutsideTable = false): Command {
  return (state, dispatch, view) => {
    if (!dispatch) return command(state, undefined, view);
    let current = state;
    const acc = state.tr;
    for (let i = 0; i < count; i++) {
      let out: Transaction | undefined;
      const ok = command(current, (tr) => { out = tr; }, view);
      if (!ok || !out) return i > 0;
      for (const step of out.steps) acc.step(step);
      current = current.apply(out);
      if (stopOutsideTable && !isInTable(current)) break;
    }
    acc.setSelection(remap(current.selection, acc.doc));
    normalizeRowGroups(acc);
    dispatch(acc.scrollIntoView());
    return true;
  };
}

register<{ rows?: number; columns?: number; style?: string | null; autoFit?: TableAutoFit; columnWidth?: number }>({
  name: 'table.insert',
  label: 'Insert table',
  args: {
    rows: 'rows, default 2', columns: 'columns, default 2', style: "the table style id; the document's TableGrid when it has one",
    autoFit: "Word's Insert Table choices (ED-005 section 12.83): 'fixed' (fixed column width, the default), 'contents' (AutoFit to contents) or 'window' (AutoFit to window)",
    columnWidth: "with 'fixed', each column's width in points; Auto (the text width shared equally) when absent",
  },
  choices: { autoFit: ['fixed', 'contents', 'window'] },
  run: ({ rows = 2, columns = 2, style, autoFit = 'fixed', columnWidth } = {}) => (state, dispatch) => {
    const { $from } = state.selection;
    // In a paragraph anywhere: the body, a cell (a nested table, as Word allows; 2026-09-25), a block control.
    if ($from.parent.type.name !== 'paragraph') return false;
    if (!(rows >= 1 && columns >= 1)) return false;
    const schema = state.schema;
    const session = sessionOf(state);
    const resolver = session?.pkg.propertyResolverOrUndefined;
    const styleId = style === undefined ? (resolver?.getStyle('TableGrid') ? 'TableGrid' : null) : style;
    // The section's text width, shared equally (A4 with 2.54 cm margins when unknown), or the width given per column.
    const width = textWidthTwips(state);
    const colTwips = autoFit === 'fixed' && columnWidth !== undefined && columnWidth > 0 ? Math.round(columnWidth * 20) : Math.floor(width / columns);
    const cellWidth = colTwips / 15;
    // What Word's dialog writes (check 42, ED-005 section 12.83): fixed column width, Auto or a value, is an automatic
    // table width with the cells' widths in twips, no layout element (Word's default table); AutoFit to contents an
    // automatic width with automatic cells, which the content sizes; AutoFit to window a 100% width with the cells as
    // shares.
    const cells = (): PMNode[] => Array.from({ length: columns }, () => schema.nodes['table_cell']!.create({ colwidth: autoFit === 'contents' ? null : [cellWidth] }, schema.nodes['paragraph']!.create()));
    const rowNodes = Array.from({ length: rows }, () => schema.nodes['table_row']!.create(null, cells()));
    const tblPr: Record<string, unknown> = { TYPE_NAME: 'org_docx4j_wml.TblPr' };
    tblPr['tblW'] = autoFit === 'window' ? { w: 5000, type: 'pct' } : { w: 0, type: 'auto' };
    if (styleId) tblPr['tblStyle'] = { val: styleId };
    const table = schema.nodes['table']!.create({ tblPr, tblGrid: Array.from({ length: columns }, () => colTwips) }, rowNodes);
    // After the caret's paragraph, at whatever depth it sits (a cell's paragraph for a nested
    // table); a caret in an empty paragraph replaces it, unless it is a cell's only block, since a
    // cell must keep a paragraph (the schema's `block+` would take the table alone; Word would not).
    const depth = $from.depth;
    const blockPos = $from.before(depth);
    const block = $from.node(depth);
    const parent = $from.node(depth - 1);
    const tr = state.tr;
    let at: number;
    if (block.content.size === 0 && !(parent.type.name === 'table_cell' && parent.childCount === 1)) { tr.replaceWith(blockPos, blockPos + block.nodeSize, table); at = blockPos; }
    else { at = blockPos + block.nodeSize; tr.insert(at, table); }
    // The caret in the first cell.
    tr.setSelection(TextSelection.create(tr.doc, at + 4));
    dispatch?.(tr.scrollIntoView());
    return true;
  },
  isEnabled: (state) => state.selection.$from.parent.type.name === 'paragraph',
});

register<{ location?: 'Before' | 'After'; count?: number }>({
  name: 'table.addRows',
  label: 'Add rows',
  args: { location: "'Before' or 'After' the row at the selection; 'After' by default", count: 'rows to add, default 1' },
  choices: { location: ['After', 'Before'] },
  run: ({ location = 'After', count = 1 } = {}) => repeated(location === 'Before' ? addRowBefore : addRowAfter, count),
  isEnabled: isInTable,
});

register<{ count?: number }>({
  name: 'table.deleteRows',
  label: 'Delete rows',
  args: { count: 'rows to delete from the selection down, default 1 (a cell selection deletes its rows)' },
  run: ({ count = 1 } = {}) => (state, dispatch, view) => {
    const tracking = storyOf(state) ? undefined : sessionOf(state)?.tracking();
    return tracking ? markRowsDeleted(state, dispatch, count, tracking) : repeated(deleteRow, count, true)(state, dispatch, view);
  },
  isEnabled: isInTable,
});

// ---------------------------------------------------------------- column widths and the layout mode (ED-005 section 12.78)

/** Word's three choices in Table Properties: fixed column widths, AutoFit to window, AutoFit to contents. */
export type TableAutoFit = 'fixed' | 'window' | 'contents';

/** The section's text width in twips: the page less its margins and gutter (A4 with 2.54 cm margins when unknown). */
export function textWidthTwips(state: EditorState, pos: number = state.selection.from): number {
  const s = (sectionAt(state, pos).sectPr ?? {}) as { pgSz?: { w?: number }; pgMar?: { left?: number; right?: number; gutter?: number } };
  return Math.max(720, (s.pgSz?.w ?? 11906) - (s.pgMar?.left ?? 1440) - (s.pgMar?.right ?? 1440) - (s.pgMar?.gutter ?? 0));
}

/** A table's AutoFit choice from its properties: `w:tblLayout fixed` is fixed; else a `w:tblW` in percent is window, anything else contents (Word's own reading of the three). */
export function tableAutoFitOf(tblPr: unknown): TableAutoFit {
  const pr = (tblPr ?? {}) as { tblLayout?: { type?: string }; tblW?: { type?: string; w?: number } };
  if (pr.tblLayout?.type === 'fixed') return 'fixed';
  return pr.tblW?.type === 'pct' ? 'window' : 'contents';
}

/** Each column's width in pixels: a cell's `colwidth` where a cell states one, else the grid's, else an equal share of the text width. */
export function columnWidthsPx(table: PMNode, map: TableMap, textTwips: number): number[] {
  const grid = table.attrs['tblGrid'] as number[] | null;
  const widths: number[] = new Array<number>(map.width).fill(0);
  const seen = new Set<number>();
  for (let r = 0; r < map.height; r++) {
    for (let c = 0; c < map.width; c++) {
      const pos = map.map[r * map.width + c]!;
      if (seen.has(pos)) continue;
      seen.add(pos);
      const cell = table.nodeAt(pos)!;
      const colwidth = cell.attrs['colwidth'] as number[] | null;
      for (let k = 0; k < (cell.attrs['colspan'] as number); k++) {
        const w = colwidth?.[k];
        if (typeof w === 'number' && w > 0 && widths[c + k] === 0) widths[c + k] = w;
      }
    }
  }
  const fallback = textTwips / Math.max(1, map.width) / TWIPS_PER_PX;
  return widths.map((w, i) => (w > 0 ? w : grid && grid.length === map.width && grid[i] ? grid[i]! / TWIPS_PER_PX : fallback));
}

/** Every cell's `colwidth` and the table's grid set from one list of column widths in pixels, by the map. */
export function setColumnWidths(tr: Transaction, tablePos: number, widths: readonly number[]): void {
  const table = tr.doc.nodeAt(tablePos)!;
  const map = TableMap.get(table);
  const seen = new Set<number>();
  for (let r = 0; r < map.height; r++) {
    for (let c = 0; c < map.width; c++) {
      const pos = map.map[r * map.width + c]!;
      if (seen.has(pos)) continue;
      seen.add(pos);
      const cell = table.nodeAt(pos)!;
      const span = cell.attrs['colspan'] as number;
      tr.setNodeMarkup(tablePos + 1 + pos, undefined, { ...cell.attrs, colwidth: Array.from({ length: span }, (_, k) => widths[c + k] ?? widths[widths.length - 1]!) });
    }
  }
  tr.setNodeMarkup(tablePos, undefined, { ...table.attrs, tblGrid: widths.map((w) => Math.round(w * TWIPS_PER_PX)) });
}

/**
 * Columns added keep the table's layout mode (Word, to be held by check 39): with fixed widths the
 * new column takes the width of the column it is added beside and the table grows, past the margin
 * if it must; with AutoFit the table keeps its width and every column is scaled down to make room.
 */
/** Whether a table's cells are automatic (AutoFit to contents, check 42): its first row states no widths. */
export function columnWidthsAside(table: PMNode): boolean {
  const first = table.firstChild;
  if (!first || first.childCount === 0) return false;
  let aside = true;
  first.forEach((cell) => { if (cell.attrs['colwidth'] !== null) aside = false; });
  return aside;
}

/** Makes every cell of the table at `tablePos` automatic (its width null); the grid stays as Word's cache. */
export function clearColumnWidths(tr: Transaction, tablePos: number): void {
  const table = tr.doc.nodeAt(tablePos)!;
  table.descendants((n, pos) => {
    if (n.type.name === 'table_cell') { if (n.attrs['colwidth'] !== null) tr.setNodeMarkup(tablePos + 1 + pos, undefined, { ...n.attrs, colwidth: null }); return false; }
    return true;
  });
}

/**
 * Inserts `count` grid columns at the boundary before column `col` of the table at `tablePos`, as Word's Insert
 * Left and Insert Right do (checks 39 and 40, ED-005 sections 12.79 and 12.82): in each row a new cell goes before
 * the cell that starts at the boundary; a cell spanning the boundary keeps its span and the new cell goes before it;
 * a row covered from above at the boundary gets its cells before the first of its own at or past the boundary (a
 * vertical merge stays one column wide, as the engine's insertion has it); a row that ends before the boundary gets
 * the cells after its last. A new cell takes, from the cell to the right of the boundary in its row when there is one, else the left
 * one, what Word copies (check 41): its shading and its first paragraph's properties (the paragraph's mark included),
 * not its borders or vertical alignment; one empty paragraph. (prosemirror-tables' `addColumn` widens a spanning
 * cell instead, which Word does not: 40b1.)
 */
export function insertGridColumns(tr: Transaction, tablePos: number, col: number, count: number): void {
  const table = tr.doc.nodeAt(tablePos)!;
  const map = TableMap.get(table);
  const cellType = table.type.schema.nodes['table_cell']!;
  // Collected first, then applied from the end so earlier positions stay valid.
  const inserts: { at: number; attrs: Record<string, unknown>; paragraph: Record<string, unknown> }[] = [];
  let rowStart = tablePos + 1;
  for (let r = 0; r < map.height; r++) {
    const row = table.child(r);
    const rowEnd = rowStart + row.nodeSize - 1;
    // The cell at a column, its own or the one covering it from above (the template for a new cell beside it).
    const template = (c: number): PMNode | undefined => (c < 0 || c >= map.width ? undefined : table.nodeAt(map.map[r * map.width + c]!) ?? undefined);
    const source = template(col) ?? template(col - 1);
    let at: number | undefined;
    if (col < map.width) {
      const pos = map.map[r * map.width + col]!;
      if (map.findCell(pos).top === r) at = tablePos + 1 + pos;   // before the cell at or spanning the boundary
      else {
        // Covered from above: before the row's first own cell at or past the boundary, else after its last.
        let offset = 0;
        row.forEach((cell, o) => { if (at === undefined && map.findCell(rowStart - tablePos - 1 + 1 + o).left >= col) at = rowStart + 1 + o; offset = o + cell.nodeSize; });
        if (at === undefined) at = rowStart + 1 + offset;
      }
    } else {
      at = rowEnd;   // past the last cell
    }
    inserts.push({ at, attrs: newCellAttrs(source), paragraph: newCellParagraphAttrs(source) });
    rowStart += row.nodeSize;
  }
  const paragraphType = table.type.schema.nodes['paragraph']!;
  for (const { at, attrs, paragraph } of inserts.sort((a, b) => b.at - a.at)) {
    for (let k = 0; k < count; k++) tr.insert(at, cellType.create(attrs, paragraphType.create(paragraph)));
  }
}

/** A new cell's attributes beside `source`: its shading alone of the cell properties (check 41: not its borders or vertical alignment). */
function newCellAttrs(source: PMNode | undefined): Record<string, unknown> {
  const tcPr = (source?.attrs['tcPr'] ?? null) as Record<string, unknown> | null;
  const shd = tcPr?.['shd'];
  return {
    sourceId: null, colspan: 1, rowspan: 1, colwidth: null, merged: null, sdt: null, revision: null,
    tcPr: shd !== undefined ? { TYPE_NAME: 'org_docx4j_wml.TcPr', shd } : null,
  };
}

/** The new cell's one paragraph: the properties of `source`'s first paragraph, as Word gives it (check 41). */
function newCellParagraphAttrs(source: PMNode | undefined): Record<string, unknown> {
  const first = source?.firstChild;
  if (!first || first.type.name !== 'paragraph') return {};
  const a = first.attrs;
  return { pStyle: a['pStyle'], jc: a['jc'], ind: a['ind'], spacing: a['spacing'], bidi: a['bidi'], numPr: a['numPr'], tabs: a['tabs'], markRPr: a['markRPr'] };
}

function addColumnsCommand(location: 'Before' | 'After', count: number): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false;
    const rect = selectedRect(state);
    const tablePos = rect.tableStart - 1;
    const textTwips = textWidthTwips(state, rect.tableStart);
    const before = columnWidthsPx(rect.table, rect.map, textTwips);
    const at = location === 'Before' ? rect.left : rect.right;
    // The width copied is the column to the right of the boundary when there is one, else the left (check 40).
    const copied = before[at] ?? before[at - 1] ?? before[0]!;
    if (dispatch) {
      const tr = state.tr;
      insertGridColumns(tr, tablePos, at, count);
      const widths = [...before.slice(0, at), ...Array.from({ length: count }, () => copied), ...before.slice(at)];
      const mode = tableAutoFitOf(rect.table.attrs['tblPr']);
      if (mode === 'window') {
        // AutoFit to window: the table keeps its width, every column scaled to make room (check 39). An automatic table
        // (Word's default, or automatic cells) grows by the copied width as a fixed one does, as the engine's insertion
        // has it (core-ts 0.4.0; the grid of an automatic-cell table is Word's cache, measured afresh at open).
        const was = before.reduce((a, b) => a + b, 0);
        const now = widths.reduce((a, b) => a + b, 0);
        const factor = now > 0 ? was / now : 1;
        for (let i = 0; i < widths.length; i++) widths[i] = widths[i]! * factor;
      }
      if (mode === 'contents' && columnWidthsAside(rect.table)) {
        // Automatic cells: the content sizes them, and the grid alone takes the new column (Word measures it afresh).
        const table = tr.doc.nodeAt(tablePos)!;
        tr.setNodeMarkup(tablePos, undefined, { ...table.attrs, tblGrid: widths.map((w) => Math.round(w * TWIPS_PER_PX)) });
      } else setColumnWidths(tr, tablePos, widths);
      if (mode === 'fixed') {
        // Word 365 (check 39, ED-005 section 12.79): a fixed table's stated width grows with its grid.
        const table = tr.doc.nodeAt(tablePos)!;
        const tblPr = (table.attrs['tblPr'] ?? {}) as { tblW?: { w?: number; type?: string } };
        if (tblPr.tblW?.type === 'dxa') {
          const sum = widths.reduce((a, b) => a + b, 0);
          tr.setNodeMarkup(tablePos, undefined, { ...table.attrs, tblPr: { ...tblPr, tblW: { ...tblPr.tblW, w: Math.round(sum * 15) } } });
        }
      }
      // The caret stays in its cell: the new cells shift the positions.
      tr.setSelection(Selection.near(tr.doc.resolve(tr.mapping.map(state.selection.from))));
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

register<{ location?: 'Before' | 'After'; count?: number }>({
  name: 'table.addColumns',
  tracking: 'untracked',
  label: 'Add columns',
  args: { location: "'Before' or 'After' the column at the selection; 'After' by default", count: 'columns to add, default 1' },
  choices: { location: ['After', 'Before'] },
  run: ({ location = 'After', count = 1 } = {}) => addColumnsCommand(location === 'Before' ? 'Before' : 'After', count),
  // Office JS: TableCell.insertColumns at the cell the selection starts in (core-ts 0.4.0, CR-002 section 42).
  api: ({ location = 'After', count = 1 } = {}) => `selection.paragraphs[0].parentTableCell.insertColumns(${JSON.stringify(location === 'Before' ? 'Before' : 'After')}, ${Math.max(1, Math.round(count))});`,
  isEnabled: isInTable,
});

register<{ autoFit: TableAutoFit }>({
  name: 'table.properties',
  tracking: 'untracked',
  label: 'Table properties',
  args: { autoFit: "'fixed' (fixed column widths), 'window' (AutoFit to window: the table the text width, w:tblW 100%) or 'contents' (AutoFit to contents)" },
  choices: { autoFit: ['fixed', 'window', 'contents'] },
  // Word's Table Properties, the AutoFit behaviour (ED-005 section 12.78). Fixed keeps the widths as they are and
  // marks the layout fixed; window scales the columns to the text width and marks the table's width 100%; contents
  // marks the layout AutoFit with the width automatic, the columns left as they are until Word lays them out.
  run: ({ autoFit }) => (state, dispatch) => {
    if (!isInTable(state)) return false;
    const rect = selectedRect(state);
    const tablePos = rect.tableStart - 1;
    const tblPr = { ...((rect.table.attrs['tblPr'] as Record<string, unknown> | null) ?? { TYPE_NAME: 'org_docx4j_wml.TblPr' }) };
    const tr = state.tr;
    if (autoFit === 'fixed') {
      tblPr['tblLayout'] = { type: 'fixed' };
    } else {
      delete tblPr['tblLayout'];
      tblPr['tblW'] = autoFit === 'window' ? { w: 5000, type: 'pct' } : { w: 0, type: 'auto' };
    }
    tr.setNodeMarkup(tablePos, undefined, { ...rect.table.attrs, tblPr });
    if (autoFit === 'window') {
      const textTwips = textWidthTwips(state, rect.tableStart);
      const widths = columnWidthsPx(rect.table, rect.map, textTwips);
      const total = widths.reduce((a, b) => a + b, 0);
      const target = textTwips / TWIPS_PER_PX;
      if (total > 0) setColumnWidths(tr, tablePos, widths.map((w) => (w * target) / total));
    } else if (autoFit === 'contents') {
      // Word's AutoFit Contents makes the cells automatic (check 42): the content sizes them from here on.
      clearColumnWidths(tr, tablePos);
    } else if (columnWidthsAside(rect.table)) {
      // Fixed over automatic cells: the grid's widths become the cells' own.
      const grid = (rect.table.attrs['tblGrid'] as number[] | null) ?? [];
      if (grid.length === rect.map.width) setColumnWidths(tr, tablePos, grid.map((w) => w / TWIPS_PER_PX));
    }
    dispatch?.(tr);
    return true;
  },
  isEnabled: isInTable,
});

/** The narrowest a column may be set, in pixels: the view's drag floor (`CELL_MIN_WIDTH` in the table view), kept equal. */
export const MIN_COLUMN_WIDTH_PX = 25;
const PX_PER_PT = 20 / TWIPS_PER_PX;

/** Each column's width in points as the table at the selection lays out, for the optimiser's "before" and a no-op check (ED-007). */
export function columnWidthsPt(state: EditorState): number[] | undefined {
  if (!isInTable(state)) return undefined;
  const rect = selectedRect(state);
  return columnWidthsPx(rect.table, rect.map, textWidthTwips(state, rect.tableStart)).map((w) => w / PX_PER_PT);
}

/** A row's `w:trHeight` from its properties attribute (plain JSON), as `{ hRule, val }`, where it has one. */
export function rowHeightOf(trPr: unknown): { hRule?: string; val?: number } | undefined {
  const items = (trPr as { cnfStyleOrDivIdOrGridBefore?: { name?: { localPart?: string }; value?: { hRule?: string; val?: number } }[] } | null | undefined)?.cnfStyleOrDivIdOrGridBefore;
  const item = items?.find((i) => i.name?.localPart === 'trHeight');
  return item ? { hRule: item.value?.hRule, val: item.value?.val } : undefined;
}

/** The row's properties without their `w:trHeight`; null when nothing else is left. */
export function withoutRowHeight(trPr: unknown): unknown | null {
  if (!rowHeightOf(trPr)) return trPr;
  const pr = trPr as Record<string, unknown> & { cnfStyleOrDivIdOrGridBefore: { name?: { localPart?: string } }[] };
  const rest = pr.cnfStyleOrDivIdOrGridBefore.filter((i) => i.name?.localPart !== 'trHeight');
  const out: Record<string, unknown> = { ...pr };
  if (rest.length > 0) out['cnfStyleOrDivIdOrGridBefore'] = rest; else delete out['cnfStyleOrDivIdOrGridBefore'];
  return Object.keys(out).some((k) => k !== 'TYPE_NAME') ? out : null;
}

register<{ widths: number[]; rowHeights?: 'content' | 'keep' }>({
  name: 'table.columnWidths',
  tracking: 'untracked',
  label: 'Column widths',
  args: {
    widths: "every column's width in points, left to right, one per grid column",
    rowHeights: "'content' removes every row's stated height, exact or minimum, so each row takes its content's height; default 'keep'",
  },
  // ED-007 section 3.1: the whole list in one transaction through `setColumnWidths` (every cell and the grid), the
  // table's own width following a fixed table's sum as Insert Right's does. A list that is the table's own writes
  // nothing, so a second press of the optimiser and an unchanged file stay as they are. Editor-only: Office JS sets
  // widths cell by cell (`TableColumn.setWidth`), a case per width for the agreement suite; untracked, as Word's own
  // column operations are (ED-005 section 12.82 item 4).
  // Row heights (section 3.5 as revised 2026-10-14): `rowHeights: 'content'` removes every row's `w:trHeight`, exact or
  // minimum, in the same transaction, so a fit's rows end at their content's height; the one the file had is on the
  // row's `trPr` attribute, and the export drops it from the rebuilt row.
  run: ({ widths, rowHeights = 'keep' }) => (state, dispatch) => {
    if (!isInTable(state) || !Array.isArray(widths)) return false;
    const rect = selectedRect(state);
    if (widths.length !== rect.map.width || widths.some((w) => typeof w !== 'number' || !Number.isFinite(w))) return false;
    const px = widths.map((w) => Math.max(MIN_COLUMN_WIDTH_PX, w * PX_PER_PT));
    const current = columnWidthsPx(rect.table, rect.map, textWidthTwips(state, rect.tableStart));
    const widthsChanged = columnWidthsAside(rect.table) || !px.every((w, i) => Math.abs(w - current[i]!) < 0.5);
    const tablePos = rect.tableStart - 1;
    const heightRows: number[] = [];
    if (rowHeights === 'content') rect.table.forEach((row, offset) => { if (rowHeightOf(row.attrs['trPr'])) heightRows.push(tablePos + 1 + offset); });
    if (!widthsChanged && heightRows.length === 0) return true;
    const tr = state.tr;
    for (const pos of heightRows) {
      const row = tr.doc.nodeAt(pos)!;
      tr.setNodeMarkup(pos, undefined, { ...row.attrs, trPr: withoutRowHeight(row.attrs['trPr']) });
    }
    if (widthsChanged) {
      setColumnWidths(tr, tablePos, px);
      const table = tr.doc.nodeAt(tablePos)!;
      const tblPr = (table.attrs['tblPr'] ?? {}) as { tblW?: { w?: number; type?: string } };
      if (tblPr.tblW?.type === 'dxa') {
        const sum = px.reduce((a, b) => a + b, 0);
        tr.setNodeMarkup(tablePos, undefined, { ...table.attrs, tblPr: { ...tblPr, tblW: { ...tblPr.tblW, w: Math.round(sum * TWIPS_PER_PX) } } });
      }
    }
    tr.setSelection(Selection.near(tr.doc.resolve(tr.mapping.map(state.selection.from))));
    dispatch?.(tr);
    return true;
  },
  isEnabled: isInTable,
});

register<{ count?: number }>({
  name: 'table.deleteColumns',
  tracking: 'untracked',
  label: 'Delete columns',
  args: { count: 'columns to delete from the selection rightwards, default 1' },
  run: ({ count = 1 } = {}) => repeated(deleteColumn, count, true),
  isEnabled: isInTable,
});

register<void>({
  name: 'table.mergeCells',
  tracking: 'untracked',
  label: 'Merge cells',
  run: () => repeated(mergeCells),
  isEnabled: (state) => state.selection instanceof CellSelection && mergeCells(state),
});

register<void>({
  name: 'table.splitCell',
  tracking: 'untracked',
  label: 'Split cell',
  run: () => repeated(splitCell),
  isEnabled: (state) => isInTable(state) && splitCell(state),
});

register<void>({
  name: 'table.delete',
  label: 'Delete table',
  run: () => deleteTable,
  isEnabled: isInTable,
});

register<{ value: string | null }>({
  name: 'table.style',
  tracking: 'untracked',
  label: 'Table style',
  args: { value: 'the table style id, or null for none' },
  run: ({ value }) => (state, dispatch) => {
    const t = tableAround(state);
    if (!t) return false;
    const tblPr = { ...((t.node.attrs['tblPr'] as Record<string, unknown> | null) ?? { TYPE_NAME: 'org_docx4j_wml.TblPr' }) };
    if (value === null) delete tblPr['tblStyle']; else tblPr['tblStyle'] = { val: value };
    dispatch?.(state.tr.setNodeMarkup(t.pos, undefined, { ...t.node.attrs, tblPr }));
    return true;
  },
  isEnabled: isInTable,
});

register<{ value?: boolean | 'toggle' }>({
  name: 'row.header',
  tracking: 'untracked',
  label: 'Header row',
  args: { value: "true, false or 'toggle' (the default): repeat this row at the top of every page (w:tblHeader)" },
  choices: { value: ['toggle', 'true', 'false'] },
  run: ({ value = 'toggle' } = {}) => (state, dispatch) => {
    if (!isInTable(state)) return false;
    const rect = selectedRect(state);
    const tr = state.tr;
    for (let r = rect.top; r < rect.bottom; r++) {
      const row = rect.table.child(r);
      let rowPos = rect.tableStart;
      for (let i = 0; i < r; i++) rowPos += rect.table.child(i).nodeSize;
      const trPr = { ...((row.attrs['trPr'] as { cnfStyleOrDivIdOrGridBefore?: unknown[] } | null) ?? { TYPE_NAME: 'org_docx4j_wml.TrPr' }) } as { cnfStyleOrDivIdOrGridBefore?: { name: { namespaceURI: string; localPart: string }; value: unknown }[] };
      const items = [...(trPr.cnfStyleOrDivIdOrGridBefore ?? [])];
      const at = items.findIndex((it) => it.name.localPart === 'tblHeader');
      const on = value === 'toggle' ? at < 0 : value;
      if (on && at < 0) items.push({ name: { namespaceURI: W_NS, localPart: 'tblHeader' }, value: { TYPE_NAME: 'org_docx4j_wml.BooleanDefaultTrue' } });
      if (!on && at >= 0) items.splice(at, 1);
      trPr.cnfStyleOrDivIdOrGridBefore = items;
      tr.setNodeMarkup(rowPos, undefined, { ...row.attrs, trPr: items.length > 0 ? trPr : (row.attrs['trPr'] ? { ...trPr, cnfStyleOrDivIdOrGridBefore: [] } : null) });
    }
    dispatch?.(tr);
    return true;
  },
  isEnabled: isInTable,
  isActive: (state) => {
    if (!isInTable(state)) return false;
    const rect = selectedRect(state);
    const row = rect.table.child(rect.top);
    return ((row.attrs['trPr'] as { cnfStyleOrDivIdOrGridBefore?: { name: { localPart: string } }[] } | null)?.cnfStyleOrDivIdOrGridBefore ?? []).some((it) => it.name.localPart === 'tblHeader');
  },
});

/**
 * A cell's `tcPr` attribute with its own shading set to `fill` (six hex digits, upper case) or taken away (null), in the form
 * Office JS's `shadingColor` writes in Word 365: `w:val="clear"` with `w:color` and `w:fill` both the colour, any theme fill
 * dropped with the colour it named (core-ts CR-002 section 67.3; ED-005 section 12.130). With `clear` the pattern colour isn't
 * drawn, so it looks as `w:color="auto"`, the form before, did.
 */
export function withCellShading(tcPr: unknown, fill: string | null): unknown | null {
  const next: Record<string, unknown> = { ...((tcPr as Record<string, unknown> | null) ?? { TYPE_NAME: 'org_docx4j_wml.TcPr' }) };
  if (fill === null) delete next['shd'];
  else next['shd'] = { TYPE_NAME: 'org_docx4j_wml.CTShd', val: 'clear', color: fill, fill };
  return Object.keys(next).every((k) => k === 'TYPE_NAME') ? null : next;
}

/** A colour as `w:fill` writes it: `#RRGGBB` or `RRGGBB` to six upper-case digits; undefined for anything else. */
function fillOf(color: string): string | undefined {
  const m = /^#?([0-9A-Fa-f]{6})$/.exec(color.trim());
  return m ? m[1]!.toUpperCase() : undefined;
}

/**
 * The shading the selected rows' cells hold as their own, as `#RRGGBB`: null when none has any, 'Mixed' when they differ
 * (a pattern or a theme fill without a hex counts as its `w:fill` alone). Undefined outside a table.
 */
export function rowShadingOf(state: EditorState): string | null | 'Mixed' | undefined {
  if (!isInTable(state)) return undefined;
  let out: string | null | undefined;
  for (const { node } of selectedRowCells(state)) {
    const fill = (node.attrs['tcPr'] as { shd?: { fill?: string } } | null)?.shd?.fill;
    const css = fill && /^[0-9A-Fa-f]{6}$/.test(fill) ? `#${fill.toUpperCase()}` : null;
    if (out === undefined) out = css;
    else if (out !== css) return 'Mixed';
  }
  return out ?? null;
}

/** Every cell in the rows the selection covers, once each (a merged cell reaching into them included), with its position. */
function selectedRowCells(state: EditorState): { node: PMNode; pos: number }[] {
  const rect = selectedRect(state);
  const seen = new Set<number>();
  const out: { node: PMNode; pos: number }[] = [];
  for (let r = rect.top; r < rect.bottom; r++) {
    for (let c = 0; c < rect.map.width; c++) {
      const rel = rect.map.map[r * rect.map.width + c]!;
      if (seen.has(rel)) continue;
      seen.add(rel);
      out.push({ node: rect.table.nodeAt(rel)!, pos: rect.tableStart + rel });
    }
  }
  return out;
}

register<{ color: string | null }>({
  name: 'row.shading',
  // Word tracks it as a cell's w:tcPrChange, a revision the model holds only as shown in Word (ED-005 section 12.128).
  tracking: 'untracked',
  label: 'Row shading',
  args: { color: "the background colour of the selected rows' cells as '#RRGGBB', or null to take their own shading away" },
  // Word's Table Design, Shading, over whole rows: every cell of each row the selection covers gets the fill as its own
  // w:shd. The api line is Office JS's TableRow.shadingColor for the first selected paragraph's row (core-ts CR-002 section
  // 67): '' takes the cells' own shading away, as null does here. A merged cell differs: the editor shades it whole from any
  // row it reaches, writing every w:tc of it, where Office JS writes the row's own w:tcs alone (ED-005 section 12.130).
  api: ({ color }) => `selection.paragraphs[0].parentTableCell.parentRow.shadingColor = ${JSON.stringify(color ?? '')};`,
  run: ({ color }) => (state, dispatch) => {
    if (!isInTable(state)) return false;
    const fill = color === null || color === undefined || color === '' ? null : fillOf(color);
    if (fill === undefined) return false;
    const tr = state.tr;
    for (const { node, pos } of selectedRowCells(state)) {
      const tcPr = withCellShading(node.attrs['tcPr'], fill);
      if (!plainEqual(tcPr, node.attrs['tcPr'] ?? null)) tr.setNodeMarkup(pos, undefined, { ...node.attrs, tcPr });
    }
    if (!tr.docChanged) return true;
    dispatch?.(tr);
    return true;
  },
  isEnabled: isInTable,
});

/** Tab in a table: the next cell, and a new row after the last cell, as Word. Outside a table it does nothing (the keymap falls through to `insertTab`). */
register<void>({
  name: 'table.nextCell',
  label: 'Next cell',
  run: () => (state, dispatch, view) => {
    if (!isInTable(state)) return false;
    if (goToNextCell(1)(state, dispatch, view)) return true;
    // The last cell: add a row, then move into it, as one transaction.
    return repeated((s, d, v) => {
      let out: Transaction | undefined;
      if (!addRowAfter(s, (tr) => { out = tr; })) return false;
      const after = s.apply(out!);
      let moved: Transaction | undefined;
      if (!goToNextCell(1)(after, (tr) => { moved = tr; }, v)) { d?.(out!); return true; }
      for (const step of moved!.steps) out!.step(step);
      out!.setSelection(remap(after.apply(moved!).selection, out!.doc));
      d?.(out!);
      return true;
    })(state, dispatch, view);
  },
  isEnabled: isInTable,
});

register<void>({
  name: 'table.previousCell',
  label: 'Previous cell',
  run: () => (state, dispatch, view) => isInTable(state) && goToNextCell(-1)(state, dispatch, view),
  isEnabled: isInTable,
});

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export { TableMap, CellSelection, isInTable };
