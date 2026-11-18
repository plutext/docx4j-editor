// SPDX-License-Identifier: GPL-3.0-or-later
//
// Import: the typed tree of a body (or any block container) to ProseMirror nodes (ED-002 section
// 3.4; ED-001 section 3). The walk is the engine's: block-level elements, then `runItemsOf` per
// paragraph. Tracked changes in the body are content since E4.c (ED-005 section 5.2): the runs of a
// `w:ins` or `w:moveTo` carry an insertion mark, those of a `w:del` or `w:moveFrom` are text with a
// deletion mark, `w:rPrChange` is a mark, a paragraph mark's, a row's and a paragraph's properties'
// changes are attributes, an atom's is its `revision` attribute. A story (a note, a header, a
// comment) keeps E1's projection of them, the insertion shown as text and the deletion a hidden
// atom (decision 2), since E4 leaves their revisions as they are (ED-005 section 10).
// Every element that becomes a node or a `runSource` mark gets a `sourceId` (ids.mts); the run
// holders between the containing node and a run (hyperlinks, insertions, smart tags) are recorded
// so that export can put a rebuilt run back inside them.
//
// `projectBlocks` takes a block range, not only a body, because ED-001 decision 20's protocol
// re-imports only the blocks a tree mutation changed.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { typeNameOf, runItemsOf, revisionKindOf, readRunOptions, selectedBranchOf, isAlternateContent, type Element, type RunFormatting } from '@docx4j/core-ts';
import type * as mce from '@docx4j/generated-objects-ts/modules/org_docx4j_mce';
import { Node as PMNode, Mark, type Schema } from 'prosemirror-model';
import { SourceIds } from './ids.mjs';
import { toPlain } from './json.mjs';
import { type Diagnostics, emptyDiagnostics } from './diagnostics.mjs';
import type { IndentAttrs, ParagraphAttrs, SpacingAttrs, Revision, RevisionInfo } from './schema.mjs';
import { revisionInfo, revisionOf, withoutRevision, propertiesChangeOf, formatChangeAttrs } from './revisions.mjs';
import { inlineOf } from './drawings.mjs';

/** What export needs to know about a run or a paragraph-level item it did not model. */
export interface SourceInfo {
  /** The run-holder elements between the containing node and this element, outermost first. */
  holders: Element[];
}

/** The comment range markers a paragraph held, so a rebuilt paragraph re-emits them. */
export interface ParagraphInfo {
  commentStarts: Map<string, Element>;
  commentEnds: Map<string, Element>;
  /** The comments whose start marker lay inside a tracked insertion or deletion (Word check 32, ED-005 section 12.49). */
  startsInRevision?: Set<string>;
}

/** What export needs to rebuild a table it did not model node for node (ED-003 section 3.2). */
export interface TableInfo {
  /** Row-level items that are not rows (bookmarks, permissions between rows), by the row they precede. */
  rowExtras: { beforeRow: number; element: Element }[];
  /** Per imported row, by its sourceId. */
  rows: Map<string, RowInfo>;
  /** A `w:vMerge` continuation cell, by the restart cell's sourceId and its row offset below it. */
  continuations: Map<string, Element>;
}

export interface RowInfo {
  /** Cell-level items that are not cells, by the cell they precede. */
  cellExtras: { beforeCell: number; element: Element }[];
  /** The columns covered from a row above at import: `${restartSourceId}:${offset}` per grid column. */
  covered: string[];
}

/** A row- or cell-level holder as imported (ED-003 section 4.2): the session writes the `sdt` ones into the `controls` map. */
export interface GroupInfo {
  form: 'row' | 'cell';
  /** `w:sdt` (a content control) or `w:customXml` (a custom XML wrapper, which has no properties). */
  kind: 'sdt' | 'customXml';
  /** `w:sdtPr` as plain JSON; null for a custom XML wrapper or a control without properties. */
  sdtPr: unknown | null;
}

export interface ProjectionContext {
  schema: Schema;
  ids: SourceIds;
  /** Runs, paragraph-level items and inline content controls, by sourceId. */
  sources: Map<string, SourceInfo>;
  paragraphs: Map<string, ParagraphInfo>;
  tables: Map<string, TableInfo>;
  /** Row- and cell-level holders by their group id (the holder's sourceId), in order of first appearance. */
  groups: Map<string, GroupInfo>;
  /** The node each block or inline content control was imported as, by sourceId: the "unchanged" test on export. */
  imported: Map<string, PMNode>;
  diagnostics: Diagnostics;
  /** Comment ranges open at the current point of the walk (a range may span paragraphs). */
  openComments: string[];
  commentIds: Set<string>;
  /** The largest annotation `w:id` seen (revisions, their range markers, bookmarks), for the revisions export writes afresh. */
  maxRevisionId: number;
  /**
   * `content` (the body): tracked changes projected as content (ED-005 section 5.2). `hidden` (a
   * story): E1's projection, an insertion's runs as text and a deletion a hidden atom.
   */
  revisions: 'content' | 'hidden';
  /** The move ranges open at the current point of the walk, by id, with their names (`w:moveFromRangeStart`, `w:moveToRangeStart`). */
  openMoves: { from: Map<string, string>; to: Map<string, string> };
}

export function newContext(schema: Schema, ids: SourceIds = new SourceIds(), revisions: ProjectionContext['revisions'] = 'content'): ProjectionContext {
  return {
    schema, ids, sources: new Map(), paragraphs: new Map(), tables: new Map(), groups: new Map(), imported: new Map(),
    diagnostics: emptyDiagnostics(), openComments: [], commentIds: new Set(), maxRevisionId: 0,
    revisions, openMoves: { from: new Map(), to: new Map() },
  };
}

/** Count a tracked change kept as it is (ED-005 section 5.2's preserved list). */
function preserved(ctx: ProjectionContext, kind: string, present: unknown): void {
  if (present === undefined || present === null) return;
  ctx.diagnostics.tracked.preserved[kind] = (ctx.diagnostics.tracked.preserved[kind] ?? 0) + 1;
  noteId(ctx, (present as { id?: unknown }).id);
}

/** An annotation id seen: the export's fresh ids start above the largest. */
function noteId(ctx: ProjectionContext, id: unknown): void {
  const n = typeof id === 'number' ? id : typeof id === 'string' ? Number(id) : NaN;
  if (Number.isFinite(n) && n > ctx.maxRevisionId) ctx.maxRevisionId = n;
}

/** The name of the move range last opened and still open, `''` for a move whose range has no name, undefined when none is open. */
function openMove(ranges: Map<string, string>): string | undefined {
  let name: string | undefined;
  for (const v of ranges.values()) name = v;
  return name;
}

/** A move range marker (inline or between blocks) opening or closing its range for the walk. */
function trackMoveRange(el: Element, ctx: ProjectionContext): void {
  const local = el.name.localPart;
  const v = el.value as { id?: unknown; name?: string };
  const id = String(v.id ?? '');
  noteId(ctx, v.id);
  if (local === 'moveFromRangeStart') ctx.openMoves.from.set(id, v.name ?? '');
  else if (local === 'moveToRangeStart') ctx.openMoves.to.set(id, v.name ?? '');
  else if (local === 'moveFromRangeEnd') ctx.openMoves.from.delete(id);
  else if (local === 'moveToRangeEnd') ctx.openMoves.to.delete(id);
  else if (local === 'bookmarkStart' || local === 'bookmarkEnd') noteId(ctx, v.id);
}

const MOVE_RANGE = new Set(['moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd', 'bookmarkStart', 'bookmarkEnd']);

/**
 * A paragraph's attributes as the projection gives them in a context: `paragraphAttrsOf`, and in
 * the body the mark's tracked change read out of `markRPr` into `markRevision` and the properties'
 * change into `pPrChange` (ED-005 section 5.2). A mark with no tracked change keeps `markRPr` as E1
 * made it, so a document without revisions projects as it always has. `moves` names the move a
 * moved mark belongs to (the ranges open at the paragraph's end).
 */
export function projectedParagraphAttrs(pPr: wml.PPr | undefined, revisions: ProjectionContext['revisions'], moves: { from: string | null; to: string | null } = { from: null, to: null }): Omit<ParagraphAttrs, 'sourceId'> {
  const attrs = paragraphAttrsOf(pPr);
  if (revisions !== 'content') return attrs;
  const markRevision = revisionOf(pPr?.rPr, moves);
  if (markRevision) { attrs.markRevision = markRevision; attrs.markRPr = withoutRevision(pPr?.rPr); }
  attrs.pPrChange = propertiesChangeOf(pPr?.pPrChange);
  return attrs;
}

/** The side of a revision a `w:ins`, `w:del`, `w:moveTo` or `w:moveFrom` element makes, with the move it belongs to. */
function sideOf(kind: 'ins' | 'del' | 'moveFrom' | 'moveTo', tc: wml.CTTrackChange, ctx: ProjectionContext): RevisionInfo {
  if (kind === 'moveTo') return revisionInfo(tc, openMove(ctx.openMoves.to) ?? '');
  if (kind === 'moveFrom') return revisionInfo(tc, openMove(ctx.openMoves.from) ?? '');
  return revisionInfo(tc, null);
}

export interface Projection {
  doc: PMNode;
  ctx: ProjectionContext;
}

/** The whole body as a `doc` node. */
export function projectBody(body: wml.Body, schema: Schema, ids?: SourceIds): Projection {
  const ctx = newContext(schema, ids);
  let blocks = projectBlocks(body.content ?? [], ctx, 'body');
  if (blocks.length === 0) blocks = [schema.nodes['paragraph']!.create()];
  ctx.diagnostics.comments = ctx.commentIds.size;
  if (ctx.revisions === 'content') preserved(ctx, 'sectPrChange', body.sectPr?.sectPrChange);
  const doc = schema.topNodeType.create({ finalSectPr: body.sectPr ? toPlain(body.sectPr) : null }, blocks);
  return { doc, ctx };
}

const P = 'org_docx4j_wml.P';
const TBL = 'org_docx4j_wml.Tbl';
const TR = 'org_docx4j_wml.Tr';
const TC = 'org_docx4j_wml.Tc';
/** Row- and cell-level holders: content controls and custom XML around rows or cells (the engine's ROW_HOLDERS / CELL_HOLDERS). */
const ROW_HOLDERS = new Set(['org_docx4j_wml.CTSdtRow', 'org_docx4j_wml.CTCustomXmlRow']);
const CELL_HOLDERS = new Set(['org_docx4j_wml.CTSdtCell', 'org_docx4j_wml.CTCustomXmlCell']);
const SDT_BLOCK = 'org_docx4j_wml.SdtBlock';
const SDT_RUN = 'org_docx4j_wml.SdtRun';
const RUN = 'org_docx4j_wml.R';
const HYPERLINK = 'org_docx4j_wml.P.Hyperlink';
const SIMPLE_FIELD = 'org_docx4j_wml.CTSimpleField';

/** Run holders that are transparent for editing: their runs are ordinary text, the wrapper is restored on export. */
const TRANSPARENT_HOLDERS = new Set([
  HYPERLINK, 'org_docx4j_wml.CTSmartTagRun', 'org_docx4j_wml.CTCustomXmlRun', 'org_docx4j_wml.P.Dir', 'org_docx4j_wml.P.Bdo',
]);

/** Block-level elements of a container (body, sdt content) to block nodes, in order. */
export function projectBlocks(elements: Element[], ctx: ProjectionContext, address: string): PMNode[] {
  const out: PMNode[] = [];
  elements.forEach((el, i) => {
    const tn = typeNameOf(el);
    const addr = `${address}/${i}`;
    let node: PMNode;
    if (tn === P) node = projectParagraph(el as Element<wml.P>, ctx, addr);
    else if (tn === SDT_BLOCK) node = projectSdtBlock(el as Element<wml.SdtBlock>, ctx, addr);
    else if (tn === TBL) node = projectTable(el as Element<wml.Tbl>, ctx, addr);
    else {
      // A move range or a bookmark between blocks opens or closes its range for what follows.
      if (MOVE_RANGE.has(el.name.localPart)) trackMoveRange(el, ctx);
      node = opaqueBlock(el, ctx, addr);
    }
    out.push(node);
  });
  return out;
}

function opaqueBlock(el: Element, ctx: ProjectionContext, address: string): PMNode {
  const sourceId = ctx.ids.idOf(el);
  const kind = el.name.localPart;
  ctx.diagnostics.opaque.push({ kind, address, inline: false, sourceId });
  const node = ctx.schema.nodes['opaqueBlock']!.create({ sourceId, kind });
  ctx.imported.set(sourceId, node);
  return node;
}

function projectSdtBlock(el: Element<wml.SdtBlock>, ctx: ProjectionContext, address: string): PMNode {
  const sourceId = ctx.ids.idOf(el);
  const sdt = el.value;
  let children = projectBlocks(sdt.sdtContent?.content ?? [], ctx, `${address}/sdt`);
  if (children.length === 0) children = [ctx.schema.nodes['paragraph']!.create()];
  const node = ctx.schema.nodes['sdtBlock']!.create({ sourceId, sdtPr: sdt.sdtPr ? toPlain(sdt.sdtPr) : null }, children);
  ctx.imported.set(sourceId, node);
  return node;
}

function on(v: { val?: boolean } | undefined): boolean {
  return v !== undefined && v.val !== false;
}

/** Twips to the CSS pixels `prosemirror-tables` keeps column widths in (96 per inch, 1440 twips per inch). */
export const TWIPS_PER_PX = 15;

/** A row-level or cell-level control group a row or cell sits in, outermost first. */
export interface SdtGroup {
  groupId: string;
  first: boolean;
  last: boolean;
}

interface RowDescriptor {
  element: Element<wml.Tr>;
  stack: string[];
  cells: CellDescriptor[];
  cellExtras: { beforeCell: number; element: Element }[];
  covered: string[];
}

interface CellDescriptor {
  element: Element<wml.Tc>;
  stack: string[];
  col: number;
  colspan: number;
  rowspan: number;
  merged: string[];
  /** A `w:vMerge` continuation: no node, folded into the restart cell above. */
  continuation: boolean;
}

/**
 * `w:tbl` to a `table` node (ED-003 section 3.3): rows through their row-level controls, cells
 * through theirs, `w:gridSpan` to `colspan`, a `w:vMerge` run to one cell with `rowspan`.
 * Anything between rows or cells that is not a row or a cell (a bookmark) is side data that
 * export puts back.
 */
function projectTable(el: Element<wml.Tbl>, ctx: ProjectionContext, address: string): PMNode {
  const sourceId = ctx.ids.idOf(el);
  const tbl = el.value;
  const info: TableInfo = { rowExtras: [], rows: new Map(), continuations: new Map() };
  ctx.tables.set(sourceId, info);
  // Rows, through the row-level holders, each with the stack of holders it sits in.
  const rows: RowDescriptor[] = [];
  const visitRows = (items: Element[] | undefined, stack: string[]): void => {
    for (const item of items ?? []) {
      const tn = typeNameOf(item);
      if (tn === TR) rows.push({ element: item as Element<wml.Tr>, stack, cells: [], cellExtras: [], covered: [] });
      else if (tn !== undefined && ROW_HOLDERS.has(tn)) visitRows(childrenOfHolder(item), [...stack, registerGroup(item, 'row', ctx)]);
      else if (item.name.localPart === 'proofErr') continue;
      else info.rowExtras.push({ beforeRow: rows.length, element: item });
    }
  };
  visitRows(tbl.content as Element[] | undefined, []);
  // Cells, with the grid column each starts at, and vertical merges folded into their restart cell.
  const open = new Map<number, CellDescriptor & { rowIndex: number }>();
  rows.forEach((row, r) => {
    const cells: CellDescriptor[] = [];
    const visitCells = (items: Element[] | undefined, stack: string[]): void => {
      for (const item of items ?? []) {
        const tn = typeNameOf(item);
        if (tn === TC) cells.push({ element: item as Element<wml.Tc>, stack, col: 0, colspan: 1, rowspan: 1, merged: [], continuation: false });
        else if (tn !== undefined && CELL_HOLDERS.has(tn)) visitCells(childrenOfHolder(item), [...stack, registerGroup(item, 'cell', ctx)]);
        else if (item.name.localPart === 'proofErr') continue;
        else row.cellExtras.push({ beforeCell: cells.length, element: item });
      }
    };
    visitCells(row.element.value.content as Element[] | undefined, []);
    let col = gridBeforeOf(row.element.value.trPr);
    const stillOpen = new Set<number>();
    for (const cell of cells) {
      const pr = cell.element.value.tcPr;
      cell.col = col;
      cell.colspan = Math.max(1, Number(pr?.gridSpan?.val ?? 1));
      const vMerge = pr?.vMerge;
      const above = open.get(col);
      if (vMerge !== undefined && vMerge.val !== 'restart' && above && above.colspan === cell.colspan) {
        // A continuation of the merge above: folded into the restart cell.
        cell.continuation = true;
        above.rowspan++;
        const cellId = ctx.ids.idOf(cell.element);
        above.merged.push(cellId);
        const restartId = ctx.ids.idOf(above.element);
        info.continuations.set(`${restartId}:${r - above.rowIndex}`, cell.element);
        for (let c = col; c < col + cell.colspan; c++) row.covered[c] = `${restartId}:${r - above.rowIndex}`;
        stillOpen.add(col);
      } else {
        if (vMerge !== undefined && vMerge.val !== 'restart') ctx.diagnostics.opaque.push({ kind: 'vMerge', address: `${address}/${r}`, inline: true, sourceId: ctx.ids.idOf(cell.element) });
        if (vMerge?.val === 'restart') { open.set(col, { ...cell, rowIndex: r }); stillOpen.add(col); }
        else open.delete(col);
      }
      col += cell.colspan;
    }
    for (const c of [...open.keys()]) if (!stillOpen.has(c)) open.delete(c);
    // The restart descriptors in `open` are copies; write the counts back to the row's own descriptors.
    for (const o of open.values()) {
      const target = rows[o.rowIndex]!.cells.find((d) => d.element === o.element);
      if (target) { target.rowspan = o.rowspan; target.merged = o.merged; }
    }
    row.cells = cells;
  });
  for (const o of open.values()) {
    const target = rows[o.rowIndex]!.cells.find((d) => d.element === o.element);
    if (target) { target.rowspan = o.rowspan; target.merged = o.merged; }
  }
  // Nodes. A group's `first` and `last` rows are the extremes of its contiguous run.
  const firstOf = new Map<string, number>();
  const lastOf = new Map<string, number>();
  rows.forEach((row, r) => { for (const g of row.stack) { if (!firstOf.has(g)) firstOf.set(g, r); lastOf.set(g, r); } });
  if (ctx.revisions === 'content') { preserved(ctx, 'tblPrChange', tbl.tblPr?.tblPrChange); preserved(ctx, 'tblGridChange', tbl.tblGrid?.tblGridChange); }
  // A cell's width in twips (`w:tcW dxa`); a percentage cell's is its columns' share of the grid (Word's cache of its
  // last layout; check 39, ED-005 section 12.79), so that the export can write the share back and the save's self-check
  // reads it the same. An automatic cell (`w:tcW auto`, AutoFit to contents: check 42, section 12.83) and a cell with no
  // `w:tcW` have none: the content sizes them.
  const gridTwips = (tbl.tblGrid?.gridCol ?? []).map((c) => c.w ?? 0);
  const rowNodes = rows.map((row, r) => {
    const rowId = ctx.ids.idOf(row.element);
    info.rows.set(rowId, { cellExtras: row.cellExtras, covered: row.covered });
    const cellNodes: PMNode[] = [];
    let column = 0;
    const cellFirst = new Map<string, number>();
    const cellLast = new Map<string, number>();
    row.cells.forEach((cell, i) => { for (const g of cell.stack) { if (!cellFirst.has(g)) cellFirst.set(g, i); cellLast.set(g, i); } });
    row.cells.forEach((cell, i) => {
      const startColumn = column;
      column += cell.colspan;
      if (cell.continuation) return;
      const cellId = ctx.ids.idOf(cell.element);
      let blocks = projectBlocks((cell.element.value.content ?? []) as Element[], ctx, `${address}/${r}/${i}`);
      if (blocks.length === 0) blocks = [ctx.schema.nodes['paragraph']!.create()];
      const tcW = cell.element.value.tcPr?.tcW;
      if (ctx.revisions === 'content') {
        const pr = cell.element.value.tcPr as (wml.TcPr & { tcPrChange?: unknown; cellIns?: unknown; cellDel?: unknown; cellMerge?: unknown }) | undefined;
        preserved(ctx, 'tcPrChange', pr?.tcPrChange);
        preserved(ctx, 'cellIns', pr?.cellIns);
        preserved(ctx, 'cellDel', pr?.cellDel);
        preserved(ctx, 'cellMerge', pr?.cellMerge);
      }
      const width = tcW?.type === 'dxa' && typeof tcW.w === 'number' && tcW.w > 0 ? tcW.w / TWIPS_PER_PX : null;
      const fromGrid = gridTwips.slice(startColumn, startColumn + cell.colspan);
      // A spanning cell's width is shared over its columns as the grid shares them (ED-007 section 8.1 item 11: an
      // equal split drew a fit table at its old widths after a reload, the first row being what a fixed layout reads),
      // equally where the grid cannot say.
      const gridKnown = fromGrid.length === cell.colspan && fromGrid.every((w) => w > 0);
      const gridSum = fromGrid.reduce((a, b) => a + b, 0);
      const colwidth = width !== null ? (gridKnown ? fromGrid.map((w) => (width * w) / gridSum) : Array.from({ length: cell.colspan }, () => width / cell.colspan))
        : tcW?.type === 'pct' && gridKnown ? fromGrid.map((w) => w / TWIPS_PER_PX) : null;
      const node = ctx.schema.nodes['table_cell']!.create({
        sourceId: cellId,
        colspan: cell.colspan,
        rowspan: cell.rowspan,
        colwidth,
        tcPr: cellPropertiesAttr(cell.element.value.tcPr),
        sdt: cell.stack.length > 0 ? cell.stack.map((g): SdtGroup => ({ groupId: g, first: cellFirst.get(g) === i, last: cellLast.get(g) === i })) : null,
        merged: cell.merged.length > 0 ? cell.merged : null,
      }, blocks);
      ctx.imported.set(cellId, node);
      cellNodes.push(node);
    });
    // The row's tracked insertion or deletion is its `revision` (ED-005 section 5.2), read out of `trPr`.
    const trPr = row.element.value.trPr;
    const revision = ctx.revisions === 'content' ? revisionOf(trPr) : null;
    if (revision) { ctx.diagnostics.tracked.shown++; noteId(ctx, trPr?.ins?.id); noteId(ctx, trPr?.del?.id); }
    if (ctx.revisions === 'content') preserved(ctx, 'trPrChange', trPr?.trPrChange);
    const node = ctx.schema.nodes['table_row']!.create({
      sourceId: rowId,
      trPr: revision ? withoutRevision(trPr) : trPr ? toPlain(trPr) : null,
      sdt: row.stack.length > 0 ? row.stack.map((g): SdtGroup => ({ groupId: g, first: firstOf.get(g) === r, last: lastOf.get(g) === r })) : null,
      ...(revision ? { revision } : {}),
    }, cellNodes);
    ctx.imported.set(rowId, node);
    return node;
  });
  if (rowNodes.length === 0) {
    // A table with no rows cannot be shown as one: opaque, reported.
    ctx.tables.delete(sourceId);
    return opaqueBlock(el, ctx, address);
  }
  if (!tbl.tblGrid?.gridCol?.length) ctx.diagnostics.opaque.push({ kind: 'tblGrid', address, inline: true, sourceId });
  const node = ctx.schema.nodes['table']!.create({
    sourceId,
    tblPr: tbl.tblPr ? toPlain(tbl.tblPr) : null,
    tblGrid: tbl.tblGrid?.gridCol ? tbl.tblGrid.gridCol.map((c) => c.w ?? 0) : null,
  }, rowNodes);
  ctx.imported.set(sourceId, node);
  return node;
}

/**
 * A cell's `w:tcPr` as the attribute: plain JSON without the three properties the node carries
 * as `colspan`, `rowspan` and `colwidth` (`w:gridSpan`, `w:vMerge`, `w:tcW`), which export
 * writes back from those, so a rebuilt cell re-projects to the same attribute.
 */
export function cellPropertiesAttr(tcPr: wml.TcPr | undefined): unknown | null {
  if (!tcPr) return null;
  const plain = toPlain(tcPr) as Record<string, unknown>;
  delete plain['gridSpan'];
  delete plain['vMerge'];
  delete plain['tcW'];
  return Object.keys(plain).every((k) => k === 'TYPE_NAME') ? null : plain;
}

/** A row- or cell-level holder's group id, with its properties recorded for the `controls` map. */
function registerGroup(holder: Element, form: 'row' | 'cell', ctx: ProjectionContext): string {
  const id = ctx.ids.idOf(holder);
  if (!ctx.groups.has(id)) {
    const v = holder.value as { sdtPr?: unknown; sdtContent?: unknown };
    const kind = v.sdtContent !== undefined || holder.name.localPart === 'sdt' ? 'sdt' : 'customXml';
    ctx.groups.set(id, { form, kind, sdtPr: kind === 'sdt' && v.sdtPr ? toPlain(v.sdtPr) : null });
  }
  return id;
}

/**
 * A table's rows and each row's cells as the projection has them, for the console's selection binding, which walks index
 * paths from the editor's nodes into the tree (ED-005 section 12.129): row- and cell-level holders unwrapped, and a
 * `w:vMerge` continuation of an open restart of the same span left out, as the walk above folds it into the restart cell.
 */
export function projectedTableCells(tbl: wml.Tbl): Element<wml.Tc>[][] {
  const rows: Element<wml.Tr>[] = [];
  const visitRows = (items: Element[] | undefined): void => {
    for (const item of items ?? []) {
      const tn = typeNameOf(item);
      if (tn === TR) rows.push(item as Element<wml.Tr>);
      else if (tn !== undefined && ROW_HOLDERS.has(tn)) visitRows(childrenOfHolder(item));
    }
  };
  visitRows(tbl.content as Element[] | undefined);
  const open = new Map<number, number>();   // the grid column a restart opens, and its span
  return rows.map((row) => {
    const cells: Element<wml.Tc>[] = [];
    const visitCells = (items: Element[] | undefined): void => {
      for (const item of items ?? []) {
        const tn = typeNameOf(item);
        if (tn === TC) cells.push(item as Element<wml.Tc>);
        else if (tn !== undefined && CELL_HOLDERS.has(tn)) visitCells(childrenOfHolder(item));
      }
    };
    visitCells(row.value.content as Element[] | undefined);
    const out: Element<wml.Tc>[] = [];
    const stillOpen = new Set<number>();
    let col = gridBeforeOf(row.value.trPr);
    for (const cell of cells) {
      const pr = cell.value.tcPr;
      const span = Math.max(1, Number(pr?.gridSpan?.val ?? 1));
      const vMerge = pr?.vMerge;
      if (vMerge !== undefined && vMerge.val !== 'restart' && open.get(col) === span) stillOpen.add(col);
      else {
        out.push(cell);
        if (vMerge?.val === 'restart') { open.set(col, span); stillOpen.add(col); }
        else open.delete(col);
      }
      col += span;
    }
    for (const c of [...open.keys()]) if (!stillOpen.has(c)) open.delete(c);
    return out;
  });
}

/** The items a row- or cell-level holder wraps (`sdtContent.content`, or `content`). */
function childrenOfHolder(holder: Element): Element[] | undefined {
  const v = holder.value as { content?: Element[]; sdtContent?: { content?: Element[] } };
  return v.sdtContent ? v.sdtContent.content : v.content;
}

/** `w:gridBefore` of a row's properties: the grid columns its first cell skips. */
function gridBeforeOf(trPr: wml.TrPr | undefined): number {
  for (const item of trPr?.cnfStyleOrDivIdOrGridBefore ?? []) {
    if (item.name.localPart === 'gridBefore') return Number((item.value as { val?: number }).val ?? 0);
  }
  return 0;
}

/** The paragraph attributes of a `w:pPr` (the editable subset; the rest rides on the copy at export). */
/**
 * The paragraph mark's own run properties as the attribute holds them: null for no `w:rPr` and for an
 * empty one alike. pandoc writes `<w:rPr/>` in every paragraph's properties; a paragraph made in the
 * editor with a copy of such an attribute (Enter's first half) is written with no element, since an
 * empty one says nothing, and a reload must then project the same attribute (the save's self-check
 * refused a split in such a paragraph; Jason, 2026-10-17).
 */
export function markPropertiesAttr(rPr: wml.ParaRPr | undefined): unknown | null {
  if (!rPr) return null;
  const plain = toPlain(rPr) as Record<string, unknown>;
  return Object.keys(plain).every((k) => k === 'TYPE_NAME') ? null : plain;
}

export function paragraphAttrsOf(pPr: wml.PPr | undefined): Omit<ParagraphAttrs, 'sourceId'> {
  const ind = pPr?.ind;
  const sp = pPr?.spacing;
  const indent: IndentAttrs | null = ind
    ? { start: ind.start ?? ind.left ?? null, end: ind.end ?? ind.right ?? null, firstLine: ind.firstLine ?? null, hanging: ind.hanging ?? null }
    : null;
  const spacing: SpacingAttrs | null = sp
    ? { before: sp.before ?? null, after: sp.after ?? null, line: sp.line ?? null, lineRule: sp.lineRule ?? null }
    : null;
  const numPr = pPr?.numPr;
  const tabs = pPr?.tabs?.tab;
  return {
    pStyle: pPr?.pStyle?.val ?? null,
    jc: pPr?.jc?.val ?? null,
    ind: indent,
    spacing,
    bidi: on(pPr?.bidi),
    numPr: numPr ? { numId: numPr.numId?.val === undefined ? null : String(numPr.numId.val), ilvl: Number(numPr.ilvl?.val ?? 0) } : null,
    tabs: tabs ? tabs.filter((t) => t.pos !== undefined).map((t) => ({ pos: Number(t.pos), val: String(t.val ?? 'left'), leader: t.leader === undefined ? null : String(t.leader) })) : null,
    sectPr: pPr?.sectPr ? toPlain(pPr.sectPr) : null,
    markRPr: markPropertiesAttr(pPr?.rPr),
    markRevision: null,
    pPrChange: null,
    copiedFrom: null,
  };
}

function projectParagraph(el: Element<wml.P>, ctx: ProjectionContext, address: string): PMNode {
  const sourceId = ctx.ids.idOf(el);
  const info: ParagraphInfo = { commentStarts: new Map(), commentEnds: new Map() };
  ctx.paragraphs.set(sourceId, info);
  const content = projectInline(runItemsOf(el.value) ?? [], [], ctx, address, info, {});
  const pPr = el.value.pPr;
  // The mark comes after the content: a moved mark belongs to the move ranges open at the paragraph's end.
  const attrs = projectedParagraphAttrs(pPr, ctx.revisions, { from: openMove(ctx.openMoves.from) ?? null, to: openMove(ctx.openMoves.to) ?? null });
  if (ctx.revisions === 'content') {
    if (attrs.markRevision) { ctx.diagnostics.tracked.shown++; for (const k of ['ins', 'del', 'moveFrom', 'moveTo'] as const) noteId(ctx, pPr?.rPr?.[k]?.id); }
    if (attrs.pPrChange) { ctx.diagnostics.tracked.shown++; noteId(ctx, pPr?.pPrChange?.id); }
    preserved(ctx, 'markRPrChange', (pPr?.rPr as { rPrChange?: unknown } | undefined)?.rPrChange);
    preserved(ctx, 'numberingChange', (pPr?.numPr as { numberingChange?: unknown } | undefined)?.numberingChange);
    preserved(ctx, 'sectPrChange', pPr?.sectPr?.sectPrChange);
  }
  const node = ctx.schema.nodes['paragraph']!.create({ sourceId, ...attrs }, content);
  ctx.imported.set(sourceId, node);
  return node;
}

/**
 * The inline items of a paragraph, hyperlink, insertion or run-level control to inline nodes.
 * `rev` is the tracked change the items sit in (the body only): an insertion, a deletion, or both
 * (a `w:del` inside another author's `w:ins`).
 */
function projectInline(items: Element[], holders: Element[], ctx: ProjectionContext, address: string, info: ParagraphInfo, rev: Revision): PMNode[] {
  const out: PMNode[] = [];
  items.forEach((el, i) => {
    const tn = typeNameOf(el);
    const local = el.name.localPart;
    if (tn === RUN) {
      out.push(...projectRun(el as Element<wml.R>, holders, ctx, `${address}/${i}`, rev));
      return;
    }
    const kind = revisionKindOf(el);
    if (kind !== undefined) {
      noteId(ctx, (el.value as wml.CTTrackChange).id);
      if (ctx.revisions === 'content') {
        // Content (ED-005 section 5.2): the runs inside carry the change; the element goes back around them at export.
        ctx.diagnostics.tracked.shown++;
        const side = sideOf(kind, el.value as wml.CTTrackChange, ctx);
        const inner: Revision = kind === 'ins' || kind === 'moveTo' ? { ...rev, insertion: side } : { ...rev, deletion: side };
        out.push(...projectInline(runItemsOf(el.value as object) ?? [], [...holders, el], ctx, `${address}/${i}`, info, inner));
      } else if (kind === 'ins' || kind === 'moveTo') {
        preserved(ctx, kind, el.value);
        out.push(...projectInline(runItemsOf(el.value as object) ?? [], [...holders, el], ctx, `${address}/${i}`, info, rev));
      } else {
        // A story (decision 2): a `w:del` / `w:moveFrom` is an invisible opaque item so that a
        // rebuilt paragraph still carries it.
        preserved(ctx, kind, el.value);
        const sourceId = ctx.ids.idOf(el);
        ctx.sources.set(sourceId, { holders });
        out.push(ctx.schema.nodes['opaqueInline']!.create({ sourceId, kind: local }));
      }
      return;
    }
    if (tn === SDT_RUN) {
      const sourceId = ctx.ids.idOf(el);
      ctx.sources.set(sourceId, { holders });
      const sdt = el.value as wml.SdtRun;
      const children = projectInline(sdt.sdtContent?.content ?? [], [], ctx, `${address}/${i}`, info, rev);
      const node = ctx.schema.nodes['sdtInline']!.create({ sourceId, sdtPr: sdt.sdtPr ? toPlain(sdt.sdtPr) : null }, children);
      ctx.imported.set(sourceId, node);
      out.push(node);
      return;
    }
    if (tn !== undefined && TRANSPARENT_HOLDERS.has(tn) && tn !== SIMPLE_FIELD) {
      out.push(...projectInline(runItemsOf(el.value as object) ?? [], [...holders, el], ctx, `${address}/${i}`, info, rev));
      return;
    }
    // A move range or a bookmark opens or closes its range for the walk; it stays an opaque atom below.
    if (MOVE_RANGE.has(local)) trackMoveRange(el, ctx);
    if (local === 'commentRangeStart' || local === 'commentRangeEnd') {
      const id = String((el.value as { id?: unknown }).id ?? '');
      if (local === 'commentRangeStart') {
        if (!ctx.openComments.includes(id)) ctx.openComments.push(id);
        info.commentStarts.set(id, el);
        if (rev.insertion || rev.deletion) (info.startsInRevision ??= new Set()).add(id);
        ctx.commentIds.add(id);
      } else {
        const at = ctx.openComments.indexOf(id);
        if (at >= 0) ctx.openComments.splice(at, 1);
        info.commentEnds.set(id, el);
      }
      return;
    }
    if (local === 'proofErr') return;   // Word's spelling marks: regenerated by Word, not worth carrying
    if (local === 'AlternateContent' && isAlternateContent(el)) {
      // ED-009: a paragraph-level kept element is a holder of its selected branch's items, as a
      // hyperlink is of its runs; the export rebuilds it around them (`rebuildAlternateContent`),
      // the other branches as they were. With no branch to select it stays opaque below.
      const branch = selectedBranchOf(el.value as mce.AlternateContent);
      if (branch) {
        out.push(...projectInline(branch, [...holders, el], ctx, `${address}/${i}`, info, rev));
        return;
      }
    }
    // Bookmarks, permissions, simple fields, math, alternate content, unknown: opaque, in place.
    // Inside a tracked change it carries the change as a run's atoms do, so the export keeps it in
    // the change's holder rather than splitting the holder around it (the C1 review, ED-005
    // section 12.3 item 11).
    const sourceId = ctx.ids.idOf(el);
    ctx.sources.set(sourceId, { holders });
    ctx.diagnostics.opaque.push({ kind: local, address: `${address}/${i}`, inline: true, sourceId });
    const revision = ctx.revisions === 'content' && (rev.insertion || rev.deletion) ? rev : null;
    out.push(ctx.schema.nodes['opaqueInline']!.create({ sourceId, kind: local, ...(revision ? { revision } : {}) }));
  });
  return out;
}

/** The insertion and deletion marks of a tracked change the content sits in. */
function revisionMarks(rev: Revision, schema: Schema): Mark[] {
  const out: Mark[] = [];
  if (rev.insertion) out.push(schema.marks['insertion']!.create(rev.insertion));
  if (rev.deletion) out.push(schema.marks['deletion']!.create(rev.deletion));
  return out;
}

/** The `link` and `comment` marks a piece of inline content carries from where it sits. */
function linkAndCommentMarks(holders: Element[], ctx: ProjectionContext): Mark[] {
  const marks: Mark[] = [];
  for (const h of holders) {
    if (typeNameOf(h) === HYPERLINK) {
      const v = h.value as wml.P.Hyperlink;
      marks.push(ctx.schema.marks['link']!.create({ sourceId: ctx.ids.idOf(h), rId: v.id ?? null, anchor: v.anchor ?? null }));
    }
  }
  for (const id of ctx.openComments) marks.push(ctx.schema.marks['comment']!.create({ id }));
  return marks;
}

/** The formatting marks of a run's direct properties: one mark per `RunOptions` field. */
export function formattingMarks(fmt: RunFormatting, rPr: wml.RPr | undefined, schema: Schema): Mark[] {
  const m = schema.marks;
  const marks: Mark[] = [];
  if (fmt.bold) marks.push(m['bold']!.create());
  if (fmt.italic) marks.push(m['italic']!.create());
  if (fmt.doubleStrikeThrough) marks.push(m['strike']!.create({ double: true }));
  else if (fmt.strikeThrough) marks.push(m['strike']!.create({ double: false }));
  if (fmt.underline !== 'None') marks.push(m['underline']!.create({ val: fmt.underline }));
  if (fmt.subscript) marks.push(m['vertAlign']!.create({ val: 'subscript' }));
  else if (fmt.superscript) marks.push(m['vertAlign']!.create({ val: 'superscript' }));
  if (fmt.name !== '' || fmt.size !== 0 || fmt.color !== '' || fmt.highlightColor !== null) {
    marks.push(m['font']!.create({
      name: fmt.name === '' ? null : fmt.name,
      size: fmt.size === 0 ? null : fmt.size,
      color: fmt.color === '' ? null : fmt.color,
      highlight: fmt.highlightColor,
    }));
  }
  if (fmt.style !== '') marks.push(m['runStyle']!.create({ id: fmt.style }));
  if (on(rPr?.rtl)) marks.push(m['rtl']!.create());
  return marks;
}

const EMU_PER_POINT = 12700;

function projectRun(el: Element<wml.R>, holders: Element[], ctx: ProjectionContext, address: string, rev: Revision): PMNode[] {
  const sourceId = ctx.ids.idOf(el);
  ctx.sources.set(sourceId, { holders });
  const run = el.value;
  const schema = ctx.schema;
  const content = ctx.revisions === 'content';
  // A run's formatting changed as a tracked change: a mark with the old `w:rPr` (ED-005 section 5.2).
  const rPrChange = (run.rPr as { rPrChange?: wml.CTRPrChange } | undefined)?.rPrChange;
  const formatChange: Mark[] = [];
  if (rPrChange) {
    noteId(ctx, rPrChange.id);
    if (content) {
      ctx.diagnostics.tracked.shown++;
      formatChange.push(schema.marks['formatChange']!.create(formatChangeAttrs(rPrChange)));
    } else preserved(ctx, 'rPrChange', rPrChange);
  }
  // Text carries the formatting as marks; atoms carry only `runSource` (UPSTREAM-2: y-prosemirror
  // keeps marks on text alone), and export takes their formatting from the source run.
  const baseMarks = [
    ...formattingMarks(readRunOptions(run.rPr), run.rPr, schema),
    schema.marks['runSource']!.create({ sourceId }),
    ...linkAndCommentMarks(holders, ctx),
    ...(content ? revisionMarks(rev, schema) : []),
    ...formatChange,
  ];
  const revision = content && (rev.insertion || rev.deletion) ? rev : null;
  const baseAtom = revision ? { runSource: sourceId, revision } : { runSource: sourceId };
  const out: PMNode[] = [];
  /**
   * One item of the run, or of the selected branch of a kept `mc:AlternateContent` among its items
   * (ED-009): then `alt` names the element, on the text's `runSource` mark and on the atoms, and the
   * export writes the branch's content back into it.
   */
  const visit = (item: Element, address: string, alt: string | undefined): void => {
    const local = item.name.localPart;
    const v = item.value as Record<string, unknown>;
    const marks = alt === undefined ? baseMarks : baseMarks.map((m) => (m.type.name === 'runSource' ? schema.marks['runSource']!.create({ sourceId, alt }) : m));
    const atom = alt === undefined ? baseAtom : { ...baseAtom, alt };
    switch (local) {
      case 't': {
        const text = String(v['value'] ?? '');
        if (text !== '') out.push(schema.text(text, marks));
        return;
      }
      case 'delText': {
        // Deleted text is text in the body (ED-005 section 5.2); elsewhere, or outside a deletion, it stays an opaque item.
        if (!content || !rev.deletion) break;
        const text = String(v['value'] ?? '');
        if (text !== '') out.push(schema.text(text, marks));
        return;
      }
      case 'tab':
        out.push(schema.nodes['tab']!.create(atom));
        return;
      case 'br': {
        const type = v['type'];
        const kind = type === 'page' || type === 'column' ? type : 'line';
        out.push(schema.nodes['hardBreak']!.create({ ...atom, kind, clear: (v['clear'] as string | undefined) ?? null }));
        return;
      }
      case 'cr':
        out.push(schema.nodes['hardBreak']!.create({ ...atom, kind: 'line' }));
        return;
      case 'lastRenderedPageBreak':
        // Word's last pagination as data (ED-001 Appendix D): an invisible atom, so it stays in
        // place through a rebuild and the view can draw a page boundary from it. Not "opaque
        // content" for the report.
        ctx.diagnostics.pageBreakMarkers++;
        out.push(schema.nodes['opaqueInline']!.create({ ...atom, sourceId: ctx.ids.idOf(item), kind: local }));
        return;
      case 'AlternateContent': {
        // A kept `mc:AlternateContent` (core-ts CR-010 phase 4): its selected branch's items are the
        // run's own, each naming the element as `alt` (ED-009; ED-005 section 12.106 item 3 had made
        // the element an image node or an opaque atom). One level: an element inside a branch, and
        // one with no branch to select, stay opaque atoms below.
        const branch = alt === undefined ? selectedBranchOf(item.value as mce.AlternateContent) : undefined;
        if (branch) {
          const id = ctx.ids.idOf(item);
          branch.forEach((b, j) => visit(b, `${address}/${j}`, id));
          return;
        }
        break;
      }
      case 'drawing': {
        // A `w:drawing` with one `wp:inline` is an image node, whatever its graphic; an anchored
        // drawing stays an opaque atom below.
        const inline = inlineOf(item);
        if (inline) {
          const itemId = ctx.ids.idOf(item);
          const cx = inline.extent?.cx;
          const cy = inline.extent?.cy;
          out.push(schema.nodes['image']!.create({
            ...atom,
            sourceId: itemId,
            width: typeof cx === 'number' ? cx / EMU_PER_POINT : null,
            height: typeof cy === 'number' ? cy / EMU_PER_POINT : null,
          }));
          return;
        }
        break;
      }
      case 'footnoteReference':
      case 'endnoteReference': {
        const itemId = ctx.ids.idOf(item);
        out.push(schema.nodes['footnoteRef']!.create({
          ...atom,
          sourceId: itemId,
          footnoteId: (item.value as wml.CTFtnEdnRef).id,
          kind: local === 'footnoteReference' ? 'footnote' : 'endnote',
        }));
        return;
      }
      default:
        break;
    }
    // fldChar, instrText, sym, pict, object, commentReference, delText, ruby, anchored drawings, ...
    // A deleted field's code is its `w:instrText` deleted: the export writes `w:delInstrText` for
    // it exactly while the atom is deleted (ED-005 section 5.3), so a rejected deletion round-trips.
    const kind = local === 'delInstrText' && content && rev.deletion ? 'instrText' : local;
    const itemId = ctx.ids.idOf(item);
    ctx.diagnostics.opaque.push({ kind, address, inline: true, sourceId: itemId });
    out.push(schema.nodes['opaqueInline']!.create({ ...atom, sourceId: itemId, kind }));
  };
  (run.content ?? []).forEach((item, i) => visit(item, `${address}/${i}`, undefined));
  // A run with no items at all (properties only, as Word sometimes leaves behind) is kept as an
  // invisible paragraph-level atom, so a rebuilt paragraph is still element for element the original.
  if (out.length === 0 && (run.content ?? []).length === 0) out.push(schema.nodes['opaqueInline']!.create({ sourceId, kind: 'r', ...(revision ? { revision } : {}) }));
  return out;
}
