// SPDX-License-Identifier: GPL-3.0-or-later
//
// Export: ProseMirror nodes back to the typed tree (ED-002 section 3.6; ED-001 section 3), where
// fidelity lives. For each block: unchanged since import (identity, then `Node.eq`, because a
// remote Yjs update rebuilds nodes) emits the original element; otherwise the element is rebuilt
// from a copy of the original, with only the fields the editor owns written back, so everything
// the editor never showed (`w:rsid`, `w:lang`, a run's `w:kern`, a paragraph's `w:keepNext`)
// survives. A node with no `sourceId` is built with the builders.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import * as wmlEl from '@docx4j/generated-objects-ts/modules/org_docx4j_wml.el';
import * as wmlFactory from '@docx4j/generated-objects-ts/modules/org_docx4j_wml.factory';
import {
  deepCopy, linkParents, readRunOptions, applyRunOptions, revisionKindOf, typeNameOf,
  runOf, t, br, tab, selectedBranchOf, isAlternateContent, type Element, type RunOptions, type RunFormatting,
} from '@docx4j/core-ts';
import type * as mce from '@docx4j/generated-objects-ts/modules/org_docx4j_mce';
import { Node as PMNode, Mark, type Schema } from 'prosemirror-model';
import { TableMap } from 'prosemirror-tables';
import { copyWithout, plainEqual, revive, toPlain } from './json.mjs';
import { FORMATTING_MARKS, type FontAttrs, type ParagraphAttrs, type Revision, type RevisionInfo } from './schema.mjs';
import { paragraphAttrsOf, projectedParagraphAttrs, formattingMarks, cellPropertiesAttr, TWIPS_PER_PX, type ProjectionContext, type SdtGroup, type TableInfo } from './import.mjs';
import { writeRevision, withoutRevision, isHolderOf, revisionMarksOf, revisionHolder, formatChangeAttrs, rPrChangeOf, dateFields, type FormatChangeAttrs } from './revisions.mjs';
import { referenceIdOf } from './comments.mjs';
import { exportedSdtPr } from './sdtpr.mjs';

/** The markers an otherwise empty paragraph may hold beside a comment's reference: nothing a reader sees. */
const EMPTY_PARAGRAPH_MARKERS = new Set(['commentRangeStart', 'commentRangeEnd', 'bookmarkStart', 'bookmarkEnd', 'proofErr', 'lastRenderedPageBreak']);

export interface ExportResult {
  /** The body's new content array; the original elements where nothing changed. */
  content: Element[];
  /** The body's final `w:sectPr`: the original object, or one revived from `doc.attrs.finalSectPr`. */
  sectPr: wml.SectPr | undefined;
  /** False when every block is the original element and the final section is unchanged: the part need not be touched. */
  changed: boolean;
  /**
   * The element each block node the export visited became (a paragraph, a table, a row, a cell, a
   * block control): the original where it was unchanged, the rebuilt one otherwise. A node inside a
   * block emitted by identity was not visited; its element is its `sourceId`'s (ED-005 section 12.27:
   * the agent's addresses).
   */
  nodes: Map<PMNode, Element>;
}

/**
 * The ids an export gives the revisions it writes afresh (ED-005 section 5.7): a new revision
 * holder, a second copy of one, a paragraph mark's or a row's new side, a rewritten
 * `w:pPrChange` or `w:rPrChange`. A save and a tree run pass the engine's annotation counter,
 * seeded over every part (`EditorDocument.revisionIds`); an export only looked at counts on from
 * the highest id its projection saw.
 */
export type RevisionIds = () => number;

/** Ids counted on from the highest annotation id the projection saw, the context left as it is. */
export function localRevisionIds(ctx: ProjectionContext): RevisionIds {
  let n = ctx.maxRevisionId;
  return () => ++n;
}

/** A row- or cell-level group's current properties, from the `controls` map (ED-003 section 4.2); undefined for a group the map does not hold. */
export type ControlLookup = (groupId: string) => { form: 'row' | 'cell'; sdtPr: unknown | null } | undefined;

/** What a story's document was imported from: the body, a `w:footnote`, a `w:hdr`, a `w:ftr` (ED-003 section 6.2). */
export interface BlockContainer {
  content?: Element[];
  /** The body's final section; a story container has none. */
  sectPr?: wml.SectPr;
}

/** Export a `doc` node against the body it was imported from. */
export function exportBody(doc: PMNode, ctx: ProjectionContext, body: wml.Body, controls: ControlLookup = () => undefined, ids: RevisionIds = localRevisionIds(ctx)): ExportResult {
  return exportBlocks(doc, ctx, body as BlockContainer, controls, ids);
}

/**
 * Export a document's blocks against the container they were imported from: the body, or a
 * story's element with the story's own context (ED-003 section 6.2). `doc.attrs.finalSectPr`
 * is compared with the container's `sectPr`; a story has neither, so it costs nothing there.
 */
export function exportBlocks(doc: PMNode, ctx: ProjectionContext, container: BlockContainer, controls: ControlLookup = () => undefined, ids: RevisionIds = localRevisionIds(ctx)): ExportResult {
  const state = new ExportState(ctx, controls, ids);
  state.commentBounds = newCommentBounds(doc, ctx);
  const content: Element[] = [];
  doc.forEach((node) => content.push(state.emitBlock(node)));
  let sectPr = container.sectPr;
  const finalSectPr = (doc.attrs['finalSectPr'] as unknown) ?? null;
  if (!plainEqual(finalSectPr, container.sectPr ? toPlain(container.sectPr) : null)) {
    sectPr = finalSectPr === null ? undefined : (revive(finalSectPr) as wml.SectPr);
    state.changed = true;
  }
  // The original content array is compared element by element: a reordering is a change too.
  const original = container.content ?? [];
  if (!state.changed && (original.length !== content.length || original.some((el, i) => el !== content[i]))) state.changed = true;
  return { content, sectPr, changed: state.changed, nodes: state.emitted };
}

/** One inline node with what export needs to place it: the holders it goes back into, and its run. */
/**
 * Where the markers of the comments made in the editor go (ED-005 section 7.4): the paragraph
 * holding each one's first anchor (a marked text node or its reference atom) and the one holding
 * its last, and the order of the reference atoms, which orders markers at one place. A comment
 * whose markers the file holds is left to `placeCommentMarkers`, which re-emits them where they
 * were, since an unedited paragraph keeps its own by identity; a comment made here is anchored by
 * edits, so every paragraph it touches is rebuilt.
 */
interface NewCommentBounds {
  first: Map<string, PMNode>;
  last: Map<string, PMNode>;
  refOrder: Map<string, number>;
  /**
   * A file comment's markers follow its text (ED-005 section 12.116). While the paragraph holding a marker
   * is written as it was, the marker stays in it; once that paragraph is rebuilt or gone, a start goes
   * before the comment's first anchor and an end after its last, as a comment made here has them. Enter
   * inside a comment makes the first half the new paragraph, a paste of paragraphs puts the rest of the
   * comment in a new one, and Backspace joins the paragraph holding them into the one before. By the
   * paragraph each now goes into, with the element the file had.
   */
  movedStarts: Map<PMNode, [string, Element][]>;
  movedEnds: Map<PMNode, [string, Element][]>;
  /** The comments whose file start or end is moved, which the paragraph that held it leaves out. */
  leftStarts: Set<string>;
  leftEnds: Set<string>;
  /** The paragraphs a marker moves into and the blocks holding them: rebuilt, though their content may be as imported. */
  forced: Set<PMNode>;
}

function newCommentBounds(doc: PMNode, ctx: ProjectionContext): NewCommentBounds {
  const marked = new Set<string>();
  for (const info of ctx.paragraphs.values()) for (const id of [...info.commentStarts.keys(), ...info.commentEnds.keys()]) marked.add(id);
  const out: NewCommentBounds = { first: new Map(), last: new Map(), refOrder: new Map(), movedStarts: new Map(), movedEnds: new Map(), leftStarts: new Set(), leftEnds: new Set(), forced: new Set() };
  // A file comment's first and last anchor, and where each paragraph stands: for its markers' moves.
  const fileFirst = new Map<string, PMNode>();
  const fileLast = new Map<string, PMNode>();
  const at = new Map<PMNode, number>();
  const bySource = new Map<string, PMNode>();
  doc.descendants((para, pos) => {
    if (para.type.name !== 'paragraph') return true;
    at.set(para, pos);
    const sourceId = para.attrs['sourceId'];
    if (typeof sourceId === 'string') bySource.set(sourceId, para);
    para.descendants((node) => {
      const ids: string[] = [];
      for (const m of node.marks) if (m.type.name === 'comment') ids.push(String(m.attrs['id']));
      const ref = referenceIdOf(node, ctx.ids);
      if (ref !== undefined) { ids.push(ref); if (!out.refOrder.has(ref)) out.refOrder.set(ref, out.refOrder.size); }
      for (const id of ids) {
        const first = marked.has(id) ? fileFirst : out.first;
        const last = marked.has(id) ? fileLast : out.last;
        if (!first.has(id)) first.set(id, para);
        last.set(id, para);
      }
      return true;
    });
    return false;
  });
  if (fileFirst.size === 0) return out;
  const move = (moved: Map<PMNode, [string, Element][]>, to: PMNode, id: string, el: Element): void => {
    const list = moved.get(to);
    if (list) list.push([id, el]); else moved.set(to, [[id, el]]);
    const $pos = doc.resolve(at.get(to)!);
    for (let d = 1; d <= $pos.depth; d++) out.forced.add($pos.node(d));
    out.forced.add(to);
  };
  for (const [paragraphId, info] of ctx.paragraphs) {
    if (info.commentStarts.size === 0 && info.commentEnds.size === 0) continue;
    const holder = bySource.get(paragraphId);
    // Written by identity, its markers with it, as the export's `unchanged` writes it.
    const imported = ctx.imported.get(paragraphId);
    if (holder && imported && (imported === holder || imported.eq(holder))) continue;
    for (const [id, el] of info.commentStarts) {
      const to = fileFirst.get(id);
      if (!to || to === holder) continue;
      move(out.movedStarts, to, id, el);
      out.leftStarts.add(id);
    }
    for (const [id, el] of info.commentEnds) {
      const to = fileLast.get(id);
      if (!to || to === holder) continue;
      move(out.movedEnds, to, id, el);
      out.leftEnds.add(id);
    }
  }
  return out;
}

interface Plan {
  node: PMNode | undefined;
  /** A paragraph-level element emitted as is (a comment marker, an opaque item), instead of `node`. */
  element?: Element;
  holders: Element[];
  runSource: string | undefined;
  /** The kept `mc:AlternateContent` inside the run this content came from, by id (ED-009); undefined for the run's own items. */
  alt?: string;
  /** The formatting marks as a comparable key; runs are grouped by (runSource, fmtKey). */
  fmtKey: string;
  marks: readonly Mark[];
  /** Inside a deletion (a `w:del` or `w:moveFrom` among its holders): text goes out as `w:delText`. */
  deleted?: boolean;
  /**
   * In the body, the run's formatting change (the text's `formatChange` mark, an atom's source
   * run's `w:rPrChange`), or null for none; undefined in a story, whose runs keep theirs as they are.
   */
  formatChange?: FormatChangeAttrs | null;
}

/**
 * Whether a run holder that is a tracked change (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`)
 * still holds content with this revision (ED-005 section 5.3, in the body): an insertion holder
 * content inserted by that author at that date, a deletion holder content so deleted. Text typed
 * next to a deletion takes the deleted run's source but not the deletion, and so goes out of the
 * holder; text another author typed into an insertion goes out of that author's. Other holders stay.
 */
function holdsRevision(h: Element, revision: { insertion?: RevisionInfo; deletion?: RevisionInfo }): boolean {
  const kind = revisionKindOf(h);
  if (kind === undefined) return true;
  return kind === 'ins' || kind === 'moveTo' ? isHolderOf(h, 'insertion', revision.insertion) : isHolderOf(h, 'deletion', revision.deletion);
}

/** The sides a piece of text is in, from its `insertion` and `deletion` marks (the earlier deletion, when two overlap). */
function revisionOfMarks(marks: readonly Mark[]): { insertion?: RevisionInfo; deletion?: RevisionInfo } {
  const out: { insertion?: RevisionInfo; deletion?: RevisionInfo } = {};
  const info = (m: Mark): RevisionInfo => ({ author: (m.attrs['author'] as string | null) ?? null, date: (m.attrs['date'] as string | null) ?? null, move: (m.attrs['move'] as string | null) ?? null });
  const { insertion, deletion } = revisionMarksOf(marks);
  if (insertion) out.insertion = info(insertion);
  if (deletion) out.deletion = info(deletion);
  return out;
}

const isDeletionHolder = (h: Element): boolean => { const k = revisionKindOf(h); return k === 'del' || k === 'moveFrom'; };

/** `w:delText` for deleted text, with `xml:space` where the text needs it, as `t` writes `w:t`. */
function delText(text: string): Element {
  const value: wml.DelText = { TYPE_NAME: 'org_docx4j_wml.DelText', value: text };
  if (/^\s|\s$|\s\s/.test(text)) value.space = 'preserve';
  return wmlEl.delText(value) as Element;
}

class ExportState {
  changed = false;
  /** Revision holders already copied once in this export: a second copy (a paragraph split inside a change) gets a fresh id. */
  private emittedHolders = new Set<Element>();
  /** The holders made in the current paragraph for content the file has none for, one per side and markup: adjacent content shares one. */
  private madeHolders = new Map<string, Element>();
  /** The kept `mc:AlternateContent` elements rebuilt once already in this export: a later piece (a split, ED-009 section 3.4) carries its other branches empty. */
  private emittedAlternates = new Set<Element>();
  /** Where the markers of the comments made in the editor go. */
  commentBounds: NewCommentBounds | undefined;
  /** The element each block node became (`ExportResult.nodes`). */
  readonly emitted = new Map<PMNode, Element>();

  constructor(readonly ctx: ProjectionContext, readonly controls: ControlLookup, readonly nextId: RevisionIds) {}

  private unchanged(node: PMNode): Element | undefined {
    const id = node.attrs['sourceId'] as string | null;
    if (!id) return undefined;
    // A file comment's marker moves into it (section 12.116).
    if (this.commentBounds?.forced.has(node)) return undefined;
    const imported = this.ctx.imported.get(id);
    if (!imported) return undefined;
    if (imported !== node && !imported.eq(node)) return undefined;
    return this.ctx.ids.get(id);
  }

  emitBlock(node: PMNode): Element {
    const element = this.emitBlockOnce(node);
    this.emitted.set(node, element);
    return element;
  }

  private emitBlockOnce(node: PMNode): Element {
    const original = this.unchanged(node);
    // A table's rows and cells carry their groups as attributes, but a group's properties live
    // in the `controls` map (ED-003 section 4.2): an unchanged node may still hold a changed group.
    if (original && !(node.type.name === 'table' && this.groupsChanged(node))) return original;
    this.changed = true;
    switch (node.type.name) {
      case 'paragraph': return this.rebuildParagraph(node);
      case 'sdtBlock': return this.rebuildSdtBlock(node);
      case 'table': return this.rebuildTable(node);
      case 'opaqueBlock': {
        const el = this.ctx.ids.get(node.attrs['sourceId'] as string);
        if (!el) throw new Error('opaque block without a source element');
        return el;
      }
      default:
        throw new Error(`unexpected block node ${node.type.name}`);
    }
  }

  private rebuildSdtBlock(node: PMNode): Element {
    const id = node.attrs['sourceId'] as string | null;
    const orig = id ? (this.ctx.ids.get(id) as Element<wml.SdtBlock> | undefined) : undefined;
    const children: Element[] = [];
    node.forEach((child) => children.push(this.emitBlock(child)));
    const value: wml.SdtBlock = orig ? copyWithout(orig.value, ['sdtContent']) : { TYPE_NAME: 'org_docx4j_wml.SdtBlock' };
    const sdtPr = node.attrs['sdtPr'] as unknown;
    if (!orig || !plainEqual(sdtPr, orig.value.sdtPr ? toPlain(orig.value.sdtPr) : null)) {
      if (sdtPr === null) delete value.sdtPr; else value.sdtPr = exportedSdtPr(sdtPr);
    }
    value.sdtContent = { TYPE_NAME: 'org_docx4j_wml.SdtContentBlock', content: children as wml.SdtContentBlock['content'] };
    const el: Element<wml.SdtBlock> = { name: orig?.name ?? { namespaceURI: W_NS, localPart: 'sdt' }, value };
    linkParents(el.value, undefined);
    return el;
  }

  // ---------------------------------------------------------------- tables (ED-003 section 3.4)

  /**
   * A changed table: a copy of the original `w:tbl` (or a new one) whose rows are the row nodes,
   * unchanged rows by identity, changed rows rebuilt from their copies, with the cells the map
   * says a row holds: its own cell nodes, and a continuation `w:tc` for every column a merge
   * from above covers. Rows go back into their row-level controls from the `sdt` stacks.
   */
  private rebuildTable(node: PMNode): Element<wml.Tbl> {
    const id = node.attrs['sourceId'] as string | null;
    const orig = id ? (this.ctx.ids.get(id) as Element<wml.Tbl> | undefined) : undefined;
    const info: TableInfo = (id && this.ctx.tables.get(id)) || { rowExtras: [], rows: new Map(), continuations: new Map() };
    const map = TableMap.get(node);
    const value: wml.Tbl = orig ? copyWithout(orig.value, ['content']) : { TYPE_NAME: 'org_docx4j_wml.Tbl', tblPr: { TYPE_NAME: 'org_docx4j_wml.TblPr' }, tblGrid: { TYPE_NAME: 'org_docx4j_wml.TblGrid' } };
    const tblPr = node.attrs['tblPr'] as unknown;
    if (!orig || !plainEqual(tblPr, orig.value.tblPr ? toPlain(orig.value.tblPr) : null)) {
      value.tblPr = tblPr === null ? { TYPE_NAME: 'org_docx4j_wml.TblPr' } : (revive(tblPr) as wml.TblPr);
    }
    const grid = node.attrs['tblGrid'] as number[] | null;
    const widths = this.columnWidths(node, map, grid);
    // Against the element's own grid, not the attribute: a command that sets the attribute (a column added, a layout
    // switch, ED-005 section 12.78) must reach the file.
    const origGrid = orig ? (orig.value.tblGrid?.gridCol ?? []).map((g) => g.w ?? 0) : null;
    if (!orig || !plainEqual(widths, origGrid)) {
      value.tblGrid = { TYPE_NAME: 'org_docx4j_wml.TblGrid', gridCol: widths.map((w) => ({ TYPE_NAME: 'org_docx4j_wml.TblGridCol', w })) };
    }
    // The rows, each with its stack, then grouped into their controls. A rebuilt cell's width is written in the
    // form its table states (check 39, ED-005 section 12.79): the frame carries the grid and the table's width type.
    const tblWType = ((value.tblPr?.tblW as { type?: string } | undefined)?.type) ?? 'dxa';
    const outerFrame = this.tableFrame;
    this.tableFrame = { grid: widths, widthType: tblWType };
    const rowElements: { element: Element; stack: SdtGroup[] }[] = [];
    try {
      for (let r = 0; r < map.height; r++) {
        const rowNode = node.child(r);
        rowElements.push({ element: this.emitRow(node, rowNode, r, map, info), stack: (rowNode.attrs['sdt'] as SdtGroup[] | null) ?? [] });
      }
    } finally {
      this.tableFrame = outerFrame;
    }
    const content = this.groupRows(rowElements, info.rowExtras);
    value.content = content as wml.Tbl['content'];
    const el: Element<wml.Tbl> = { name: orig?.name ?? { namespaceURI: W_NS, localPart: 'tbl' }, value };
    linkParents(el.value, undefined);
    return el;
  }

  /** The grid's column widths in twips: the imported grid while every cell keeps its width, else measured from the cells. */
  private columnWidths(table: PMNode, map: TableMap, grid: number[] | null): number[] {
    if (grid && grid.length === map.width) {
      let resized = false;
      table.descendants((n) => {
        if (n.type.name !== 'table_cell' || resized) return !resized;
        const imported = this.ctx.imported.get(n.attrs['sourceId'] as string);
        if (!imported || !plainEqual(imported.attrs['colwidth'], n.attrs['colwidth'])) resized = true;
        return false;
      });
      if (!resized) return grid;
    }
    const widths: number[] = new Array<number>(map.width).fill(0);
    const seen = new Set<number>();
    for (let r = 0; r < map.height; r++) {
      for (let c = 0; c < map.width; c++) {
        const pos = map.map[r * map.width + c]!;
        if (seen.has(pos)) continue;
        seen.add(pos);
        const cell = table.nodeAt(pos)!;
        const colwidth = cell.attrs['colwidth'] as number[] | null;
        const span = cell.attrs['colspan'] as number;
        for (let k = 0; k < span; k++) {
          const w = colwidth?.[k];
          if (typeof w === 'number' && w > 0 && widths[c + k] === 0) widths[c + k] = Math.round(w * TWIPS_PER_PX);
        }
      }
    }
    const fallback = grid?.length ? Math.round(grid.reduce((a, b) => a + b, 0) / Math.max(1, map.width)) : Math.round(9026 / Math.max(1, map.width));
    return widths.map((w, i) => (w > 0 ? w : grid?.[i] ?? fallback));
  }

  /** A row: the original element when unchanged and covered as it was imported; else a copy with the cells the map says. */
  private emitRow(table: PMNode, rowNode: PMNode, r: number, map: TableMap, info: TableInfo): Element {
    const id = rowNode.attrs['sourceId'] as string | null;
    const orig = id ? (this.ctx.ids.get(id) as Element<wml.Tr> | undefined) : undefined;
    const rowInfo = id ? info.rows.get(id) : undefined;
    // Which columns a merge from above covers in this row, and by which restart cell at which offset.
    const covered: string[] = [];
    // A merge whose shading the editor changed: its continuation here follows it (continuationCell), so the row is rebuilt.
    let shadingFollows = false;
    for (let c = 0; c < map.width; c++) {
      const pos = map.map[r * map.width + c]!;
      const rect = map.findCell(pos);
      if (rect.top < r) {
        const restart = table.nodeAt(pos)!;
        covered[c] = `${String(restart.attrs['sourceId'] ?? `new@${pos}`)}:${r - rect.top}`;
        if (!shadingFollows && this.changedShading(restart) !== undefined) shadingFollows = true;
      }
    }
    if (orig && rowInfo && !shadingFollows && this.unchanged(rowNode) && plainEqual(covered, rowInfo.covered) && !this.groupsChanged(rowNode)) return orig;
    this.changed = true;
    const value: wml.Tr = orig ? copyWithout(orig.value, ['content']) : { TYPE_NAME: 'org_docx4j_wml.Tr' };
    const trPr = rowNode.attrs['trPr'] as unknown;
    // In the body the attribute is `w:trPr` less the row's tracked change, which is `revision` (ED-005 section 5.2).
    const imported = id ? this.ctx.imported.get(id) : undefined;
    const content = this.ctx.revisions === 'content';
    const baseTrPr = content && imported ? imported.attrs['trPr'] : orig?.value.trPr ? toPlain(orig.value.trPr) : null;
    const trPrChanged = !orig || !plainEqual(trPr, baseTrPr);
    if (trPrChanged) {
      if (trPr === null) delete value.trPr; else value.trPr = revive(trPr) as wml.TrPr;
    }
    if (content) {
      const revision = (rowNode.attrs['revision'] as Revision | null) ?? null;
      const baseRevision = (imported?.attrs['revision'] as Revision | null | undefined) ?? null;
      if (trPrChanged || !plainEqual(revision, baseRevision)) {
        if (revision) writeRevision((value.trPr ??= { TYPE_NAME: 'org_docx4j_wml.TrPr' }), revision, baseRevision, orig?.value.trPr, this.nextId);
        else if (value.trPr) { delete value.trPr.ins; delete value.trPr.del; if (withoutRevision(value.trPr) === null && !value.trPr.trPrChange) delete value.trPr; }
      }
    }
    // Cells in grid order: this row's own, and a continuation where a merge covers.
    const cells: { element: Element; stack: SdtGroup[] }[] = [];
    let ownIndex = 0;
    for (let c = 0; c < map.width;) {
      const pos = map.map[r * map.width + c]!;
      const rect = map.findCell(pos);
      const cell = table.nodeAt(pos)!;
      const span = cell.attrs['colspan'] as number;
      if (rect.top === r) {
        cells.push({ element: this.emitCell(cell, c), stack: (cell.attrs['sdt'] as SdtGroup[] | null) ?? [] });
        ownIndex++;
      } else {
        cells.push({ element: this.continuationCell(cell, r - rect.top, c, info), stack: [] });
      }
      c += span;
    }
    // Cell-level extras (a bookmark between cells) go back before the cell they preceded.
    const extras = rowInfo?.cellExtras ?? [];
    const grouped = this.groupCells(cells, extras);
    value.content = grouped as wml.Tr['content'];
    const el: Element<wml.Tr> = { name: orig?.name ?? { namespaceURI: W_NS, localPart: 'tr' }, value };
    void ownIndex;
    return el;
  }

  /** A cell: the original when unchanged; else a copy with span, merge and width written from the attributes and the content rebuilt. */
  /** The table being rebuilt: its grid in twips and the type of its `w:tblW` (the form a new cell's width takes). */
  private tableFrame: { grid: number[]; widthType: string } | undefined;

  private emitCell(cell: PMNode, column: number): Element {
    const element = this.emitCellOnce(cell, column);
    this.emitted.set(cell, element);
    return element;
  }

  /**
   * A rebuilt cell's `w:tcW`, in the form the cell had, else the form its table's width has (Word 365, check 39,
   * ED-005 section 12.79): a percentage cell its share of the grid in fiftieths of a percent (the cumulative shares
   * floored, so three equal columns are 1666, 1667, 1667 as Word writes them); else twips.
   */
  private cellWidth(colwidth: number[], column: number, was: string | undefined): wml.TcPr['tcW'] {
    // A cell with a width of its own is never automatic (an automatic cell has none: `emitCellOnce`); the percentage
    // form is the cell's own or its table's (an automatic table width holds twips cells, Word's default table: check 42).
    const type = was === 'pct' || (was === undefined && this.tableFrame?.widthType === 'pct') ? 'pct' : 'dxa';
    const twips = Math.round(colwidth.reduce((a, b) => a + b, 0) * TWIPS_PER_PX);
    const grid = this.tableFrame?.grid;
    if (type === 'pct' && grid && grid.length) {
      const total = grid.reduce((a, b) => a + b, 0);
      if (total > 0) {
        const start = grid.slice(0, column).reduce((a, b) => a + b, 0);
        const end = Math.min(total, start + grid.slice(column, column + colwidth.length).reduce((a, b) => a + b, 0));
        return { w: Math.floor((end / total) * 5000) - Math.floor((start / total) * 5000), type: 'pct' };
      }
    }
    return { w: twips, type: 'dxa' };
  }

  private emitCellOnce(cell: PMNode, column: number): Element {
    const original = this.unchanged(cell);
    if (original) return original;
    this.changed = true;
    const id = cell.attrs['sourceId'] as string | null;
    const orig = id ? (this.ctx.ids.get(id) as Element<wml.Tc> | undefined) : undefined;
    const value: wml.Tc = orig ? copyWithout(orig.value, ['content']) : { TYPE_NAME: 'org_docx4j_wml.Tc' };
    // The attribute holds the properties less the three the node carries (import's
    // `cellPropertiesAttr`); an unchanged attribute keeps the original's copy, those three included.
    const tcPrAttr = cell.attrs['tcPr'] as unknown;
    const attrChanged = !orig || !plainEqual(tcPrAttr, cellPropertiesAttr(orig.value.tcPr));
    if (attrChanged) value.tcPr = tcPrAttr === null ? { TYPE_NAME: 'org_docx4j_wml.TcPr' } : (revive(tcPrAttr) as wml.TcPr);
    else if (orig.value.tcPr) value.tcPr = deepCopy(orig.value.tcPr);
    const colspan = cell.attrs['colspan'] as number;
    const rowspan = cell.attrs['rowspan'] as number;
    const colwidth = cell.attrs['colwidth'] as number[] | null;
    const imported = id ? this.ctx.imported.get(id) : undefined;
    const tcPr = (value.tcPr ??= { TYPE_NAME: 'org_docx4j_wml.TcPr' });
    if (colspan > 1) tcPr.gridSpan = { val: colspan }; else delete tcPr.gridSpan;
    if (rowspan > 1) tcPr.vMerge = { val: 'restart' }; else delete tcPr.vMerge;
    if (colwidth && (attrChanged || !imported || !plainEqual(imported.attrs['colwidth'], colwidth))) {
      tcPr.tcW = this.cellWidth(colwidth, column, (orig?.value.tcPr?.tcW as { type?: string } | undefined)?.type);
    } else if (!colwidth && (orig ? imported?.attrs['colwidth'] != null : this.tableFrame?.widthType === 'auto')) {
      // A cell with no width of its own, new in an automatic table or one whose width was taken away (AutoFit
      // Contents): automatic, as Word's AutoFit to contents writes its cells (check 42, ED-005 section 12.83).
      tcPr.tcW = { w: 0, type: 'auto' };
    }
    if (Object.keys(tcPr).every((k) => k === 'TYPE_NAME' || k === 'PARENT')) delete value.tcPr;
    const blocks: Element[] = [];
    cell.forEach((child) => blocks.push(this.emitBlock(child)));
    value.content = blocks as wml.Tc['content'];
    return { name: orig?.name ?? { namespaceURI: W_NS, localPart: 'tc' }, value };
  }

  /** The `w:tc` a merge from above occupies in this row: the imported one when there was one, else built from the restart cell. */
  private continuationCell(restart: PMNode, offset: number, column: number, info: TableInfo): Element {
    const restartId = restart.attrs['sourceId'] as string | null;
    const existing = restartId ? info.continuations.get(`${restartId}:${offset}`) as Element<wml.Tc> | undefined : undefined;
    const shading = this.changedShading(restart);
    if (existing) {
      // The restart cell's width changed (a column resized, a fit): the continuation's follows it, as Word writes every
      // cell of a column (check 39); an unchanged one is re-emitted as it was (ED-007 section 8.1 item 11).
      let tcPr: wml.TcPr | undefined;
      const edited = (): wml.TcPr => (tcPr ??= existing.value.tcPr ? deepCopy(existing.value.tcPr) : { TYPE_NAME: 'org_docx4j_wml.TcPr' });
      const colwidth = restart.attrs['colwidth'] as number[] | null;
      const imported = restartId ? this.ctx.imported.get(restartId) : undefined;
      if (colwidth && imported && !plainEqual(imported.attrs['colwidth'], colwidth)) {
        const have = existing.value.tcPr?.tcW;
        const want = this.cellWidth(colwidth, column, (have as { type?: string } | undefined)?.type);
        if (!have || have.w !== want?.w || have.type !== want?.type) edited().tcW = want;
      }
      // Its shading follows the restart cell's too, when the editor changed that (row.shading): one cell to the person, so
      // every `w:tc` of the merge says the same, whichever of them a reader draws the merged cell from (ED-005 section 12.128).
      if (shading !== undefined && !plainEqual(toPlain(existing.value.tcPr?.shd ?? null), shading === null ? null : toPlain(shading))) {
        if (shading === null) delete edited().shd; else edited().shd = shading;
      }
      if (!tcPr) return existing;
      this.changed = true;
      return { name: existing.name, value: { ...copyWithout(existing.value, ['tcPr']), tcPr } };
    }
    this.changed = true;
    const source = restartId ? (this.ctx.ids.get(restartId) as Element<wml.Tc> | undefined) : undefined;
    const tcPr: wml.TcPr = source?.value.tcPr ? copyWithout(deepCopy(source.value.tcPr), ['vMerge', 'gridSpan']) : { TYPE_NAME: 'org_docx4j_wml.TcPr' };
    if (shading === null) delete tcPr.shd; else if (shading !== undefined) tcPr.shd = shading;
    const colspan = restart.attrs['colspan'] as number;
    if (colspan > 1) tcPr.gridSpan = { val: colspan };
    tcPr.vMerge = {};
    const value: wml.Tc = { TYPE_NAME: 'org_docx4j_wml.Tc', tcPr, content: [{ name: { namespaceURI: W_NS, localPart: 'p' }, value: { TYPE_NAME: 'org_docx4j_wml.P' } }] as wml.Tc['content'] };
    return { name: { namespaceURI: W_NS, localPart: 'tc' }, value };
  }

  /**
   * A restart cell's own shading when the editor changed it since the import, for its continuations to follow: the
   * `w:shd` (a fresh copy), null when it was taken away, undefined when unchanged (or a new cell with none).
   */
  private changedShading(restart: PMNode): wml.CTShd | null | undefined {
    const restartId = restart.attrs['sourceId'] as string | null;
    const imported = restartId ? this.ctx.imported.get(restartId) : undefined;
    const shdOf = (node: PMNode): unknown => (node.attrs['tcPr'] as { shd?: unknown } | null)?.shd ?? null;
    const now = shdOf(restart);
    if (imported ? plainEqual(now, shdOf(imported)) : now === null) return undefined;
    return now === null ? null : (revive(now) as wml.CTShd);
  }

  /** Whether a group on the rows or cells of a table, or on a row and its cells, has properties in the map that differ from its source element's. */
  private groupsChanged(node: PMNode): boolean {
    let changed = false;
    const check = (stack: SdtGroup[] | null): void => {
      for (const g of stack ?? []) if (!changed && this.groupChanged(g.groupId)) changed = true;
    };
    const row = (r: PMNode): void => {
      check(r.attrs['sdt'] as SdtGroup[] | null);
      r.forEach((cell) => check(cell.attrs['sdt'] as SdtGroup[] | null));
    };
    if (node.type.name === 'table_row') row(node); else node.forEach(row);
    return changed;
  }

  private readonly groupState = new Map<string, boolean>();

  /** A group made in the editor has no source element and counts as changed; an imported one is changed when its entry's properties differ from the element's. */
  private groupChanged(groupId: string): boolean {
    let v = this.groupState.get(groupId);
    if (v === undefined) {
      const orig = this.ctx.ids.get(groupId);
      const entry = this.controls(groupId);
      const origPr = (orig?.value as { sdtPr?: unknown } | undefined)?.sdtPr;
      v = !orig || (entry !== undefined && !plainEqual(entry.sdtPr, origPr ? toPlain(origPr) : null));
      this.groupState.set(groupId, v);
    }
    return v;
  }

  /** Rows back into their row-level controls: a run of rows sharing a group becomes one holder around them, nested inside out. */
  private groupRows(rows: { element: Element; stack: SdtGroup[] }[], extras: { beforeRow: number; element: Element }[]): Element[] {
    return this.groupItems(rows, extras.map((e) => ({ before: e.beforeRow, element: e.element })), 'sdtContent');
  }

  private groupCells(cells: { element: Element; stack: SdtGroup[] }[], extras: { beforeCell: number; element: Element }[]): Element[] {
    return this.groupItems(cells, extras.map((e) => ({ before: e.beforeCell, element: e.element })), 'sdtContent');
  }

  private groupItems(items: { element: Element; stack: SdtGroup[] }[], extras: { before: number; element: Element }[], _contentKey: string): Element[] {
    interface Open { groupId: string; items: Element[] }
    const top: Element[] = [];
    const open: Open[] = [];
    const target = (): Element[] => (open.length > 0 ? open[open.length - 1]!.items : top);
    const close = (): void => {
      const g = open.pop()!;
      target().push(this.wrapGroup(g.groupId, g.items));
    };
    items.forEach((item, i) => {
      for (const e of extras) if (e.before === i) target().push(e.element);
      const stack = item.stack.map((s) => s.groupId);
      let common = 0;
      while (common < open.length && common < stack.length && open[common]!.groupId === stack[common]) common++;
      while (open.length > common) close();
      for (let d = common; d < stack.length; d++) open.push({ groupId: stack[d]!, items: [] });
      target().push(item.element);
    });
    for (const e of extras) if (e.before >= items.length) target().push(e.element);
    while (open.length > 0) close();
    return top;
  }

  /**
   * A row- or cell-level holder around its items: the original element when its content is the
   * same elements and its properties are what the `controls` map holds, else a copy; a group
   * created in the editor (no source element) is built from its map entry (ED-003 section 4.2).
   */
  private wrapGroup(groupId: string, items: Element[]): Element {
    const orig = this.ctx.ids.get(groupId);
    const entry = this.controls(groupId);
    if (!orig) {
      if (!entry) throw new Error(`content control group ${groupId} has neither a source element nor a controls entry`);
      this.changed = true;
      const [sdtType, contentType] = entry.form === 'cell' ? ['org_docx4j_wml.CTSdtCell', 'org_docx4j_wml.CTSdtContentCell'] : ['org_docx4j_wml.CTSdtRow', 'org_docx4j_wml.CTSdtContentRow'];
      const value: Record<string, unknown> = { TYPE_NAME: sdtType };
      if (entry.sdtPr !== null) value['sdtPr'] = exportedSdtPr(entry.sdtPr);
      value['sdtContent'] = { TYPE_NAME: contentType, content: items };
      const el: Element = { name: { namespaceURI: W_NS, localPart: 'sdt' }, value };
      linkParents(el.value, undefined);
      return el;
    }
    const value = orig.value as { sdtPr?: unknown; sdtContent?: { content?: Element[] }; content?: Element[] };
    const current = value.sdtContent ? value.sdtContent.content ?? [] : value.content ?? [];
    const sameContent = current.length === items.length && current.every((el, i) => el === items[i]);
    // A custom XML wrapper has no entry; a control's entry is compared with the original's properties.
    const sameProperties = !entry || plainEqual(entry.sdtPr, value.sdtPr ? toPlain(value.sdtPr) : null);
    if (sameContent && sameProperties) return orig;
    this.changed = true;
    if (value.sdtContent) {
      const copy = copyWithout(value as Record<string, unknown>, ['sdtContent', ...(sameProperties ? [] : ['sdtPr'])]);
      if (!sameProperties && entry && entry.sdtPr !== null) copy['sdtPr'] = exportedSdtPr(entry.sdtPr);
      copy['sdtContent'] = { TYPE_NAME: (value.sdtContent as { TYPE_NAME?: string }).TYPE_NAME, content: items };
      const el: Element = { name: orig.name, value: copy };
      if (!sameProperties) linkParents(el.value, undefined);
      return el;
    }
    const copy = copyWithout(value as Record<string, unknown>, ['content']);
    copy['content'] = items;
    return { name: orig.name, value: copy };
  }

  private rebuildParagraph(node: PMNode): Element<wml.P> {
    const id = node.attrs['sourceId'] as string | null;
    const orig = id ? (this.ctx.ids.get(id) as Element<wml.P> | undefined) : undefined;
    const value: wml.P = orig ? copyWithout(orig.value, ['content']) : { TYPE_NAME: 'org_docx4j_wml.P' };
    const attrs = node.attrs as unknown as ParagraphAttrs;
    const copied = orig ? undefined : copiedPropertiesOf(attrs.copiedFrom, this.ctx);
    const base = copied ? projectedParagraphAttrs(copied, this.ctx.revisions) : this.paragraphBase(id, orig);
    let pPr = applyParagraphAttrs(orig?.value.pPr ? deepCopy(orig.value.pPr) : copied, attrs, base);
    if (this.ctx.revisions === 'content') pPr = this.writeParagraphRevisions(pPr, attrs, base, orig?.value.pPr);
    if (pPr) value.pPr = pPr; else delete value.pPr;
    this.madeHolders = new Map();
    const plans = this.plansOf(node);
    this.placeCommentMarkers(plans, id, node);
    this.placeNewCommentMarkers(plans, node);
    value.content = this.rebuildInline(plans, 0) as wml.P['content'];
    const el: Element<wml.P> = { name: orig?.name ?? { namespaceURI: W_NS, localPart: 'p' }, value };
    linkParents(el.value, undefined);
    return el;
  }

  /**
   * What a paragraph's attributes were at import, for the comparison that decides what to write:
   * the imported node's where there is one (in the body its `markRPr` less the mark's revision,
   * its `markRevision` with the move it belongs to), else read from the original's properties.
   */
  private paragraphBase(id: string | null, orig: Element<wml.P> | undefined): Omit<ParagraphAttrs, 'sourceId'> {
    return paragraphBaseOf(id, orig, this.ctx);
  }

  /**
   * A paragraph mark's tracked change and the properties' change written back (ED-005 section
   * 5.3): into the mark's `w:rPr` when `markRevision` or `markRPr` differs from the import (an
   * unchanged side keeps its element and id), and `w:pPrChange` from the attribute when it differs.
   */
  private writeParagraphRevisions(pPr: wml.PPr | undefined, attrs: ParagraphAttrs, base: Omit<ParagraphAttrs, 'sourceId'>, origPPr: wml.PPr | undefined): wml.PPr | undefined {
    const markRPrChanged = !plainEqual(attrs.markRPr, base.markRPr);
    if (markRPrChanged || !plainEqual(attrs.markRevision, base.markRevision)) {
      if (attrs.markRevision) {
        const out = (pPr ??= { TYPE_NAME: 'org_docx4j_wml.PPr' });
        const rPr = (out.rPr ??= { TYPE_NAME: 'org_docx4j_wml.ParaRPr' });
        writeRevision(rPr, attrs.markRevision, base.markRevision, origPPr?.rPr, this.nextId);
      } else if (pPr?.rPr) {
        for (const k of ['ins', 'del', 'moveFrom', 'moveTo'] as const) delete pPr.rPr[k];
        if (Object.keys(pPr.rPr).every((k) => k === 'TYPE_NAME' || k === 'PARENT')) delete pPr.rPr;
      }
    }
    if (!plainEqual(attrs.pPrChange, base.pPrChange)) {
      if (attrs.pPrChange === null) delete pPr?.pPrChange;
      else {
        const out = (pPr ??= { TYPE_NAME: 'org_docx4j_wml.PPr' });
        const change: wml.CTPPrChange = { TYPE_NAME: 'org_docx4j_wml.CTPPrChange', id: this.nextId(), author: attrs.pPrChange.author ?? '', pPr: (attrs.pPrChange.before === null ? { TYPE_NAME: 'org_docx4j_wml.PPrBase' } : revive(attrs.pPrChange.before)) as wml.PPrBase };
        if (attrs.pPrChange.date !== null) Object.assign(change, dateFields(attrs.pPrChange.date));
        out.pPrChange = change;
      }
    }
    if (pPr && Object.keys(pPr).every((k) => k === 'TYPE_NAME' || k === 'PARENT')) return undefined;
    return pPr;
  }

  private plansOf(container: PMNode): Plan[] {
    const plans: Plan[] = [];
    this.newLink = undefined;
    container.forEach((n) => plans.push(this.planOf(n)));
    return plans;
  }

  /**
   * The `w:hyperlink` holder being built for a link made in the editor (a `link` mark with no
   * `sourceId`, ED-003 section 5.5): one holder per run of adjacent inline nodes carrying an
   * equal mark, as Word's dialog makes one. `r:id` and `w:anchor` from the mark; `w:history`
   * as Word writes it.
   */
  private newLink: { mark: Mark; el: Element } | undefined;

  private newLinkHolder(mark: Mark): Element {
    if (this.newLink && this.newLink.mark.eq(mark)) return this.newLink.el;
    const value: wml.P.Hyperlink = { TYPE_NAME: HYPERLINK, history: true, content: [] };
    const rId = mark.attrs['rId'] as string | null;
    const anchor = mark.attrs['anchor'] as string | null;
    if (rId) value.id = rId;
    if (anchor) value.anchor = anchor;
    const el: Element = { name: { namespaceURI: W_NS, localPart: 'hyperlink' }, value };
    this.newLink = { mark, el };
    return el;
  }

  private planOf(n: PMNode): Plan {
    if (ATOMS.has(n.type.name)) {
      // UPSTREAM-2: an atom's run and formatting are attributes, not marks.
      const runSource = (n.attrs['runSource'] as string | null) ?? undefined;
      const alt = (n.attrs['alt'] as string | null) ?? undefined;
      const fmt = n.attrs['fmt'] as string | null;
      let marks: readonly Mark[] = [];
      if (typeof fmt === 'string') marks = marksFromKey(fmt, this.ctx.schema);
      else if (runSource) {
        const src = this.ctx.ids.get(runSource) as Element<wml.R> | undefined;
        if (src) marks = formattingMarks(readRunOptions(src.value.rPr), src.value.rPr, this.ctx.schema);
      }
      let holders: Element[] = [];
      if (runSource) holders = this.ctx.sources.get(runSource)?.holders ?? [];
      else if (n.type.name === 'opaqueInline') holders = this.ctx.sources.get(n.attrs['sourceId'] as string)?.holders ?? [];
      let formatChange: FormatChangeAttrs | null | undefined;
      if (this.ctx.revisions === 'content') {
        const revision = (n.attrs['revision'] as Revision | null) ?? {};
        holders = this.withRevisionHolders(holders.filter((h) => holdsRevision(h, revision)), revision);
        // An atom has no marks (UPSTREAM-2): its run's formatting change is its source run's.
        const change = runSource ? (this.ctx.ids.get(runSource) as Element<wml.R> | undefined)?.value.rPr?.rPrChange : undefined;
        formatChange = change ? formatChangeAttrs(change) : null;
      }
      return { node: n, holders, runSource, alt, fmtKey: formattingKey(marks), marks, deleted: holders.some(isDeletionHolder), formatChange };
    }
    if (n.type.name === 'sdtInline') {
      return { node: n, holders: this.ctx.sources.get(n.attrs['sourceId'] as string)?.holders ?? [], runSource: undefined, fmtKey: '', marks: [] };
    }
    // Text: the run's holders, with hyperlinks following the `link` marks strictly, so that what
    // the file says matches what the screen shows. Text typed at the end of a link inherits the
    // run (and so its unseen properties) but not the non-inclusive link mark; it goes after the
    // hyperlink, splitting it when the link continues. Link editing proper is E2.
    const runSourceMark = n.marks.find((m) => m.type.name === 'runSource');
    const runSource = (runSourceMark?.attrs['sourceId'] as string | null | undefined) ?? undefined;
    const alt = (runSourceMark?.attrs['alt'] as string | null | undefined) ?? undefined;
    const links = n.marks.filter((m) => m.type.name === 'link');
    const linked = new Set(links.map((m) => m.attrs['sourceId'] as string | null).filter((id): id is string => id !== null));
    let holders = (runSource ? this.ctx.sources.get(runSource)?.holders ?? [] : []).filter((h) => {
      if (typeNameOf(h) !== HYPERLINK) return true;
      const id = this.ctx.ids.existingIdOf(h);
      return id !== undefined && linked.has(id);
    });
    for (const id of linked) {
      const h = this.ctx.ids.get(id);
      if (h && !holders.includes(h)) holders = [...holders, h];
    }
    // A link made in the editor: a holder of its own, shared with the adjacent nodes of the same link.
    const made = links.find((m) => m.attrs['sourceId'] === null);
    if (made) { holders = [...holders, this.newLinkHolder(made)]; this.changed = true; } else this.newLink = undefined;
    let formatChange: FormatChangeAttrs | null | undefined;
    if (this.ctx.revisions === 'content') {
      const revision = revisionOfMarks(n.marks);
      holders = this.withRevisionHolders(holders.filter((h) => holdsRevision(h, revision)), revision);
      const change = n.marks.find((m) => m.type.name === 'formatChange');
      formatChange = change ? { author: (change.attrs['author'] as string | null) ?? null, date: (change.attrs['date'] as string | null) ?? null, before: change.attrs['before'] as string } : null;
    }
    return { node: n, holders, runSource, alt, fmtKey: formattingKey(n.marks), marks: n.marks, deleted: holders.some(isDeletionHolder), formatChange };
  }

  /**
   * The holders content goes back into, with one made for each side it is in that none of them
   * is (ED-005 section 5.3): content the file has no change for, as the tracker's, a peer's or a
   * tree run's is. An insertion goes around a deletion (another author's insertion deleted is a
   * `w:ins` holding a `w:del`), and both go inside every other holder, since a hyperlink, a
   * simple field or a control holds a change and never the other way round.
   */
  private withRevisionHolders(holders: Element[], revision: { insertion?: RevisionInfo; deletion?: RevisionInfo }): Element[] {
    let out = holders;
    if (revision.insertion && !out.some((h) => isHolderOf(h, 'insertion', revision.insertion))) {
      const made = this.madeHolder('insertion', revision.insertion);
      const at = out.findIndex(isDeletionHolder);
      out = at < 0 ? [...out, made] : [...out.slice(0, at), made, ...out.slice(at)];
    }
    if (revision.deletion && !out.some((h) => isHolderOf(h, 'deletion', revision.deletion))) out = [...out, this.madeHolder('deletion', revision.deletion)];
    return out;
  }

  /** The holder made in this paragraph for one side and its markup, with a fresh id when first made. */
  private madeHolder(side: 'insertion' | 'deletion', info: RevisionInfo): Element {
    const key = JSON.stringify([side, info.author, info.date, info.move]);
    let h = this.madeHolders.get(key);
    if (!h) { h = revisionHolder(side, info, this.nextId()); this.madeHolders.set(key, h); }
    return h;
  }

  /**
   * Re-emit the file's comment range markers this paragraph holds around the content that still carries
   * the mark: its own, less those moved out to follow their text, and those moved in (section 12.116).
   */
  private placeCommentMarkers(plans: Plan[], paragraphId: string | null, node: PMNode): void {
    const info = paragraphId ? this.ctx.paragraphs.get(paragraphId) : undefined;
    const bounds = this.commentBounds;
    const commentEnds = [...(info?.commentEnds ?? [])].filter(([id]) => !bounds?.leftEnds.has(id)).concat(bounds?.movedEnds.get(node) ?? []);
    const commentStarts = [...(info?.commentStarts ?? [])].filter(([id]) => !bounds?.leftStarts.has(id)).concat(bounds?.movedStarts.get(node) ?? []);
    if (commentStarts.length === 0 && commentEnds.length === 0) return;
    const carries = (p: Plan, id: string): boolean => p.node !== undefined && p.node.marks.some((m) => m.type.name === 'comment' && m.attrs['id'] === id);
    const isReference = (p: Plan, id: string): boolean => {
      if (p.node?.type.name !== 'opaqueInline' || p.node.attrs['kind'] !== 'commentReference') return false;
      const el = this.ctx.ids.get(p.node.attrs['sourceId'] as string);
      return String((el?.value as { id?: unknown })?.id ?? '') === id;
    };
    const marker = (el: Element, like: Plan | undefined): Plan => ({ node: undefined, element: el, holders: like?.holders ?? [], runSource: undefined, fmtKey: '', marks: [] });
    // Ends first, so that an inserted start does not shift the index an end was computed from.
    // With no text left carrying the comment, the end goes just before its reference (the first plan
    // too), else at the paragraph's end.
    for (const [id, el] of commentEnds) {
      let last = -1;
      for (let i = plans.length - 1; i >= 0; i--) if (carries(plans[i]!, id)) { last = i; break; }
      const reference = last < 0 ? plans.findIndex((p) => isReference(p, id)) : -1;
      const at = last >= 0 ? last + 1 : reference >= 0 ? reference : plans.length;
      plans.splice(at, 0, marker(el, last >= 0 ? plans[last] : plans[at]));
    }
    // With no text left carrying the comment, the start goes just before its end, then before its
    // reference: the end is placed before the reference above, so a start put at the reference would
    // follow its own end, and the range would run on to the paragraph's end (ED-005 section 12.47).
    for (const [id, el] of commentStarts) {
      let first = plans.findIndex((p) => carries(p, id));
      if (first < 0) first = plans.findIndex((p) => p.element !== undefined && p.element.name.localPart === 'commentRangeEnd' && String((p.element.value as { id?: unknown }).id ?? '') === id);
      if (first < 0) first = plans.findIndex((p) => isReference(p, id));
      if (first < 0) first = plans.length;
      plans.splice(first, 0, marker(el, plans[first]));
    }
  }

  /**
   * The markers of the comments made in the editor that start or end in this paragraph (ED-005
   * section 7.4): a start before the first text carrying the comment, an end after the last and
   * after any end already there, so a reply's follow its parent's as Word and the engine write
   * them; markers at one place in the order of their reference atoms. An empty range (the
   * reference alone) has its markers just before its reference, and empty ranges side by side
   * nest, the later inside the earlier, as Word writes them: `start 0, start 1, end 1, reference
   * 1, end 0, reference 0` (Jason's Word file, 2026-09-29).
   */
  private placeNewCommentMarkers(plans: Plan[], node: PMNode): void {
    const bounds = this.commentBounds;
    if (!bounds || bounds.first.size === 0) return;
    const starts: string[] = [];
    const ends: string[] = [];
    for (const [id, p] of bounds.first) if (p === node) starts.push(id);
    for (const [id, p] of bounds.last) if (p === node) ends.push(id);
    // A comment on an empty paragraph is its reference run alone, with no range markers (Word 2010 and
    // Word 15, ED-005 section 12.33 item 2), where an empty range before a paragraph's first word has
    // them (section 7.3). Empty: nothing but comment references and markers.
    const bare = node.textContent === '' && node.childCount > 0 && Array.from({ length: node.childCount }, (_v, i) => node.child(i))
      .every((c) => c.type.name === 'opaqueInline' && (c.attrs['kind'] === 'commentReference' || EMPTY_PARAGRAPH_MARKERS.has(c.attrs['kind'] as string)));
    if (bare) {
      const whole = new Set(starts.filter((id) => ends.includes(id)));
      starts.splice(0, starts.length, ...starts.filter((id) => !whole.has(id)));
      ends.splice(0, ends.length, ...ends.filter((id) => !whole.has(id)));
    }
    if (starts.length === 0 && ends.length === 0) return;
    const byRef = (a: string, b: string): number => (bounds.refOrder.get(a) ?? Infinity) - (bounds.refOrder.get(b) ?? Infinity) || Number(a) - Number(b);
    const carries = (p: Plan, id: string): boolean => p.node !== undefined && p.node.marks.some((m) => m.type.name === 'comment' && String(m.attrs['id']) === id);
    const isRef = (p: Plan, id: string): boolean => p.node !== undefined && referenceIdOf(p.node, this.ctx.ids) === id;
    const markerOf = (p: Plan, local: string): string | undefined => p.element?.name.localPart === local ? String((p.element.value as { id?: unknown }).id ?? '') : undefined;
    const marker = (el: Element, like: Plan | undefined): Plan => ({ node: undefined, element: el, holders: like?.holders ?? [], runSource: undefined, fmtKey: '', marks: [] });
    const start = (id: string): Element => wmlEl.commentRangeStart(wmlFactory.createCommentRangeStart({ id: Number(id) })) as Element;
    const end = (id: string): Element => wmlEl.commentRangeEnd(wmlFactory.createCommentRangeEnd({ id: Number(id) })) as Element;
    const empty = new Set<string>();
    for (const id of ends.sort(byRef)) {
      let last = -1;
      for (let i = plans.length - 1; i >= 0; i--) if (carries(plans[i]!, id)) { last = i; break; }
      if (last >= 0) {
        let at = last + 1;
        while (at < plans.length && markerOf(plans[at]!, 'commentRangeEnd') !== undefined) at++;
        plans.splice(at, 0, marker(end(id), plans[last]));
        continue;
      }
      const r = plans.findIndex((p) => isRef(p, id));
      if (r < 0) continue;
      plans.splice(r, 0, marker(end(id), plans[r]));
      empty.add(id);
    }
    // A piece of an empty range side by side with this one: its start, its end or its reference.
    const emptyPiece = (p: Plan): boolean => {
      const own = markerOf(p, 'commentRangeStart') ?? markerOf(p, 'commentRangeEnd') ?? (p.node ? referenceIdOf(p.node, this.ctx.ids) : undefined);
      return own !== undefined && empty.has(own);
    };
    for (const id of starts.sort(byRef)) {
      const first = plans.findIndex((p) => carries(p, id));
      if (first >= 0) { plans.splice(first, 0, marker(start(id), plans[first])); continue; }
      let at = plans.findIndex((p) => markerOf(p, 'commentRangeEnd') === id);
      if (at < 0) at = plans.findIndex((p) => isRef(p, id));
      if (at < 0) continue;
      while (at > 0 && emptyPiece(plans[at - 1]!)) at--;
      plans.splice(at, 0, marker(start(id), plans[at]));
    }
  }

  /** Group consecutive plans by their holder at `depth`; a group with a holder becomes a copy of it around the rest. */
  private rebuildInline(plans: Plan[], depth: number): Element[] {
    const out: Element[] = [];
    let i = 0;
    while (i < plans.length) {
      const h = plans[i]!.holders[depth];
      let j = i;
      while (j < plans.length && plans[j]!.holders[depth] === h) j++;
      const group = plans.slice(i, j);
      if (h) out.push(this.rebuildHolder(h, group, depth + 1));
      else out.push(...this.rebuildRuns(group));
      i = j;
    }
    return out;
  }

  private rebuildHolder(h: Element, group: Plan[], depth: number): Element {
    // A paragraph-level kept `mc:AlternateContent` holds its selected branch's items (ED-009).
    if (isAlternateContent(h)) return this.rebuildAlternateContent(h, this.rebuildInline(group, depth));
    const key = itemsKeyOf(h.value as object);
    const value = copyWithout(h.value as Record<string, unknown>, [key]);
    value[key] = this.rebuildInline(group, depth);
    if (revisionKindOf(h) !== undefined) {
      if (this.emittedHolders.has(h)) value['id'] = this.nextId();
      this.emittedHolders.add(h);
    }
    return { name: h.name, value };
  }

  /** Plans with no holder left: runs, paragraph-level elements, inline content controls. */
  private rebuildRuns(group: Plan[]): Element[] {
    const out: Element[] = [];
    let cur: { runSource: string | undefined; fmtKey: string; changeKey: string; formatChange: FormatChangeAttrs | null | undefined; marks: readonly Mark[]; items: Element[]; text: string; deleted: boolean } | undefined;
    // The run-level kept `mc:AlternateContent` the items being collected belong to (ED-009): its
    // text and items gather here and go into the run as the element rebuilt around them. A change of
    // run (a formatting split) or of element closes it; the same element met again opens a later piece.
    let bucket: { alt: string; el: Element; items: Element[]; text: string } | undefined;
    // Consecutive text pieces of one run (split by source runs that merged, or by typing) become one
    // w:t, or one w:delText inside a deletion (ED-005 section 5.3).
    const flushText = (): void => {
      if (bucket && bucket.text !== '') { bucket.items.push(cur?.deleted ? delText(bucket.text) : t(bucket.text)); bucket.text = ''; }
      if (cur && cur.text !== '') { cur.items.push(cur.deleted ? delText(cur.text) : t(cur.text)); cur.text = ''; }
    };
    const closeBucket = (): void => {
      if (!bucket) return;
      flushText();
      if (bucket.items.length > 0 && cur) cur.items.push(this.rebuildAlternateContent(bucket.el, bucket.items));
      bucket = undefined;
    };
    const flush = (): void => {
      closeBucket();
      flushText();
      if (cur && cur.items.length > 0) out.push(this.buildRun(cur.runSource, cur.marks, cur.items, cur.formatChange));
      cur = undefined;
    };
    for (const plan of group) {
      if (plan.element) { flush(); out.push(plan.element); continue; }
      const n = plan.node!;
      if (n.type.name === 'opaqueInline' && plan.runSource === undefined) {
        flush();
        const el = this.opaqueElement(n);
        if (el) out.push(el);
        continue;
      }
      if (n.type.name === 'sdtInline') { flush(); out.push(this.rebuildSdtInline(n)); continue; }
      // A run is one source, one formatting, one formatting change and one side of a deletion.
      const changeKey = plan.formatChange === undefined ? '' : JSON.stringify(plan.formatChange);
      if (!cur || cur.runSource !== plan.runSource || cur.fmtKey !== plan.fmtKey || cur.changeKey !== changeKey || cur.deleted !== (plan.deleted === true)) {
        flush();
        cur = { runSource: plan.runSource, fmtKey: plan.fmtKey, changeKey, formatChange: plan.formatChange, marks: plan.marks, items: [], text: '', deleted: plan.deleted === true };
      }
      if (bucket && bucket.alt !== plan.alt) closeBucket();
      if (plan.alt !== undefined && !bucket) {
        flushText();
        const el = this.ctx.ids.get(plan.alt);
        if (el && isAlternateContent(el)) bucket = { alt: plan.alt, el, items: [], text: '' };
      }
      const into = bucket ?? cur;
      if (n.isText) { into.text += n.text ?? ''; continue; }
      flushText();
      const item = this.itemOf(n, this.ctx.revisions === 'content' ? cur.deleted : undefined);
      if (item) into.items.push(item);
    }
    flush();
    return out;
  }

  /**
   * A kept `mc:AlternateContent` written back with its selected branch's content replaced by what
   * the editor holds for it and every other branch, however many Choices and the Fallback, the
   * original branch object in its place (ED-009 section 3.3). A later piece of the same element in
   * one export (a split, section 3.4) carries every other branch present and empty, so a reader
   * taking any branch takes the same branch in every piece.
   */
  private rebuildAlternateContent(el: Element, content: Element[]): Element {
    const v = el.value as mce.AlternateContent;
    const branch = selectedBranchOf(v);
    const first = !this.emittedAlternates.has(el);
    this.emittedAlternates.add(el);
    const choice = (v.choice ?? []).map((c) => (c.any === branch ? { ...c, any: content } : first ? c : { ...c, any: [] }));
    const value: mce.AlternateContent = { ...v, choice };
    if (v.fallback) value.fallback = v.fallback.any === branch ? { ...v.fallback, any: content } : first ? v.fallback : { ...v.fallback, any: [] };
    this.changed = true;
    return { name: el.name, value };
  }

  private rebuildSdtInline(n: PMNode): Element {
    const original = this.unchanged(n);
    if (original) return original;
    const id = n.attrs['sourceId'] as string | null;
    const orig = id ? (this.ctx.ids.get(id) as Element<wml.SdtRun> | undefined) : undefined;
    const value: wml.SdtRun = orig ? copyWithout(orig.value, ['sdtContent']) : { TYPE_NAME: 'org_docx4j_wml.SdtRun' };
    const sdtPr = n.attrs['sdtPr'] as unknown;
    if (!orig || !plainEqual(sdtPr, orig.value.sdtPr ? toPlain(orig.value.sdtPr) : null)) {
      if (sdtPr === null) delete value.sdtPr; else value.sdtPr = exportedSdtPr(sdtPr);
    }
    value.sdtContent = { TYPE_NAME: 'org_docx4j_wml.CTSdtContentRun', content: this.rebuildInline(this.plansOf(n), 0) as wml.CTSdtContentRun['content'] };
    return { name: orig?.name ?? { namespaceURI: W_NS, localPart: 'sdt' }, value };
  }

  /**
   * The `w:r` item for an inline node: text, tab, break, or the original element of an atom; in
   * the body a field code as its run's deletion says (`deleted`), `w:delInstrText` inside a
   * deletion and `w:instrText` outside one (ED-005 section 5.3).
   */
  private itemOf(n: PMNode, deleted?: boolean): Element | undefined {
    switch (n.type.name) {
      case 'text': return t(n.text ?? '');
      case 'tab': return tab();
      case 'hardBreak': {
        const kind = n.attrs['kind'] as string;
        const el = br(kind === 'page' || kind === 'column' ? kind : undefined);
        const clear = n.attrs['clear'] as string | null;
        if (clear) (el.value as wml.Br).clear = clear as wml.Br['clear'];
        return el;
      }
      case 'image':
      case 'footnoteRef':
        return this.ctx.ids.get(n.attrs['sourceId'] as string);
      case 'opaqueInline': {
        const el = this.opaqueElement(n);
        if (!el || deleted === undefined) return el;
        const local = el.name.localPart;
        if (local === 'instrText' && deleted) return wmlEl.delInstrText(el.value as wml.Text) as Element;
        if (local === 'delInstrText' && !deleted) return wmlEl.instrText(el.value as wml.Text) as Element;
        return el;
      }
      default:
        throw new Error(`unexpected inline node ${n.type.name}`);
    }
  }

  /**
   * An opaque atom's element: its original, or, for a page-break marker made by Re-paginate in a
   * room (ED-004 section 3.8), which has no source, `<w:lastRenderedPageBreak/>` built from its kind.
   */
  private opaqueElement(n: PMNode): Element | undefined {
    const id = n.attrs['sourceId'] as string | null;
    const original = id ? this.ctx.ids.get(id) : undefined;
    if (original) return original;
    if (n.attrs['kind'] === 'lastRenderedPageBreak') return wmlEl.lastRenderedPageBreak({ TYPE_NAME: 'org_docx4j_wml.R.LastRenderedPageBreak' }) as Element;
    return undefined;
  }

  /**
   * A run: the source run's `rPr` copied with only the differing formatting applied, or one from
   * the marks alone; in the body its `w:rPrChange` as its formatting change says (ED-005 section
   * 5.3), the source's own element and id while the change is the one read from it.
   */
  private buildRun(runSource: string | undefined, marks: readonly Mark[], items: Element[], formatChange?: FormatChangeAttrs | null): Element<wml.R> {
    const source = runSource ? (this.ctx.ids.get(runSource) as Element<wml.R> | undefined) : undefined;
    const run = runOf(items, runPropertiesFor(source?.value.rPr, marks));
    if (source) {
      // Word's revision-save ids: not formatting, but part of what a run "is"; keep them on a rebuilt run.
      for (const k of ['rsidR', 'rsidRPr', 'rsidDel'] as const) if (source.value[k] !== undefined) run.value[k] = source.value[k];
    }
    if (formatChange !== undefined) {
      const kept = source?.value.rPr?.rPrChange;
      if (formatChange === null) {
        if (run.value.rPr?.rPrChange) {
          delete run.value.rPr.rPrChange;
          if (Object.keys(run.value.rPr).every((k) => k === 'TYPE_NAME' || k === 'PARENT')) delete run.value.rPr;
        }
      } else if (!kept || !plainEqual(formatChangeAttrs(kept), formatChange)) {
        (run.value.rPr ??= { TYPE_NAME: 'org_docx4j_wml.RPr' }).rPrChange = rPrChangeOf(formatChange, this.nextId());
      }
    }
    return run;
  }
}

/** What a paragraph's attributes were at import (`ExportState.paragraphBase`). */
function paragraphBaseOf(id: string | null, orig: Element<wml.P> | undefined, ctx: ProjectionContext): Omit<ParagraphAttrs, 'sourceId'> {
  const imported = id ? ctx.imported.get(id) : undefined;
  if (imported?.type.name === 'paragraph') {
    const { sourceId: _ignored, ...attrs } = imported.attrs as unknown as ParagraphAttrs;
    void _ignored;
    return attrs;
  }
  return orig ? projectedParagraphAttrs(orig.value.pPr, ctx.revisions) : paragraphAttrsOf(undefined);
}

/**
 * The `w:pPr` a paragraph node's attributes mean, as the export writes it but for the tracked
 * changes: the original's copied with the attributes that differ from the import applied. What
 * the tracker records as a paragraph's properties before and after a change (ED-005 section 5.4).
 */
export function paragraphPropertiesOf(node: PMNode, ctx: ProjectionContext): wml.PPr | undefined {
  const id = node.attrs['sourceId'] as string | null;
  const orig = id ? (ctx.ids.get(id) as Element<wml.P> | undefined) : undefined;
  const attrs = node.attrs as unknown as ParagraphAttrs;
  const copied = orig ? undefined : copiedPropertiesOf(attrs.copiedFrom, ctx);
  const base = copied ? projectedParagraphAttrs(copied, ctx.revisions) : paragraphBaseOf(id, orig, ctx);
  return applyParagraphAttrs(orig?.value.pPr ? deepCopy(orig.value.pPr) : copied, attrs, base);
}

/**
 * The `w:pPr` a paragraph made with a copy of another's attributes starts from (`copiedFrom`, Enter's
 * first half; ED-005 section 12.102): a deep copy of that element's, as Word copies the whole
 * properties to the new mark, less what belongs to the one paragraph alone: the section it ends, the
 * mark's tracked change and the properties' change with their ids (the new mark's own revision is
 * written afresh from the attribute), and the preserved changes. Undefined when there is no such
 * element (the paragraph was made from nothing, or a tree run replaced its source), and the
 * paragraph is then built from its attributes alone, as before.
 */
function copiedPropertiesOf(copiedFrom: string | null, ctx: ProjectionContext): wml.PPr | undefined {
  const el = copiedFrom ? (ctx.ids.get(copiedFrom) as Element<wml.P> | undefined) : undefined;
  if (!el?.value.pPr) return undefined;
  const pPr = deepCopy(el.value.pPr) as wml.PPr & { rPr?: Record<string, unknown>; numPr?: Record<string, unknown> };
  delete pPr.sectPr;
  delete pPr.pPrChange;
  if (pPr.rPr) {
    for (const k of ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange']) delete pPr.rPr[k];
    if (Object.keys(pPr.rPr).every((k) => k === 'TYPE_NAME' || k === 'PARENT')) delete pPr.rPr;
  }
  if (pPr.numPr) delete pPr.numPr['numberingChange'];
  return pPr;
}

/**
 * The `w:rPr` a set of marks means: a copy of the source run's with only the formatting that
 * differs applied (so unseen properties survive), or, with no source, one built from the marks
 * alone. Shared by export and by the effective-formatting reads (E1.b), which resolve it.
 */
export function runPropertiesFor(sourceRPr: wml.RPr | undefined, marks: readonly Mark[]): wml.RPr | undefined {
  const desired = formattingOf(marks);
  let rPr: wml.RPr | undefined;
  if (sourceRPr) {
    rPr = deepCopy(sourceRPr);
    const current = readRunOptions(rPr);
    const diff: RunOptions = {};
    for (const key of RUN_OPTION_KEYS) {
      if (desired.fmt[key] !== current[key]) (diff as Record<string, unknown>)[key] = desired.fmt[key];
    }
    if (diff.size === 0) { delete rPr.sz; delete rPr.szCs; delete diff.size; }
    if (diff.underline === 'Mixed') delete diff.underline;
    applyRunOptions(rPr, diff);
    const rtlNow = rPr.rtl !== undefined && rPr.rtl.val !== false;
    if (desired.rtl !== rtlNow) { if (desired.rtl) rPr.rtl = {}; else delete rPr.rtl; }
    return rPr;
  }
  const opts: RunOptions = {};
  const f = desired.fmt;
  if (f.bold) opts.bold = true;
  if (f.italic) opts.italic = true;
  if (f.underline !== 'None' && f.underline !== 'Mixed') opts.underline = f.underline;
  if (f.strikeThrough) opts.strikeThrough = true;
  if (f.doubleStrikeThrough) opts.doubleStrikeThrough = true;
  if (f.subscript) opts.subscript = true;
  if (f.superscript) opts.superscript = true;
  if (f.name !== '') opts.name = f.name;
  if (f.size !== 0) opts.size = f.size;
  if (f.color !== '') opts.color = f.color;
  if (f.highlightColor !== null) opts.highlightColor = f.highlightColor;
  if (f.style !== '') opts.style = f.style;
  rPr = applyRunOptions({ TYPE_NAME: 'org_docx4j_wml.RPr' }, opts);
  if (desired.rtl) rPr.rtl = {};
  // A run built from marks alone gets no empty w:rPr.
  if (Object.keys(rPr).every((k) => k === 'TYPE_NAME' || k === 'PARENT')) return undefined;
  return rPr;
}

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** The inline node types that carry `runSource` and `fmt` as attributes (UPSTREAM-2). */
const ATOMS = new Set(['tab', 'hardBreak', 'image', 'footnoteRef', 'opaqueInline']);

const HYPERLINK = 'org_docx4j_wml.P.Hyperlink';

const RUN_OPTION_KEYS = ['bold', 'italic', 'underline', 'strikeThrough', 'doubleStrikeThrough', 'subscript', 'superscript', 'name', 'size', 'color', 'highlightColor', 'style'] as const;

/** The array a run holder keeps its items in (the model's three names for it). */
function itemsKeyOf(value: object): 'content' | 'customXmlOrSmartTagOrSdt' | 'accOrBarOrBox' {
  const v = value as Record<string, unknown>;
  if (Array.isArray(v['customXmlOrSmartTagOrSdt'])) return 'customXmlOrSmartTagOrSdt';
  if (Array.isArray(v['accOrBarOrBox'])) return 'accOrBarOrBox';
  if (Array.isArray(v['content'])) return 'content';
  const tn = typeNameOf(value);
  return tn === 'org_docx4j_wml.RunIns' || tn === 'org_docx4j_wml.RunDel' ? 'customXmlOrSmartTagOrSdt' : tn === 'org_docx4j_wml.RunTrackChange' ? 'accOrBarOrBox' : 'content';
}

/** The formatting marks of a node as a comparable key, for grouping into runs. */
export function formattingKey(marks: readonly Mark[]): string {
  const parts = marks
    .filter((m) => (FORMATTING_MARKS as readonly string[]).includes(m.type.name))
    .map((m) => [m.type.name, m.attrs] as const)
    .sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0));
  return JSON.stringify(parts);
}

/** The inverse of `formattingKey`: the marks a key describes (an atom's `fmt` attribute). */
export function marksFromKey(key: string, schema: Schema): Mark[] {
  if (key === '') return [];
  const parts = JSON.parse(key) as [string, Record<string, unknown>][];
  return parts.map(([name, attrs]) => schema.marks[name]!.create(attrs));
}

/** What the marks say the run's direct formatting should be, in `readRunOptions`' vocabulary. */
export function formattingOf(marks: readonly Mark[]): { fmt: RunFormatting; rtl: boolean } {
  const by = new Map(marks.map((m) => [m.type.name, m] as const));
  const font = by.get('font')?.attrs as FontAttrs | undefined;
  const strike = by.get('strike');
  const vertAlign = by.get('vertAlign')?.attrs['val'];
  return {
    fmt: {
      bold: by.has('bold'),
      italic: by.has('italic'),
      underline: (by.get('underline')?.attrs['val'] as RunFormatting['underline'] | undefined) ?? 'None',
      strikeThrough: strike !== undefined && strike.attrs['double'] !== true,
      doubleStrikeThrough: strike !== undefined && strike.attrs['double'] === true,
      subscript: vertAlign === 'subscript',
      superscript: vertAlign === 'superscript',
      name: font?.name ?? '',
      size: font?.size ?? 0,
      color: font?.color ?? '',
      highlightColor: font?.highlight ?? null,
      style: (by.get('runStyle')?.attrs['id'] as string | null | undefined) ?? '',
    },
    rtl: by.has('rtl'),
  };
}

/** Write the paragraph attributes that differ from the imported ones onto a copy of `w:pPr`. */
export function applyParagraphAttrs(pPr: wml.PPr | undefined, attrs: ParagraphAttrs, base: Omit<ParagraphAttrs, 'sourceId'>): wml.PPr | undefined {
  let out = pPr;
  const ensure = (): wml.PPr => (out ??= { TYPE_NAME: 'org_docx4j_wml.PPr' });
  if (attrs.pStyle !== base.pStyle) {
    if (attrs.pStyle === null) delete out?.pStyle; else ensure().pStyle = { val: attrs.pStyle };
  }
  if (attrs.jc !== base.jc) {
    if (attrs.jc === null) delete out?.jc; else ensure().jc = { val: attrs.jc as wml.Jc['val'] };
  }
  if (!plainEqual(attrs.ind, base.ind)) {
    if (attrs.ind === null) delete out?.ind;
    else {
      const p = ensure();
      const ind = (p.ind ??= {});
      setIndent(ind, 'start', attrs.ind.start);
      setIndent(ind, 'end', attrs.ind.end);
      if (attrs.ind.firstLine === null) delete ind.firstLine; else { ind.firstLine = attrs.ind.firstLine; delete ind.hanging; }
      if (attrs.ind.hanging === null) delete ind.hanging; else { ind.hanging = attrs.ind.hanging; delete ind.firstLine; }
    }
  }
  if (!plainEqual(attrs.spacing, base.spacing)) {
    if (attrs.spacing === null) delete out?.spacing;
    else {
      const p = ensure();
      const sp = (p.spacing ??= {});
      if (attrs.spacing.before === null) delete sp.before; else sp.before = attrs.spacing.before;
      if (attrs.spacing.after === null) delete sp.after; else sp.after = attrs.spacing.after;
      if (attrs.spacing.line === null) delete sp.line; else sp.line = attrs.spacing.line;
      if (attrs.spacing.lineRule === null) delete sp.lineRule; else sp.lineRule = attrs.spacing.lineRule as wml.PPrBase.Spacing['lineRule'];
    }
  }
  if (attrs.bidi !== base.bidi) {
    if (attrs.bidi) ensure().bidi = {}; else delete out?.bidi;
  }
  if (!plainEqual(attrs.numPr, base.numPr)) {
    if (attrs.numPr === null) delete out?.numPr;
    else {
      // The element's other children (a tracked `w:ins`, `w:numberingChange`) stay as they were.
      const numPr = (ensure().numPr ??= { TYPE_NAME: 'org_docx4j_wml.PPrBase.NumPr' });
      numPr.ilvl = { val: attrs.numPr.ilvl };
      if (attrs.numPr.numId === null) delete numPr.numId; else numPr.numId = { val: Number(attrs.numPr.numId) };
    }
  }
  if (!plainEqual(attrs.tabs, base.tabs)) {
    // The element is replaced whole (ED-003 section 8.5): its stops are the attribute's, in its order.
    if (attrs.tabs === null || attrs.tabs.length === 0) delete out?.tabs;
    else {
      ensure().tabs = {
        TYPE_NAME: 'org_docx4j_wml.Tabs',
        tab: attrs.tabs.map((t) => ({ TYPE_NAME: 'org_docx4j_wml.CTTabStop', pos: t.pos, val: t.val as wml.STTabJc, ...(t.leader === null ? {} : { leader: t.leader as wml.STTabTlc }) })),
      };
    }
  }
  if (!plainEqual(attrs.sectPr, base.sectPr)) {
    if (attrs.sectPr === null) delete out?.sectPr; else ensure().sectPr = revive(attrs.sectPr) as wml.SectPr;
  }
  if (!plainEqual(attrs.markRPr, base.markRPr)) {
    if (attrs.markRPr === null) delete out?.rPr; else ensure().rPr = revive(attrs.markRPr) as wml.ParaRPr;
  }
  if (out && Object.keys(out).every((k) => k === 'TYPE_NAME' || k === 'PARENT')) return undefined;
  return out;
}

/**
 * `w:ind` start or end: written to the attribute the element already uses (`w:start` or the
 * transitional `w:left`), and to the transitional one when it uses neither, since that is what
 * Word writes and what every reader understands (ED-001 section 3.7 reads both as aliases).
 */
function setIndent(ind: wml.PPrBase.Ind, side: 'start' | 'end', value: number | null): void {
  const logical = side;
  const physical = side === 'start' ? 'left' : 'right';
  if (value === null) { delete ind[logical]; delete ind[physical]; return; }
  if (ind[logical] !== undefined && ind[physical] === undefined) ind[logical] = value;
  else ind[physical] = value;
}
