// SPDX-License-Identifier: GPL-3.0-or-later
//
// The console's `selection` binding (ED-002 section 6.3; ED-003 section 7.3, section 9 item 31):
// what the view's selection is, as content-API objects over the tree. Two halves, because the
// tree a run edits is built after the export step and elements of edited paragraphs are rebuilt
// on the way in (ED-002 section 10.4 item 7), so nothing the state holds can find them:
// `selectionSpec` reads the editor state into indexes and offsets, and `selectionBinding` builds
// the objects over the tree once the run has it. The two are what the `api` lines of section 7.2
// and the shim's `getSelection()` run over, so the palette's test can build the same objects
// headless.
import { Body, Font, Range, Paragraph, childrenOf, typeNameOf, type Element, type FontTracking } from '@docx4j/core-ts';
import type { EditorState } from 'prosemirror-state';
import type { Node as PMNode, ResolvedPos } from 'prosemirror-model';
import { paragraphText, type FindContext } from './find.mjs';
import { isDeleted } from './revisions.mjs';
import type { SourceIds } from './ids.mjs';
import { sessionOf } from './commands/spec.mjs';
import { projectedTableCells } from './import.mjs';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';

/** The selection as indexes and offsets: the top-level blocks it touches, and the paragraph and span when it lies within one paragraph. */
export interface SelectionSpec {
  /** Top-level block indexes the selection touches, first to last. */
  blocks: number[];
  /**
   * The paragraph holding both ends, as the index path from the top-level block down through
   * block-level controls and table parts, with the span in the paragraph's text (the content
   * API's offsets: tabs and breaks count one, other atoms none). Undefined across paragraphs.
   */
  range?: { path: number[]; start: number; end: number };
  /**
   * The selection across paragraphs, as one of the engine's ranges since core-ts 0.6.0 (CR-014; ED-008
   * section 8.8 item 8): the first paragraph and the offset in it, the last and the offset in it, and
   * whether the selection takes the last paragraph's mark.
   */
  span?: { path: number[]; start: number; endPath: number[]; end: number; holdsMark: boolean };
  /**
   * The paragraphs the selection covers, first to last, as index paths like `range`'s: a caret's paragraph alone, and
   * none a selection only reaches the start of. Office JS's `selection.paragraphs`, where `blocks` would give every
   * paragraph of a table the caret is in (ED-005 section 12.129). A cell selection's are its cells' paragraphs.
   */
  paragraphs?: number[][];
}

/** The selection over the tree: the paragraphs, their text and addresses, the range and the font. */
export interface SelectionBinding {
  paragraphs: Paragraph[];
  text: string;
  addresses: string[];
  /** The content API's `Range` over the caret or the selected text when it lies within one paragraph. */
  range: Range | undefined;
  /**
   * The selection as one range: `range` within one paragraph, else the engine's range from the first
   * selected paragraph's offset to the last's (`Range.spansParagraphs`), the last mark held when the
   * selection takes it. What the console's `getSelection()` reads, deletes and replaces (ED-008 section
   * 8.8 item 8); `range` and `font` keep their one-paragraph form for the `api` lines.
   */
  selectionRange: Range | undefined;
  /** The range's font, or a font over every run of the selected paragraphs. */
  font: Font;
}

/**
 * The caret's offset in a paragraph's text as the content API counts it: `pos` is relative to the
 * paragraph's content start. An atom counts what the engine reads from it (`atomText`: a tab or a
 * break one, a symbol its character, a picture or a note reference nothing); `ids` finds a
 * symbol's or a simple field's element, without which they count nothing (ED-005 section 12.1).
 */
export function textOffsetIn(paragraph: PMNode, pos: number, ids?: SourceIds, context?: FindContext): number {
  // The same reading as find's (`paragraphText`): deleted and hidden text and a placeholder count nothing.
  let offset = 0;
  for (const piece of paragraphText(paragraph, 0, ids, context).pieces) {
    if (piece.pos >= pos) break;
    offset += piece.text ? Math.min(piece.end - piece.start, pos - piece.pos) : piece.end - piece.start;
  }
  return offset;
}

/** The nearest paragraph up the chain from a position, with its depth. */
function paragraphDepth($pos: ResolvedPos): number {
  let depth = $pos.depth;
  while (depth >= 1 && $pos.node(depth).type.name !== 'paragraph') depth--;
  return depth;
}

export function selectionSpec(state: EditorState): SelectionSpec {
  const { from, to, $from, $to } = state.selection;
  const blocks: number[] = [];
  // A block the selection only reaches the start of is not selected (the selection bar's runs to the next
  // paragraph's start to take the mark; Word counts nothing of that paragraph): the same rule as `blocksEnd` in
  // `commands/paragraph.mts` (ED-005 section 12.81 item 5).
  const end = to > from && $to.parentOffset === 0 && $to.parent.isTextblock && $to.depth > 0 ? $to.before() : to;
  state.doc.forEach((node, offset, index) => { if (offset + node.nodeSize > from && offset < end) blocks.push(index); });
  const spec: SelectionSpec = { blocks };
  const session = sessionOf(state);
  const ids = session?.ctx.ids;
  const context = { resolver: session?.pkg.propertyResolverOrUndefined };
  const pathOf = ($pos: ResolvedPos, depth: number): number[] => { const path: number[] = []; for (let d = 1; d <= depth; d++) path.push($pos.index(d - 1)); return path; };
  const depth = paragraphDepth($from);
  // The paragraphs, through every range (a cell selection's ranges are its cells; a text selection has one).
  const paragraphs: number[][] = [];
  if (state.selection.empty) {
    if (depth >= 1) paragraphs.push(pathOf($from, depth));
  } else {
    const ranges = state.selection.ranges;
    ranges.forEach((r, i) => {
      const rangeEnd = i === ranges.length - 1 && ranges.length === 1 ? end : r.$to.pos;
      if (rangeEnd <= r.$from.pos) return;
      state.doc.nodesBetween(r.$from.pos, rangeEnd, (node, pos) => {
        if (node.type.name !== 'paragraph') return true;
        const $in = state.doc.resolve(pos + 1);
        paragraphs.push(pathOf($in, $in.depth));
        return false;
      });
    });
  }
  spec.paragraphs = paragraphs;
  if (depth >= 1 && paragraphDepth($to) === depth && $from.start(depth) === $to.start(depth)) {
    const paragraph = $from.node(depth);
    const start = $from.start(depth);
    spec.range = { path: pathOf($from, depth), start: textOffsetIn(paragraph, from - start, ids, context), end: textOffsetIn(paragraph, to - start, ids, context) };
  } else if (depth >= 1 && to > from) {
    // Across paragraphs: the last paragraph is the one holding `end`, or, where the selection runs to the
    // next block's start (the mark taken), the paragraph before that start.
    const holdsMark = end !== to;
    const $end = state.doc.resolve(holdsMark ? Math.max(end - 1, 0) : to);
    const endDepth = paragraphDepth($end);
    if (endDepth >= 1 && $end.start(endDepth) !== $from.start(depth)) {
      const first = $from.node(depth);
      const last = $end.node(endDepth);
      const lastStart = $end.start(endDepth);
      const endOffset = holdsMark ? textOffsetIn(last, last.content.size, ids, context) : textOffsetIn(last, to - lastStart, ids, context);
      spec.span = { path: pathOf($from, depth), start: textOffsetIn(first, from - $from.start(depth), ids, context), endPath: pathOf($end, endDepth), end: endOffset, holdsMark };
    }
  }
  return spec;
}

/** The block-level children of an element as the projection counts them (no proof marks or bookmark ends). */
function projectedChildren(element: Element): Element[] {
  return (childrenOf(element.value as object) ?? []).filter((e) => {
    const local = e.name.localPart;
    return typeNameOf(e) !== undefined && local !== 'proofErr' && local !== 'bookmarkStart' && local !== 'bookmarkEnd';
  });
}

/** Every paragraph within an element, in order: the element itself, or those of a control, a table, a row or a cell. */
export function paragraphsWithin(element: Element): Element[] {
  if (typeNameOf(element) === 'org_docx4j_wml.P') return [element];
  const out: Element[] = [];
  for (const child of childrenOf(element.value as object) ?? []) out.push(...paragraphsWithin(child));
  return out;
}

/** The selection's objects over the tree the run edits (`body` is the run's content API). */
export function selectionBinding(body: Body, spec: SelectionSpec): SelectionBinding {
  // A path steps through a table as the projection has it: a row, then a cell of it, with a vertical merge's
  // continuation `w:tc`s left out (`projectedTableCells`), so a path through a row a merge reaches lands on its cell.
  const tables = new Map<object, Element<wml.Tc>[][]>();
  const cellsOf = (tbl: Element): Element<wml.Tc>[][] => {
    let rows = tables.get(tbl.value as object);
    if (!rows) tables.set(tbl.value as object, rows = projectedTableCells(tbl.value as wml.Tbl));
    return rows;
  };
  const paragraphAt = (path: readonly number[]): Paragraph | undefined => {
    let el: Element | undefined = body.content[path[0]!];
    for (let d = 1; d < path.length && el; d++) {
      if (typeNameOf(el) === 'org_docx4j_wml.Tbl' && d + 1 < path.length) { el = cellsOf(el)[path[d]!]?.[path[d + 1]!]; d++; }
      else el = projectedChildren(el)[path[d]!];
    }
    return el && typeNameOf(el) === 'org_docx4j_wml.P' ? body.paragraphFor(el.value as never) : undefined;
  };
  const paragraphs: Paragraph[] = [];
  const addresses: string[] = [];
  const add = (view: Paragraph | undefined): void => {
    if (!view || paragraphs.includes(view)) return;
    paragraphs.push(view);
    const a = body.addressOf(view);
    if (a) addresses.push(a);
  };
  if (spec.paragraphs) for (const path of spec.paragraphs) add(paragraphAt(path));
  else {
    for (const i of spec.blocks) {
      const el = body.content[i];
      if (!el) continue;
      for (const p of paragraphsWithin(el)) add(body.paragraphFor(p.value as never));
    }
  }
  let range: Range | undefined;
  let selectionRange: Range | undefined;
  if (spec.range) {
    const view = paragraphAt(spec.range.path);
    if (view) range = selectionRange = new Range(view, spec.range.start, spec.range.end);
  } else if (spec.span) {
    const first = paragraphAt(spec.span.path);
    const last = paragraphAt(spec.span.endPath);
    if (first && last) selectionRange = new Range(first, spec.span.start, spec.span.end, false, spec.span.holdsMark, last);
  }
  // Across paragraphs, one Font over all their runs, tracked as each paragraph's own is when the
  // package is tracking: the tracker, and this author's insertions from every paragraph (ED-005
  // section 12.7: without it a formatting change there went untracked, the agreement suite found).
  const tracking = (): FontTracking | undefined => {
    const each = paragraphs.map((p) => p.fontTracking()).filter((t): t is FontTracking => t !== undefined);
    return each.length === 0 ? undefined : { tracker: each[0]!.tracker, ownInsertions: new Set(each.flatMap((t) => [...t.ownInsertions])) };
  };
  const font = range ? range.font : new Font(() => paragraphs.flatMap((p) => p.runs.map((r) => r.value as never)), tracking);
  return { paragraphs, text: paragraphs.map((p) => p.text).join('\n'), addresses, range, selectionRange, font };
}
