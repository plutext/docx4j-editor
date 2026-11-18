// SPDX-License-Identifier: GPL-3.0-or-later
//
// The revisions shown in Word only (ED-005 section 6.5) resolved through the engine, which lists and
// resolves them since core-ts 0.3.0 (CR-002 section 35): a table's property records, a cell inserted
// or deleted, a section's properties, a paragraph mark's formatting. The edit layer keeps them as
// they are, some in node attributes and some only in the element a node came from (a grid's record,
// a list's numbering), so they are resolved on the tree in a tree run, after the command's
// transaction has resolved the edit layer's own changes: the four commands' `settle` half.
//
// A panel entry names its revision by where it is, since a tree run's elements are not the
// projection's: its top-level block (the export writes one element per block), its kind, and its
// place among that block's revisions of the kind. A table goes by its order among the block's
// tables, since the engine lists nothing for a table whose records change nothing.
import type { Node as PMNode } from 'prosemirror-model';
import { childrenOf, typeNameOf, type Body, type TrackedChange, type Element } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { PreservedGroup } from './review.mjs';

/** The kinds the engine lists, and so resolves one at a time; a cell merge and a list's numbering it resolves only with the rest ("all"). */
export type PreservedTargetKind = 'tableProperties' | 'cellIns' | 'cellDel' | 'sectPrChange' | 'markRPrChange' | 'mathIns' | 'mathDel' | 'holderIns' | 'holderDel';

/** Where a revision shown in Word only is, for the engine to find it in a tree run. */
export interface PreservedTarget {
  /** The top-level block's index, or the body's last section. */
  block: number | 'final';
  kind: PreservedTargetKind;
  /** Its place among the block's revisions of the kind; a table's among the block's tables. */
  ordinal: number;
}

const LISTED = new Set<string>(['tableProperties', 'cellIns', 'cellDel', 'sectPrChange', 'markRPrChange', 'mathIns', 'mathDel', 'holderIns', 'holderDel']);

/** The top-level block holding a position, with its index and bounds. */
function blockAt(doc: PMNode, pos: number): { index: number; from: number; node: PMNode } | undefined {
  let found: { index: number; from: number; node: PMNode } | undefined;
  doc.forEach((node, offset, index) => { if (!found && pos >= offset && pos < offset + node.nodeSize) found = { index, from: offset, node }; });
  return found;
}

/**
 * Where a panel entry's revision is, among `groups` (the document's, `preservedGroupsOf`);
 * undefined for one the engine does not list (a cell merge, a list's numbering), which only
 * Accept all and Reject all resolve, as in Word's API.
 */
export function preservedTargetOf(doc: PMNode, group: PreservedGroup, groups: readonly PreservedGroup[]): PreservedTarget | undefined {
  if (!LISTED.has(group.kind)) return undefined;
  const kind = group.kind as PreservedTargetKind;
  if (group.pos >= doc.content.size) return kind === 'sectPrChange' ? { block: 'final', kind, ordinal: 0 } : undefined;
  const block = blockAt(doc, group.pos);
  if (!block) return undefined;
  if (kind === 'tableProperties') {
    const tables: number[] = block.node.type.name === 'table' ? [block.from] : [];
    block.node.descendants((node, pos) => { if (node.type.name === 'table') tables.push(block.from + 1 + pos); return true; });
    const ordinal = tables.indexOf(group.pos);
    return ordinal < 0 ? undefined : { block: block.index, kind, ordinal };
  }
  const end = block.from + block.node.nodeSize;
  const ordinal = groups.filter((g) => g.kind === kind && g.pos >= block.from && g.pos < end && g.pos < group.pos).length;
  return { block: block.index, kind, ordinal };
}

/** A tree item's value: an element's `value`, or the item itself. */
function valueOf(item: unknown): object | undefined {
  if (item === null || typeof item !== 'object') return undefined;
  const v = (item as { value?: unknown }).value;
  return 'name' in item && v !== null && typeof v === 'object' ? v : item;
}

/** The tables under a top-level element (it included), in document order, as values. */
function tablesIn(item: Element): object[] {
  const out: object[] = [];
  const visit = (items: readonly Element[]): void => {
    for (const el of items) {
      const v = valueOf(el);
      if (!v) continue;
      if (typeNameOf(el) === 'org_docx4j_wml.Tbl') out.push(v);
      const children = childrenOf(v);
      if (children) visit(children);
    }
  };
  visit([item]);
  return out;
}

/** One revision shown in Word only as the engine lists it, with where it is. */
export interface ListedPreserved extends PreservedTarget {
  change: TrackedChange;
  author: string;
}

/**
 * The revisions shown in Word only that the engine lists in a body, each with where it is as
 * `preservedTargetOf` names it: a paragraph mark's formatting out of the group Office JS makes of
 * it with its runs' (core-ts check 27), the rest as listed.
 */
export function listedPreserved(body: Body, container: { content?: Element[]; sectPr?: wml.SectPr }): ListedPreserved[] {
  const content = container.content ?? [];
  const top = new Map<object, number>();
  content.forEach((el, i) => { const v = valueOf(el); if (v) top.set(v, i); });
  const blockOf = (start: object | undefined): number | undefined => {
    let v = start as { PARENT?: object } | undefined;
    for (let depth = 0; v && depth < 10_000; depth++) {
      const i = top.get(v);
      if (i !== undefined) return i;
      v = v.PARENT as { PARENT?: object } | undefined;
    }
    return undefined;
  };
  const pieces: TrackedChange[] = [];
  for (const c of body.getTrackedChanges()) {
    if (c.target.kind === 'group') { for (const p of c.target.pieces) if (p.target.kind === 'markProperties') pieces.push(p); }
    else pieces.push(c);
  }
  const counts = new Map<string, number>();
  // An equation's revisions are one entry in the panel (the equation's), many changes to the engine (each math
  // run's and structure's): every change of an equation takes the equation's ordinal among the block's revised
  // equations of the kind, so one target resolves them all.
  const equations = new Map<string, Map<object, number>>();
  const out: ListedPreserved[] = [];
  for (const change of pieces) {
    const t = change.target;
    let kind: PreservedTargetKind;
    let block: number | 'final' | undefined;
    let equation: object | undefined;
    switch (t.kind) {
      case 'tableProperties': kind = 'tableProperties'; block = blockOf(valueOf(t.tbl)); break;
      case 'cell': kind = t.cell === 'ins' ? 'cellIns' : 'cellDel'; block = blockOf(valueOf(t.tc)); break;
      case 'sectionProperties': kind = 'sectPrChange'; block = t.sectPr === container.sectPr ? 'final' : blockOf(change.paragraph?.p ?? t.sectPr); break;
      case 'markProperties': kind = 'markRPrChange'; block = blockOf(change.paragraph?.p); break;
      case 'holder': kind = t.revision === 'ins' ? 'holderIns' : 'holderDel'; block = blockOf(valueOf(t.holder)); break;
      case 'mathRun': case 'mathStructure': {
        kind = t.revision === 'ins' ? 'mathIns' : 'mathDel';
        equation = valueOf(t.equation) ?? valueOf(t.kind === 'mathRun' ? t.run : t.structure) ?? undefined;
        block = blockOf(equation);
        break;
      }
      default: continue;
    }
    if (block === undefined) continue;
    let ordinal: number;
    if (t.kind === 'tableProperties' && block !== 'final') ordinal = tablesIn(content[block]!).indexOf(valueOf(t.tbl)!);
    else if (equation !== undefined) {
      const key = `${block}:${kind}`;
      const seen = equations.get(key) ?? equations.set(key, new Map()).get(key)!;
      if (!seen.has(equation)) seen.set(equation, seen.size);
      ordinal = seen.get(equation)!;
    } else {
      const key = `${block}:${kind}`;
      ordinal = counts.get(key) ?? 0;
      counts.set(key, ordinal + 1);
    }
    // No `w:author` reads as undefined, whatever the type says: `''`, the unknown author a caller names (ED-005 section 12.113).
    out.push({ change, author: (change.author as string | undefined) ?? '', block, kind, ordinal });
  }
  return out;
}

/**
 * Resolve, on a tree run's body, one entry's revision (`target`) or every listed one of an
 * author's, last first as the engine's "all" does. The number resolved.
 */
export function resolvePreservedOnTree(body: Body, container: { content?: Element[]; sectPr?: wml.SectPr }, action: 'accept' | 'reject', which: { target: PreservedTarget } | { author: string }): number {
  const picked = listedPreserved(body, container).filter((l) => ('target' in which
    ? l.block === which.target.block && l.kind === which.target.kind && l.ordinal === which.target.ordinal
    : l.author === which.author));
  for (let i = picked.length - 1; i >= 0; i--) picked[i]!.change[action]();
  return picked.length;
}
