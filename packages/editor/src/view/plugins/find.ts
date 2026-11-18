// SPDX-License-Identifier: GPL-3.0-or-later
//
// The find bar's hits on a surface (ED-005 sections 3.4 and 3.6): the query's `Finder` from the
// model, its hits memoised per top-level block node (a keystroke searches only the blocks it
// changed), the current hit, and the highlights as inline decorations in one set, the current one
// with a class of its own. A query matching everywhere keeps its count exact and draws only the
// thousand hits nearest the current one, since a set of a hundred thousand decorations is the cost
// CLAUDE.md warns of. A keystroke that changes no block's hits maps the set; one that does patches
// the blocks it changed, unless the query runs across paragraph marks (`Finder.across`: Word's
// wildcards, `^p`), when a block's hits depend on its neighbours and the set is drawn again. Every
// surface carries the plugin; the bar sets the query on each.
import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import type { BlockHits, EditorDocument, Finder, Hit } from '@docx4j/editor-model';

export const findKey = new PluginKey<FindState>('docx4j-find');

/** The meta that sets the query (a finder, or null to clear it) or moves the current hit. */
export const FIND_META = 'docx4j-find';

export interface FindMeta {
  /** A new query: its finder, or null to clear the highlights. */
  finder?: Finder | null;
  /** The current hit by its index in document order; -1 for none (the current hit is in another story). */
  current?: number;
  /** With a new query: the current hit is the first at or after this position (else the selection's start). */
  near?: number;
}

export interface FindState {
  finder: Finder | null;
  /** The blocks holding hits, in order. */
  blocks: BlockHits[];
  total: number;
  /** The current hit's index in document order; -1 for none. */
  current: number;
  set: DecorationSet;
  /** The positions between which hits are drawn: everything when the query has fewer than `DRAWN` hits. */
  window: [number, number];
}

/** The most hits drawn at once, around the current one. */
export const DRAWN = 1000;

const EMPTY: FindState = { finder: null, blocks: [], total: 0, current: -1, set: DecorationSet.empty, window: [0, 0] };

/** The query a surface made later (a header at a boundary, a note's view) starts with: the bar's, per session. */
const active = new WeakMap<EditorDocument, Finder>();

/** Remember the bar's query for the session's surfaces made from now on; null forgets it. */
export function setActiveFinder(session: EditorDocument, finder: Finder | null): void {
  if (finder) active.set(session, finder); else active.delete(session);
}

/** The blocks with hits, and the ranges of the blocks searched afresh (the ones a transaction changed). */
function collect(finder: Finder, doc: PMNode): { blocks: BlockHits[]; total: number; changed: [number, number][]; across: boolean } {
  const { blocks, changed, across } = finder.byBlock(doc);
  return { blocks, total: blocks.reduce((n, b) => n + b.hits.length, 0), changed, across };
}

/** The hit of an index, at its position in the document. */
export function hitAt(state: Pick<FindState, 'blocks'>, index: number): Hit | undefined {
  if (index < 0) return undefined;
  let i = index;
  for (const b of state.blocks) {
    if (i < b.hits.length) { const h = b.hits[i]!; return { ...h, from: h.from + b.pos, to: h.to + b.pos }; }
    i -= b.hits.length;
  }
  return undefined;
}

/** The index of the first hit at or after a position; -1 when none follows. */
export function indexAtOrAfter(blocks: readonly BlockHits[], pos: number): number {
  let index = 0;
  for (const b of blocks) {
    if (b.pos + b.size <= pos) { index += b.hits.length; continue; }
    for (const h of b.hits) {
      if (h.from + b.pos >= pos) return index;
      index++;
    }
  }
  return -1;
}

function decoration(hit: Hit, current: boolean): Decoration {
  return Decoration.inline(hit.from, hit.to, { class: current ? 'find-hit current' : 'find-hit' });
}

/** The set drawn whole: every hit, or the `DRAWN` nearest the current one. */
function draw(doc: PMNode, finder: Finder, blocks: BlockHits[], total: number, current: number): FindState {
  const all = total <= DRAWN;
  const centre = Math.max(0, current);
  const first = all ? 0 : Math.max(0, Math.min(centre - DRAWN / 2, total - DRAWN));
  const last = all ? total - 1 : first + DRAWN - 1;
  const decos: Decoration[] = [];
  let index = 0;
  let from = Number.POSITIVE_INFINITY;
  let to = 0;
  for (const b of blocks) {
    if (index + b.hits.length <= first) { index += b.hits.length; continue; }
    if (index > last) break;
    for (const h of b.hits) {
      if (index >= first && index <= last) {
        const at = { ...h, from: h.from + b.pos, to: h.to + b.pos };
        decos.push(decoration(at, index === current));
        from = Math.min(from, at.from);
        to = Math.max(to, at.to);
      }
      index++;
    }
  }
  const window: [number, number] = all ? [0, doc.content.size] : [from === Number.POSITIVE_INFINITY ? 0 : from, to];
  return { finder, blocks, total, current, set: DecorationSet.create(doc, decos), window };
}

/**
 * After a transaction that changed the document: the blocks it changed are searched again (the
 * memo keeps the rest), the current hit follows its text, and the set is mapped, patched over the
 * changed blocks, or drawn again when the current hit or the drawn window moved.
 */
function afterEdit(value: FindState, tr: Transaction, state: EditorState): FindState {
  const finder = value.finder!;
  const { blocks, total, changed, across } = collect(finder, state.doc);
  const old = hitAt(value, value.current);
  let current = -1;
  if (old && total > 0) {
    current = indexAtOrAfter(blocks, tr.mapping.map(old.from, -1));
    if (current < 0) current = 0;
  }
  const next = hitAt({ blocks }, current);
  const sameCurrent = old !== undefined && next !== undefined && tr.mapping.map(old.from, -1) === next.from && tr.mapping.map(old.to, 1) === next.to;
  // A window that drew everything and still does, or a partial one whose count is unchanged, can be patched; a query
  // across paragraph marks is drawn again, since a hit may start in a block the transaction left alone.
  const patchable = !across && ((total <= DRAWN && value.total <= DRAWN) || (total === value.total && total > DRAWN));
  if (!patchable || (!sameCurrent && (old !== undefined || next !== undefined))) return draw(state.doc, finder, blocks, total, current);
  let set = value.set.map(tr.mapping, tr.doc);
  const window: [number, number] = total <= DRAWN ? [0, state.doc.content.size] : [tr.mapping.map(value.window[0], -1), tr.mapping.map(value.window[1], 1)];
  if (changed.length > 0) {
    const add: Decoration[] = [];
    let index = 0;
    for (const b of blocks) {
      const fresh = changed.some(([f]) => f === b.pos);
      for (const h of b.hits) {
        if (fresh) {
          const at = { ...h, from: h.from + b.pos, to: h.to + b.pos };
          if (at.from >= window[0] && at.to <= window[1]) add.push(decoration(at, index === current));
        }
        index++;
      }
    }
    for (const [from, to] of changed) set = set.remove(set.find(from, to));
    set = set.add(tr.doc, add);
  }
  return { finder, blocks, total, current, set, window };
}

export function findPlugin(session: EditorDocument): Plugin<FindState> {
  return new Plugin<FindState>({
    key: findKey,
    state: {
      init: (_config, state) => {
        const finder = active.get(session);
        if (!finder) return EMPTY;
        const { blocks, total } = collect(finder, state.doc);
        return draw(state.doc, finder, blocks, total, -1);
      },
      apply: (tr, value, _old, state) => {
        const meta = tr.getMeta(FIND_META) as FindMeta | undefined;
        if (meta?.finder !== undefined) {
          if (meta.finder === null) return EMPTY;
          const { blocks, total } = collect(meta.finder, state.doc);
          let current = meta.current ?? (total === 0 ? -1 : indexAtOrAfter(blocks, meta.near ?? state.selection.from));
          if (meta.current === undefined && current < 0 && total > 0) current = 0;
          return draw(state.doc, meta.finder, blocks, total, current);
        }
        if (!value.finder) return value;
        if (meta?.current !== undefined) return draw(state.doc, value.finder, value.blocks, value.total, meta.current);
        if (!tr.docChanged) return value;
        return afterEdit(value, tr, state);
      },
    },
    props: {
      decorations: (state) => findKey.getState(state)?.set ?? DecorationSet.empty,
    },
  });
}
