// SPDX-License-Identifier: GPL-3.0-or-later
//
// A bookmark keeps both its markers (ED-005 section 12.115). A bookmark is two atoms in the edit layer, a
// `w:bookmarkStart` and a `w:bookmarkEnd` paired by `w:id`, both invisible, so a selection can take one and not
// the other: Shift+Down from the start of the paragraph before a heading ends after the heading's leading
// starts, and a Delete then took the starts and left the ends. Such a file names a bookmark it does not open:
// docx4j's TOC update refuses it and a TOC entry's link to it is broken. The save's self-check cannot see it,
// since the edit layer and the file agree.
//
// The rule, applied to what a local transaction removes after the tracker has run: a marker removed while its
// partner stays goes back where the removal happened, a start after it (before what is left of the bookmark),
// an end before it, so the bookmark loses the text deleted and nothing else. A bookmark both of whose markers
// went, as when its text is deleted whole with them, goes. With tracking on the tracker keeps what it deletes,
// marked, so the rule acts when the deletion is accepted (or an insertion holding one rejected); a marker put
// back is never a revision.

import { Fragment, type Node as PMNode } from 'prosemirror-model';
import { Selection, type Transaction } from 'prosemirror-state';
import { insertPoint } from 'prosemirror-transform';

const START = 'bookmarkStart';
const END = 'bookmarkEnd';

/** A bookmark's marker, inline or between blocks. */
function isMarker(node: PMNode): boolean {
  const kind = node.attrs['kind'];
  return (kind === START || kind === END) && (node.type.name === 'opaqueInline' || node.type.name === 'opaqueBlock');
}

interface Removed {
  node: PMNode;
  sourceId: string;
  start: boolean;
  /** The bookmark's `w:id`; undefined for a marker whose element has none. */
  id: string | undefined;
  /** The step that removed it, and the range it removed, in that step's document. */
  step: number;
  from: number;
  to: number;
  /** Where it stood in the transaction's first document, and among markers removed there, for markers put back at one place. */
  origin: number;
  order: number;
}

/**
 * Put back each bookmark marker the transaction removes while its partner stays (see above). `idOf` is the
 * bookmark's `w:id` for a marker's source id. Nothing is done, and nothing walked beyond the removed ranges,
 * unless a step removes a marker.
 */
export function keepBookmarkMarkers(tr: Transaction, idOf: (sourceId: string) => string | undefined): number {
  const removed = new Map<string, Removed>();
  let order = 0;
  tr.steps.forEach((step, i) => {
    const doc = tr.docs[i]!;
    step.getMap().forEach((from, to) => {
      if (to <= from) return;
      doc.nodesBetween(from, to, (node, pos) => {
        if (!isMarker(node)) return true;
        const sourceId = node.attrs['sourceId'];
        // The last removal wins: a marker moved and then removed again goes back where it was last.
        if (pos >= from && pos + node.nodeSize <= to && typeof sourceId === 'string') {
          removed.set(sourceId, { node, sourceId, start: node.attrs['kind'] === START, id: idOf(sourceId), step: i, from, to, origin: originOf(tr, i, pos), order: order++ });
        }
        return false;
      });
    });
  });
  if (removed.size === 0) return 0;
  // What a step removed may stand in the document again: the tracker puts a deletion back, a move inserts it elsewhere.
  const present = new Set<string>();
  tr.doc.descendants((node) => {
    if (isMarker(node) && typeof node.attrs['sourceId'] === 'string') present.add(node.attrs['sourceId']);
    return true;
  });
  const lost = [...removed.values()].filter((r) => !present.has(r.sourceId));
  const gone = new Set(lost.map((r) => `${r.start ? START : END}:${r.id}`));
  const kept = lost.filter((r) => r.id === undefined || !gone.has(`${r.start ? END : START}:${r.id}`));
  if (kept.length === 0) return 0;
  const placed = kept
    .map((r) => ({ r, at: placeFor(tr, r) }))
    .filter((p): p is { r: Removed; at: number } => p.at !== undefined)
    // From the end of the document, so a place still stands when the ones after it are filled; at one place an
    // end before a start, so two bookmarks that met there do not overlap, and each kind in document order (the
    // steps that removed them may have run from the end, as a resolution does).
    .sort((a, b) => b.at - a.at || Number(a.r.start) - Number(b.r.start) || a.r.origin - b.r.origin || a.r.order - b.r.order);
  for (let i = 0; i < placed.length;) {
    const at = placed[i]!.at;
    const group: PMNode[] = [];
    for (; i < placed.length && placed[i]!.at === at; i++) group.push(putBack(placed[i]!.r.node));
    tr.insert(at, Fragment.from(group));
  }
  return placed.length;
}

/**
 * A position in step `step`'s document mapped back to the transaction's first, one map at a time: `Mapping.invert()`
 * inverts every map of the mapping, a slice's included.
 */
function originOf(tr: Transaction, step: number, pos: number): number {
  // UPSTREAM-26: not `tr.mapping.slice(0, step).invert()`.
  for (let j = step - 1; j >= 0; j--) pos = tr.mapping.maps[j]!.invert().map(pos);
  return pos;
}

/** The marker as it stood, its revision gone: what removed it was a deletion accepted or an insertion rejected, or no revision at all. */
function putBack(node: PMNode): PMNode {
  if (node.attrs['revision'] === undefined || node.attrs['revision'] === null) return node;
  return node.type.create({ ...node.attrs, revision: null }, null, node.marks);
}

/**
 * Where a removed marker goes back: where it was removed, in the transaction's document, a start after what
 * replaced it and an end before. An inline marker goes into the nearest paragraph on its side, a marker
 * between blocks where a block can stand. Undefined when there is no such place.
 */
function placeFor(tr: Transaction, r: Removed): number | undefined {
  const doc = tr.doc;
  // The removed range's own edges, since a position on an edge maps to that edge whatever its side.
  const pos = r.start ? tr.mapping.slice(r.step).map(r.to, 1) : tr.mapping.slice(r.step).map(r.from, -1);
  if (r.node.isBlock) return insertPoint(doc, pos, r.node.type) ?? undefined;
  const $pos = doc.resolve(pos);
  if ($pos.parent.inlineContent && $pos.parent.canReplaceWith($pos.index(), $pos.index(), r.node.type)) return pos;
  const near = Selection.findFrom($pos, r.start ? 1 : -1, true) ?? Selection.findFrom($pos, r.start ? -1 : 1, true);
  return near && near.$from.parent.inlineContent ? near.from : undefined;
}
