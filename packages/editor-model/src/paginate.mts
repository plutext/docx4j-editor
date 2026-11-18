// SPDX-License-Identifier: GPL-3.0-or-later
//
// Re-paginate in a room (ED-004 section 3.8): the companion's layout moves the page-break
// markers alone. Outside a room the answer's body replaces the tree's through a tree run
// (`adoptPaginated`); in a room that would copy the whole document into `sources` on every
// Re-paginate. Here the answer is projected on its own, its paragraphs matched to the store's in
// order (Paginate keeps the structure), and each paragraph whose text agrees gets the answer's
// markers at the answer's offsets. A marker made here has no `sourceId`: export builds
// `<w:lastRenderedPageBreak/>` from its kind, the one opaque kind that can be built that way.
import { Fragment, type Node as PMNode } from 'prosemirror-model';
import { Transform } from 'prosemirror-transform';

export const MARKER_KIND = 'lastRenderedPageBreak';

export function isMarker(node: PMNode): boolean {
  return node.type.name === 'opaqueInline' && node.attrs['kind'] === MARKER_KIND;
}

/** The document's textblocks in order, with their positions. */
function textblocks(doc: PMNode): { node: PMNode; pos: number }[] {
  const out: { node: PMNode; pos: number }[] = [];
  doc.descendants((node, pos) => {
    if (node.isTextblock) { out.push({ node, pos }); return false; }
    return true;
  });
  return out;
}

/** A paragraph's content with its markers left out, as what must agree between the store and the answer. */
function signature(block: PMNode): string {
  const parts: string[] = [];
  block.forEach((n) => { if (!isMarker(n)) parts.push(n.isText ? `t:${n.text}` : n.type.name); });
  return `${block.type.name}|${parts.join('|')}`;
}

/** The offsets of a paragraph's markers, counted over its content without them. */
function markerOffsets(block: PMNode): number[] {
  const out: number[] = [];
  let at = 0;
  block.forEach((n) => { if (isMarker(n)) out.push(at); else at += n.nodeSize; });
  return out;
}

/** The run a marker at this point belongs to: the run of the content before it, else after it. */
function runSourceOf(n: PMNode | undefined): string | null {
  if (!n) return null;
  if (n.isText) return (n.marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'] as string | undefined) ?? null;
  return (n.attrs['runSource'] as string | null | undefined) ?? null;
}

/** The paragraph's content with markers at `offsets`, reusing its own markers where one already stands at an offset. */
function withMarkers(block: PMNode, offsets: readonly number[]): Fragment {
  const schema = block.type.schema;
  const kept = new Map<number, PMNode[]>();
  let at = 0;
  const content: PMNode[] = [];
  block.forEach((n) => {
    if (isMarker(n)) { const list = kept.get(at) ?? []; list.push(n); kept.set(at, list); } else { content.push(n); at += n.nodeSize; }
  });
  const want = [...offsets];
  const out: PMNode[] = [];
  const marker = (offset: number, before: PMNode | undefined, after: PMNode | undefined): PMNode => {
    const own = kept.get(offset);
    if (own && own.length > 0) return own.shift()!;
    return schema.nodes['opaqueInline']!.create({ kind: MARKER_KIND, sourceId: null, runSource: runSourceOf(before) ?? runSourceOf(after) });
  };
  let pos = 0;
  let previous: PMNode | undefined;
  for (const n of content) {
    let rest: PMNode = n;
    let restStart = pos;
    while (want.length > 0 && want[0]! < pos + n.nodeSize) {
      const o = want.shift()!;
      if (o <= restStart) { out.push(marker(o, previous, rest)); continue; }
      // Inside a text node: split it.
      const head = rest.cut(0, o - restStart);
      out.push(head);
      previous = head;
      rest = rest.cut(o - restStart);
      restStart = o;
      out.push(marker(o, previous, rest));
    }
    out.push(rest);
    previous = rest;
    pos += n.nodeSize;
  }
  for (const o of want) out.push(marker(o, previous, undefined));
  return Fragment.from(out);
}

export interface MarkerMove {
  /** The store's document with the answer's markers. */
  doc: PMNode;
  /** Paragraphs whose markers changed. */
  moved: number;
  /** Paragraphs whose text differs from the answer's (edited since the save the companion laid out): their markers left alone. */
  skipped: number;
}

/**
 * The store's document with each paragraph's markers where the answer has them. Undefined when
 * the two have different numbers of paragraphs, which Re-paginate refuses (section 3.8).
 */
export function moveMarkers(store: PMNode, answer: PMNode): MarkerMove | undefined {
  const ours = textblocks(store);
  const theirs = textblocks(answer);
  if (ours.length !== theirs.length) return undefined;
  const tr = new Transform(store);
  let moved = 0;
  let skipped = 0;
  // From the end, so the positions read before any change stay valid.
  for (let i = ours.length - 1; i >= 0; i--) {
    const { node, pos } = ours[i]!;
    const their = theirs[i]!.node;
    if (signature(node) !== signature(their)) { skipped++; continue; }
    const want = markerOffsets(their);
    const have = markerOffsets(node);
    if (want.length === have.length && want.every((o, k) => o === have[k])) continue;
    tr.replaceWith(pos + 1, pos + 1 + node.content.size, withMarkers(node, want));
    moved++;
  }
  return { doc: tr.doc, moved, skipped };
}
