// SPDX-License-Identifier: GPL-3.0-or-later
//
// A comment stays one range (ED-005 section 12.116). In the file a comment is a start and an end marker, so it holds
// everything between them; in the edit layer it is a mark, which a paste does not take: text pasted inside a comment
// came in plain, the comment became two pieces with that text between, and the save wrote one range over all of it.
// On reload the pasted text was in the comment, and the self-check refused every save. Typed text takes the marks
// around it already. So what a local transaction inserts between two pieces of one comment takes that comment, as
// in Word, where text put inside a commented range is in the range.

import type { Mark, Node as PMNode } from 'prosemirror-model';
import { Selection, type Transaction } from 'prosemirror-state';

/** How many paragraph edges the search for a neighbour crosses: a comment spanning empty paragraphs is still found. */
const HOPS = 8;

/** Give what the transaction inserts inside a comment that comment's mark. Nothing is walked unless a step inserts. */
export function keepCommentsWhole(tr: Transaction): void {
  const ranges: [number, number][] = [];
  tr.steps.forEach((step, i) => {
    const rest = tr.mapping.slice(i + 1);
    step.getMap().forEach((_oldStart, _oldEnd, newStart, newEnd) => {
      if (newEnd <= newStart) return;
      const from = rest.map(newStart, 1);
      const to = rest.map(newEnd, -1);
      if (to > from) ranges.push([from, to]);
    });
  });
  for (const [from, to] of ranges) {
    const before = commentsBeside(tr.doc, from, -1);
    if (before.length === 0) continue;
    const after = commentsBeside(tr.doc, to, 1);
    for (const mark of before) {
      if (!after.some((m) => m.eq(mark))) continue;
      // Already whole: typed text took the mark itself.
      if (tr.doc.rangeHasMark(from, to, mark.type) && covered(tr.doc, from, to, mark)) continue;
      tr.addMark(from, to, mark);
    }
  }
}

/** Whether every text node in the range carries the mark. */
function covered(doc: PMNode, from: number, to: number, mark: Mark): boolean {
  let all = true;
  doc.nodesBetween(from, to, (node) => {
    if (node.isText && !mark.isInSet(node.marks)) all = false;
    return all;
  });
  return all;
}

/**
 * The comment marks of the nearest text on one side of a position, over the invisible atoms and paragraph edges
 * between (a comment's text can end one paragraph and go on in the next). None at a block that is not a paragraph.
 */
function commentsBeside(doc: PMNode, pos: number, dir: -1 | 1): readonly Mark[] {
  let at = pos;
  for (let hops = 0; hops <= HOPS;) {
    const $at = doc.resolve(at);
    const node = dir < 0 ? $at.nodeBefore : $at.nodeAfter;
    if (node) {
      if (node.isText) return node.marks.filter((m) => m.type.name === 'comment');
      if (!node.isInline) return [];
      at += dir * node.nodeSize;
      continue;
    }
    if (!$at.parent.isTextblock || $at.depth === 0) return [];
    const near = Selection.findFrom(doc.resolve(dir < 0 ? $at.before() : $at.after()), dir, true);
    if (!near) return [];
    at = near.from;
    hops++;
  }
  return [];
}
