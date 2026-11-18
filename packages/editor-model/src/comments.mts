// SPDX-License-Identifier: GPL-3.0-or-later
//
// Comments edited (ED-005 section 7): what a comment is made of, as Word writes one, and where
// the body anchors it. A comment's anchor is its `comment` marks and its reference atom (the
// `commentReference` run, an opaque atom); a comment with an empty range has the atom alone
// (section 7.3). The anchors are read per top-level block and kept while the block's node lives,
// as the margin's ranges are (ED-003 section 8.2), since the listing asks on every change.
import type { Node as PMNode } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';
import { p, runOf, linkParents, wordCalendarOf, type Element } from '@docx4j/core-ts';
import * as el from '@docx4j/generated-objects-ts/modules/org_docx4j_wml.el';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { SourceIds } from './ids.mjs';
import type { ParagraphInfo } from './import.mjs';
import { runItemsFor } from './excluded.mjs';

/** The style Word gives a comment's paragraphs, and the character style of its reference runs (the engine's `COMMENT_TEXT_STYLE`, `COMMENT_REFERENCE_STYLE`). */
export const COMMENT_STYLES = { text: 'CommentText', reference: 'CommentReference' } as const;

/**
 * A comment's paragraphs for its text, a paragraph per line, as Word and the engine write them:
 * the `CommentText` style, the first opened by the annotation reference run.
 */
export function commentParagraphs(text: string): Element<wml.P>[] {
  return text.split('\n').map((line, i) => {
    const runs: Element[] = [];
    if (i === 0) runs.push(runOf([el.annotationRef({}) as Element], { rStyle: { val: COMMENT_STYLES.reference } }));
    if (line !== '') runs.push(runOf(runItemsFor(line)));
    return p(runs, { style: COMMENT_STYLES.text }) as Element<wml.P>;
  });
}

/**
 * A `w:comment` for a comment made in the editor, what its story is projected from and what a peer
 * adopts; the engine writes the entry itself at save (`addCommentEntry`), with the `w14:paraId` and
 * the date as CR-002 section 29 has it (ED-005 section 7.4).
 */
export function commentElement(id: number, author: string, initials: string, date: Date, text: string): wml.Comments.Comment {
  const element: wml.Comments.Comment = {
    TYPE_NAME: 'org_docx4j_wml.Comments.Comment',
    // `w:date` as Word writes it: the local wall clock, marked Z (core-ts 0.3.0, CR-002 section 29).
    id, author, initials, date: wordCalendarOf(date),
    content: commentParagraphs(text) as wml.Comments.Comment['content'],
  };
  linkParents(element, undefined);
  return element;
}

/** The initials Word shows for a name: the first letter of each word, at most three, as the engine derives them. */
export function initialsOf(name: string): string {
  return name.split(/\s+/).filter((w) => w.length > 0).slice(0, 3).map((w) => w[0]!.toUpperCase()).join('');
}

/** The comment id a `commentReference` atom stands for, read through the registry that resolves its element. */
export function referenceIdOf(node: PMNode, ids: SourceIds): string | undefined {
  if (node.type.name !== 'opaqueInline' || node.attrs['kind'] !== 'commentReference') return undefined;
  const element = ids.get(node.attrs['sourceId'] as string);
  const id = (element?.value as { id?: unknown } | undefined)?.id;
  return id === undefined || id === null ? undefined : String(id);
}

const blockAnchors = new WeakMap<PMNode, Set<string>>();
const docAnchors = new WeakMap<PMNode, Set<string>>();

/** The ids of the comments a block anchors: a `comment` mark, or a reference atom. Memoised per block. */
function commentAnchorsOf(block: PMNode, ids: SourceIds): Set<string> {
  let within = blockAnchors.get(block);
  if (!within) {
    const found = new Set<string>();
    block.descendants((node) => {
      for (const m of node.marks) if (m.type.name === 'comment') found.add(String(m.attrs['id']));
      const ref = referenceIdOf(node, ids);
      if (ref !== undefined) found.add(ref);
      return true;
    });
    within = found;
    blockAnchors.set(block, within);
  }
  return within;
}

/** The ids of the comments a document anchors: a `comment` mark, or a reference atom. Memoised per block, then per document. */
export function commentAnchors(doc: PMNode, ids: SourceIds): Set<string> {
  let out = docAnchors.get(doc);
  if (out) return out;
  out = new Set();
  doc.forEach((block) => { for (const id of commentAnchorsOf(block, ids)) out!.add(id); });
  docAnchors.set(doc, out);
  return out;
}

const blockTextComments = new WeakMap<PMNode, Set<string>>();

/** The comments whose mark some text of a block carries (a reference alone is not text). Memoised per block. */
function textCommentsOf(block: PMNode): Set<string> {
  let out = blockTextComments.get(block);
  if (out) return out;
  const found = new Set<string>();
  block.descendants((node) => {
    if (node.isText) for (const m of node.marks) if (m.type.name === 'comment') found.add(String(m.attrs['id']));
    return true;
  });
  blockTextComments.set(block, found);
  return found;
}

/** Whether the file put a comment's start marker outside any tracked insertion or deletion. */
export function startedOutsideRevision(paragraphs: ReadonlyMap<string, ParagraphInfo>, id: string): boolean {
  for (const info of paragraphs.values()) if (info.commentStarts.has(id)) return !info.startsInRevision?.has(id);
  return false;
}

/**
 * Word check 32 (ED-005 section 12.49): a resolution that takes away all the text a comment covered
 * (an insertion rejected, a deletion accepted, the one change or all of them) takes the comment with
 * it, when the comment's start marker went with the revision. Word removes its markers, its
 * reference run, plain and outside the revision as it is, and its entries. Where the file put the
 * start marker outside the revision (the form core-ts 0.3.0 wrote), Word 2010 and Word 15 kept the
 * comment as its reference alone (their Reject All saves of `review-for-word.docx`). A comment made
 * in the editor has its start marker where its first word is, so inside the revision it lies in
 * (`placeNewCommentMarkers`), as Word's own form has it.
 *
 * So in `tr`, after the resolution, the reference of every such comment whose text `before` had and
 * `tr.doc` has none is deleted; the save then drops the comment through the engine's
 * `Comment.delete`, its markers and replies with it (`Stories.exportComments`), and an undo brings it
 * back. A comment that keeps some text stays, and one that had none (a reference alone) is left;
 * neither is measured. Only the blocks the resolution changed are read for what went. `startedOutside`
 * names a file comment whose start marker lay outside any revision. The ids of the comments taken.
 */
export function dropCommentsResolvedAway(before: PMNode, tr: Transaction, ids: SourceIds, startedOutside: (id: string) => boolean = () => false): string[] {
  const now = new Set<PMNode>();
  tr.doc.forEach((block) => { now.add(block); });
  const gone = new Set<string>();
  before.forEach((block) => { if (!now.has(block)) for (const id of textCommentsOf(block)) gone.add(id); });
  if (gone.size === 0) return [];
  tr.doc.forEach((block) => { for (const id of textCommentsOf(block)) gone.delete(id); });
  for (const id of gone) if (startedOutside(id)) gone.delete(id);
  if (gone.size === 0) return [];
  const references: number[] = [];
  tr.doc.forEach((block, offset) => {
    if (![...gone].some((id) => commentAnchorsOf(block, ids).has(id))) return;
    block.descendants((node, pos) => {
      const id = referenceIdOf(node, ids);
      if (id !== undefined && gone.has(id)) references.push(offset + 1 + pos);
      return true;
    });
  });
  for (const at of references.reverse()) tr.delete(at, at + 1);
  return [...gone];
}

