// SPDX-License-Identifier: GPL-3.0-or-later
//
// A place in one version of a document found again in another (ED-011 section 3.5): the caret and the view kept
// across a reload, and carried into the disk version's tab. A place is a paragraph and an offset in its text, so it
// outlives the positions of the version it was read from; the two versions' paragraphs are aligned by `compareText`.
// Both lists are read from the edit layer by one reading, so an index in one is a paragraph the view has. The banner's
// lists are the engine's, which are not these: over the fixtures they differ in 15 of 67 (tables.docx has 193 here, 200 there).

import type { Node as PMNode } from 'prosemirror-model';
import { textOffsetIn } from './bindings.mjs';
import { compareText } from './compare.mjs';
import { paragraphText, type ParagraphText } from './find.mjs';
import { graphemeBoundary } from './text.mjs';

/** A place in a document's text: which of its paragraphs, in document order, and the offset in that paragraph's text. */
export interface TextPlace {
  paragraph: number;
  offset: number;
}

/**
 * Where a place landed, and what became of the text it was in:
 * - `kept`: its paragraph is there, as it was or edited (an empty paragraph has no text to lose, so it is kept where it stood);
 * - `moved`: its paragraph's text stands elsewhere;
 * - `gone`: its text is not there, and the place is where the text stood;
 * - `unrelated`: the versions have nothing in common, and the place is the start.
 */
export interface Landing extends TextPlace {
  found: 'kept' | 'moved' | 'gone' | 'unrelated';
}

/** The shortest text followed to another place: a shorter one repeats by chance ("Notes", "Yes"). */
const MOVED_LENGTH = 12;

/** The length of the pieces two texts are compared by: long enough that unrelated words in one language rarely share them. */
const WINDOW = 4;

function paragraphsOf(doc: PMNode): { node: PMNode; pos: number }[] {
  const out: { node: PMNode; pos: number }[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== 'paragraph') return true;
    out.push({ node, pos });
    return false;
  });
  return out;
}

/** The text of each of a document's paragraphs, in document order, table cells and controls included. */
export function paragraphStrings(doc: PMNode): string[] {
  return paragraphsOf(doc).map((p) => paragraphText(p.node, 0).text);
}

/** The place of a position: the paragraph holding it, else the next after it, else the end of the last. */
export function textPlaceOf(doc: PMNode, pos: number): TextPlace {
  const list = paragraphsOf(doc);
  const index = list.findIndex((p) => pos < p.pos + p.node.nodeSize);
  if (index < 0) {
    const last = list[list.length - 1];
    return last ? { paragraph: list.length - 1, offset: paragraphText(last.node, 0).text.length } : { paragraph: 0, offset: 0 };
  }
  const { node, pos: at } = list[index]!;
  return { paragraph: index, offset: pos <= at ? 0 : textOffsetIn(node, pos - at - 1) };
}

/** The position of an offset: the character it points at, or after the last one. An atom of text is never entered. */
function positionIn({ pieces }: ParagraphText, base: number, offset: number): number {
  let after = base;
  for (const p of pieces) {
    if (p.start === p.end) continue;
    if (offset < p.end) return p.text ? p.pos + (offset - p.start) : p.pos;
    after = p.pos + p.size;
  }
  return after;
}

/** The position of a place, the inverse of `textPlaceOf`; a paragraph past the last is the document's end. */
export function positionOfPlace(doc: PMNode, place: TextPlace): number {
  const target = paragraphsOf(doc)[place.paragraph];
  if (!target) return doc.content.size;
  const base = target.pos + 1;
  return positionIn(paragraphText(target.node, base), base, place.offset);
}

const hasText = (text: string): boolean => text.trim() !== '';

/** Whether at least half of a text survives in another, read in pieces, so that two edits far apart leave it standing. */
function survives(text: string, other: string): boolean {
  if (text.length <= WINDOW) return other.includes(text);
  const there = new Set<string>();
  for (let i = 0; i + WINDOW <= other.length; i++) there.add(other.slice(i, i + WINDOW));
  const pieces = text.length - WINDOW + 1;
  let kept = 0;
  for (let i = 0; i < pieces; i++) if (there.has(text.slice(i, i + WINDOW))) kept++;
  return kept * 2 >= pieces;
}

/** An offset carried across an edit: in the text before or after the edit it stays with that text; inside it, the edit's start. */
function offsetAcross(text: string, edited: string, offset: number): number {
  const shorter = Math.min(text.length, edited.length);
  let prefix = 0;
  while (prefix < shorter && text[prefix] === edited[prefix]) prefix++;
  let suffix = 0;
  while (suffix < shorter - prefix && text[text.length - 1 - suffix] === edited[edited.length - 1 - suffix]) suffix++;
  const carried = offset <= prefix ? offset : offset >= text.length - suffix ? edited.length - (text.length - offset) : prefix;
  return graphemeBoundary(edited, carried);
}

/**
 * Finds places of one version in another, given each version's `paragraphStrings`. The alignment is made once for
 * the places asked after. A paragraph the alignment calls `changed` only stands where the other stood, so whether it
 * is the same paragraph edited is read from the texts.
 */
export function relocator(before: readonly string[], after: readonly string[]): (place: TextPlace) => Landing {
  const { pairs } = compareText(before, after);
  const pairOf = new Map<number, number>();
  const matched = new Set<number>();
  pairs.forEach((p, k) => {
    if (p.a !== undefined) pairOf.set(p.a, k);
    if (p.kind === 'equal') matched.add(p.b!);
  });
  // A blank paragraph matches anywhere, so it says nothing of whether the versions are related.
  let related: boolean | undefined;
  const isRelated = (): boolean => related ??= !before.some(hasText) || pairs.some((p) => p.a !== undefined && p.b !== undefined && hasText(before[p.a]!) && (p.kind === 'equal' || survives(before[p.a]!, after[p.b]!)));
  return ({ paragraph, offset }) => {
    const k = pairOf.get(paragraph);
    if (k === undefined || after.length === 0) return { paragraph: 0, offset: 0, found: 'kept' };
    const pair = pairs[k]!;
    if (pair.kind === 'equal') return { paragraph: pair.b!, offset, found: 'kept' };
    const text = before[paragraph]!;
    if (text.trim().length >= MOVED_LENGTH) {
      const there = after.flatMap((s, j) => (s === text && !matched.has(j) ? [j] : []));
      if (there.length === 1) return { paragraph: there[0]!, offset, found: 'moved' };
    }
    if (pair.kind === 'changed' && hasText(text) && survives(text, after[pair.b!]!)) return { paragraph: pair.b!, offset: offsetAcross(text, after[pair.b!]!, offset), found: 'kept' };
    if (!isRelated()) return { paragraph: 0, offset: 0, found: 'unrelated' };
    const found = hasText(text) ? 'gone' : 'kept';
    if (pair.kind === 'changed') return { paragraph: pair.b!, offset: 0, found };
    const next = pairs.slice(k + 1).find((p) => p.b !== undefined);
    return next ? { paragraph: next.b!, offset: 0, found } : { paragraph: after.length - 1, offset: after[after.length - 1]!.length, found };
  };
}
