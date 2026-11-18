// SPDX-License-Identifier: GPL-3.0-or-later
//
// A document's text against another's (ED-011 section 3.4): the paragraph strings of both read through the
// engine's text model, one definition on both sides (hidden text and placeholders out, deleted text out in the
// accepted view, a kept `mc:AlternateContent` read as no text), aligned by a sequence diff over the strings
// (Myers' algorithm in its linear-space form, as a line diff uses it), and counted. The counts are the
// disk-offer banner's; the aligned pairs are phase 2's starting point.

import { WordprocessingMLPackage, Body, type RepairMode } from '@docx4j/core-ts';
import type { EditorDocument } from './session.mjs';

/** One row of the alignment: an index into each list where the row has one there. */
export interface AlignedPair {
  kind: 'equal' | 'changed' | 'added' | 'removed';
  a?: number;
  b?: number;
}

export interface TextComparison {
  pairs: AlignedPair[];
  /** Paragraphs with a counterpart whose text differs. */
  changed: number;
  /** Paragraphs in `b` with no counterpart in `a`. */
  added: number;
  /** Paragraphs in `a` with no counterpart in `b`. */
  removed: number;
}

/**
 * The alignment of two lists of paragraph strings and its counts. Within a run of differences between
 * two matches, the deleted and the inserted strings are paired in order as "changed", a line diff's
 * display of a replaced block; those left over are removed or added. A paragraph moved elsewhere is one
 * removed and one added.
 */
export function compareText(a: readonly string[], b: readonly string[]): TextComparison {
  // A string the other list does not have can match nothing, so the search runs over the rest, as a line diff
  // discards such lines: its cost grows with the differences, and a rewritten file would take seconds on the page's
  // thread. What is left out is deleted or inserted in the walk below.
  const inA = new Set(a);
  const inB = new Set(b);
  const keptA = a.flatMap((s, i) => (inB.has(s) ? [i] : []));
  const keptB = b.flatMap((s, j) => (inA.has(s) ? [j] : []));
  const ops: Op[] = [];
  diff(keptA.map((i) => a[i]!), 0, keptA.length, keptB.map((j) => b[j]!), 0, keptB.length, ops);
  const pairs: AlignedPair[] = [];
  let changed = 0;
  let added = 0;
  let removed = 0;
  let nextA = 0;
  let nextB = 0;
  /** The run of differences up to these indices. */
  const flush = (toA: number, toB: number): void => {
    const deleted = toA - nextA;
    const inserted = toB - nextB;
    const paired = Math.min(deleted, inserted);
    for (let i = 0; i < paired; i++) pairs.push({ kind: 'changed', a: nextA + i, b: nextB + i });
    for (let i = paired; i < deleted; i++) pairs.push({ kind: 'removed', a: nextA + i });
    for (let i = paired; i < inserted; i++) pairs.push({ kind: 'added', b: nextB + i });
    changed += paired;
    removed += deleted - paired;
    added += inserted - paired;
    nextA = toA;
    nextB = toB;
  };
  for (const op of ops) {
    if (op.kind !== 'equal') continue;
    const x = keptA[op.a]!;
    const y = keptB[op.b]!;
    flush(x, y);
    pairs.push({ kind: 'equal', a: x, b: y });
    nextA = x + 1;
    nextB = y + 1;
  }
  flush(a.length, b.length);
  return { pairs, changed, added, removed };
}

type Op = { kind: 'equal'; a: number; b: number } | { kind: 'delete'; a: number } | { kind: 'insert'; b: number };

/** The diff of `a[aLo, aHi)` against `b[bLo, bHi)` appended to `out`, by recursion on the middle snake. */
function diff(a: readonly string[], aLo: number, aHi: number, b: readonly string[], bLo: number, bHi: number, out: Op[]): void {
  while (aLo < aHi && bLo < bHi && a[aLo] === b[bLo]) { out.push({ kind: 'equal', a: aLo, b: bLo }); aLo++; bLo++; }
  let suffix = 0;
  while (aLo < aHi && bLo < bHi && a[aHi - 1] === b[bHi - 1]) { aHi--; bHi--; suffix++; }
  if (aLo === aHi) for (let j = bLo; j < bHi; j++) out.push({ kind: 'insert', b: j });
  else if (bLo === bHi) for (let i = aLo; i < aHi; i++) out.push({ kind: 'delete', a: i });
  else {
    const [x, y, u, v] = middleSnake(a, aLo, aHi, b, bLo, bHi);
    diff(a, aLo, x, b, bLo, y, out);
    for (let i = 0; i < u - x; i++) out.push({ kind: 'equal', a: x + i, b: y + i });
    diff(a, u, aHi, b, v, bHi, out);
  }
  for (let i = 0; i < suffix; i++) out.push({ kind: 'equal', a: aHi + i, b: bHi + i });
}

/**
 * Myers' middle snake over the ranges (absolute indices in): the forward search from the start and the
 * reverse search from the end meet on a diagonal, and the snake is the run of equal strings there, possibly
 * empty. Linear space: the two frontier arrays alone.
 */
function middleSnake(a: readonly string[], aLo: number, aHi: number, b: readonly string[], bLo: number, bHi: number): [number, number, number, number] {
  const n = aHi - aLo;
  const m = bHi - bLo;
  const delta = n - m;
  const odd = (delta & 1) !== 0;
  const max = Math.ceil((n + m) / 2) + 1;
  const offset = max + 1;
  const vf = new Int32Array(2 * offset + 1);
  const vb = new Int32Array(2 * offset + 1);
  vf[offset + 1] = 0;
  vb[offset + 1] = 0;
  for (let d = 0; d <= max; d++) {
    for (let k = -d; k <= d; k += 2) {
      let x = k === -d || (k !== d && vf[offset + k - 1]! < vf[offset + k + 1]!) ? vf[offset + k + 1]! : vf[offset + k - 1]! + 1;
      let y = x - k;
      const xs = x;
      const ys = y;
      while (x < n && y < m && a[aLo + x] === b[bLo + y]) { x++; y++; }
      vf[offset + k] = x;
      const kr = delta - k;
      if (odd && kr >= -(d - 1) && kr <= d - 1 && x + vb[offset + kr]! >= n) return [aLo + xs, bLo + ys, aLo + x, bLo + y];
    }
    for (let k = -d; k <= d; k += 2) {
      let x = k === -d || (k !== d && vb[offset + k - 1]! < vb[offset + k + 1]!) ? vb[offset + k + 1]! : vb[offset + k - 1]! + 1;
      let y = x - k;
      const xs = x;
      const ys = y;
      while (x < n && y < m && a[aHi - 1 - x] === b[bHi - 1 - y]) { x++; y++; }
      vb[offset + k] = x;
      const kf = delta - k;
      if (!odd && kf >= -d && kf <= d && vf[offset + kf]! + x >= n) return [aHi - x, bHi - y, aHi - xs, bHi - ys];
    }
  }
  // Unreachable: the searches meet by d = ceil((n + m) / 2).
  throw new Error('compareText: the middle snake was not found');
}

/** The paragraph strings of a body as the engine's text model reads them, the accepted view. */
export function paragraphTextsOf(body: Body): string[] {
  return body.paragraphs.map((p) => p.text);
}

/**
 * The paragraph strings of a file's body, read from its bytes through the engine and never projected: the
 * disk version's side (ED-011 section 3.4). The package is loaded for the read and dropped.
 */
export async function paragraphTextsOfBytes(bytes: Uint8Array, repair: RepairMode = 'refuse'): Promise<string[]> {
  const pkg = await WordprocessingMLPackage.load(bytes, { repair });
  const main = pkg.getMainDocumentPart();
  const document = await main.readContents();
  document.body ??= { TYPE_NAME: 'org_docx4j_wml.Body' };
  // The resolver, since a hidden run may be hidden by its style; `getBody()` builds it the same way.
  await pkg.getPropertyResolver();
  return paragraphTextsOf(new Body(undefined, document.body, 'body', pkg));
}

/**
 * The live document's paragraph strings, read through `readTree` over a copy as the companion's dry run
 * reads them (ED-011 section 3.4): the edit layer exported into the private tree, nothing written.
 */
export function paragraphTextsOfSession(session: EditorDocument): Promise<string[]> {
  return session.readTree(({ body }) => paragraphTextsOf(body), { copy: true });
}

/** The live document against a file's bytes: the two reads, then `compareText`. */
export async function compareWithBytes(session: EditorDocument, bytes: Uint8Array): Promise<TextComparison> {
  // The disk version read as the session read its file, through Word's repair where it took one.
  const [live, disk] = await Promise.all([paragraphTextsOfSession(session), paragraphTextsOfBytes(bytes, session.pkg.loadOptions.repair ?? 'refuse')]);
  return compareText(live, disk);
}
