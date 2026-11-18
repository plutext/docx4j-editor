// SPDX-License-Identifier: GPL-3.0-or-later
//
// A tree run's own change, applied to the store as it is when the run ends (ED-004 section 10.2
// item 1). The run exports the store, awaits the script and re-projects; meanwhile another peer
// (or the same user, typing during a macro's await) may have changed the store. Handing the fresh
// projection to `replaceDocument` would diff it against the store as it is now and so revert
// every such change. Instead the run's change is taken block by block (the document before the
// run against the fresh projection, ids ignored) and only those blocks are substituted into the
// current document. The blocks before the run are told apart in the current document by their
// Yjs element, which an edit inside a block keeps. A block both sides changed goes to the run.
import type { Node as PMNode, Schema } from 'prosemirror-model';
import type * as Y from 'yjs';
import { normalizedJson } from './selfcheck.mjs';

/** A block's content with ids and local attributes left out: what "the run changed it" compares. */
export function blockKey(schema: Schema, block: PMNode): string {
  return JSON.stringify(normalizedJson(schema.topNodeType.create(null, [block])));
}

/**
 * Pairs of equal blocks in order (a longest common subsequence), as [index in a, index in b]. The
 * common prefix and suffix are paired first, so the quadratic part covers only the region a run
 * changed. `same` compares two blocks by index; `a` and `b` are their keys, read only past the
 * prefix and suffix.
 */
function align(n0: number, m0: number, same: (i: number, j: number) => boolean, a: (i: number) => string, b: (j: number) => string): [number, number][] {
  let start = 0;
  while (start < n0 && start < m0 && same(start, start)) start++;
  let endA = n0;
  let endB = m0;
  while (endA > start && endB > start && same(endA - 1, endB - 1)) { endA--; endB--; }
  const pairs: [number, number][] = [];
  for (let i = 0; i < start; i++) pairs.push([i, i]);
  const n = endA - start;
  const m = endB - start;
  if (n > 0 && m > 0) {
    // lengths[i][j]: the longest common subsequence of a[start + i..endA) and b[start + j..endB).
    const lengths: Uint32Array[] = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
    for (let i = n - 1; i >= 0; i--) {
      for (let j = m - 1; j >= 0; j--) {
        lengths[i]![j] = a(start + i) === b(start + j) ? lengths[i + 1]![j + 1]! + 1 : Math.max(lengths[i + 1]![j]!, lengths[i]![j + 1]!);
      }
    }
    let i = 0;
    let j = 0;
    while (i < n && j < m) {
      if (a(start + i) === b(start + j)) { pairs.push([start + i, start + j]); i++; j++; }
      else if (lengths[i + 1]![j]! >= lengths[i]![j + 1]!) i++;
      else j++;
    }
  }
  for (let k = 0; k < n0 - endA; k++) pairs.push([endA + k, endB + k]);
  return pairs;
}

export interface RunMerge {
  /** The document to hand to `replaceDocument`. */
  doc: PMNode;
  /** Blocks kept as the store has them, since another change reached them during the run and the run did not touch them. */
  kept: PMNode[];
  /**
   * The blocks the run replaced or removed that another change reached during the run: each
   * before-block's index, and the store's form of it now (none when another change removed it).
   * The merge puts the run's blocks in their place, so the other change is lost there (ED-004
   * section 10.2's rule for a script); an agent's call is refused instead (ED-005 proposal 42).
   */
  conflicts: { before: number; now: PMNode | undefined }[];
}

/**
 * The run's change over the current document. `before` and `beforeItems` are the store's blocks
 * and their Yjs elements when the run started, `fresh` the projection after it, `current` and
 * `currentItems` the store's when it ended. Undefined when the store's blocks and elements do
 * not line up (the caller then replaces the document whole, as before).
 */
export function mergeRun(
  schema: Schema,
  before: PMNode, beforeItems: readonly unknown[],
  fresh: PMNode,
  current: PMNode, currentItems: readonly unknown[],
): RunMerge | undefined {
  if (before.childCount !== beforeItems.length || current.childCount !== currentItems.length) return undefined;
  const beforeBlocks = blocks(before);
  const freshBlocks = blocks(fresh);
  const currentBlocks = blocks(current);
  // Keys are read lazily: most blocks are equal as nodes (ids included), which is cheaper to see.
  const beforeKey = cached((i) => blockKey(schema, beforeBlocks[i]!));
  const freshKey = cached((j) => blockKey(schema, freshBlocks[j]!));
  const same = (i: number, j: number): boolean => beforeBlocks[i]!.eq(freshBlocks[j]!) || beforeKey(i) === freshKey(j);

  // Where each block before the run is now, and the blocks another change added, after the
  // before-block they follow (-1: at the start).
  const origin = new Map<unknown, number>();
  beforeItems.forEach((item, i) => origin.set(item as Y.XmlElement, i));
  const now = new Map<number, number>();
  const added = new Map<number, PMNode[]>();
  let anchor = -1;
  currentItems.forEach((item, c) => {
    const i = origin.get(item);
    if (i !== undefined) { now.set(i, c); anchor = i; return; }
    const list = added.get(anchor) ?? [];
    list.push(currentBlocks[c]!);
    added.set(anchor, list);
  });

  const out: PMNode[] = [...(added.get(-1) ?? [])];
  const kept: PMNode[] = [];
  const conflicts: RunMerge['conflicts'] = [];
  /** Whether another change reached before-block `k` during the run. */
  const changedMeanwhile = (k: number): { now: PMNode | undefined } | undefined => {
    const c = now.get(k);
    if (c === undefined) return { now: undefined };
    const block = currentBlocks[c]!;
    return block.eq(beforeBlocks[k]!) || blockKey(schema, block) === beforeKey(k) ? undefined : { now: block };
  };
  const pairs = align(beforeBlocks.length, freshBlocks.length, same, beforeKey, freshKey);
  let i = 0;
  let j = 0;
  const flushTo = (iEnd: number, jEnd: number): void => {
    // A hunk: the run removed before-blocks i..iEnd and put fresh blocks j..jEnd in their place.
    for (; j < jEnd; j++) out.push(freshBlocks[j]!);
    for (; i < iEnd; i++) {
      const meanwhile = changedMeanwhile(i);
      if (meanwhile) conflicts.push({ before: i, now: meanwhile.now });
      out.push(...(added.get(i) ?? []));
    }
  };
  for (const [pi, pj] of pairs) {
    flushTo(pi, pj);
    // A block the run left as it was: the store's form if another change reached it, else the
    // fresh one (whose ids are what the run's projection gave); gone if another change removed it.
    const c = now.get(pi);
    if (c !== undefined) {
      const block = currentBlocks[c]!;
      if (block.eq(beforeBlocks[pi]!) || blockKey(schema, block) === beforeKey(pi)) out.push(freshBlocks[pj]!);
      else { out.push(block); kept.push(block); }
    }
    out.push(...(added.get(pi) ?? []));
    i = pi + 1;
    j = pj + 1;
  }
  flushTo(beforeBlocks.length, freshBlocks.length);

  // The final section: the run's if it changed it, else the store's.
  const runSect = fresh.attrs['finalSectPr'];
  const sectPr = JSON.stringify(runSect) !== JSON.stringify(before.attrs['finalSectPr']) ? runSect : current.attrs['finalSectPr'];
  if (out.length === 0) out.push(schema.nodes['paragraph']!.create());
  return { doc: schema.topNodeType.create({ ...fresh.attrs, finalSectPr: sectPr }, out), kept, conflicts };
}

function cached(f: (i: number) => string): (i: number) => string {
  const memo = new Map<number, string>();
  return (i) => { let v = memo.get(i); if (v === undefined) { v = f(i); memo.set(i, v); } return v; };
}

function blocks(doc: PMNode): PMNode[] {
  const out: PMNode[] = [];
  doc.forEach((b) => out.push(b));
  return out;
}
