// SPDX-License-Identifier: GPL-3.0-or-later
//
// The `sources` map (ED-004 section 3.7): every element that is not in the base file, in a room
// and out of one, as its XML with the ids of it and every element under it keyed by path. A tree
// run writes an entry for each block it made or changed; a session replays every entry at open
// (a join, a recovery) and registers the ones that arrive later, so a node's `sourceId` always
// resolves to the same element on every peer and in every recovered session.
//
// The path of an object is the property names and array indexes that reach it from the entry's
// root element, `PARENT` pointers skipped: `value/content/2/value/content/0`. The receiver
// unmarshals the entry's XML and assigns each id to the object at its path before projecting, so
// the projection's side tables (run holders, comment ranges, continuation cells, the imported
// nodes) come out with the sender's ids whatever order the projection walks in. The sender puts
// the same round trip through its own tree first (`roundTrip` in the session), so both sides hold
// the objects the XML unmarshals to and the paths agree by construction.
import type { Element } from '@docx4j/core-ts';
import type { SourceIds } from './ids.mjs';

/** One entry: a block's XML and the ids of every element in it that has one, by path. */
export interface SourceEntry {
  xml: string;
  ids: Record<string, string>;
  /** A run (a note's reference), registered as a run source rather than projected as a block. */
  inline?: true;
}

/** Visit every object of a typed tree once, with its path from the root; `PARENT` pointers are not followed. */
function walk(root: unknown, visit: (value: object, path: string) => void): void {
  const seen = new Set<object>();
  const go = (value: unknown, path: string): void => {
    if (value === null || typeof value !== 'object' || seen.has(value)) return;
    seen.add(value);
    visit(value, path);
    if (Array.isArray(value)) {
      value.forEach((v, i) => go(v, path === '' ? String(i) : `${path}/${i}`));
      return;
    }
    for (const [k, v] of Object.entries(value)) {
      if (k === 'PARENT') continue;
      go(v, path === '' ? k : `${path}/${k}`);
    }
  };
  go(root, '');
}

/** The ids of an element and everything under it, by path (the root's path is the empty string). */
export function pathIds(root: Element, ids: SourceIds): Record<string, string> {
  const out: Record<string, string> = {};
  walk(root, (value, path) => {
    const id = ids.existingIdOf(value as Element);
    if (id !== undefined) out[path] = id;
  });
  return out;
}

/** Give the objects of an unmarshalled entry the ids its sender gave them; the paths an entry names that the tree lacks, returned. */
export function assignPathIds(root: Element, entryIds: Record<string, string>, ids: SourceIds): string[] {
  const byPath = new Map<string, object>();
  walk(root, (value, path) => byPath.set(path, value));
  const missing: string[] = [];
  for (const [path, id] of Object.entries(entryIds)) {
    const target = byPath.get(path);
    if (target) ids.assign(target as Element, id); else missing.push(path);
  }
  return missing;
}

/** Drop the ids of an element and everything under it, so the next projection mints new ones. */
export function forgetTree(root: Element, ids: SourceIds): void {
  walk(root, (value) => ids.forget(value as Element));
}

/**
 * Structural equality of two pieces of the typed tree, `PARENT` pointers skipped and undefined
 * properties ignored: whether a tree run changed a block in place (the block compared with its
 * copy from before the run).
 */
export function treeEqual(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (typeof a !== typeof b || a === null || b === null || typeof a !== 'object') return false;
  if (Array.isArray(a) !== Array.isArray(b)) return false;
  if (Array.isArray(a)) {
    const bb = b as unknown[];
    return a.length === bb.length && a.every((v, i) => treeEqual(v, bb[i]));
  }
  const ao = a as Record<string, unknown>;
  const bo = b as Record<string, unknown>;
  const ak = Object.keys(ao).filter((k) => k !== 'PARENT' && ao[k] !== undefined);
  const bk = Object.keys(bo).filter((k) => k !== 'PARENT' && bo[k] !== undefined);
  if (ak.length !== bk.length) return false;
  return ak.every((k) => Object.prototype.hasOwnProperty.call(bo, k) && treeEqual(ao[k], bo[k]));
}
