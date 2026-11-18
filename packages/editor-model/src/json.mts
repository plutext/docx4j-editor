// SPDX-License-Identifier: GPL-3.0-or-later
//
// Plain JSON of a piece of the typed tree, for node attributes that carry an element the editor
// does not model field by field (`sectPr`, `sdtPr`, the paragraph mark's `rPr`): PARENT pointers
// stripped, TYPE_NAME kept so that the value can be put back into the tree and marshalled.

/** A structural copy without PARENT pointers; TYPE_NAME and every attribute survive. */
export function toPlain<T>(value: T): T {
  return JSON.parse(JSON.stringify(value, (key, v) => (key === 'PARENT' ? undefined : v))) as T;
}

/** A fresh tree value from plain JSON (a deep clone, so the attribute's object is never shared with the tree). */
export function revive<T>(plain: T): T {
  return JSON.parse(JSON.stringify(plain)) as T;
}

/** Deep equality of two plain JSON values (key order insensitive). */
export function plainEqual(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (typeof a !== typeof b || a === null || b === null || typeof a !== 'object') return false;
  if (Array.isArray(a) !== Array.isArray(b)) return false;
  if (Array.isArray(a)) {
    const bb = b as unknown[];
    return a.length === bb.length && a.every((v, i) => plainEqual(v, bb[i]));
  }
  const ao = a as Record<string, unknown>;
  const bo = b as Record<string, unknown>;
  const ak = Object.keys(ao).filter((k) => ao[k] !== undefined);
  const bk = Object.keys(bo).filter((k) => bo[k] !== undefined);
  if (ak.length !== bk.length) return false;
  return ak.every((k) => plainEqual(ao[k], bo[k]));
}

/** A shallow copy of a tree value without some keys and without PARENT. */
export function copyWithout<T extends object>(value: T, keys: readonly string[]): T {
  const out: Record<string, unknown> = {};
  for (const k of Object.keys(value)) {
    if (k === 'PARENT' || keys.includes(k)) continue;
    out[k] = (value as Record<string, unknown>)[k];
  }
  return out as T;
}

/**
 * Whether two pieces of the typed tree are the same by value: key order, `TYPE_NAME` tags (which
 * a value built in the editor may lack) and the unmarshaller's extra QName fields aside, as the
 * self-check compares. A formatting or properties change compares its `before` this way
 * (ED-005 section 12.3 item 12).
 */
export function sameValue(a: unknown, b: unknown): boolean {
  return JSON.stringify(canonical(a)) === JSON.stringify(canonical(b));
}

/** A tree value as `sameValue` compares it: names as `{ localPart, namespaceURI }` whatever else a parsed one carries, keys sorted, no `TYPE_NAME`. */
export function canonical(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonical);
  if (typeof value === 'object' && value !== null) {
    const v = value as Record<string, unknown>;
    if (typeof v['localPart'] === 'string' && typeof v['namespaceURI'] === 'string') return { localPart: v['localPart'], namespaceURI: v['namespaceURI'] };
    const out: Record<string, unknown> = {};
    for (const k of Object.keys(v).sort()) if (k !== 'TYPE_NAME' && k !== 'PARENT' && v[k] !== undefined) out[k] = canonical(v[k]);
    return out;
  }
  return value;
}
