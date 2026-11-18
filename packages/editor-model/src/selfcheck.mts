// SPDX-License-Identifier: GPL-3.0-or-later
//
// The round-trip self-check ED-001 section 11 item 1 runs before every save: export to bytes,
// reload, re-project, compare with the live projection. Session-local things are stripped before
// the comparison: `sourceId`s (fresh per import) and `runSource` marks (a rebuilt paragraph merges
// runs the live document still holds apart); adjacent text with equal marks is merged.
import type { Node as PMNode, Schema } from 'prosemirror-model';
import { canonical } from './json.mjs';

interface JsonNode {
  type: string;
  attrs?: Record<string, unknown>;
  content?: JsonNode[];
  marks?: { type: string; attrs?: Record<string, unknown> }[];
  text?: string;
}

/** The document as JSON with everything session-local removed. */
export function normalizedJson(doc: PMNode): JsonNode {
  return normalize(doc.toJSON() as JsonNode);
}

/**
 * Session-local attributes: ids assigned per import, and an atom's `fmt`, which a reload folds
 * back into its run's properties (the run itself is compared through the text marks around it).
 * A table's `tblGrid` is the grid as imported; export derives the written grid from the cells
 * once a column came or went, so it is compared through the cells' widths, not the attribute.
 */
const LOCAL_ATTRS = new Set(['sourceId', 'copiedFrom', 'runSource', 'alt', 'fmt', 'merged', 'tblGrid']);

/**
 * JSON with object keys sorted at every level, so key order (which unmarshalling does not fix)
 * cannot differ, and without TYPE_NAME tags, which a value built in the editor may lack and a
 * reload always carries.
 */
function sortKeys(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(sortKeys);
  if (typeof value === 'object' && value !== null) {
    const v = value as Record<string, unknown>;
    // An element name: the unmarshaller's QName carries `prefix`, `key` and `string` beside the
    // two parts; one built in the editor carries the two parts. Only the two parts name it.
    if (typeof v['localPart'] === 'string' && typeof v['namespaceURI'] === 'string') return { localPart: v['localPart'], namespaceURI: v['namespaceURI'] };
    const out: Record<string, unknown> = {};
    for (const k of Object.keys(v).sort()) if (k !== 'TYPE_NAME') out[k] = sortKeys(v[k]);
    return out;
  }
  return value;
}

/** Content control group ids are source ids: session-local; they become ordinals in order of first appearance. */
let groupOrdinals = new Map<string, number>();

function normalizeSdt(value: unknown): unknown {
  if (!Array.isArray(value)) return value;
  return value.map((g) => {
    const group = g as { groupId?: string; first?: boolean; last?: boolean };
    const id = String(group.groupId);
    if (!groupOrdinals.has(id)) groupOrdinals.set(id, groupOrdinals.size);
    return { first: group.first, groupId: groupOrdinals.get(id), last: group.last };
  });
}

/** The order `deletionOf` picks the written deletion by: the date, then the author and the move. */
const deletionKey = (m: { attrs?: Record<string, unknown> }): string => `${String(m.attrs?.['date'] ?? '\uffff')}\u0000${String(m.attrs?.['author'] ?? '')}\u0000${String(m.attrs?.['move'] ?? '')}`;

function normalize(node: JsonNode): JsonNode {
  if (node.type === 'doc') groupOrdinals = new Map();
  const out: JsonNode = { type: node.type };
  if (node.attrs) {
    const attrs: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(node.attrs)) {
      if (LOCAL_ATTRS.has(k) || v === null || v === undefined) continue;
      // A cell's column widths are stored as one `w:tcW`, so only their sum (in twips) survives a reload.
      if (k === 'colwidth' && Array.isArray(v)) { attrs[k] = Math.round((v as number[]).reduce((a, b) => a + b, 0) * 15); continue; }
      attrs[k] = k === 'sdt' ? normalizeSdt(v) : sortKeys(v);
    }
    if (Object.keys(attrs).length > 0) out.attrs = attrs;
  }
  if (node.text !== undefined) out.text = node.text;
  if (node.marks) {
    const marks = node.marks
      .filter((m) => m.type !== 'runSource')
      .map((m) => {
        const attrs: Record<string, unknown> = {};
        for (const [k, v] of Object.entries(m.attrs ?? {})) {
          if (LOCAL_ATTRS.has(k) || v === null || v === undefined) continue;
          // A formatting change's properties before, by value: the tracker's names carry a namespace and a
          // local name, a reloaded file's its prefix too (ED-005 section 12.40).
          attrs[k] = m.type === 'formatChange' && k === 'before' && typeof v === 'string' ? canonical(JSON.parse(v)) : sortKeys(v);
        }
        return Object.keys(attrs).length > 0 ? { type: m.type, attrs } : { type: m.type };
      })
      // Marks of one type that may overlap (comments) are a set: their order is how they were added
      // here and how their ranges opened in a reload, which the file does not record (ED-005 section 12.20).
      .sort((a, b) => (a.type < b.type ? -1 : a.type > b.type ? 1 : 0) || JSON.stringify(a.attrs ?? {}).localeCompare(JSON.stringify(b.attrs ?? {})))
      // Two authors' deletions of one text in a room (ED-005 section 12.26): the file holds the earlier alone, Word nesting no w:del in a w:del (`deletionOf`).
      .filter((m, _i, all) => m.type !== 'deletion' || m === all.filter((d) => d.type === 'deletion').reduce((a, b) => (deletionKey(b) < deletionKey(a) ? b : a)));
    if (marks.length > 0) out.marks = marks;
  }
  if (node.content) {
    const content: JsonNode[] = [];
    for (const child of node.content.map(normalize)) {
      const last = content[content.length - 1];
      if (last && last.type === 'text' && child.type === 'text' && JSON.stringify(last.marks ?? []) === JSON.stringify(child.marks ?? [])) {
        last.text = (last.text ?? '') + (child.text ?? '');
      } else {
        content.push(child);
      }
    }
    out.content = content;
  }
  return out;
}

export interface SelfCheckResult {
  equal: boolean;
  /**
   * The index of the first top-level block that differs, with both sides as JSON, when not
   * equal; `story` names the story (`footnote:3`, `header:/word/header1.xml`) when the block is
   * not the body's (ED-003 section 6.2), or says the reload lacks the story altogether.
   */
  difference?: { index: number; live: string; reloaded: string; story?: string };
}

/** Compare the live projection with the one a reload gives. */
export function compareProjections(live: PMNode, reloaded: PMNode): SelfCheckResult {
  const a = normalizedJson(live);
  const b = normalizedJson(reloaded);
  const ac = a.content ?? [];
  const bc = b.content ?? [];
  const n = Math.max(ac.length, bc.length);
  for (let i = 0; i < n; i++) {
    const l = JSON.stringify(ac[i] ?? null);
    const r = JSON.stringify(bc[i] ?? null);
    if (l !== r) return { equal: false, difference: { index: i, live: l, reloaded: r } };
  }
  if (JSON.stringify(a.attrs ?? {}) !== JSON.stringify(b.attrs ?? {})) {
    return { equal: false, difference: { index: -1, live: JSON.stringify(a.attrs), reloaded: JSON.stringify(b.attrs) } };
  }
  return { equal: true };
}

export class SelfCheckError extends Error {
  constructor(readonly result: SelfCheckResult) {
    super(`round-trip self-check failed at block ${result.difference?.index}${result.difference?.story ? ` of ${result.difference.story}` : ''}: the saved document would not re-project to what is on screen`);
    this.name = 'SelfCheckError';
  }
}

export type { Schema };
