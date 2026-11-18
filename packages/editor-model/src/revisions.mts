// SPDX-License-Identifier: GPL-3.0-or-later
//
// Tracked changes in the edit layer (ED-005 section 5.2, E4.c): the file's revision elements read
// into the representation the import projects (marks on text, attributes on paragraphs, rows and
// atoms) and written back by the export. Revision ids are never stored (section 5.7): an element
// the import read keeps its id when it goes back unchanged, and one written afresh takes the next
// id the export allocates.
import { calendarOf, wordCalendarOf, revisionDateOf, revisionKindOf, type Element } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { rPrToElements } from '@docx4j/generated-objects-ts/builders/wml';
import type { Mark, Node as PMNode } from 'prosemirror-model';
import { copyWithout, revive, toPlain } from './json.mjs';
import type { Revision, RevisionInfo, PropertiesChange } from './schema.mjs';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** A revision element's dates: `w:date` and, where Word wrote it, `w16du:dateUtc`. */
type RevisionDates = { date?: wml.XmlCalendar; dateUtc?: wml.XmlCalendar };

/**
 * A revision's date as the edit layer keeps it: the instant, an ISO string in UTC to the second
 * (`2026-09-28T05:36:00Z`), read as Word means it: `w16du:dateUtc` when the element has one, else
 * `w:date` as the local wall clock, its `Z` notwithstanding (core-ts 0.3.0's `revisionDateOf`,
 * CR-002 section 29; ED-005 section 12.5 item 12).
 */
export function dateString(value: RevisionDates | undefined): string | null {
  const d = value ? revisionDateOf(value) : undefined;
  return d ? d.toISOString().replace(/\.000Z$/, 'Z') : null;
}

/** How many revision dates the editor has written with `w16du:dateUtc`: an export that wrote one declares the prefix on its part's root. */
let datesWritten = 0;
export function dateUtcWrites(): number {
  return datesWritten;
}

/** A revision's dates as Word writes them: `w:date` on the local wall clock (marked Z) and the instant in `w16du:dateUtc`. */
export function dateFields(iso: string): { date: wml.XmlCalendar; dateUtc: wml.XmlCalendar } {
  const d = new Date(iso);
  datesWritten++;
  return { date: wordCalendarOf(d), dateUtc: calendarOf(d) };
}

/** A revision element's author and date, and the move it belongs to. */
export function revisionInfo(tc: { author?: string } & RevisionDates, move: string | null): RevisionInfo {
  return { author: tc.author ?? null, date: dateString(tc), move };
}

/** A `w:ins`, `w:del`, `w:moveFrom` or `w:moveTo` value (`CTTrackChange`) for a side of a revision, with its id. */
export function trackChangeOf(info: RevisionInfo, id: number): wml.CTTrackChange {
  const out: wml.CTTrackChange = { TYPE_NAME: 'org_docx4j_wml.CTTrackChange', id, author: info.author ?? '' };
  if (info.date !== null) Object.assign(out, dateFields(info.date));
  return out;
}

/** A revision's `w:id`, `w:author` and dates, the fields every change element shares. */
function markupOf(info: RevisionInfo, id: number): { id: number; author: string } & RevisionDates {
  return info.date !== null ? { id, author: info.author ?? '', ...dateFields(info.date) } : { id, author: info.author ?? '' };
}

/** The four revision fields a paragraph mark's `w:rPr` or (`ins` and `del`) a row's `w:trPr` can hold. */
type RevisionHolder = { ins?: wml.CTTrackChange; del?: wml.CTTrackChange; moveFrom?: wml.CTTrackChange; moveTo?: wml.CTTrackChange };

/** A mark's or a row's revision, read out of its properties; null when there is none. */
export function revisionOf(pr: RevisionHolder | undefined, move: { from: string | null; to: string | null } = { from: null, to: null }): Revision | null {
  if (!pr) return null;
  const out: Revision = {};
  if (pr.ins) out.insertion = revisionInfo(pr.ins, null);
  else if (pr.moveTo) out.insertion = revisionInfo(pr.moveTo, move.to ?? '');
  if (pr.del) out.deletion = revisionInfo(pr.del, null);
  else if (pr.moveFrom) out.deletion = revisionInfo(pr.moveFrom, move.from ?? '');
  return out.insertion || out.deletion ? out : null;
}

/** The properties as plain JSON less their revision fields: what the attribute beside `markRevision` or a row's `revision` holds. */
export function withoutRevision(pr: object | undefined): unknown | null {
  if (!pr) return null;
  const plain = toPlain(pr) as Record<string, unknown>;
  for (const k of ['ins', 'del', 'moveFrom', 'moveTo']) delete plain[k];
  return Object.keys(plain).every((k) => k === 'TYPE_NAME') ? null : plain;
}

/**
 * Write a revision into properties (a paragraph mark's `w:rPr`, a row's `w:trPr`), replacing
 * whatever revision fields they held: a side that the original held as it is keeps the original's
 * element and id, a new one takes `nextId()`.
 */
export function writeRevision(pr: RevisionHolder, revision: Revision | null, original: Revision | null, originalPr: RevisionHolder | undefined, nextId: () => number): void {
  for (const k of ['ins', 'del', 'moveFrom', 'moveTo'] as const) delete pr[k];
  const side = (info: RevisionInfo | undefined, was: RevisionInfo | undefined, plain: 'ins' | 'del', move: 'moveTo' | 'moveFrom'): void => {
    if (!info) return;
    const key = info.move !== null ? move : plain;
    const same = was && was.author === info.author && was.date === info.date && was.move === info.move ? originalPr?.[key] : undefined;
    pr[key] = same ? copyWithout(same, []) : trackChangeOf(info, nextId());
  };
  side(revision?.insertion, original?.insertion, 'ins', 'moveTo');
  side(revision?.deletion, original?.deletion, 'del', 'moveFrom');
}

/**
 * Whether a run-level revision holder (`w:ins`, `w:del`, `w:moveTo`, `w:moveFrom`) is the side a
 * piece of content says it is in: the same side, a move's holder for a move, the same author and
 * date. Content typed into another author's insertion carries its own insertion, and so leaves
 * that author's holder (ED-005 section 5.3).
 */
export function isHolderOf(h: Element, side: 'insertion' | 'deletion', info: RevisionInfo | undefined): boolean {
  const kind = revisionKindOf(h);
  if (kind === undefined || info === undefined) return false;
  const expected = side === 'insertion' ? (info.move !== null ? 'moveTo' : 'ins') : (info.move !== null ? 'moveFrom' : 'del');
  if (kind !== expected) return false;
  const tc = h.value as wml.CTTrackChange;
  return (tc.author ?? null) === info.author && dateString(tc) === info.date;
}

/**
 * A new run-level holder for one side of a revision, as the engine's tracker writes one: `w:ins`
 * (`RunIns`) or `w:del` (`RunDel`), or `w:moveTo` or `w:moveFrom` (`RunTrackChange`) for a side
 * that belongs to a move, with `w:id`, `w:author` and `w:date`.
 */
export function revisionHolder(side: 'insertion' | 'deletion', info: RevisionInfo, id: number): Element {
  const markup = markupOf(info, id);
  if (info.move !== null) {
    const value: wml.RunTrackChange = { TYPE_NAME: 'org_docx4j_wml.RunTrackChange', ...markup, accOrBarOrBox: [] };
    return { name: { namespaceURI: W_NS, localPart: side === 'insertion' ? 'moveTo' : 'moveFrom' }, value };
  }
  const value: wml.RunIns | wml.RunDel = side === 'insertion'
    ? { TYPE_NAME: 'org_docx4j_wml.RunIns', ...markup, customXmlOrSmartTagOrSdt: [] }
    : { TYPE_NAME: 'org_docx4j_wml.RunDel', ...markup, customXmlOrSmartTagOrSdt: [] };
  return { name: { namespaceURI: W_NS, localPart: side === 'insertion' ? 'ins' : 'del' }, value };
}

/** A `formatChange` mark's attributes. */
export interface FormatChangeAttrs { author: string | null; date: string | null; before: string }

/**
 * The `before` of a `formatChange` mark: a run's old properties as `w:rPrChange` holds them (the
 * element list the engine's tracker writes), as a JSON string, so the mark stays flat (ED-005
 * section 12.3 item 2). An empty list is left out, as an unmarshalled `<w:rPr/>` has none.
 */
export function formatChangeBefore(rPr: wml.CTRPrChange['rPr'] | undefined): string {
  if (!rPr) return 'null';
  const plain = toPlain(rPr) as { egrPrBase?: unknown[] };
  if (Array.isArray(plain.egrPrBase) && plain.egrPrBase.length === 0) delete plain.egrPrBase;
  return JSON.stringify(plain);
}

/** The `before` for a run whose properties were `rPr` before the change (a tracked formatting change made in the editor). */
export function formatChangeBeforeRPr(rPr: wml.RPr | undefined): string {
  return formatChangeBefore({ TYPE_NAME: 'org_docx4j_wml.CTRPrChange.RPr', egrPrBase: rPrToElements(rPr) });
}

/** A run's `w:rPrChange` as a `formatChange` mark's attributes. */
export function formatChangeAttrs(change: wml.CTRPrChange): FormatChangeAttrs {
  const info = revisionInfo(change, null);
  return { author: info.author, date: info.date, before: formatChangeBefore(change.rPr) };
}

/** The `w:rPrChange` a `formatChange` mark writes, with a fresh id. */
export function rPrChangeOf(attrs: FormatChangeAttrs, id: number): wml.CTRPrChange {
  const before = JSON.parse(attrs.before) as unknown;
  const markup = markupOf({ author: attrs.author, date: attrs.date, move: null }, id);
  return { TYPE_NAME: 'org_docx4j_wml.CTRPrChange', ...markup, rPr: before === null ? { TYPE_NAME: 'org_docx4j_wml.CTRPrChange.RPr' } : (revive(before) as wml.CTRPrChange['rPr']) };
}

/** A paragraph's `w:pPrChange` as the attribute holds it; null when there is none. */
export function propertiesChangeOf(change: wml.CTPPrChange | undefined): PropertiesChange | null {
  if (!change) return null;
  return { author: change.author ?? null, date: dateString(change), before: change.pPr ? toPlain(change.pPr) : null };
}

/** Whether a node is deleted content in the accepted view: text with a deletion mark, an atom whose run was deleted. */
export function isDeleted(node: PMNode): boolean {
  if (node.isText) return node.marks.some((m) => m.type.name === 'deletion');
  const r = node.attrs['revision'] as Revision | null | undefined;
  return !!r?.deletion;
}

/**
 * The deletion a piece of text goes out in when two overlap (ED-005 section 12.7 item 5 (a)): Word
 * nests no `w:del` in a `w:del`, so the earlier, then by author and move, which every peer picks
 * alike whatever order its marks came in.
 */
export function deletionOf(marks: readonly Mark[]): Mark | undefined {
  let out: Mark | undefined;
  const key = (m: Mark): string => `${String(m.attrs['date'] ?? '\uffff')}\u0000${String(m.attrs['author'] ?? '')}\u0000${String(m.attrs['move'] ?? '')}`;
  for (const m of marks) if (m.type.name === 'deletion' && (!out || key(m) < key(out))) out = m;
  return out;
}

/** The revision marks a piece of text carries: its insertion and its deletion (the earlier, when two overlap). */
export function revisionMarksOf(marks: readonly Mark[]): { insertion?: Mark; deletion?: Mark } {
  const out: { insertion?: Mark; deletion?: Mark } = {};
  for (const m of marks) if (m.type.name === 'insertion') out.insertion = m;
  const deletion = deletionOf(marks);
  if (deletion) out.deletion = deletion;
  return out;
}
