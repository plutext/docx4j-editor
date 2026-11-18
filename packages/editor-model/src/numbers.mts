// SPDX-License-Identifier: GPL-3.0-or-later
//
// Number formats for what the page shows beside the lists (ED-003 sections 6.3 and 6.4): the
// page number from `w:pgNumType`, a note's number from `w:footnotePr`. The formats Word uses
// for those: decimal, roman, letters, and the chicago symbols footnotes may take. The engine's
// list-label formatter covers many more but is not exported; asked of core-ts, this stays small.
import type { Node as PMNode } from 'prosemirror-model';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { NoteKind } from './stories.mjs';

const ROMAN: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];

function roman(n: number): string {
  let out = '';
  let rest = n;
  for (const [value, digits] of ROMAN) while (rest >= value) { out += digits; rest -= value; }
  return out;
}

/** Word's letters: a to z, then aa, bb, cc (not ab). */
function letters(n: number): string {
  const index = (n - 1) % 26;
  const repeat = Math.floor((n - 1) / 26) + 1;
  return String.fromCharCode(97 + index).repeat(repeat);
}

/** Word's footnote symbols: *, †, ‡, §, then doubled, tripled. */
const CHICAGO = ['*', '†', '‡', '§'];

/** A number in a `w:numFmt` the page shows; formats this does not know show decimal. */
export function formatNumber(n: number, fmt: wml.NumberFormat | undefined): string {
  if (n < 1) return String(n);
  switch (fmt) {
    case 'lowerRoman': return roman(n);
    case 'upperRoman': return roman(n).toUpperCase();
    case 'lowerLetter': return letters(n);
    case 'upperLetter': return letters(n).toUpperCase();
    case 'chicago': return CHICAGO[(n - 1) % 4]!.repeat(Math.floor((n - 1) / 4) + 1);
    case 'decimalZero': return n < 10 ? `0${n}` : String(n);
    case 'none': return '';
    default: return String(n);
  }
}

// ---------------------------------------------------------------- note numbering (ED-003 section 6.3)

/** `w:footnotePr` / `w:endnotePr` as the display reads them: the settings part's under the section's. */
export interface NoteNumbering {
  fmt: wml.NumberFormat | undefined;
  start: number;
  restart: 'continuous' | 'eachSect' | 'eachPage';
}

type NoteProps = { numFmt?: { val?: wml.NumberFormat }; numStart?: { val?: number }; numRestart?: { val?: string } } | undefined;

/** The section's properties over the document's, property by property (Word's rule). */
export function noteNumberingFor(kind: NoteKind, sectPr: wml.SectPr | null | undefined, settings: { footnotePr?: unknown; endnotePr?: unknown } | undefined): NoteNumbering {
  const key = kind === 'footnote' ? 'footnotePr' : 'endnotePr';
  const doc = settings?.[key] as NoteProps;
  const sec = (sectPr as { footnotePr?: unknown; endnotePr?: unknown } | null | undefined)?.[key] as NoteProps;
  const restart = sec?.numRestart?.val ?? doc?.numRestart?.val;
  return {
    // Unstated, Word numbers footnotes 1, 2, 3 and endnotes i, ii, iii (ED-006 section 9.6).
    fmt: sec?.numFmt?.val ?? doc?.numFmt?.val ?? (kind === 'endnote' ? 'lowerRoman' : undefined),
    start: sec?.numStart?.val ?? doc?.numStart?.val ?? 1,
    restart: restart === 'eachSect' || restart === 'eachPage' ? restart : 'continuous',
  };
}

/**
 * The number each reference shows, by `kind:id` of its first reference, over the document: the
 * counts run through the body in reference order, restarting where the section says (each
 * section, or each page from the page-break markers and the page-starting section breaks), from
 * the section's start, in its format. A document without notes costs one test.
 */
export function noteLabels(doc: PMNode, sections: (wml.SectPr | null)[], settings: { footnotePr?: unknown; endnotePr?: unknown } | undefined): Map<string, string> {
  const out = new Map<string, string>();
  let section = 0;
  const counts: Record<NoteKind, number> = { footnote: 0, endnote: 0 };
  const props = (kind: NoteKind): NoteNumbering => noteNumberingFor(kind, sections[section], settings);
  const restartFor = (when: 'eachSect' | 'eachPage'): void => {
    for (const kind of ['footnote', 'endnote'] as const) if (props(kind).restart === when || (when === 'eachSect' && props(kind).restart === 'eachPage')) counts[kind] = 0;
  };
  doc.descendants((node) => {
    if (node.type.name === 'paragraph') {
      const sectPr = node.attrs['sectPr'] as wml.SectPr | null;
      if (sectPr) {
        // A new section restarts a per-section count whether or not the break starts a page (Word's rule).
        section++;
        restartFor('eachSect');
      }
      return true;
    }
    if (node.type.name === 'opaqueInline' && node.attrs['kind'] === 'lastRenderedPageBreak') { restartFor('eachPage'); return false; }
    if (node.type.name === 'footnoteRef') {
      const kind: NoteKind = node.attrs['kind'] === 'endnote' ? 'endnote' : 'footnote';
      const key = `${kind}:${String(node.attrs['footnoteId'])}`;
      if (!out.has(key)) {
        const p = props(kind);
        out.set(key, formatNumber(p.start + counts[kind], p.fmt));
        counts[kind]++;
      }
      return false;
    }
    return !node.isLeaf;
  });
  return out;
}
