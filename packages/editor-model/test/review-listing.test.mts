// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 6.5, step D1: the listing of tracked changes made incremental. Each top-level
// block's pieces are memoised per block node and the grouping is one pass over the blocks' lists,
// where step C4's function walked the whole document and, between two pieces of one side, every
// block between them. That function is kept here as the oracle: over every fixture, the check
// folders, and each step of a sequence of tracked edits, the two list the same changes, in
// positions and grouping alike. A keystroke walks one block again. The revisions shown in Word
// only (`preservedChangesOf`) are read from the Word files.
import { describe, test, expect } from 'vitest';
import { readdir } from 'node:fs/promises';
import { join } from 'node:path';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { EditorDocument, runCommand, runCommandAsync, stateFor, recordTransaction, trackedChangesOf, preservedChangesOf, preservedGroupsOf, listingBlockWalks, type Revision, type ParagraphAttrs, type TrackedChangeInfo, type ChangePiece } from '../src/index.mjs';
import { fixture, fixturesDir, allFixtures, wordDate } from './helpers.mjs';
import { rowTextOf } from '../src/review.mjs';

// ---------------------------------------------------------------------------------------------
// The oracle: step C4's `trackedChangesOf`, as it was before step D1 (ED-005 section 12.6).
// ---------------------------------------------------------------------------------------------

type Side = 'insertion' | 'deletion' | 'format' | 'properties';

const MARKERS = new Set(['bookmarkStart', 'bookmarkEnd', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd', 'commentRangeStart', 'commentRangeEnd', 'commentReference', 'permStart', 'permEnd', 'lastRenderedPageBreak', 'proofErr', 'r']);
const isMarker = (node: PMNode): boolean => (node.type.name === 'opaqueInline' || node.type.name === 'opaqueBlock') && MARKERS.has(node.attrs['kind'] as string);

function atomText(node: PMNode): string {
  if (node.type.name === 'tab') return '\t';
  if (node.type.name === 'hardBreak') return node.attrs['kind'] === 'page' ? '\f' : '\v';
  return '';
}

const sidesOf = (r: Revision | null | undefined): ('insertion' | 'deletion')[] => [...(r?.insertion ? ['insertion' as const] : []), ...(r?.deletion ? ['deletion' as const] : [])];

function piecesOf(doc: PMNode): ChangePiece[] {
  const pieces: ChangePiece[] = [];
  const inline = (node: PMNode, at: number, into: ChangePiece[], holder: string): void => {
    if (node.isText) {
      const link = node.marks.find((m) => m.type.name === 'link');
      const own = link ? `${holder}link:${String(link.attrs['sourceId'])}|${String(link.attrs['rId'])}|${String(link.attrs['anchor'])};` : holder;
      for (const m of node.marks) {
        const name = m.type.name;
        if (name !== 'insertion' && name !== 'deletion' && name !== 'formatChange') continue;
        const side: Side = name === 'formatChange' ? 'format' : name;
        into.push({ what: side === 'format' ? 'format' : 'text', side, author: (m.attrs['author'] as string | null) ?? null, date: (m.attrs['date'] as string | null) ?? null, move: side === 'format' ? null : ((m.attrs['move'] as string | null) ?? null), from: at, to: at + node.nodeSize, pos: at, mark: m, ...(side === 'format' ? { before: m.attrs['before'] } : {}), text: node.text ?? '', holder: own });
      }
      return;
    }
    const own = node.type.name === 'opaqueInline' && node.attrs['kind'] === 'fldSimple' ? `${holder}fld@${at};` : holder;
    const r = node.attrs['revision'] as Revision | null | undefined;
    for (const side of sidesOf(r)) {
      const info = r![side]!;
      into.push({ what: 'atom', side, author: info.author, date: info.date, move: info.move, from: at, to: at + node.nodeSize, pos: at, text: atomText(node), holder: own });
    }
  };
  /** A paragraph's pieces: its properties' change, its inline content (an inline control's inside with a holder of its own), its mark with `markHolder`. */
  const paragraph = (node: PMNode, pos: number, into: ChangePiece[], holder = '', markHolder = holder): void => {
    const attrs = node.attrs as unknown as ParagraphAttrs;
    if (attrs.pPrChange) into.push({ what: 'properties', side: 'properties', author: attrs.pPrChange.author, date: attrs.pPrChange.date, move: null, from: pos, to: pos + node.nodeSize, pos, before: attrs.pPrChange.before, text: '', holder });
    const walk = (parent: PMNode, base: number, h: string): void => {
      parent.forEach((child, offset) => {
        const at = base + offset;
        inline(child, at, into, h);
        if (child.type.name === 'sdtInline') walk(child, at + 1, `${h}sdt@${at};`);
      });
    };
    walk(node, pos + 1, holder);
    const close = pos + node.nodeSize - 1;
    for (const side of sidesOf(attrs.markRevision)) {
      const info = attrs.markRevision![side]!;
      into.push({ what: 'mark', side, author: info.author, date: info.date, move: info.move, from: close, to: close + 1, pos, text: '\r', holder: markHolder });
    }
  };
  /**
   * The blocks under `parent`, with the holder their pieces are in: a block control's inside has one
   * of its own, but its last paragraph's mark is the outside's, as Office JS lists it (core-ts check 30).
   */
  const blocks = (parent: PMNode, base: number, holder: string): void => parent.forEach((node, offset) => {
    const pos = base + offset;
    if (node.type.name === 'sdtBlock') {
      const inside = `${holder}sdtb@${pos};`;
      node.forEach((child, childOffset, index) => {
        const at = pos + 1 + childOffset;
        if (child.type.name === 'paragraph') paragraph(child, at, pieces, inside, index === node.childCount - 1 ? holder : inside);
        else blocks(child, at + 1, inside);
      });
      return;
    }
    if (node.type.name === 'table_row') {
      const r = node.attrs['revision'] as Revision | null;
      const side = r?.insertion ? 'insertion' : r?.deletion ? 'deletion' : undefined;
      // A plain row keeps its cells' changes to itself; a row that is a revision is one piece (core-ts check 31).
      if (!side) { blocks(node, pos + 1, `${holder}row@${pos};`); return; }
      const inner: ChangePiece[] = [];
      node.descendants((child, rel) => { if (child.type.name === 'paragraph') paragraph(child, pos + 1 + rel, inner); return child.type.name !== 'paragraph'; });
      const info = r![side]!;
      pieces.push({ what: 'row', side, author: info.author, date: info.date, move: null, from: pos, to: pos + node.nodeSize, pos, text: rowTextOf(node, side), holder, inner: inner.filter((p) => p.side === side) });
      return;
    }
    if (node.type.name === 'paragraph') { paragraph(node, pos, pieces, holder); return; }
    blocks(node, pos + 1, holder);
  });
  blocks(doc, 0, '');
  return pieces;
}

function touches(doc: PMNode, a: number, b: number): boolean {
  if (b <= a) return true;
  let ok = true;
  doc.nodesBetween(a, b, (node, pos) => {
    if (!ok) return false;
    const end = pos + node.nodeSize;
    if (node.isText) { if (pos < b && end > a) ok = false; return false; }
    if (node.isLeaf) { if (pos >= a && pos < b && !isMarker(node)) ok = false; return false; }
    // A block control's end is no content (core-ts check 30), nor a table's (check 31).
    if (node.type.name !== 'sdtBlock' && node.type.name !== 'table' && end - 1 >= a && end - 1 < b) { ok = false; return false; }
    return true;
  });
  return ok;
}

type ChangeType = TrackedChangeInfo['type'];
const typeOf = (side: Side): ChangeType => (side === 'insertion' ? 'Added' : side === 'deletion' ? 'Deleted' : 'Formatted');

function oracle(doc: PMNode): TrackedChangeInfo[] {
  const pieces = piecesOf(doc);
  const out: TrackedChangeInfo[] = [];
  const open = new Map<Side, { change: TrackedChangeInfo; last: ChangePiece }>();
  const moves = new Map<string, TrackedChangeInfo>();
  const start = (p: ChangePiece): TrackedChangeInfo => ({ type: typeOf(p.side), author: p.author, date: p.date, text: '', from: p.from, to: p.to, pieces: [], ...(p.what === 'row' ? { row: true } : {}) });
  const add = (c: TrackedChangeInfo, p: ChangePiece): void => { c.pieces.push(p, ...(p.inner ?? [])); c.text += p.text; c.to = Math.max(c.to, p.to); if (p.what !== 'row') delete c.row; };
  for (const p of pieces) {
    if (p.side === 'properties') { const c = start(p); add(c, p); out.push(c); continue; }
    if (p.move !== null && p.move !== '') {
      const key = `${p.side}:${p.move}`;
      let c = moves.get(key);
      if (!c) { c = { ...start(p), move: p.move }; moves.set(key, c); out.push(c); }
      add(c, p);
      continue;
    }
    const o = open.get(p.side);
    if (o && o.change.author === p.author && touches(doc, o.last.to, p.from) && (o.last.holder ?? '') === (p.holder ?? '')) { add(o.change, p); o.last = p; continue; }
    const c = start(p);
    add(c, p);
    out.push(c);
    open.set(p.side, { change: c, last: p });
  }
  const order: Record<ChangeType, number> = { Added: 0, Deleted: 1, Formatted: 2 };
  return out.sort((a, b) => a.from - b.from || order[a.type] - order[b.type]);
}

// ---------------------------------------------------------------------------------------------

/** A listing as plain values: a mark instance by its name and attributes, so two listings compare without walking the schema. */
function plainListing(changes: readonly TrackedChangeInfo[]): unknown[] {
  return changes.map((c) => ({
    ...c,
    // `holder` keys by position, absolute in the oracle and within a block in the listing: the grouping it makes is compared, not the key.
    pieces: c.pieces.map((p) => ({ ...p, mark: p.mark ? [p.mark.type.name, p.mark.attrs] : undefined, holder: undefined, inner: p.inner?.map((q) => ({ ...q, mark: q.mark ? [q.mark.type.name, q.mark.attrs] : undefined, holder: undefined })) })),
  }));
}

function agrees(doc: PMNode, label: string): void {
  expect(plainListing(trackedChangesOf(doc)), label).toEqual(plainListing(oracle(doc)));
}

/** The .docx files of the check folders (Word's accepts and rejects, ED-005 section 12.6), which `allFixtures` does not read. */
async function checkFiles(): Promise<string[]> {
  const out: string[] = [];
  for (const dir of ['check18', 'check20', 'check22', 'check24', 'check25', 'check26', 'check28']) {
    for (const name of await readdir(join(fixturesDir, 'revisions', dir))) if (name.endsWith('.docx')) out.push(`revisions/${dir}/${name}`);
  }
  return out.sort();
}

describe('the listing, incremental, held to its oracle (ED-005 section 6.5, step D1)', () => {
  test('every fixture and every check file lists the same changes', async () => {
    // The oracle is the function as step D1 found it, without section 12.113's rules. The two fixtures those rules change, both
    // ONLYOFFICE's shape and neither Word's, are held to them and to a fresh walk in `empty-revisions.test.mts`.
    const changedSince = new Set(['authorless-contents.docx', 'numbering-clauses.docx']);
    const names = [...(await allFixtures()), ...(await checkFiles())].filter((name) => !changedSince.has(name));
    let listed = 0;
    for (const name of names) {
      const session = await EditorDocument.open(await fixture(name));
      const doc = session.doc;
      agrees(doc, name);
      listed += trackedChangesOf(doc).length;
    }
    // The Word files and the check folders hold changes enough for the comparison to mean something.
    expect(listed).toBeGreaterThan(200);
  }, 300_000);

  /** A transaction as the view's input makes one, recorded as the view records it, and applied. */
  function input(state: EditorState, make: (state: EditorState) => Transaction): EditorState {
    return state.apply(recordTransaction(make(state), state).tr);
  }

  /** Where `text` starts in the document, inside one text node. */
  function at(doc: PMNode, text: string): number {
    let pos = -1;
    doc.descendants((n, p) => { if (pos < 0 && n.isText && n.text!.includes(text)) pos = p + n.text!.indexOf(text); return pos < 0; });
    if (pos < 0) throw new Error(`"${text}" is not in one text node`);
    return pos;
  }

  /** The first text node carrying a mark of this name (and, for a move, one with its name set). */
  function marked(doc: PMNode, mark: 'insertion' | 'deletion' | 'formatChange', move = false): { from: number; to: number } {
    let found: { from: number; to: number } | undefined;
    doc.descendants((n, p) => {
      if (found || !n.isText) return !found;
      if (n.marks.some((m) => m.type.name === mark && (!move || m.attrs['move']))) found = { from: p, to: p + n.nodeSize };
      return false;
    });
    if (!found) throw new Error(`no text with ${mark}`);
    return found;
  }

  test('each step of a sequence of tracked edits on revisions-word15.docx lists the same changes', async () => {
    const session = await EditorDocument.open(await fixture('revisions/revisions-word15.docx'));
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author C';
    session.trackedChangeDate = new Date('2026-09-28T10:00:00Z');
    let state = stateFor(session);
    agrees(state.doc, 'as opened');
    const steps: [string, (s: EditorState) => Transaction][] = [
      ['typing inside an insertion', (s) => { const r = marked(s.doc, 'insertion'); return s.tr.insertText('X', r.from + 1); }],
      ['typing inside a deletion', (s) => { const r = marked(s.doc, 'deletion'); return s.tr.insertText('Y', r.from + 1); }],
      ['typing over a range across a revision', (s) => { const r = marked(s.doc, 'insertion'); return s.tr.insertText('Z', r.to - 1, r.to + 2); }],
      ['typing inside moved text', (s) => { const r = marked(s.doc, 'insertion', true); return s.tr.insertText('M', r.from + 1); }],
      ['deleting inside moved text', (s) => { const r = marked(s.doc, 'insertion', true); return s.tr.delete(r.from + 2, r.from + 4); }],
      ['bold over a formatting change', (s) => { const r = marked(s.doc, 'formatChange'); return s.tr.addMark(r.from, r.to, s.schema.marks['bold']!.create()); }],
    ];
    for (const [label, make] of steps) {
      state = input(state, make);
      agrees(state.doc, label);
    }
    // Backspace across a paragraph mark: the mark deleted, the two paragraphs not joined.
    session.replace(state.doc);
    const second = at(state.doc, 'bravo');
    let boundary = -1;
    state.doc.forEach((node, pos) => { if (boundary < 0 && pos + node.nodeSize > second) boundary = pos; });
    state = input(state, (s) => s.tr.join(boundary));
    agrees(state.doc, 'Backspace at a paragraph\'s start');
    // A row deleted, through the command, which marks it rather than removing it.
    session.replace(state.doc);
    let cell = -1;
    state.doc.descendants((n, p) => { if (cell < 0 && n.type.name === 'table_cell') cell = p + 2; return cell < 0; });
    expect(runCommand(session, 'table.deleteRows', undefined, { from: cell }).ok).toBe(true);
    agrees(session.doc, 'a row deleted');
    // Everything accepted, then everything rejected, from the edited document.
    const accepted = await EditorDocument.open(await fixture('revisions/revisions-word15.docx'));
    accepted.replace(session.doc);
    expect((await runCommandAsync(accepted, 'trackedChanges.acceptAll')).ok).toBe(true);
    agrees(accepted.doc, 'all accepted');
    expect(trackedChangesOf(accepted.doc)).toEqual([]);
  }, 120_000);

  test('after a keystroke, the listing walks the one block the keystroke touched', async () => {
    const session = await EditorDocument.open(await fixture('revisions/revisions-word15.docx'));
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author C';
    const state = stateFor(session);
    const first = trackedChangesOf(state.doc);
    const walks = listingBlockWalks();
    // The same document again is the memo's: nothing walked.
    expect(trackedChangesOf(state.doc)).toEqual(first);
    expect(listingBlockWalks()).toBe(walks);
    const next = input(state, (s) => s.tr.insertText('X', at(s.doc, 'bravo')));
    const second = trackedChangesOf(next.doc);
    expect(listingBlockWalks()).toBe(walks + 1);
    expect(second.length).toBe(first.length + 1);
    agrees(next.doc, 'after the keystroke');
  });
});

describe('the revisions shown in Word only (ED-005 section 6.5, step D1)', () => {
  for (const name of ['revisions-word2010.docx', 'revisions-word15.docx']) {
    test(`${name}: the section's, the table's, the grid's and the cells' changes, in document order, with their authors and dates`, async () => {
      const session = await EditorDocument.open(await fixture(`revisions/${name}`));
      const doc = session.doc;
      const found = preservedChangesOf(doc, (id) => session.ctx.ids.get(id)?.value);
      expect(found.map((p) => p.kind)).toEqual(['tblPrChange', 'tblGridChange', 'tcPrChange', 'tcPrChange', 'tcPrChange', 'tcPrChange', 'sectPrChange']);
      // The seven the diagnostics count in the body, and not the footnote's revision (a story's).
      const counted = Object.entries(session.diagnostics.tracked.preserved).reduce((n, [, v]) => n + v, 0);
      expect(found.length).toBe(counted);
      for (const p of found) {
        expect(p.author).toBe('Author A');
        expect(wordDate(p.date!)).toMatch(/^2026-09-28T\d\d:\d\d:00Z$/);
      }
      // A table's change is on the table, a cell's on its cell, the final section's at the end.
      expect(doc.nodeAt(found[0]!.pos)!.type.name).toBe('table');
      expect(doc.nodeAt(found[2]!.pos)!.type.name).toBe('table_cell');
      expect(found[6]!.pos).toBe(doc.content.size);
      // As Office JS lists them (core-ts check 27): the table's records one change, at the table; the section's another.
      const groups = preservedGroupsOf(doc, found);
      expect(groups.map((g) => [g.kind, g.records.length])).toEqual([['tableProperties', 6], ['sectPrChange', 1]]);
      expect(groups[0]!.pos).toBe(found[0]!.pos);
    });
  }
});
