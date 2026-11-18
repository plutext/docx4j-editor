// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-010 section 3.3: the rules a pasted slice goes through, pure over the slice. The view's paste
// and a dragged copy both run them (packages/editor/test/clipboard.test.ts is the view's end to
// end, over the Word 365 probe of section 9.1).
import { describe, test, expect } from 'vitest';
import { Slice, type Node as PMNode } from 'prosemirror-model';
import { EditorDocument, pastedSlice, closedAtMark, commentLengthsIn, commentLengthIn, readControl, stateFor } from '../src/index.mjs';
import { TextSelection } from 'prosemirror-state';
import { fixture } from './helpers.mjs';

const OFF = { tracking: false, story: false };
const ON = { tracking: true, story: false };

/** Every node of a type in a fragment, in order. */
function nodesOf(root: PMNode | Slice, type: string): PMNode[] {
  const out: PMNode[] = [];
  const content = root instanceof Slice ? root.content : root.content;
  content.descendants((n) => { if (n.type.name === type) out.push(n); return true; });
  return out;
}

const markNames = (n: PMNode): string[] => [...new Set(n.content.content.flatMap((c) => c.marks.map((m) => m.type.name)))];

describe('pastedSlice: the rules (ED-010 section 3.3)', () => {
  test('a paragraph drops its source id and names it as copiedFrom; its style, numbering and runs stay; its section goes; a bookmark goes; a tab stays', async () => {
    const session = await EditorDocument.open(await fixture('numbering-clauses.docx'));
    const doc = session.doc;
    const slice = new Slice(doc.content.cut(0, Math.min(doc.content.size, 2000)), 0, 0);
    const out = pastedSlice(slice, OFF);
    const before = nodesOf(slice, 'paragraph');
    const after = nodesOf(out, 'paragraph');
    expect(after.length).toBe(before.length);
    for (let i = 0; i < before.length; i++) {
      const b = before[i]!;
      const a = after[i]!;
      expect(a.attrs['sourceId']).toBeNull();
      expect(a.attrs['copiedFrom']).toBe(b.attrs['sourceId']);
      expect(a.attrs['pStyle']).toBe(b.attrs['pStyle']);
      expect(a.attrs['numPr']).toEqual(b.attrs['numPr']);
      expect(a.attrs['sectPr']).toBeNull();
      expect(a.textContent).toBe(b.textContent);
      // the runs keep their source, the export rebuilding from a copy of its properties
      expect(markNames(a)).toEqual(markNames(b));
    }
    expect(nodesOf(slice, 'opaqueInline').some((n) => n.attrs['kind'] === 'bookmarkStart')).toBe(true);
    expect(nodesOf(out, 'opaqueInline').filter((n) => n.attrs['kind'] === 'bookmarkStart' || n.attrs['kind'] === 'bookmarkEnd')).toHaveLength(0);
    expect(nodesOf(out, 'tab').length).toBe(nodesOf(slice, 'tab').length);
    // the same openness kept
    expect([out.openStart, out.openEnd]).toEqual([0, 0]);
    const open = new Slice(doc.content.cut(5, 60), 1, 1);
    expect([pastedSlice(open, OFF).openStart, pastedSlice(open, OFF).openEnd]).toEqual([1, 1]);
  });

  test('a paragraph already copied keeps the element it was copied from', async () => {
    const session = await EditorDocument.open(await fixture('numbering-clauses.docx'));
    const first = pastedSlice(new Slice(session.doc.content.cut(0, 400), 0, 0), OFF);
    const again = pastedSlice(first, OFF);
    const a = nodesOf(first, 'paragraph')[1]!;
    const b = nodesOf(again, 'paragraph')[1]!;
    expect(a.attrs['copiedFrom']).toBeTruthy();
    expect(b.attrs['copiedFrom']).toBe(a.attrs['copiedFrom']);
  });

  test('revisions: kept with tracking off (Word, case 3); taken as accepted with tracking on (case 4) and into a story', async () => {
    const session = await EditorDocument.open(await fixture('review/review-for-word-word15.docx'));
    const doc = session.doc;
    const slice = new Slice(doc.content, 0, 0);
    const text = (s: Slice, mark: string): string => { let t = ''; s.content.descendants((n) => { if (n.isText && n.marks.some((m) => m.type.name === mark)) t += n.text; return true; }); return t; };
    expect(text(slice, 'insertion').length).toBeGreaterThan(0);
    expect(text(slice, 'deletion').length).toBeGreaterThan(0);
    const off = pastedSlice(slice, OFF);
    expect(text(off, 'insertion')).toBe(text(slice, 'insertion'));
    expect(text(off, 'deletion')).toBe(text(slice, 'deletion'));
    const on = pastedSlice(slice, ON);
    expect(text(on, 'insertion')).toBe('');
    expect(text(on, 'deletion')).toBe('');
    // the deleted text is out, the inserted text in, plain
    const plain = (s: Slice): string => { let t = ''; s.content.descendants((n) => { if (n.isText) t += n.text; return true; }); return t; };
    expect(plain(on).length).toBe(plain(slice).length - text(slice, 'deletion').length);
    const story = pastedSlice(slice, { tracking: false, story: true });
    expect(text(story, 'insertion')).toBe('');
    expect(text(story, 'deletion')).toBe('');
    // a paragraph's mark revision follows the same rule
    const marked = nodesOf(slice, 'paragraph').filter((p) => p.attrs['markRevision']).length;
    expect(marked).toBeGreaterThan(0);
    expect(nodesOf(off, 'paragraph').filter((p) => p.attrs['markRevision']).length).toBe(marked);
    expect(nodesOf(on, 'paragraph').filter((p) => p.attrs['markRevision']).length).toBe(0);
  });

  test('comments: the mark takes the id the caller gives it, or goes; the whole-or-part lengths', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    const doc = session.doc;
    const slice = new Slice(doc.content, 0, 0);
    const lengths = commentLengthsIn(slice.content);
    expect(lengths.size).toBeGreaterThan(0);
    for (const [id, length] of lengths) expect(commentLengthIn(doc, id)).toBe(length);
    const ids = [...lengths.keys()];
    const mapped = pastedSlice(slice, { ...OFF, commentId: (id) => (id === ids[0] ? 'copy-1' : undefined) });
    const after = commentLengthsIn(mapped.content);
    expect(after.get('copy-1')).toBe(lengths.get(ids[0]!));
    expect(after.has(ids[0]!)).toBe(false);
    for (const id of ids.slice(1)) expect(after.has(id)).toBe(false);
    // no mapping: every comment mark goes, the text stays
    const dropped = pastedSlice(slice, OFF);
    expect(commentLengthsIn(dropped.content).size).toBe(0);
    expect(dropped.content.textBetween(0, dropped.content.size)).toBe(slice.content.textBetween(0, slice.content.size));
    // a part of a comment's range is shorter than the whole
    let at = -1;
    doc.descendants((n, pos) => { if (at < 0 && n.isText && n.marks.some((m) => m.type.name === 'comment' && String(m.attrs['id']) === ids[0])) at = pos; return at < 0; });
    const part = doc.slice(at, at + 1);
    expect(commentLengthsIn(part.content).get(ids[0]!)).toBe(1);
    expect(commentLengthIn(doc, ids[0]!)).toBeGreaterThan(1);
  });

  test('a table, its rows and cells drop their source ids and a row its control stack; into a story the table goes', async () => {
    const session = await EditorDocument.open(await fixture('tables.docx'));
    const slice = new Slice(session.doc.content, 0, 0);
    const out = pastedSlice(slice, OFF);
    for (const type of ['table', 'table_row', 'table_cell']) {
      expect(nodesOf(out, type).length).toBe(nodesOf(slice, type).length);
      for (const n of nodesOf(out, type)) expect(n.attrs['sourceId']).toBeNull();
    }
    for (const row of nodesOf(out, 'table_row')) expect(row.attrs['sdt']).toBeNull();
    const cellPr = nodesOf(slice, 'table_cell').map((c) => c.attrs['tcPr']);
    expect(nodesOf(out, 'table_cell').map((c) => c.attrs['tcPr'])).toEqual(cellPr);
    const story = pastedSlice(slice, { tracking: false, story: true });
    expect(nodesOf(story, 'table')).toHaveLength(0);
    expect(nodesOf(story, 'paragraph').length).toBeLessThan(nodesOf(slice, 'paragraph').length);
  });

  test('a control gets the id the caller allocates; a link keeps its address and drops its holder', async () => {
    const session = await EditorDocument.open(await fixture('invoice2013.docx'));
    const slice = new Slice(session.doc.content, 0, 0);
    const controls = [...nodesOf(slice, 'sdtBlock'), ...nodesOf(slice, 'sdtInline')].filter((n) => n.attrs['sdtPr']);
    expect(controls.length).toBeGreaterThan(0);
    let next = 9000;
    const out = pastedSlice(slice, { ...OFF, controlId: () => next++ });
    const after = [...nodesOf(out, 'sdtBlock'), ...nodesOf(out, 'sdtInline')].filter((n) => n.attrs['sdtPr']);
    expect(after.length).toBe(controls.length);
    const ids = after.map((n) => readControl(n.attrs['sdtPr']).id);
    expect(new Set(ids).size).toBe(ids.length);
    for (const id of ids) expect(id).toBeGreaterThanOrEqual(9000);
    expect(after.map((n) => readControl(n.attrs['sdtPr']).tag)).toEqual(controls.map((n) => readControl(n.attrs['sdtPr']).tag));
    for (const n of after) expect(n.attrs['sourceId']).toBeNull();

    const linked = await EditorDocument.open(await fixture('hyperlink_dupe.docx'));
    const links = pastedSlice(new Slice(linked.doc.content, 0, 0), OFF);
    let seen = 0;
    links.content.descendants((n) => { for (const m of n.marks) if (m.type.name === 'link') { seen++; expect(m.attrs['sourceId']).toBeNull(); expect(m.attrs['rId'] ?? m.attrs['anchor']).toBeTruthy(); } return true; });
    expect(seen).toBeGreaterThan(0);
  });

  test('atoms go to the caller: dropped, replaced or kept', async () => {
    const session = await EditorDocument.open(await fixture('footnote.docx'));
    const slice = new Slice(session.doc.content, 0, 0);
    const refs = nodesOf(slice, 'footnoteRef');
    expect(refs.length).toBeGreaterThan(0);
    const dropped = pastedSlice(slice, { ...OFF, atom: (n) => (n.type.name === 'footnoteRef' ? null : undefined) });
    expect(nodesOf(dropped, 'footnoteRef')).toHaveLength(0);
    const replaced = pastedSlice(slice, { ...OFF, atom: (n) => (n.type.name === 'footnoteRef' ? n.type.create({ ...n.attrs, footnoteId: 77 }) : undefined) });
    expect(nodesOf(replaced, 'footnoteRef').map((n) => n.attrs['footnoteId'])).toEqual(refs.map(() => 77));
    const kept = pastedSlice(slice, OFF);
    expect(nodesOf(kept, 'footnoteRef').map((n) => n.attrs['footnoteId'])).toEqual(refs.map((n) => n.attrs['footnoteId']));
  });
});

describe('a selection that took the paragraph mark pastes whole paragraphs (Jason, 2026-10-29)', () => {
  /** The selection bar's and Shift+Down's shape: from the paragraph's content start to the next paragraph's. */
  function markedSelection(session: EditorDocument, i: number): { from: number; to: number } {
    let pos = 0;
    for (let k = 0; k < i; k++) pos += session.doc.child(k).nodeSize;
    return { from: pos + 1, to: pos + session.doc.child(i).nodeSize + 1 };
  }

  test('the slice ends with an empty open paragraph and is closed with the mark dropped', async () => {
    const session = await EditorDocument.open(await fixture('numbering-clauses.docx'));
    const { from, to } = markedSelection(session, 1);
    const slice = session.doc.slice(from, to);
    expect([slice.openStart, slice.openEnd, slice.content.childCount]).toEqual([1, 1, 2]);
    const closed = closedAtMark(slice);
    expect([closed.openStart, closed.openEnd, closed.content.childCount]).toEqual([0, 0, 1]);
    expect(closed.content.firstChild!.attrs['pStyle']).toBe(session.doc.child(1).attrs['pStyle']);
    // A slice without the mark (to the end of the paragraph's text) is left as it is.
    const plain = session.doc.slice(from, to - 2);
    expect(closedAtMark(plain)).toBe(plain);
    // A slice of text across two paragraphs' middles is left as it is.
    const across = session.doc.slice(from + 2, to + 2);
    expect(closedAtMark(across)).toBe(across);
  });

  test('cut from the selection bar and pasted into an empty paragraph, the paragraph keeps its style', async () => {
    const session = await EditorDocument.open(await fixture('numbering-clauses.docx'));
    const style = session.doc.child(1).attrs['pStyle'];
    expect(style).toBeTruthy();
    const { from, to } = markedSelection(session, 1);
    const state0 = stateFor(session);
    const out = pastedSlice(closedAtMark(state0.doc.slice(from, to)), OFF);
    let tr = state0.tr.setSelection(TextSelection.create(state0.doc, from, to)).deleteSelection();
    tr = tr.insert(tr.doc.content.size, state0.schema.nodes['paragraph']!.create());
    const caret = tr.doc.content.size - 1;
    tr = tr.setSelection(TextSelection.create(tr.doc, caret)).replaceSelection(out);
    const doc = state0.apply(tr).doc;
    const pasted = doc.child(doc.childCount - 1);
    expect(pasted.attrs['pStyle']).toBe(style);
    expect(pasted.attrs['copiedFrom']).toBe(session.doc.child(1).attrs['sourceId']);
    expect(pasted.textContent).toBe(session.doc.child(1).textContent);
    // The empty target paragraph was replaced, not kept before the pasted one.
    expect(doc.childCount).toBe(session.doc.childCount);
  });
});

describe('a bookmark comes across when its name is free (a cut), and is dropped when taken (a copy)', () => {
  function bookmarkAtoms(root: PMNode | Slice): PMNode[] {
    return nodesOf(root, 'opaqueInline').filter((n) => n.attrs['kind'] === 'bookmarkStart' || n.attrs['kind'] === 'bookmarkEnd');
  }
  /** A `bookmarkOf` over the fixture's elements, `taken` as the test says. */
  function rule(session: EditorDocument, taken: (name: string) => boolean): NonNullable<import('../src/index.mjs').PasteRules['bookmarkOf']> {
    return (node) => {
      const v = session.ctx.ids.get(String(node.attrs['sourceId']))?.value as { id?: number | string; name?: string } | undefined;
      if (!v || v.id === undefined) return undefined;
      return { id: String(v.id), name: v.name ?? '', taken: v.name !== undefined && taken(v.name) };
    };
  }

  test('both markers in the slice', async () => {
    const session = await EditorDocument.open(await fixture('numbering-clauses.docx'));
    const slice = new Slice(session.doc.content.cut(0, 400), 0, 0);
    const inSlice = bookmarkAtoms(slice);
    expect(inSlice.length).toBeGreaterThan(0);
    expect(bookmarkAtoms(pastedSlice(slice, OFF))).toHaveLength(0);
    expect(bookmarkAtoms(pastedSlice(slice, { ...OFF, bookmarkOf: rule(session, () => true) }))).toHaveLength(0);
    const kept = bookmarkAtoms(pastedSlice(slice, { ...OFF, bookmarkOf: rule(session, () => false) }));
    // every pair whose start and end are both in the slice, by identity
    const ids = (nodes: PMNode[], kind: string) => new Set(nodes.filter((n) => n.attrs['kind'] === kind).map((n) => String((session.ctx.ids.get(String(n.attrs['sourceId']))!.value as { id: number }).id)));
    const starts = ids(inSlice, 'bookmarkStart');
    const ends = ids(inSlice, 'bookmarkEnd');
    const paired = [...starts].filter((id) => ends.has(id));
    expect(paired.length).toBeGreaterThan(0);
    expect([...ids(kept, 'bookmarkStart')].sort()).toEqual(paired.sort());
    expect([...ids(kept, 'bookmarkEnd')].sort()).toEqual(paired.sort());
    // kept by identity of their source: the same elements go out
    const sources = new Set(inSlice.map((n) => n.attrs['sourceId']));
    for (const n of kept) expect(sources.has(n.attrs['sourceId'])).toBe(true);
  });

  test('a start whose end is outside the slice is dropped', async () => {
    const session = await EditorDocument.open(await fixture('numbering-clauses.docx'));
    // cut the slice just after the first bookmarkStart
    let cutAt = -1;
    session.doc.descendants((n, pos) => { if (cutAt < 0 && n.type.name === 'opaqueInline' && n.attrs['kind'] === 'bookmarkStart') cutAt = pos + n.nodeSize; return cutAt < 0; });
    expect(cutAt).toBeGreaterThan(0);
    const slice = session.doc.slice(0, cutAt);
    expect(bookmarkAtoms(slice).some((n) => n.attrs['kind'] === 'bookmarkStart')).toBe(true);
    expect(bookmarkAtoms(pastedSlice(slice, { ...OFF, bookmarkOf: rule(session, () => false) }))).toHaveLength(0);
  });
});
