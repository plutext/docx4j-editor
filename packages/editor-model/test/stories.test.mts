// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 7.7 step E1: footnotes, comments and headers as read-only projections, read
// privately, so the untouched round trip holds after every read.
import { describe, test, expect } from 'vitest';
import { EditorDocument } from '../src/index.mjs';
import { fixture, comparePackages, textOf } from './helpers.mjs';

describe('footnotes', () => {
  test('notes come in reference order, projected with the body schema; the separators never; the round trip holds', async () => {
    const original = await fixture('footnote.docx');
    const session = await EditorDocument.open(original);
    const notes = await session.stories.footnotes();
    expect(notes.length).toBeGreaterThan(0);
    const refs: number[] = [];
    session.doc.descendants((n) => { if (n.type.name === 'footnoteRef' && n.attrs['kind'] === 'footnote') refs.push(Number(n.attrs['footnoteId'])); return true; });
    expect(notes.filter((n) => !n.orphan).map((n) => n.id)).toEqual([...new Set(refs)]);
    for (const n of notes) {
      expect(n.id).toBeGreaterThan(0);
      expect(n.doc.type.name).toBe('doc');
      expect(n.doc.childCount).toBeGreaterThan(0);
      expect(n.doc.firstChild!.type.name).toBe('paragraph');
      // The note's own `w:footnoteRef` mark is an opaque inline ahead of the text; an orphan may hold nothing else.
      if (!n.orphan) {
        expect(textOf(n.doc.firstChild!).length).toBeGreaterThan(0);
        expect(session.ctx.ids.get(session.stories.referenceOf(session.doc, 'footnote', n.id)!)).toBeDefined();
      }
    }
    expect(await session.stories.footnotes()).toBe(notes);   // cached
    const saved = await session.save();
    expect(comparePackages(original, saved.bytes)).toEqual([]);
  });

  test('a document without notes gives none', async () => {
    const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
    expect(await session.stories.footnotes()).toEqual([]);
    expect(await session.stories.endnotes()).toEqual([]);
  });
});

describe('comments', () => {
  test('threads with author, date and text; the ids match the body marks; the round trip holds', async () => {
    const original = await fixture('comments-two.docx');
    const session = await EditorDocument.open(original);
    const threads = await session.stories.comments();
    expect(threads.length).toBeGreaterThan(0);
    const ids = new Set<string>();
    session.doc.descendants((n) => { for (const m of n.marks) if (m.type.name === 'comment') ids.add(String(m.attrs['id'])); return true; });
    const all = threads.flatMap((t) => [t, ...t.replies]);
    expect(all.filter((t) => !t.orphan).map((t) => t.id).sort()).toEqual([...ids].sort());
    for (const t of all) {
      expect(t.author).not.toBe('');
      expect(t.doc.childCount).toBeGreaterThan(0);
      expect(textOf(t.doc.firstChild!).length).toBeGreaterThan(0);
    }
    const saved = await session.save();
    expect(comparePackages(original, saved.bytes)).toEqual([]);
  });
});

describe('headers and footers', () => {
  test('every header and footer part projects; the variant rules pick first, even and default; the round trip holds', async () => {
    const original = await fixture('headers.docx');   // six header parts over five sections
    const session = await EditorDocument.open(original);
    const parts = await session.stories.headers();
    expect(parts.length).toBe(6);
    for (const h of parts) { expect(['header', 'footer']).toContain(h.kind); expect(h.rId).toMatch(/^rId/); expect(h.doc.childCount).toBeGreaterThan(0); }
    const sections = session.stories.sections();
    expect(sections.length).toBeGreaterThan(0);
    const first = await session.stories.headerFooterFor(0, 1, true);
    const second = await session.stories.headerFooterFor(0, 2, false);
    const third = await session.stories.headerFooterFor(0, 3, false);
    expect(['first', 'even', 'default']).toContain(first.variant);
    expect(second.variant).not.toBe('first');
    expect(third.variant).toBe('default');
    // Whatever the variants, a page in a section with references gets a header.
    expect(first.header ?? second.header ?? third.header).toBeDefined();
    const saved = await session.save();
    expect(comparePackages(original, saved.bytes)).toEqual([]);
  });

  test('odd, even and first pages of a title-page document differ', async () => {
    const session = await EditorDocument.open(await fixture('odd-even-first-page.docx'));
    const p1 = await session.stories.headerFooterFor(0, 1, true);
    const p2 = await session.stories.headerFooterFor(0, 2, false);
    const p3 = await session.stories.headerFooterFor(0, 3, false);
    expect(p1.variant).toBe('first');
    expect(p2.variant).toBe('even');
    expect(p3.variant).toBe('default');
    expect(p1.header?.rId).not.toBe(p2.header?.rId);
    expect(p2.header?.rId).not.toBe(p3.header?.rId);
  });

  test('same as previous: a later section without references inherits', async () => {
    const session = await EditorDocument.open(await fixture('sections-multiple.docx'));
    const sections = session.stories.sections();
    expect(sections.length).toBeGreaterThan(1);
    const last = sections.length - 1;
    const early = await session.stories.headerFooterFor(0, 1, true);
    const late = await session.stories.headerFooterFor(last, 5, false);
    if (early.header || early.footer) expect(late.header ?? late.footer).toBeDefined();
  });
});
