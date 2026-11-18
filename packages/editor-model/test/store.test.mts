// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 5.6 step C4: adopting a persisted store, which relies on sourceIds being
// deterministic for the same bytes.
import { describe, test, expect } from 'vitest';
import * as Y from 'yjs';
import { EditorDocument, newEditStore, runCommand, yAttributeKey, editScript, trackedChangesOf } from '../src/index.mjs';
import { fixture, comparePackages, allFixtures } from './helpers.mjs';

describe('deterministic ids', () => {
  test('the same bytes give the same sourceIds and the same document', async () => {
    for (const name of (await allFixtures()).slice(0, 8)) {
      const bytes = await fixture(name);
      const a = await EditorDocument.open(bytes, { resolver: false });
      const b = await EditorDocument.open(bytes, { resolver: false });
      expect(a.imported.eq(b.imported)).toBe(true);
      expect(a.ctx.ids.size).toBe(b.ctx.ids.size);
    }
  });
});

describe('OpenOptions.store', () => {
  test('an empty store is loaded; a store with content is adopted, and edits in it export', async () => {
    const original = await fixture('tracked-changes.docx');
    const first = await EditorDocument.open(original);
    let pos = 0;
    for (let k = 0; k < 4; k++) pos += first.doc.child(k).nodeSize;
    runCommand(first, 'insertText', { text: ' (recovered)' }, { from: pos + 1 + 10 });
    // Persist the store as y-indexeddb would: the update log.
    const update = Y.encodeStateAsUpdate(first.store.ydoc);

    // A new session over the same bytes adopting the persisted store.
    const ydoc = new Y.Doc();
    Y.applyUpdate(ydoc, update);
    const store = newEditStore(ydoc);
    expect(store.body.length).toBeGreaterThan(0);
    const second = await EditorDocument.open(original, { store });
    expect(second.doc.child(4).textContent).toBe('Some text. (recovered)');
    expect(second.doc.eq(first.doc)).toBe(true);
    const saved = await second.save();
    expect(saved.changed).toBe(true);
    expect(comparePackages(original, saved.bytes, [second.main.partName.storeName, 'docProps/app.xml']).filter((d) => d.name !== 'docProps/app.xml')).toEqual([]);
    const reopened = await EditorDocument.open(saved.bytes);
    expect(reopened.doc.child(4).textContent).toBe('Some text. (recovered)');

    // An empty store given explicitly is loaded as usual.
    const empty = newEditStore();
    const third = await EditorDocument.open(original, { store: empty });
    expect(third.doc.eq(first.imported)).toBe(true);
  });

  test('setOptions turns the app.xml write off after open', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    session.setOptions({ writeApplication: false });
    runCommand(session, 'paragraph.alignment', { value: 'Centered' }, { from: 1 });
    const saved = await session.save();
    expect(saved.changed).toBe(true);
    expect(comparePackages(original, saved.bytes, [session.main.partName.storeName])).toEqual([]);
  });
});

describe('the Yjs attribute key of a mark (ED-005 section 12.26, UPSTREAM-12)', () => {
  test('yAttributeKey is the key y-prosemirror stores each revision mark by: a deletion and a comment by a hash, an insertion by its name', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const schema = session.doc.type.schema;
    const marks = [
      schema.marks['deletion']!.create({ author: 'Author A', date: '2026-09-29T01:00:00Z', move: null }),
      schema.marks['deletion']!.create({ author: 'Author B', date: '2026-09-29T02:00:00Z', move: null }),
      schema.marks['insertion']!.create({ author: 'Author C', date: '2026-09-29T03:00:00Z', move: null }),
      schema.marks['comment']!.create({ id: '7' }),
    ];
    const p = session.doc.child(4);
    const text = schema.text('both', marks);
    session.replace(session.doc.copy(session.doc.content.replaceChild(4, p.type.create(p.attrs, text))));
    let keys: string[] = [];
    const visit = (t: Y.XmlElement | Y.XmlFragment | Y.XmlText): void => {
      if (t instanceof Y.XmlText) {
        for (const op of t.toDelta() as { insert: unknown; attributes?: Record<string, unknown> }[]) if (op.insert === 'both') keys = Object.keys(op.attributes ?? {}).sort();
        return;
      }
      for (const child of t.toArray()) visit(child as Y.XmlElement | Y.XmlText);
    };
    visit(session.store.body);
    expect(keys).toEqual(marks.map((m) => yAttributeKey(m)).sort());
    expect(keys.filter((k) => k.startsWith('deletion--')).length).toBe(2);
    expect(keys).toContain('insertion');
  });
});

// UPSTREAM-16 (ED-005 section 12.47): before y-prosemirror's diff, the store's children are matched to the new ones by
// the shortest edit script over their kinds and texts, so the diff updates what is kept in place instead of putting a
// copy of a neighbour where a removed child was.
describe('the store\'s children matched to the new ones before the diff (UPSTREAM-16)', () => {
  /** The length of the longest common subsequence, by the textbook table. */
  function lcs(a: string[], b: string[]): number {
    const t = Array.from({ length: a.length + 1 }, () => new Array<number>(b.length + 1).fill(0));
    for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++) t[i]![j] = a[i - 1] === b[j - 1] ? t[i - 1]![j - 1]! + 1 : Math.max(t[i - 1]![j]!, t[i]![j - 1]!);
    return t[a.length]![b.length]!;
  }

  test('the edit script turns one sequence into the other in the fewest edits', () => {
    let seed = 7;
    const random = (): number => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
    for (let round = 0; round < 500; round++) {
      const a = Array.from({ length: Math.floor(random() * 12) }, () => 'pqt'[Math.floor(random() * 3)]!);
      const b = Array.from({ length: Math.floor(random() * 12) }, () => 'pqt'[Math.floor(random() * 3)]!);
      const ops = editScript(a, b)!;
      const out: string[] = [];
      let i = 0;
      let j = 0;
      for (const op of ops) {
        if (op === 0) { expect(a[i]).toBe(b[j]); out.push(a[i]!); i++; j++; } else if (op === 1) i++; else { out.push(b[j]!); j++; }
      }
      expect([i, j]).toEqual([a.length, b.length]);
      expect(out).toEqual(b);
      expect(ops.filter((op) => op !== 0).length).toBe(a.length + b.length - 2 * lcs(a, b));
    }
    expect(editScript(['p', 'p', 'p'], ['q', 'q', 'q'], 5)).toBeUndefined();
  });

  test('a move accepted alone: the kept paragraph and the moved paragraph\'s text are the store\'s own, never copies', async () => {
    const session = await EditorDocument.open(await fixture('revisions/revisions-word15.docx'));
    const move = trackedChangesOf(session.doc).find((c) => c.move !== undefined && c.type === 'Added')!;
    const texts = (): string[] => session.store.body.toArray().map((y) => (y instanceof Y.XmlElement ? y.toString().replace(/<[^>]*>/g, '') : ''));
    const kept = session.store.body.get(texts().indexOf('The first paragraph of the move: it stays where it is.')) as Y.XmlElement;
    const moved = (session.store.body.get(texts().findIndex((t) => t.startsWith('The second paragraph'))) as Y.XmlElement).toArray().find((c) => c instanceof Y.XmlText)!;
    expect(runCommand(session, 'trackedChange.accept', { key: session.changeKey(move) }).ok).toBe(true);
    const after = session.store.body.toArray();
    expect(after).toContain(kept);
    const holder = after.find((y) => y instanceof Y.XmlElement && y.toArray().includes(moved as Y.XmlText));
    expect(holder).toBeDefined();
    expect(texts().filter((t) => t.startsWith('The second paragraph'))).toHaveLength(1);
  });
});
