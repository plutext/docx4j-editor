// SPDX-License-Identifier: GPL-3.0-or-later
//
// A document's text against another's (ED-011 section 3.4, step S2): `compareText` over lists of strings,
// and the two readers over fixtures, the engine's text model on both sides.
import { describe, test, expect } from 'vitest';
import { EditorDocument, compareText, paragraphTextsOfBytes, paragraphTextsOfSession, compareWithBytes } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

function counts(a: string[], b: string[]): { changed: number; added: number; removed: number } {
  const { changed, added, removed } = compareText(a, b);
  return { changed, added, removed };
}

describe('compareText', () => {
  test('a rewritten file costs no search: it takes seconds without the strings left out, and this bound has room on any machine', () => {
    const a = Array.from({ length: 20_000 }, (_, i) => `before ${i}`);
    const b = Array.from({ length: 20_000 }, (_, i) => `after ${i}`);
    const start = performance.now();
    expect(compareText(a, b).changed).toBe(20_000);
    expect(performance.now() - start).toBeLessThan(1_000);
  });

  test('finds a longest match over any lists: strings the other list lacks are left out of the search, never out of the answer', () => {
    // A longest common subsequence's length, the plain way, as the oracle.
    const longest = (a: string[], b: string[]): number => {
      let row = new Array<number>(b.length + 1).fill(0);
      for (const s of a) {
        const next = [0];
        for (let j = 1; j <= b.length; j++) next[j] = s === b[j - 1] ? row[j - 1]! + 1 : Math.max(row[j]!, next[j - 1]!);
        row = next;
      }
      return row[b.length]!;
    };
    let seed = 20261107;
    const random = (n: number): number => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % n; };
    for (let run = 0; run < 400; run++) {
      const alphabet = 2 + random(12);
      const list = (): string[] => Array.from({ length: random(30) }, () => `s${random(alphabet)}`);
      const a = list();
      const b = list();
      const { pairs, changed, added, removed } = compareText(a, b);
      const equal = pairs.filter((p) => p.kind === 'equal');
      expect(equal.length).toBe(longest(a, b));
      for (const p of equal) expect(a[p.a!]).toBe(b[p.b!]);
      // Every paragraph of each side stands once, in order.
      expect(pairs.flatMap((p) => (p.a === undefined ? [] : [p.a]))).toEqual(a.map((_, i) => i));
      expect(pairs.flatMap((p) => (p.b === undefined ? [] : [p.b]))).toEqual(b.map((_, j) => j));
      expect(equal.length + changed + removed).toBe(a.length);
      expect(equal.length + changed + added).toBe(b.length);
    }
  });

  test('equal lists align pair by pair with zero counts', () => {
    const c = compareText(['a', 'b', 'c'], ['a', 'b', 'c']);
    expect(c.pairs).toEqual([{ kind: 'equal', a: 0, b: 0 }, { kind: 'equal', a: 1, b: 1 }, { kind: 'equal', a: 2, b: 2 }]);
    expect(counts(['a', 'b', 'c'], ['a', 'b', 'c'])).toEqual({ changed: 0, added: 0, removed: 0 });
  });

  test('one paragraph changed is one changed pair', () => {
    const c = compareText(['a', 'b', 'c'], ['a', 'B', 'c']);
    expect(c.pairs).toEqual([{ kind: 'equal', a: 0, b: 0 }, { kind: 'changed', a: 1, b: 1 }, { kind: 'equal', a: 2, b: 2 }]);
    expect(counts(['a', 'b', 'c'], ['a', 'B', 'c'])).toEqual({ changed: 1, added: 0, removed: 0 });
  });

  test('one added, one removed', () => {
    expect(compareText(['a', 'c'], ['a', 'b', 'c']).pairs).toEqual([{ kind: 'equal', a: 0, b: 0 }, { kind: 'added', b: 1 }, { kind: 'equal', a: 1, b: 2 }]);
    expect(counts(['a', 'c'], ['a', 'b', 'c'])).toEqual({ changed: 0, added: 1, removed: 0 });
    expect(compareText(['a', 'b', 'c'], ['a', 'c']).pairs).toEqual([{ kind: 'equal', a: 0, b: 0 }, { kind: 'removed', a: 1 }, { kind: 'equal', a: 2, b: 1 }]);
    expect(counts(['a', 'b', 'c'], ['a', 'c'])).toEqual({ changed: 0, added: 0, removed: 1 });
  });

  test('a moved paragraph is one removed and one added', () => {
    expect(counts(['x', 'y', 'z'], ['y', 'z', 'x'])).toEqual({ changed: 0, added: 1, removed: 1 });
  });

  test('a block replaced by a longer one pairs in order, the rest added', () => {
    const c = compareText(['a', 'b', 'c', 'd'], ['a', 'B', 'C', 'C2', 'd']);
    expect(c.pairs).toEqual([
      { kind: 'equal', a: 0, b: 0 }, { kind: 'changed', a: 1, b: 1 }, { kind: 'changed', a: 2, b: 2 }, { kind: 'added', b: 3 }, { kind: 'equal', a: 3, b: 4 },
    ]);
    expect(counts(['a', 'b', 'c', 'd'], ['a', 'B', 'C', 'C2', 'd'])).toEqual({ changed: 2, added: 1, removed: 0 });
  });

  test('empty lists', () => {
    expect(compareText([], []).pairs).toEqual([]);
    expect(counts([], ['a'])).toEqual({ changed: 0, added: 1, removed: 0 });
    expect(counts(['a'], [])).toEqual({ changed: 0, added: 0, removed: 1 });
    expect(counts(['a', 'b'], ['c', 'd', 'e'])).toEqual({ changed: 2, added: 1, removed: 0 });
  });

  test('the alignment is a valid edit script over longer lists with repeats', () => {
    const a = 'the quick brown fox jumps over the lazy dog the end'.split(' ');
    const b = 'the slow brown fox leaps over the dog and the cat the end'.split(' ');
    const c = compareText(a, b);
    // Every index of each side appears exactly once, in order.
    expect(c.pairs.filter((p) => p.a !== undefined).map((p) => p.a)).toEqual(a.map((_, i) => i));
    expect(c.pairs.filter((p) => p.b !== undefined).map((p) => p.b)).toEqual(b.map((_, i) => i));
    for (const p of c.pairs) if (p.kind === 'equal') expect(a[p.a!]).toBe(b[p.b!]);
    // The common subsequence is the longest: the, brown, fox, over, the, dog, the, end (8).
    expect(c.pairs.filter((p) => p.kind === 'equal').length).toBe(8);
  });
});

describe('the readers (the engine\'s text on both sides)', () => {
  test('a fixture against itself gives zero counts, the live side through readTree', async () => {
    const bytes = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(bytes);
    const c = await compareWithBytes(session, bytes);
    expect({ changed: c.changed, added: c.added, removed: c.removed }).toEqual({ changed: 0, added: 0, removed: 0 });
    expect(c.pairs.length).toBeGreaterThan(0);
    // A tracked deletion reads out on both sides: the accepted view.
    const live = await paragraphTextsOfSession(session);
    expect(live).toEqual(await paragraphTextsOfBytes(bytes));
    expect(live.join('\n')).not.toContain('deleted');
    // The read wrote nothing: an unedited save is byte for byte.
    expect((await session.save()).changed).toBe(false);
  });

  test('a copy with one paragraph edited gives one changed', async () => {
    const bytes = await fixture('HelloWordOnline.docx');
    const session = await EditorDocument.open(bytes);
    await session.runOnTree(({ body }) => { body.paragraphs[0]!.insertText(' (edited)', 'End'); });
    const edited = (await session.save()).bytes;
    const fresh = await EditorDocument.open(bytes);
    const c = await compareWithBytes(fresh, edited);
    expect({ changed: c.changed, added: c.added, removed: c.removed }).toEqual({ changed: 1, added: 0, removed: 0 });
    // And the other way, the edited session against the original file.
    const back = await compareWithBytes(session, bytes);
    expect({ changed: back.changed, added: back.added, removed: back.removed }).toEqual({ changed: 1, added: 0, removed: 0 });
  });

  test('a kept mc:AlternateContent reads as no text on both sides', async () => {
    const bytes = await fixture('mc/mc-branches.docx');
    const session = await EditorDocument.open(bytes);
    const c = await compareWithBytes(session, bytes);
    expect({ changed: c.changed, added: c.added, removed: c.removed }).toEqual({ changed: 0, added: 0, removed: 0 });
  });

  test('two different fixtures', async () => {
    const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
    const c = await compareWithBytes(session, await fixture('tables.docx'));
    const live = await paragraphTextsOfSession(session);
    const disk = await paragraphTextsOfBytes(await fixture('tables.docx'));
    expect(c.changed + c.removed).toBeLessThanOrEqual(live.length);
    expect(c.changed + c.added).toBe(disk.length - c.pairs.filter((p) => p.kind === 'equal').length);
    console.log(`HelloWordOnline (${live.length}) against tables.docx (${disk.length}): ${c.changed} changed, ${c.added} added, ${c.removed} removed`);
  });
});
