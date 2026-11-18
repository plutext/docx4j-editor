// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 3.2: the editor's find holds to the engine's `body.search` over every fixture,
// for a fixed set of queries (Word's options, wildcards, regular expressions, Hebrew, Arabic,
// Japanese) and queries drawn from each fixture's own text (its words in other cases, a word inside
// a longer one, a phrase across a run boundary). A difference is a finding recorded in ED-005
// section 12 and fixed, never a tolerance here (section 2.1).
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, Finder, findIn, atomText, paragraphText, textOffsetIn, hitSelection, runCommand, schema, stateFor, type FindOptions } from '../src/index.mjs';
import type { Node as PMNode } from 'prosemirror-model';
import { fixture, allFixtures, textOf } from './helpers.mjs';

type Query = [string, FindOptions];

const FIXED: Query[] = [
  ['the', {}], ['The', { matchCase: true }], ['the', { matchWholeWord: true }], ['THE', { matchCase: true }], ['e', {}], ['in', { matchWholeWord: true }],
  ['?e', { matchWildcards: true }], ['t*e', { matchWildcards: true }], ['[a-z]@', { matchWildcards: true }], ['[!a-z ]', { matchWildcards: true }],
  ['<th', { matchWildcards: true }], ['ing>', { matchWildcards: true }], ['<[A-Z][a-z]@>', { matchWildcards: true, matchCase: true }],
  ['(<[a-z]@>) (<[a-z]@>)', { matchWildcards: true }], ['[0-9]{2}', { matchWildcards: true }],
  ['\\b\\w+ing\\b', { matchRegExp: true }], ['(\\w+) (\\w+)', { matchRegExp: true }], ['\\d+', { matchRegExp: true }], ['\\t', { matchRegExp: true }],
  ['\\p{L}+', { matchRegExp: true }], ['[^\\x00-\\x7f]+', { matchRegExp: true }], ['text|table', { matchRegExp: true, matchWholeWord: true }],
  ['^t', {}], ['.', {}], ['(', {}], ['a b', {}],
  ['שלום', {}], ['זהו משפט', {}], ['مرحبا', {}], ['نص تجريبي', { matchWholeWord: true }], ['日本語', {}], ['テキスト日本', {}], ['の', {}],
  ['😀', {}], ['→', {}],
  // The accepted view (section 3.2, step C5): across a deletion and over deleted text, from the Word files and tracked-changes.docx.
  ['near the stone bridge.', {}], ['old stone', {}], ['bridge at dawn', {}], ['the field', { matchWholeWord: true }], ['green field', {}],
  ['Followed by.', {}], ['A deletion', {}], ['the*bridge', { matchWildcards: true }], ['by[.]', { matchWildcards: true }], ['the\\s+stone', { matchRegExp: true }],
  // Across paragraph marks (ED-008 section 8.8 item 11): Word's `^p` in a plain query, and the wildcards that take a mark (check 59).
  ['^p', {}], ['.^p', {}], ['^p^p', {}], ['e^pT', {}], ['^pThe', {}], ['s.^p', { matchWholeWord: true }], ['^13', { matchWildcards: true }], ['e^13?', { matchWildcards: true }],
  ['[!a-z ]^13', { matchWildcards: true }], ['.*e', { matchWildcards: true }], ['?^13?', { matchWildcards: true }], ['[.!?]^13[A-Z]', { matchWildcards: true }],
];

/** Queries from the fixture's own text: a few words as they are, upper-cased with and without match case, their middles with and without whole words, and two words together. */
function derived(text: string): Query[] {
  const words = [...new Set(text.match(/[\p{L}\p{N}]{4,}/gu) ?? [])];
  const picked = [words[0], words[Math.floor(words.length / 2)], words[words.length - 1]].filter((w): w is string => w !== undefined);
  const out: Query[] = [];
  for (const w of new Set(picked)) {
    out.push([w, {}], [w.toUpperCase(), {}], [w.toUpperCase(), { matchCase: true }], [w.slice(1, -1), {}], [w.slice(1, -1), { matchWholeWord: true }]);
  }
  const pair = /([\p{L}]{3,}) ([\p{L}]{3,})/u.exec(text);
  if (pair) out.push([pair[0], {}]);
  return out;
}

const boundaryCache = new Map<string, Set<number>>();

/** Whether an offset of a text is on a grapheme boundary. */
function onBoundary(text: string, offset: number): boolean {
  let bounds = boundaryCache.get(text);
  if (!bounds) {
    bounds = new Set<number>([text.length]);
    for (const { index } of new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text)) bounds.add(index);
    boundaryCache.set(text, bounds);
  }
  return bounds.has(offset);
}

/**
 * The engine's hits, as text, or its error's kind. A hit that splits a grapheme cluster (half an
 * emoji's surrogate pair under wildcards, an Arabic letter without its combining mark under
 * `\p{L}+`) is never the editor's (ED-001 section 3.7). Up to core-ts 0.2.1 the engine produced
 * 12 over the fixtures (ED-005 section 12.1 item 2); 0.2.2 drops them as the editor does (section
 * 1.7). One would be left out of the engine's list here and counted, and the count is pinned at 0.
 */
function engineHits(body: Awaited<ReturnType<WordprocessingMLPackage['getBody']>>, [q, o]: Query, split: { count: number }): string[] | 'error' {
  try {
    // `Body.search` whole: its wildcards run across paragraph marks (check 59, CR-002 section 32 K7) and a plain `^p`
    // matches one (CR-014 phase 3), as the finder's do since ED-008 section 8.8 item 11; a hit's end is in its last
    // paragraph (`endParagraph`, where it spans).
    return body.search(q, o).filter((r) => {
      const last = (r as unknown as { endParagraph?: { text: string } }).endParagraph ?? r.paragraph;
      const ok = onBoundary(r.paragraph.text, r.start) && onBoundary(last.text, r.end);
      if (!ok) split.count++;
      return ok;
    }).map((r) => r.text);
  } catch { return 'error'; }
}

function editorHits(session: EditorDocument, [q, o]: Query): string[] | 'error' {
  let finder: Finder;
  try { finder = new Finder(q, o, session.ctx.ids); } catch { return 'error'; }
  return finder.all(session.doc).map((h) => h.text);
}

/**
 * Where the engine's `Body.search` differs from the measured rule and the editor keeps the rule: a comparison here must
 * differ while the engine does, and agree (the entry then goes) once a release fixes it. Empty since core-ts 0.7.0, which
 * fixed the two reported 2026-10-26 (ED-008 section 8.8 item 11; core-ts 7f00fad, CR-002 section 64): a moved-from mark
 * read as `\r` (nine queries over the two Word revision files) and a wildcard hit across a table's edge in invoice2013.
 */
const KNOWN_ENGINE: { fixture: RegExp; queries: string[]; reason: string }[] = [];

describe('find agrees with the engine (ED-005 section 3.2)', () => {
  test('every fixture, every query: the same hits in the same order', async () => {
    const findings: string[] = [];
    let compared = 0;
    let hits = 0;
    let known = 0;
    const split = { count: 0 };
    for (const name of await allFixtures()) {
      const bytes = await fixture(name);
      const session = await EditorDocument.open(bytes);
      const pkg = await WordprocessingMLPackage.load(bytes);
      const body = await pkg.getBody();
      for (const query of [...FIXED, ...derived(body.text)]) {
        const theirs = engineHits(body, query, split);
        const ours = editorHits(session, query);
        compared++;
        if (Array.isArray(theirs)) hits += theirs.length;
        const pinned = KNOWN_ENGINE.find((k) => k.fixture.test(name) && k.queries.includes(query[0]));
        if (pinned) {
          if (JSON.stringify(theirs) === JSON.stringify(ours)) findings.push(`${name} ${JSON.stringify(query)}: the known engine difference (${pinned.reason}) no longer shows; drop its pin`);
          else known++;
          continue;
        }
        if (JSON.stringify(theirs) !== JSON.stringify(ours)) {
          const [a, b] = [Array.isArray(theirs) ? theirs : [theirs], Array.isArray(ours) ? ours : [ours]];
          let i = 0;
          while (i < a.length && a[i] === b[i]) i++;
          findings.push(`${name} ${JSON.stringify(query)}: engine ${a.length} hits, editor ${b.length}; first difference at ${i}: engine ${JSON.stringify(a.slice(i, i + 3))} editor ${JSON.stringify(b.slice(i, i + 3))}`);
        }
      }
    }
    expect(compared).toBeGreaterThan(1000);
    expect(hits).toBeGreaterThan(10000);
    expect(findings).toEqual([]);
    expect(known).toBe(0);
    // The engine's hits inside a cluster: 12 in fonts-symbol-and-emoji.docx and loadAndSave.docx up to core-ts 0.2.1 (section 12.1 item 2).
    expect(split.count).toBe(SPLIT_BY_ENGINE);
  });
});

describe('the accepted view (ED-005 section 3.2, step C5)', () => {
  // The agreement above holds only if both sides skip deleted text; this says what skipping is, over the Word files and ours.
  test('a hit reads across a deletion and never covers deleted text', async () => {
    const hits = async (name: string, q: string): Promise<string[]> => {
      const session = await EditorDocument.open(await fixture(name));
      return findIn(session.doc, q, {}, session.ctx.ids).map((h) => h.text);
    };
    for (const name of ['revisions/revisions-word2010.docx', 'revisions/revisions-word15.docx']) {
      expect(await hits(name, 'near the stone bridge.')).toEqual(['near the stone bridge.']);
      expect(await hits(name, 'old stone')).toEqual([]);
      expect(await hits(name, 'bridge at dawn')).toEqual([]);   // The instructions quote " at dawn".
      expect(await hits(name, 'the field beside')).toEqual(['the field beside']);
      expect(await hits(name, 'the green field')).toEqual([]);
      // Inserted text is text: "the very lazy dog" as Word shows it with the insertion.
      expect(await hits(name, 'the very lazy')).toEqual(['the very lazy']);
    }
    expect(await hits('tracked-changes.docx', 'Followed by.')).toEqual(['Followed by.']);
    expect(await hits('tracked-changes.docx', 'A deletion')).toEqual([]);
  });
});

/** How many of the engine's hits over the fixtures split a grapheme cluster: none since core-ts 0.2.2 (ED-005 section 1.7). */
const SPLIT_BY_ENGINE = 0;

describe('hits and positions', () => {
  test('a hit covers its text in the document; atoms without text inside it are covered, not counted', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const doc = session.doc;
    for (const h of findIn(doc, 'e', {}, session.ctx.ids)) expect(doc.textBetween(h.from, h.to)).toBe('e');
    // Block 23: "A short " + two bookmark markers + "para" + "graph with ..." (two runs): the engine reads "short para".
    const [hit] = findIn(doc, 'short para', {}, session.ctx.ids);
    expect(hit?.text).toBe('short para');
    const kinds: string[] = [];
    doc.nodesBetween(hit!.from, hit!.to, (n) => { if (n.type.name === 'opaqueInline') kinds.push(n.attrs['kind'] as string); return true; });
    expect(kinds).toEqual(['bookmarkStart', 'bookmarkEnd']);
    expect(doc.textBetween(hit!.from, hit!.to)).toBe('short para');
  });

  test("an atom's text is Office JS's (core-ts 0.6.0, check 48): any symbol \"(\", the hyphens U+001E and U+001F, a note reference U+0002, the breaks, a picture or a field character nothing; the console's offsets count the same", () => {
    const ids = { get: (id: string) => (id === 's' ? { value: { _char: 'F04A' } } : undefined) } as unknown as Parameters<typeof atomText>[1];
    const sym = schema.nodes['opaqueInline']!.create({ kind: 'sym', sourceId: 's' });
    const nbh = schema.nodes['opaqueInline']!.create({ kind: 'noBreakHyphen' });
    const field = schema.nodes['opaqueInline']!.create({ kind: 'fldChar' });
    expect(atomText(sym, ids)).toBe('(');
    expect(atomText(sym)).toBe('(');
    expect(atomText(nbh)).toBe('\u001E');
    expect(atomText(schema.nodes['opaqueInline']!.create({ kind: 'softHyphen' }))).toBe('\u001F');
    expect(atomText(field, ids)).toBe('');
    expect(atomText(schema.nodes['hardBreak']!.create())).toBe('\u000B');
    expect(atomText(schema.nodes['hardBreak']!.create({ kind: 'page' }))).toBe('\u000C');
    expect(atomText(schema.nodes['footnoteRef']!.create({ footnoteId: 1 }))).toBe('\u0002');
    const paragraph = schema.nodes['paragraph']!.create(null, [schema.text('a'), sym, nbh, field, schema.text('b')]);
    expect(textOffsetIn(paragraph, paragraph.content.size, ids)).toBe(4);
    expect(paragraphText(paragraph, 1, ids).text).toBe('a(\u001Eb');
  });

  test('a tab is U+0009 to a wildcard or an expression; a plain query or an expression never crosses a paragraph', () => {
    const p = (content: PMNode[]): PMNode => schema.nodes['paragraph']!.create(null, content);
    const doc = schema.topNodeType.create(null, [
      p([schema.text('one'), schema.nodes['tab']!.create(), schema.text('two')]),
      p([schema.text('three')]),
    ]);
    expect(findIn(doc, 'one\\ttwo', { matchRegExp: true }).map((h) => h.text)).toEqual(['one\ttwo']);
    expect(findIn(doc, 'two three', {}).length).toBe(0);
    expect(findIn(doc, 'two.three', { matchRegExp: true }).length).toBe(0);
    expect(findIn(doc, 'two[\\s\\S]three', { matchRegExp: true }).length).toBe(0);
  });

  test("a wildcard crosses a paragraph mark and ^13 is one, a plain ^p matches one (check 59, CR-014 phase 3; ED-008 section 8.8 item 11): the hit's text holds \\r, its selection runs into the next paragraph, and it is never replaced", async () => {
    const p = (content: PMNode[], attrs: Record<string, unknown> | null = null): PMNode => schema.nodes['paragraph']!.create(attrs, content);
    const doc = schema.topNodeType.create(null, [
      p([schema.text('one'), schema.nodes['tab']!.create(), schema.text('two')]),
      p([schema.text('three')]),
      p([schema.text('four')]),
    ]);
    // The lazy `*` runs from "o" into the next paragraph where the paragraph alone gave nothing: the engine's K7.
    const across = findIn(doc, 'o*t', { matchWildcards: true });
    expect(across.map((h) => h.text)).toEqual(['one\tt', 'o\rt']);
    expect(across[1]!.mark).toBe(true);
    expect(across[0]!.mark).toBeUndefined();
    expect(doc.textBetween(across[1]!.from, across[1]!.to, '|')).toBe('o|t');
    const sel = hitSelection(doc, across[1]!);
    expect([sel.from, sel.to]).toEqual([across[1]!.from, doc.child(0).nodeSize + 2]);   // into "three"'s "t" exclusive: the next paragraph's start plus one
    expect(doc.textBetween(sel.from, sel.to, '|')).toBe('o|t');
    // `^13` under wildcards and `^p` plain: every paragraph's own mark but, under wildcards, the document's last (the engine's Content range).
    expect(findIn(doc, '^13', { matchWildcards: true }).map((h) => h.text)).toEqual(['\r', '\r']);
    expect(findIn(doc, '^p', {}).map((h) => h.text)).toEqual(['\r', '\r', '\r']);
    expect(findIn(doc, 'two^pthree', {}).map((h) => h.text)).toEqual(['two\rthree']);
    expect(findIn(doc, 'two^pthree^p', {}).map((h) => h.text)).toEqual(['two\rthree\r']);
    expect(findIn(doc, 'o^pthree^pf', {}).map((h) => h.text)).toEqual(['o\rthree\rf']);
    expect(findIn(doc, 'two^pfour', {}).length).toBe(0);
    expect(findIn(doc, '?^13?', { matchWildcards: true }).map((h) => h.text)).toEqual(['o\rt', 'e\rf']);
    // A hit of the mark alone is at the paragraph's end, selected into the next paragraph.
    const bare = findIn(doc, '^p', {})[0]!;
    expect(bare.from).toBe(doc.child(0).nodeSize - 1);
    expect(hitSelection(doc, bare).to).toBe(doc.child(0).nodeSize + 1);
    // A deleted mark reads as nothing and joins its paragraphs (CR-012 section 19): `^p` does not match it, a wildcard runs over it.
    const joined = schema.topNodeType.create(null, [
      p([schema.text('one')], { markRevision: { deletion: { author: 'A', date: '2026-10-26T00:00:00Z' } } }),
      p([schema.text('two')]),
    ]);
    expect(findIn(joined, '^p', {}).map((h) => h.text)).toEqual(['\r']);
    expect(findIn(joined, 'e?t', { matchWildcards: true }).map((h) => h.text)).toEqual([]);
    expect(findIn(joined, 'e*t', { matchWildcards: true }).map((h) => h.text)).toEqual(['et']);
    // Never replaced: the engine's `replaceText` refuses before changing anything, and so do the commands.
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const before = session.doc;
    expect(findIn(session.doc, '^p', {}, session.ctx.ids).length).toBeGreaterThan(10);
    expect(runCommand(session, 'find.replaceAll', { text: '.^p', replacement: 'X' }).ok).toBe(false);
    expect(runCommand(session, 'find.replaceAll', { text: '.*e', replacement: 'X', matchWildcards: true }).ok).toBe(false);
    expect(session.doc.eq(before)).toBe(true);
    // Replace on a selected mark hit: refused, where Replace on a plain hit of the same query replaces it.
    const marked = findIn(session.doc, '?^13?', { matchWildcards: true }, session.ctx.ids)[0]!;
    expect(marked.mark).toBe(true);
    const markedSel = hitSelection(session.doc, marked);
    expect(runCommand(session, 'find.replace', { text: '?^13?', replacement: 'X', matchWildcards: true }, { from: markedSel.from, to: markedSel.to }).ok).toBe(false);
    expect(session.doc.eq(before)).toBe(true);
    const plain = findIn(session.doc, 'e?t', { matchWildcards: true }, session.ctx.ids).find((h) => !h.mark)!;
    expect(runCommand(session, 'find.replace', { text: 'e?t', replacement: 'X', matchWildcards: true }, { from: plain.from, to: plain.to }).ok).toBe(true);
    expect(session.doc.eq(before)).toBe(false);
  });

  test('a hit that would split a grapheme cluster is not produced (ED-001 section 3.7), as the engine does not produce one since core-ts 0.2.2', () => {
    const composed = `e${String.fromCharCode(0x301)} e`;
    const doc = schema.topNodeType.create(null, [schema.nodes['paragraph']!.create(null, [schema.text(composed)])]);
    expect(findIn(doc, 'e').map((h) => h.from)).toEqual([4]);
    const skin = String.fromCodePoint(0x1f44d, 0x1f3fd);
    const emoji = schema.topNodeType.create(null, [schema.nodes['paragraph']!.create(null, [schema.text(`${skin} ${String.fromCodePoint(0x1f44d)}`)])]);
    expect(findIn(emoji, String.fromCodePoint(0x1f44d)).length).toBe(1);
  });

  test("an invalid expression or wildcard pattern throws the engine's error; wildcards and an expression together are refused", () => {
    expect(() => new Finder('(', { matchRegExp: true })).toThrow();
    expect(() => new Finder('(a', { matchWildcards: true })).toThrow(/wildcard/);
    expect(() => new Finder('a', { matchRegExp: true, matchWildcards: true })).toThrow(/two languages/);
    expect(new Finder('', {}).all(schema.topNodeType.create(null, [schema.nodes['paragraph']!.create(null, [schema.text('x')])]))).toEqual([]);
  });

  test('the memo: after an edit only the changed block is searched again', async () => {
    // Over a state, as the view holds it: `session.doc` is read from the store afresh, every node new.
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const state = stateFor(session);
    const finder = new Finder('text', {}, session.ctx.ids);
    const before = finder.all(state.doc).length;
    const blocks = state.doc.childCount;
    expect(finder.misses).toBe(blocks);
    const next = state.apply(state.tr.insertText('text ', 3));
    expect(finder.all(next.doc).length).toBe(before + 1);
    expect(finder.misses).toBe(blocks + 1);
  });
});

describe('find.replace and find.replaceAll (ED-005 section 3.3)', () => {
  test('replace all is one transaction and one undo step, and takes the first character\'s formatting', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const count = findIn(session.doc, 'indented', {}, session.ctx.ids).length;
    expect(count).toBe(20);
    const r = runCommand(session, 'find.replaceAll', { text: 'indented', replacement: 'moved' });
    expect(r.ok).toBe(true);
    expect(findIn(session.doc, 'indented', {}, session.ctx.ids).length).toBe(0);
    expect(findIn(session.doc, 'moved', {}, session.ctx.ids).length).toBe(20);
    // Block 26: "Font styles " then "Aerial Black" in a font; a hit across the change takes the first run's formatting.
    runCommand(session, 'find.replaceAll', { text: 'styles Aerial', replacement: 'X' });
    const b26 = session.doc.child(26);
    expect(textOf(b26)).toBe('Font X Black');
    const marksOf = (i: number): string[] => { let at = 0; let out: string[] = []; b26.forEach((c) => { if (at <= i && i < at + c.nodeSize) out = c.marks.map((m) => m.type.name); at += c.nodeSize; }); return out; };
    expect(marksOf(5)).not.toContain('font');
    expect(marksOf(7)).toContain('font');
    session.undoManager.undo();
    expect(textOf(session.doc.child(26))).toBe('Font styles Aerial Black');
    session.undoManager.undo();
    expect(findIn(session.doc, 'indented', {}, session.ctx.ids).length).toBe(20);
  });

  test('group references: $1 in an expression, \\1 with wildcards; atoms without text stay', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    runCommand(session, 'find.replaceAll', { text: '(Some) (text)', replacement: '$2 $1', matchRegExp: true });
    expect(textOf(session.doc.child(4))).toBe('text Some.');
    runCommand(session, 'find.replaceAll', { text: '(Left) (indent)', replacement: '\\2 \\1', matchWildcards: true });
    expect(textOf(session.doc.child(12))).toBe('indent Left');
    // "short " and the markers and "para": the text goes, the markers stay after the replacement, as the engine's `splice` leaves them; "X" is in the run "short " was in.
    runCommand(session, 'find.replaceAll', { text: 'short para', replacement: 'X' });
    const order: string[] = [];
    session.doc.child(23).forEach((n) => order.push(n.isText ? n.text! : String(n.attrs['kind'])));
    expect(order).toEqual(['A ', 'X', 'bookmarkStart', 'bookmarkEnd', 'graph with 10 points spacing before, 20 points after.']);
  });

  test('replace: the first press selects the next hit, the second replaces it and selects the one after, wrapping', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    let state = runCommand(session, 'find.replace', { text: 'aaa', replacement: 'Z' }, { from: 1 }).state!;
    const first = findIn(session.doc, 'aaa')[0]!;
    expect([state.selection.from, state.selection.to]).toEqual([first.from, first.to]);
    state = runCommand(session, 'find.replace', { text: 'aaa', replacement: 'Z' }, state.selection).state!;
    expect(findIn(session.doc, 'aaa').length).toBe(13);
    expect(state.doc.textBetween(state.selection.from, state.selection.to)).toBe('aaa');
    expect(state.selection.from).toBeGreaterThan(first.from);
    // Nothing matches: the command does not apply.
    expect(runCommand(session, 'find.replace', { text: 'no such words', replacement: 'Z' }, { from: 1 }).ok).toBe(false);
    expect(runCommand(session, 'find.replaceAll', { text: 'no such words', replacement: 'Z' }).ok).toBe(false);
    expect(runCommand(session, 'find.replaceAll', { text: '(', replacement: 'Z', matchRegExp: true }).ok).toBe(false);
  });

  test('a replacement that itself matches is not found again inside itself', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const n = findIn(session.doc, 'bbb').length;
    runCommand(session, 'find.replaceAll', { text: 'bbb', replacement: 'bbbbbb' });
    expect(findIn(session.doc, 'bbbbbb').length).toBe(n);
    const state = stateFor(session, { from: 1 });
    expect(state.doc.textContent.includes('bbbbbb')).toBe(true);
  });
});
