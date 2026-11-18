// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 sections 5.5 and 5.6, step C4: accept and reject, held to what Word and Office JS do.
// Check 23's cases (core-ts's, the listing Office JS gave and the paragraphs after accepting or
// rejecting the first change); check 22's file (a deleted paragraph, a lone deleted mark, a
// deletion across a mark), each change resolved and compared with the file Word saved; check 18's
// move accepted and rejected from either half and all at once; check 20's edited and broken moves.
// check 25's list turned to bullets and rejected. Word's files are core-ts's `test/fixtures/revisions/check18`, `check20`,
// `check22` and `check25`, copied.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, runCommandAsync, trackedChangesOf, preservedChangesOf, preservedGroupsOf, formattingKey, type Revision, type TrackedChangeInfo } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** Every paragraph as [its text accepted, its text as it was, its mark's revision], as core-ts's check 22 test reads them. */
function paragraphsOf(doc: PMNode): [string, string, string][] {
  const out: [string, string, string][] = [];
  doc.descendants((node) => {
    if (node.type.name !== 'paragraph') return true;
    let accepted = '';
    let original = '';
    node.descendants((n) => {
      if (!n.isText) return true;
      const names = n.marks.map((m) => m.type.name);
      if (!names.includes('deletion')) accepted += n.text;
      if (!names.includes('insertion')) original += n.text;
      return false;
    });
    const r = node.attrs['markRevision'] as Revision | null;
    out.push([accepted, original, r?.insertion ? 'ins' : r?.deletion ? 'del' : '']);
    return false;
  });
  return out;
}

/** Every paragraph's properties and its text's formatting, runs merged where the formatting is equal: what a formatting or properties change resolved leaves. */
function formattingOf(doc: PMNode): string[] {
  const out: string[] = [];
  doc.descendants((node) => {
    if (node.type.name !== 'paragraph') return true;
    const runs: [string, string][] = [];
    node.descendants((n) => {
      if (!n.isText) return true;
      const key = formattingKey(n.marks) + (n.marks.some((m) => m.type.name === 'formatChange') ? ' changed' : '');
      const last = runs[runs.length - 1];
      if (last && last[1] === key) last[0] += n.text; else runs.push([n.text!, key]);
      return false;
    });
    const a = node.attrs;
    out.push(JSON.stringify([a['pStyle'], a['jc'], a['ind'], a['numPr'], a['pPrChange'] !== null, runs]));
    return false;
  });
  return out;
}

/** The tables' and sections' properties of a saved file, as the engine writes them: what the revisions shown in Word only resolved leave (ED-005 section 12.26). */
async function tablesAndSectionsOf(bytes: Uint8Array): Promise<string[]> {
  const main = (await WordprocessingMLPackage.load(bytes)).getMainDocumentPart();
  await main.getContents();
  const xml = (await main.getXml()).replace(/ w:rsid\w*="[^"]*"/g, '');
  return [...xml.matchAll(/<w:(tblPr|tblGrid|trPr|tcPr|sectPr)\b[^>]*\/>|<w:(tblPr|tblGrid|trPr|tcPr|sectPr)\b.*?<\/w:\2>/g)].map((m) => m[0]);
}

const listing = (changes: readonly TrackedChangeInfo[]): [string, string | null, string][] => changes.map((c) => [c.type, c.author, c.text]);

/** Resolve the change at `index` of the listing through the command, over its range. */
function resolve(session: EditorDocument, index: number, action: 'accept' | 'reject'): void {
  const c = trackedChangesOf(session.doc)[index]!;
  const r = runCommand(session, action === 'accept' ? 'trackedChange.accept' : 'trackedChange.reject', { from: c.from, to: c.to });
  expect(r.ok, `${action} ${index}`).toBe(true);
}

describe('what one change is, and what resolving it takes (core-ts check 23, Office JS)', () => {
  const D1 = '2026-09-28T01:00:00Z';
  const D2 = '2026-09-28T02:00:00Z';
  let nextId = 100;
  const who = (author: string, date: string): string => `w:id="${nextId++}" w:author="Author ${author}" w:date="${date}"`;
  const p = (inner: string, mark?: string): string => `<w:p>${mark ? `<w:pPr><w:rPr>${mark}</w:rPr></w:pPr>` : ''}${inner}</w:p>`;
  const r = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const ins = (a: string, d: string, text: string): string => `<w:ins ${who(a, d)}>${r(text)}</w:ins>`;
  const del = (a: string, d: string, text: string): string => `<w:del ${who(a, d)}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
  const insMark = (a: string, d: string): string => `<w:ins ${who(a, d)}/>`;
  const delMark = (a: string, d: string): string => `<w:del ${who(a, d)}/>`;
  // [markup, Office JS's listing, the paragraphs' text after accepting the first change, after rejecting it]
  const CASES: Record<string, [string, [string, string, string][], string[], string[]]> = {
    'ins-paragraph': [p(r('One.')) + p(ins('A', D1, 'Two.'), insMark('A', D1)) + p(r('Three.')),
      [['Added', 'Author A', 'Two.\r']], ['One.', 'Two.', 'Three.'], ['One.', 'Three.']],
    'ins-across': [p(r('Four ') + ins('A', D1, 'five.'), insMark('A', D1)) + p(ins('A', D1, 'Six ') + r('seven.')),
      [['Added', 'Author A', 'five.\rSix ']], ['Four five.', 'Six seven.'], ['Four seven.']],
    'ins-mark-alone': [p(r('Eight.'), insMark('A', D1)) + p(r('Nine.')),
      [['Added', 'Author A', '\r']], ['Eight.', 'Nine.'], ['Eight.Nine.']],
    'del-paragraph-dates': [p(r('One.')) + p(del('A', D1, 'Two.'), delMark('A', D2)) + p(r('Three.')),
      [['Deleted', 'Author A', 'Two.\r']], ['One.', 'Three.'], ['One.', 'Two.', 'Three.']],
    'del-paragraph-authors': [p(r('One.')) + p(del('A', D1, 'Two.'), delMark('B', D1)) + p(r('Three.')),
      [['Deleted', 'Author A', 'Two.'], ['Deleted', 'Author B', '\r']], ['One.', '', 'Three.'], ['One.', 'Two.', 'Three.']],
    'del-same': [p(r('Ten ') + del('A', D1, 'eleven ') + del('A', D1, 'twelve') + r('.')),
      [['Deleted', 'Author A', 'eleven twelve']], ['Ten .'], ['Ten eleven twelve.']],
    'del-dates': [p(r('Ten ') + del('A', D1, 'eleven ') + del('A', D2, 'twelve') + r('.')),
      [['Deleted', 'Author A', 'eleven twelve']], ['Ten .'], ['Ten eleven twelve.']],
    'del-authors': [p(r('Ten ') + del('A', D1, 'eleven ') + del('B', D1, 'twelve') + r('.')),
      [['Deleted', 'Author A', 'eleven '], ['Deleted', 'Author B', 'twelve']], ['Ten .'], ['Ten eleven .']],
    'del-then-ins': [p(r('Ten ') + del('A', D1, 'eleven') + ins('A', D1, 'twelve') + r('.')),
      [['Deleted', 'Author A', 'eleven'], ['Added', 'Author A', 'twelve']], ['Ten twelve.'], ['Ten eleventwelve.']],
  };
  const open = async (blocks: string): Promise<EditorDocument> => {
    const pkg = await WordprocessingMLPackage.createPackage();
    pkg.getMainDocumentPart().setXml(`<w:document xmlns:w="${W}"><w:body>${blocks}</w:body></w:document>`);
    await pkg.getMainDocumentPart().getContents();
    return EditorDocument.open(await pkg.save());
  };
  for (const [name, [blocks, expected, accepted, rejected]] of Object.entries(CASES)) {
    test(name, async () => {
      expect(listing(trackedChangesOf((await open(blocks)).doc))).toEqual(expected);
      for (const [action, after] of [['accept', accepted], ['reject', rejected]] as const) {
        const session = await open(blocks);
        resolve(session, 0, action);
        // A paragraph's text as Office JS reads it: accepted, its deleted mark not yet joining it.
        expect(paragraphsOf(session.doc).map(([text]) => text), `${name} ${action}`).toEqual(after);
        expect(listing(trackedChangesOf(session.doc)), `${name} ${action}: the rest left`).toEqual(expected.slice(1));
        await session.save();
      }
    });
  }
});

describe("a deleted paragraph, a lone deleted mark, a deletion across a mark (core-ts check 22, Word's files)", () => {
  const FILE = 'revisions/check22/22-deleted.docx';
  test('listed as Office JS lists them: three changes, a mark "\\r"', async () => {
    const session = await EditorDocument.open(await fixture(FILE));
    expect(listing(trackedChangesOf(session.doc))).toEqual([['Deleted', 'Author A', 'Two.\r'], ['Deleted', 'Author A', '\r'], ['Deleted', 'Author A', 'six.\rSeven ']]);
  });

  for (const action of ['accept', 'reject'] as const) {
    test(`${action}: each of the three as Word resolved it`, async () => {
      const original = paragraphsOf((await EditorDocument.open(await fixture(FILE))).doc);
      // The lone mark's result is the JSON's (Word saved no file for it), as core-ts reads it.
      const lone = [...original];
      if (action === 'accept') lone.splice(3, 2, ['Four.Five ', 'Four.Five six.', 'del']); else lone.splice(3, 1, ['Four.', 'Four.', '']);
      for (const [index, expected] of [[0, `22-${action}-whole`], [1, lone], [2, `22-${action}-across`]] as const) {
        const session = await EditorDocument.open(await fixture(FILE));
        resolve(session, index, action);
        const word = typeof expected === 'string' ? paragraphsOf((await EditorDocument.open(await fixture(`revisions/check22/${expected}.docx`))).doc) : expected;
        expect(paragraphsOf(session.doc), `${action} ${index}`).toEqual(word);
        expect(trackedChangesOf(session.doc), `${action} ${index}: the other two left`).toHaveLength(2);
        await session.save();
      }
    });
  }
});

/** The Word file of a check, as paragraphs. */
const wordParagraphs = async (name: string): Promise<[string, string, string][]> => paragraphsOf((await EditorDocument.open(await fixture(name))).doc);

describe("a move resolved from either half, and all at once (core-ts check 18, Word's files)", () => {
  const INPUT = 'revisions/revisions-word15.docx';
  for (const [action, half] of [['accept', 'destination'], ['accept', 'source'], ['reject', 'destination'], ['reject', 'source']] as const) {
    test(`${action} on the ${half}: the whole move, as Word's ${action}-${half}.docx`, async () => {
      const session = await EditorDocument.open(await fixture(INPUT));
      session.trackingMode = 'Off';
      const change = trackedChangesOf(session.doc).find((c) => c.move !== undefined && c.type === (half === 'destination' ? 'Added' : 'Deleted'))!;
      expect(change).toBeDefined();
      runCommand(session, `trackedChange.${action}`, { from: change.from, to: change.to });
      expect(trackedChangesOf(session.doc).some((c) => c.move !== undefined)).toBe(false);
      let markers = 0;
      session.doc.descendants((n) => { if (String(n.attrs['kind'] ?? '').startsWith('move')) markers++; return true; });
      expect(markers).toBe(0);
      expect(paragraphsOf(session.doc)).toEqual(await wordParagraphs(`revisions/check18/${action}-${half}.docx`));
      await session.save();
    });
  }
  for (const action of ['accept', 'reject'] as const) {
    test(`${action} all: as Word's ${action}-all.docx, the table's and the section's properties too`, async () => {
      const session = await EditorDocument.open(await fixture(INPUT));
      session.trackingMode = 'Off';
      expect((await runCommandAsync(session, action === 'accept' ? 'trackedChanges.acceptAll' : 'trackedChanges.rejectAll')).ok).toBe(true);
      expect(trackedChangesOf(session.doc)).toEqual([]);
      const word = await EditorDocument.open(await fixture(`revisions/check18/${action}-all.docx`));
      expect(paragraphsOf(session.doc)).toEqual(paragraphsOf(word.doc));
      // The formatting changes (section 4) and properties changes (5, 9) resolved as Word resolved them.
      expect(formattingOf(session.doc)).toEqual(formattingOf(word.doc));
      // And the table's and the section's, through the engine (core-ts 0.3.0, ED-005 section 12.26).
      expect(preservedChangesOf(session.doc, (id) => session.ctx.ids.get(id)?.value)).toEqual([]);
      const saved = await session.save();
      expect(await tablesAndSectionsOf(saved.bytes)).toEqual(await tablesAndSectionsOf(await fixture(`revisions/check18/${action}-all.docx`)));
    });
  }
});

describe("a list turned to bullets, rejected (core-ts check 25, Word's files)", () => {
  // Word recorded "One." and "Three." with hanging=360 alone, part of their level's left=720 hanging=360, and wrote no
  // indent for them on reject, as for "Two.", whose whole indent was its level's (25b): an indent is dropped when every
  // attribute it states is the restored level's.
  // 25e and 25f: "Three." recorded numId 1 with no level while it stood at level 1 of another list; Word put back the
  // recorded level, writing its w:ilvl 0, not the current one (core-ts d68e803).
  for (const [input, word] of [['25a-bullets', '25b-bullets-reject'], ['25e-bullet-demote', '25f-bullet-demote-reject']]) {
    test(`reject all: as Word's ${word}.docx, the recorded numbering and level, no indent`, async () => {
      const session = await EditorDocument.open(await fixture(`revisions/check25/${input}.docx`));
      session.trackingMode = 'Off';
      expect(runCommand(session, 'trackedChanges.rejectAll').ok).toBe(true);
      expect(trackedChangesOf(session.doc)).toEqual([]);
      const theirs = await EditorDocument.open(await fixture(`revisions/check25/${word}.docx`));
      expect(formattingOf(session.doc)).toEqual(formattingOf(theirs.doc));
      expect(await (async () => { const saved = await session.save(); return (await (await WordprocessingMLPackage.load(saved.bytes)).getMainDocumentPart().getXml()); })())
        .toMatch(/<w:numPr><w:ilvl w:val="0"\/><w:numId w:val="1"\/><\/w:numPr><\/w:pPr><w:r(?: [^>]*)?><w:t>Three\./);
    });
  }
});

describe("touching formatting changes, and a table's property records (core-ts check 28, Word's files)", () => {
  // Bold applied over an italic run and a plain one: each run recorded different formatting, and Office JS
  // listed one change, Formatted "alpha beta" (core-ts 57cab79). Word's accept and reject resolve the
  // runs' formatting; the two tables' property records are shown in Word only here, one entry per table.
  test('the listing: one formatting change over the two runs; the tables\' records one entry each', async () => {
    const session = await EditorDocument.open(await fixture('revisions/check28/28a-changes.docx'));
    const formatted = trackedChangesOf(session.doc).filter((c) => c.type === 'Formatted');
    expect(formatted.map((c) => c.text)).toEqual(['alpha beta']);
    expect(formatted[0]!.pieces.length).toBe(2);
    const groups = preservedGroupsOf(session.doc, preservedChangesOf(session.doc, (id) => session.ctx.ids.get(id)?.value));
    expect(groups.map((g) => g.kind)).toEqual(['tableProperties', 'tableProperties']);
    expect(groups.reduce((n, g) => n + g.records.length, 0)).toBe(14);
  });

  for (const [action, word] of [['accept', '28a-accept-all'], ['reject', '28a-reject-all']] as const) {
    test(`${action} all: the runs' formatting and the tables' properties as Word's ${word}.docx`, async () => {
      const session = await EditorDocument.open(await fixture('revisions/check28/28a-changes.docx'));
      session.trackingMode = 'Off';
      expect((await runCommandAsync(session, action === 'accept' ? 'trackedChanges.acceptAll' : 'trackedChanges.rejectAll')).ok).toBe(true);
      expect(trackedChangesOf(session.doc)).toEqual([]);
      const theirs = await EditorDocument.open(await fixture(`revisions/check28/${word}.docx`));
      expect(formattingOf(session.doc)).toEqual(formattingOf(theirs.doc));
      // The two tables' records, through the engine (ED-005 section 12.26).
      expect(preservedChangesOf(session.doc, (id) => session.ctx.ids.get(id)?.value)).toEqual([]);
      const saved = await session.save();
      expect(await tablesAndSectionsOf(saved.bytes)).toEqual(await tablesAndSectionsOf(await fixture(`revisions/check28/${word}.docx`)));
    });
  }
});

describe("edited and broken moves (core-ts check 20, Word's files)", () => {
  const CASES = ['01-type-inside-a', '03-type-at-start', '04-type-at-end', '06-delete-part-a', '08-delete-all', '12-partner-missing', '13-moveTo-no-ranges'];
  for (const name of CASES) {
    for (const action of ['accept', 'reject'] as const) {
      test(`${name}, ${action} on the destination: as Word's file`, async () => {
        const session = await EditorDocument.open(await fixture(`revisions/check20/${name}.docx`));
        session.trackingMode = 'Off';
        const changes = trackedChangesOf(session.doc);
        const change = changes.find((c) => c.move !== undefined && c.type === 'Added') ?? changes.find((c) => c.type === 'Added' && c.text.includes('The second paragraph'))!;
        expect(change, `${name}: the destination`).toBeDefined();
        // By its key, as Word's Accept This Change acts on the one change: the destination's range
        // also holds what was typed or deleted inside it, which are changes of their own.
        expect(runCommand(session, `trackedChange.${action}`, { key: session.changeKey(change) }).ok).toBe(true);
        const word = await EditorDocument.open(await fixture(`revisions/check20/${name}-${action}.docx`));
        expect(paragraphsOf(session.doc)).toEqual(paragraphsOf(word.doc));
        // What is left pending, a move or a plain change, as Word left it.
        const left = (doc: PMNode): string[] => trackedChangesOf(doc).map((c) => `${c.type} ${c.author ?? ''} ${c.move !== undefined ? 'move' : 'plain'} ${JSON.stringify(c.text)}`);
        expect(left(session.doc)).toEqual(left(word.doc));
        await session.save();
      });
    }
  }
});

describe('the commands (ED-005 section 5.5)', () => {
  test('with tracking on, accepting all is not itself tracked, and is one undo step', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author C';
    const before = listing(trackedChangesOf(session.doc));
    expect(before.length).toBeGreaterThan(0);
    const r = runCommand(session, 'trackedChanges.acceptAll');
    expect(r.ok).toBe(true);
    expect(trackedChangesOf(session.doc)).toEqual([]);
    runCommand(session, 'history.undo');
    expect(listing(trackedChangesOf(session.doc))).toEqual(before);
  });

  test("an author's changes alone, and a key that names nothing any more", async () => {
    const session = await EditorDocument.open(await fixture('revisions/revisions-word15.docx'));
    session.trackingMode = 'Off';
    const b = trackedChangesOf(session.doc).filter((c) => c.author === 'Author B');
    expect(b.length).toBeGreaterThan(0);
    const key = session.changeKey(b[0]!);
    expect(runCommand(session, 'trackedChanges.rejectAll', { author: 'Author B' }).ok).toBe(true);
    expect(trackedChangesOf(session.doc).some((c) => c.author === 'Author B')).toBe(false);
    expect(trackedChangesOf(session.doc).some((c) => c.author === 'Author A')).toBe(true);
    expect(runCommand(session, 'trackedChange.accept', { key }).ok).toBe(false);
    await session.save();
  });
});
