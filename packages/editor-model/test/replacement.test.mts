// SPDX-License-Identifier: GPL-3.0-or-later
//
// A replacement as one (ED-005 proposal 41, section 9.7): the pairing the review panel shows
// (`replacementPairs`), held to its rule over every fixture and in the cases the plan names, and the
// `keys` of `trackedChange.accept` and `.reject`, which resolve those changes and no others, in one
// transaction and one undo step. The room cases are in `concurrency.test.mts`.
import { describe, test, expect } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, trackedChangesOf, replacementPairs, type TrackedChangeInfo } from '../src/index.mjs';
import { fixture, allFixtures } from './helpers.mjs';

const pairsOf = (doc: PMNode): [string, string, string][] =>
  [...replacementPairs(doc, trackedChangesOf(doc))].map(([d, i]) => [d.author ?? '', d.text, i.text]);

/** The position of block `i`'s first character. */
function blockFrom(s: EditorDocument, i: number): number {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += s.doc.child(k).nodeSize;
  return pos + 1;
}

/** The changes this test makes: the fixture has two of its own, by another author. */
const ours = (doc: PMNode): string[] => trackedChangesOf(doc).filter((c) => c.author === 'Author A' || c.author === 'Author B').map((c) => `${c.type}:${c.author}:${c.text}`);

async function tracking(author: string): Promise<EditorDocument> {
  const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
  session.trackingMode = 'TrackAll';
  session.trackingAuthor = author;
  return session;
}

describe('the pairing (replacementPairs)', () => {
  test('over every fixture: a deletion and the same author\'s insertion beside it in either order, nothing with text between, a block\'s end only when the first took its mark', async () => {
    let pairs = 0;
    for (const name of await allFixtures()) {
      const doc = (await EditorDocument.open(await fixture(name))).doc;
      for (const [deleted, inserted] of replacementPairs(doc, trackedChangesOf(doc))) {
        pairs++;
        expect(deleted.type, name).toBe('Deleted');
        expect(inserted.type, name).toBe('Added');
        expect(inserted.author, name).toBe(deleted.author);
        const [first, second] = inserted.from >= deleted.to ? [deleted, inserted] : [inserted, deleted];
        expect(second.from, name).toBeGreaterThanOrEqual(first.to);
        expect(doc.textBetween(first.to, second.from), name).toBe('');
        if (doc.resolve(first.to).parent !== doc.resolve(second.from).parent) expect(first.pieces.at(-1)!.what, name).toBe('mark');
      }
    }
    expect(pairs).toBeGreaterThan(0);
  });

  test('review-for-word.docx: words typed over, the agent\'s paragraph and its replacement; an insertion first, or a tab deleted and another inserted further on, never pair', async () => {
    const doc = (await EditorDocument.open(await fixture('review/review-for-word.docx'))).doc;
    expect(pairsOf(doc)).toEqual([
      ['Author A', '7 days', '14 days'],
      ['Author B', 'colour', 'color'],
      ['Author B', 'colour', 'color'],
      ['Author B', 'colour', 'color'],
      // A paragraph deleted with its mark, and the paragraphs inserted after it (up to the link, a holder of its own).
      ['Claude', 'The tenant pays rent monthly in arrears.\r', 'Rent\rThe tenant pays rent monthly in advance, as the '],
      ['Claude', '30 days', '60 days'],
    ]);
  });

  test('typing over a selection pairs', async () => {
    const s = await tracking('Author A');
    const b = blockFrom(s, 4);
    expect(runCommand(s, 'insertText', { text: 'Many' }, { from: b, to: b + 4 }).ok).toBe(true);
    expect(pairsOf(s.doc)).toEqual([['Author A', 'Some', 'Many']]);
  });

  test('an insertion first and the deletion after it pair too: the order a script\'s Replace writes through the engine (core-ts 0.6.0, check 54 case 4; ED-008 section 8.8 item 9)', async () => {
    const s = await tracking('Author A');
    const b = blockFrom(s, 4);
    expect(runCommand(s, 'insertText', { text: 'Many' }, { from: b }).ok).toBe(true);
    expect(runCommand(s, 'insertText', { text: '' }, { from: b + 4, to: b + 8 }).ok).toBe(true);
    expect(ours(s.doc)).toEqual(['Added:Author A:Many', 'Deleted:Author A:Some']);
    expect(pairsOf(s.doc)).toEqual([['Author A', 'Some', 'Many']]);
  });

  test('a deletion and another author\'s insertion after it stay two', async () => {
    const s = await tracking('Author A');
    const b = blockFrom(s, 4);
    expect(runCommand(s, 'insertText', { text: '' }, { from: b, to: b + 4 }).ok).toBe(true);
    s.trackingAuthor = 'Author B';
    expect(runCommand(s, 'insertText', { text: 'Many' }, { from: b + 4 }).ok).toBe(true);
    expect(ours(s.doc)).toEqual(['Deleted:Author A:Some', 'Added:Author B:Many']);
    expect(pairsOf(s.doc)).toEqual([]);
  });

  test('the tail of a paragraph deleted, its mark kept, and words inserted at the start of the next: two acts, two entries', async () => {
    const s = await tracking('Author A');
    const b = blockFrom(s, 4);
    // "text." deleted up to the paragraph's end, the mark kept; then "New " at the start of block 5.
    expect(runCommand(s, 'insertText', { text: '' }, { from: b + 5, to: b + 10 }).ok).toBe(true);
    expect(runCommand(s, 'insertText', { text: 'New ' }, { from: blockFrom(s, 5) }).ok).toBe(true);
    expect(ours(s.doc)).toEqual(['Deleted:Author A:text.', 'Added:Author A:New ']);
    expect(pairsOf(s.doc)).toEqual([]);
  });
});

describe('keys: several changes resolved as one (trackedChange.accept and .reject)', () => {
  /** A replacement by Author A in block 4, then Author B's bold over its inserted word. */
  async function replacedAndFormatted(): Promise<{ s: EditorDocument; deleted: TrackedChangeInfo; inserted: TrackedChangeInfo }> {
    const s = await tracking('Author A');
    const b = blockFrom(s, 4);
    expect(runCommand(s, 'insertText', { text: 'Many' }, { from: b, to: b + 4 }).ok).toBe(true);
    s.trackingAuthor = 'Author B';
    s.undoManager.stopCapturing();
    expect(runCommand(s, 'font.bold', { value: true }, { from: b + 4, to: b + 8 }).ok).toBe(true);
    const [deleted, inserted] = [...replacementPairs(s.doc, trackedChangesOf(s.doc))][0]!;
    return { s, deleted, inserted };
  }
  const listing = (s: EditorDocument): string[] => ours(s.doc);

  for (const [action, text] of [['accept', 'Many text.'], ['reject', 'Some text.']] as const) {
    test(`${action}: both halves in one transaction and one undo step; another author's formatting on the inserted word left`, async () => {
      const { s, deleted, inserted } = await replacedAndFormatted();
      const before = listing(s);
      expect(before).toContain('Formatted:Author B:Many');
      s.undoManager.stopCapturing();
      expect(runCommand(s, `trackedChange.${action}`, { keys: s.changeKeys([deleted, inserted]) }).ok).toBe(true);
      expect(s.doc.child(4).textContent).toBe(text);
      // B's formatting change stays where there is still text to carry it: a range over both halves would have taken it too.
      expect(listing(s)).toEqual(action === 'accept' ? ['Formatted:Author B:Many'] : []);
      s.undoManager.undo();
      expect(listing(s)).toEqual(before);
    });
  }

  test('a key twice, or a key whose change is gone, resolves what is there once', async () => {
    const { s, deleted, inserted } = await replacedAndFormatted();
    const [dk, ik] = s.changeKeys([deleted, inserted]);
    expect(runCommand(s, 'trackedChange.accept', { keys: [dk!, dk!] }).ok).toBe(true);
    expect(listing(s)).toEqual(['Added:Author A:Many', 'Formatted:Author B:Many']);
    expect(runCommand(s, 'trackedChange.accept', { keys: [dk!, ik!] }).ok).toBe(true);
    expect(listing(s)).toEqual(['Formatted:Author B:Many']);
  });
});
