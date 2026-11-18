// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 9.2 item 1, E4.g step G1: the editor's file for Jason's session in Word, made in a
// room by a person's session as two authors and by the agent through the companion's tools. The
// builder runs as it runs for the session (`scripts/review-for-word.mjs`); the file passes its
// self-check, round-trips unedited, and lists what it made. The listing is pinned whole: it is what
// Jason compares with Word's Reviewing Pane, so a change to it is a change to what he is told.
import { describe, expect, test } from 'vitest';
import { ZipPartStore } from '@docx4j/core-ts';
import { EditorDocument } from '@docx4j/editor-model';
import { buildReviewForWord, describe as listing } from '../scripts/review-for-word.mjs';

/** The parts that differ between two packages, bar those the engine regenerates by design. */
function differences(before: Uint8Array, after: Uint8Array): string[] {
  const a = new ZipPartStore(before);
  const b = new ZipPartStore(after);
  const an = new Set(a.partNames());
  const bn = new Set(b.partNames());
  const out: string[] = [];
  for (const n of new Set([...an, ...bn])) {
    if (n === '[Content_Types].xml' || n.endsWith('.rels')) continue;
    if (!an.has(n) || !bn.has(n)) { out.push(`${n} (missing on one side)`); continue; }
    const x = a.loadSync(n);
    const y = b.loadSync(n);
    if (x.length !== y.length || x.some((v, i) => v !== y[i])) out.push(n);
  }
  return out;
}

describe('the editor\'s file for Word (section 9.2 item 1)', () => {
  test('made in a room by two people and the agent: the self-check passes, an unedited round trip holds, and the listing is what the session made', async () => {
    const result = await buildReviewForWord();
    expect(result.selfCheck).toBe(true);

    const reopened = await EditorDocument.open(result.bytes);
    expect(differences(result.bytes, (await reopened.save()).bytes)).toEqual([]);

    const A = 'Author A';
    const B = 'Author B';
    const C = 'Claude';
    expect(result.changes).toEqual([
      // 1, 2: text inserted and deleted by two authors; an insertion deleted by another.
      { type: 'Added', author: A, text: 'very ' },
      { type: 'Deleted', author: B, text: 'old ' },
      { type: 'Added', author: A, text: 'green ' },
      { type: 'Deleted', author: B, text: 'green ' },
      // 3: typed over a selection (section 9.7's replacement).
      { type: 'Deleted', author: A, text: '7 days' },
      { type: 'Added', author: A, text: '14 days' },
      // 4: run formatting.
      { type: 'Formatted', author: A, text: 'alpha' },
      { type: 'Formatted', author: B, text: 'bravo' },
      // 5: paragraph properties (a paragraph's properties change lists with no text).
      { type: 'Formatted', author: A, text: '' },
      { type: 'Formatted', author: B, text: '' },
      { type: 'Formatted', author: A, text: '' },
      // 6: a paragraph mark inserted (the split) and one deleted (the join).
      { type: 'Added', author: A, text: '\r' },
      { type: 'Deleted', author: B, text: '\r' },
      // 7: a row inserted and typed in, and one deleted; a row's text is its cells with a tab between, then "\r\n",
      // as Office JS gives it (core-ts checks 30 and 31).
      { type: 'Added', author: A, text: 'new\tnew\tnew\r\n', row: true },
      { type: 'Deleted', author: B, text: '3\tDelete\tthis row is deleted\r\n', row: true },
      // 8: a tab deleted and one inserted; a footnote reference inserted, and one deleted.
      { type: 'Deleted', author: A, text: '\t' },
      { type: 'Added', author: A, text: '\t' },
      { type: 'Added', author: B, text: '' },
      { type: 'Deleted', author: B, text: '' },
      // 9: a list started (a paragraph's properties); a link made over words (its character style).
      { type: 'Formatted', author: A, text: '' },
      { type: 'Formatted', author: B, text: 'docx4j site' },
      // 10: Replace all, three matches (one in section 11's paragraph).
      { type: 'Deleted', author: B, text: 'colour' },
      { type: 'Added', author: B, text: 'color' },
      { type: 'Deleted', author: B, text: 'colour' },
      { type: 'Added', author: B, text: 'color' },
      { type: 'Deleted', author: B, text: 'colour' },
      { type: 'Added', author: B, text: 'color' },
      // 12: the agent. Its paragraph replaced by a heading and a paragraph, then a list; its comment's
      // reference parts the insertion after the run holding "in advance" (section 12.29 item 3), and the link's
      // text is a change of its own, as Office JS and Word's pane list it (core-ts check 30, section 12.32 item 3).
      { type: 'Deleted', author: C, text: 'The tenant pays rent monthly in arrears.\r' },
      { type: 'Added', author: C, text: 'Rent\rThe tenant pays rent monthly in advance, as the ' },
      { type: 'Added', author: C, text: 'lease' },
      { type: 'Added', author: C, text: ' says.\rThe first obligation.\rThe second obligation.\r' },
      // the agent's replace_text through the engine writes the insertion first (core-ts 0.6.0, CR-014 phase 3; ED-008 section 8.8 item 9)
      { type: 'Added', author: C, text: '60 days' },
      { type: 'Deleted', author: C, text: '30 days' },
    ]);
    expect(result.preserved).toEqual([]);
    expect(result.comments).toEqual([
      { author: C, text: 'Changed from arrears, as agreed.', done: false },
      { author: A, text: 'A comment by Author A, replied to and resolved.', done: true },
      { author: B, text: 'A reply by Author B.', reply: true },
      { author: B, text: 'A comment by Author B.', done: false },
      { author: A, text: 'A comment by Author A over two paragraphs.', done: false },
    ]);
    expect(listing(result)).toContain('By type: Added 14, Deleted 12, Formatted 7. Authors: Author A, Author B, Claude.');
  }, 90_000);
});
