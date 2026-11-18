// SPDX-License-Identifier: GPL-3.0-or-later
//
// E4.g step G5 (ED-005 section 9.2 item 1): the editor's Accept all and Reject all over its own file,
// `review/review-for-word.docx` (made in a room by a person's session and the agent's tools, section
// 12.30), against Word 2010's and Word 15's Accept All and Reject All saves of the same file (Jason's
// session, section 12.34), and Word 365's Reject All (2026-10-21, through core-ts 277d118). Compared block by block, as `review.test.mts` compares the editor with
// checks 18 to 29: each paragraph's text, style, alignment, indent, list level, direction and runs'
// formatting, each table's rows, the comments and the footnotes. The session found no difference;
// a difference here is a finding, fixed or sent to core-ts, never a tolerance (section 2.1).
import { describe, test, expect } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { ZipPartStore } from '@docx4j/core-ts';
import { EditorDocument, runCommandAsync, trackedChangesOf } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

/** The markers that carry no content of their own, and Word adds or drops as it saves. */
const MARKERS = new Set(['commentRangeStart', 'commentRangeEnd', 'bookmarkStart', 'bookmarkEnd', 'proofErr', 'lastRenderedPageBreak']);
/** The marks that are revisions or anchors, not formatting. */
const NOT_FORMATTING = new Set(['comment', 'insertion', 'deletion', 'formatChange', 'runSource']);

/** A mark as formatting: its name, and its attributes that are values (not ids). */
function markKey(m: PMNode['marks'][number]): string {
  if (m.type.name === 'link') return 'link';
  const attrs = Object.entries(m.attrs).filter(([k, v]) => v !== null && k !== 'sourceId' && k !== 'rId');
  return attrs.length > 0 ? `${m.type.name}${JSON.stringify(Object.fromEntries(attrs))}` : m.type.name;
}

/** A document as comparable data: per block, its kind, text, paragraph properties and runs' formatting. */
function shape(doc: PMNode): unknown[] {
  const paragraph = (n: PMNode): unknown => {
    const runs: { key: string; text: string }[] = [];
    n.forEach((c) => {
      if (c.isText) {
        const key = c.marks.filter((m) => !NOT_FORMATTING.has(m.type.name)).map(markKey).sort().join(',');
        const last = runs[runs.length - 1];
        if (last && last.key === key) last.text += c.text; else runs.push({ key, text: c.text! });
      } else if (c.type.name !== 'opaqueInline' || !MARKERS.has(c.attrs['kind'] as string)) {
        runs.push({ key: `[${c.type.name}${c.attrs['kind'] ? `:${String(c.attrs['kind'])}` : ''}]`, text: '' });
      }
    });
    const a = n.attrs;
    return {
      style: a['pStyle'], jc: a['jc'], ind: a['ind'], list: a['numPr'] ? (a['numPr'] as { ilvl: number }).ilvl : null, bidi: a['bidi'],
      runs: runs.map((r) => `${r.text}{${r.key}}`),
    };
  };
  // A table's cells as their blocks, compared as the body's are, so that the formatting typed in a row with
  // tracking on is held to Word's too (the third review point, ED-005 section 12.47).
  const block = (b: PMNode): unknown => {
    if (b.type.name === 'paragraph') return paragraph(b);
    if (b.type.name === 'table') {
      const rows: unknown[][] = [];
      b.forEach((row) => { const cells: unknown[] = []; row.forEach((cell) => { const blocks: unknown[] = []; cell.forEach((c) => blocks.push(block(c))); cells.push(blocks); }); rows.push(cells); });
      return { table: rows };
    }
    return { [b.type.name]: b.textContent };
  };
  const out: unknown[] = [];
  doc.forEach((b) => out.push(block(b)));
  return out;
}

async function comments(session: EditorDocument): Promise<string[]> {
  const threads = await session.stories.comments();
  return threads.flatMap((t) => [`${t.author}${t.done ? ' (done)' : ''}: ${session.stories.currentDoc(t).textContent}`, ...t.replies.map((r) => `  reply ${r.author}: ${session.stories.currentDoc(r).textContent}`)]);
}

async function footnotes(session: EditorDocument): Promise<string[]> {
  return (await session.stories.footnotes()).map((n) => n.doc.textContent);
}

/** The editor's own resolution of its file, saved and read again as Word's save is read. */
async function resolvedByTheEditor(action: 'accept' | 'reject'): Promise<EditorDocument> {
  const session = await EditorDocument.open(await fixture('review/review-for-word.docx'));
  await session.stories.comments();
  expect((await runCommandAsync(session, `trackedChanges.${action}All`, {})).ok).toBe(true);
  return EditorDocument.open((await session.save()).bytes);
}

describe('what the file changed through Apply, kept by Word\'s saves (ED-005 section 12.47)', () => {
  // The file's parts were changed through E4.b's Apply, untracked (review-for-word.mjs, section 11): Heading 1's
  // colour, the numbering format, the default tab stop. Each of Word's saves keeps all three.
  for (const name of ['word2010', 'word15', 'word2010-accepted', 'word15-accepted', 'word2010-rejected', 'word15-rejected', 'word365-rejected']) {
    test(`review-for-word-${name}.docx`, async () => {
      const bytes = await fixture(`review/review-for-word-${name}.docx`);
      const session = await EditorDocument.open(bytes);
      const heading = session.pkg.propertyResolverOrUndefined!.getStyle('Heading1');
      expect(heading?.rPr?.color?.val).toBe('C00000');
      expect(session.defaultTabStop).toBe(1440);
      const numbering = new TextDecoder().decode(new ZipPartStore(bytes).loadSync('word/numbering.xml'));
      expect(numbering).toMatch(/w:numFmt w:val="upperRoman"/);
      expect(numbering).not.toMatch(/w:numFmt w:val="decimal"/);
    });
  }
});

describe('the editor\'s Accept all and Reject all against Word\'s, over the editor\'s own file (ED-005 section 12.34)', () => {
  test('the comparison sees a difference: Word\'s accepted save is not its rejected one', async () => {
    const accepted = await EditorDocument.open(await fixture('review/review-for-word-word15-accepted.docx'));
    const rejected = await EditorDocument.open(await fixture('review/review-for-word-word15-rejected.docx'));
    expect(shape(accepted.doc)).not.toEqual(shape(rejected.doc));
  });

  for (const action of ['accept', 'reject'] as const) {
    const suffix = action === 'accept' ? 'accepted' : 'rejected';
    for (const word of action === 'reject' ? (['word2010', 'word15', 'word365'] as const) : (['word2010', 'word15'] as const)) {
      test(`${action} all: the same as ${word === 'word2010' ? 'Word 2010' : word === 'word15' ? 'Word 15' : 'Word 365'}'s save`, async () => {
        const mine = await resolvedByTheEditor(action);
        const theirs = await EditorDocument.open(await fixture(`review/review-for-word-${word}-${suffix}.docx`));
        expect(trackedChangesOf(mine.doc)).toEqual([]);
        expect(trackedChangesOf(theirs.doc)).toEqual([]);
        let myShape = shape(mine.doc);
        let myComments = await comments(mine);
        if (word === 'word365') {
          // The known difference (core-ts 277d118, CR-002 section 54; ED-005 section 12.49 addendum): Word 365's Reject
          // All removes the agent's comment, whose markers core-ts 0.3.0 put outside the w:ins, where Word 2010 and
          // Word 15 kept it as its reference alone and the editor keeps it as they do. Whether the editor follows
          // Word 365 is Jason's decision; until then the difference is pinned here, exactly, and nothing else differs.
          const reference = '{[opaqueInline:commentReference]}';
          const its = (b: unknown): boolean => { const runs = (b as { runs?: string[] }).runs; return runs?.[0] === reference && runs[1] === 'Notice must be given within 30 days.{}'; };
          expect(myShape.filter(its)).toHaveLength(1);
          myShape = myShape.map((b) => its(b) ? { ...(b as object), runs: (b as { runs: string[] }).runs.slice(1) } : b);
          expect(myComments).toContain('Claude: Changed from arrears, as agreed.');
          myComments = myComments.filter((l) => l !== 'Claude: Changed from arrears, as agreed.');
        }
        expect(myShape).toEqual(shape(theirs.doc));
        // Word 2010 keeps no threads: a reply is a comment of its own, and none is done (section 12.23).
        const flat = (lines: string[]): string[] => lines.map((l) => l.replace(/^ {2}reply /, '').replace(/ \(done\)/, ''));
        expect(word === 'word2010' ? flat(myComments) : myComments).toEqual(await comments(theirs));
        expect(await footnotes(mine)).toEqual(await footnotes(theirs));
      });
    }
  }
});
