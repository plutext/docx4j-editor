// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 9.2, E4.g step G1: the files for Jason's session in Word open in the editor,
// round-trip unedited, and hold what their builders say. (The editor's file, made in a room with
// the agent, is the companion's test: `editor-mcp/test/word-files.test.mts`.) The builders run from
// the built package, as they run for the session.
import { describe, expect, test } from 'vitest';
import { EditorDocument, trackedChangesOf } from '../src/index.mjs';
import { comparePackages } from './helpers.mjs';
import { HEBREW, buildMeasureBase } from './fixtures/build/measure-base.mjs';
import { IDS, buildIdsForWord } from './fixtures/build/ids-for-word.mjs';

async function unedited(bytes: Uint8Array): Promise<void> {
  const session = await EditorDocument.open(bytes);
  const again = await session.save();
  expect(comparePackages(bytes, again.bytes)).toEqual([]);
}

describe('the measurements file (section 9.2 item 2)', () => {
  test('its one revision is Claude\'s pending insertion; the right-to-left and centred paragraphs are there; it round-trips', async () => {
    const bytes = await buildMeasureBase();
    const session = await EditorDocument.open(bytes);
    const changes = trackedChangesOf(session.doc);
    expect(changes).toHaveLength(1);
    expect(changes[0]).toMatchObject({ type: 'Added', author: 'Claude' });
    expect(changes[0]!.text).toContain('The tenant pays rent monthly in advance.');
    const paragraphs: { text: string; bidi: boolean; jc: string | null }[] = [];
    session.doc.forEach((n) => { if (n.type.name === 'paragraph') paragraphs.push({ text: n.textContent, bidi: n.attrs['bidi'] as boolean, jc: n.attrs['jc'] as string | null }); });
    expect(paragraphs.filter((p) => p.text === HEBREW)).toEqual([{ text: HEBREW, bidi: true, jc: null }, { text: HEBREW, bidi: true, jc: null }]);
    expect(paragraphs.find((p) => p.text.startsWith('This paragraph is centred'))?.jc).toBe('center');
    expect(paragraphs.some((p) => p.text === '')).toBe(true);
    await unedited(bytes);
  });
});

describe('the id bisection (section 9.2 item 3)', () => {
  test('a footnote and a comment at each id: the id where the file names it, and an unedited round trip', async () => {
    const files = await buildIdsForWord();
    expect(files.map((f) => f.name)).toEqual([...IDS.map((n) => `footnote-${n}.docx`), ...IDS.map((n) => `comment-${n}.docx`)]);
    for (const f of files) {
      const session = await EditorDocument.open(f.bytes);
      if (f.kind === 'footnote') {
        const refs: unknown[] = [];
        session.doc.descendants((n) => { if (n.type.name === 'footnoteRef') refs.push(n.attrs['footnoteId']); return true; });
        expect(refs.map(String), f.name).toEqual([String(f.id)]);
      } else {
        const threads = await session.stories.comments();
        expect(threads.map((t) => t.id), f.name).toEqual([String(f.id)]);
        expect(threads[0]!.author).toBe('Author A');
      }
      await unedited(f.bytes);
    }
  }, 60_000);
});
