// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 6.6 item 5's model half (step D5): the companion's answer adopted through the
// tree run. The double moves every page-break marker one paragraph down; after the reload the
// text is what it was, the markers are where the double put them, the notes' references still
// resolve, one undo step takes the run back, and the export carries the new markers.
import { describe, test, expect } from 'vitest';
import { EditorDocument } from '../src/index.mjs';
import { fixture, textOf } from './helpers.mjs';
import { paginateDouble } from './fixtures/build/companion-double.mjs';

/** The indexes of the top-level blocks holding a page-break marker. */
function markerBlocks(session: EditorDocument): number[] {
  const out: number[] = [];
  session.doc.forEach((block, _offset, index) => {
    block.descendants((n) => { if (n.type.name === 'opaqueInline' && n.attrs['kind'] === 'lastRenderedPageBreak' && !out.includes(index)) out.push(index); return true; });
  });
  return out;
}

describe('adoptPaginated', () => {
  test('notes-numbering.docx: the marker moves down a paragraph, the text and the references stay, undo restores, the save carries the new markers', async () => {
    const original = await fixture('notes-numbering.docx');
    const session = await EditorDocument.open(original);
    await session.stories.footnotes();
    const before = markerBlocks(session);
    expect(before).toEqual([4]);
    const texts = Array.from({ length: session.doc.childCount }, (_, i) => textOf(session.doc.child(i)));
    const saved = await session.save({ selfCheck: false });
    const answer = await paginateDouble(saved.bytes);
    const result = await session.adoptPaginated(answer);
    expect(result.error).toBeUndefined();
    expect(result.changed).toBe(true);
    expect(session.paginationSource).toBe('companion');
    expect(markerBlocks(session)).toEqual([5]);
    expect(Array.from({ length: session.doc.childCount }, (_, i) => textOf(session.doc.child(i)))).toEqual(texts);
    // The references resolve to the notes as before (their ids are new, the note ids the same).
    for (const id of [1, 2, 3, 4, 5]) expect(session.stories.referenceOf(session.doc, 'footnote', id)).toBeDefined();
    expect(session.stories.orderedNotes('footnote', session.doc).map((n) => n.id)).toEqual([1, 2, 3, 4, 5]);
    // The main part is live (a tree run is an edit); the saved document carries the moved marker.
    expect(session.treeLive).toBe(true);
    const after = await session.save();
    expect(after.selfCheck?.equal).toBe(true);
    const reopened = await EditorDocument.open(after.bytes);
    expect(markerBlocks(reopened)).toEqual([5]);
    // One undo step takes the run back.
    session.undoManager.undo();
    expect(markerBlocks(session)).toEqual([4]);
    expect(session.doc.eq(session.imported)).toBe(false);   // the baseline is the paginated tree; the screen is the old one
  });

  test('a document without markers gains none from the double, and the run changes nothing on screen', async () => {
    const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
    const saved = await session.save({ selfCheck: false });
    const result = await session.adoptPaginated(await paginateDouble(saved.bytes));
    expect(result.error).toBeUndefined();
    expect(result.changed).toBe(false);
    expect(markerBlocks(session)).toEqual([]);
  });
});
