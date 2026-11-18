// SPDX-License-Identifier: GPL-3.0-or-later
//
// core-ts check 33 (run by Jason in current Word, 2026-10-02; `fixtures/revisions/check33/`, copied from
// docx4j-core-ts's `test/fixtures/revisions/check33/`): a whole Heading 2 paragraph, its text and its mark, deleted
// with tracking on before a body paragraph (`33a-deleted.docx`), then Accept All (`33b-accepted.docx`). Word carries
// nothing: the body paragraph keeps its own properties, with no record. The carry of ED-005 section 12.33 item 3 is
// for a mark deleted alone, its paragraph's text kept (`word-measure.test.mts`). ED-005 section 12.51.
import { describe, expect, test } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, runCommandAsync, trackedChangesOf } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

/** Each paragraph's text, style, mark revision and properties record, as Word's files are compared. */
function shape(doc: PMNode): unknown[] {
  const out: unknown[] = [];
  doc.forEach((b) => {
    const a = b.attrs;
    const mark = a['markRevision'] as { insertion?: unknown; deletion?: unknown } | null;
    out.push({ text: b.textContent, style: a['pStyle'] ?? null, markDeleted: !!mark?.deletion, pPrChange: a['pPrChange'] !== null && a['pPrChange'] !== undefined });
  });
  return out;
}

describe('core-ts check 33: a whole paragraph deleted with tracking on hands the next paragraph nothing, as Word does', () => {
  test("the heading's text and mark deleted: as Word's save has it, then Accept All as Word's", async () => {
    const session = await EditorDocument.open(await fixture('revisions/check33/33-input.docx'));
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author A';
    // The heading is the second block; the selection runs from its first character to the next paragraph's, mark included.
    const heading = session.doc.child(0).nodeSize;
    const next = heading + session.doc.child(1).nodeSize;
    expect(runCommand(session, 'insertText', { text: '' }, { from: heading + 1, to: next + 1 }).ok).toBe(true);
    const word = await EditorDocument.open(await fixture('revisions/check33/33a-deleted.docx'));
    expect(shape(session.doc)).toEqual(shape(word.doc));
    expect(trackedChangesOf(session.doc).map((c) => c.type)).toEqual(['Deleted']);
    expect((await runCommandAsync(session, 'trackedChanges.acceptAll', {})).ok).toBe(true);
    const accepted = await EditorDocument.open(await fixture('revisions/check33/33b-accepted.docx'));
    expect(shape(session.doc)).toEqual(shape(accepted.doc));
  });
});
