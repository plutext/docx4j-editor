// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-004 section 1.3 item 8: recovery after a tree run. The recovery banner reopens the file's
// original bytes and adopts the persisted store (ED-002 section 5.4); ids a tree run minted in
// the earlier session are unknown to the fresh projection, so the recovered session's export
// loses what the run made, or resolves an id to an element a story projected there later.
// Marked `test.fails` until E3.a's `sources` map fixes it (ED-004 section 3.7, the invariant in
// and out of a room); the fix removes `.fails`, never the expectations (ED-004 section 2.1).
import { describe, test, expect } from 'vitest';
import * as Y from 'yjs';
import { EditorDocument, newEditStore } from '../src/index.mjs';
import { fixture, comparePackages } from './helpers.mjs';

/** The persisted state as recovery finds it: the store's updates applied to a fresh `Y.Doc`. */
function persisted(session: EditorDocument): ReturnType<typeof newEditStore> {
  const ydoc = new Y.Doc();
  Y.applyUpdate(ydoc, Y.encodeStateAsUpdate(session.store.ydoc));
  return newEditStore(ydoc);
}

describe('recovery after a tree run (ED-004 section 1.3 item 8)', () => {
  test('a recovered session exports what the session did before the close', async () => {
    const bytes = await fixture('tracked-changes.docx');
    const first = await EditorDocument.open(bytes);
    // A paragraph with properties the editor does not model (the line-numbering and hyphenation
    // switches, an embossed run): what a rebuild from the builders would lose.
    const xml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:suppressLineNumbers/><w:suppressAutoHyphens/></w:pPr><w:r><w:rPr><w:emboss/></w:rPr><w:t>Inserted by a script</w:t></w:r></w:p>';
    const run = await first.runOnTree(async ({ body }) => {
      const p = body.paragraphs[4]!;
      p.insertText(' via a script', 'End');
      await body.insertXml(xml, 'After', p);
    });
    expect(run.error).toBeUndefined();
    const before = await first.save({ selfCheck: false });
    const recovered = await EditorDocument.open(bytes, { store: persisted(first) });
    expect(recovered.doc.eq(first.doc)).toBe(true);
    const after = await recovered.save({ selfCheck: false });
    expect(comparePackages(before.bytes, after.bytes)).toEqual([]);
  });

  test('a story read after recovery does not take over an id the tree run minted', async () => {
    const bytes = await fixture('footnote.docx');
    const first = await EditorDocument.open(bytes);
    await first.runOnTree(({ body }) => { body.paragraphs[0]!.insertParagraph('Inserted by a script', 'After'); });
    const before = await first.save({ selfCheck: false });
    const recovered = await EditorDocument.open(bytes, { store: persisted(first) });
    // Reading the notes projects them with the counter the run's ids came from (section 1.3 item 3).
    await recovered.stories.footnotes();
    const after = await recovered.save({ selfCheck: false });
    expect(comparePackages(before.bytes, after.bytes)).toEqual([]);
  });
});

describe('story ids (ED-004 section 3.6)', () => {
  test('a story gets the same ids whatever order a session reads the stories in', async () => {
    const bytes = await fixture('numbering-stories.docx');
    const a = await EditorDocument.open(bytes);
    const b = await EditorDocument.open(bytes);
    const ids = (s: { doc: { descendants(f: (n: { attrs: Record<string, unknown> }) => boolean | void): void } }): string[] => {
      const out: string[] = [];
      s.doc.descendants((n) => { const id = n.attrs['sourceId']; if (typeof id === 'string') out.push(id); });
      return out;
    };
    // One session reads the headers first, the other the notes first.
    const headersA = await a.stories.headers();
    const notesA = await a.stories.footnotes();
    const notesB = await b.stories.footnotes();
    const headersB = await b.stories.headers();
    expect(headersA.length).toBeGreaterThan(0);
    expect(notesA.length).toBeGreaterThan(0);
    expect(headersA.map(ids)).toEqual(headersB.map(ids));
    expect(notesA.map(ids)).toEqual(notesB.map(ids));
    expect(ids(headersA[0]!)[0]).toMatch(/^(header|footer):\/word\/[a-z]+\d*\.xml\/e\d+$/);
    expect(ids(notesA[0]!)[0]).toMatch(/^footnote:\d+\/e\d+$/);
  });
});
