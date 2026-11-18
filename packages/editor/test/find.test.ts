// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 3.6: the find plugin's hits are memoised per top-level block, so a keystroke
// searches only the block it changed and maps or patches the highlights; a query matching
// everywhere keeps its count exact and draws the thousand nearest the current hit; the current
// hit follows its text through an edit; every surface made while a query is active starts with it.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { TextSelection } from 'prosemirror-state';
import { EditorDocument, Finder, findIn } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { EditableStoryView } from '../src/view/StoryView.js';
import { findKey, FIND_META, DRAWN, hitAt, setActiveFinder, type FindMeta } from '../src/view/plugins/find.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

async function open(name: string): Promise<{ session: EditorDocument; editor: DocxEditorView }> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));
  const container = document.createElement('div');
  document.body.appendChild(container);
  return { session, editor: new DocxEditorView(container, session, { headersAndFooters: false }) };
}

function search(editor: DocxEditorView, finder: Finder | null, meta: Omit<FindMeta, 'finder'> = {}): void {
  editor.view.dispatch(editor.state.tr.setMeta(FIND_META, { finder, ...meta } satisfies FindMeta));
}

function blockStart(editor: DocxEditorView, i: number): number {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += editor.state.doc.child(k).nodeSize;
  return pos;
}

describe('the find plugin', () => {
  test('the hits and the current one; a keystroke searches its block alone and the highlights follow', async () => {
    const { session, editor } = await open('tracked-changes.docx');
    const finder = new Finder('indented', {}, session.ctx.ids);
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1)));
    search(editor, finder);
    let st = findKey.getState(editor.state)!;
    expect(st.total).toBe(20);
    expect(st.current).toBe(0);
    expect(st.set.find().length).toBe(20);
    expect(st.set.find().filter((d) => (d.spec as object) && (d as unknown as { type: { attrs: { class: string } } }).type.attrs.class === 'find-hit current').length).toBe(1);
    expect(editor.view.dom.querySelectorAll('.find-hit').length).toBe(20);
    const misses = finder.misses;
    // A keystroke in block 4, which holds no hit: one block searched, the set mapped.
    editor.view.dispatch(editor.state.tr.insertText('x', blockStart(editor, 4) + 2));
    expect(finder.misses).toBe(misses + 1);
    st = findKey.getState(editor.state)!;
    expect(st.total).toBe(20);
    // A keystroke that makes a hit: counted and drawn.
    editor.view.dispatch(editor.state.tr.insertText(' indented', blockStart(editor, 4) + 2));
    st = findKey.getState(editor.state)!;
    expect(st.total).toBe(21);
    expect(editor.view.dom.querySelectorAll('.find-hit').length).toBe(21);
    // The current hit (the first of block 17, "Indented") follows its text: the new hit before it makes it the second.
    expect(st.current).toBe(1);
    const current = hitAt(st, st.current)!;
    expect(editor.state.doc.textBetween(current.from, current.to)).toBe('Indented');
    expect(current.from).toBeGreaterThan(blockStart(editor, 17));
    search(editor, null);
    expect(findKey.getState(editor.state)!.total).toBe(0);
    expect(editor.view.dom.querySelectorAll('.find-hit').length).toBe(0);
    editor.destroy();
  });

  test('a query matching everywhere keeps its count and draws the nearest thousand', async () => {
    const { session, editor } = await open('fonts-cjk-linebox.docx');
    const finder = new Finder('e', {}, session.ctx.ids);
    const expected = findIn(editor.state.doc, 'e', {}, session.ctx.ids).length;
    expect(expected).toBeGreaterThan(DRAWN);
    search(editor, finder, { current: expected - 1 });
    const st = findKey.getState(editor.state)!;
    expect(st.total).toBe(expected);
    const drawn = st.set.find();
    expect(drawn.length).toBe(DRAWN);
    // The window holds the current (the last) hit.
    const last = hitAt(st, expected - 1)!;
    expect(drawn.some((d) => d.from === last.from && d.to === last.to)).toBe(true);
    editor.destroy();
  });

  test('a story view made while a query is active highlights it from the start', async () => {
    const { session, editor } = await open('footnote.docx');
    const notes = await session.stories.footnotes();
    const note = notes[0]!;
    const word = /\p{L}{3,}/u.exec(note.doc.textContent)![0];
    const finder = new Finder(word, {}, session.ctx.ids);
    setActiveFinder(session, finder);
    const container = document.createElement('div');
    document.body.appendChild(container);
    const view = new EditableStoryView(container, session, note, { stylesheet: editor });
    expect(findKey.getState(view.state)!.total).toBeGreaterThan(0);
    setActiveFinder(session, null);
    view.destroy();
    editor.destroy();
  });
});
