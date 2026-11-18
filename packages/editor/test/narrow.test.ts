// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-004 section 6.3 (UPSTREAM-9): a whole-document replace narrowed to the blocks that differ,
// the document the same; and a colleague's keystroke reaching a view as a step over its block.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import * as Y from 'yjs';
import { Fragment, Slice, type Node as PMNode } from 'prosemirror-model';
import { EditorState, type Transaction } from 'prosemirror-state';
import { ReplaceStep } from 'prosemirror-transform';
import { EditorDocument, newEditStore, runCommand } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { narrowed } from '../src/view/narrow.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

describe('narrowed remote replaces (ED-004 section 6.3)', () => {
  test('a whole-document replace keeping every block but one becomes a step over that block, with the same document', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'tracked-changes.docx'))));
    const doc = session.doc;
    const blocks: PMNode[] = [];
    doc.forEach((b, _o, i) => { blocks.push(i === 3 ? b.type.create(b.attrs, doc.type.schema.text('Changed.')) : b); });
    const state = EditorState.create({ doc });
    const whole = state.tr.replace(0, doc.content.size, new Slice(Fragment.from(blocks), 0, 0));
    const narrow = narrowed(state.tr).replace(0, doc.content.size, new Slice(Fragment.from(blocks), 0, 0));
    expect(narrow.doc.eq(whole.doc)).toBe(true);
    const step = narrow.steps[0] as ReplaceStep;
    let from = 0;
    for (let i = 0; i < 3; i++) from += doc.child(i).nodeSize;
    expect(step.from).toBe(from);
    expect(step.to).toBe(from + doc.child(3).nodeSize);
    // The same blocks: no step at all. Any other replace is untouched.
    expect(narrowed(state.tr).replace(0, doc.content.size, new Slice(doc.content, 0, 0)).steps).toHaveLength(0);
    expect((narrowed(state.tr).replace(1, 3, Slice.empty).steps[0] as ReplaceStep).from).toBe(1);
  });

  test('a colleague\'s keystroke reaches the view as a step over its block, not the whole document', async () => {
    const bytes = new Uint8Array(await readFile(join(fixtures, 'tracked-changes.docx')));
    const creator = await EditorDocument.open(bytes);
    await creator.share(bytes, 'x.docx');
    const store = newEditStore();
    Y.applyUpdate(store.ydoc, Y.encodeStateAsUpdate(creator.store.ydoc));
    const tag = {};
    creator.store.ydoc.on('update', (u: Uint8Array, o: unknown) => { if (o !== tag) Y.applyUpdate(store.ydoc, u, tag); });
    store.ydoc.on('update', (u: Uint8Array, o: unknown) => { if (o !== tag) Y.applyUpdate(creator.store.ydoc, u, tag); });
    const joiner = await EditorDocument.open(bytes, { store, mode: 'join' });
    const container = document.createElement('div');
    document.body.appendChild(container);
    const view = new DocxEditorView(container, creator);
    const seen: Transaction[] = [];
    view.subscribe((_s, trs) => { seen.push(...trs); });
    let pos = 0;
    for (let k = 0; k < 4; k++) pos += joiner.doc.child(k).nodeSize;
    expect(runCommand(joiner, 'insertText', { text: 'Remote ' }, { from: pos + 1 }).ok).toBe(true);
    const remote = seen.filter((tr) => tr.docChanged);
    expect(remote.length).toBeGreaterThan(0);
    const size = view.state.doc.content.size;
    for (const tr of remote) for (const s of tr.steps as ReplaceStep[]) {
      expect(s.from).toBeGreaterThan(0);
      expect(s.to).toBeLessThan(size);
    }
    expect(view.state.doc.child(4).textContent.startsWith('Remote ')).toBe(true);
    expect(view.state.doc.eq(joiner.doc)).toBe(true);
    view.destroy();
  });
});
