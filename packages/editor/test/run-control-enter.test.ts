// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.123 (Jason, 2026-11-12, tmp/testcase_sdt-replicates.docx): Enter through the view's keymaps never splits
// a run-level control. With the caret at the control's inner end, the Word keymap's `insertParagraph` refused (the caret was
// not directly in a paragraph) and ProseMirror's `splitBlock` split the control, an empty copy with its `w:id` in the new
// paragraph.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { TextSelection } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, readControl } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

/** A view over a document whose first long paragraph ends with a control of `kind` over its last six characters. */
async function controlAtEnd(kind: string): Promise<{ editor: DocxEditorView; paragraph: number }> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'HelloWordOnline.docx'))));
  const container = document.createElement('div');
  document.body.appendChild(container);
  const editor = new DocxEditorView(container, session);
  let at: { pos: number; node: PMNode; index: number } | undefined;
  editor.state.doc.forEach((node, pos, index) => { if (!at && node.type.name === 'paragraph' && node.textContent.length > 10) at = { pos, node, index }; });
  const end = at!.pos + 1 + at!.node.content.size;
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, end - 6, end)));
  expect(editor.run('contentControl.insert', { kind })).toBe(true);
  return { editor, paragraph: at!.index };
}

/** The control's place: where its content starts and ends. */
function place(doc: PMNode): { start: number; end: number } {
  let out: { start: number; end: number } | undefined;
  doc.descendants((n, pos) => { if (!out && n.type.name === 'sdtInline') out = { start: pos + 1, end: pos + 1 + n.content.size }; return !out; });
  return out!;
}

const ids = (doc: PMNode): (number | undefined)[] => { const out: (number | undefined)[] = []; doc.descendants((n) => { if (n.type.name === 'sdtInline') out.push(readControl(n.attrs['sdtPr']).id); return true; }); return out; };

function pressEnter(editor: DocxEditorView): void {
  editor.view.someProp('handleKeyDown', (f) => f(editor.view, new KeyboardEvent('keydown', { key: 'Enter' })));
}

describe('Enter through the keymaps and a run-level control (ED-005 section 12.123)', () => {
  test("at the control's inner end: a new empty paragraph, and no copy of the control", async () => {
    const { editor, paragraph } = await controlAtEnd('PlainText');
    const before = ids(editor.state.doc);
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, place(editor.state.doc).end)));
    expect(editor.state.selection.$from.parent.type.name).toBe('sdtInline');
    pressEnter(editor);
    expect(ids(editor.state.doc)).toEqual(before);
    expect(editor.state.doc.child(paragraph + 1).childCount).toBe(0);
    expect(editor.state.selection.$from.parent).toBe(editor.state.doc.child(paragraph + 1));
  });

  test('in the middle of a rich text control: nothing, where ProseMirror would split it', async () => {
    const { editor } = await controlAtEnd('RichText');
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, place(editor.state.doc).start + 3)));
    const doc = editor.state.doc;
    pressEnter(editor);
    expect(editor.state.doc.eq(doc)).toBe(true);
  });
});
