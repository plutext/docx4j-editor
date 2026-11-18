// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 5.4 (step C3) in the view: the labels come from the labels plugin's node walk;
// a keystroke in a list item carries them forward without a walk, a new paragraph or a level
// change walks again; the widget takes the hanging indent and the paragraph the level's indent.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { TextSelection } from 'prosemirror-state';
import { EditorDocument, type ParagraphAttrs } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { labelsKey } from '../src/view/plugins/labels.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

async function open(name: string): Promise<{ session: EditorDocument; editor: DocxEditorView }> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));
  const container = document.createElement('div');
  document.body.appendChild(container);
  return { session, editor: new DocxEditorView(container, session, { headersAndFooters: false }) };
}

const labelTexts = (editor: DocxEditorView): string[] => Array.from(editor.container.querySelectorAll('.list-label')).map((el) => el.textContent ?? '');
const walks = (editor: DocxEditorView): number => labelsKey.getState(editor.state)!.walks;

function firstListItem(editor: DocxEditorView): { pos: number; end: number } {
  let found: { pos: number; end: number } | undefined;
  editor.state.doc.descendants((node, pos) => {
    if (found) return false;
    if (node.type.name === 'paragraph' && (node.attrs as ParagraphAttrs).numPr) found = { pos, end: pos + 1 + node.content.size };
    return !found;
  });
  return found!;
}

describe('live labels in the view', () => {
  test('typing carries the labels; Enter numbers the new item and renumbers the rest; detaching drops one', async () => {
    const { session, editor } = await open('numbering-stories.docx');
    expect(labelTexts(editor)).toEqual(['1.', '2.', '3.', '4.', '5.', '6.']);
    expect(walks(editor)).toBe(1);
    const item = firstListItem(editor);
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, item.end)));
    editor.run('insertText', { text: ' typed' });
    expect(walks(editor)).toBe(1);
    expect(labelTexts(editor)).toEqual(['1.', '2.', '3.', '4.', '5.', '6.']);
    editor.run('insertParagraph');
    expect(walks(editor)).toBe(2);
    expect(labelTexts(editor)).toEqual(['1.', '2.', '3.', '4.', '5.', '6.', '7.']);
    editor.run('list.indent');
    expect(walks(editor)).toBe(3);
    expect(labelTexts(editor)).toHaveLength(7);
    // The view's edits coalesce in the undo manager's capture window; the detach is its own step here.
    session.undoManager.stopCapturing();
    editor.run('paragraph.detachFromList');
    expect(labelTexts(editor)).toEqual(['1.', '2.', '3.', '4.', '5.', '6.']);
    editor.run('history.undo');
    expect(labelTexts(editor)).toHaveLength(7);
    editor.destroy();
  });

  test('the widget fills the hanging indent and the paragraph lays out at the level\'s indent', async () => {
    const { editor } = await open('numbering-stories.docx');
    const item = firstListItem(editor);
    const attrs = editor.state.doc.nodeAt(item.pos)!.attrs as ParagraphAttrs;
    expect(attrs.ind).toBeNull();   // the indent comes from the level alone
    const p = editor.container.querySelector('.ProseMirror p .list-label')!.parentElement as HTMLElement;
    expect(p.style.paddingInlineStart).toMatch(/pt$/);
    expect(p.style.textIndent).toMatch(/^-/);
    const label = p.querySelector('.list-label') as HTMLElement;
    expect(label.style.inlineSize).toBe(p.style.textIndent.slice(1));
    editor.destroy();
  });
});
