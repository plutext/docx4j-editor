// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.55: Word's mapping both ways, through the view. Typing in a bound control writes its node, and the
// other control bound to that node follows once the dispatch is over; an undo on the page writes the old value back
// (the view sees the undo as the undo manager's, `isUndoRedoOperation`); the refresh is no undo step of its own.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

const zeroRect = (): DOMRect => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON: () => ({}) }) as DOMRect;
if (typeof Range !== 'undefined') {
  Range.prototype.getClientRects ??= () => ([] as unknown as DOMRectList);
  Range.prototype.getBoundingClientRect ??= zeroRect;
}
if (typeof Element !== 'undefined') Element.prototype.getClientRects ??= () => ([] as unknown as DOMRectList);

const settle = (): Promise<void> => new Promise((r) => setTimeout(r, 0));

describe('two-way binding through the view (ED-005 section 12.55)', () => {
  test('typing in one control writes the node and the other control follows; an undo on the page writes the old value back', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'HelloWordOnline.docx'))));
    const part = session.pkg.customXmlParts.add('<data><name>Ann</name></data>');
    await session.template.load();
    // Two plain text controls over "Hel" and "lo", both bound to /data/name.
    expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: 1, to: 4 }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/name' }, { from: 3 }).ok).toBe(true);
    const second = 1 + session.doc.child(0).child(0).nodeSize;
    expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: second, to: second + 2 }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/name' }, { from: second + 1 }).ok).toBe(true);
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const texts = (doc: PMNode): string[] => { const out: string[] = []; doc.child(0).forEach((n) => { if (n.type.name === 'sdtInline') out.push(n.textContent); }); return out; };
    const node = (): string => part.selectSingleNode('/data/name')!.text;
    expect(texts(editor.state.doc)).toEqual(['Ann', 'Ann']);
    // The end of the first control's text: after its bookmark's two markers and "Ann".
    let end = -1;
    editor.state.doc.child(0).child(0).descendants((n, p) => { if (n.isText) end = 1 + 1 + p + n.nodeSize; return true; });
    session.undoManager.stopCapturing();
    editor.view.dispatch(editor.state.tr.insertText('e', end));
    expect(node()).toBe('Anne');
    await settle();
    expect(texts(editor.state.doc)).toEqual(['Anne', 'Anne']);
    expect(texts(session.doc)).toEqual(['Anne', 'Anne']);
    // One undo: the typing goes, the node is written back, and the other control follows.
    session.undoManager.undo();
    await settle();
    expect(texts(editor.state.doc)[0]).toBe('Ann');
    expect(node()).toBe('Ann');
    expect(texts(editor.state.doc)).toEqual(['Ann', 'Ann']);
    editor.destroy();
  });
});
