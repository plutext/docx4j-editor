// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.75: an XHTML-bound control shown. The node's markup is converted into the
// control (a block-level one takes the blocks, a run-level one the first paragraph's inlines),
// out of the undo history; an edit inside writes nothing back; an unchanged value converts nothing.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, isXhtmlControl } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { refreshXhtmlControls } from '../src/view/xhtmlbind.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');
if (typeof Element !== 'undefined') Element.prototype.getClientRects ??= () => ([] as unknown as DOMRectList);
const escape = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

async function open(markup: string): Promise<{ session: EditorDocument; node: () => string; setNode: (v: string) => void }> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'HelloWordOnline.docx'))));
  const part = session.pkg.customXmlParts.add(`<data><body>${escape(markup)}</body></data>`);
  await session.template.load();
  return { session, node: () => part.selectSingleNode('/data/body')!.text, setNode: (v) => { part.selectSingleNode('/data/body')!.text = v; } };
}

describe('XHTML-bound controls (ED-005 section 12.75)', () => {
  test('a block-level control takes the blocks: a heading, bold, a list; typing inside writes nothing back; the same value again converts nothing', async () => {
    const { session, node } = await open('<h1>Title</h1><p>Hello <b>bold</b></p><ul><li>one</li><li>two</li></ul>');
    const para = session.doc.child(0);
    expect(runCommand(session, 'contentControl.insert', {}, { from: 1, to: 1 + para.content.size }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/body' }, { from: 2 }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.xhtml', { on: true }, { from: 2 }).ok).toBe(true);
    expect(isXhtmlControl(session.doc.child(0).attrs['sdtPr'])).toBe(true);
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const first = await refreshXhtmlControls(editor.view, session);
    expect(first.changed).toBe(1);
    const control = (): PMNode => editor.state.doc.child(0);
    expect(control().type.name).toBe('sdtBlock');
    const texts = (): string[] => { const out: string[] = []; control().forEach((n) => out.push(n.textContent)); return out; };
    expect(texts()).toEqual(['Title', 'Hello bold', 'one', 'two']);
    expect(control().child(0).attrs['pStyle']).toBe('Heading1');
    expect(control().child(2).attrs['numPr']).toBeTruthy();
    let bold = false;
    control().child(1).descendants((n) => { if (n.isText && n.marks.some((m) => m.type.name === 'bold')) bold = true; return true; });
    expect(bold).toBe(true);
    // Nothing is written back: the node is as it was after an edit inside the control.
    const before = node();
    editor.caretAt(control().child(1).nodeSize + 3);
    editor.view.dispatch(editor.state.tr.insertText('X'));
    await new Promise((r) => setTimeout(r, 10));
    expect(node()).toBe(before);
    // The value unchanged: nothing converts, the edit stays.
    const again = await refreshXhtmlControls(editor.view, session);
    expect(again.changed).toBe(0);
    expect(editor.state.doc.child(0).textContent).toContain('X');
    editor.destroy();
  });

  test('a run-level control takes the first paragraph\'s inline content and reports the rest; a changed node value is converted again', async () => {
    const { session, setNode } = await open('<p>First <i>one</i></p><p>Second</p>');
    expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: 1, to: 4 }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/body' }, { from: 2 }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.xhtml', { on: true }, { from: 2 }).ok).toBe(true);
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const result = await refreshXhtmlControls(editor.view, session);
    expect(result.changed).toBe(1);
    expect(result.notes).toHaveLength(1);
    const inline = (): PMNode => editor.state.doc.child(0).child(0);
    expect(inline().type.name).toBe('sdtInline');
    expect(inline().textContent).toBe('First one');
    setNode('<p>Changed</p>');
    session.template.invalidate();
    const next = await refreshXhtmlControls(editor.view, session);
    expect(next.changed).toBe(1);
    expect(inline().textContent).toBe('Changed');
    editor.destroy();
  });
});
