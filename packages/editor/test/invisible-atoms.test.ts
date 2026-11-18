// SPDX-License-Identifier: GPL-3.0-or-later
//
// A selection covering a paragraph's visible text takes the invisible atoms at its edges with it
// (ED-010 section 9.3): the markers before a heading's first character go with a cut of the heading,
// so the paste finds their names free and keeps them. `numbering-clauses.docx`'s clause headings
// start with an empty run and a `_Toc` bookmark, both drawn as nothing.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { TextSelection } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { isInvisible } from '../src/view/plugins/invisibleAtoms.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

async function editorOver(name: string): Promise<{ session: EditorDocument; editor: DocxEditorView }> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));
  const container = document.createElement('div');
  document.body.appendChild(container);
  return { session, editor: new DocxEditorView(container, session, { headersAndFooters: false }) };
}

/** Block `i`: its content's start, the first visible position, the last visible position, the content's end. */
function bounds(doc: PMNode, i: number): { start: number; visibleStart: number; visibleEnd: number; end: number } {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += doc.child(k).nodeSize;
  const p = doc.child(i);
  const start = pos + 1;
  const end = start + p.content.size;
  let visibleStart = start;
  for (let k = 0; k < p.childCount && isInvisible(p.child(k)); k++) visibleStart += p.child(k).nodeSize;
  let visibleEnd = end;
  for (let k = p.childCount - 1; k >= 0 && isInvisible(p.child(k)); k--) visibleEnd -= p.child(k).nodeSize;
  return { start, visibleStart, visibleEnd, end };
}

describe('a selection over a paragraph\'s visible text takes its invisible edges', () => {
  test('from the first visible character to the end: widened back over the leading markers; part of the text: not', async () => {
    const { editor } = await editorOver('numbering-clauses.docx');
    const doc = editor.state.doc;
    const b = bounds(doc, 1);
    expect(b.visibleStart).toBeGreaterThan(b.start);
    const set = (from: number, to: number): { from: number; to: number; anchor: number; head: number } => {
      editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)));
      const s = editor.state.selection;
      return { from: s.from, to: s.to, anchor: s.anchor, head: s.head };
    };
    // the visible text whole: the markers come in at the start
    expect(set(b.visibleStart, b.visibleEnd)).toMatchObject({ from: b.start, to: b.end });
    // backwards too, the direction kept
    expect(set(b.visibleEnd, b.visibleStart)).toMatchObject({ from: b.start, to: b.end, anchor: b.end, head: b.start });
    // to the next paragraph's start (the selection bar's shape): the same
    expect(set(b.visibleStart, b.end + 2)).toMatchObject({ from: b.start, to: b.end + 2 });
    // part of the text: left alone
    expect(set(b.visibleStart + 1, b.visibleEnd)).toMatchObject({ from: b.visibleStart + 1, to: b.visibleEnd });
    expect(set(b.visibleStart, b.visibleEnd - 1)).toMatchObject({ from: b.visibleStart, to: b.visibleEnd - 1 });
    // a caret after the markers stays where it is
    expect(set(b.visibleStart, b.visibleStart)).toMatchObject({ from: b.visibleStart, to: b.visibleStart });
    // from an earlier paragraph into the middle of this one: the end is not dragged to the paragraph's end
    expect(set(b.start - 3, b.visibleStart + 2)).toMatchObject({ to: b.visibleStart + 2 });
  });

  test('a cut from the first visible character takes the markers, and the paste keeps the bookmarks (tracking off)', async () => {
    const { session, editor } = await editorOver('numbering-clauses.docx');
    await session.setChangeTrackingMode('Off');
    const view = editor.view;
    const names = (p: PMNode): string[] => { const out: string[] = []; p.descendants((n) => { if (n.type.name === 'opaqueInline' && n.attrs['kind'] === 'bookmarkStart') out.push(String((session.ctx.ids.get(String(n.attrs['sourceId']))?.value as { name?: string } | undefined)?.name)); return true; }); return out; };
    const b = bounds(view.state.doc, 1);
    const before = names(view.state.doc.child(1));
    expect(before.length).toBeGreaterThan(0);
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, b.visibleStart, b.end + 2)));
    const slice = view.state.selection.content();
    const html = view.serializeForClipboard(slice).dom.innerHTML;
    view.dispatch(view.state.tr.deleteSelection());
    let left = 0;
    view.state.doc.descendants((n) => { if (n.type.name === 'opaqueInline' && n.attrs['kind'] === 'bookmarkStart' && before.includes(String((session.ctx.ids.get(String(n.attrs['sourceId']))?.value as { name?: string } | undefined)?.name))) left++; return true; });
    expect(left).toBe(0);
    let tr = view.state.tr.insert(view.state.doc.content.size, view.state.schema.nodes['paragraph']!.create());
    tr = tr.setSelection(TextSelection.create(tr.doc, tr.doc.content.size - 1));
    view.dispatch(tr);
    const event = { clipboardData: { getData: (t: string) => (t === 'text/html' ? html : '') }, preventDefault: () => {} } as unknown as ClipboardEvent;
    expect(view.someProp('handlePaste', (f) => f(view, event, slice))).toBe(true);
    await new Promise((r) => setTimeout(r, 300));
    const pasted = view.state.doc.child(view.state.doc.childCount - 1);
    expect(names(pasted)).toEqual(before);
    expect(pasted.attrs['pStyle']).toBe('736');
  });
});
