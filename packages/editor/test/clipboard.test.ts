// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-010: the editor's own clipboard, end to end in jsdom over the Word 365 probe of section 9.1
// (editor-model/test/fixtures/clipboard/clipboard-probe.docx): copy writes the slice's JSON beside
// ProseMirror's HTML, stamped with the session; the paste rebuilds it, the copies made (a comment
// copied whole, a note, a picture), and the save holds what Word's own paste gave.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { TextSelection } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { unzipSync, strFromU8 } from 'fflate';
import { EditorDocument, CLIPBOARD_SLICE_ATTR, CLIPBOARD_SESSION_ATTR } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { editorSliceOf } from '../src/view/clipboard.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures', 'clipboard');

async function editorOver(name = 'clipboard-probe.docx'): Promise<{ session: EditorDocument; editor: DocxEditorView }> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));
  const container = document.createElement('div');
  document.body.appendChild(container);
  const editor = new DocxEditorView(container, session, { headersAndFooters: false });
  return { session, editor };
}

const blocks = (doc: PMNode): PMNode[] => { const out: PMNode[] = []; doc.forEach((n) => out.push(n)); return out; };
const blockStarting = (doc: PMNode, text: string): { node: PMNode; pos: number } => {
  let found: { node: PMNode; pos: number } | undefined;
  doc.forEach((n, pos) => { if (!found && n.textContent.startsWith(text)) found = { node: n, pos }; });
  if (!found) throw new Error(`no block starting "${text}"`);
  return found;
};

/** The clipboard's HTML for a selection, as ProseMirror's copy writes it through the plugin's serializer. */
function copy(editor: DocxEditorView, from: number, to: number): string {
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)));
  const { dom } = editor.view.serializeForClipboard(editor.state.selection.content());
  return dom.innerHTML;
}

/** A paste event of `html` at `pos`, through the plugin's handler; the document after it settles. */
async function paste(editor: DocxEditorView, html: string, pos: number): Promise<boolean> {
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos)));
  const event = { clipboardData: { getData: (type: string) => (type === 'text/html' ? html : type === 'text/plain' ? 'plain' : '') }, preventDefault: () => {} } as unknown as ClipboardEvent;
  const before = editor.state.doc;
  const handled = editor.view.someProp('handlePaste', (f) => f(editor.view, event, editor.state.selection.content())) ?? false;
  for (let i = 0; i < 200 && editor.state.doc === before; i++) await new Promise((r) => setTimeout(r, 5));
  return handled;
}

const endOf = (doc: PMNode, text: string): number => { const { node, pos } = blockStarting(doc, text); return pos + node.nodeSize - 1; };

function documentXml(bytes: Uint8Array): string {
  return strFromU8(unzipSync(bytes)['word/document.xml']!);
}

describe('the carrier (ED-010 section 3.1)', () => {
  test('copy writes the slice as JSON and the session token beside data-pm-slice; another token falls through', async () => {
    const { session, editor } = await editorOver();
    const { node, pos } = blockStarting(editor.state.doc, 'Case 9');
    const html = copy(editor, pos + 1, pos + node.nodeSize - 1);
    expect(html).toContain('data-pm-slice');
    expect(html).toContain(CLIPBOARD_SLICE_ATTR);
    expect(html).toContain(`${CLIPBOARD_SESSION_ATTR}="${session.clipboardToken}"`);
    const json = editorSliceOf(html, session.clipboardToken) as { content: { type: string; attrs: { pStyle: string } }[] };
    expect(json.content[0]!.type).toBe('paragraph');
    expect(json.content[0]!.attrs.pStyle).toBe('Heading2');
    expect(editorSliceOf(html, 'another-session')).toBeNull();
    expect(editorSliceOf('<p data-pm-slice="1 1 []">x</p>', session.clipboardToken)).toBeNull();
  });
});

describe('the paste (ED-010 sections 3.2 and 3.3, Word 365 section 9.1)', () => {
  test('a heading pasted whole keeps its style through copiedFrom; its words alone take the target (case 9)', async () => {
    const { session, editor } = await editorOver();
    const heading = blockStarting(editor.state.doc, 'Case 9');
    // whole: the paragraph with its mark
    const whole = copy(editor, heading.pos, heading.pos + heading.node.nodeSize);
    expect(await paste(editor, whole, endOf(editor.state.doc, '[T9'))).toBe(true);
    const pasted = blocks(editor.state.doc).find((b, i, all) => i > 0 && all[i - 1]!.textContent.startsWith('[T9') && b.textContent.startsWith('Case 9'))!;
    expect(pasted).toBeDefined();
    expect(pasted.attrs['pStyle']).toBe('Heading2');
    expect(pasted.attrs['sourceId']).toBeNull();
    expect(pasted.attrs['copiedFrom']).toBe(heading.node.attrs['sourceId']);
    // P8: the target keeps its own properties, no empty paragraph after
    const target = blockStarting(editor.state.doc, '[T9');
    expect(target.node.attrs['pStyle']).toBeNull();
    expect(target.node.textContent).toBe('[T9: paste case 9 here, a plain paragraph: paste the heading words into it]');
    // the words alone into a plain paragraph: the target's style kept (case 9)
    const words = copy(editor, heading.pos + 1 + 'Case 9: '.length, heading.pos + 1 + 'Case 9: HEADING TWO TEXT'.length);
    expect(await paste(editor, words, endOf(editor.state.doc, '[T8'))).toBe(true);
    const t8 = blockStarting(editor.state.doc, '[T8');
    expect(t8.node.textContent.endsWith('HEADING TWO TEXT')).toBe(true);
    expect(t8.node.attrs['pStyle']).toBeNull();
    // the save holds the heading's style for the pasted paragraph and the plain one for the target
    const saved = await session.save({ selfCheck: false });
    expect(saved.changed).toBe(true);
    const xml = documentXml(saved.bytes);
    expect(xml.match(/<w:pStyle w:val="Heading2"\/>/g)?.length).toBe(2);
    expect(xml.match(/HEADING TWO TEXT/g)?.length).toBe(3);   // the source, the heading pasted whole, the words
  });

  test('a numbered item pasted whole keeps its numbering (P9: the list continues); bold stays', async () => {
    const { session, editor } = await editorOver();
    const item = blockStarting(editor.state.doc, 'Case 8');
    const html = copy(editor, item.pos, item.pos + item.node.nodeSize);
    expect(await paste(editor, html, endOf(editor.state.doc, '[T8'))).toBe(true);
    const pasted = blocks(editor.state.doc).filter((b) => b.textContent.startsWith('Case 8'));
    expect(pasted).toHaveLength(2);
    expect(pasted[1]!.attrs['numPr']).toEqual(item.node.attrs['numPr']);
    expect(pasted[1]!.attrs['pStyle']).toBe('ListParagraph');
    const xml = documentXml((await session.save({ selfCheck: false })).bytes);
    expect(xml.match(/<w:numId w:val="1"\/>/g)?.length).toBe(2);
  });

  test('a bookmark is not copied (case 5); the tracked paragraph keeps Author B\'s revisions with tracking off (case 3)', async () => {
    const { session, editor } = await editorOver();
    const marked = blockStarting(editor.state.doc, 'Case 5');
    expect(await paste(editor, copy(editor, marked.pos, marked.pos + marked.node.nodeSize), endOf(editor.state.doc, '[T5'))).toBe(true);
    const pasted = blocks(editor.state.doc).filter((b) => b.textContent.startsWith('Case 5'));
    expect(pasted).toHaveLength(2);
    const atoms = (b: PMNode): string[] => { const out: string[] = []; b.descendants((n) => { if (n.type.name === 'opaqueInline') out.push(String(n.attrs['kind'])); return true; }); return out; };
    expect(atoms(pasted[0]!)).toEqual(['bookmarkStart', 'bookmarkEnd']);
    expect(atoms(pasted[1]!)).toEqual([]);

    const tracked = blockStarting(editor.state.doc, 'Case 3');
    expect(await paste(editor, copy(editor, tracked.pos, tracked.pos + tracked.node.nodeSize), endOf(editor.state.doc, '[T3'))).toBe(true);
    const both = blocks(editor.state.doc).filter((b) => b.textContent.startsWith('Case 3'));
    expect(both).toHaveLength(2);
    const revisions = (b: PMNode): string[] => { const out: string[] = []; b.descendants((n) => { for (const m of n.marks) if (m.type.name === 'insertion' || m.type.name === 'deletion') out.push(`${m.type.name}:${m.attrs['author']}:${n.text}`); return true; }); return out; };
    expect(revisions(both[1]!)).toEqual(revisions(both[0]!));
    expect(revisions(both[1]!)).toEqual(['insertion:Author B:INSERTED ', 'deletion:Author B:DELETED ']);
    const xml = documentXml((await session.save({ selfCheck: false })).bytes);
    expect(xml.match(/<w:bookmarkStart /g)?.length).toBe(1);
    expect(xml.match(/w:author="Author B"/g)?.length).toBe(4);
  });

  test('with tracking on the paste is one insertion of the accepted text by this author, the pasted mark included (case 4)', async () => {
    const { session, editor } = await editorOver();
    const tracked = blockStarting(editor.state.doc, 'Case 3');
    const html = copy(editor, tracked.pos, tracked.pos + tracked.node.nodeSize);
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author A';
    expect(await paste(editor, html, endOf(editor.state.doc, '[T4'))).toBe(true);
    const pasted = blocks(editor.state.doc).filter((b) => b.textContent.startsWith('Case 3'))[1]!;
    expect(pasted.textContent).toBe('Case 3 and 4: before INSERTED after. (copy the whole paragraph, its mark included)');
    const authors = new Set<string>();
    pasted.descendants((n) => { for (const m of n.marks) if (m.type.name === 'insertion' || m.type.name === 'deletion') authors.add(`${m.type.name}:${m.attrs['author']}`); return true; });
    expect([...authors]).toEqual(['insertion:Author A']);
    expect((pasted.attrs['markRevision'] as { insertion?: { author: string } } | null)?.insertion?.author).toBe('Author A');
    const xml = documentXml((await session.save({ selfCheck: false })).bytes);
    // the source's deletion alone: the paste carries none
    expect(xml.match(/DELETED/g)?.length).toBe(1);
    expect(xml.match(/INSERTED/g)?.length).toBe(2);
    expect(xml.match(/w:author="Author A"/g)!.length).toBeGreaterThanOrEqual(2);
  });

  test('a comment copied whole is made again with its author (case 1); copied in part it goes (case 2)', async () => {
    const { session, editor } = await editorOver();
    await session.stories.comments();
    const whole = blockStarting(editor.state.doc, 'Case 1');
    const text = 'Case 1: '.length;
    const html = copy(editor, whole.pos + 1 + text, whole.pos + 1 + text + 'COMMENTED WHOLE'.length + 1);   // the text and its reference atom
    expect(await paste(editor, html, endOf(editor.state.doc, '[T1'))).toBe(true);
    const t1 = blockStarting(editor.state.doc, '[T1');
    const ids = new Set<string>();
    t1.node.descendants((n) => { for (const m of n.marks) if (m.type.name === 'comment') ids.add(String(m.attrs['id'])); return true; });
    expect(ids.size).toBe(1);
    const [newId] = [...ids];
    expect(newId).not.toBe('0');
    const threads = await session.stories.comments();
    const made = threads.find((t) => t.id === newId)!;
    expect(made).toBeDefined();
    expect(made.author).toBe('Author B');
    expect(session.stories.currentDoc(made.record).textContent).toContain('Comment on case 1.');
    // part: the word PART from inside case 2's range
    const part = blockStarting(editor.state.doc, 'Case 2');
    const at = part.pos + 1 + 'Case 2: COMMENTED '.length;
    expect(await paste(editor, copy(editor, at, at + 'PART'.length), endOf(editor.state.doc, '[T2'))).toBe(true);
    const t2 = blockStarting(editor.state.doc, '[T2');
    expect(t2.node.textContent.endsWith('PART')).toBe(true);
    let marks = 0;
    t2.node.descendants((n) => { marks += n.marks.filter((m) => m.type.name === 'comment').length; return true; });
    expect(marks).toBe(0);
    const xml = documentXml((await session.save({ selfCheck: false })).bytes);
    expect(xml.match(/<w:commentReference /g)?.length).toBe(3);
  });

  test('a picture pasted is a copy with a fresh docPr id over the same image part (case 6); a footnote a new note with the same text (case 7)', async () => {
    const { session, editor } = await editorOver();
    await session.stories.footnotes();
    const picture = blockStarting(editor.state.doc, 'Case 6');
    let imageAt = -1;
    picture.node.forEach((n, offset) => { if (n.type.name === 'image') imageAt = picture.pos + 1 + offset; });
    expect(imageAt).toBeGreaterThan(0);
    expect(await paste(editor, copy(editor, imageAt, imageAt + 1), endOf(editor.state.doc, '[T6'))).toBe(true);
    const images: PMNode[] = [];
    editor.state.doc.descendants((n) => { if (n.type.name === 'image') images.push(n); return true; });
    expect(images).toHaveLength(2);
    expect(images[1]!.attrs['sourceId']).not.toBe(images[0]!.attrs['sourceId']);

    const note = blockStarting(editor.state.doc, 'Case 7');
    expect(await paste(editor, copy(editor, note.pos, note.pos + note.node.nodeSize), endOf(editor.state.doc, '[T7'))).toBe(true);
    const refs: PMNode[] = [];
    editor.state.doc.descendants((n) => { if (n.type.name === 'footnoteRef') refs.push(n); return true; });
    expect(refs).toHaveLength(2);
    expect(refs[1]!.attrs['footnoteId']).not.toBe(refs[0]!.attrs['footnoteId']);
    const copied = session.stories.noteOf('footnote', Number(refs[1]!.attrs['footnoteId']))!;
    expect(session.stories.currentDoc(copied).textContent).toContain('The footnote of case 7.');

    const saved = await session.save({ selfCheck: false });
    const zip = unzipSync(saved.bytes);
    const xml = strFromU8(zip['word/document.xml']!);
    const docPr = [...xml.matchAll(/<wp:docPr id="(\d+)" name="([^"]*)"/g)].map((m) => [m[1], m[2]]);
    expect(docPr).toHaveLength(2);
    expect(docPr[0]).toEqual(['1', 'Picture 1']);
    expect(docPr[1]![0]).not.toBe('1');
    expect(docPr[1]![1]).toBe(`Picture ${docPr[1]![0]}`);
    expect(Object.keys(zip).filter((k) => k.startsWith('word/media/'))).toEqual(['word/media/image1.png']);
    const footnotes = strFromU8(zip['word/footnotes.xml']!);
    expect(footnotes.match(/The footnote of case 7\./g)?.length).toBe(2);
  });

  test('undo takes the paste back whole, the document as it was', async () => {
    const { session, editor } = await editorOver();
    const before = editor.state.doc;
    const item = blockStarting(editor.state.doc, 'Case 8');
    expect(await paste(editor, copy(editor, item.pos, item.pos + item.node.nodeSize), endOf(editor.state.doc, '[T8'))).toBe(true);
    expect(editor.state.doc.eq(before)).toBe(false);
    session.undoManager.undo();
    expect(editor.state.doc.eq(before)).toBe(true);
    expect((await session.save({ selfCheck: false })).changed).toBe(false);
  });
});
