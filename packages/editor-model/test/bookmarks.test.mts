// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.115, from Jason's save of docx4j's Getting Started guide: Shift+Down from the start of the paragraph
// before a heading ended after the heading's ten leading `w:bookmarkStart`s, and Delete took them and left their ends, so
// docx4j's TOC update refused the file. A marker removed while its partner stays goes back where it was removed.
import { describe, test, expect } from 'vitest';
import { Fragment, type Node as PMNode } from 'prosemirror-model';
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import { ySyncPluginKey } from 'y-prosemirror';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, LOCAL_ORIGIN, recordTransaction, runCommandAsync, stateFor } from '../src/index.mjs';
import { mainXml } from './helpers.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function documentOf(markup: string): Promise<EditorDocument> {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.getMainDocumentPart().setXml(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${markup}<w:sectPr/></w:body></w:document>`);
  return EditorDocument.open(new Uint8Array(await pkg.save()));
}

const run = (s: string): string => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
const start = (id: number): string => `<w:bookmarkStart w:id="${id}" w:name="_Toc${id}"/>`;
const end = (id: number): string => `<w:bookmarkEnd w:id="${id}"/>`;
const p = (inner: string, style?: string): string => `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}${inner}</w:p>`;

/** The guide's shape: a paragraph, then a heading whose two TOC bookmarks open before its text and close after it, then a paragraph. */
const GUIDE = p(run('One source change.')) + p(start(1) + start(2) + run('OpenXML concepts') + end(1) + end(2), 'Heading1') + p(run('After the heading.'));

/** A block's children: text as it stands, a bookmark marker as `[1` and `1]`, another atom by kind. */
function shape(node: PMNode, session: EditorDocument): string[] {
  const out: string[] = [];
  node.forEach((child) => {
    if (child.isText) { out.push(child.text!); return; }
    const kind = child.attrs['kind'] as string;
    const id = (session.ctx.ids.get(String(child.attrs['sourceId']))?.value as { id?: unknown } | undefined)?.id;
    out.push(kind === 'bookmarkStart' ? `[${String(id)}` : kind === 'bookmarkEnd' ? `${String(id)}]` : kind);
  });
  return out;
}

/** The markers of the saved main part, in order: `[1` for a start, `1]` for an end, `|` for a paragraph's end. */
function markersOf(xml: string): string {
  return [...xml.matchAll(/<w:bookmarkStart [^>]*w:id="(\d+)"[^>]*\/>|<w:bookmarkEnd [^>]*w:id="(\d+)"[^>]*\/>|<\/w:p>/g)]
    .map((m) => (m[1] !== undefined ? `[${m[1]}` : m[2] !== undefined ? `${m[2]}]` : '|')).join(' ');
}

/** Every start has its end and every end its start. */
function paired(xml: string): boolean {
  const starts = [...xml.matchAll(/<w:bookmarkStart [^>]*w:id="(\d+)"/g)].map((m) => m[1]).sort();
  const ends = [...xml.matchAll(/<w:bookmarkEnd [^>]*w:id="(\d+)"/g)].map((m) => m[1]).sort();
  return JSON.stringify(starts) === JSON.stringify(ends);
}

/** The position of block `i`'s first child, and of its first child after its leading atoms (where the caret is drawn). */
function blockStart(doc: PMNode, i: number): number {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += doc.child(k).nodeSize;
  return pos + 1;
}
function visibleStart(doc: PMNode, i: number): number {
  let pos = blockStart(doc, i);
  const block = doc.child(i);
  for (let k = 0; k < block.childCount && block.child(k).type.name === 'opaqueInline'; k++) pos += block.child(k).nodeSize;
  return pos;
}

/** A local transaction recorded as the view records it and written to the store, one undo step. */
function edit(session: EditorDocument, make: (state: EditorState) => Transaction): Transaction {
  const state = stateFor(session);
  const { tr } = recordTransaction(make(state), state);
  session.undoManager.stopCapturing();
  session.replace(state.apply(tr).doc, LOCAL_ORIGIN, [], undefined, state.doc);
  return tr;
}

/** A selection deleted as Delete deletes it: ProseMirror's `deleteSelection`, which widens to the blocks it covers. */
const deleteSelection = (session: EditorDocument, from: number, to: number): Transaction =>
  edit(session, (state) => state.tr.setSelection(TextSelection.create(state.doc, from, to)).deleteSelection());

const saved = async (session: EditorDocument): Promise<string> => mainXml((await session.save()).bytes);

describe('a bookmark keeps both its markers (ED-005 section 12.115)', () => {
  test('the guide: the paragraph before a heading deleted with the heading\'s leading starts; the starts stay at the heading\'s start', async () => {
    const session = await documentOf(GUIDE);
    const doc = session.doc;
    deleteSelection(session, blockStart(doc, 0), visibleStart(doc, 1));
    expect(session.doc.childCount).toBe(doc.childCount - 1);
    expect(shape(session.doc.child(0), session)).toEqual(['[1', '[2', 'OpenXML concepts', '1]', '2]']);
    const xml = await saved(session);
    expect(paired(xml)).toBe(true);
    expect(markersOf(xml)).toBe('[1 [2 1] 2] | |');
    expect(xml).toMatch(/<w:pStyle w:val="Heading1"\/><\/w:pPr><w:bookmarkStart/);
    expect(xml).not.toContain('One source change');
  });

  test('before the fix\'s rule, the same deletion is what lost them: ProseMirror takes the heading\'s content up to the selection\'s end', async () => {
    const session = await documentOf(GUIDE);
    const state = stateFor(session);
    const tr = state.tr.setSelection(TextSelection.create(state.doc, blockStart(state.doc, 0), visibleStart(state.doc, 1))).deleteSelection();
    expect(shape(tr.doc.child(0), session)).toEqual(['OpenXML concepts', '1]', '2]']);
  });

  test('ends taken with the text before them go back before what follows, and the bookmark keeps the text left', async () => {
    const session = await documentOf(GUIDE);
    const doc = session.doc;
    const into = visibleStart(doc, 1) + 'OpenXML'.length;
    deleteSelection(session, into, visibleStart(doc, 2) + 'After '.length);
    expect(shape(session.doc.child(1), session)).toEqual(['[1', '[2', 'OpenXML', '1]', '2]', 'the heading.']);
    expect(paired(await saved(session))).toBe(true);
  });

  test('replaced by typed text: a start goes after what was typed, an end before it, so the bookmark loses what was deleted alone', async () => {
    const session = await documentOf(GUIDE);
    const doc = session.doc;
    edit(session, (state) => state.tr.insertText('New ', blockStart(doc, 1), visibleStart(doc, 1) + 'Open'.length));
    expect(shape(session.doc.child(1), session)).toEqual(['New ', '[1', '[2', 'XML concepts', '1]', '2]']);
    edit(session, (state) => {
      const heading = blockStart(state.doc, 1);
      return state.tr.insertText('!', heading + 'New '.length + 2 + 'XML con'.length, heading + state.doc.child(1).content.size);
    });
    expect(shape(session.doc.child(1), session)).toEqual(['New ', '[1', '[2', 'XML con', '1]', '2]', '!']);
    expect(paired(await saved(session))).toBe(true);
  });

  test('a bookmark deleted with both its markers goes, and so does its name', async () => {
    const session = await documentOf(GUIDE);
    const doc = session.doc;
    deleteSelection(session, blockStart(doc, 0), blockStart(doc, 2));
    const xml = await saved(session);
    expect(xml).not.toContain('bookmarkStart');
    expect(xml).not.toContain('bookmarkEnd');
    expect(xml).not.toContain('_Toc1');
  });

  test('one bookmark whole and one by its start: the whole one goes, the other\'s start stays', async () => {
    const session = await documentOf(p(start(1) + start(2) + run('Gone') + end(1)) + p(run('Kept') + end(2)));
    const doc = session.doc;
    deleteSelection(session, blockStart(doc, 0), blockStart(doc, 0) + doc.child(0).content.size);
    expect(shape(session.doc.child(0), session)).toEqual(['[2']);
    const xml = await saved(session);
    expect(markersOf(xml)).toBe('[2 | 2] |');
  });

  test('with tracking on the deletion keeps them, marked, and accepting it leaves them plain at the heading\'s start', async () => {
    const session = await documentOf(GUIDE);
    await runCommandAsync(session, 'document.setChangeTrackingMode', { value: 'TrackAll' });
    const doc = session.doc;
    deleteSelection(session, blockStart(doc, 0), visibleStart(doc, 1));
    expect(paired(await saved(session))).toBe(true);
    await runCommandAsync(session, 'trackedChanges.acceptAll', undefined);
    const heading = session.doc.child(0);
    expect(shape(heading, session)).toEqual(['[1', '[2', 'OpenXML concepts', '1]', '2]']);
    heading.forEach((child) => { if (!child.isText) expect(child.attrs['revision']).toBeNull(); });
    const xml = await saved(session);
    expect(paired(xml)).toBe(true);
    expect(xml).not.toContain('One source change');
    expect(xml).not.toMatch(/<w:(ins|del) /);
  });

  test('rejecting an insertion that held a start puts it back; the insertion\'s text goes', async () => {
    const session = await documentOf(p(`<w:ins w:id="90" w:author="A" w:date="2026-11-08T00:00:00Z">${run('Inserted ')}${start(1)}${run('also')}</w:ins>${run(' kept')}`) + p(run('After') + end(1)));
    const before = shape(session.doc.child(0), session);
    expect(before).toContain('[1');
    await runCommandAsync(session, 'trackedChanges.rejectAll', undefined);
    expect(shape(session.doc.child(0), session)).toEqual(['[1', ' kept']);
    expect(paired(await saved(session))).toBe(true);
  });

  test('undo takes the deletion and the markers put back as one step', async () => {
    const session = await documentOf(GUIDE);
    const doc = session.doc;
    deleteSelection(session, blockStart(doc, 0), visibleStart(doc, 1));
    session.undoManager.undo();
    expect(session.doc.eq(doc)).toBe(true);
  });

  test('a marker a transaction moves is in the document once', async () => {
    const session = await documentOf(GUIDE);
    const doc = session.doc;
    const marker = doc.child(1).child(0);
    const tr = edit(session, (state) => state.tr.delete(blockStart(doc, 1), blockStart(doc, 1) + 1).insert(blockStart(doc, 0), marker));
    expect(tr.steps).toHaveLength(2);
    expect(shape(session.doc.child(0), session)).toEqual(['[1', 'One source change.']);
    expect(shape(session.doc.child(1), session)).toEqual(['[2', 'OpenXML concepts', '1]', '2]']);
    expect(paired(await saved(session))).toBe(true);
  });

  test('a transaction that removes no marker is left as it was', async () => {
    const session = await documentOf(GUIDE);
    const state = stateFor(session);
    const typed = state.tr.insertText('x', visibleStart(state.doc, 1)).delete(blockStart(state.doc, 0), blockStart(state.doc, 0) + 3);
    const steps = typed.steps.length;
    expect(recordTransaction(typed, state).tr.steps).toHaveLength(steps);
  });

  test('a paragraph deleted whole: its start goes into the paragraph after it, its end into the one before', async () => {
    const starts = await documentOf(p(start(2) + run('Gone')) + p(run('Kept') + end(2)));
    edit(starts, (state) => state.tr.delete(0, state.doc.child(0).nodeSize));
    expect(shape(starts.doc.child(0), starts)).toEqual(['[2', 'Kept', '2]']);
    const ends = await documentOf(p(start(3) + run('Kept')) + p(run('Gone') + end(3)) + p(run('After')));
    edit(ends, (state) => state.tr.delete(state.doc.child(0).nodeSize, state.doc.child(0).nodeSize + state.doc.child(1).nodeSize));
    expect(shape(ends.doc.child(0), ends)).toEqual(['[3', 'Kept', '3]']);
    expect(shape(ends.doc.child(1), ends)).toEqual(['After']);
    expect(paired(await saved(ends))).toBe(true);
  });

  test('a marker moved and then removed again goes back where it was removed last', async () => {
    const session = await documentOf(GUIDE);
    const doc = session.doc;
    const marker = doc.child(1).child(0);
    edit(session, (state) => {
      const tr = state.tr.delete(blockStart(doc, 1), blockStart(doc, 1) + 1).insert(blockStart(doc, 0) + 3, marker);
      return tr.delete(blockStart(doc, 0) + 3, blockStart(doc, 0) + 4);
    });
    expect(shape(session.doc.child(0), session)).toEqual(['One', '[1', ' source change.']);
    expect(paired(await saved(session))).toBe(true);
  });

  test('a start between blocks deleted with the paragraph before it goes back between the blocks', async () => {
    const session = await documentOf(p(run('One source change.')) + start(5) + p(run('OpenXML concepts') + end(5), 'Heading1'));
    const doc = session.doc;
    expect(doc.child(1).type.name).toBe('opaqueBlock');
    edit(session, (state) => state.tr.delete(0, doc.child(0).nodeSize + doc.child(1).nodeSize));
    expect(session.doc.child(0).type.name).toBe('opaqueBlock');
    const xml = await saved(session);
    expect(paired(xml)).toBe(true);
    expect(xml).toMatch(/<w:body><w:bookmarkStart w:id="5"/);
  });

  test('a peer\'s transaction is left as it came: its own store already holds what it did', async () => {
    const session = await documentOf(GUIDE);
    const state = stateFor(session);
    const tr = state.tr.setSelection(TextSelection.create(state.doc, blockStart(state.doc, 0), visibleStart(state.doc, 1))).deleteSelection().setMeta(ySyncPluginKey, { isChangeOrigin: true });
    const steps = tr.steps.length;
    expect(recordTransaction(tr, state).tr.steps).toHaveLength(steps);
  });

  test('markers removed by separate steps come back in their document order, whatever ran first and moved them', async () => {
    const session = await documentOf(p(run('A') + start(2) + start(1) + run('B')) + p(run('C') + end(1) + end(2)));
    const doc = session.doc;
    const b = blockStart(doc, 0);
    edit(session, (state) => {
      const before = state.schema.nodes['paragraph']!.create(null, state.schema.text('Inserted before'));
      const tr = state.tr.delete(b + 2, b + 3).insert(0, before);
      const shift = before.nodeSize;
      return tr.delete(b + 1 + shift, b + 2 + shift);
    });
    expect(shape(session.doc.child(1), session)).toEqual(['A', '[2', '[1', 'B']);
    expect(paired(await saved(session))).toBe(true);
  });

  test('markers put back at one place keep their order, ends before starts', async () => {
    const session = await documentOf(p(run('A') + start(1) + run('B') + end(1) + start(2) + start(3) + run('C')) + p(run('D') + end(2) + end(3)));
    const doc = session.doc;
    const b = blockStart(doc, 0);
    // From inside "B" to inside "C": end 1 and starts 2 and 3 go, their partners stay.
    edit(session, (state) => state.tr.replaceWith(b + 2, b + 6, Fragment.empty));
    expect(shape(session.doc.child(0), session)).toEqual(['A', '[1', '1]', '[2', '[3', 'C']);
    expect(paired(await saved(session))).toBe(true);
  });
});
