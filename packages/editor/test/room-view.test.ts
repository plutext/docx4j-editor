// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-004 section 3.5: the view's fixing plugins repair on local changes only. A table broken by
// a peer arrives unrepaired, the first binding of a broken store repairs nothing, an undo back to
// a broken table repairs nothing (and keeps the redo), and a local change to the table repairs it.
// ED-005 step C5: the tracker too acts on local changes only, so a peer's typing stays the peer's.
// ED-005 section 12.47 (UPSTREAM-16): two views resolving the same move at once leave what one leaves.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import * as Y from 'yjs';
import { Fragment, type Node as PMNode } from 'prosemirror-model';
import { TextSelection } from 'prosemirror-state';
import { EditorDocument, newEditStore, runCommand, stateFor, trackedChangesOf } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

/** Two stores, each one's updates applied to the other at once: what a provider does. */
function link(a: Y.Doc, b: Y.Doc): void {
  const tag = {};
  a.on('update', (u: Uint8Array, origin: unknown) => { if (origin !== tag) Y.applyUpdate(b, u, tag); });
  b.on('update', (u: Uint8Array, origin: unknown) => { if (origin !== tag) Y.applyUpdate(a, u, tag); });
}

async function room(): Promise<{ creator: EditorDocument; joiner: EditorDocument; base: Uint8Array }> {
  const base = new Uint8Array(await readFile(join(fixtures, 'tables.docx')));
  const creator = await EditorDocument.open(base);
  await creator.share(base, 'tables.docx');
  const store = newEditStore();
  Y.applyUpdate(store.ydoc, Y.encodeStateAsUpdate(creator.store.ydoc));
  link(creator.store.ydoc, store.ydoc);
  const joiner = await EditorDocument.open(base, { store, mode: 'join' });
  return { creator, joiner, base };
}

function mount(session: EditorDocument): DocxEditorView {
  const container = document.createElement('div');
  document.body.appendChild(container);
  return new DocxEditorView(container, session);
}

/** The first table's cell count per row. */
function rowWidths(doc: PMNode): number[] {
  let widths: number[] = [];
  doc.descendants((n) => {
    if (widths.length > 0) return false;
    if (n.type.name !== 'table') return true;
    n.forEach((row) => { widths.push(row.childCount); });
    return false;
  });
  if (widths.length === 0) widths = [];
  return widths;
}

/** The document with the first table's first row missing its last cell: a table `fixTables` would repair. */
function broken(doc: PMNode): PMNode {
  const blocks: PMNode[] = [];
  let done = false;
  doc.forEach((b) => {
    if (done || b.type.name !== 'table') { blocks.push(b); return; }
    done = true;
    const rows: PMNode[] = [];
    b.forEach((row, _o, i) => {
      if (i !== 0) { rows.push(row); return; }
      const cells: PMNode[] = [];
      row.forEach((cell, _c, k) => { if (k < row.childCount - 1) cells.push(cell); });
      rows.push(row.copy(Fragment.from(cells)));
    });
    blocks.push(b.copy(Fragment.from(rows)));
  });
  return doc.copy(Fragment.from(blocks));
}

/** The position just inside the first table's first cell's paragraph. */
function inFirstCell(doc: PMNode): number {
  let pos = -1;
  doc.descendants((n, p) => {
    if (pos >= 0) return false;
    if (n.type.name === 'paragraph' && p > 0 && doc.resolve(p).depth >= 3) { pos = p + 1; return false; }
    return true;
  });
  return pos;
}

describe('fixing plugins run on local changes only (ED-004 section 3.5)', () => {
  test('a table a peer broke arrives unrepaired: nothing is appended and nothing written back', async () => {
    const { creator, joiner } = await room();
    const view = mount(creator);
    const widths = rowWidths(view.state.doc);
    expect(widths[0]).toBeGreaterThan(1);
    joiner.replace(broken(joiner.doc));
    const vector = Y.encodeStateVector(creator.store.ydoc);
    expect(rowWidths(view.state.doc)[0]).toBe(widths[0]! - 1);
    // A repair would have been a transaction of the creator's, written to the store and sent back.
    expect(Y.encodeStateVector(creator.store.ydoc)).toEqual(vector);
    expect(rowWidths(joiner.doc)[0]).toBe(widths[0]! - 1);
    view.destroy();
  });

  test('the first binding of a store holding a broken table repairs nothing', async () => {
    const base = new Uint8Array(await readFile(join(fixtures, 'tables.docx')));
    const session = await EditorDocument.open(base);
    const widths = rowWidths(session.doc);
    session.replace(broken(session.doc));
    const vector = Y.encodeStateVector(session.store.ydoc);
    const view = mount(session);
    expect(rowWidths(view.state.doc)[0]).toBe(widths[0]! - 1);
    expect(Y.encodeStateVector(session.store.ydoc)).toEqual(vector);
    view.destroy();
  });

  test('a local change to the table repairs it; its undo goes back to the broken table without repairing it, and keeps the redo', async () => {
    const { creator, joiner } = await room();
    const view = mount(creator);
    const widths = rowWidths(view.state.doc);
    joiner.replace(broken(joiner.doc));
    expect(rowWidths(view.state.doc)[0]).toBe(widths[0]! - 1);
    // Typing in the table: the fix-up runs on this local change and the row is whole again, on both peers.
    const at = inFirstCell(view.state.doc);
    view.view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at)).insertText('x'));
    expect(rowWidths(view.state.doc)[0]).toBe(widths[0]);
    expect(rowWidths(joiner.doc)[0]).toBe(widths[0]);
    // Undo: the state before the keystroke, the peer's broken table, is restored and left alone.
    expect(view.run('history.undo')).toBe(true);
    expect(rowWidths(view.state.doc)[0]).toBe(widths[0]! - 1);
    expect(rowWidths(joiner.doc)[0]).toBe(widths[0]! - 1);
    expect(creator.undoManager.canRedo()).toBe(true);
    view.destroy();
  });
});

describe("a cell's last block stays a paragraph (ED-005 section 12.109)", () => {
  /** The first table's first cell: its position and the names of its children. */
  function firstCell(doc: PMNode): { pos: number; shape: string[] } {
    let pos = -1;
    doc.descendants((n, p) => { if (pos >= 0) return false; if (n.type.name === 'table_cell') { pos = p; return false; } return true; });
    const cell = doc.nodeAt(pos)!;
    const shape: string[] = [];
    cell.forEach((n) => shape.push(n.type.name));
    return { pos, shape };
  }
  /** A table node for a cell's end: the document's second table, or its first. */
  function someTable(doc: PMNode): PMNode {
    const tables: PMNode[] = [];
    doc.forEach((b) => { if (b.type.name === 'table') tables.push(b); });
    return tables[1] ?? tables[0]!;
  }

  test('a table landed at a cell\'s end by a local transaction gets a paragraph after it; a peer\'s arrives as it is', async () => {
    const { creator, joiner } = await room();
    const view = mount(creator);
    const { pos, shape } = firstCell(view.state.doc);
    expect(shape).toEqual(['paragraph']);
    const cell = view.state.doc.nodeAt(pos)!;
    view.view.dispatch(view.state.tr.insert(pos + cell.nodeSize - 1, someTable(view.state.doc)));
    expect(firstCell(view.state.doc).shape).toEqual(['paragraph', 'table', 'paragraph']);
    expect(firstCell(joiner.doc).shape).toEqual(['paragraph', 'table', 'paragraph']);
    view.destroy();
    // The other way: the joiner leaves a table last in the cell; the creator's view appends nothing (a save does).
    const { creator: c2, joiner: j2 } = await room();
    const v2 = mount(c2);
    const first = firstCell(j2.doc);
    const jc = j2.doc.nodeAt(first.pos)!;
    j2.replace(stateFor(j2).tr.insert(first.pos + jc.nodeSize - 1, someTable(j2.doc)).doc);
    const vector = Y.encodeStateVector(c2.store.ydoc);
    expect(firstCell(v2.state.doc).shape).toEqual(['paragraph', 'table']);
    expect(firstCell(j2.doc).shape).toEqual(['paragraph', 'table']);
    // A repair would have been a transaction of the creator's, written to the store and sent back.
    expect(Y.encodeStateVector(c2.store.ydoc)).toEqual(vector);
    v2.destroy();
  });
});

describe('a peer\'s tracked edits are not tracked again (ED-005 section 5.4, step C5)', () => {
  test('with tracking on in both, the joiner\'s typing arrives in the creator\'s view as the joiner\'s insertion, and nothing is written back', async () => {
    const { creator, joiner } = await room();
    creator.trackingMode = 'TrackAll';
    creator.trackingAuthor = 'Author A';
    joiner.trackingMode = 'TrackAll';
    joiner.trackingAuthor = 'Author B';
    const view = mount(creator);
    const at = inFirstCell(joiner.doc);
    expect(runCommand(joiner, 'insertText', { text: 'peer' }, { from: at }).ok).toBe(true);
    const vector = Y.encodeStateVector(creator.store.ydoc);
    const authors: string[] = [];
    view.state.doc.descendants((n) => { const m = n.marks.find((k) => k.type.name === 'insertion'); if (n.isText && m && n.text!.includes('peer')) authors.push(m.attrs['author'] as string); });
    expect(authors).toEqual(['Author B']);
    expect(Y.encodeStateVector(creator.store.ydoc)).toEqual(vector);
    // The creator's own typing is the creator's.
    const mine = inFirstCell(view.state.doc);
    view.view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, mine)).insertText('own'));
    const own: string[] = [];
    joiner.doc.descendants((n) => { const m = n.marks.find((k) => k.type.name === 'insertion'); if (n.isText && m && n.text!.includes('own')) own.push(m.attrs['author'] as string); });
    expect(own).toEqual(['Author A']);
    view.destroy();
  });
});

// UPSTREAM-16 (ED-005 section 12.47): y-prosemirror's diff put a copy of a kept paragraph where a move's range marker had
// been, and a copy of a paragraph's text where its inline marker had been, once per view, so two people accepting the
// same move at once had the paragraph twice, and two rejecting it had its words twice over. The view's binding matches
// the store's children to the new ones before its diff (`view/textformat.ts`), as the headless runner does.
describe('two views resolving the same move at once (UPSTREAM-16)', () => {
  for (const action of ['accept', 'reject'] as const) {
    test(`both ${action}: the document one ${action} leaves, on both`, async () => {
      const base = new Uint8Array(await readFile(join(fixtures, 'revisions', 'revisions-word15.docx')));
      const creator = await EditorDocument.open(base);
      await creator.share(base, 'revisions-word15.docx');
      const store = newEditStore();
      Y.applyUpdate(store.ydoc, Y.encodeStateAsUpdate(creator.store.ydoc));
      // Held: each side's updates wait until both have resolved, as two presses within the relay's latency do.
      const held: [Y.Doc, Uint8Array][] = [];
      const tag = {};
      creator.store.ydoc.on('update', (u: Uint8Array, origin: unknown) => { if (origin !== tag) held.push([store.ydoc, u]); });
      store.ydoc.on('update', (u: Uint8Array, origin: unknown) => { if (origin !== tag) held.push([creator.store.ydoc, u]); });
      const joiner = await EditorDocument.open(base, { store, mode: 'join' });
      const views = [mount(creator), mount(joiner)];
      held.length = 0;
      for (const [session, view] of [[creator, views[0]!], [joiner, views[1]!]] as const) {
        const move = trackedChangesOf(session.doc).find((c) => c.move !== undefined)!;
        expect(view.run(`trackedChange.${action}`, { key: session.changeKey(move) })).toBe(true);
      }
      for (const [doc, u] of held.splice(0)) Y.applyUpdate(doc, u, tag);
      const alone = await EditorDocument.open(base);
      const move = trackedChangesOf(alone.doc).find((c) => c.move !== undefined)!;
      expect(runCommand(alone, `trackedChange.${action}`, { key: alone.changeKey(move) }).ok).toBe(true);
      const texts = (doc: PMNode): string[] => { const out: string[] = []; doc.forEach((b) => out.push(b.textContent)); return out; };
      for (const session of [creator, joiner]) expect(texts(session.doc)).toEqual(texts(alone.doc));
      for (const view of views) expect(texts(view.state.doc)).toEqual(texts(alone.doc));
      for (const view of views) view.destroy();
    });
  }
});
