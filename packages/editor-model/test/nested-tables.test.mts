// SPDX-License-Identifier: GPL-3.0-or-later
//
// Nested tables (ED-003 section 11.13, 2026-09-25): a table inside a cell projects, round-trips,
// takes the row and column commands and typing without touching the outer table, is deleted
// alone or with its outer row, and a table can be inserted inside a cell.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, runCommand, compareProjections } from '../src/index.mjs';
import { fixture, comparePackages, textOf } from './helpers.mjs';

/** The position inside the first text block whose text is `text`. */
function positionOf(session: EditorDocument, text: string): number {
  let found = -1;
  session.doc.descendants((n, pos) => { if (found < 0 && n.isTextblock && n.textContent === text) found = pos + 1; return found < 0; });
  if (found < 0) throw new Error(`no block with the text ${text}`);
  return found;
}

const innerTable = (session: EditorDocument) => session.doc.child(1).child(0).child(1).child(0);

async function tablesInMain(bytes: Uint8Array): Promise<number> {
  const xml = await (await WordprocessingMLPackage.load(bytes)).getMainDocumentPart().getXml();
  return (xml.match(/<w:tbl>/g) ?? []).length;
}

describe('nested tables', () => {
  test('projection and the untouched round trip', async () => {
    const original = await fixture('nested-table.docx');
    const session = await EditorDocument.open(original);
    const outer = session.doc.child(1);
    expect(outer.type.name).toBe('table');
    const inner = innerTable(session);
    expect(inner.type.name).toBe('table');
    expect(inner.childCount).toBe(2);
    expect(inner.child(0).childCount).toBe(2);
    expect(textOf(inner.child(0).child(0).child(0))).toBe('inner a1');
    expect(textOf(session.doc.child(1).child(0).child(1).child(1))).toBe('after the inner table');
    expect(session.diagnostics.opaque).toEqual([]);
    const saved = await session.save();
    expect(saved.changed).toBe(false);
    expect(comparePackages(original, saved.bytes)).toEqual([]);
  });

  test('rows and columns added to the inner table, typing in it, a save and a reload; the outer table unchanged', async () => {
    const original = await fixture('nested-table.docx');
    const session = await EditorDocument.open(original);
    const at = positionOf(session, 'inner a1');
    expect(runCommand(session, 'table.addRows', { location: 'After' }, { from: at }).ok).toBe(true);
    expect(innerTable(session).childCount).toBe(3);
    expect(session.doc.child(1).childCount).toBe(2);
    expect(runCommand(session, 'table.addColumns', { location: 'After' }, { from: at }).ok).toBe(true);
    expect(innerTable(session).child(0).childCount).toBe(3);
    expect(session.doc.child(1).child(0).childCount).toBe(2);
    expect(runCommand(session, 'insertText', { text: 'X' }, { from: at }).ok).toBe(true);
    const saved = await session.save();
    expect(saved.selfCheck?.equal).toBe(true);
    expect(comparePackages(original, saved.bytes, [session.main.partName.storeName, 'docProps/app.xml'])).toEqual([]);
    expect(await tablesInMain(saved.bytes)).toBe(2);
    const reopened = await EditorDocument.open(saved.bytes);
    const inner = innerTable(reopened);
    expect(inner.childCount).toBe(3);
    expect(inner.child(0).childCount).toBe(3);
    expect(textOf(inner.child(0).child(0).child(0))).toBe('Xinner a1');
    expect(compareProjections(session.doc, reopened.imported).equal).toBe(true);
  });

  test('the inner table deleted alone keeps the cell\'s paragraph; an outer row deleted takes the nested table with it', async () => {
    const original = await fixture('nested-table.docx');
    const session = await EditorDocument.open(original);
    expect(runCommand(session, 'table.delete', undefined, { from: positionOf(session, 'inner a1') }).ok).toBe(true);
    const cell = session.doc.child(1).child(0).child(1);
    expect(cell.childCount).toBe(1);
    expect(textOf(cell.child(0))).toBe('after the inner table');
    expect(session.doc.child(1).type.name).toBe('table');
    const saved = await session.save();
    expect(saved.selfCheck?.equal).toBe(true);
    expect(await tablesInMain(saved.bytes)).toBe(1);

    const fresh = await EditorDocument.open(original);
    expect(runCommand(fresh, 'table.deleteRows', undefined, { from: positionOf(fresh, 'outer a1') }).ok).toBe(true);
    expect(fresh.doc.child(1).childCount).toBe(1);
    const saved2 = await fresh.save();
    expect(saved2.selfCheck?.equal).toBe(true);
    expect(await tablesInMain(saved2.bytes)).toBe(1);
  });

  test('table.insert inside a cell nests a table after the caret\'s paragraph, an empty paragraph after it (ED-005 section 12.109); in an empty cell the paragraph stays', async () => {
    const original = await fixture('nested-table.docx');
    const session = await EditorDocument.open(original);
    // In "outer b2", a plain cell: the new table follows the paragraph.
    expect(runCommand(session, 'table.insert', { rows: 2, columns: 3 }, { from: positionOf(session, 'outer b2') }).ok).toBe(true);
    const cell = session.doc.child(1).child(1).child(1);
    expect(cell.childCount).toBe(3);   // the paragraph, the table, and the empty paragraph Word requires after it (ED-005 section 12.109)
    expect(textOf(cell.child(0))).toBe('outer b2');
    expect(cell.child(1).type.name).toBe('table');
    expect(cell.child(2).type.name).toBe('paragraph');
    expect(cell.child(1).childCount).toBe(2);
    expect(cell.child(1).child(0).childCount).toBe(3);
    // Inside the inner table's cell too (three deep).
    expect(runCommand(session, 'table.insert', undefined, { from: positionOf(session, 'inner b2') }).ok).toBe(true);
    expect(innerTable(session).child(1).child(1).child(1).type.name).toBe('table');
    // An empty paragraph that is a cell's only block is kept beside the new table.
    const emptyCellPos = (() => { let pos = -1; session.doc.descendants((n, p) => { if (pos < 0 && n.type.name === 'table_cell' && n.childCount === 1 && n.child(0).content.size === 0) pos = p + 2; return pos < 0; }); return pos; })();
    expect(emptyCellPos).toBeGreaterThan(0);
    expect(runCommand(session, 'table.insert', undefined, { from: emptyCellPos }).ok).toBe(true);
    const emptyCell = session.doc.resolve(emptyCellPos).node(-1);
    expect(emptyCell.type.name).toBe('table_cell');
    expect(emptyCell.childCount).toBe(3);
    expect(emptyCell.child(0).type.name).toBe('paragraph');
    expect(emptyCell.child(1).type.name).toBe('table');
    expect(emptyCell.child(2).type.name).toBe('paragraph');
    const saved = await session.save();
    expect(saved.selfCheck?.equal).toBe(true);
    expect(await tablesInMain(saved.bytes)).toBe(5);
    const reopened = await EditorDocument.open(saved.bytes);
    expect(compareProjections(session.doc, reopened.imported).equal).toBe(true);
    // In a body paragraph as before: after the paragraph, at the top level.
    expect(runCommand(session, 'table.insert', undefined, { from: 1 }).ok).toBe(true);
    expect(session.doc.child(1).type.name).toBe('table');
    expect(session.doc.child(2).type.name).toBe('table');
  });
});
