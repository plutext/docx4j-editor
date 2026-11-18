// Editor check 42 (Jason, Word 365, 2026-10-09; `fixtures/tables/check42/`, the probe by `fixtures/build/insert-table-probe.mjs`):
// what Word's Insert Table dialog writes for each of its choices, and the editor's `table.insert` held to it
// (ED-005 section 12.83 item 5):
//   A. fixed column width Auto (5 x 2): an automatic table width, NO layout element, the cells in twips sharing the text
//      width (Word's 1803/1803/1803/1804/1804 over 9,017; the editor shares 9,026);
//   B. fixed 72 pt (3 x 1): the same form, every cell and grid column 1,440;
//   C. AutoFit to contents (3 x 1): an automatic width, automatic cells, the grid Word's measurement (222 each);
//   D. AutoFit to window (3 x 1): a 100% width, the cells 1666/1667/1667 pct, the grid 3005/3006/3006.
import { describe, expect, test } from 'vitest';
import { WordprocessingMLPackage, rowsOf, cellsOf, typeNameOf } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { EditorDocument, runCommand } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

async function reloadBody(bytes: Uint8Array): Promise<wml.Body> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  return (await pkg.getMainDocumentPart().readContents()).body!;
}
const tables = (body: wml.Body): wml.Tbl[] => body.content!.filter((el) => typeNameOf(el) === 'org_docx4j_wml.Tbl').map((el) => el.value as wml.Tbl);
const grid = (t: wml.Tbl): number[] => t.tblGrid!.gridCol!.map((g) => Number(g.w));
const tcWs = (t: wml.Tbl): { w: number; type: string }[] => cellsOf(rowsOf(t)[0]!.element.value).map((c) => { const w = c.element.value.tcPr!.tcW!; return { w: Number(w.w), type: String(w.type) }; });
const tblW = (t: wml.Tbl): { w: number; type: string } => ({ w: Number(t.tblPr!.tblW!.w), type: String(t.tblPr!.tblW!.type) });

describe('check 42: the Insert Table dialog, held to Word 365', () => {
  test('fixed Auto, fixed 72 pt, AutoFit to contents, AutoFit to window', async () => {
    const word = tables(await reloadBody(await fixture('tables/check42/insert-table-word365.docx')));
    expect(word).toHaveLength(4);
    const session = await EditorDocument.open(await fixture('tables/check42/insert-table-probe.docx'));
    // The empty paragraph under each heading: a table replaces it.
    const empties: number[] = [];
    let pos = 0;
    session.doc.forEach((b) => { if (b.type.name === 'paragraph' && b.content.size === 0) empties.push(pos + 1); pos += b.nodeSize; });
    expect(empties).toHaveLength(4);
    const choices = [
      { rows: 2, columns: 5 },
      { rows: 1, columns: 3, autoFit: 'fixed', columnWidth: 72 },
      { rows: 1, columns: 3, autoFit: 'contents' },
      { rows: 1, columns: 3, autoFit: 'window' },
    ] as const;
    // From the last up, so each position stays as computed.
    for (let i = 3; i >= 0; i--) expect(runCommand(session, 'table.insert', choices[i], { from: empties[i]! }).ok).toBe(true);
    const mine = tables(await reloadBody((await session.save()).bytes));
    expect(mine).toHaveLength(4);
    // A: automatic width, no layout, twips cells sharing the text width.
    expect(tblW(word[0]!)).toEqual({ w: 0, type: 'auto' });
    expect(word[0]!.tblPr!.tblLayout).toBeUndefined();
    expect(tcWs(word[0]!).map((c) => c.type)).toEqual(['dxa', 'dxa', 'dxa', 'dxa', 'dxa']);
    expect(tblW(mine[0]!)).toEqual(tblW(word[0]!));
    expect(mine[0]!.tblPr!.tblLayout).toBeUndefined();
    expect(tcWs(mine[0]!).map((c) => c.type)).toEqual(tcWs(word[0]!).map((c) => c.type));
    for (let i = 0; i < 5; i++) { expect(Math.abs(tcWs(mine[0]!)[i]!.w - tcWs(word[0]!)[i]!.w)).toBeLessThanOrEqual(10); expect(Math.abs(grid(mine[0]!)[i]! - grid(word[0]!)[i]!)).toBeLessThanOrEqual(10); }
    // B: the same form, 1,440 everywhere.
    expect(tblW(word[1]!)).toEqual({ w: 0, type: 'auto' });
    expect(grid(word[1]!)).toEqual([1440, 1440, 1440]);
    expect(tcWs(word[1]!)).toEqual([{ w: 1440, type: 'dxa' }, { w: 1440, type: 'dxa' }, { w: 1440, type: 'dxa' }]);
    expect(tblW(mine[1]!)).toEqual(tblW(word[1]!));
    expect(mine[1]!.tblPr!.tblLayout).toBeUndefined();
    expect(grid(mine[1]!)).toEqual(grid(word[1]!));
    expect(tcWs(mine[1]!)).toEqual(tcWs(word[1]!));
    // C: automatic throughout; Word's grid is its measurement of empty cells (222), the editor's a share, not held.
    expect(tblW(word[2]!)).toEqual({ w: 0, type: 'auto' });
    expect(tcWs(word[2]!)).toEqual([{ w: 0, type: 'auto' }, { w: 0, type: 'auto' }, { w: 0, type: 'auto' }]);
    expect(grid(word[2]!)).toEqual([222, 222, 222]);
    expect(tblW(mine[2]!)).toEqual(tblW(word[2]!));
    expect(tcWs(mine[2]!)).toEqual(tcWs(word[2]!));
    expect(grid(mine[2]!)).toHaveLength(3);
    // D: 100%, the cells as shares.
    expect(tblW(word[3]!)).toEqual({ w: 5000, type: 'pct' });
    expect(tcWs(word[3]!)).toEqual([{ w: 1666, type: 'pct' }, { w: 1667, type: 'pct' }, { w: 1667, type: 'pct' }]);
    expect(tblW(mine[3]!)).toEqual(tblW(word[3]!));
    expect(tcWs(mine[3]!)).toEqual(tcWs(word[3]!));
    for (let i = 0; i < 3; i++) expect(Math.abs(grid(mine[3]!)[i]! - grid(word[3]!)[i]!)).toBeLessThanOrEqual(10);
  });
});
