// Editor check 39 (run by Jason in Word 365, 2026-10-09; `fixtures/tables/check39/`, the probe built by
// `fixtures/build/table-autofit-probe.mjs`): what Insert Right writes in each of Word's table layout modes
// (ED-005 section 12.78 item 4, the answer in section 12.79). The editor's `table.addColumns` is held to Word's save:
//   fixed: the new column copies the width beside it, the table grows by it (past the margin);
//   AutoFit to window: the table keeps its 100%, the three cells an equal share (1666/1667/1667 pct, the grid a third each);
//   AutoFit to contents: the cells and the table stay `auto`; the grid Word writes is its measurement of the text
//   (593/222/597 twips), which the editor cannot make in Node, so only its shape is held.
import { describe, expect, test } from 'vitest';
import { TableMap } from 'prosemirror-tables';
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

function cellStart(session: EditorDocument, block: number): number {
  let pos = 0;
  for (let k = 0; k < block; k++) pos += session.doc.child(k).nodeSize;
  const map = TableMap.get(session.doc.child(block));
  return pos + 1 + map.map[0]! + 2;
}

describe('check 39: a column inserted in each layout mode, held to Word 365', () => {
  test('fixed, AutoFit to window, AutoFit to contents', async () => {
    const word = tables(await reloadBody(await fixture('tables/check39/table-autofit-probe-word365.docx')));
    const session = await EditorDocument.open(await fixture('tables/check39/table-autofit-probe.docx'));
    const blocks: number[] = [];
    session.doc.forEach((b, _o, i) => { if (b.type.name === 'table') blocks.push(i); });
    expect(blocks).toHaveLength(3);
    for (const b of blocks) expect(runCommand(session, 'table.addColumns', { location: 'After' }, { from: cellStart(session, b) }).ok).toBe(true);
    const saved = await session.save();
    const mine = tables(await reloadBody(saved.bytes));
    expect(mine).toHaveLength(3);
    // Fixed: Word 365 copies the 4513 beside the caret and grows the table to 13539 dxa.
    expect(tblW(word[0]!)).toEqual({ w: 13539, type: 'dxa' });
    expect(grid(word[0]!)).toEqual([4513, 4513, 4513]);
    expect(tblW(mine[0]!)).toEqual(tblW(word[0]!));
    expect(grid(mine[0]!)).toEqual(grid(word[0]!));
    expect(tcWs(mine[0]!)).toEqual(tcWs(word[0]!));
    // Window: 100% kept, an equal share each (the probe's two columns were equal, so a proportional rescale gives the same).
    expect(tblW(word[1]!)).toEqual({ w: 5000, type: 'pct' });
    expect(tcWs(word[1]!)).toEqual([{ w: 1666, type: 'pct' }, { w: 1667, type: 'pct' }, { w: 1667, type: 'pct' }]);
    expect(tblW(mine[1]!)).toEqual(tblW(word[1]!));
    expect(tcWs(mine[1]!)).toEqual(tcWs(word[1]!));
    expect(grid(mine[1]!)).toHaveLength(3);
    for (let i = 0; i < 3; i++) expect(Math.abs(grid(mine[1]!)[i]! - grid(word[1]!)[i]!)).toBeLessThanOrEqual(10);
    // Contents: the table and the cells automatic; Word's grid is its text measurement (593/222/597), not held.
    expect(tblW(word[2]!)).toEqual({ w: 0, type: 'auto' });
    expect(tcWs(word[2]!)).toEqual([{ w: 0, type: 'auto' }, { w: 0, type: 'auto' }, { w: 0, type: 'auto' }]);
    expect(grid(word[2]!)).toEqual([593, 222, 597]);
    expect(tblW(mine[2]!)).toEqual(tblW(word[2]!));
    expect(tcWs(mine[2]!)).toEqual(tcWs(word[2]!));
    expect(grid(mine[2]!)).toHaveLength(3);
    console.log('editor:', mine.map((t) => ({ tblW: tblW(t), grid: grid(t), tcW: tcWs(t) })));
  });
});
