// Word check 40 (Jason, Word 365, 2026-10-09; core-ts CR-002 section 42.6, the files copied from its
// `test/fixtures/check40/` into `fixtures/tables/check40/`): the three points check 39 left open, held on the
// editor's `table.addColumns` (ED-005 section 12.82).
//   40a1, 40a2: an AutoFit-to-window table of 3,000 and 6,026 twips; Insert Right in the narrow first cell, and Insert
//     Left in the wide last cell: both copy the column to the RIGHT of the boundary (6,026) and rescale proportionally,
//     the cells 996, 2002, 2002 pct; Word's grid 1797, 3610, 3610 (its layout's measurement, 9 twips short).
//   40b1: a fixed table, row 1 a cell spanning columns 1 and 2; Insert Right in row 2's first cell: the boundary lies
//     inside the span; Word keeps the span and puts the new cell before it; the width copied is column 2's (3,009).
//   40b2: Insert Right in the spanning cell: the boundary at its right edge; column 3's width (3,008) is copied.
import { describe, expect, test } from 'vitest';
import { TableMap } from 'prosemirror-tables';
import { WordprocessingMLPackage, rowsOf, cellsOf, typeNameOf, textOf } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { EditorDocument, runCommand } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

async function reloadBody(bytes: Uint8Array): Promise<wml.Body> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  return (await pkg.getMainDocumentPart().readContents()).body!;
}
const firstTable = (body: wml.Body): wml.Tbl => body.content!.find((el) => typeNameOf(el) === 'org_docx4j_wml.Tbl')!.value as wml.Tbl;
const grid = (t: wml.Tbl): number[] => t.tblGrid!.gridCol!.map((g) => Number(g.w));
const tblW = (t: wml.Tbl): { w: number; type: string } => ({ w: Number(t.tblPr!.tblW!.w), type: String(t.tblPr!.tblW!.type) });
const rows = (t: wml.Tbl): { w: number; type: string; span: number; text: string }[][] => rowsOf(t).map((r) => cellsOf(r.element.value).map((c) => {
  const tcPr = c.element.value.tcPr!;
  return { w: Number(tcPr.tcW!.w), type: String(tcPr.tcW!.type), span: tcPr.gridSpan?.val ?? 1, text: textOf(c.element).trim() };
}));

/** The first text position inside the cell holding `text`. */
function cellWith(session: EditorDocument, text: string): number {
  let found = -1;
  session.doc.descendants((n, pos) => {
    if (found >= 0) return false;
    if (n.type.name === 'table_cell' && n.textContent.startsWith(text)) { found = pos + 2; return false; }
    return true;
  });
  expect(found).toBeGreaterThan(0);
  return found;
}

async function run(name: string, cell: string, location: 'Before' | 'After'): Promise<{ mine: wml.Tbl; word: wml.Tbl }> {
  const session = await EditorDocument.open(await fixture(`tables/check40/${name}`));
  expect(runCommand(session, 'table.addColumns', { location }, { from: cellWith(session, cell) }).ok).toBe(true);
  expect(TableMap.get(session.doc.child(session.doc.childCount - 2).type.name === 'table' ? session.doc.child(session.doc.childCount - 2) : session.doc.child(1)).width).toBeGreaterThan(0);
  const saved = await session.save();
  const word = firstTable(await reloadBody(await fixture(`tables/check40/${name.replace(/-.*$/, '')}-word365.docx`)));
  return { mine: firstTable(await reloadBody(saved.bytes)), word };
}

describe('check 40: the window rescale with unequal columns, a column beside and inside a span', () => {
  test('40a1: Insert Right in the narrow cell copies the wide column to its right and rescales proportionally', async () => {
    const { mine, word } = await run('40a1-window-unequal.docx', 'W1A', 'After');
    expect(tblW(word)).toEqual({ w: 5000, type: 'pct' });
    expect(rows(word)[0]!.map((c) => c.w)).toEqual([996, 2002, 2002]);
    expect(tblW(mine)).toEqual(tblW(word));
    expect(rows(mine)).toEqual(rows(word));
    for (let i = 0; i < 3; i++) expect(Math.abs(grid(mine)[i]! - grid(word)[i]!)).toBeLessThanOrEqual(10);
  });
  test('40a2: Insert Left in the wide cell copies it, the column to the boundary\'s right', async () => {
    const { mine, word } = await run('40a2-window-unequal.docx', 'W1B', 'Before');
    expect(rows(mine)).toEqual(rows(word));
    for (let i = 0; i < 3; i++) expect(Math.abs(grid(mine)[i]! - grid(word)[i]!)).toBeLessThanOrEqual(10);
  });
  test('40b1: a boundary inside a spanning cell: the span kept, the new cell before it; column 2\'s width copied', async () => {
    const { mine, word } = await run('40b1-span.docx', 'S2A', 'After');
    expect(tblW(word)).toEqual({ w: 12035, type: 'dxa' });
    expect(grid(word)).toEqual([3009, 3009, 3009, 3008]);
    expect(rows(word).map((r) => r.map((c) => [c.span, c.text.slice(0, 4)]))).toEqual([[[1, ''], [2, 'S1AB'], [1, 'S1C']], [[1, 'S2A'], [1, ''], [1, 'S2B'], [1, 'S2C']]]);
    expect(tblW(mine)).toEqual(tblW(word));
    expect(grid(mine)).toEqual(grid(word));
    expect(rows(mine)).toEqual(rows(word));
  });
  test('40b2: Insert Right in the spanning cell: a plain cell after it, column 3\'s width copied', async () => {
    const { mine, word } = await run('40b2-span.docx', 'S1AB', 'After');
    expect(tblW(word)).toEqual({ w: 12034, type: 'dxa' });
    expect(grid(word)).toEqual([3009, 3009, 3008, 3008]);
    expect(tblW(mine)).toEqual(tblW(word));
    expect(grid(mine)).toEqual(grid(word));
    expect(rows(mine)).toEqual(rows(word));
  });
});
