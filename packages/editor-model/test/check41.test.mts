// Word check 41 (Jason, Word 365, 2026-10-09; core-ts CR-002 section 42.6, the files copied from its
// `test/fixtures/check41/` into `fixtures/tables/check41/`): what a new column's cells take from their neighbour
// (ED-005 section 12.82 item 6). Column 1 is yellow, centred vertically, with a thick red left border and centred
// bold text; column 2 blue, bottom-aligned, with a dashed green right border and right-aligned italic text. Insert
// Right from P1A (41a) and Insert Left from P1B (41b) both make a blue cell with no border and no vertical
// alignment, its one paragraph right-aligned with an italic mark: the right-hand neighbour's shading and paragraph
// properties, nothing else.
import { describe, expect, test } from 'vitest';
import { WordprocessingMLPackage, rowsOf, cellsOf, typeNameOf } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { EditorDocument, runCommand } from '../src/index.mjs';
import { toPlain } from '../src/json.mjs';
import { fixture } from './helpers.mjs';

async function reloadBody(bytes: Uint8Array): Promise<wml.Body> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  return (await pkg.getMainDocumentPart().readContents()).body!;
}
const firstTable = (body: wml.Body): wml.Tbl => body.content!.find((el) => typeNameOf(el) === 'org_docx4j_wml.Tbl')!.value as wml.Tbl;
/** The first row's cells: their properties less the width, and their first paragraph's properties, as plain objects. */
function firstRow(t: wml.Tbl): { tcPr: unknown; pPr: unknown }[] {
  return cellsOf(rowsOf(t)[0]!.element.value).map((c) => {
    const tcPr = strip(toPlain(c.element.value.tcPr ?? {})) as Record<string, unknown>;
    delete tcPr['tcW'];
    const p = (c.element.value.content ?? []).find((el) => typeNameOf(el) === 'org_docx4j_wml.P')?.value as wml.P | undefined;
    return { tcPr, pPr: strip(toPlain(p?.pPr ?? {})) };
  });
}
/** The plain object without its type names, at every level. */
function strip(v: unknown): unknown {
  if (Array.isArray(v)) return v.map(strip);
  if (v && typeof v === 'object') return Object.fromEntries(Object.entries(v as Record<string, unknown>).filter(([k]) => k !== 'TYPE_NAME').map(([k, x]) => [k, strip(x)]));
  return v;
}

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

describe('check 41: the new cell takes the right-hand neighbour\'s shading and paragraph properties, nothing else', () => {
  for (const [name, cell, location] of [['41a', 'P1A', 'After'], ['41b', 'P1B', 'Before']] as const) {
    test(`${name}: Insert ${location === 'After' ? 'Right' : 'Left'} from ${cell}`, async () => {
      const session = await EditorDocument.open(await fixture(`tables/check41/${name}-properties.docx`));
      expect(runCommand(session, 'table.addColumns', { location }, { from: cellWith(session, cell) }).ok).toBe(true);
      const mine = firstRow(firstTable(await reloadBody((await session.save()).bytes)));
      const word = firstRow(firstTable(await reloadBody(await fixture(`tables/check41/${name}-word365.docx`))));
      expect(word).toHaveLength(3);
      expect(word[1]!.tcPr).toEqual({ shd: { val: 'clear', color: 'auto', fill: '99CCFF' } });
      expect(word[1]!.pPr).toEqual({ jc: { val: 'right' }, rPr: { i: {} } });
      expect(mine).toHaveLength(3);
      expect(mine[1]).toEqual(word[1]);
      expect(mine[0]).toEqual(word[0]);
      expect(mine[2]).toEqual(word[2]);
    });
  }
});
