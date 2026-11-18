// SPDX-License-Identifier: GPL-3.0-or-later
//
// A table style's conditional formatting for the page (ED-005 section 12.65): the conditions `cellConditionsOf` finds
// for each cell, over a stand-in table built from the table node, are the engine's over the file's own `w:tbl` (core-ts
// CR-007), for every cell of every table of the fixtures that exercise them: core-ts's probes (every look, band sizes,
// `w:cnfStyle` caches, spans, `w:gridBefore`, a vertical merge, a nested table, controls around rows and cells),
// `headers-footers.docx`'s styled table and the 255-page document's.
import { describe, expect, test } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { Element } from '@docx4j/core-ts';
import { EditorDocument, cellConditionsOf, tableConditionsOf, conditionalTcPr } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

const FIXTURES = ['tables/tables-conditions.docx', 'tables/tables-conditions-mode12.docx', 'tables/tables-banding-merged-row.docx', 'tables/tables-banding-band-size.docx', 'headers-footers.docx', 'large/anon-250.docx'];

/** Every table node in a document, nested ones included. */
function tablesOf(doc: PMNode): PMNode[] {
  const out: PMNode[] = [];
  doc.descendants((n) => { if (n.type.name === 'table') out.push(n); return true; });
  return out;
}

describe('cellConditionsOf: the engine\'s conditions for each cell, over a stand-in of the table node', () => {
  for (const name of FIXTURES) {
    test(name, async () => {
      const session = await EditorDocument.open(await fixture(name));
      const resolver = session.pkg.propertyResolverOrUndefined!;
      const ids = session.ctx.ids;
      let cells = 0;
      let formatted = 0;
      let texts = 0;
      for (const table of tablesOf(session.doc)) {
        const tbl = ids.get(table.attrs['sourceId'] as string) as Element<wml.Tbl> | undefined;
        if (!tbl) continue;
        const engine = resolver.tableContext(tbl.value);
        const ours = cellConditionsOf(session, table);
        // Every paragraph directly in a cell: its text's context is the engine's for it (its own w:cnfStyle honoured).
        const { paragraphs } = tableConditionsOf(session, table);
        table.forEach((row) => row.forEach((cell) => cell.forEach((p) => {
          if (p.type.name !== 'paragraph') return;
          const real = ids.get(p.attrs['sourceId'] as string) as Element<wml.P> | undefined;
          if (!real) return;
          const expected = resolver.cellContextOf(real.value);
          const got = paragraphs.get(p)?.context ?? ours.get(cell)?.cell;
          if (!expected?.formatsText) return;
          texts++;
          expect(got?.key.equals(expected.key), `${name} paragraph "${p.textContent.slice(0, 30)}"`).toBe(true);
        })));
        table.forEach((row) => {
          const tr = ids.get(row.attrs['sourceId'] as string) as Element<wml.Tr> | undefined;
          row.forEach((cell) => {
            const tc = ids.get(cell.attrs['sourceId'] as string) as Element<wml.Tc> | undefined;
            if (!tr || !tc) return;
            cells++;
            const expected = engine.forCell(tr.value, tc.value);
            const formats = expected.formatsText || engine.tableStyle.tcPr !== undefined || engine.applicable(expected.conditions).some((s) => s.tcPr !== undefined);
            const got = ours.get(cell);
            const where = `${name} cell "${cell.textContent.slice(0, 30)}"`;
            if (!formats) { expect(got, where).toBeUndefined(); return; }
            formatted++;
            expect(got, where).toBeDefined();
            expect([...got!.cell.conditions].sort(), where).toEqual([...expected.conditions].sort());
            expect(got!.cell.tableStyleId, where).toBe(expected.tableStyleId);
            expect(got!.key.startsWith(`tc-${(expected.tableStyleId ?? 'none').replace(/[^A-Za-z0-9_-]/g, (c) => `_${c.charCodeAt(0).toString(16)}`)}-`), where).toBe(true);
            expect(got!.cell.key.equals(expected.key), where).toBe(true);
          });
        });
      }
      expect(cells).toBeGreaterThan(0);
      if (!name.startsWith('tables/tables-banding')) { expect(formatted).toBeGreaterThan(0); expect(texts).toBeGreaterThan(0); }
    });
  }

  test('conditionalTcPr: the style\'s cell properties, then each conditional format\'s in order', async () => {
    const session = await EditorDocument.open(await fixture('tables/tables-conditions.docx'));
    let checked = 0;
    for (const table of tablesOf(session.doc)) {
      for (const [, c] of cellConditionsOf(session, table)) {
        const merged = conditionalTcPr(c);
        const last = [c.table.tableStyle.tcPr, ...c.table.applicable(c.cell.conditions).map((s) => s.tcPr)].filter((p) => p?.shd).pop();
        expect(merged.shd).toEqual(last?.shd);
        checked++;
      }
    }
    expect(checked).toBeGreaterThan(0);
  });
});
