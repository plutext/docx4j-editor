// SPDX-License-Identifier: GPL-3.0-or-later
//
// A table style's conditional formatting on the page (ED-005 section 12.65), held to the engine (core-ts CR-007): for every
// cell the page gives a class of conditions, its first paragraph's computed bold, italic, colour and size are what the
// engine resolves for that paragraph in its cell (`getEffectiveRPr` with `cellContextOf`), and the cell's background is the
// shading that applies to it (the style's, then each conditional format's in order, then the cell's own).
import { test, expect, type Page } from '@playwright/test';
import { openFixture } from './helpers.js';

interface Finding { cell: string; property: string; page: string; engine: string }

/** Compare every classed cell of the document on the page with the engine; the findings, and how many cells were compared. */
async function compare(page: Page): Promise<{ checked: number; findings: Finding[] }> {
  return page.evaluate(() => {
    type Any = any; // eslint-disable-line @typescript-eslint/no-explicit-any
    const v = (window as unknown as { docx4jEditor: Any }).docx4jEditor;
    const session = v.session;
    const resolver = session.pkg.propertyResolverOrUndefined;
    const ids = session.ctx.ids;
    const findings: { cell: string; property: string; page: string; engine: string }[] = [];
    let checked = 0;
    const rgb = (hex: string): string => `rgb(${parseInt(hex.slice(0, 2), 16)}, ${parseInt(hex.slice(2, 4), 16)}, ${parseInt(hex.slice(4, 6), 16)})`;
    const on = (b: Any): boolean => b !== undefined && b !== null && b.val !== false && b.val !== 0 && b.val !== '0' && b.val !== 'false';
    v.state.doc.descendants((node: Any, pos: number) => {
      if (node.type.name !== 'table') return true;
      const tbl = ids.get(node.attrs.sourceId)?.value;
      if (!tbl) return true;
      const context = resolver.tableContext(tbl);
      node.forEach((row: Any, rowOffset: number) => {
        row.forEach((cell: Any, cellOffset: number) => {
          const cellPos = pos + 1 + rowOffset + 1 + cellOffset;
          const td = v.view.nodeDOM(cellPos) as HTMLElement | null;
          if (!td || ![...td.classList].some((c) => c.startsWith('tc-'))) return;
          const p = cell.firstChild;
          const realP = p?.type.name === 'paragraph' ? ids.get(p.attrs.sourceId)?.value : undefined;
          const tc = ids.get(cell.attrs.sourceId)?.value;
          const tr = ids.get(row.attrs.sourceId)?.value;
          if (!realP || !tc || !tr) return;
          checked++;
          const name = `${cell.textContent.slice(0, 24)}`;
          const rPr = resolver.getEffectiveRPr(undefined, realP.pPr, resolver.cellContextOf(realP));
          const pDom = v.view.nodeDOM(cellPos + 1) as HTMLElement;
          const style = getComputedStyle(pDom);
          const bold = Number(style.fontWeight) >= 600;
          if (bold !== on(rPr.b)) findings.push({ cell: name, property: 'bold', page: style.fontWeight, engine: JSON.stringify(rPr.b ?? null) });
          const italic = style.fontStyle === 'italic';
          if (italic !== on(rPr.i)) findings.push({ cell: name, property: 'italic', page: style.fontStyle, engine: JSON.stringify(rPr.i ?? null) });
          const color = rPr.color?.val;
          if (typeof color === 'string' && /^[0-9A-Fa-f]{6}$/.test(color) && style.color !== rgb(color)) findings.push({ cell: name, property: 'color', page: style.color, engine: color });
          const sz = rPr.sz?.val;
          if (typeof sz === 'number' && Math.abs(parseFloat(style.fontSize) - (sz / 2) * (4 / 3)) > 0.2) findings.push({ cell: name, property: 'size', page: style.fontSize, engine: `${sz / 2}pt` });
          // The cell's shading: the style's, then the conditional formats', then its own.
          const cellContext = context.forCell(tr, tc);
          const layers = [context.tableStyle.tcPr, ...context.applicable(cellContext.conditions).map((s: Any) => s.tcPr), tc.tcPr].filter((l: Any) => l?.shd);
          const fill = layers.length > 0 ? layers[layers.length - 1].shd.fill : undefined;
          if (typeof fill === 'string' && /^[0-9A-Fa-f]{6}$/.test(fill)) {
            const bg = getComputedStyle(td).backgroundColor;
            if (bg !== rgb(fill)) findings.push({ cell: name, property: 'shading', page: bg, engine: fill });
          }
        });
      });
      return true;
    });
    return { checked, findings };
  });
}

for (const name of ['tables/tables-conditions.docx', 'tables/tables-banding-band-size.docx', 'headers-footers.docx']) {
  test(`${name}: each classed cell's text and shading as the engine resolves them`, async ({ page }) => {
    await openFixture(page, name);
    const { checked, findings } = await compare(page);
    expect(checked).toBeGreaterThan(0);
    expect(findings).toEqual([]);
  });
}

test('the classes follow the table as it changes: a banded row deleted, the row after it takes the band', async ({ page }) => {
  await openFixture(page, 'tables/tables-banding-band-size.docx');
  const rows = (): Promise<string[]> => page.evaluate(() => {
    const table = document.querySelectorAll('.docx-editor .ProseMirror[aria-label="Document"] table')[3]!;
    return [...table.querySelectorAll(':scope > tbody > tr')].map((tr) => [...tr.children].map((td) => [...td.classList].find((c) => c.startsWith('tc-')) ?? '-').join(' '));
  });
  const before = await rows();
  // The fourth table: a header row, then rows banded one by one (the first band formatted), a merge in row 1.
  expect(before[3]).toBe('tc-ProbeBands-128 tc-ProbeBands-128');
  expect(before[4]).toBe('- -');
  // Row 3 deleted: the row that was 4 is the third band row now, formatted as such; the last moves up into the unformatted band.
  await page.evaluate(() => {
    type V = { state: { doc: { descendants(f: (n: { type: { name: string }; child(i: number): { nodeSize: number } }, p: number) => boolean | void): void } }; caretAt(p: number): void; run(n: string, a?: unknown): boolean };
    const v = (window as unknown as { docx4jEditor: V }).docx4jEditor;
    let seen = 0;
    let at = -1;
    v.state.doc.descendants((n, p) => {
      if (at >= 0 || n.type.name !== 'table') return at < 0;
      if (seen++ === 3) { let row = p + 1; for (let r = 0; r < 3; r++) row += n.child(r).nodeSize; at = row + 3; }
      return false;
    });
    v.caretAt(at);
    v.run('table.deleteRows');
  });
  await expect.poll(rows).toEqual([before[0], before[1], before[2], 'tc-ProbeBands-128 tc-ProbeBands-128', '- -']);
});
