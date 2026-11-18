// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-007's fixtures (section 5): five tables whose column widths an optimiser can improve, in two files, one
// with fixed widths and one automatic (Word's default table). Run from packages/editor-model after
// `npm run build`:
//
//   node test/fixtures/build/table-fit.mjs [out-dir]
//
// (the default output is test/fixtures/tables/, where table-fit-fixed.docx and table-fit-auto.docx go)
//
// The tables, each under a heading paragraph: A, three columns of short identifiers, long descriptions and short
// dates at equal widths (the pasted example); B, a cell merged over two columns and a vertical merge; C, a row of
// exact height, a row of minimum height taller than its content, and a long identifier that cannot break (the fit
// releases both heights, section 3.5); D, a wide table of five columns with one long column; E, ten narrow columns
// whose longest words together are wider than the table (section 8.1 item 12).
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '@docx4j/core-ts';

const here = dirname(fileURLToPath(import.meta.url));
const outDir = process.argv[2] ?? join(here, '..', 'tables');
const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;');
const p = (text) => `<w:p><w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
const TEXT = 9026;   // A4 with Word's margins
const LONG = 'The tenant pays the rent monthly in advance on the first business day of each month, by electronic transfer to the account the landlord names in writing, and the landlord gives a receipt within seven days.';
const LONG2 = 'Where the premises are damaged by fire, storm or flood so that they are wholly or partly unfit for occupation, the rent abates in proportion until the premises are made fit again.';

/** A cell: text paragraphs, a width in twips (or null for automatic), a span, a vertical merge value. */
function tc(texts, { width = null, span = 1, vMerge, exact } = {}) {
  const pr = [
    width === null ? '<w:tcW w:w="0" w:type="auto"/>' : `<w:tcW w:w="${width}" w:type="dxa"/>`,
    span > 1 ? `<w:gridSpan w:val="${span}"/>` : '',
    vMerge === 'restart' ? '<w:vMerge w:val="restart"/>' : vMerge === 'continue' ? '<w:vMerge/>' : '',
  ].join('');
  return `<w:tc><w:tcPr>${pr}</w:tcPr>${(texts.length ? texts : ['']).map(p).join('')}</w:tc>`;
}
function tr(cells, { exact, atLeast } = {}) {
  const height = exact ? `<w:trHeight w:val="${exact}" w:hRule="exact"/>` : atLeast ? `<w:trHeight w:val="${atLeast}" w:hRule="atLeast"/>` : '';
  return `<w:tr>${height ? `<w:trPr>${height}</w:trPr>` : ''}${cells.join('')}</w:tr>`;
}
function tbl(grid, rows, { fixed }) {
  const sum = grid.reduce((a, b) => a + b, 0);
  const pr = fixed
    ? `<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${sum}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>`
    : `<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>`;
  return `<w:tbl>${pr}<w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows.join('')}</w:tbl>`;
}

function tables(fixed) {
  const w3 = [3009, 3009, 3008];
  const A = tbl(w3, [
    tr([tc(['Id'], { width: w3[0] }), tc(['Obligation'], { width: w3[1] }), tc(['Due'], { width: w3[2] })]),
    tr([tc(['R1'], { width: w3[0] }), tc([LONG], { width: w3[1] }), tc(['1 July'], { width: w3[2] })]),
    tr([tc(['R2'], { width: w3[0] }), tc([LONG2], { width: w3[1] }), tc(['1 August'], { width: w3[2] })]),
    tr([tc(['R3'], { width: w3[0] }), tc(['Keep the premises clean.'], { width: w3[1] }), tc(['Always'], { width: w3[2] })]),
  ], { fixed });
  const B = tbl(w3, [
    tr([tc(['Merged over two columns: the heading of the clause'], { width: w3[0] + w3[1], span: 2 }), tc(['Note'], { width: w3[2] })]),
    tr([tc(['R1'], { width: w3[0], vMerge: 'restart' }), tc([LONG], { width: w3[1] }), tc(['Short'], { width: w3[2] })]),
    tr([tc([], { width: w3[0], vMerge: 'continue' }), tc([LONG2], { width: w3[1] }), tc(['Short'], { width: w3[2] })]),
  ], { fixed });
  const C = tbl(w3, [
    tr([tc(['Exact'], { width: w3[0] }), tc([LONG], { width: w3[1] }), tc(['x'], { width: w3[2] })], { exact: 600 }),
    tr([tc(['ChartStylePart/ChartColorStylePart'], { width: w3[0] }), tc([LONG2], { width: w3[1] }), tc(['y'], { width: w3[2] })], { atLeast: 1800 }),
  ], { fixed });
  const w5 = [1805, 1805, 1805, 1805, 1806];
  const D = tbl(w5, [
    tr(w5.map((w, i) => tc([`H${i + 1}`], { width: w }))),
    tr([tc(['a'], { width: w5[0] }), tc(['b'], { width: w5[1] }), tc([LONG], { width: w5[2] }), tc(['d'], { width: w5[3] }), tc(['e'], { width: w5[4] })]),
    tr([tc(['f'], { width: w5[0] }), tc(['g'], { width: w5[1] }), tc([LONG2], { width: w5[2] }), tc(['h'], { width: w5[3] }), tc(['i'], { width: w5[4] })]),
  ], { fixed });
  // E: ten narrow columns under a heading row of product names, whose longest words together are wider than the table
  // (section 8.1 item 12): no widths keep every word whole, and the body's "price" breaks in the narrowest columns.
  const w10 = [1309, 1122, 1636, 561, 1215, 654, 561, 561, 795, 612];
  const names = ['', 'Alpha Enterprise', 'Beta.Woods', 'Gamma.Den', 'DeltaBay', 'EPSILON DsWood', 'Zetaforest', 'EtaEastern', 'Theta', 'IotaContour'];
  const priced = (label, second) => tr([tc([label], { width: w10[0] }), tc([second], { width: w10[1] }), ...w10.slice(2).map((w) => tc(['in the price'], { width: w }))]);
  const E = tbl(w10, [
    tr(names.map((n, i) => tc(n ? [n] : [], { width: w10[i] }))),
    priced('Engine', 'free (ASLv2)'),
    tr([tc(['Internal use, 1 developer'], { width: w10[0] }), tc(['$1,250'], { width: w10[1] }), tc(['$1,100 perp'], { width: w10[2] }), tc(['$80'], { width: w10[3] }), tc(['$800 perp'], { width: w10[4] }), tc(['$900'], { width: w10[5] }), ...w10.slice(6).map((w) => tc(['—'], { width: w }))]),
    priced('Support', 'community'),
    priced('Updates', 'free'),
    priced('Source', 'yes'),
    priced('Licence', 'ASLv2'),
  ], { fixed });
  return [['A. Three columns, equal widths, one long', A], ['B. A merged cell over two columns and a vertical merge', B], ['C. A row of exact height, one of minimum height, and an identifier that cannot break', C], ['D. Five columns, one long', D], ['E. Ten narrow columns whose longest words together are wider than the table', E]];
}

for (const fixed of [true, false]) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = await pkg.getBody();
  const xml = tables(fixed).map(([heading, t]) => p(heading) + t + '<w:p/>').join('') + p('End.');
  await body.insertXml(xml, 'End');
  const out = join(outDir, fixed ? 'table-fit-fixed.docx' : 'table-fit-auto.docx');
  await mkdir(dirname(out), { recursive: true });
  await writeFile(out, await pkg.save());
  console.log(`written ${out}`);
}
