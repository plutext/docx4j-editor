// SPDX-License-Identifier: GPL-3.0-or-later
//
// Word check 39 (ED-005 section 12.78): what Insert Column writes in each of Word's table layout modes.
// Builds one document with three two-column tables, each a paragraph apart: fixed column widths
// (w:tblLayout fixed, the grid two equal halves of the text width), AutoFit to window (w:tblW 100%, no
// layout) and AutoFit to contents (w:tblW auto, no layout). Run from packages/editor-model after `npm run build`:
//
//   node test/fixtures/build/table-autofit-probe.mjs [out.docx]
//
// (the default output is test-results/table-autofit-probe.docx, which git ignores)
//
// In Word 2010 and Word 15: put the caret in the first cell of each table, Table Tools > Layout > Insert Right,
// then Save As beside the original. Read back: the new column's w:tcW and the grid in each table; in the
// AutoFit-to-window table whether the three widths are the old two scaled alike or an equal share; what the
// AutoFit-to-contents table's cells say (auto or dxa) and its grid.
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
const here = dirname(fileURLToPath(import.meta.url));
const out = process.argv[2] ?? join(here, '..', '..', '..', 'test-results', 'table-autofit-probe.docx');
const pkg = await WordprocessingMLPackage.createPackage();
const body = await pkg.getBody();
body.insertParagraph('Check 39: Insert Right in the first cell of each table, then Save As. Each table is two columns of the text width.', 'End');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const half = Math.floor(9026 / 2);
const table = (label, tblPrInner, tcW) => `<w:p><w:r><w:t xml:space="preserve">${label}</w:t></w:r></w:p>` +
  `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>${tblPrInner}</w:tblPr><w:tblGrid><w:gridCol w:w="${half}"/><w:gridCol w:w="${half}"/></w:tblGrid>` +
  [1, 2].map((r) => `<w:tr>${['A', 'B'].map((c) => `<w:tc><w:tcPr>${tcW}</w:tcPr><w:p><w:r><w:t>${label[0]}${r}${c}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`).join('') + '</w:tbl>';
const xml = table('Fixed column width', `<w:tblW w:w="${half * 2}" w:type="dxa"/><w:tblLayout w:type="fixed"/>`, `<w:tcW w:w="${half}" w:type="dxa"/>`)
  + table('AutoFit to window', '<w:tblW w:w="5000" w:type="pct"/>', '<w:tcW w:w="2500" w:type="pct"/>')
  + table('AutoFit to contents', '<w:tblW w:w="0" w:type="auto"/>', '<w:tcW w:w="0" w:type="auto"/>')
  + '<w:p><w:r><w:t>End.</w:t></w:r></w:p>';
void W;
await body.insertXml(xml, 'End');
await pkg.styles.ensure(['TableGrid']).catch(() => {});
await mkdir(dirname(out), { recursive: true });
await writeFile(out, await pkg.save());
console.log(`written ${out}`);
