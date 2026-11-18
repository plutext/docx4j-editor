// SPDX-License-Identifier: GPL-3.0-or-later
//
// Builds `nested-table.docx` with the engine (Jason, 2026-09-25, after E2.d: nothing in the suites
// held a table inside a table): a two-by-two table whose second cell of the first row holds a
// two-by-two table followed by a paragraph. Run from packages/editor-model:
//
//   node test/fixtures/build/nested-table.mjs
//
// and commit the output; the suites read the committed document.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '@docx4j/core-ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const cell = (inner, w) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
const p = (text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const table = (cols, rows) => `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${cols.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows.map((cells) => `<w:tr>${cells.map((c, i) => cell(c, cols[i])).join('')}</w:tr>`).join('')}</w:tbl>`;

const inner = table([2000, 2000], [[p('inner a1'), p('inner a2')], [p('inner b1'), p('inner b2')]]);
const outer = table([4500, 4500], [[p('outer a1'), inner + p('after the inner table')], [p('outer b1'), p('outer b2')]]);
const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}"><w:body>${p('Before the table.')}${outer}${p('After the table.')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;

const pkg = await WordprocessingMLPackage.createPackage();
pkg.getMainDocumentPart().setXml(document);
const out = join(dirname(fileURLToPath(import.meta.url)), '..', 'nested-table.docx');
await writeFile(out, await pkg.save());
console.log(`wrote ${out}`);
