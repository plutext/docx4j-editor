// SPDX-License-Identifier: GPL-3.0-or-later
//
// Check 44 (ED-007 section 8.1 item 13): how Word breaks a word in a cell too narrow for it, with and without an
// optional break: a zero-width space (U+200B, Word's "No-Width Optional Break") and w:softHyphen (Word's "Optional
// Hyphen"). Run from packages/editor-model after `npm run build`:
//
//   node test/fixtures/build/optional-breaks-probe.mjs [out-dir]
//
// (the default output is test/fixtures/tables/check44/optional-breaks-probe.docx). Open it in Word 365 and Word 2010:
// the expectation is that column 1 breaks mid-word wherever Word must, column 2 at the period or the capital with
// nothing drawn, and column 3 at the capital with a hyphen drawn; a save back into the same folder records it.
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '@docx4j/core-ts';

const here = dirname(fileURLToPath(import.meta.url));
const outDir = process.argv[2] ?? join(here, '..', 'tables', 'check44');
const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;');
const p = (text) => `<w:p><w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
const ZWSP = '​';
const W = 700;
const tc = (xml) => `<w:tc><w:tcPr><w:tcW w:w="${W}" w:type="dxa"/></w:tcPr>${xml}</w:tc>`;
const softHyphen = (before, after) => `<w:p><w:r><w:t>${esc(before)}</w:t><w:softHyphen/><w:t>${esc(after)}</w:t></w:r></w:p>`;
const grid = [W, W, W, W];
const rows = [
  [p('plain'), p('U+200B'), p('w:softHyphen'), p('hyphen')],
  [p('Gamma.Den'), p(`Gamma.${ZWSP}Den`), softHyphen('Gamma.', 'Den'), p('Gamma.-Den')],
  [p('EtaEastern'), p(`Eta${ZWSP}Eastern`), softHyphen('Eta', 'Eastern'), p('Eta-Eastern')],
  [p('Lambda.Woods'), p(`Lambda.${ZWSP}Woods`), softHyphen('Lambda.', 'Woods'), p('Lambda.-Woods')],
  [p('in the price'), p(`in the pri${ZWSP}ce`), softHyphen('in the pri', 'ce'), p('in the pri-ce')],
];
const table = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${W * 4}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows.map((cells) => `<w:tr>${cells.map(tc).join('')}</w:tr>`).join('')}</w:tbl>`;
const pkg = await WordprocessingMLPackage.createPackage();
const body = await pkg.getBody();
await body.insertXml(p('Check 44: a word in a cell too narrow for it. Column 1 is the plain word; column 2 has a zero-width space (U+200B) at the period or the capital; column 3 has w:softHyphen there; column 4 a hyphen. Save this file back into the same folder after opening it, and screenshot the table.') + table + '<w:p/>' + p('End.'), 'End');
const out = join(outDir, 'optional-breaks-probe.docx');
await mkdir(dirname(out), { recursive: true });
await writeFile(out, await pkg.save());
console.log(`written ${out}`);
