// SPDX-License-Identifier: GPL-3.0-or-later
//
// Word check 42 (ED-005 section 12.83): what Word's Insert Table dialog writes for each of its choices.
// Builds one document of four labelled paragraphs, each followed by an empty paragraph where Jason inserts a
// table from the dialog in Word 365. Run from packages/editor-model after `npm run build`:
//
//   node test/fixtures/build/insert-table-probe.mjs [out.docx]
//
// (the default output is test-results/insert-table-probe.docx, which git ignores)
//
// Read back from the save: each table's w:tblW, w:tblLayout (present or not), w:tblGrid and the cells' w:tcW,
// against what the editor's table.insert writes for the same choice (section 12.83 item 1).
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
const here = dirname(fileURLToPath(import.meta.url));
const out = process.argv[2] ?? join(here, '..', '..', '..', 'test-results', 'insert-table-probe.docx');
const pkg = await WordprocessingMLPackage.createPackage();
const body = await pkg.getBody();
const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;');
const steps = [
  'Check 42: under each heading, put the caret in the empty paragraph and use Insert > Table > Insert Table..., then Save As. Leave the tables empty.',
  'A. Fixed column width: Auto (the defaults: 5 columns, 2 rows)',
  'B. Fixed column width: 1" (72 pt), 3 columns, 1 row',
  'C. AutoFit to contents, 3 columns, 1 row',
  'D. AutoFit to window, 3 columns, 1 row',
];
const xml = steps.map((s, i) => `<w:p><w:r><w:t xml:space="preserve">${esc(s)}</w:t></w:r></w:p>` + (i > 0 ? '<w:p/>' : '')).join('') + '<w:p><w:r><w:t>End.</w:t></w:r></w:p>';
await body.insertXml(xml, 'End');
await mkdir(dirname(out), { recursive: true });
await writeFile(out, await pkg.save());
console.log(`written ${out}`);
