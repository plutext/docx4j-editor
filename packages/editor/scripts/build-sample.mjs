// SPDX-License-Identifier: GPL-3.0-or-later
//
// The site's sample document (the release plan's phase C, step 2; Jason, 2026-11-19): "Hello World" in Heading 1, then
// "from docx4j-editor" in Normal, made with the engine. It is written to public/samples/hello-world.docx, which the
// site serves and the drop zone's "Open the sample document" opens.
//
//   node packages/editor/scripts/build-sample.mjs
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '@docx4j/core-ts';

const out = join(dirname(fileURLToPath(import.meta.url)), '..', 'public', 'samples', 'hello-world.docx');
const pkg = await WordprocessingMLPackage.createPackage({ pageSize: 'A4' });
await pkg.styles.ensure('Heading1');
const body = await pkg.getBody();
await body.insertXml(
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Hello World</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t>from docx4j-editor</w:t></w:r></w:p>',
  'End',
);
await mkdir(dirname(out), { recursive: true });
await writeFile(out, await pkg.save());
console.log(`written ${out}`);
