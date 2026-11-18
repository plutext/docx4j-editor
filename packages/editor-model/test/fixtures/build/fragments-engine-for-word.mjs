// SPDX-License-Identifier: GPL-3.0-or-later
//
// The files for ED-005 proposal 46's Word check, part 6(b): the engine's own output (core-ts CR-005 section 11, from f3ce1cc)
// over check 46c's template, for Word 365 through the portfolio's Word bridge. The template holds OpenDoPE's od:progid
// control and Word's w15-mapped rich text, both bound to the same Flat OPC node, and a plain text control mapped to it. From
// the repository root, with that engine installed:
//
//   node packages/editor-model/test/fixtures/build/fragments-engine-for-word.mjs [out-dir]
//
// 46d: applyBindings with the fragment pass's defaults (resolved; the w15 mapping kept);
// 46e: route 'altChunk', docx4j's form, for Word to resolve;
// 46f: resolved with keepSourceFormatting (Word's matchSrc).
import { readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '@docx4j/core-ts';

const here = dirname(fileURLToPath(import.meta.url));
const out = process.argv[2] ?? join(here, '..', 'fragments');
const template = await readFile(join(here, '..', 'fragments', '46c-template.docx'));

for (const [name, fragments] of [
  ['46d-engine-resolved.docx', {}],
  ['46e-engine-altchunk.docx', { route: 'altChunk' }],
  ['46f-engine-keepsource.docx', { keepSourceFormatting: true }],
]) {
  const pkg = await WordprocessingMLPackage.load(new Uint8Array(template));
  await pkg.customXmlParts.load();
  const result = await pkg.customXmlParts.applyBindings({ fragments });
  await writeFile(join(out, name), await pkg.save());
  console.log(name, JSON.stringify({ bound: result.bound, updated: result.updated, skipped: result.skipped, notes: result.notes ?? [] }));
}
