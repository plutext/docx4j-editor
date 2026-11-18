// SPDX-License-Identifier: GPL-3.0-or-later
//
// Keeps only the redistributable font in a document Word saved with fonts embedded (ED-003
// section 11.18 item 5): Microsoft 365 Word embedded the document's default font, Aptos, whole
// beside the Noto Sans subsets, and Aptos is Microsoft's, so the save cannot be a fixture as it
// is. Every `w:embed*` element on a `w:font` other than the ones kept is removed, with its
// relationship and its part; the `w:font` entries themselves stay, since the document's styles
// name them. Nothing else is touched. Run from packages/editor-model:
//
//   node test/fixtures/build/fonts-embedded-strip.mjs <in.docx> <out.docx> [family ...]
//
// The families to keep default to Noto Sans. The output is the fixture
// (`fonts-embedded-subset-m365.docx` was made this way from Jason's Microsoft 365 save).
import { readFile, writeFile } from 'node:fs/promises';
import { WordprocessingMLPackage } from '@docx4j/core-ts';

const [input, output, ...families] = process.argv.slice(2);
if (!input || !output) { console.error('usage: fonts-embedded-strip.mjs <in.docx> <out.docx> [family ...]'); process.exit(2); }
const keep = new Set(families.length > 0 ? families : ['Noto Sans']);
const EMBEDS = ['embedRegular', 'embedBold', 'embedItalic', 'embedBoldItalic'];

const pkg = await WordprocessingMLPackage.load(new Uint8Array(await readFile(input)));
const table = pkg.mainDocumentPart.fontTablePart;
if (!table) throw new Error('no font table part');
const fonts = await table.readContents();
const rels = table.getRelationshipsPart();
const removed = [];
for (const font of fonts.font ?? []) {
  if (keep.has(font.name)) continue;
  for (const el of EMBEDS) {
    const rel = font[el];
    if (!rel?.id) continue;
    const part = rels.getPart(rel.id);
    if (part) removed.push(...rels.removePart(part.partName).map(String));
    else rels.removeRelationship(rel.id);
    delete font[el];
    console.log(`${font.name}: ${el} removed (${rel.id})`);
  }
}
table.setContents(fonts);
await writeFile(output, await pkg.save());
console.log(`removed parts: ${removed.join(', ') || 'none'}; written ${output}`);
