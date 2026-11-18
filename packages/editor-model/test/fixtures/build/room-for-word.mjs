// SPDX-License-Identifier: GPL-3.0-or-later
//
// The file for ED-004 section 6.4's check in Word 2010 on the reference machine: a document saved
// from a room, so its new identifiers are the room's random ones (section 3.4): a list (random
// `w:numId` and `w:abstractNumId`), two footnotes (random note ids), two links (`rId` and eight
// base-36 characters) and two content controls. Word opens it, saves it, and the editor reopens
// Word's save. Run from packages/editor-model after `npm run build`:
//
//   node test/fixtures/build/room-for-word.mjs [out.docx] [--no-notes]
//
// (the default output is test-results/room-for-word-2010.docx, which git ignores)
//
// `--no-notes` leaves the two footnotes out (Word 2010 refused the room's random note ids,
// ED-004 section 10.8 item 6), so the list, links and controls can be checked on their own.
// It prints the identifiers it wrote, to compare with what Word's save keeps.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { EditorDocument, runCommandAsync } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const args = process.argv.slice(2);
const notes = !args.includes('--no-notes');
// By default under the package's `test-results/`, which git ignores: the file is for Word, not the repository.
const out = args.find((a) => !a.startsWith('--')) ?? join(here, '..', '..', '..', 'test-results', notes ? 'room-for-word-2010.docx' : 'room-for-word-2010-no-notes.docx');
await mkdir(dirname(out), { recursive: true });
const bytes = new Uint8Array(await readFile(join(here, '..', 'tracked-changes.docx')));
const session = await EditorDocument.open(bytes);
await session.share(bytes, 'tracked-changes.docx');

const blockFrom = (i) => { let pos = 0; for (let k = 0; k < i; k++) pos += session.doc.child(k).nodeSize; return pos + 1; };
const run = async (name, args, from, to) => {
  const result = await runCommandAsync(session, name, args, { from, ...(to !== undefined ? { to } : {}) });
  if (!result.ok) throw new Error(`${name} did not apply`);
};

await run('list.toggle', { bullet: false }, blockFrom(4));
if (notes) {
  await run('range.insertFootnote', { text: 'The first note, made in a room.' }, blockFrom(1) + 3);
  await run('range.insertFootnote', { text: 'The second note, made in a room.' }, blockFrom(4) + 3);
}
await run('range.hyperlink', { address: 'https://www.docx4java.org/' }, blockFrom(1) + 5, blockFrom(1) + 9);
await run('range.hyperlink', { address: 'https://github.com/plutext/docx4j-ts-editor' }, blockFrom(4) + 6, blockFrom(4) + 10);
await run('contentControl.insert', { kind: 'RichText', tag: 'first' }, blockFrom(12), blockFrom(12) + 4);
await run('contentControl.insert', { kind: 'PlainText', tag: 'second' }, blockFrom(13), blockFrom(13) + 4);

const saved = await session.save();
await writeFile(out, saved.bytes);
const ids = { numId: session.doc.child(4).attrs.numPr?.numId, notes: [], links: new Set(), controls: [] };
session.doc.descendants((n) => {
  if (n.type.name === 'footnoteRef') ids.notes.push(n.attrs.footnoteId);
  for (const m of n.marks) if (m.type.name === 'link' && m.attrs.rId) ids.links.add(m.attrs.rId);
  const id = (n.attrs?.sdtPr?.rPrOrAliasOrLock ?? []).find((e) => e.name?.localPart === 'id');
  if (id) ids.controls.push(id.value.val);
  return true;
});
console.log(`wrote ${out}: numId ${ids.numId}; notes ${ids.notes.join(', ')}; links ${[...ids.links].join(', ')}; controls ${ids.controls.join(', ')}`);
