// SPDX-License-Identifier: GPL-3.0-or-later
//
// The file for core-ts check 35's part 2 in Word 2010 (ED-005 section 12.58): what the editor writes for rich text
// mapping, opened in a Word that knows no `w15`. Made from check 35's `35-input.docx`, which declares no `w15`, so the
// save itself must declare it ignorable. Run from packages/editor-model after `npm run build`:
//
//   node test/fixtures/build/check35-for-word.mjs [out.docx]
//
// (the default output is test-results/35d-editor.docx, which git ignores)
//
// In it: 1a's rich text control mapped to `name` with `w15:dataBinding` (showing "Ann"), and inside it a plain text
// control the Data tab made, bound to `rich2` with `w:dataBinding` (empty: the placeholder); 1b's block-level rich text
// control mapped to `rich1` with `w15:dataBinding` (empty: the placeholder); 1f's rich text control as the engine bound
// it (`w:dataBinding` to `bound`, showing "Old text", the node "Bea"), untouched; "After." wrapped in a repeating
// section (`w15:repeatingSection`). Word 2010 should open it without complaint, ignore the `w15` bindings and the
// repeating section, and show 1f as the node's "Bea" if it takes `w:dataBinding` on rich text as Word 15 does.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { EditorDocument, runCommand } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const out = process.argv[2] ?? join(here, '..', '..', '..', 'test-results', '35d-editor.docx');
await mkdir(dirname(out), { recursive: true });
const session = await EditorDocument.open(new Uint8Array(await readFile(join(here, '..', 'revisions', 'check35', '35-input.docx'))));
await session.template.load();
const part = session.template.dataPart.part;

/** Where `text` starts in the document. */
const at = (text) => {
  let found = -1;
  session.doc.descendants((n, pos) => { if (found < 0 && n.isText && n.text.includes(text)) found = pos + n.text.indexOf(text); return found < 0; });
  if (found < 0) throw new Error(`no "${text}"`);
  return found;
};
const run = (name, args, selection) => {
  const result = runCommand(session, name, args, selection);
  if (!result.ok) throw new Error(`${name} refused`);
};
const map = (xpath) => ({ xpath, storeItemID: part.id });

run('contentControl.bind', map('/data[1]/name[1]'), { from: at('quick') + 1 });          // 1a, as Map to Selected
run('contentControl.insertBound', map('/data[1]/rich2[1]'), { from: at('1a. The ') + '1a. The '.length + 1 + 2 });   // the Data tab, inside 1a: "An|n"
run('contentControl.bind', map('/data[1]/rich1[1]'), { from: at('First paragraph') + 1 });   // 1b
const after = at('After.');
run('contentControl.insert', { kind: 'RepeatingSection' }, { from: after, to: after + 'After.'.length });

const saved = await session.save();
await writeFile(out, saved.bytes);
console.log(`wrote ${out}`);
