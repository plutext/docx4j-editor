// SPDX-License-Identifier: GPL-3.0-or-later
//
// The file for ED-009 section 4's check in Word 365 (step S6, Jason's): the shapes the editor writes
// for a kept `mc:AlternateContent` whose selected branch was edited. Run from packages/editor-model
// after `npm run build`:
//
//   node test/fixtures/build/mc-for-word.mjs [out.docx]
//
// (the default output is test-results/mc-for-word.docx, which git ignores). From `mc/mc-branches.docx`:
//   1 (case 1) "XX" typed into the Choice, tracking off: the Choice changed, the Fallback as it was;
//   2 (case 6) "w1" of w14 made bold, tracking off: two elements, the first with its nine other Choices
//     and the Fallback, the second ("4") with every other branch present and empty;
//   3 (case 2) "after the tab" deleted with tracking on: w:delText in the Choice inside w:del, the
//     Fallback of the first piece as it was, the deleted piece's empty;
//   4 (case 8) "XX" typed into the paragraph-level element's second run, tracking off.
// What to see in Word: the Choice's text in every case (never "The Earth is round", "Fallback text",
// "yang"); item 3 as a tracked deletion. Then save in Word and reopen the save in the editor.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { EditorDocument, runCommand, runCommandAsync } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const out = process.argv[2] ?? join(here, '..', '..', '..', 'test-results', 'mc-for-word.docx');
await mkdir(dirname(out), { recursive: true });
const session = await EditorDocument.open(new Uint8Array(await readFile(join(here, '..', 'mc', 'mc-branches.docx'))));

/** Where `text` starts in the document, and a selection over it. */
const find = (text) => {
  let at = -1;
  session.doc.descendants((n, pos) => { if (at < 0 && n.isText && n.text.includes(text)) at = pos + n.text.indexOf(text); return at < 0; });
  if (at < 0) throw new Error(`no "${text}"`);
  return { from: at, to: at + text.length };
};
const ok = (r, what) => { if (!r.ok) throw new Error(`${what}: ${r.reason ?? 'refused'}`); };

ok(runCommand(session, 'insertText', { text: 'XX' }, { from: find('The Earth is flat').from + 4 }), 'item 1');
const w14 = find('w14');
ok(runCommand(session, 'font.bold', { value: true }, { from: w14.from, to: w14.from + 2 }), 'item 2');
ok(runCommand(session, 'insertText', { text: 'XX' }, { from: find(' plain ying').from + 1 }), 'item 4');
session.trackingAuthor = 'Editor';
ok(await runCommandAsync(session, 'document.setChangeTrackingMode', { value: 'TrackAll' }), 'tracking');
ok(runCommand(session, 'insertText', { text: '' }, find('after the tab')), 'item 3');
const result = await session.save();
if (!result.selfCheck?.equal) throw new Error('self-check failed');
await writeFile(out, result.bytes);
console.log('wrote', out);
