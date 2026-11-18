// SPDX-License-Identifier: GPL-3.0-or-later
//
// The measurements file of ED-005 section 9.2 item 2 (E4.g step G1): what earlier sub-phases left to
// measure in Word, one section per question, each with its instructions in blue italics and the text
// to act on, on `revisions-base.mjs`'s pattern. Jason opens it on the reference machine, does what
// each section says, and saves it; `fixtures/review/README.md` says where the saves go. The one
// revision it holds, Claude's pending insertion in section 1, is the engine's own (`withTracking`),
// since the question is what Word does beside it; the right-to-left paragraphs are set by the
// editor's `paragraph.bidi`. Run from the repository root after `npm run build`:
//
//   node packages/editor-model/test/fixtures/build/measure-base.mjs [out-dir]
//
// It writes measure-base.docx into out-dir, by default packages/editor-mcp/test-results/word-session/
// (with the rest of the session's files), which git ignores.
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, runCommandAsync } from '../../../dist/index.mjs';

/** Two Hebrew words, "peace" and "world", one space between them. */
export const HEBREW = 'שלום עולם';

/** Builds the file: its bytes. */
export async function buildMeasureBase() {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const heading = (text) => { const h = body.insertParagraph(text, 'End'); h.styleBuiltIn = 'Heading1'; return h; };
  const plain = (text) => body.insertParagraph(text, 'End');
  /** An instruction: blue italics, never acted on. */
  const say = (text) => { const i = body.insertParagraph(text, 'End'); i.font.italic = true; i.font.color = '#1F4E79'; return i; };

  body.insertParagraph('Measurements in Word for the docx4j editor (ED-005 section 9.2)', 'End').styleBuiltIn = 'Title';
  say('Built by docx4j for E4.g. Each section asks one question of Word. Do sections 1 to 3 in Word 15 and save as measure-word15.docx; then, from this file again, do sections 1 to 3 in Word 2010 and save as measure-word2010.docx. Sections 4 and 5 are optional, in Word 15 only, and need screenshots rather than a save.');
  say('Before you start: File > Options > General, User name "Author A", Initials "AA". Act only on the black text; never on a heading or these blue instructions. Accept and reject nothing.');

  heading('1. A comment on one word of a pending insertion');
  say('The paragraph below is Claude\'s tracked insertion, not yet accepted. Select only the word "monthly" (double-click it), Review > New Comment, and type "M1". Track Changes may be on or off: Word does not track comments.');
  await pkg.withTracking({ author: { name: 'Claude' }, mode: 'TrackAll' }, () => { body.insertParagraph('The tenant pays rent monthly in advance.', 'End'); });

  heading('2. Align Left with Track Changes off');
  say('Make sure Review > Track Changes is OFF. Click in the centred paragraph below and press Ctrl+L. It is centred by its own formatting; its style (Normal) is left-aligned.');
  plain('This paragraph is centred by direct formatting and then aligned left.').alignment = 'Centered';

  heading('3. New Comment at a caret, right to left and empty');
  say('The two paragraphs below run right to left: two Hebrew words, the first on the right, one space between them. Do not select anything: click to put the caret where each step says, then Review > New Comment.');
  say('3a. In the first paragraph, click just to the left of the first (rightmost) word, so that the caret touches that word and the space is on the caret\'s left. New Comment, type "M3a".');
  plain(HEBREW);
  say('3b. In the second paragraph, click just to the right of the second (leftmost) word, so that the caret touches that word and the space is on the caret\'s right. New Comment, type "M3b".');
  plain(HEBREW);
  say('3c. Click in the empty paragraph below (between the two lines of dashes), New Comment, type "M3c".');
  plain('- - -');
  plain('');
  plain('- - -');

  heading('4. Optional: No Markup over a deleted paragraph mark');
  say('Turn Track Changes ON. Click at the very end of the heading "A heading joined to the next paragraph" and press Delete, so that its paragraph mark is deleted. Then Review > Display for Review > No Markup, take a screenshot of the two lines, and go back to All Markup.');
  const joined = body.insertParagraph('A heading joined to the next paragraph', 'End');
  joined.styleBuiltIn = 'Heading2';
  plain('The body paragraph after the heading.');

  heading('5. Optional: Original over a formatting change');
  say('With Track Changes ON, make the word "bold" below bold. Then Review > Display for Review > Original, take a screenshot of the line, and go back to All Markup.');
  plain('This word becomes bold with tracking on.');

  // The right-to-left paragraphs: set by the editor, whose export writes `w:bidi` as it writes it for a person.
  const session = await EditorDocument.open(new Uint8Array(await pkg.save()));
  const hebrew = [];
  session.doc.forEach((node, pos) => { if (node.type.name === 'paragraph' && node.textContent === HEBREW) hebrew.push(pos + 1); });
  if (hebrew.length !== 2) throw new Error(`expected two Hebrew paragraphs, found ${hebrew.length}`);
  for (const from of hebrew) {
    const result = await runCommandAsync(session, 'paragraph.bidi', { value: true }, { from });
    if (!result.ok) throw new Error('paragraph.bidi did not apply');
  }
  const saved = await session.save();
  if (saved.selfCheck && !saved.selfCheck.equal) throw new Error('the self-check failed');
  return saved.bytes;
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const here = dirname(fileURLToPath(import.meta.url));
  const out = process.argv[2] ?? join(here, '..', '..', '..', '..', 'editor-mcp', 'test-results', 'word-session');
  await mkdir(out, { recursive: true });
  await writeFile(join(out, 'measure-base.docx'), await buildMeasureBase());
  console.log(`wrote ${join(out, 'measure-base.docx')}`);
}
