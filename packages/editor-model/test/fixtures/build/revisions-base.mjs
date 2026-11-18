// SPDX-License-Identifier: GPL-3.0-or-later
//
// Builds `revisions/revisions-base.docx` with the engine (ED-005 section 13, for E4.c step C1): the
// document Jason opens in Word 2010 and in Word 15 on the reference machine and revises there with
// tracking on, as two authors, following the instructions it carries. Word makes every revision,
// since revision markup a test wrote would only test the writer against itself; this file holds
// none. Each numbered section is one kind of revision of ED-005 section 5.2 (and two of its
// "preserved" list): its instructions in blue italics, then the text to revise. Run from
// packages/editor-model:
//
//   node test/fixtures/build/revisions-base.mjs
//
// and commit the output; `revisions/README.md` says what to do in Word and where the saves go.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { p, r } from '@docx4j/generated-objects-ts/builders/wml';

const pkg = await WordprocessingMLPackage.createPackage();
const body = pkg.body;

const heading = (text, level = 1) => { const h = body.insertParagraph(text, 'End'); h.styleBuiltIn = `Heading${level}`; return h; };
const plain = (text) => body.insertParagraph(text, 'End');
/** An instruction: blue italics, never revised. */
const say = (text) => { const i = body.insertParagraph(text, 'End'); i.font.italic = true; i.font.color = '#1F4E79'; return i; };

const title = body.insertParagraph('Revision fixtures for the docx4j editor', 'End');
title.styleBuiltIn = 'Title';
say('This document was built by docx4j (core-ts) for ED-005 section 13 and holds no revisions: Word makes them all. Do it twice, from this file each time: once in Word 2010, saved as revisions-word2010.docx, and once in Word 15, saved as revisions-word15.docx.');
say('Before you start: File > Options > General, User name "Author A", Initials "AA" (in Word 15 also tick "Always use these values regardless of sign in to Office"). In Word 15, if the title bar says [Compatibility Mode], use File > Info > Convert first. Then Review > Track Changes on, and in the change tracking options keep "Track moves" and "Track formatting" ticked.');
say('Then do every step marked [A] in order; set the user name to "Author B" (initials "BB") and do every step marked [B]; set it back to "Author A" and do the steps marked [A again]. Revise only the black text, never a heading or these blue instructions. Accept and reject nothing. Save as a Word Document (.docx).');

heading('1. Insertions and deletions by two authors');
say('[A] Type the word "very " before "lazy", and delete the word "old " with the space after it.');
say('[B] Type the word "red " before "fox", and delete the words " at dawn" (keep the full stop).');
plain('The quick brown fox jumps over the lazy dog near the old stone bridge at dawn.');

heading('2. An insertion deleted by another author');
say('[A] Type the word "green " before "field".');
say('[B] Select the word "green " that Author A typed and delete it.');
plain('Cows graze in the field beside the river.');

heading('3. A move');
say('[A] Triple-click the second paragraph below (it is selected with its paragraph mark), cut it (Ctrl+X), click at the start of the first paragraph below and paste (Ctrl+V). Word marks a move with a double strikethrough where the text was and a double underline where it went, in green. If it shows as an ordinary deletion and insertion, undo, check "Track moves", and try again, or drag the selected paragraph instead.');
plain('The first paragraph of the move: it stays where it is.');
plain('The second paragraph of the move: it goes before the first.');
plain('The third paragraph of the move: it stays where it is.');

heading('4. Run formatting changed');
say('[A] Make the word "alpha" bold, the word "bravo" red (Font Color, Red), and the word "charlie" italic and underlined.');
plain('The word alpha becomes bold, the word bravo becomes red, and the word charlie becomes italic and underlined.');

heading('5. Paragraph properties changed');
say('[A] Centre the first paragraph below (Ctrl+E), indent the second by one step (Increase Indent), and give the third the style Heading 2.');
plain('Centre this paragraph.');
plain('Indent this paragraph by one step.');
plain('Give this paragraph the style Heading 2.');

heading('6. Paragraph marks inserted and deleted');
say('[A] Click right after the colon in the first paragraph below and press Enter. Then click at the very end of the second paragraph and press Delete, joining the third to it.');
plain('Split this paragraph after the colon: the rest becomes a paragraph of its own.');
plain('Join this paragraph with the next one.');
plain('This is the paragraph it joins.');

heading('7. Table rows inserted and deleted');
say('[A] Click in the row "2", then Table Tools, Layout, Insert Below, and type "new" in each cell of the new row. Then click in the row "3" and use Layout, Delete, Delete Rows.');
const rows = body.insertTable(4, 3, 'End', [['Row', 'Action', 'Note'], ['1', 'Keep', 'row one stays'], ['2', 'Keep', 'a row is inserted below this one'], ['3', 'Delete', 'this row is deleted']]);
rows.styleBuiltIn = 'TableGrid';
rows.headerRowCount = 1;
plain('The paragraph after the table of rows.');

heading('8. Table properties changed');
say('[A] Click in the table below and use Table Properties, Table, Alignment: Center. Then shade its top left cell yellow (Table Tools, Design, Shading).');
const props = body.insertTable(2, 2, 'End', [['Top left', 'Top right'], ['Bottom left', 'Bottom right']]);
props.styleBuiltIn = 'TableGrid';
plain('The paragraph after the table of properties.');

heading('9. Numbering changed');
say('[A] Select the three numbered items below and click Bullets, so the list becomes a bulleted one.');
const first = plain('The first numbered item.');
const list = await first.startNewList();
plain('The second numbered item.').attachToList(list.id, 0);
plain('The third numbered item.').attachToList(list.id, 0);
plain('The paragraph after the list.');

heading('10. Inline items inserted and deleted');
say('[A] Delete the tab between "Left" and "right". Put a line break (Shift+Enter) after the word "break". Insert a footnote at the end of the last sentence (References, Insert Footnote) and type "A tracked footnote." in it.');
// The builders' `r` makes a `w:tab` of the `\t` (the content API's `insertText` would put the character in the `w:t`).
body.insertElement(p([r('Left\tright: the tab between them is deleted. A line break goes after the word break here. The footnote goes at the end of this sentence.')]), 'End');

heading('11. Comments (Word 15; in Word 2010 only the two comments)');
say('[A] Select the first sentence below, Review, New Comment, and type "A: the first comment."');
say('[B] Reply to Author A\'s comment (the Reply button in the comment, Word 15 only) with "B: the reply." Then select the second sentence, New Comment, and type "B: a comment of B\'s own."');
say('[A again] Right-click Author B\'s reply and choose Mark Comment Done (Word 15 only).');
body.insertElement(p([r('Author A comments on this sentence, and Author B answers. '), r('Author B comments on this sentence alone.')]), 'End');

heading('12. Section properties changed');
say('[A] Page Layout (Layout in Word 15), Margins, Narrow.');
plain('The last paragraph of the document.');

const out = join(dirname(fileURLToPath(import.meta.url)), '..', 'revisions', 'revisions-base.docx');
await writeFile(out, await pkg.save());
console.log(`written ${out}`);
