// SPDX-License-Identifier: GPL-3.0-or-later
//
// Builds `paste/paste-samples.docx` with the engine (ED-003 section 8.3, step F2): the document
// Jason opens in Word to capture the clipboard HTML the paste corpus needs. Each numbered part
// is one capture; the whole document is the last. Run from packages/editor-model:
//
//   node test/fixtures/build/paste-samples.mjs
//
// and commit the output; `paste/README.md` says how to capture. The document is built with the
// engine so it is ours to redistribute; what Word puts on the clipboard for it is the fixture.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { p, r } from '@docx4j/generated-objects-ts/builders/wml';

const pkg = await WordprocessingMLPackage.createPackage();
const body = pkg.body;

const heading = (text, level) => { const p = body.insertParagraph(text, 'End'); p.styleBuiltIn = `Heading${level}`; return p; };
const plain = (text) => body.insertParagraph(text, 'End');

// 0. How to use the document (not a capture).
const title = body.insertParagraph('Paste samples for the docx4j editor', 'End');
title.styleBuiltIn = 'Title';
plain('Each numbered section below is one capture: select from its heading to the line before the next heading, copy, and save the clipboard HTML as paste/README.md says. The last capture is the whole document.');

// 1. Headings and body text.
heading('1. Headings and body text', 1);
heading('A second-level heading', 2);
plain('Body text under the second-level heading, one plain sentence.');
heading('A third-level heading', 3);
plain('Body text under the third-level heading, with a second sentence to make a longer paragraph that Word will wrap on its own.');

// 2. Inline formatting in one paragraph, each mark on its own run (an `insertText` at the end of
// a paragraph inherits the last run's formatting, so the runs are built with the builders).
heading('2. Inline formatting', 1);
body.insertElement(p([
  r('This paragraph has '), r('bold', { bold: true }), r(', '), r('italic', { italic: true }), r(', '),
  r('underlined', { underline: true }), r(', '), r('struck through', { strikeThrough: true }), r(', a super'),
  r('script', { superscript: true }), r(' and a sub'), r('script', { subscript: true }), r(' run, then '),
  r('bold italic together', { bold: true, italic: true }), r('.'),
]), 'End');
const breaks = plain('A line break follows this text');
breaks.insertBreak('Line', 'End');
breaks.insertText('and this is the same paragraph after it, then a tab', 'End');
breaks.insertText('\t', 'End');
breaks.insertText('and text after the tab.', 'End');

// 3. A bulleted list, two levels deep.
heading('3. A bulleted list', 1);
const b1 = plain('First bullet');
const bullets = await b1.startNewList({ bullet: true });
plain('Second bullet').attachToList(bullets.id, 0);
plain('A nested bullet under the second').attachToList(bullets.id, 1);
plain('Another nested bullet').attachToList(bullets.id, 1);
plain('Third bullet, back at the first level').attachToList(bullets.id, 0);
plain('A plain paragraph after the list.');

// 4. A numbered list, two levels deep.
heading('4. A numbered list', 1);
const n1 = plain('First item');
const numbers = await n1.startNewList();
plain('Second item').attachToList(numbers.id, 0);
plain('A nested item under the second').attachToList(numbers.id, 1);
plain('Another nested item').attachToList(numbers.id, 1);
plain('Third item, back at the first level').attachToList(numbers.id, 0);
plain('A plain paragraph after the list.');

// 5. A simple table.
heading('5. A simple table', 1);
plain('A three-by-three table with a header row:');
const table = body.insertTable(3, 3, 'End', [['Name', 'Quantity', 'Price'], ['Widget', '2', '3.50'], ['Gadget', '10', '12.00']]);
table.styleBuiltIn = 'TableGrid';
table.headerRowCount = 1;
plain('A plain paragraph after the table.');

// 6. Everything the converter drops, so the status line's report can be checked.
heading('6. What the editor drops on paste', 1);
const dropped = plain('This paragraph is centred, indented and in 14-point red Arial; the editor keeps the text and drops the rest, and says so.');
dropped.alignment = 'Centered';
dropped.leftIndent = 36;
dropped.font.size = 14;
dropped.font.color = '#FF0000';
dropped.font.name = 'Arial';
const hl = plain('This one has a ');
hl.insertText('highlighted', 'End').font.highlightColor = 'yellow';
hl.insertText(' word.', 'End');

await writeFile(join(dirname(fileURLToPath(import.meta.url)), '..', 'paste', 'paste-samples.docx'), await pkg.save());
console.log('written paste/paste-samples.docx');
