// SPDX-License-Identifier: GPL-3.0-or-later
//
// Builds the two documents Jason opens in Word 2010 on the reference machine and resaves with a
// font embedded (ED-003 section 8.4, step F3): the same content twice, one to be saved with the
// entire font, one with the characters used only. Noto Sans (OFL) is the font, so the fixtures
// stay redistributable with the font inside them; it must be installed on the machine that
// embeds. Run from packages/editor-model:
//
//   node test/fixtures/build/fonts-embedded.mjs [outputDirectory]
//
// The output goes to the directory given (Jason's Downloads by default), not to the fixtures:
// what Word writes back is the fixture (`fonts-embedded.docx`, `fonts-embedded-subset.docx`).
//
// Word subsets only when a font is used for fewer than 32 characters; past that it embeds the
// entire font whatever the option says (Microsoft's `Document.SaveSubsetFonts` reference), which
// is why the first subset attempt, over the pangrams, came back whole. The subset variant
// therefore sets short words in Noto Sans (13 distinct characters over the four faces, `x`
// among the letters it lacks) and keeps the pangrams for the entire-font variant.
import { writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { homedir } from 'node:os';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { p, r } from '@docx4j/generated-objects-ts/builders/wml';

const out = process.argv[2] ?? join(homedir(), 'Downloads');
const FONT = 'Noto Sans';

async function build(variant) {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const title = body.insertParagraph(`Embedded font sample: ${variant}`, 'End');
  title.styleBuiltIn = 'Heading1';
  body.insertParagraph(`This paragraph is in the document's default font, for comparison. Save this document from Word with File > Options > Save > "Embed fonts in the file"${variant === 'entire font' ? ', with "Embed only the characters used" off' : ', with "Embed only the characters used" on'}.`, 'End');
  if (variant === 'entire font') {
    // Noto Sans on whole paragraphs, in the four faces Word embeds separately (w:embedRegular, w:embedBold, w:embedItalic, w:embedBoldItalic).
    body.insertElement(p([r(`This paragraph is set in ${FONT}, regular: the quick brown fox jumps over the lazy dog. 0123456789`, { name: FONT })]), 'End');
    body.insertElement(p([r(`This paragraph is set in ${FONT}, bold: sphinx of black quartz, judge my vow.`, { name: FONT, bold: true })]), 'End');
    body.insertElement(p([r(`This paragraph is set in ${FONT}, italic: pack my box with five dozen liquor jugs.`, { name: FONT, italic: true })]), 'End');
    body.insertElement(p([r(`This paragraph is set in ${FONT}, bold italic: how vexingly quick daft zebras jump.`, { name: FONT, bold: true, italic: true })]), 'End');
    body.insertElement(p([r('abc', { name: FONT }), r(' (three letters in the embedded font)')]), 'End');
  } else {
    // Under 32 distinct characters of Noto Sans in the whole document, so Word makes a subset: the
    // face's name in the face, the rest of each line in the default font. The subset lacks `x`,
    // among others, so typing one into a Noto Sans run later shows the fallback.
    body.insertElement(p([r('regular', { name: FONT }), r(` is set in ${FONT}, regular; the rest of the line is the default font.`)]), 'End');
    body.insertElement(p([r('bold', { name: FONT, bold: true }), r(` is set in ${FONT}, bold.`)]), 'End');
    body.insertElement(p([r('italic', { name: FONT, italic: true }), r(` is set in ${FONT}, italic.`)]), 'End');
    body.insertElement(p([r('bold italic', { name: FONT, bold: true, italic: true }), r(` is set in ${FONT}, bold italic.`)]), 'End');
    body.insertElement(p([r('abc', { name: FONT }), r(' (three letters in the embedded font; a glyph the subset lacks, such as x, falls back)')]), 'End');
  }
  body.insertParagraph('A closing paragraph in the default font.', 'End');
  return pkg.save();
}

await writeFile(join(out, 'fonts-embedded.docx'), await build('entire font'));
await writeFile(join(out, 'fonts-embedded-subset.docx'), await build('characters used only'));
console.log(`written fonts-embedded.docx and fonts-embedded-subset.docx to ${out}`);
