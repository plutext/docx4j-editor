// SPDX-License-Identifier: GPL-3.0-or-later
//
// The files for ED-010 section 4's Word 365 session: what Word's own copy and paste does with a
// comment, another author's revisions, a bookmark, a picture, a footnote, a numbered item and a
// heading, and whether Word opens a file holding the duplicates the editor's paste-by-identity
// would write (two drawings sharing a `wp:docPr` id, two bookmarks sharing a name and id). Written
// as WordprocessingML by hand, so every element is exactly what the question needs; zipped with
// fflate (core-ts's dependency, hoisted). Run from packages/editor-model:
//
//   node test/fixtures/build/clipboard-probe.mjs [out-dir]
//
// (the default output is test-results/clipboard/, which git ignores; the README there says what to
// do in Word, and goes to the share with the files)
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { zipSync, strToU8 } from 'fflate';

const out = process.argv[2] ?? join('test-results', 'clipboard');
await mkdir(out, { recursive: true });

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const PIC = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const DATE = '2026-10-22T00:00:00Z';

// a 1 by 1 pixel PNG, shown at one inch
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==', 'base64');

const t = (text) => `<w:t xml:space="preserve">${text}</w:t>`;
const r = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${t(text)}</w:r>`;
const p = (inner, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${inner}</w:p>`;

/** An inline picture, its `wp:docPr` id and name as given. */
const picture = (id, name, rId = 'rIdImage') => `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
  `<wp:extent cx="914400" cy="914400"/><wp:docPr id="${id}" name="${name}"/>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
  `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const body = (paragraphs) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<w:document ${W} ${R} ${WP} ${A} ${PIC}><w:body>${paragraphs.join('')}` +
  `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>` +
  `</w:body></w:document>`;

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles ${W}>` +
  `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>` +
  `<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
  `<w:pPr><w:keepNext/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="28"/></w:rPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>` +
  `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style>` +
  `<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>` +
  `<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:rPr><w:sz w:val="16"/></w:rPr></w:style>` +
  `</w:styles>`;

const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering ${W}>` +
  `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>` +
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>` +
  `</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

// the footnote separators are declared only where there is a footnotes part: declared without one, Word 365
// reports the file unreadable and repairs it (Jason's first run of the two duplicate files, 2026-10-22)
const settings = (footnotes) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings ${W}>` +
  (footnotes ? `<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>` : '') +
  `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;

const footnotesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:footnotes ${W}>` +
  `<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>` +
  `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
  `<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>${r(' The footnote of case 7.')}</w:p></w:footnote>` +
  `</w:footnotes>`;

const comment = (id, text) => `<w:comment w:id="${id}" w:author="Author B" w:date="${DATE}" w:initials="B"><w:p><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>${r(text)}</w:p></w:comment>`;
const commentsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments ${W}>${comment(0, 'Comment on case 1.')}${comment(1, 'Comment on case 2.')}</w:comments>`;

const commented = (id, text) => `<w:commentRangeStart w:id="${id}"/>${r(text)}<w:commentRangeEnd w:id="${id}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;

const contentTypes = (extra) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
  `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>` +
  extra + `</Types>`;
const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
const rel = (id, type, target) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
const docRels = (extra) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  rel('rIdStyles', 'styles', 'styles.xml') + rel('rIdSettings', 'settings', 'settings.xml') + extra + `</Relationships>`;

/** Writes a package from its parts (strings, or bytes for the image). */
const save = async (name, parts) => {
  const files = {};
  for (const [path, content] of Object.entries(parts)) files[path] = typeof content === 'string' ? strToU8(content) : content;
  await writeFile(join(out, name), zipSync(files, { level: 6 }));
  console.log(join(out, name));
};

// ---------------------------------------------------------------- the probe

const target = (n, note) => p(r(`[T${n}: paste case ${n} here${note ? `, ${note}` : ''}]`));
const probe = [
  p(r('clipboard probe (ED-010, 2026-10-22). Sources first, targets below. Copy each source as the README says, paste at its target.'), ''),
  p(r('Case 1: ') + commented(0, 'COMMENTED WHOLE') + r(' end.')),
  p(r('Case 2: ') + commented(1, 'COMMENTED PART ONE') + r(' end. (copy the word PART alone)')),
  p(r('Case 3 and 4: before ') + `<w:ins w:id="10" w:author="Author B" w:date="${DATE}">${r('INSERTED ')}</w:ins>` +
    `<w:del w:id="11" w:author="Author B" w:date="${DATE}"><w:r><w:delText xml:space="preserve">DELETED </w:delText></w:r></w:del>` + r('after. (copy the whole paragraph, its mark included)')),
  p(r('Case 5: ') + `<w:bookmarkStart w:id="0" w:name="bm1"/>${r('BOOKMARKED')}<w:bookmarkEnd w:id="0"/>` + r(' end. (copy the whole paragraph)')),
  p(r('Case 6: ') + picture(1, 'Picture 1') + r(' (copy the picture alone)')),
  p(r('Case 7: a footnote reference') + `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>` + r(' here. (copy the whole paragraph)')),
  p(r('Case 8: NUMBERED ITEM (copy the whole paragraph, its mark included)'), `<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>`),
  p(r('Case 9: HEADING TWO TEXT (copy the words alone, not the mark)'), `<w:pStyle w:val="Heading2"/>`),
  p(r('--- targets ---')),
  target(1), target(2), target(3, 'tracking off'), target(4, 'tracking ON as Author A, in the second run'),
  target(5), target(6), target(7), target(8), target(9, 'a plain paragraph: paste the heading words into it'),
  p(r('end.')),
];
await save('clipboard-probe.docx', {
  '[Content_Types].xml': contentTypes(
    `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>` +
    `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>` +
    `<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>`),
  '_rels/.rels': rootRels,
  'word/document.xml': body(probe),
  'word/_rels/document.xml.rels': docRels(rel('rIdNumbering', 'numbering', 'numbering.xml') + rel('rIdComments', 'comments', 'comments.xml') + rel('rIdFootnotes', 'footnotes', 'footnotes.xml') + rel('rIdImage', 'image', 'media/image1.png')),
  'word/styles.xml': stylesXml,
  'word/settings.xml': settings(true),
  'word/numbering.xml': numberingXml,
  'word/comments.xml': commentsXml,
  'word/footnotes.xml': footnotesXml,
  'word/media/image1.png': new Uint8Array(PNG),
});

// ---------------------------------------------------------------- the duplicates

await save('dup-docpr-2.docx', {
  '[Content_Types].xml': contentTypes(''),
  '_rels/.rels': rootRels,
  'word/document.xml': body([
    p(r('dup-docpr: two inline pictures below share wp:docPr id="1" and name="Picture 1", as a paste by identity would write them.')),
    p(r('First: ') + picture(1, 'Picture 1')),
    p(r('Second: ') + picture(1, 'Picture 1')),
    p(r('end.')),
  ]),
  'word/_rels/document.xml.rels': docRels(rel('rIdImage', 'image', 'media/image1.png')),
  'word/styles.xml': stylesXml,
  'word/settings.xml': settings(false),
  'word/media/image1.png': new Uint8Array(PNG),
});

await save('dup-bookmark-2.docx', {
  '[Content_Types].xml': contentTypes(''),
  '_rels/.rels': rootRels,
  'word/document.xml': body([
    p(r('dup-bookmark: the two paragraphs below each hold a bookmark named bm1 with w:id="0".')),
    p(r('First: ') + `<w:bookmarkStart w:id="0" w:name="bm1"/>${r('BOOKMARKED')}<w:bookmarkEnd w:id="0"/>` + r(' end.')),
    p(r('Second: ') + `<w:bookmarkStart w:id="0" w:name="bm1"/>${r('BOOKMARKED')}<w:bookmarkEnd w:id="0"/>` + r(' end.')),
    p(r('end.')),
  ]),
  'word/_rels/document.xml.rels': docRels(''),
  'word/styles.xml': stylesXml,
  'word/settings.xml': settings(false),
});

await writeFile(join(out, 'README.txt'), `clipboard probes (docx4j-ts-editor ED-010 section 4, 2026-10-22), for Word 365

Question: what Word's own copy and paste does with a comment, another author's revisions, a
bookmark, a picture, a footnote, a numbered item and a heading; and whether Word opens a file
holding the duplicates a paste by identity would write.

A. clipboard-probe.docx, run 1, tracking OFF
   Open it. Review > Track Changes off. For each case, select what the source paragraph says to
   copy, Ctrl+C, click into its target paragraph [Tn] (at the end of the bracketed text is fine),
   Ctrl+V, and take Word's default paste option (Keep Source Formatting). Skip case 4.
     1  the whole commented range "COMMENTED WHOLE" (select from C to E exactly)
     2  the word PART alone, from inside the commented range
     3  the whole paragraph "Case 3 and 4: ..." including its paragraph mark
     5  the whole paragraph "Case 5: ..." including its mark
     6  the picture alone
     7  the whole paragraph "Case 7: ..." including its mark
     8  the whole paragraph "Case 8: ..." including its mark
     9  the words "HEADING TWO TEXT" alone, without the mark
   Save As clipboard-probe-word365-off.docx (same folder).

B. clipboard-probe.docx, run 2, tracking ON
   Open the original again. File > Options > General: user name "Author A". Review > Track Changes
   on. Copy the whole paragraph "Case 3 and 4: ..." (its mark included), paste into [T4].
   Save As clipboard-probe-word365-on.docx.

C. dup-docpr-2.docx: open. Note: opens clean / unreadable content (repair) / other. If it opens,
   Save As dup-docpr-2-word365.docx without editing. (dup-docpr.docx, the first build, declared
   footnote separators without a footnotes part, which is what Word repaired; -2 has no such part.)

D. dup-bookmark-2.docx: the same; Save As dup-bookmark-2-word365.docx.

Notes welcome below (what the Paste Options showed, anything odd):
`);
