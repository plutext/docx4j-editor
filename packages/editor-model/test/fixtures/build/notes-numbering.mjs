// SPDX-License-Identifier: GPL-3.0-or-later
//
// Builds `notes-numbering.docx` with the engine (ED-003 section 6.6 item 1): footnotes and
// endnotes over two sections, `w:footnotePr` and `w:endnotePr` with a format, a start and a
// per-section restart, `w:pgNumType` with a start and a format, and one `w:lastRenderedPageBreak`
// in the second section. Run from packages/editor-model:
//
//   node test/fixtures/build/notes-numbering.mjs
//
// and commit the output; the suites read the committed document. Expected display: footnotes
// i, ii in section 1 and i, ii, iii in section 2 (lower roman, restarting each section);
// endnotes C, D (upper letter from 3, continuous); page numbers 5, then i and ii.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage, FootnotesPart, EndnotesPart } from '@docx4j/core-ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const fnRef = (id) => `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
const enRef = (id) => `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;
const t = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>
<w:p>${t('First section, first note')}${fnRef(1)}${t(' and second note')}${fnRef(2)}${t('.')}</w:p>
<w:p>${t('An endnote')}${enRef(1)}${t(' in the first section.')}</w:p>
<w:p><w:pPr><w:sectPr><w:footnotePr><w:numRestart w:val="eachSect"/></w:footnotePr><w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:pgNumType w:start="5"/></w:sectPr></w:pPr>${t('Last paragraph of the first section.')}</w:p>
<w:p>${t('Second section, third note')}${fnRef(3)}${t(' and fourth')}${fnRef(4)}${t('.')}</w:p>
<w:p><w:r><w:lastRenderedPageBreak/><w:t xml:space="preserve">Page two of the second section, fifth note</w:t></w:r>${fnRef(5)}${t('.')}</w:p>
<w:p>${t('A second endnote')}${enRef(2)}${t('.')}</w:p>
<w:sectPr><w:footnotePr><w:numRestart w:val="eachSect"/></w:footnotePr><w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:pgNumType w:fmt="lowerRoman" w:start="1"/></w:sectPr>
</w:body></w:document>`;

const sepNotes = (tag) => `<w:${tag} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:${tag}><w:${tag} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:${tag}>`;
const note = (tag, id, text) => `<w:${tag} w:id="${id}"><w:p><w:pPr><w:pStyle w:val="${tag === 'footnote' ? 'FootnoteText' : 'EndnoteText'}"/></w:pPr><w:r><w:rPr><w:rStyle w:val="${tag === 'footnote' ? 'FootnoteReference' : 'EndnoteReference'}"/></w:rPr><w:${tag}Ref/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:${tag}>`;

const footnotes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="${W}" xmlns:r="${R}">${sepNotes('footnote')}${[1, 2, 3, 4, 5].map((i) => note('footnote', i, `Footnote ${i}.`)).join('')}</w:footnotes>`;
const endnotes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="${W}" xmlns:r="${R}">${sepNotes('endnote')}${[1, 2].map((i) => note('endnote', i, `Endnote ${i}.`)).join('')}</w:endnotes>`;

const settings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${W}"><w:defaultTabStop w:val="720"/><w:footnotePr><w:numFmt w:val="lowerRoman"/><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:endnotePr><w:numFmt w:val="upperLetter"/><w:numStart w:val="3"/><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;

const pkg = await WordprocessingMLPackage.createPackage();
const main = pkg.getMainDocumentPart();
main.setXml(document);
main.documentSettingsPart.setXml(settings);
const fn = new FootnotesPart();
fn.setXml(footnotes);
main.addTargetPart(fn);
const en = new EndnotesPart();
en.setXml(endnotes);
main.addTargetPart(en);
const out = join(dirname(fileURLToPath(import.meta.url)), '..', 'notes-numbering.docx');
await writeFile(out, await pkg.save());
console.log(`wrote ${out}`);
