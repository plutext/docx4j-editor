// SPDX-License-Identifier: GPL-3.0-or-later
//
// Builds `print-pages.docx` with the engine (ED-006 section 5): what printing must honour, in four
// sections. Run from packages/editor-model:
//
//   node test/fixtures/build/print-pages.mjs
//
// and commit the output; the suites read the committed document.
//
// 1. A4 portrait: a heading (its style keeps with next), a page break inside a paragraph, a paragraph
//    with its own "page break before", one that keeps with the next, a heading turning its style's
//    "keep with next" off, one with widow control off, a footnote and an endnote, a tracked insertion
//    and deletion, a content control showing its placeholder, and a long table with two header rows
//    and a row that cannot split.
// 2. A4 landscape, after a next-page break.
// 3. The same landscape page, after a continuous break: printed on the same page.
// 4. US Letter with a gutter, after a next-page break. The settings part asks for mirror margins.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage, FootnotesPart, EndnotesPart } from '@docx4j/core-ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const t = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (text, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${t(text)}</w:p>`;
const fnRef = (id) => `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
const enRef = (id) => `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;
const date = '2026-10-06T00:00:00Z';

const a4 = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>';
const a4Landscape = '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>';
const letter = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="708" w:footer="708" w:gutter="720"/>';

const cell = (text, width = 4513) => `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr><w:p>${t(text)}</w:p></w:tc>`;
const row = (cells, trPr = '') => `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells.join('')}</w:tr>`;
const table = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid>
${row([cell('Header one'), cell('Header two')], '<w:tblHeader/>')}
${row([cell('Second header row'), cell('also repeated')], '<w:tblHeader/>')}
${Array.from({ length: 40 }, (_, i) => row([cell(`Row ${i + 1}`), cell(`Value ${i + 1}`)])).join('\n')}
${row([`<w:tc><w:tcPr><w:tcW w:w="4513" w:type="dxa"/></w:tcPr>${Array.from({ length: 6 }, (_, i) => `<w:p>${t(`Unsplit line ${i + 1}`)}</w:p>`).join('')}</w:tc>`, cell('A row that cannot split')], '<w:cantSplit/>')}
</w:tbl>`;

const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>
${p('Printing samples', '<w:pStyle w:val="Title"/>')}
${p('Section one: A4 portrait', '<w:pStyle w:val="Heading1"/>')}
<w:p>${t('Before the page break.')}<w:r><w:br w:type="page"/></w:r>${t('After the page break.')}</w:p>
${p('This paragraph starts a page of its own.', '<w:pageBreakBefore/>')}
${p('This paragraph keeps with the next.', '<w:keepNext/>')}
${p('The next paragraph.')}
${p('A heading that does not keep with the next', '<w:pStyle w:val="Heading1"/><w:keepNext w:val="0"/>')}
${p('A paragraph with widow control off.', '<w:widowControl w:val="0"/>')}
<w:p>${t('A footnote')}${fnRef(1)}${t(' and an endnote')}${enRef(1)}${t('.')}</w:p>
<w:p>${t('A tracked change: ')}<w:ins w:id="101" w:author="Reviewer" w:date="${date}">${t('inserted words')}</w:ins>${t(' and ')}<w:del w:id="102" w:author="Reviewer" w:date="${date}"><w:r><w:delText xml:space="preserve">deleted words</w:delText></w:r></w:del>${t('.')}</w:p>
<w:p>${t('A content control: ')}<w:sdt><w:sdtPr><w:alias w:val="Name"/><w:tag w:val="name"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Click or tap here to enter text.</w:t></w:r></w:sdtContent></w:sdt></w:p>
${p('A table with two header rows and a row that cannot split:')}
${table}
<w:p><w:pPr><w:sectPr>${a4}</w:sectPr></w:pPr>${t('The last paragraph of section one.')}</w:p>
${p('Section two: A4 landscape', '<w:pStyle w:val="Heading1"/>')}
<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/>${a4Landscape}</w:sectPr></w:pPr>${t('The last paragraph of section two.')}</w:p>
<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/>${a4Landscape}</w:sectPr></w:pPr>${t('Section three follows on the same landscape page: its break is continuous.')}</w:p>
${p('Section four: US Letter with a gutter', '<w:pStyle w:val="Heading1"/>')}
${p('The last paragraph of the document.')}
<w:sectPr><w:type w:val="nextPage"/>${letter}</w:sectPr>
</w:body></w:document>`;

const sepNotes = (tag) => `<w:${tag} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:${tag}><w:${tag} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:${tag}>`;
const note = (tag, id, text) => `<w:${tag} w:id="${id}"><w:p><w:pPr><w:pStyle w:val="${tag === 'footnote' ? 'FootnoteText' : 'EndnoteText'}"/></w:pPr><w:r><w:rPr><w:rStyle w:val="${tag === 'footnote' ? 'FootnoteReference' : 'EndnoteReference'}"/></w:rPr><w:${tag}Ref/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:${tag}>`;
const footnotes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="${W}" xmlns:r="${R}">${sepNotes('footnote')}${note('footnote', 1, 'The footnote, printed after the body.')}</w:footnotes>`;
const endnotes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="${W}" xmlns:r="${R}">${sepNotes('endnote')}${note('endnote', 1, 'The endnote, printed after the footnotes.')}</w:endnotes>`;
const settings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${W}"><w:mirrorMargins/><w:defaultTabStop w:val="720"/><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;

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
// The styles the document names that a created package lacks (the defaults carry Heading1 and Title).
await pkg.styles.ensure(['FootnoteText', 'FootnoteReference', 'EndnoteText', 'EndnoteReference', 'PlaceholderText']);
const out = join(dirname(fileURLToPath(import.meta.url)), '..', 'print-pages.docx');
await writeFile(out, await pkg.save());
console.log(`wrote ${out}`);
