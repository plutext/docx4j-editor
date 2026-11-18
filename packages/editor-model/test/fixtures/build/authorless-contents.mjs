// SPDX-License-Identifier: GPL-3.0-or-later
//
// `authorless-contents.docx` (ED-005 section 12.113): the revisions ONLYOFFICE 9.4 leaves when it refreshes a Contents with
// tracking on, in its shape, with text of our own. Run from packages/editor-model:
//
//   node test/fixtures/build/authorless-contents.mjs
//
// Each Contents line is a hyperlink of `w:ins` elements with no author and no date, one per run, among them runs with nothing
// in them; its tab, its field characters and an empty run after the link; its paragraph mark `w:author=""`. Each heading keeps
// its text plain, with three `w:ins` around empty runs beside its bookmark. A named reviewer's real changes come after, one
// of them with an empty wrapper of no author inside it. Part of the shape: `w:author` is required by the schema and missing.
import { writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '@docx4j/core-ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const here = dirname(fileURLToPath(import.meta.url));

let id = 100;
const t = (s) => `<w:t xml:space="preserve">${s}</w:t>`;
/** A `w:ins` of no author and no date around a run's items, as ONLYOFFICE writes it. */
const ins = (items) => `<w:ins w:id="${id++}"><w:r>${items}</w:r></w:ins>`;
const wrapper = () => ins('');
const page = (n, mark) => `<w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve">PAGEREF ${mark} \\h</w:instrText><w:fldChar w:fldCharType="separate"/>${t(n)}<w:fldChar w:fldCharType="end"/>`;

const contentsLine = (number, title, pageNumber, mark) => `<w:p><w:pPr><w:rPr><w:ins w:id="${id++}" w:author=""/></w:rPr></w:pPr>`
  + `<w:hyperlink w:anchor="${mark}" w:history="1">${ins(t(number))}${ins('<w:tab/>')}${wrapper()}${ins(t(title))}${wrapper()}${ins('<w:tab/>')}${ins(page(pageNumber, mark))}</w:hyperlink>`
  + `${wrapper()}</w:p>`;

const heading = (title, mark, bookmark) => `<w:p>${wrapper()}<w:bookmarkStart w:id="${bookmark}" w:name="${mark}"/><w:r>${t(title)}</w:r>${wrapper()}<w:bookmarkEnd w:id="${bookmark}"/>${wrapper()}<w:r></w:r></w:p>`;

const named = 'w:author="Reviewer" w:date="2026-10-13T12:45:00Z"';
const body = [
  contentsLine('2', 'Membership generally', '4', '_Toc4'),
  contentsLine('3', 'Membership applications', '4', '_Toc5'),
  contentsLine('4', 'Register of members', '5', '_Toc6'),
  heading('Membership generally', '_Toc4', 4),
  `<w:p><w:r>${t('An individual is a member if ')}</w:r><w:ins w:id="${id++}" ${named}><w:r>${t('the committee agreed')}</w:r></w:ins>${wrapper()}<w:ins w:id="${id++}" ${named}><w:r>${t(' in writing')}</w:r></w:ins><w:r>${t('.')}</w:r></w:p>`,
  heading('Membership applications', '_Toc5', 5),
  `<w:p><w:r>${t('An application must be ')}</w:r><w:del w:id="${id++}" ${named}><w:r><w:delText xml:space="preserve">signed</w:delText></w:r></w:del><w:ins w:id="${id++}" ${named}><w:r>${t('made in writing')}</w:r></w:ins><w:r>${t('.')}</w:r></w:p>`,
  heading('Register of members', '_Toc6', 6),
  `<w:p><w:r>${t('The secretary keeps the register.')}</w:r></w:p>`,
].join('');

const pkg = await WordprocessingMLPackage.createPackage();
pkg.getMainDocumentPart().setXml(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
await writeFile(join(here, '..', 'authorless-contents.docx'), new Uint8Array(await pkg.save()));
