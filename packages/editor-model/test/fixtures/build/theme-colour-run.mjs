// SPDX-License-Identifier: GPL-3.0-or-later
//
// Two fixtures on the package of `odd-even-first-page.docx` (a Word file with three headers and the Office theme,
// accent1 `4F81BD`), with its body and the first paragraph of each header replaced. Run from packages/editor-model:
//
//   node test/fixtures/build/theme-colour-run.mjs
//
// `theme-colour-run.docx` (ED-008 section 8.8, a run's own theme colour): runs whose `w:val` disagrees with the theme
// reference beside it, in the places the view reaches a run: a paragraph, a tint, a tracked insertion and deletion, a
// table cell, an inline control, a header; and one run whose value agrees.
//
// `normal-form-values.docx` (ED-008 section 8.8, the normalisation report): values a save writes as read and normal
// form would change: a lower-case colour, a named colour, a highlight outside Word's list, one in a header.
import { readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '@docx4j/core-ts';

const here = dirname(fileURLToPath(import.meta.url));
const src = join(here, '..', 'odd-even-first-page.docx');

const STALE = '<w:color w:val="FF00FF" w:themeColor="accent1"/>';
const t = (s) => `<w:t xml:space="preserve">${s}</w:t>`;
const r = (text, rPr) => `<w:r><w:rPr>${rPr}</w:rPr>${t(text)}</w:r>`;
const p = (inner) => `<w:p>${inner}</w:p>`;
const revision = 'w:author="Reviewer" w:date="2026-11-05T09:00:00Z"';

const themed = [
  p(r('stale run', STALE)),
  p(r('agreed run', '<w:color w:val="4F81BD" w:themeColor="accent1"/>')),
  p(r('tinted run', '<w:color w:val="FF00FF" w:themeColor="accent1" w:themeTint="99"/>')),
  p(`<w:ins w:id="901" ${revision}>${r('inserted run', STALE)}</w:ins>`),
  p(`<w:del w:id="902" ${revision}><w:r><w:rPr>${STALE}</w:rPr><w:delText xml:space="preserve">deleted run</w:delText></w:r></w:del>`),
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(r('cell run', STALE))}</w:tc></w:tr></w:tbl>`,
  p(`<w:sdt><w:sdtPr><w:id w:val="903"/></w:sdtPr><w:sdtContent>${r('control run', STALE)}</w:sdtContent></w:sdt>`),
  p(r('last run', '')),
].join('');

const values = [
  p(r('lower case colour', '<w:color w:val="ff0000"/>')),
  p(r('named colour', '<w:color w:val="red"/>')),
  p(r('unknown highlight', '<w:highlight w:val="purple"/>')),
  p(r('removed in the test', '<w:color w:val="00ff00"/>')),
  p(r('last run', '')),
].join('');

async function build(name, body, headerRPr) {
  const pkg = await WordprocessingMLPackage.load(new Uint8Array(await readFile(src)));
  pkg.normaliseOnSave = false;
  const main = pkg.getMainDocumentPart();
  main.setXml((await main.getXml()).replace(/<w:body>.*?(<w:sectPr)/s, `<w:body>${body}$1`));
  for (const n of [1, 2, 3]) {
    const part = pkg.getPart(`/word/header${n}.xml`);
    part.setXml((await part.getXml()).replace(/<w:r( [^>]*)?>(<w:t[ >])/, `<w:r$1><w:rPr>${headerRPr}</w:rPr>$2`));
  }
  const out = join(here, '..', name);
  await writeFile(out, new Uint8Array(await pkg.save()));
  console.log('wrote', out);
}

await build('theme-colour-run.docx', themed, STALE);
await build('normal-form-values.docx', values, '<w:color w:val="0000ff"/>');
