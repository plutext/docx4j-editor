// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-009's fixture: kept `mc:AlternateContent` elements whose branches hold run content, built on the
// package of Jason's probe `mc/mc-simple.docx` (Word 365, 2026-10-20; the share's `Office 2016/mc-fun/`)
// with its body replaced by the cases of ED-009 section 7. Run from packages/editor-model:
//
//   node test/fixtures/build/mc-branches.mjs
//
// writes test/fixtures/mc/mc-branches.docx. The cases, each under a plain paragraph naming it:
//   1 the run-level shape, text branches (mc-simple's own paragraph);
//   2 a tab and a line break in the Choice, other text in the Fallback;
//   3 a Choice requiring Excel's x14ac, with a Fallback: not understood by the engine before core-ts CR-010
//     section 15 (the Fallback selected), understood by Word 365 and by the engine since (the Choice);
//   4 the same with no Fallback (the first Choice either way);
//   5 text and an inline picture in the Choice (the package's rId17);
//   6 Jason's ten Choices, w14 to wp14, and a Fallback;
//   7 the same ten with w16du first (the engine selects the next Choice it understands);
//   8 the paragraph-level shape: two runs of different formatting in the Choice, one run in the Fallback.
//   9 the ten with w16sdtdh first, 10 the ten with w16sdtfl first (core-ts CR-010 section 15's gate: Word 365
//     understands both, Jason, 2026-10-20; the engine too since section 15);
//   11 a Choice requiring a private prefix (zz9), so the Fallback is selected; 12 the same with no Fallback (the first
//     Choice, mcBranchOf's rule): the unknown cases 3 and 4 had been, kept declared since objects 0.4.3 (jsonix 3.5.0).
import { readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { WordprocessingMLPackage } from '@docx4j/core-ts';

const here = dirname(fileURLToPath(import.meta.url));
const src = join(here, '..', 'mc', 'mc-simple.docx');
const out = process.argv[2] ?? join(here, '..', 'mc', 'mc-branches.docx');

const t = (s) => `<w:t xml:space="preserve">${s}</w:t>`;
const r = (inner, rPr = '') => `<w:r>${rPr}${inner}</w:r>`;
const p = (inner, pPr = '') => `<w:p>${pPr}${inner}</w:p>`;
const label = (s) => p(r(t(s)));
const ac = (choices, fallback) => `<mc:AlternateContent>${choices.map(([req, c]) => `<mc:Choice Requires="${req}">${c}</mc:Choice>`).join('')}${fallback === undefined ? '' : `<mc:Fallback>${fallback}</mc:Fallback>`}</mc:AlternateContent>`;
const TEN = ['w14', 'w15', 'w16se', 'w16cid', 'w16', 'w16cex', 'w16sdtdh', 'w16sdtfl', 'w16du', 'wp14'];
const picture = `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="901" name="Picture in a Choice"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="document_image_rId17.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId17"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;

const body = [
  p(r(t('ED-009 mc branches')), '<w:pPr><w:pStyle w:val="Title"/></w:pPr>'),
  label('Case 1: run-level, text branches'),
  p(r(ac([['wps', t('The Earth is flat')]], t('The Earth is round')))),
  label('Case 2: a tab and a break in the Choice'),
  p(r(ac([['wps', `${t('Before')}<w:tab/>${t('after the tab')}<w:br/>${t('after the break')}`]], t('Fallback text')))),
  label('Case 3: an x14ac Choice, understood since CR-010 section 15'),
  p(r(ac([['x14ac', t('unknown choice')]], t('fallback shown')))),
  label('Case 4: an x14ac Choice and no Fallback'),
  p(r(ac([['x14ac', t('nothing to select')]]))),
  label('Case 5: text and a picture in the Choice'),
  p(r(ac([['wps', `${t('A picture: ')}${picture}${t(' after it')}`]], t('no picture here')))),
  label('Case 6: ten Choices'),
  p(r(t('I prefer ')) + r(ac(TEN.map((x) => [x, t(x)]), t('The Earth is round')))),
  label('Case 7: ten Choices, w16du first'),
  p(r(t('I prefer ')) + r(ac([TEN[8], ...TEN.slice(0, 8), TEN[9]].map((x) => [x, t(x)]), t('The Earth is round')))),
  label('Case 8: paragraph-level, runs in the branches'),
  p(ac([['wps', r(t('bold ying'), '<w:rPr><w:b/></w:rPr>') + r(t(' plain ying'))]], r(t('yang')))),
  label('Case 9: ten Choices, w16sdtdh first'),
  p(r(t('I prefer ')) + r(ac([TEN[6], ...TEN.slice(0, 6), ...TEN.slice(7)].map((x) => [x, t(x)]), t('The Earth is round')))),
  label('Case 10: ten Choices, w16sdtfl first'),
  p(r(t('I prefer ')) + r(ac([TEN[7], ...TEN.slice(0, 7), ...TEN.slice(8)].map((x) => [x, t(x)]), t('The Earth is round')))),
  label('Case 11: a private Choice, the Fallback selected'),
  p(r(ac([['zz9', t('private choice')]], t('private fallback')))),
  label('Case 12: a private Choice and no Fallback'),
  p(r(ac([['zz9', t('private only')]]))),
  '<w:p/>',
].join('');

const pkg = await WordprocessingMLPackage.load(new Uint8Array(await readFile(src)));
const part = pkg.getMainDocumentPart();
let xml = await part.getXml();
const decls = {
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  w16se: 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
  w16cid: 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
  w16: 'http://schemas.microsoft.com/office/word/2018/wordml',
  w16cex: 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
  w16sdtdh: 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
  w16sdtfl: 'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock',
  w16du: 'http://schemas.microsoft.com/office/word/2023/wordml/word16du',
  wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  x14ac: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',   // Excel's: understood by Word 365 (ED-009 section 9.2) and by the engine since CR-010 section 15; until objects 0.4.3 the one unlisted prefix the engine's marshal kept declared
  zz9: 'http://example.com/zz9',   // private: understood by nothing; its declaration kept since objects 0.4.3 (jsonix 3.5.0)
};
const extra = Object.entries(decls).filter(([k]) => !xml.includes(`xmlns:${k}=`)).map(([k, v]) => ` xmlns:${k}="${v}"`).join('');
xml = xml.replace('<w:document ', `<w:document${extra} `);
xml = xml.replace(/<w:body>.*?(<w:sectPr)/s, `<w:body>${body}$1`);
part.setXml(xml);
await writeFile(out, new Uint8Array(await pkg.save()));
console.log('wrote', out);
