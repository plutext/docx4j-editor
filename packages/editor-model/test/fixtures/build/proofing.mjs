// SPDX-License-Identifier: GPL-3.0-or-later
//
// The fixture for the editor's proofing (view/proofing.ts; Jason, 2026-11-11): a document in Australian English with a run
// in French, a run Word marks "Do not check spelling or grammar" (w:noProof), a paragraph in a style that sets w:noProof
// (Word's HTML Code is one such), a run in Japanese whose w:lang names ja-JP for East Asian text, and a character style in
// German. Written by hand as XML. From the repository root:
//
//   node packages/editor-model/test/fixtures/build/proofing.mjs [out.docx]
import { writeFile, mkdir } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync, strToU8 } from 'fflate';

const here = dirname(fileURLToPath(import.meta.url));
const out = process.argv[2] ?? join(here, '..', 'proofing', 'proofing.docx');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const run = (text, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

const document = `<w:document xmlns:w="${W}"><w:body>`
  + `<w:p>${run('Plain Australian text, its colour and its organisation. ')}${run('Une phrase en français.', '<w:lang w:val="fr-FR"/>')}${run(' Then Quarnby and Velmor, ')}${run('marked Do not check spelling', '<w:noProof/>')}${run('.')}</w:p>`
  + `<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr>${run('const gizmo = frobnicate(800);')}</w:p>`
  + `<w:p>${run('Japanese: ')}${run('日本語の文章です。', '<w:lang w:val="en-AU" w:eastAsia="ja-JP"/>')}${run(' and German by style: ')}${run('Ein deutscher Satz.', '<w:rStyle w:val="German"/>')}</w:p>`
  + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
  + '</w:body></w:document>';

const styles = `<w:styles xmlns:w="${W}">`
  + '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:lang w:val="en-AU" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault></w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:noProof/></w:rPr></w:style>'
  + '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>'
  + '<w:style w:type="character" w:customStyle="1" w:styleId="German"><w:name w:val="German"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:lang w:val="de-DE"/></w:rPr></w:style>'
  + '</w:styles>';

const rels = (items) => `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`).join('')}</Relationships>`;
const zip = zipSync({
  '[Content_Types].xml': strToU8(`${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`),
  '_rels/.rels': strToU8(rels([['rId1', 'officeDocument', 'word/document.xml']])),
  'word/document.xml': strToU8(XML + document),
  'word/_rels/document.xml.rels': strToU8(rels([['rId1', 'styles', 'styles.xml']])),
  'word/styles.xml': strToU8(XML + styles),
}, { level: 6 });
await mkdir(dirname(out), { recursive: true });
await writeFile(out, zip);
console.log(`wrote ${out}`);
