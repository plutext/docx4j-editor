// SPDX-License-Identifier: GPL-3.0-or-later
//
// The files for ED-005 proposal 46's Word check, part 6(a): how Word 365 resolves a Word fragment
// (a Flat OPC `pkg:package`) given to it as docx4j's processor gives it, an `w:altChunk` in a
// control tagged `od:progid=Word.Document`; and what Word does with the template before any
// processor ran. Run through the portfolio's Word bridge (Jason, 2026-11-11: Word 365 only, and
// perhaps Word for the Mac). Written by hand as XML, with no engine, so that nothing of the
// editor's or the engine's own writing is in what Word is asked about. From the repository root:
//
//   node packages/editor-model/test/fixtures/build/fragments-for-word.mjs [out-dir]
//
// (the default output is packages/editor-model/test/fixtures/fragments/; README.md there says what
// to look for).
//
// The fragment holds what part 2's rule must settle: a Heading 1 that differs from the document's,
// a style the document lacks, a numbered list whose numId the document also uses, a picture, a
// footnote, a table, a control inside it, and a section of another page size.
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { deflateSync } from 'node:zlib';
import { zipSync, strToU8 } from 'fflate';

const here = dirname(fileURLToPath(import.meta.url));
const out = process.argv[2] ?? join(here, '..', 'fragments');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const DATA_ID = '{4F6A2C1E-8B3D-4E5F-9A7B-1C2D3E4F5A6B}';
const XPATHS_ID = '{5A7B3D2F-9C4E-4F6A-8B1C-2D3E4F5A6B7C}';

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const t = (text) => `<w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
const p = (text, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${t(text)}</w:p>`;
const style = (id) => `<w:pStyle w:val="${id}"/>`;
const list = (numId) => `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`;

/** A PNG of w by h pixels in one colour. */
function png(w, h, [r, g, b]) {
  const table = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
  const crc = (buf) => { let c = 0xffffffff; for (const x of buf) c = table[(c ^ x) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const chunk = (type, data) => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type, 'ascii'), data]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([len, td, c]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  const raw = Buffer.alloc((1 + w * 3) * h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { const o = y * (1 + w * 3) + 1 + x * 3; raw[o] = r; raw[o + 1] = g; raw[o + 2] = b; }
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

// ---------------------------------------------------------------- the fragment, as Flat OPC

const PICTURE = png(160, 80, [0x2e, 0x75, 0xb6]);
const inlinePicture = (rid) => `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="1524000" cy="762000"/><wp:docPr id="1" name="Fragment picture"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="fragment.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1524000" cy="762000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

const fragmentDocument = `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>`
  + p('Fragment heading (the fragment\'s Heading 1: red, 20 pt)', style('Heading1'))
  + p('A paragraph in FragmentNote, a style the document lacks (green, italic).', style('FragmentNote'))
  + p('Fragment list, item one (the fragment\'s numId 1: upper-case Roman)', list(1))
  + p('Fragment list, item two', list(1))
  + `<w:p>${t('A footnote follows this sentence.')}<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>`
  + inlinePicture('rIdImage')
  + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>`
  + `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p('Service')}</w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p('Days')}</w:tc></w:tr>`
  + `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p('Standard')}</w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p('14')}</w:tc></w:tr></w:tbl>`
  + `<w:sdt><w:sdtPr><w:alias w:val="Inner control"/><w:tag w:val="inner"/><w:id w:val="201"/></w:sdtPr><w:sdtContent>${p('A control inside the fragment (tag "inner").')}</w:sdtContent></w:sdt>`
  + p('Last paragraph of the fragment. Its section is landscape A5; the document\'s is portrait A4.')
  + `<w:sectPr><w:pgSz w:w="11906" w:h="8391" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>`
  + '</w:body></w:document>';

const fragmentStyles = `<w:styles xmlns:w="${W}">`
  + '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="C00000"/><w:sz w:val="40"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:customStyle="1" w:styleId="FragmentNote"><w:name w:val="Fragment Note"/><w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:i/><w:color w:val="00B050"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style>'
  + '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
  + '</w:styles>';

const fragmentNumbering = `<w:numbering xmlns:w="${W}">`
  + '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
  + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
  + '</w:numbering>';

const fragmentFootnotes = `<w:footnotes xmlns:w="${W}">`
  + '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>'
  + '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
  + `<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>${t(' The fragment\'s footnote.')}</w:p></w:footnote>`
  + '</w:footnotes>';

const rels = (items) => `<Relationships xmlns="${PR}">${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${RT}/${type}" Target="${target}"/>`).join('')}</Relationships>`;
const xmlPart = (name, type, xml) => `<pkg:part pkg:name="${name}" pkg:contentType="${type}"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`;
const base64Lines = (bytes) => Buffer.from(bytes).toString('base64').replace(/.{76}/g, '$&\n');

/** The fragment as Word writes a Flat OPC document: the declarations, then the package's parts. */
const FRAGMENT = `${XML}<?mso-application progid="Word.Document"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">`
  + xmlPart('/_rels/.rels', 'application/vnd.openxmlformats-package.relationships+xml', rels([['rId1', 'officeDocument', 'word/document.xml']]))
  + xmlPart('/word/document.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', fragmentDocument)
  + xmlPart('/word/_rels/document.xml.rels', 'application/vnd.openxmlformats-package.relationships+xml', rels([
    ['rIdStyles', 'styles', 'styles.xml'], ['rIdNumbering', 'numbering', 'numbering.xml'], ['rIdFootnotes', 'footnotes', 'footnotes.xml'], ['rIdImage', 'image', 'media/image1.png'],
  ]))
  + xmlPart('/word/styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml', fragmentStyles)
  + xmlPart('/word/numbering.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml', fragmentNumbering)
  + xmlPart('/word/footnotes.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml', fragmentFootnotes)
  + `<pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png" pkg:compression="store"><pkg:binaryData>${base64Lines(PICTURE)}</pkg:binaryData></pkg:part>`
  + '</pkg:package>';

// ---------------------------------------------------------------- the document

const documentStyles = `<w:styles xmlns:w="${W}">`
  + '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
  + '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="28"/></w:rPr></w:style>'
  + '</w:styles>';

const documentNumbering = `<w:numbering xmlns:w="${W}">`
  + '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
  + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
  + '</w:numbering>';

/**
 * Word 2013's compatibility mode, as a new document in Word 365 has it. Without a settings part Word 365 opens a file in
 * mode 12 (Word 2007) and drops a w15 mapping on save: the first run (2026-11-11, the saves `-word365-mode12`) did.
 */
const settings = `<w:settings xmlns:w="${W}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;

const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/** The document around the fragment control, its own Heading 1 (blue, 14 pt) and its own list (decimal, numId 1). */
function documentXml(title, control) {
  return `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w15"><w:body>`
    + p(title)
    + p('Document heading (the document\'s Heading 1: blue, 14 pt)', style('Heading1'))
    + p('Document list, item one (decimal, numId 1)', list(1))
    + p('Document list, item two', list(1))
    + control
    + p('Document list, item three (after the fragment: 3, or does it count the fragment\'s items?)', list(1))
    + p('After the control.')
    + sectPr + '</w:body></w:document>';
}

const fragmentControl = (content) => `<w:sdt><w:sdtPr><w:alias w:val="Word fragment"/><w:tag w:val="od:xpath=x1&amp;od:progid=Word.Document"/><w:id w:val="101"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

/** The data part holding the fragment escaped as the text of /data/fragment, and the XPaths part naming it x1. */
const dataPart = `${XML}<data><fragment>${esc(FRAGMENT)}</fragment></data>`;
const xpathsPart = `${XML}<xpaths xmlns="http://opendope.org/xpaths"><xpath id="x1"><dataBinding xmlns:w="${W}" w:storeItemID="${DATA_ID}" w:xpath="/data/fragment" w:prefixMappings=""/></xpath></xpaths>`;
const itemProps = (id) => `${XML}<ds:datastoreItem ds:itemID="${id}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>`;

function docx(document, { chunk } = {}) {
  const docRels = [['rIdStyles', 'styles', 'styles.xml'], ['rIdNumbering', 'numbering', 'numbering.xml'], ['rIdSettings', 'settings', 'settings.xml'], ['rIdData', 'customXml', '../customXml/item1.xml'], ['rIdXPaths', 'customXml', '../customXml/item2.xml']];
  if (chunk) docRels.push(['rIdChunk', 'aFChunk', 'chunk1.xml']);
  const files = {
    '[Content_Types].xml': strToU8(`${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      + '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      + '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>'
      + '<Override PartName="/customXml/itemProps2.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>'
      + '</Types>'),
    '_rels/.rels': strToU8(XML + rels([['rId1', 'officeDocument', 'word/document.xml']])),
    'word/document.xml': strToU8(XML + document),
    'word/_rels/document.xml.rels': strToU8(XML + rels(docRels)),
    'word/styles.xml': strToU8(XML + documentStyles),
    'word/numbering.xml': strToU8(XML + documentNumbering),
    'word/settings.xml': strToU8(XML + settings),
    'customXml/item1.xml': strToU8(dataPart),
    'customXml/itemProps1.xml': strToU8(itemProps(DATA_ID)),
    'customXml/_rels/item1.xml.rels': strToU8(XML + rels([['rId1', 'customXmlProps', 'itemProps1.xml']])),
    'customXml/item2.xml': strToU8(xpathsPart),
    'customXml/itemProps2.xml': strToU8(itemProps(XPATHS_ID)),
    'customXml/_rels/item2.xml.rels': strToU8(XML + rels([['rId1', 'customXmlProps', 'itemProps2.xml']])),
  };
  // docx4j's AltChunkType.Xml: the Flat OPC as an application/xml part, which the Default above covers.
  if (chunk) files['word/chunk1.xml'] = strToU8(FRAGMENT);
  return zipSync(files, { level: 6 });
}

await mkdir(out, { recursive: true });

// 46a: docx4j's output, the control's content an altChunk (BindingTraverserXSLT.convertFlatOPC).
await writeFile(join(out, '46a-altchunk.docx'), docx(documentXml(
  '46a. A block-level control tagged od:progid=Word.Document whose content is a w:altChunk holding the fragment as Flat OPC (docx4j\'s form). Word resolves the chunk when it opens the file.',
  fragmentControl('<w:altChunk r:id="rIdChunk"/>'),
), { chunk: true }));

// 46b: the same with w:matchSrc, Word's "keep the source formatting" for a chunk.
await writeFile(join(out, '46b-altchunk-matchsrc.docx'), docx(documentXml(
  '46b. As 46a, the altChunk carrying w:altChunkPr/w:matchSrc (the chunk\'s own formatting asked to be kept).',
  fragmentControl('<w:altChunk r:id="rIdChunk"><w:altChunkPr><w:matchSrc/></w:altChunkPr></w:altChunk>'),
), { chunk: true }));

// 46c: the template before any processor: the fragment control with its placeholder, a plain text control mapped to the
// same node, and a rich text control mapped to it with w15:dataBinding (does Word fill it from the Flat OPC on open?).
await writeFile(join(out, '46c-template.docx'), docx(documentXml(
  '46c. The template as authored, before any processor: (1) the od:progid control holding its placeholder; (2) a plain text control mapped to the same node by w:dataBinding; (3) a rich text control mapped to it by w15:dataBinding.',
  fragmentControl(p('(1) Word fragment placeholder: the processor puts the fragment here.'))
    + `<w:p><w:sdt><w:sdtPr><w:alias w:val="Plain text mapped"/><w:id w:val="102"/><w:dataBinding w:prefixMappings="" w:xpath="/data[1]/fragment[1]" w:storeItemID="${DATA_ID}"/><w:text w:multiLine="1"/></w:sdtPr><w:sdtContent>${t('(2) plain text control placeholder')}</w:sdtContent></w:sdt></w:p>`
    + `<w:sdt><w:sdtPr><w:alias w:val="Rich text mapped"/><w:id w:val="103"/><w15:dataBinding w:prefixMappings="" w:xpath="/data[1]/fragment[1]" w:storeItemID="${DATA_ID}"/></w:sdtPr><w:sdtContent>${p('(3) rich text control placeholder')}</w:sdtContent></w:sdt>`,
)));

console.log(`wrote 46a-altchunk.docx, 46b-altchunk-matchsrc.docx and 46c-template.docx to ${out}`);
