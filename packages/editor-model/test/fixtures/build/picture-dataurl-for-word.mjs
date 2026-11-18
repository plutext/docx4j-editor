// SPDX-License-Identifier: GPL-3.0-or-later
//
// The files for ED-005 section 12.125's Word check: what Word 365 does with a picture content control mapped (`w:dataBinding`)
// to a node holding a data URL (`data:image/png;base64,...`), where Word itself writes bare base64. The editor and the engine
// read a data URL as its base64; whether Word does is what this asks. Run through the portfolio's Word bridge (Jason,
// 2026-11-12: "try the word bridge"). Written by hand as XML, with no engine. From the repository root:
//
//   node packages/editor-model/test/fixtures/build/picture-dataurl-for-word.mjs [out-dir]
//
// (the default output is packages/editor-model/test/fixtures/pictures-dataurl/; README.md there says what to look for).
//
// Four files, each one picture control mapped to /data/pic, the node holding a 120 by 60 orange PNG:
//   pdu1  bare base64, the control holding the editor's placeholder (the control: Word fills a mapped picture control on open);
//   pdu2  a data URL, the control holding the editor's placeholder;
//   pdu3  a data URL, the control holding a 120 by 60 blue picture (the editor's filled form, another picture so a refill shows);
//   pdu4  bare base64, the control holding the blue picture (pdu3's control).
// Round 2 (Jason, 2026-11-12, "yes stage the round", after run 1 found pdu1 left as it was), each over bare base64, to say why
// Word filled pdu4 and not pdu1: the flag, or no picture in the control to swap.
//   pdu5  the placeholder text without w:showingPlcHdr;
//   pdu6  Word's form as the ports write it (core-ts CR-005 section 10.1): the flag and a drawing of the ports' placeholder
//         picture, docx4j's picture-placeholder.png;
//   pdu7  nothing in the control, no flag.
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { deflateSync } from 'node:zlib';
import { zipSync, strToU8 } from 'fflate';

const here = dirname(fileURLToPath(import.meta.url));
const out = process.argv[2] ?? join(here, '..', 'pictures-dataurl');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const DATA_ID = '{6B8C4D2E-1A3F-4B5C-8D7E-9F0A1B2C3D4E}';

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const p = (text) => `<w:p><w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;

/** A PNG of w by h pixels in one colour, at 96 dpi (no pHYs). */
function png(w, h, [r, g, b]) {
  const table = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
  const crc = (buf) => { let c = 0xffffffff; for (const x of buf) c = table[(c ^ x) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const chunk = (type, data) => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type, 'ascii'), data]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([len, td, c]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  const raw = Buffer.alloc((1 + w * 3) * h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { const o = y * (1 + w * 3) + 1 + x * 3; raw[o] = r; raw[o + 1] = g; raw[o + 2] = b; }
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

/** The node's picture, orange; the control's own, blue. 120 by 60 pixels, 90 by 45 points at 96 dpi. */
const ORANGE = png(120, 60, [0xed, 0x7d, 0x31]);
const BLUE = png(120, 60, [0x2e, 0x75, 0xb6]);
const BASE64 = ORANGE.toString('base64');
const DATA_URL = `data:image/png;base64,${BASE64}`;

/**
 * The ports' placeholder picture: docx4j's `picture-placeholder.png` (`docx4j-core/src/main/resources/org/docx4j/model/
 * datastorage/`, 394589c18; core-ts's `PICTURE_PLACEHOLDER_BASE64`), 200 by 200 pixels, 676 bytes, SHA-1 e9e4ee21363e5c8d.
 */
const PLACEHOLDER = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAMgAAADICAIAAAAiOjnJAAAACXBIWXMAABYlAAAWJQFJUiTwAAACVklEQVR42u3cwU3jYBCGYRtx50gDNOUOckwDOaWBHNOBm6IAuEEBSByQIkv5B1njgGLP8xx3vasVeuVh82m3f/v46uDWHnwJEBbCQlggLISFsEBYCAthgbAQFsICYSEshAXCQlgU87jkF59Px+indvuDx9b+2PTH/y+sX/6gXdd9vr/O+U08ds+PnU/HdFtOIX+iT/8rncsbaxgGX8eNGcdx4TX0xsLfChEWwgJhISwq65f8pyAzP39jpZ6eX7yxcAoRFuTcYIQ26WyPSQenEGGBsBAWwgJhISxqMEITMkLjFCIsSDNC02CExilEWCAshIWwQFgIixqM0ISM0DiFCAvSjNA0GKFxChEWCAthISwQFsKiBiM0ISM0TiHCgjQjNA1GaJxChAXCQlgIC4SFsKjBCE3ICI1TiLAgzQhNgxEapxBhgbAQFsICYSEsajBCEzJC4xQiLEgzQtNghMYpRFggLISFsEBYCIsajNCEjNA4hQgL0ozQNBihcQoRFggLYSEsEBbCogYjNCEjNE4hwoI0IzQNRmicQoQFwkJYCAuEhbCowQhNyAiNU4iwIM0ITYMRGqcQYYGwEBbCAmEhLGowQhMyQuMUIixIM0LTYITGKURYICyEhbBAWAiLGozQhIzQOIUIC9KM0DQYoXEKERYIC2EhLBAWwqIGIzQhIzROIcKCNCM0DUZoNvfGuq57avoaaz7gsdU95o3Fmt9Yu/3h8m3WtZkfgXjsnh9Lf4PVLfyAFJxChIWwQFgIC2GBsBAWwgJhISyEBcJCWAgLhIWwEBb8+AbsaRgh6xNNrAAAAABJRU5ErkJggg==', 'base64');

const binding = `<w:dataBinding w:storeItemID="${DATA_ID}" w:xpath="/data/pic"/>`;

/** The editor's picture control holding its placeholder (contentControl.insert's Picture, as in Jason's file). */
const emptyControl = `<w:p><w:sdt><w:sdtPr><w:id w:val="235915003"/><w:picture/><w:showingPlcHdr/><w:alias w:val="Picture"/>${binding}</w:sdtPr>`
  + '<w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Click or tap here to enter text.</w:t></w:r></w:sdtContent></w:sdt></w:p>';

/** The control holding a picture, as the editor's refresh writes one: an inline drawing at the picture's natural size. */
const filledControl = `<w:p><w:sdt><w:sdtPr><w:id w:val="235915003"/><w:picture/><w:alias w:val="Picture"/>${binding}</w:sdtPr><w:sdtContent>`
  + '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1143000" cy="571500"/><wp:docPr id="1" name="Picture 1"/>'
  + '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
  + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="blue.png"/><pic:cNvPicPr/></pic:nvPicPr>'
  + '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
  + '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="571500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
  + '</wp:inline></w:drawing></w:r></w:sdtContent></w:sdt></w:p>';

/** The placeholder text, as the editor writes it, without the flag. */
const textNoFlagControl = `<w:p><w:sdt><w:sdtPr><w:id w:val="235915003"/><w:picture/><w:alias w:val="Picture"/>${binding}</w:sdtPr>`
  + '<w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Click or tap here to enter text.</w:t></w:r></w:sdtContent></w:sdt></w:p>';

/** Word's form, as the ports write it (core-ts `showPicturePlaceholder`): the flag, one run holding the placeholder drawing, 100 points square. */
const placeholderControl = `<w:p><w:sdt><w:sdtPr><w:id w:val="235915003"/><w:picture/><w:showingPlcHdr/><w:alias w:val="Picture"/>${binding}</w:sdtPr><w:sdtContent>`
  + '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1270000" cy="1270000"/><wp:docPr id="1" name="Picture 1" descr=""/>'
  + '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
  + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="picture-placeholder.png"/><pic:cNvPicPr/></pic:nvPicPr>'
  + '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
  + '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="1270000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
  + '</wp:inline></w:drawing></w:r></w:sdtContent></w:sdt></w:p>';

/** Nothing in the control, and no flag. */
const emptyContentControl = `<w:p><w:sdt><w:sdtPr><w:id w:val="235915003"/><w:picture/><w:alias w:val="Picture"/>${binding}</w:sdtPr><w:sdtContent/></w:sdt></w:p>`;

const documentXml = (title, control) => `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>`
  + p(title) + control + p('After the control.')
  + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>';

const styles = `<w:styles xmlns:w="${W}">`
  + '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
  + '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rPr><w:color w:val="808080"/></w:rPr></w:style>'
  + '</w:styles>';

/** Word 2013's compatibility mode, as a new document in Word 365 has it (section 12.118's run 1 found mode 12 without it). */
const settings = `<w:settings xmlns:w="${W}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;
const rels = (items) => `<Relationships xmlns="${PR}">${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${RT}/${type}" Target="${target}"/>`).join('')}</Relationships>`;

function docx(document, value, picture) {
  const docRels = [['rIdStyles', 'styles', 'styles.xml'], ['rIdSettings', 'settings', 'settings.xml'], ['rIdData', 'customXml', '../customXml/item1.xml']];
  if (picture) docRels.push(['rIdImage', 'image', 'media/image1.png']);
  const files = {
    '[Content_Types].xml': strToU8(`${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>`
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      + '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>'
      + '</Types>'),
    '_rels/.rels': strToU8(XML + rels([['rId1', 'officeDocument', 'word/document.xml']])),
    'word/document.xml': strToU8(XML + document),
    'word/_rels/document.xml.rels': strToU8(XML + rels(docRels)),
    'word/styles.xml': strToU8(XML + styles),
    'word/settings.xml': strToU8(XML + settings),
    'customXml/item1.xml': strToU8(`${XML}<data><pic>${value}</pic></data>`),
    'customXml/itemProps1.xml': strToU8(`${XML}<ds:datastoreItem ds:itemID="${DATA_ID}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>`),
    'customXml/_rels/item1.xml.rels': strToU8(XML + `<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${RT}/customXmlProps" Target="itemProps1.xml"/></Relationships>`),
  };
  if (picture) files['word/media/image1.png'] = new Uint8Array(picture);
  return zipSync(files, { mtime: new Date('2026-11-12T00:00:00Z') });
}

const FILES = [
  ['pdu1-base64-empty', 'pdu1: the node holds bare base64; the control holds its placeholder.', BASE64, emptyControl, undefined],
  ['pdu2-dataurl-empty', 'pdu2: the node holds a data URL; the control holds its placeholder.', DATA_URL, emptyControl, undefined],
  ['pdu3-dataurl-filled', 'pdu3: the node holds a data URL (orange); the control holds a blue picture.', DATA_URL, filledControl, BLUE],
  ['pdu4-base64-filled', 'pdu4: the node holds bare base64 (orange); the control holds a blue picture.', BASE64, filledControl, BLUE],
  // Round 2
  ['pdu5-base64-text-noflag', 'pdu5: the node holds bare base64; the control holds the placeholder text without the flag.', BASE64, textNoFlagControl, undefined],
  ['pdu6-base64-placeholder-flag', 'pdu6: the node holds bare base64; the control holds the ports\' placeholder picture with the flag.', BASE64, placeholderControl, PLACEHOLDER],
  ['pdu7-base64-empty-content', 'pdu7: the node holds bare base64; the control holds nothing.', BASE64, emptyContentControl, undefined],
];

await mkdir(out, { recursive: true });
for (const [name, title, value, control, picture] of FILES) {
  await writeFile(join(out, `${name}.docx`), docx(documentXml(title, control), value, picture));
  console.log(`${name}.docx`);
}
