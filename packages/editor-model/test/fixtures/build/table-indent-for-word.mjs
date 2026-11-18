// SPDX-License-Identifier: GPL-3.0-or-later
//
// The files for ED-005 section 12.132 item 5's Word check (Jason, 2026-11-18: "build a small test file then ask portfolio
// agent to have Word run it via the bridge"): where Word puts a left-aligned table's edge, and so its cells' text, against
// the body text, in compatibility modes 15 and 14 and with no mode stated. The editor outdents every left-aligned table by
// the table's default left cell margin (`tableVars`, view/tables.ts), Word 2010's placement; Word 365's PDF of a mode 15
// form put a table whose cells' own left margin is 0 at the margin. Written by hand as XML, with no engine:
//
//   node packages/editor-model/test/fixtures/build/table-indent-for-word.mjs [out-dir]
//
// (the default output is packages/editor-model/test/fixtures/table-indent/; README.md there says what to look for).
//
// Each file holds the same body, a reference paragraph and seven bordered tables T1 to T7, and differs only in its
// settings: ind15 compatibility mode 15 (Word 2013 and later, a new document in Word 365), ind14 mode 14 (Word 2010),
// ind00 no compatibility setting at all (as many generated files have).
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync, strToU8 } from 'fflate';

const here = dirname(fileURLToPath(import.meta.url));
const out = process.argv[2] ?? join(here, '..', 'table-indent');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const p = (text) => `<w:p><w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;

const COL = 4000;
const BORDERS = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
/** A cell of one paragraph; `left` is its own `w:tcMar` left margin in twips, none when undefined. */
const tc = (text, left) => `<w:tc><w:tcPr><w:tcW w:w="${COL}" w:type="dxa"/>${left === undefined ? '' : `<w:tcMar><w:left w:w="${left}" w:type="dxa"/></w:tcMar>`}</w:tcPr>${p(text)}</w:tc>`;
const tr = (...cells) => `<w:tr>${cells.join('')}</w:tr>`;
/** A table: `ind` its own `w:tblInd`, `cellLeft` its own `w:tblCellMar` left margin (twips), each none when undefined. */
const tbl = ({ ind, cellLeft }, rows) => '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
  + (ind === undefined ? '' : `<w:tblInd w:w="${ind}" w:type="dxa"/>`)
  + BORDERS
  + (cellLeft === undefined ? '' : `<w:tblCellMar><w:left w:w="${cellLeft}" w:type="dxa"/></w:tblCellMar>`)
  + '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>'
  + `<w:tblGrid><w:gridCol w:w="${COL}"/><w:gridCol w:w="${COL}"/></w:tblGrid>${rows.join('')}</w:tbl>`;

// Every text starts with "T", the reference paragraph's too, so a glyph's side bearing is the same everywhere.
const body = p('T0 body text at the margin: every table\'s first cell text is measured against this line\'s left edge.')
  + p('T1: no indent, the default cell margins (108).')
  + tbl({}, [tr(tc('T1 first cell'), tc('T1 second'))])
  + p('T2: no indent, the table\'s own left cell margin 400.')
  + tbl({ cellLeft: 400 }, [tr(tc('T2 first cell'), tc('T2 second'))])
  + p('T3: no indent, the default cell margins, the first cell\'s own left margin 0 (the share application\'s form).')
  + tbl({}, [tr(tc('T3 first cell', 0), tc('T3 second', 0))])
  + p('T4: no indent, the default cell margins, the first cell\'s own left margin 400.')
  + tbl({}, [tr(tc('T4 first cell', 400), tc('T4 second'))])
  + p('T5: indent 720, the default cell margins.')
  + tbl({ ind: 720 }, [tr(tc('T5 first cell'), tc('T5 second'))])
  + p('T6: indent 720, the first cell\'s own left margin 400.')
  + tbl({ ind: 720 }, [tr(tc('T6 first cell', 400), tc('T6 second'))])
  + p('T7: no indent, two rows: row 1\'s first cell\'s own left margin 400, row 2\'s none (the default 108).')
  + tbl({}, [tr(tc('T7 row 1', 400), tc('T7 r1 second')), tr(tc('T7 row 2'), tc('T7 r2 second'))])
  + p('T8 after the tables.');

const documentXml = `<w:document xmlns:w="${W}"><w:body>${body}`
  + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>';

const styles = `<w:styles xmlns:w="${W}">`
  + '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="120" w:after="60"/></w:pPr></w:pPrDefault></w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
  + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>'
  + '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
  + '</w:styles>';

/** The settings part: a compatibility mode, or none stated. */
const settings = (mode) => `<w:settings xmlns:w="${W}">`
  + (mode === undefined ? '' : `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/></w:compat>`)
  + '</w:settings>';
const rels = (items) => `<Relationships xmlns="${PR}">${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${RT}/${type}" Target="${target}"/>`).join('')}</Relationships>`;

function docx(mode) {
  const files = {
    '[Content_Types].xml': strToU8(`${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'),
    '_rels/.rels': strToU8(XML + rels([['rId1', 'officeDocument', 'word/document.xml']])),
    'word/document.xml': strToU8(XML + documentXml),
    'word/_rels/document.xml.rels': strToU8(XML + rels([['rIdStyles', 'styles', 'styles.xml'], ['rIdSettings', 'settings', 'settings.xml']])),
    'word/styles.xml': strToU8(XML + styles),
    'word/settings.xml': strToU8(XML + settings(mode)),
  };
  // A fixed date, so a rebuild is byte for byte.
  return zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, [v, { mtime: new Date('2026-11-18T00:00:00Z') }]])));
}

await mkdir(out, { recursive: true });
await writeFile(join(out, 'ind15-mode15.docx'), docx(15));
await writeFile(join(out, 'ind14-mode14.docx'), docx(14));
await writeFile(join(out, 'ind00-no-mode.docx'), docx(undefined));
console.log(`wrote ind15, ind14 and ind00 in ${out}`);
