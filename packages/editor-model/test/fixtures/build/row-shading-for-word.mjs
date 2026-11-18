// SPDX-License-Identifier: GPL-3.0-or-later
//
// The files for ED-005 section 12.128's Word check (Jason, 2026-11-15: "set up the Word bridge round"): what Word 365 shows
// and saves for a table cell's own shading in the forms a "no colour" could take, under a table style that shades the cell;
// which `w:tc` of a vertical merge Word draws the merged cell's shading from; and whether Word reads a file the editor's
// row.shading wrote as the editor draws it. Run from the repository root after `npm run build -w @docx4j/editor-model`
// (shd3 is the model's own save; the others are written by hand as XML, with no engine):
//
//   node packages/editor-model/test/fixtures/build/row-shading-for-word.mjs [out-dir]
//
// (the default output is packages/editor-model/test/fixtures/row-shading/; README.md there says what to look for).
//
//   shd0  the base the editor edits for shd3 (not a job);
//   shd1  two styled tables, a row per form: a style shading the whole table orange (table A), and a style banding its
//         odd rows orange (table B);
//   shd2  three unstyled two-by-two tables, the second column merged down: its shading on the restart cell alone (C), on
//         the continuation alone (D), on both (E, the editor's form);
//   shd3  shd0 after four row.shading commands through the model's headless runner, saved by the model.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync, strToU8 } from 'fflate';

const here = dirname(fileURLToPath(import.meta.url));
const out = process.argv[2] ?? join(here, '..', 'row-shading');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const p = (text) => `<w:p><w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;

const ORANGE = 'FFC000';
/** A cell's own shading in each form the round asks about; undefined writes none. */
const SHD = {
  none: undefined,
  fillAuto: '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>',
  nil: '<w:shd w:val="nil"/>',
  white: '<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/>',
  yellow: '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>',
};

const COL = 4513;   // half the A4 text width with one-inch margins, in twips
/** A cell: its width, a merge, its shading, one paragraph of text. */
const tc = (text, { shd, vMerge } = {}) => `<w:tc><w:tcPr><w:tcW w:w="${COL}" w:type="dxa"/>${vMerge ?? ''}${shd ?? ''}</w:tcPr>${p(text)}</w:tc>`;
const tr = (...cells) => `<w:tr>${cells.join('')}</w:tr>`;
const BORDERS = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>';
/** Horizontal banding on, nothing else of the style's conditions (both of Word's forms of the look). */
const LOOK = '<w:tblLook w:val="0400" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>';
const tbl = (style, rows) => `<w:tbl><w:tblPr>${style ? `<w:tblStyle w:val="${style}"/>` : ''}<w:tblW w:w="0" w:type="auto"/>${style ? '' : BORDERS}${LOOK}</w:tblPr>`
  + `<w:tblGrid><w:gridCol w:w="${COL}"/><w:gridCol w:w="${COL}"/></w:tblGrid>${rows.join('')}</w:tbl>`;
/** A row of table A or B: the form's name, and what the cell holds. */
const formRow = (name, form) => tr(tc(name, { shd: SHD[form] }), tc(form === 'none' ? 'no w:shd of its own' : SHD[form], { shd: SHD[form] }));

const documentXml = (body) => `<w:document xmlns:w="${W}"><w:body>${body}`
  + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>';

const styles = `<w:styles xmlns:w="${W}">`
  + '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
  + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>'
  + '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
  // Table A's style: every cell shaded orange, from the style's own w:tcPr.
  + '<w:style w:type="table" w:styleId="ShadedTable"><w:name w:val="Shaded Table"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="40"/>'
  + `<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr>${BORDERS}</w:tblPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${ORANGE}"/></w:tcPr></w:style>`
  // Table B's style: the odd rows orange, a band of one row.
  + '<w:style w:type="table" w:styleId="BandedTable"><w:name w:val="Banded Table"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="41"/>'
  + `<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblStyleRowBandSize w:val="1"/>${BORDERS}</w:tblPr>`
  + `<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${ORANGE}"/></w:tcPr></w:tblStylePr></w:style>`
  + '</w:styles>';

/** Word 2013's compatibility mode, as a new document in Word 365 has it. */
const settings = `<w:settings xmlns:w="${W}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;
const rels = (items) => `<Relationships xmlns="${PR}">${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${RT}/${type}" Target="${target}"/>`).join('')}</Relationships>`;

function docx(body) {
  const files = {
    '[Content_Types].xml': strToU8(`${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'),
    '_rels/.rels': strToU8(XML + rels([['rId1', 'officeDocument', 'word/document.xml']])),
    'word/document.xml': strToU8(XML + documentXml(body)),
    'word/_rels/document.xml.rels': strToU8(XML + rels([['rIdStyles', 'styles', 'styles.xml'], ['rIdSettings', 'settings', 'settings.xml']])),
    'word/styles.xml': strToU8(XML + styles),
    'word/settings.xml': strToU8(XML + settings),
  };
  // A fixed date, so a rebuild is byte for byte.
  return zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, [v, { mtime: new Date('2026-11-15T00:00:00Z') }]])));
}

const restart = '<w:vMerge w:val="restart"/>';
const cont = '<w:vMerge/>';

const shd1 = docx(
  p('Table A: the style shades every cell orange. Each row is one form of a cell\'s own shading.')
  + tbl('ShadedTable', [formRow('A1 none', 'none'), formRow('A2 fill auto', 'fillAuto'), formRow('A3 val nil', 'nil'), formRow('A4 white', 'white'), formRow('A5 yellow', 'yellow')])
  + p('Table B: the style shades the odd rows orange (B1, B3, B5); the even rows are not banded.')
  + tbl('BandedTable', [formRow('B1 none, band', 'none'), formRow('B2 none, no band', 'none'), formRow('B3 fill auto, band', 'fillAuto'), formRow('B4 none, no band', 'none'), formRow('B5 val nil, band', 'nil'), formRow('B6 none, no band', 'none')])
  + p('After the tables.'));

const merged = (label, top, below) => tbl(null, [
  tr(tc(`${label} r1 c1`), tc(`${label} merged cell`, { vMerge: restart, shd: top })),
  tr(tc(`${label} r2 c1`), tc('', { vMerge: cont, shd: below })),
]);
const shd2 = docx(
  p('Table C: the merged cell\'s restart w:tc yellow, its continuation none.') + merged('C', SHD.yellow, undefined)
  + p('Table D: the restart none, the continuation yellow.') + merged('D', undefined, SHD.yellow)
  + p('Table E: both yellow, as the editor writes a merge.') + merged('E', SHD.yellow, SHD.yellow)
  + p('After the tables.'));

// The editor's base: what each cell's text says is what the editor shows once shd3's commands have run.
const plainRow = (a, b) => tr(tc(a), tc(b));
const shd0 = docx(
  p('Table 1, plain: row 1 yellow at a caret; rows 2 and 3 grey from a cell selection of their second column.')
  + tbl(null, [plainRow('1 r1: yellow', 'yellow'), plainRow('1 r2: grey', 'grey'), plainRow('1 r3: grey', 'grey')])
  + p('Table 2, the second column merged down: row 1 light blue at a caret; the merged cell light blue to its foot; r2 c1 white.')
  + tbl(null, [tr(tc('2 r1: light blue'), tc('2 merged: light blue', { vMerge: restart })), tr(tc('2 r2 c1: white'), tc('', { vMerge: cont }))])
  + p('Table 3, the orange style: row 1 held yellow of its own, taken away (No colour), so orange as row 2.')
  + tbl('ShadedTable', [tr(tc('3 r1: orange', { shd: SHD.yellow }), tc('orange', { shd: SHD.yellow })), plainRow('3 r2: orange', 'orange')])
  + p('After the tables.'));

await mkdir(out, { recursive: true });
await writeFile(join(out, 'shd0-base.docx'), shd0);
await writeFile(join(out, 'shd1-style-forms.docx'), shd1);
await writeFile(join(out, 'shd2-merged.docx'), shd2);

// shd3: the model's own save of shd0 after the four commands, as the menu runs them.
const { EditorDocument, runCommand, CellSelection, TableMap } = await import('../../../dist/index.mjs');
const session = await EditorDocument.open(new Uint8Array(await readFile(join(out, 'shd0-base.docx'))));
/** The position of the cell node at (r, c) of the table that is block `i`. */
const cellAt = (i, r, c) => {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += session.doc.child(k).nodeSize;
  const map = TableMap.get(session.doc.child(i));
  return pos + 1 + map.map[r * map.width + c];
};
const run = (args, selection) => {
  const result = runCommand(session, 'row.shading', args, selection);
  if (!result.ok) throw new Error(`row.shading ${JSON.stringify(args)} did not apply`);
};
run({ color: '#FFFF00' }, { from: cellAt(1, 0, 0) + 2 });
run({ color: '#D9D9D9' }, CellSelection.create(session.doc, cellAt(1, 1, 1), cellAt(1, 2, 1)));
run({ color: '#00B0F0' }, { from: cellAt(3, 0, 0) + 2 });
run({ color: null }, { from: cellAt(5, 0, 0) + 2 });
const saved = await session.save();
if (!saved.changed) throw new Error('shd3: the save wrote no change');
await writeFile(join(out, 'shd3-editor.docx'), saved.bytes);
console.log(`wrote shd0 to shd3 in ${out}`);
