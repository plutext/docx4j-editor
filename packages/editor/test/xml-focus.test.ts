// SPDX-License-Identifier: GPL-3.0-or-later
//
// Reveal codes in a table (ED-005 section 12.99): the outline of an XML text with its tags' lines, the folding regions
// made from it, the caret's chain of tables from the state (the grid column through TableMap, past a vertical merge),
// and the focus in the text: the caret's cell marked, the other rows and cells folded, through a row's control
// wrapper and a nested table.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { TextSelection } from 'prosemirror-state';
import { EditorDocument } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { outlineOf, foldingRangesOf, cellFocusOf, tableChainOf, pretty, xmlOf } from '../src/app/xml.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

/** The one-based lines of a text matching a pattern, in order. */
const linesOf = (text: string, re: RegExp): number[] => text.split('\n').flatMap((l, i) => (re.test(l) ? [i + 1] : []));

describe('outlineOf', () => {
  test('every element with the lines of its tags, its attributes, and its children; the rest of the markup passed over', () => {
    const text = '<?xml version="1.0"?>\n<a x="1">\n  <!-- c -->\n  <b y=\'2\'/>\n  <c>\n    <d>t</d>\n  </c>\n</a>\n';
    const [a] = outlineOf(text);
    expect(a).toMatchObject({ name: 'a', local: 'a', attrs: { x: '1' }, line: 2, endLine: 8 });
    expect(a!.children.map((c) => [c.name, c.line, c.endLine])).toEqual([['b', 4, 4], ['c', 5, 7]]);
    expect(a!.children[0]!.attrs).toEqual({ y: '2' });
    expect(a!.children[1]!.children[0]).toMatchObject({ name: 'd', line: 6, endLine: 6, children: [] });
    expect(outlineOf('<w:p/>\n<w:tbl>\n</w:tbl>').map((e) => [e.local, e.line, e.endLine])).toEqual([['p', 1, 1], ['tbl', 2, 3]]);
    // The regions: the elements that span lines, start tag to end tag.
    expect(foldingRangesOf(text)).toEqual([{ start: 2, end: 8 }, { start: 5, end: 7 }]);
  });
});

describe('cellFocusOf', () => {
  const table = pretty('<w:tbl xmlns:w="urn:w"><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1"/><w:gridCol w:w="2"/><w:gridCol w:w="3"/></w:tblGrid>'
    + '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>heading</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr>'
    + '<w:sdt><w:sdtPr><w:tag w:val="row"/></w:sdtPr><w:sdtContent><w:tr>'
    + '<w:tc><w:p/></w:tc><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>'
    + '<w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p/></w:tc><w:tc><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc>'
    + '</w:tr></w:sdtContent></w:sdt></w:tbl>');
  const rows = linesOf(table, /^\s*<w:tr>/);
  const cells = linesOf(table, /^\s*<w:tc>/);
  const ends = linesOf(table, /^\s*<\/w:tc>/);

  test('the cell at a grid column, past a merged cell and a continuation, in a row inside a control: the other rows and cells folded', () => {
    expect(rows).toHaveLength(3);
    expect(cells).toHaveLength(7);
    // Row 1 (in the sdt), column 2: the third cell (the first at column 0, the continuation at 1).
    const f = cellFocusOf(outlineOf(table), [{ table: 0, row: 1, column: 2 }])!;
    expect(f.line).toBe(cells[4]);
    expect(f.endLine).toBe(ends[6]);   // its end tag closes after the nested table's cells
    expect(f.fold).toEqual([rows[0], cells[2], cells[3]]);
    // Row 0, column 2: the second cell, since the first spans two.
    const g = cellFocusOf(outlineOf(table), [{ table: 0, row: 0, column: 2 }])!;
    expect(g.line).toBe(cells[1]);
    expect(g.fold).toEqual([rows[1], cells[0]]);
    // A column no cell starts at: the row is the focus, its rows folded.
    const h = cellFocusOf(outlineOf(table), [{ table: 0, row: 0, column: 1 }])!;
    expect(h.line).toBe(rows[0]);
    expect(h.fold).toEqual([rows[1]]);
  });

  test("a nested table: the step's table among the cell's, its other cells folded too", () => {
    const f = cellFocusOf(outlineOf(table), [{ table: 0, row: 1, column: 2 }, { table: 0, row: 0, column: 1 }])!;
    expect(f.line).toBe(cells[6]);
    expect(f.endLine).toBe(ends[5]);
    expect(f.fold).toEqual([rows[0], cells[2], cells[3], cells[5]]);
    // A row the XML has not: the chain ends with what was found.
    expect(cellFocusOf(outlineOf(table), [{ table: 0, row: 1, column: 2 }, { table: 0, row: 3, column: 0 }])).toMatchObject({ line: cells[4], fold: [rows[0], cells[2], cells[3]] });
    expect(cellFocusOf(outlineOf(table), [{ table: 1, row: 0, column: 0 }])).toBeUndefined();
    expect(cellFocusOf(outlineOf(table), [])).toBeUndefined();
  });

  test('a row with cells before its grid starts counts from w:gridBefore', () => {
    const text = pretty('<w:tbl xmlns:w="urn:w"><w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>');
    const tcs = linesOf(text, /^\s*<w:tc>/);
    expect(cellFocusOf(outlineOf(text), [{ table: 0, row: 0, column: 1 }])!.line).toBe(tcs[0]);
    expect(cellFocusOf(outlineOf(text), [{ table: 0, row: 0, column: 2 }])!.line).toBe(tcs[1]);
  });
});

describe('tableChainOf', () => {
  test('the caret in a fixture table: the row, the grid column past a vertical merge, and the focus in the block\'s XML', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'tables', 'table-fit-fixed.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session, { headersAndFooters: false });
    const doc = editor.state.doc;
    const caretInCell = (block: number, row: number, cell: number): void => {
      let pos = 0;
      for (let k = 0; k < block; k++) pos += doc.child(k).nodeSize;
      const table = doc.child(block);
      pos += 1;
      for (let k = 0; k < row; k++) pos += table.child(k).nodeSize;
      const tr = table.child(row);
      pos += 1;
      for (let k = 0; k < cell; k++) pos += tr.child(k).nodeSize;
      editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos + 2)));
    };
    // Table B (block 4): a heading over two columns, then a vertical merge in the first column.
    caretInCell(4, 0, 1);
    expect(tableChainOf(editor.state)).toEqual([{ table: 0, row: 0, column: 2 }]);
    caretInCell(4, 2, 0);
    const chain = tableChainOf(editor.state);
    expect(chain).toEqual([{ table: 0, row: 2, column: 1 }]);
    const text = await xmlOf(editor.exportBlock(doc.child(4)));
    const focus = cellFocusOf(outlineOf(text), chain)!;
    const lines = text.split('\n');
    expect(lines.slice(focus.line - 1, focus.endLine).join('\n')).toContain('Where the premises are damaged');
    expect(lines[focus.line - 1]).toMatch(/<w:tc>/);
    expect(lines[focus.endLine - 1]).toMatch(/<\/w:tc>/);
    // The other two rows and, in this row, the continuation cell and the last cell are folded.
    const rows = linesOf(text, /^\s*<w:tr[ >]/);
    expect(focus.fold.slice(0, 2)).toEqual([rows[0], rows[1]]);
    expect(focus.fold).toHaveLength(4);
    expect(focus.fold).not.toContain(focus.line);
    // Outside a table: no chain.
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1)));
    expect(tableChainOf(editor.state)).toEqual([]);
    editor.destroy();
  });
});
