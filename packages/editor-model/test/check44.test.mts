// SPDX-License-Identifier: GPL-3.0-or-later
//
// Check 44 (ED-007 section 8.1 item 13): Jason's Word 365 save of the optional-breaks probe. Word kept every U+200B
// where it was and every w:softHyphen; the editor reads the one as text and the other as an atom, and writes the file
// back unedited part for part.
import { describe, test, expect } from 'vitest';
import { rowsOf, cellsOf } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { EditorDocument, OPTIONAL_BREAK, TableMap } from '../src/index.mjs';
import { fixture, comparePackages } from './helpers.mjs';

describe('check 44: optional breaks in Word 365', () => {
  test('the save keeps U+200B as text and w:softHyphen as an element; the editor round-trips it untouched', async () => {
    const original = await fixture('tables/check44/optional-breaks-probe-Word365.docx');
    const session = await EditorDocument.open(original);
    const tbl = session.body.content!.find((e) => e.name.localPart === 'tbl')!.value as wml.Tbl;
    const cell = (r: number, c: number) => cellsOf(rowsOf(tbl)[r]!.element.value)[c]!.element;
    const xmlOf = (r: number, c: number): string => JSON.stringify(cell(r, c));
    expect(xmlOf(1, 1)).toContain(`Gamma.${OPTIONAL_BREAK}Den`);
    expect(xmlOf(2, 1)).toContain(`Eta${OPTIONAL_BREAK}Eastern`);
    expect(xmlOf(4, 1)).toContain(`pri${OPTIONAL_BREAK}ce`);
    expect(xmlOf(1, 2)).toContain('softHyphen');
    expect(xmlOf(1, 1)).not.toContain('softHyphen');
    // In the editor: the character inside the cell's text; the soft hyphen a node of its own between two texts.
    const table = session.doc.child(1);
    const map = TableMap.get(table);
    const at = (r: number, c: number) => table.nodeAt(map.map[r * map.width + c]!)!;
    expect(at(1, 1).textContent).toBe(`Gamma.${OPTIONAL_BREAK}Den`);
    const soft = at(1, 2).firstChild!;
    expect(soft.childCount).toBe(3);
    expect(soft.child(0).text).toBe('Gamma.');
    expect(soft.child(1).isText).toBe(false);
    expect(soft.child(2).text).toBe('Den');
    const saved = await session.save();
    expect(comparePackages(original, saved.bytes, ['[Content_Types].xml'])).toEqual([]);
  });
});
