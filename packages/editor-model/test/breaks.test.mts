// SPDX-License-Identifier: GPL-3.0-or-later
//
// Optional breaks (ED-007 section 8.1 item 13): where a word may take one, the table command putting them into the
// words named, once, and the file carrying the character; the caret's own.
import { describe, test, expect } from 'vitest';
import { rowsOf, cellsOf, textOf as textOfElement } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { EditorDocument, runCommand, breakOffsetsIn, OPTIONAL_BREAK, WORD_RUN, TableMap } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

describe('breakOffsetsIn', () => {
  test('after an inner period before a letter, before a capital after a lower-case letter; nowhere else', () => {
    expect(breakOffsetsIn('Gamma.Den')).toEqual([6]);
    expect(breakOffsetsIn('EtaEastern')).toEqual([3]);
    expect(breakOffsetsIn('Lambda.Woods')).toEqual([7]);
    expect(breakOffsetsIn('EPSILON')).toEqual([]);
    expect(breakOffsetsIn('DsWood')).toEqual([2]);
    expect(breakOffsetsIn('Redistribution')).toEqual([]);
    expect(breakOffsetsIn('e.g.')).toEqual([2]);
    expect(breakOffsetsIn('3.5')).toEqual([]);
    expect(breakOffsetsIn('docx4j.org')).toEqual([7]);
    expect(breakOffsetsIn('ÜberStraße')).toEqual([4]);
  });

  test('a word is a run with no blank, hyphen, dash, soft hyphen or optional break', () => {
    const words = (s: string): string[] => Array.from(s.matchAll(WORD_RUN), (m) => m[0]);
    expect(words('in the price')).toEqual(['in', 'the', 'price']);
    expect(words('$1,250–$2,500/yr sub')).toEqual(['$1,250', '$2,500/yr', 'sub']);
    expect(words(`Gamma.${OPTIONAL_BREAK}Den Eta­Eastern a-b`)).toEqual(['Gamma.', 'Den', 'Eta', 'Eastern']);
  });
});

describe('table.breakWords and range.insertOptionalBreak', () => {
  test('the words named take their breaks in every cell, once; the file carries the character', async () => {
    const original = await fixture('tables/table-fit-fixed.docx');
    const session = await EditorDocument.open(original);
    // Table E (block 13): "Gamma.Den" at (0, 3), "EtaEastern" at (0, 7), "Zetaforest" at (0, 6) has nowhere to break.
    let pos = 0;
    for (let k = 0; k < 13; k++) pos += session.doc.child(k).nodeSize;
    // The map is read afresh each time: an inserted character moves the cells after it.
    const cellText = (r: number, c: number): string => { const t = session.doc.child(13); const map = TableMap.get(t); return t.nodeAt(map.map[r * map.width + c]!)!.textContent; };
    const inside = pos + 1 + TableMap.get(session.doc.child(13)).map[TableMap.get(session.doc.child(13)).width + 2]! + 2;
    expect(runCommand(session, 'table.breakWords', { words: ['Gamma.Den', 'EtaEastern', 'Zetaforest'] }, { from: inside }).ok).toBe(true);
    expect(cellText(0, 3)).toBe(`Gamma.${OPTIONAL_BREAK}Den`);
    expect(cellText(0, 7)).toBe(`Eta${OPTIONAL_BREAK}Eastern`);
    expect(cellText(0, 6)).toBe('Zetaforest');
    expect(cellText(0, 2)).toBe('Beta.Woods');
    // Again: the broken words read as two words each now, so nothing is added.
    expect(runCommand(session, 'table.breakWords', { words: ['Gamma.Den', 'EtaEastern'] }, { from: inside }).ok).toBe(false);
    expect(cellText(0, 3)).toBe(`Gamma.${OPTIONAL_BREAK}Den`);
    // Outside a table nothing happens.
    expect(runCommand(session, 'table.breakWords', { words: ['Gamma.Den'] }, { from: 1 }).ok).toBe(false);
    // The caret's own.
    expect(runCommand(session, 'range.insertOptionalBreak', undefined, { from: inside }).ok).toBe(true);
    expect(cellText(1, 2)).toContain(OPTIONAL_BREAK);   // the caret's cell, row 1's third
    const saved = await session.save();
    const again = await EditorDocument.open(saved.bytes);
    const body = again.body;
    const tbl = body.content!.filter((e) => e.name.localPart === 'tbl')[4]!.value as wml.Tbl;
    const text = (r: number, c: number): string => textOfElement(cellsOf(rowsOf(tbl)[r]!.element.value)[c]!.element);
    expect(text(0, 3)).toBe(`Gamma.${OPTIONAL_BREAK}Den`);
    expect(text(0, 7)).toBe(`Eta${OPTIONAL_BREAK}Eastern`);
  });
});
