// SPDX-License-Identifier: GPL-3.0-or-later
//
// What a formatting or properties change changed, with its values (ED-005 proposal 43, section 9.7):
// `changedValues` over the Word-made fixtures' changes and `review-for-word.docx`'s, and over changes
// the commands make with tracking on, a property taken away included.
import { describe, test, expect } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, trackedChangesOf, changedValues, whatChanged } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

/** Each formatting and properties change: its text and its values. */
function formatted(doc: PMNode): [string, Record<string, unknown>][] {
  return trackedChangesOf(doc).filter((c) => c.type === 'Formatted').map((c) => [c.text.trim(), Object.fromEntries(changedValues(doc, c).map((v) => [v.property, v.value]))]);
}

const LIST = { numPr: { numId: '4', ilvl: 0 }, 'ind.start': null, 'ind.firstLine': null };

describe('changedValues over the Word-made files', () => {
  test.each(['revisions/revisions-word2010.docx', 'revisions/revisions-word15.docx'])('%s', async (name) => {
    const doc = (await EditorDocument.open(await fixture(name))).doc;
    expect(formatted(doc)).toEqual([
      ['alpha', { bold: true }],
      ['bravo', { 'font.color': '#FF0000' }],
      ['charlie', { italic: true, underline: 'Single' }],
      ['', { jc: 'center' }],
      ['', { 'ind.start': 720 }],
      ['', { pStyle: 'Heading2' }],
      // Word took the paragraph's own indents away as it made the list, which states its own.
      ['', LIST], ['', LIST], ['', LIST],
    ]);
  });

  test.each([
    ['review/review-for-word.docx', '1271868263'],
    ['review/review-for-word-word15.docx', '4'],
  ])('%s', async (name, numId) => {
    const doc = (await EditorDocument.open(await fixture(name))).doc;
    expect(formatted(doc)).toEqual([
      ['alpha', { bold: true }],
      ['bravo', { 'font.size': 16 }],
      ['', { jc: 'center' }],
      ['', { 'ind.start': 720 }],
      ['', { pStyle: 'Heading2' }],
      ['', { numPr: { numId, ilvl: 0 } }],
      // The link's text takes the Hyperlink character style (core-ts 0.2.1).
      ['docx4j site', { runStyle: 'Hyperlink' }],
    ]);
  });

  test('the properties are whatChanged\'s, in its order', async () => {
    const doc = (await EditorDocument.open(await fixture('revisions/revisions-word15.docx'))).doc;
    for (const c of trackedChangesOf(doc)) expect(changedValues(doc, c).map((v) => v.property)).toEqual(whatChanged(doc, c));
  });
});

describe('changedValues over the commands\' changes', () => {
  async function tracked(): Promise<{ session: EditorDocument; from: number; to: number }> {
    const session = await EditorDocument.open(await fixture('revisions/revisions-base.docx'));
    let from = -1;
    let to = -1;
    // "quick" in a Normal paragraph (the title's style is bold of its own).
    session.doc.descendants((n, pos) => {
      if (from < 0 && n.isText && n.text!.includes('The quick brown')) { from = pos + n.text!.indexOf('quick'); to = from + 5; }
      return from < 0;
    });
    // Formatting stated before tracking starts, so that taking it away is a change.
    runCommand(session, 'font.bold', { value: true }, { from, to });
    runCommand(session, 'font.size', { value: 14 }, { from, to });
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author C';
    return { session, from, to };
  }

  test('run formatting: a property taken away is null, a value its own', async () => {
    const { session, from, to } = await tracked();
    runCommand(session, 'font.bold', { value: false }, { from, to });
    runCommand(session, 'font.size', { value: 0 }, { from, to });
    runCommand(session, 'font.highlightColor', { value: 'yellow' }, { from, to });
    runCommand(session, 'font.doubleStrikeThrough', { value: true }, { from, to });
    runCommand(session, 'font.subscript', { value: true }, { from, to });
    expect(formatted(session.doc)).toEqual([['quick', { bold: null, strike: 'double', vertAlign: 'subscript', 'font.size': null, 'font.highlight': '#FFFF00' }]]);
  });

  test('a paragraph\'s properties: a hanging indent negative, a line spacing with its rule', async () => {
    const { session, from, to } = await tracked();
    runCommand(session, 'paragraph.firstLineIndent', { value: -18 }, { from, to });
    runCommand(session, 'paragraph.lineSpacing', { value: 12 }, { from, to });
    runCommand(session, 'paragraph.spaceAfter', { value: 6 }, { from, to });
    const values = formatted(session.doc).find(([text]) => text === '')?.[1];
    expect(values).toEqual({ 'ind.firstLine': -360, 'spacing.after': 120, 'spacing.line': { line: 240, rule: 'auto' } });
  });
});
