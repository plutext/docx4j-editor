// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 4.9 step B6: the direct and effective formatting readings under the toolbar.
import { describe, test, expect } from 'vitest';
import { EditorDocument, runCommand, stateFor, selectionFormatting, effectiveFormatting, WORD_KEYMAP, commandNamed } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

function block(session: EditorDocument, i: number): { from: number; to: number } {
  const doc = session.doc;
  let pos = 0;
  for (let k = 0; k < i; k++) pos += doc.child(k).nodeSize;
  return { from: pos + 1, to: pos + 1 + doc.child(i).content.size };
}

describe('formatting reads', () => {
  test('a heading reads bold effective and not direct, when its style says so', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const i = 2;   // Heading1
    const state = stateFor(session, block(session, i));
    const direct = selectionFormatting(state);
    const effective = effectiveFormatting(state);
    expect(direct.paragraph.style).toBe('Heading1');
    expect(direct.font.bold).toBe(false);
    const styleRPr = session.pkg.propertyResolver.getEffectiveRPr('Heading1');
    const styleBold = styleRPr?.b !== undefined && styleRPr.b.val !== false;
    expect(effective.font.bold).toBe(styleBold);
    expect(effective.font.size).toBeGreaterThan(0);
    // `name` is '' when the face comes from the theme (`w:asciiTheme`); E1.c's font box resolves that through the theme, as the CSS does.
    expect(effective.paragraph.style).toBe('Heading1');
    expect(effective.paragraph.alignment).not.toBe('Unknown');
  });

  test("'Mixed' across paragraphs of different styles and runs of different formatting", async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const state = stateFor(session, { from: block(session, 0).from, to: block(session, 1).to });
    expect(selectionFormatting(state).paragraph.style).toBe('Mixed');
    expect(effectiveFormatting(state).paragraph.style).toBe('Mixed');
    const { from, to } = block(session, 4);
    runCommand(session, 'font.bold', { value: true }, { from, to: from + 4 });
    const mixed = stateFor(session, { from, to });
    expect(selectionFormatting(mixed).font.bold).toBe('Mixed');
    expect(effectiveFormatting(mixed).font.bold).toBe('Mixed');
    expect(selectionFormatting(stateFor(session, { from, to: from + 4 })).font.bold).toBe(true);
  });

  test('a caret reads the marks at the position; typed text in a bold run reads bold', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const { from, to } = block(session, 4);
    runCommand(session, 'font.bold', { value: true }, { from, to });
    runCommand(session, 'insertText', { text: '!' }, { from: to });
    const caret = stateFor(session, { from: to + 1 });
    expect(selectionFormatting(caret).font.bold).toBe(true);
    expect(effectiveFormatting(caret).font.bold).toBe(true);
    expect(effectiveFormatting(caret).font.size).toBeGreaterThan(0);
  });

  test('an unsaved paragraph change is in the effective reading', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const at = { from: block(session, 4).from };
    expect(effectiveFormatting(stateFor(session, at)).paragraph.alignment).toBe('Left');
    runCommand(session, 'paragraph.alignment', { value: 'Centered' }, at);
    expect(effectiveFormatting(stateFor(session, at)).paragraph.alignment).toBe('Centered');
    runCommand(session, 'paragraph.leftIndent', { value: 36 }, at);
    expect(effectiveFormatting(stateFor(session, at)).paragraph.leftIndent).toBe(36);
    expect(selectionFormatting(stateFor(session, at)).paragraph.leftIndent).toBe(36);
  });

  test('without a resolver the effective reading is the direct one', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'), { resolver: false });
    const state = stateFor(session, block(session, 2));
    expect(effectiveFormatting(state)).toEqual(selectionFormatting(state));
  });
});

describe('the keymap as data', () => {
  test('every binding names a registered command; a key bound twice is a chain, and only Tab and Shift-Tab are (the list level, else the next cell, else a tab)', () => {
    const keys = WORD_KEYMAP.map((b) => b.key);
    const repeated = keys.filter((k, i) => keys.indexOf(k) !== i);
    expect([...new Set(repeated)]).toEqual(['Tab', 'Shift-Tab']);
    expect(WORD_KEYMAP.filter((b) => b.key === 'Tab').map((b) => b.command)).toEqual(['list.indent', 'table.nextCell', 'insertTab']);
    for (const b of WORD_KEYMAP) expect(commandNamed(b.command), b.key).toBeDefined();
  });
});
