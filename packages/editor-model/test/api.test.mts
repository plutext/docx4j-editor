// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 7.2 (section 9 item 20): every command with an `api` line is run twice over a
// fixture, once through the registry and once as its line evaluated inside `runOnTree` against
// the console's `selection` binding, and the two exports must agree part by part. A line whose
// export differs is a finding (in the line or in the engine), recorded in section 11, never
// papered over here.
import { describe, test, expect } from 'vitest';
import { EditorDocument, listCommands, commandNamed, runCommand, stateFor, selectionSpec, selectionBinding, textOffsetIn } from '../src/index.mjs';
import { block, compareLines } from './api-cases.mjs';
import { fixture } from './helpers.mjs';

describe('the portable form of a command (ED-003 section 7.2)', () => {
  test('every api line exports what its command exports', async () => {
    const { findings, checked } = await compareLines(await fixture('tracked-changes.docx'), { known: (c) => c.known });
    expect(checked).toBeGreaterThan(20);
    expect(findings).toEqual([]);
  });

  test('the commands that have no api line are the edit-layer and part-writing ones', () => {
    const without = listCommands().filter((c) => !c.api).map((c) => c.name).sort();
    for (const name of ['history.undo', 'insertParagraph', 'insertTab', 'list.toggle', 'list.indent', 'contentControl.wrapRows', 'contentControl.select', 'table.mergeCells', 'section.orientation', 'paragraph.tabStops', 'document.setChangeTrackingMode']) {
      expect(without).toContain(name);
    }
    for (const name of ['font.bold', 'paragraph.alignment', 'insertText', 'range.hyperlink', 'find.replace', 'find.replaceAll', 'trackedChange.accept', 'trackedChange.reject', 'trackedChanges.acceptAll', 'trackedChanges.rejectAll']) expect(without).not.toContain(name);
  });

  test("every choices entry names an argument the command has, lists values the description mentions, and the palette's select values parse", () => {
    for (const spec of listCommands()) {
      for (const [arg, choices] of Object.entries(spec.choices ?? {})) {
        expect(spec.args?.[arg], `${spec.name}.${arg}`).toBeDefined();
        expect(choices.length, `${spec.name}.${arg}`).toBeGreaterThan(1);
        for (const c of choices) expect(spec.args![arg]!.includes(c) || ['true', 'false', 'toggle'].includes(c) || spec.name === 'font.underline', `${spec.name}.${arg} ${c}`).toBe(true);
      }
    }
    expect(commandNamed('paragraph.alignment')!.choices).toEqual({ value: ['Left', 'Centered', 'Right', 'Justified', 'Unknown'] });
  });

  test('a replace with a group reference is the engine\'s replaceText, or has no portable form for one hit', () => {
    expect(commandNamed('find.replaceAll')!.api!({ text: '(a)', replacement: '$1', matchRegExp: true })).toBe('body.replaceText("(a)", "$1", {"matchRegExp":true});');
    expect(commandNamed('find.replaceAll')!.api!({ text: '(a)', replacement: '$1' })).toMatch(/^const hits = body\.search\("\(a\)"\);/);
    expect(commandNamed('find.replace')!.api!({ text: '(a)', replacement: '\\1', matchWildcards: true })).toBeUndefined();
    expect(commandNamed('find.replaceAll')!.api!({ text: '', replacement: 'x' })).toBeUndefined();
  });

  test('a cleared value has no portable form', () => {
    expect(commandNamed('font.name')!.api!({ value: '' })).toBeUndefined();
    expect(commandNamed('font.size')!.api!({ value: 0 })).toBeUndefined();
    expect(commandNamed('paragraph.style')!.api!({ value: null })).toBeUndefined();
    expect(commandNamed('range.hyperlink')!.api!({ text: 'x', address: 'https://a' })).toBeUndefined();
    expect(commandNamed('range.hyperlink')!.api!({ address: 'https://a', subAddress: 'top' })).toBe('selection.range.hyperlink = "https://a#top";');
  });
});

describe("the console's selection binding (section 9 item 31)", () => {
  test('a caret and a text selection within one paragraph give a range and its font; across paragraphs the range is undefined', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const b4 = block(session, 4);
    const spec = selectionSpec(stateFor(session, { from: b4.from + 5, to: b4.from + 9 }));
    expect(spec.blocks).toEqual([4]);
    expect(spec.range).toEqual({ path: [4], start: 5, end: 9 });
    const bound = await session.runOnTree(({ body }) => {
      const s = selectionBinding(body, spec);
      return { text: s.text, rangeText: s.range?.text, paragraphs: s.paragraphs.length, bold: s.font.bold };
    });
    expect(bound.value).toEqual({ text: 'Some text.', rangeText: 'text', paragraphs: 1, bold: false });
    const across = selectionSpec(stateFor(session, { from: b4.from, to: block(session, 6).to }));
    // Block 6 is empty, and a selection that reaches a paragraph's start selects none of it (ED-005 section 12.81 item 5).
    expect(across.blocks).toEqual([4, 5]);
    expect(across.range).toBeUndefined();
    // The selection as one of the engine's ranges across paragraphs (core-ts 0.6.0, CR-014; ED-008 section 8.8 item 8):
    // from block 4's start to block 5's end, its mark taken, since the selection runs to block 6's start.
    expect(across.span).toEqual({ path: [4], start: 0, endPath: [5], end: block(session, 5).to - block(session, 5).from, holdsMark: true });
    const acrossBound = await session.runOnTree(({ body }) => { const s = selectionBinding(body, across); return { range: s.range, paragraphs: s.paragraphs.length, font: typeof s.font.bold, spans: s.selectionRange?.spansParagraphs, text: s.selectionRange?.text }; });
    expect(acrossBound.value).toEqual({ range: undefined, paragraphs: 2, font: 'boolean', spans: true, text: `Some text.\r${session.doc.child(5).textContent}\r` });
  });

  test("a caret in a table cell selects that cell's paragraph alone, as Office JS does, through a vertical merge's rows too (ED-005 section 12.129)", async () => {
    // Block 7: four rows; row 2's first cell is merged down into row 3, whose w:tr holds the continuation w:tc first, which
    // the editor folds into the merged cell; row 4's second cell spans two columns.
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const table = session.doc.child(7);
    const caretIn = (row: number, cell: number): number => {
      let at = block(session, 7).from;
      for (let r = 0; r < row; r++) at += table.child(r).nodeSize;
      at += 1;   // into the row
      for (let c = 0; c < cell; c++) at += table.child(row).child(c).nodeSize;
      return at + 2;   // into the cell and its first paragraph
    };
    // The editor's cell, and where the engine has it: row 3's cells come after the continuation, which the engine counts.
    const expected: [number, number, number, number][] = [[0, 0, 0, 0], [0, 2, 0, 2], [1, 0, 1, 0], [2, 0, 2, 1], [2, 1, 2, 2], [3, 1, 3, 1]];
    for (const [row, cell, rowIndex, cellIndex] of expected) {
      const spec = selectionSpec(stateFor(session, { from: caretIn(row, cell) }));
      expect(spec.blocks).toEqual([7]);
      const bound = await session.runOnTree(({ body }) => {
        const s = selectionBinding(body, spec);
        return { paragraphs: s.paragraphs.map((p) => [p.parentTableCell?.rowIndex, p.parentTableCell?.cellIndex]), range: s.range ? [s.range.paragraph.parentTableCell?.rowIndex, s.range.paragraph.parentTableCell?.cellIndex] : undefined };
      });
      expect(bound.value, `row ${row}, cell ${cell}`).toEqual({ paragraphs: [[rowIndex, cellIndex]], range: [rowIndex, cellIndex] });
    }
  });

  test('paragraphs inside a block-level content control are in the selection, and the range reaches them (ED-002 section 10.4 item 7 closed)', async () => {
    const session = await EditorDocument.open(await fixture('invoice.docx'));
    // Block 3 is a content control holding "Joe Bloggs".
    const b3 = block(session, 3);
    const spec = selectionSpec(stateFor(session, { from: b3.from + 3 }));
    expect(spec.blocks).toEqual([3]);
    expect(spec.range?.path.length).toBeGreaterThan(1);
    const bound = await session.runOnTree(({ body }) => { const s = selectionBinding(body, spec); return { paragraphs: s.paragraphs.map((p) => p.text), range: s.range?.paragraph.text }; });
    expect(bound.value?.paragraphs.length).toBe(1);
    expect(bound.value?.range).toBe(bound.value?.paragraphs[0]);
  });

  test('text offsets count tabs, breaks and note references as one character (Office JS\'s text, core-ts 0.6.0) and other atoms as none', async () => {
    const session = await EditorDocument.open(await fixture('footnote.docx'));
    const doc = session.doc;
    let found = false;
    doc.forEach((node) => {
      if (found || node.type.name !== 'paragraph') return;
      let hasRef = false;
      node.descendants((c) => { if (c.type.name === 'footnoteRef') hasRef = true; });
      if (!hasRef) return;
      found = true;
      // The offset at the paragraph's end is the text length as the content API counts it.
      expect(textOffsetIn(node, node.content.size)).toBe(node.textBetween(0, node.content.size, '', (leaf) => (leaf.type.name === 'tab' ? '\t' : leaf.type.name === 'hardBreak' ? '\u000B' : leaf.type.name === 'footnoteRef' ? '\u0002' : '')).length);
    });
    expect(found).toBe(true);
  });
});
