// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 4.9, steps B1 to B3: the registry and the headless runner, undo, run formatting
// with grapheme-safe edges and atom fmt, paragraph formatting. Every edit exports through
// `session.save()` (so the self-check runs) and is checked on the reloaded tree.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage, textOf as textOfElement, typeNameOf, runItemsOf, type Element } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { EditorDocument, listCommands, commandNamed, runCommand, stateFor } from '../src/index.mjs';
import { fixture, comparePackages, plain, textOf } from './helpers.mjs';

async function reloadBody(bytes: Uint8Array): Promise<wml.Body> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  return (await pkg.getMainDocumentPart().readContents()).body!;
}

/** Positions of block `i`: its start, the start of its content, the end of its content. */
function block(session: EditorDocument, i: number): { start: number; from: number; to: number } {
  const doc = session.doc;
  let pos = 0;
  for (let k = 0; k < i; k++) pos += doc.child(k).nodeSize;
  return { start: pos, from: pos + 1, to: pos + 1 + doc.child(i).content.size };
}

async function saveAndReload(session: EditorDocument, original: Uint8Array): Promise<wml.Body> {
  const saved = await session.save();
  expect(saved.changed).toBe(true);
  const diffs = comparePackages(original, saved.bytes, [session.main.partName.storeName, 'docProps/app.xml']).filter((d) => d.name !== 'docProps/app.xml');
  expect(diffs).toEqual([]);
  return reloadBody(saved.bytes);
}

const runsOf = (p: Element): Element<wml.R>[] => (runItemsOf(p.value as object) ?? []).filter((e) => typeNameOf(e) === 'org_docx4j_wml.R') as Element<wml.R>[];

function expectOthersEqual(before: Element[], after: Element[], changed: number[]): void {
  expect(after.length).toBe(before.length);
  before.forEach((el, i) => { if (!changed.includes(i)) expect(plain(after[i])).toEqual(plain(el)); });
}

describe('registry and runner (B1)', () => {
  test('every command has a unique content-API name and a label', () => {
    const names = listCommands().map((c) => c.name);
    expect(names.length).toBeGreaterThan(30);
    expect(new Set(names).size).toBe(names.length);
    for (const c of listCommands()) expect(c.label).not.toBe('');
    expect(commandNamed('font.bold')).toBeDefined();
    expect(commandNamed('no.such')).toBeUndefined();
  });

  test('a command runs by name with no view and exports deep-equal', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    const i = 4;   // "Some text."
    const { from } = block(session, i);
    const r = runCommand(session, 'font.bold', { value: true }, { from: from + 5, to: from + 9 });
    expect(r.ok).toBe(true);
    const after = (await saveAndReload(session, original)).content!;
    expectOthersEqual(before, after, [i]);
    const runs = runsOf(after[i]!);
    expect(runs.map((x) => textOfElement(x))).toEqual(['Some ', 'text', '.']);
    expect(runs[1]!.value.rPr?.b).toBeDefined();
    expect(runs[0]!.value.rPr?.b).toBeUndefined();
  });

  test('an unknown command throws; a command that does not apply returns ok false', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    expect(() => runCommand(session, 'no.such')).toThrow(/no command/);
    // decreaseIndent at 0 has nothing to do.
    expect(runCommand(session, 'paragraph.decreaseIndent', undefined, { from: block(session, 4).from }).ok).toBe(false);
  });
});

describe('undo (B1)', () => {
  test('undo restores the export byte for byte; redo brings the edit back; a remote change is not undone', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const { from, to } = block(session, 4);
    expect(session.undoManager.canUndo()).toBe(false);
    runCommand(session, 'font.italic', { value: true }, { from, to });
    expect(session.undoManager.canUndo()).toBe(true);
    expect(session.doc.eq(session.imported)).toBe(false);

    expect(runCommand(session, 'history.undo').ok).toBe(true);
    expect(session.doc.eq(session.imported)).toBe(true);
    const saved = await session.save();
    expect(saved.changed).toBe(false);
    expect(comparePackages(original, saved.bytes)).toEqual([]);

    expect(runCommand(session, 'history.redo').ok).toBe(true);
    expect(session.doc.eq(session.imported)).toBe(false);
    runCommand(session, 'history.undo');

    // Two commands are two steps.
    runCommand(session, 'font.bold', { value: true }, { from, to });
    runCommand(session, 'paragraph.alignment', { value: 'Centered' }, { from });
    expect(session.doc.child(4).attrs['jc']).toBe('center');
    runCommand(session, 'history.undo');
    expect(session.doc.child(4).attrs['jc']).toBeNull();
    expect(session.doc.child(4).firstChild!.marks.some((m) => m.type.name === 'bold')).toBe(true);
    runCommand(session, 'history.undo');
    expect(session.doc.eq(session.imported)).toBe(true);

    // A change with another origin (a peer, the agent) is outside the user's history.
    const state = stateFor(session, { from, to });
    const remote = state.tr.addMark(from, to, session.schema.marks['bold']!.create()).doc;
    session.replace(remote, 'peer');
    expect(session.undoManager.canUndo()).toBe(false);
    expect(runCommand(session, 'history.undo').ok).toBe(false);
    expect(session.doc.child(4).firstChild!.marks.some((m) => m.type.name === 'bold')).toBe(true);
  });
});

describe('run formatting (B2)', () => {
  test.each([
    ['font.underline', { value: 'Single' }, (r: wml.RPr) => r.u?.val === 'single'],
    ['font.underline', { value: 'Double' }, (r: wml.RPr) => r.u?.val === 'double'],
    ['font.strikeThrough', { value: true }, (r: wml.RPr) => r.strike !== undefined && r.dstrike === undefined],
    ['font.doubleStrikeThrough', { value: true }, (r: wml.RPr) => r.dstrike !== undefined],
    ['font.subscript', { value: true }, (r: wml.RPr) => r.vertAlign?.val === 'subscript'],
    ['font.superscript', { value: true }, (r: wml.RPr) => r.vertAlign?.val === 'superscript'],
    ['font.name', { value: 'Georgia' }, (r: wml.RPr) => r.rFonts?.ascii === 'Georgia' && r.rFonts?.hAnsi === 'Georgia'],
    ['font.size', { value: 14 }, (r: wml.RPr) => r.sz?.val === 28 && r.szCs?.val === 28],
    ['font.color', { value: '#FF0000' }, (r: wml.RPr) => r.color?.val === 'FF0000'],
    ['font.highlightColor', { value: '#FFFF00' }, (r: wml.RPr) => r.highlight?.val === 'yellow'],
    ['font.style', { value: 'Strong' }, (r: wml.RPr) => r.rStyle?.val === 'Strong'],
    ['font.rtl', { value: true }, (r: wml.RPr) => r.rtl !== undefined],
  ])('%s %j writes the property on every run of the selection', async (name, args, check) => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    const i = 1;   // two runs
    const { from, to } = block(session, i);
    expect(runCommand(session, name, args, { from, to }).ok).toBe(true);
    const after = (await saveAndReload(session, original)).content!;
    expectOthersEqual(before, after, [i]);
    const runs = runsOf(after[i]!);
    expect(runs.length).toBe(runsOf(before[i]!).length);
    for (const r of runs) expect(check(r.value.rPr ?? {})).toBe(true);
    // Everything else in each run's properties is what it was.
    expect(textOfElement(after[i]!)).toBe(textOfElement(before[i]!));
  });

  test("'toggle' turns a uniform property off, and isActive follows the marks", async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const { from, to } = block(session, 4);
    const bold = commandNamed('font.bold')!;
    expect(bold.isActive!(stateFor(session, { from, to }))).toBe(false);
    runCommand(session, 'font.bold', { value: 'toggle' }, { from, to });
    expect(bold.isActive!(stateFor(session, { from, to }))).toBe(true);
    runCommand(session, 'font.bold', { value: 'toggle' }, { from, to });
    expect(bold.isActive!(stateFor(session, { from, to }))).toBe(false);
    expect(session.doc.eq(session.imported)).toBe(true);
    // Partly bold: not active, and toggle turns it on everywhere.
    runCommand(session, 'font.bold', { value: true }, { from, to: from + 4 });
    expect(bold.isActive!(stateFor(session, { from, to }))).toBe(false);
    runCommand(session, 'font.bold', { value: 'toggle' }, { from, to });
    expect(bold.isActive!(stateFor(session, { from, to }))).toBe(true);
  });

  test('a collapsed selection sets the stored marks and keeps runSource', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const { from } = block(session, 4);
    const r = runCommand(session, 'font.bold', { value: true }, { from: from + 2 });
    expect(r.ok).toBe(true);
    const stored = r.state!.storedMarks!;
    expect(stored.some((m) => m.type.name === 'bold')).toBe(true);
    expect(stored.some((m) => m.type.name === 'runSource')).toBe(true);
    expect(session.doc.eq(session.imported)).toBe(true);   // nothing in the document changed
  });

  test('font.clear removes every formatting mark and nothing else; the run properties follow', async () => {
    const original = await fixture('tracked-changes.docx');
    const first = await EditorDocument.open(original);
    const { from, to } = block(first, 0);   // the Title paragraph
    runCommand(first, 'font.bold', { value: true }, { from, to });
    runCommand(first, 'font.color', { value: '#0000FF' }, { from, to });
    runCommand(first, 'font.clear', undefined, { from, to });
    // Clearing what was just added restores the imported marks exactly: nothing to save.
    expect(first.doc.eq(first.imported)).toBe(true);
    expect(first.doc.child(0).firstChild!.marks.map((m) => m.type.name)).toEqual(['runSource']);
    // Direct formatting that came from the file goes the same way.
    runCommand(first, 'font.bold', { value: true }, { from, to });
    runCommand(first, 'font.color', { value: '#0000FF' }, { from, to });
    const formatted = (await first.save()).bytes;
    const session = await EditorDocument.open(formatted);
    for (const r of runsOf(session.body.content![0]!)) expect(r.value.rPr?.b).toBeDefined();
    runCommand(session, 'font.clear', undefined, block(session, 0));
    const after = (await saveAndReload(session, formatted)).content!;
    for (const r of runsOf(after[0]!)) { expect(r.value.rPr?.b).toBeUndefined(); expect(r.value.rPr?.color).toBeUndefined(); }
    expect(textOfElement(after[0]!)).toBe(textOfElement(session.body.content![0]!));
  });

  test('font.grow and font.shrink step along Word\'s size list from the effective size', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const { from, to } = block(session, 4);
    runCommand(session, 'font.size', { value: 11 }, { from, to });
    runCommand(session, 'font.grow', undefined, { from, to });
    expect(session.doc.child(4).firstChild!.marks.find((m) => m.type.name === 'font')!.attrs['size']).toBe(12);
    runCommand(session, 'font.grow', undefined, { from, to });
    expect(session.doc.child(4).firstChild!.marks.find((m) => m.type.name === 'font')!.attrs['size']).toBe(14);
    runCommand(session, 'font.shrink', undefined, { from, to });
    runCommand(session, 'font.shrink', undefined, { from, to });
    runCommand(session, 'font.shrink', undefined, { from, to });
    expect(session.doc.child(4).firstChild!.marks.find((m) => m.type.name === 'font')!.attrs['size']).toBe(10);
  });

  test('a mark never splits a grapheme: the range grows to the cluster', async () => {
    const original = await fixture('fonts-symbol-and-emoji.docx');
    const session = await EditorDocument.open(original);
    // Find a text node with a surrogate pair and select from inside the pair.
    let target: { pos: number; text: string; offset: number } | undefined;
    session.doc.descendants((node, pos) => {
      if (target || !node.isText) return true;
      const m = /[\uD800-\uDBFF][\uDC00-\uDFFF]/.exec(node.text ?? '');
      if (m) target = { pos, text: node.text!, offset: m.index };
      return true;
    });
    expect(target).toBeDefined();
    const insidePair = target!.pos + target!.offset + 1;   // between the two halves
    runCommand(session, 'font.bold', { value: true }, { from: insidePair, to: insidePair + 3 });
    let boldText = '';
    session.doc.descendants((node) => { if (node.isText && node.marks.some((m) => m.type.name === 'bold')) boldText += node.text; return true; });
    expect(boldText.isWellFormed()).toBe(true);
    expect(boldText.startsWith(target!.text.slice(target!.offset, target!.offset + 2))).toBe(true);
    const after = (await saveAndReload(session, original)).content!;
    for (const p of after) for (const r of runsOf(p)) expect(textOfElement(r).isWellFormed()).toBe(true);
  });

  test('atoms in the range take the formatting: a tab inside a bolded hyperlink is bold too', async () => {
    const original = await fixture('hyperlink_dupe.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    const i = 1;   // "H1", a tab, a PAGEREF field, "1", all inside a hyperlink
    const { from, to } = block(session, i);
    expect(runCommand(session, 'font.bold', { value: true }, { from, to }).ok).toBe(true);
    const after = (await saveAndReload(session, original)).content!;
    expectOthersEqual(before, after, [i]);
    const link = (runItemsOf(after[i]!.value as object) ?? []).find((e) => e.name.localPart === 'hyperlink')!;
    const runs = runsOf(link);
    expect(runs.length).toBe(runsOf((runItemsOf(before[i]!.value as object) ?? []).find((e) => e.name.localPart === 'hyperlink')!).length);
    // Every run with content is bold, the tab's included; a run with no items has nothing to format.
    for (const r of runs) if ((r.value.content ?? []).length > 0) expect(r.value.rPr?.b, `${(r.value.content ?? []).map((c) => c.name.localPart).join('+')} is bold`).toBeDefined();
    expect(runs.some((r) => (r.value.content ?? []).some((c) => c.name.localPart === 'tab') && r.value.rPr?.b !== undefined)).toBe(true);
  });
});

describe('paragraph formatting (B3)', () => {
  test.each([
    ['paragraph.alignment', { value: 'Centered' }, (p: wml.PPr) => p.jc?.val === 'center'],
    ['paragraph.alignment', { value: 'Justified' }, (p: wml.PPr) => p.jc?.val === 'both'],
    ['paragraph.style', { value: 'Heading2' }, (p: wml.PPr) => p.pStyle?.val === 'Heading2'],
    ['paragraph.leftIndent', { value: 36 }, (p: wml.PPr) => p.ind?.left === 720 && p.ind?.start === undefined],
    ['paragraph.rightIndent', { value: 18 }, (p: wml.PPr) => p.ind?.right === 360],
    ['paragraph.firstLineIndent', { value: 18 }, (p: wml.PPr) => p.ind?.firstLine === 360 && p.ind?.hanging === undefined],
    ['paragraph.firstLineIndent', { value: -18 }, (p: wml.PPr) => p.ind?.hanging === 360 && p.ind?.firstLine === undefined],
    ['paragraph.spaceBefore', { value: 6 }, (p: wml.PPr) => p.spacing?.before === 120],
    ['paragraph.spaceAfter', { value: 12 }, (p: wml.PPr) => p.spacing?.after === 240],
    ['paragraph.lineSpacing', { value: 15 }, (p: wml.PPr) => p.spacing?.line === 300 && p.spacing?.lineRule === 'auto'],
    ['paragraph.bidi', { value: true }, (p: wml.PPr) => p.bidi !== undefined],
  ])('%s %j writes w:pPr and leaves the runs alone', async (name, args, check) => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    const i = 4;
    expect(runCommand(session, name, args, { from: block(session, i).from + 2 }).ok).toBe(true);
    const after = (await saveAndReload(session, original)).content!;
    expectOthersEqual(before, after, [i]);
    expect(check((after[i]!.value as wml.P).pPr ?? {})).toBe(true);
    expect(plain(runsOf(after[i]!))).toEqual(plain(runsOf(before[i]!)));
  });

  test('increase and decrease indent step by half an inch and floor at zero', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const at = { from: block(session, 4).from };
    runCommand(session, 'paragraph.increaseIndent', undefined, at);
    runCommand(session, 'paragraph.increaseIndent', undefined, at);
    expect(session.doc.child(4).attrs['ind']).toEqual({ start: 1440, end: null, firstLine: null, hanging: null });
    runCommand(session, 'paragraph.decreaseIndent', undefined, at);
    expect(session.doc.child(4).attrs['ind']).toMatchObject({ start: 720 });
    runCommand(session, 'paragraph.decreaseIndent', undefined, at);
    expect(session.doc.child(4).attrs['ind']).toMatchObject({ start: 0 });
    expect(commandNamed('paragraph.decreaseIndent')!.isEnabled!(stateFor(session, at))).toBe(false);
    const after = (await saveAndReload(session, original)).content!;
    expect((after[4]!.value as wml.P).pPr?.ind?.left).toBe(0);
  });

  test('a paragraph command applies to every paragraph the selection touches; clearFormatting resets a heading', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const range = { from: block(session, 1).from, to: block(session, 2).to };
    runCommand(session, 'paragraph.alignment', { value: 'Right' }, range);
    expect(session.doc.child(1).attrs['jc']).toBe('right');
    expect(session.doc.child(2).attrs['jc']).toBe('right');
    expect(session.doc.child(0).attrs['jc']).toBeNull();
    runCommand(session, 'paragraph.clearFormatting', undefined, { from: block(session, 2).from });
    expect(session.doc.child(2).attrs['pStyle']).toBeNull();
    expect(session.doc.child(2).attrs['jc']).toBeNull();
    const after = (await saveAndReload(session, original)).content!;
    expect((after[2]!.value as wml.P).pPr?.pStyle).toBeUndefined();
    expect(textOf(session.doc.child(2))).toBe(textOfElement(after[2]!));
  });

  test('the alignment isActive reads the direct value', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const at = { from: block(session, 4).from };
    const alignment = commandNamed('paragraph.alignment')!;
    expect(alignment.isActive!(stateFor(session, at), { value: 'Unknown' })).toBe(true);
    runCommand(session, 'paragraph.alignment', { value: 'Centered' }, at);
    expect(alignment.isActive!(stateFor(session, at), { value: 'Centered' })).toBe(true);
    expect(alignment.isActive!(stateFor(session, at), { value: 'Left' })).toBe(false);
  });
});

describe('exportBlock (reveal codes)', () => {
  test('an unchanged block gives its original element; a changed one the rebuilt element, and the tree is untouched', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    expect(session.exportBlock(session.doc.child(4))).toBe(before[4]);
    const { from, to } = block(session, 4);
    runCommand(session, 'font.bold', { value: true }, { from, to });
    const rebuilt = session.exportBlock(session.doc.child(4));
    expect(rebuilt).not.toBe(before[4]);
    expect(runsOf(rebuilt)[0]!.value.rPr?.b).toBeDefined();
    expect(session.body.content![4]).toBe(before[4]);
    const saved = await session.save();
    expect(saved.changed).toBe(true);
  });
});
