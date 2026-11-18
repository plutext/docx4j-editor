// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 5.8 items 1 and 2, the attribute half (step C1): `numPr` in the editable subset,
// and attach, detach, level, indent, outdent and continue through the headless runner with the
// export reloaded. The untouched round trip of every fixture is roundtrip.test.mts; here each
// command saves through `session.save()` (so the self-check runs) with only the main part changed.
import { describe, test, expect } from 'vitest';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, stateFor, commandNamed, paragraphAttrsOf, applyParagraphAttrs, clampLevel, numberingOf, listRefOf, WORD_KEYMAP, type ParagraphAttrs } from '../src/index.mjs';
import { fixture, comparePackages, plain } from './helpers.mjs';

type Found = { node: PMNode; pos: number; from: number };

/** The nth paragraph (in story order) matching `pred`, with the position of its content. */
function paragraph(doc: PMNode, pred: (attrs: ParagraphAttrs, node: PMNode) => boolean, n = 0): Found {
  let found: Found | undefined;
  let seen = 0;
  doc.descendants((node, pos) => {
    if (found || node.type.name !== 'paragraph') return !found;
    if (pred(node.attrs as ParagraphAttrs, node) && seen++ === n) found = { node, pos, from: pos + 1 };
    return false;
  });
  if (!found) throw new Error('no such paragraph');
  return found;
}

const numPrOf = (doc: PMNode, pos: number): ParagraphAttrs['numPr'] => (doc.nodeAt(pos)!.attrs as ParagraphAttrs).numPr;
/** How often each paragraph text occurs in the document. */
function textCounts(doc: PMNode): Map<string, number> {
  const out = new Map<string, number>();
  doc.descendants((node) => {
    if (node.type.name !== 'paragraph') return true;
    out.set(node.textContent, (out.get(node.textContent) ?? 0) + 1);
    return false;
  });
  return out;
}

/** A paragraph in no list, with text unique in the document, so that it is found again by it after a reload. */
function plainParagraph(doc: PMNode): (a: ParagraphAttrs, node: PMNode) => boolean {
  const counts = textCounts(doc);
  return (a, node) => a.numPr === null && a.pStyle === null && a.sourceId !== null && node.textContent.length > 20 && counts.get(node.textContent) === 1;
}

/**
 * Save (self-check included), require that only the main part changed, and hand back the saved
 * document reopened: the same positions, since no command here adds or removes a block, and
 * `numPr` read from the written `w:numPr` by the projection.
 */
async function reload(session: EditorDocument, original: Uint8Array): Promise<PMNode> {
  const saved = await session.save();
  expect(saved.changed).toBe(true);
  const diffs = comparePackages(original, saved.bytes, [session.main.partName.storeName, 'docProps/app.xml']);
  expect(diffs).toEqual([]);
  return (await EditorDocument.open(saved.bytes)).doc;
}

describe('numPr in the editable subset (5.2)', () => {
  test('the paragraph attributes carry the paragraph\'s own w:numPr, numId as a string', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const numbered: ParagraphAttrs['numPr'][] = [];
    session.doc.descendants((node) => {
      if (node.type.name !== 'paragraph') return true;
      const numPr = (node.attrs as ParagraphAttrs).numPr;
      if (numPr) numbered.push(numPr);
      return false;
    });
    expect(numbered).toEqual([
      { numId: '1', ilvl: 0 }, { numId: '1', ilvl: 1 },
      { numId: '2', ilvl: 0 }, { numId: '2', ilvl: 1 }, { numId: '2', ilvl: 2 },
    ]);
  });

  test('a w:numId of 0 is the string \'0\'; a level-only w:numPr has numId null and writes back without one', () => {
    expect(paragraphAttrsOf({ numPr: { numId: { val: 0 }, ilvl: { val: 0 } } }).numPr).toEqual({ numId: '0', ilvl: 0 });
    expect(paragraphAttrsOf({ numPr: { ilvl: { val: 1 } } }).numPr).toEqual({ numId: null, ilvl: 1 });
    expect(paragraphAttrsOf({ numPr: { numId: { val: 3 } } }).numPr).toEqual({ numId: '3', ilvl: 0 });
    const base = paragraphAttrsOf(undefined);
    const written = applyParagraphAttrs(undefined, { ...base, sourceId: null, numPr: { numId: null, ilvl: 1 } }, base);
    expect(plain(written?.numPr)).toEqual({ TYPE_NAME: 'org_docx4j_wml.PPrBase.NumPr', ilvl: { val: 1 } });
  });

  test('a changed numPr replaces the element and keeps its other children; an unchanged one leaves w:pPr alone', () => {
    const original: wml.PPr = { numPr: { ilvl: { val: 0 }, numId: { val: 1 }, ins: { id: 7, author: 'x' } as wml.CTTrackChange }, jc: { val: 'center' } };
    const base = paragraphAttrsOf(original);
    const same = applyParagraphAttrs(structuredClone(original), { ...base, sourceId: null }, base);
    expect(plain(same)).toEqual(plain(original));
    const changed = applyParagraphAttrs(structuredClone(original), { ...base, sourceId: null, numPr: { numId: '5', ilvl: 2 } }, base);
    expect(plain(changed?.numPr)).toEqual({ ilvl: { val: 2 }, numId: { val: 5 }, ins: { id: 7, author: 'x' } });
    const removed = applyParagraphAttrs(structuredClone(original), { ...base, sourceId: null, numPr: null }, base);
    expect(removed?.numPr).toBeUndefined();
    expect(removed?.jc?.val).toBe('center');
  });

  test('clampLevel keeps to the levels the list defines', async () => {
    const session = await EditorDocument.open(await fixture('headers-footers.docx'));
    const { definitions } = numberingOf(stateFor(session))!;
    expect(clampLevel('26', 4, definitions)).toBe(4);       // abstractNum 13: nine levels
    expect(clampLevel('26', 12, definitions)).toBe(8);
    expect(clampLevel('26', -3, definitions)).toBe(0);
    expect(clampLevel('2', 3, definitions)).toBe(0);        // abstractNum 4: one level
    expect(clampLevel('99', 3, definitions)).toBe(3);       // unknown list: the range alone
    expect(clampLevel('26', 4, undefined)).toBe(4);
  });
});

describe('attach and detach (5.3)', () => {
  test('attachToList writes numId and ilvl on every selected paragraph; an unknown list is refused', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const a = paragraph(session.doc, plainParagraph(session.doc), 0);
    const b = paragraph(session.doc, plainParagraph(session.doc), 1);
    expect(runCommand(session, 'paragraph.attachToList', { listId: 99 }, { from: a.from }).ok).toBe(false);
    // Into b's text: a selection ending at b's start touches b and no further, and selects none of it (ED-005 section 12.81 item 5).
    expect(runCommand(session, 'paragraph.attachToList', { listId: 2, level: 1 }, { from: a.from, to: b.from + 1 }).ok).toBe(true);
    expect(numPrOf(session.doc, a.pos)).toEqual({ numId: '2', ilvl: 1 });
    expect(numPrOf(session.doc, b.pos)).toEqual({ numId: '2', ilvl: 1 });
    // The level is clamped to the list's levels.
    runCommand(session, 'paragraph.attachToList', { listId: '1', level: 12 }, { from: a.from });
    expect(numPrOf(session.doc, a.pos)).toEqual({ numId: '1', ilvl: 8 });
    const after = await reload(session, original);
    expect(numPrOf(after, a.pos)).toEqual({ numId: '1', ilvl: 8 });
    expect(numPrOf(after, b.pos)).toEqual({ numId: '2', ilvl: 1 });
  });

  test('detachFromList removes the paragraph\'s numPr, and writes numId 0 where the style numbers it', async () => {
    const original = await fixture('headers-footers.docx');
    const session = await EditorDocument.open(original);
    const direct = paragraph(session.doc, (a) => a.numPr?.numId === '26');
    const styled = paragraph(session.doc, (a) => a.pStyle === 'ListBullet' && a.numPr === null);
    const state = stateFor(session, { from: styled.from });
    expect(listRefOf(styled.node.attrs as ParagraphAttrs, numberingOf(state)!.emulator)).toEqual({ numId: '2', ilvl: 0 });
    expect(commandNamed('paragraph.detachFromList')!.isEnabled!(state)).toBe(true);
    expect(runCommand(session, 'paragraph.detachFromList', undefined, { from: direct.from }).ok).toBe(true);
    expect(numPrOf(session.doc, direct.pos)).toBeNull();
    expect(runCommand(session, 'paragraph.detachFromList', undefined, { from: styled.from }).ok).toBe(true);
    expect(numPrOf(session.doc, styled.pos)).toEqual({ numId: '0', ilvl: 0 });
    expect(listRefOf(session.doc.nodeAt(styled.pos)!.attrs as ParagraphAttrs, numberingOf(state)!.emulator)).toBeUndefined();
    // A paragraph that is in no list is not a target.
    const plainOne = paragraph(session.doc, plainParagraph(session.doc));
    expect(runCommand(session, 'paragraph.detachFromList', undefined, { from: plainOne.from }).ok).toBe(false);
    const after = await reload(session, original);
    expect(numPrOf(after, direct.pos)).toBeNull();
    expect(numPrOf(after, styled.pos)).toEqual({ numId: '0', ilvl: 0 });
  });

  test('each command is one undo step and undo restores the imported document', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const item = paragraph(session.doc, (a) => a.numPr?.numId === '1' && a.numPr.ilvl === 0);
    runCommand(session, 'list.indent', undefined, { from: item.from });
    runCommand(session, 'paragraph.detachFromList', undefined, { from: item.from });
    expect(numPrOf(session.doc, item.pos)).toBeNull();
    runCommand(session, 'history.undo');
    expect(numPrOf(session.doc, item.pos)).toEqual({ numId: '1', ilvl: 1 });
    runCommand(session, 'history.undo');
    expect(session.doc.eq(session.imported)).toBe(true);
    const saved = await session.save();
    expect(saved.changed).toBe(false);
    expect(comparePackages(original, saved.bytes)).toEqual([]);
  });
});

describe('the level (5.3)', () => {
  test('listItem.level sets the level, clamped; on a style-numbered paragraph it writes the style\'s numId beside it', async () => {
    const original = await fixture('headers-footers.docx');
    const session = await EditorDocument.open(original);
    const direct = paragraph(session.doc, (a) => a.numPr?.numId === '26');
    const styled = paragraph(session.doc, (a) => a.pStyle === 'ListBullet' && a.numPr === null);
    const plainOne = paragraph(session.doc, plainParagraph(session.doc));
    expect(runCommand(session, 'listItem.level', { value: 4 }, { from: direct.from }).ok).toBe(true);
    expect(numPrOf(session.doc, direct.pos)).toEqual({ numId: '26', ilvl: 4 });
    runCommand(session, 'listItem.level', { value: 12 }, { from: direct.from });
    expect(numPrOf(session.doc, direct.pos)).toEqual({ numId: '26', ilvl: 8 });
    expect(commandNamed('listItem.level')!.isActive!(stateFor(session, { from: direct.from }), { value: 8 })).toBe(true);
    // ListBullet's list (numId 2) defines one level: the level clamps to where the paragraph resolves already, and nothing is written.
    expect(runCommand(session, 'listItem.level', { value: 1 }, { from: styled.from }).ok).toBe(true);
    expect(numPrOf(session.doc, styled.pos)).toBeNull();
    expect(runCommand(session, 'listItem.level', { value: 1 }, { from: plainOne.from }).ok).toBe(false);
    // Appendix1 numbers through w:num 8 (nine levels): the level change writes the style's numId beside the level.
    runCommand(session, 'paragraph.style', { value: 'Appendix1' }, { from: plainOne.from });
    expect(listRefOf(session.doc.nodeAt(plainOne.pos)!.attrs as ParagraphAttrs, numberingOf(stateFor(session))!.emulator)).toEqual({ numId: '8', ilvl: 0 });
    expect(runCommand(session, 'listItem.level', { value: 2 }, { from: plainOne.from }).ok).toBe(true);
    expect(numPrOf(session.doc, plainOne.pos)).toEqual({ numId: '8', ilvl: 2 });
    const after = await reload(session, original);
    expect(numPrOf(after, direct.pos)).toEqual({ numId: '26', ilvl: 8 });
    expect(numPrOf(after, styled.pos)).toBeNull();
    expect(numPrOf(after, plainOne.pos)).toEqual({ numId: '8', ilvl: 2 });
  });

  test('indent and outdent step the level; outdent at level 0 detaches', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const item = paragraph(session.doc, (a) => a.numPr?.numId === '2' && a.numPr.ilvl === 0);
    const at = { from: item.from };
    expect(runCommand(session, 'list.indent', undefined, at).ok).toBe(true);
    expect(numPrOf(session.doc, item.pos)).toEqual({ numId: '2', ilvl: 1 });
    runCommand(session, 'list.indent', undefined, at);
    expect(numPrOf(session.doc, item.pos)).toEqual({ numId: '2', ilvl: 2 });
    runCommand(session, 'list.outdent', undefined, at);
    runCommand(session, 'list.outdent', undefined, at);
    expect(numPrOf(session.doc, item.pos)).toEqual({ numId: '2', ilvl: 0 });
    // Level 0: outdent detaches (ListParagraph numbers nothing by itself), and the paragraph is no list item.
    expect(runCommand(session, 'list.outdent', undefined, at).ok).toBe(true);
    expect(numPrOf(session.doc, item.pos)).toBeNull();
    expect(runCommand(session, 'list.outdent', undefined, at).ok).toBe(false);
    expect(runCommand(session, 'list.indent', undefined, at).ok).toBe(false);
    expect(commandNamed('list.indent')!.isEnabled!(stateFor(session, at))).toBe(false);
    const after = await reload(session, original);
    expect(numPrOf(after, item.pos)).toBeNull();
  });

  test('outdent at level 0 of a style-numbered paragraph switches the list off with numId 0', async () => {
    const session = await EditorDocument.open(await fixture('headers-footers.docx'));
    const styled = paragraph(session.doc, (a) => a.pStyle === 'ListBullet' && a.numPr === null);
    expect(runCommand(session, 'list.outdent', undefined, { from: styled.from }).ok).toBe(true);
    expect(numPrOf(session.doc, styled.pos)).toEqual({ numId: '0', ilvl: 0 });
  });

  test('at the top level, indent is handled with nothing to do rather than falling through', async () => {
    const session = await EditorDocument.open(await fixture('headers-footers.docx'));
    const styled = paragraph(session.doc, (a) => a.pStyle === 'ListBullet' && a.numPr === null);
    // numId 2 defines one level: there is no level 1 to go to; the key is still consumed.
    expect(runCommand(session, 'list.indent', undefined, { from: styled.from }).ok).toBe(true);
    expect(numPrOf(session.doc, styled.pos)).toBeNull();
  });

  test('with atStart, Tab and Shift+Tab apply only at the start of a list item, and the keymap binds them first', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const item = paragraph(session.doc, (a) => a.numPr?.numId === '2' && a.numPr.ilvl === 1);
    const plainOne = paragraph(session.doc, plainParagraph(session.doc));
    expect(runCommand(session, 'list.indent', { atStart: true }, { from: item.from + 2 }).ok).toBe(false);
    expect(runCommand(session, 'list.indent', { atStart: true }, { from: item.from, to: item.from + 2 }).ok).toBe(false);
    expect(runCommand(session, 'list.indent', { atStart: true }, { from: plainOne.from }).ok).toBe(false);
    expect(runCommand(session, 'list.indent', { atStart: true }, { from: item.from }).ok).toBe(true);
    expect(numPrOf(session.doc, item.pos)).toEqual({ numId: '2', ilvl: 2 });
    expect(runCommand(session, 'list.outdent', { atStart: true }, { from: item.from }).ok).toBe(true);
    expect(numPrOf(session.doc, item.pos)).toEqual({ numId: '2', ilvl: 1 });
    const tabs = WORD_KEYMAP.filter((b) => b.key === 'Tab').map((b) => b.command);
    expect(tabs).toEqual(['list.indent', 'table.nextCell', 'insertTab']);
    expect(WORD_KEYMAP.filter((b) => b.key === 'Shift-Tab').map((b) => b.command)).toEqual(['list.outdent', 'table.previousCell']);
  });
});

describe('continue (5.3)', () => {
  test('the caret\'s item and the later items of its w:num move to the nearest preceding list of the same abstract definition', async () => {
    const original = await fixture('headers-footers.docx');
    const session = await EditorDocument.open(original);
    // w:num 21 and 22 share w:abstractNum 12 (22 restarts it); 26 is another definition.
    const [a, b, c, d] = [0, 1, 2, 3].map((n) => paragraph(session.doc, plainParagraph(session.doc), n));
    runCommand(session, 'paragraph.attachToList', { listId: 21 }, { from: a!.from });
    runCommand(session, 'paragraph.attachToList', { listId: 22, level: 1 }, { from: b!.from });
    runCommand(session, 'paragraph.attachToList', { listId: 22 }, { from: c!.from });
    runCommand(session, 'paragraph.attachToList', { listId: 26 }, { from: d!.from });
    // Nothing precedes a's list: refused.
    expect(runCommand(session, 'list.continue', undefined, { from: a!.from }).ok).toBe(false);
    expect(runCommand(session, 'list.continue', undefined, { from: d!.from }).ok).toBe(false);
    expect(runCommand(session, 'list.continue', undefined, { from: b!.from }).ok).toBe(true);
    expect(numPrOf(session.doc, b!.pos)).toEqual({ numId: '21', ilvl: 1 });
    expect(numPrOf(session.doc, c!.pos)).toEqual({ numId: '21', ilvl: 0 });
    expect(numPrOf(session.doc, d!.pos)).toEqual({ numId: '26', ilvl: 0 });
    // Already continuing: nothing to do.
    expect(runCommand(session, 'list.continue', undefined, { from: c!.from }).ok).toBe(false);
    const after = await reload(session, original);
    expect(numPrOf(after, c!.pos)).toEqual({ numId: '21', ilvl: 0 });
  });
});

describe('the rest of the paragraph commands with numPr', () => {
  test('clearFormatting removes the paragraph\'s numPr; Enter carries it to the new paragraph', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const item = paragraph(session.doc, (a) => a.numPr?.numId === '1' && a.numPr.ilvl === 1);
    runCommand(session, 'insertParagraph', undefined, { from: item.from + 1 });
    expect(numPrOf(session.doc, item.pos)).toEqual({ numId: '1', ilvl: 1 });
    const next = item.pos + session.doc.nodeAt(item.pos)!.nodeSize;
    expect(numPrOf(session.doc, next)).toEqual({ numId: '1', ilvl: 1 });
    runCommand(session, 'paragraph.clearFormatting', undefined, { from: item.from });
    expect(numPrOf(session.doc, item.pos)).toBeNull();
  });
});
