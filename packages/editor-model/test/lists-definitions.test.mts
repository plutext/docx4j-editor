// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 5.8 item 2, the definition half (step C2): a list started in a document with no
// numbering part gets the part and its relationship and the paragraph its w:numPr; restart (a new
// w:num with a w:startOverride, the later paragraphs moved); toggle in each of its three cases;
// Enter on an empty item and Backspace at the start of one; undo per command with the part's
// write kept and said so.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, runCommandAsync, stateFor, commandNamed, toggleCase, liveLabels, listLabels, type ParagraphAttrs } from '../src/index.mjs';
import { fixture, comparePackages } from './helpers.mjs';

type Found = { node: PMNode; pos: number; from: number };

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
const plain = (a: ParagraphAttrs, node: PMNode): boolean => a.numPr === null && a.pStyle === null && a.sourceId !== null && node.textContent.length > 0;

/** The numbering part of saved bytes. */
async function numberingOf(bytes: Uint8Array): Promise<wml.Numbering> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  return pkg.getMainDocumentPart().numberingDefinitionsPart!.readContents();
}

// A selection to a paragraph's start touches that paragraph and no further (ED-005 section 12.81 item 5): the selections
// here reach one position into their last paragraph.
describe('start (5.3)', () => {
  test('HelloWordOnline.docx has no numbering part: a list started gets the part and its relationship, and the paragraphs their numPr; undo keeps the part write and says so', async () => {
    const original = await fixture('HelloWordOnline.docx');
    const session = await EditorDocument.open(original);
    expect(session.main.numberingDefinitionsPart).toBeUndefined();
    expect(() => runCommand(session, 'list.start', { bullet: true })).toThrow(/runCommandAsync/);
    const a = paragraph(session.doc, plain, 0);
    const result = await runCommandAsync(session, 'list.start', { bullet: true }, { from: a.from });
    expect(result.ok).toBe(true);
    const numPr = numPrOf(session.doc, a.pos)!;
    expect(numPr.ilvl).toBe(0);
    expect(session.main.numberingDefinitionsPart).toBeDefined();
    expect(session.touchedParts.has('/word/numbering.xml')).toBe(true);
    // Live: the label is a bullet at once (one document held: `session.doc` builds a fresh tree each time).
    const doc = session.doc;
    const label = liveLabels(doc, session).get(doc.nodeAt(a.pos)!);
    expect(label?.isBullet).toBe(true);
    expect(label?.listString).toBe('•');   // docx4j's default bullet is Symbol's F0B7, shown as the Unicode bullet
    const saved = await session.save();
    const diffs = comparePackages(original, saved.bytes, [session.main.partName.storeName, 'docProps/app.xml']);
    expect(diffs).toEqual([{ name: 'word/numbering.xml', reason: 'added' }]);
    const numbering = await numberingOf(saved.bytes);
    expect(numbering.num!.some((n) => String(n.numId) === numPr.numId)).toBe(true);
    const reopened = await EditorDocument.open(saved.bytes);
    expect(listLabels(reopened).size).toBe(1);
    // Undo takes the numPr off; the part stays, and the session says it was touched.
    runCommand(session, 'history.undo');
    expect(numPrOf(session.doc, a.pos)).toBeNull();
    expect(session.doc.eq(session.imported)).toBe(true);
    expect(session.touchedParts.has('/word/numbering.xml')).toBe(true);
    expect((await session.save()).changed).toBe(true);
  });

  test('a numbered list over several paragraphs at a level; a second start of the same kind reuses the unreferenced w:num', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const a = paragraph(session.doc, plain, 0);
    const b = paragraph(session.doc, plain, 1);
    const before = session.main.numberingDefinitionsPart!.definitions.instanceListDefinitions.size;
    expect((await runCommandAsync(session, 'list.start', { level: 1 }, { from: a.from, to: b.from + 1 })).ok).toBe(true);
    const first = numPrOf(session.doc, a.pos)!;
    expect(numPrOf(session.doc, b.pos)).toEqual(first);
    expect(first.ilvl).toBe(1);
    expect(session.main.numberingDefinitionsPart!.definitions.instanceListDefinitions.size).toBe(before + 1);
    const doc = session.doc;
    expect(liveLabels(doc, session).get(doc.nodeAt(b.pos)!)?.isBullet).toBe(false);
    // Off, then on again: the same w:num, since nothing names it any more.
    runCommand(session, 'paragraph.detachFromList', undefined, { from: a.from, to: b.from + 1 });
    await runCommandAsync(session, 'list.start', {}, { from: a.from });
    expect(numPrOf(session.doc, a.pos)).toEqual({ numId: first.numId, ilvl: 0 });
    expect(session.main.numberingDefinitionsPart!.definitions.instanceListDefinitions.size).toBe(before + 1);
    // While it is in use, a new start makes another.
    await runCommandAsync(session, 'list.start', {}, { from: b.from });
    expect(numPrOf(session.doc, b.pos)!.numId).not.toBe(first.numId);
    expect(session.main.numberingDefinitionsPart!.definitions.instanceListDefinitions.size).toBe(before + 2);
    const saved = await session.save();
    expect(comparePackages(original, saved.bytes, [session.main.partName.storeName, 'docProps/app.xml', 'word/numbering.xml'])).toEqual([]);
  });
});

describe('restart (5.3)', () => {
  test('a new w:num over the same definition with a w:startOverride; this item and the later items move to it, and count from 1 again', async () => {
    const original = await fixture('numbering-stories.docx');
    const session = await EditorDocument.open(original);
    const third = paragraph(session.doc, (a) => a.numPr?.numId === '70', 2);
    expect((await runCommandAsync(session, 'list.restart', undefined, { from: third.from })).ok).toBe(true);
    const restarted = numPrOf(session.doc, third.pos)!.numId!;
    expect(restarted).not.toBe('70');
    const labels = liveLabels(session.doc, session);
    const strings = [...labels.values()].map((l) => l.listString);
    expect(strings).toEqual(['1.', '2.', '1.', '2.', '3.', '4.']);
    const saved = await session.save();
    expect(comparePackages(original, saved.bytes, [session.main.partName.storeName, 'docProps/app.xml', 'word/numbering.xml'])).toEqual([]);
    const numbering = await numberingOf(saved.bytes);
    const num = numbering.num!.find((n) => String(n.numId) === restarted)!;
    expect(String(num.abstractNumId.val)).toBe('70');
    expect(num.lvlOverride?.[0]?.startOverride?.val).toBe(1);
    const reopened = await EditorDocument.open(saved.bytes);
    expect([...listLabels(reopened).values()].map((l) => l.listString)).toEqual(strings);
    // Not a list item: refused, and nothing written.
    const touched = session.touchedParts.size;
    const plainOne = paragraph(session.doc, plain);
    expect((await runCommandAsync(session, 'list.restart', undefined, { from: plainOne.from })).ok).toBe(false);
    expect(session.touchedParts.size).toBe(touched);
  });
});

describe('toggle (5.3)', () => {
  test('its three cases: start, attach and continue the list ending just before, detach', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const a = paragraph(session.doc, plain, 0);
    const b = paragraph(session.doc, plain, 1);
    // a and b are consecutive plain paragraphs? Find b as the paragraph right after a.
    const after = session.doc.resolve(a.pos + a.node.nodeSize).nodeAfter;
    const next = after && after.type.name === 'paragraph' ? { pos: a.pos + a.node.nodeSize, from: a.pos + a.node.nodeSize + 1 } : b;
    expect(toggleCase(stateFor(session, { from: a.from }), true).kind).toBe('start');
    expect((await runCommandAsync(session, 'list.toggle', { bullet: true }, { from: a.from })).ok).toBe(true);
    const numId = numPrOf(session.doc, a.pos)!.numId!;
    expect(commandNamed('list.toggle')!.isActive!(stateFor(session, { from: a.from }), { bullet: true })).toBe(true);
    expect(commandNamed('list.toggle')!.isActive!(stateFor(session, { from: a.from }), { bullet: false })).toBe(false);
    // The next paragraph: a bulleted list ends just before it, so it attaches and continues.
    expect(toggleCase(stateFor(session, { from: next.from }), true)).toEqual({ kind: 'attach', numId, ilvl: 0 });
    expect((await runCommandAsync(session, 'list.toggle', { bullet: true }, { from: next.from })).ok).toBe(true);
    expect(numPrOf(session.doc, next.pos)).toEqual({ numId, ilvl: 0 });
    // Numbering on a bulleted item: a numbered list starts (a different kind).
    expect(toggleCase(stateFor(session, { from: next.from }), false).kind).toBe('start');
    // Bullets on both bulleted items: detach.
    expect(toggleCase(stateFor(session, { from: a.from, to: next.from + 1 }), true).kind).toBe('detach');
    expect((await runCommandAsync(session, 'list.toggle', { bullet: true }, { from: a.from, to: next.from + 1 })).ok).toBe(true);
    expect(numPrOf(session.doc, a.pos)).toBeNull();
    expect(numPrOf(session.doc, next.pos)).toBeNull();
  });
});

describe('Enter and Backspace on a list item (5.3)', () => {
  test('Enter carries the list to the new item; Enter on an empty item ends the list; Backspace at the start of an item detaches it first', async () => {
    const session = await EditorDocument.open(await fixture('numbering-stories.docx'));
    const first = paragraph(session.doc, (a) => a.numPr?.numId === '70');
    const end = first.from + first.node.content.size;
    runCommand(session, 'insertParagraph', undefined, { from: end });
    const made = first.pos + session.doc.nodeAt(first.pos)!.nodeSize;
    expect(numPrOf(session.doc, made)).toEqual({ numId: '70', ilvl: 0 });
    expect(session.doc.nodeAt(made)!.content.size).toBe(0);
    // Enter on the empty item: no new paragraph, the item leaves the list.
    const count = session.doc.childCount;
    runCommand(session, 'insertParagraph', undefined, { from: made + 1 });
    expect(session.doc.childCount).toBe(count);
    expect(numPrOf(session.doc, made)).toBeNull();
    // Backspace at the start of the next item: it leaves the list; the paragraphs are not joined.
    const item = paragraph(session.doc, (a) => a.numPr?.numId === '70', 1);
    runCommand(session, 'joinBackward', undefined, { from: item.from });
    expect(session.doc.childCount).toBe(count);
    expect(numPrOf(session.doc, item.pos)).toBeNull();
    // A second Backspace, now out of the list, joins as before.
    runCommand(session, 'joinBackward', undefined, { from: item.from });
    expect(session.doc.childCount).toBe(count - 1);
    // A style-numbered paragraph is switched off with numId 0 by both rules.
    const styled = await EditorDocument.open(await fixture('headers-footers.docx'));
    const bullet = paragraph(styled.doc, (a) => a.pStyle === 'ListBullet' && a.numPr === null);
    runCommand(styled, 'joinBackward', undefined, { from: bullet.from });
    expect(numPrOf(styled.doc, bullet.pos)).toEqual({ numId: '0', ilvl: 0 });
  });
});
