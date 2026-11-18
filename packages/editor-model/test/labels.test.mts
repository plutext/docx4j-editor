// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 5.8 item 3 (step C3): the node walk (`liveLabels`) agrees with the tree walk
// (`listLabels`) on every fixture untouched, and with the tree walk over the saved and reopened
// document after each list command, which is the oracle for live renumbering: the emulator is
// the same, only the input differs. And the level's indent merged with the paragraph's own.
import { describe, test, expect } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, listLabels, liveLabels, levelIndentOf, numberingParts, runCommand, type ListLabel, type ParagraphAttrs } from '../src/index.mjs';
import { allFixtures, fixture } from './helpers.mjs';

const names = await allFixtures();

/** The labels in story order, one entry per paragraph node (undefined where it has none), from either walk. */
function sequence(doc: PMNode, lookup: (node: PMNode) => ListLabel | undefined): (Omit<ListLabel, never> | undefined)[] {
  const out: (ListLabel | undefined)[] = [];
  doc.descendants((node) => {
    if (node.type.name === 'opaqueBlock') return false;
    if (node.type.name !== 'paragraph') return true;
    const l = lookup(node);
    out.push(l && { listString: l.listString, level: l.level, numId: l.numId, siblingIndex: l.siblingIndex, isBullet: l.isBullet });
    return false;
  });
  return out;
}

const treeSequence = (session: EditorDocument): (ListLabel | undefined)[] => {
  const byId = listLabels(session);
  return sequence(session.doc, (node) => { const id = node.attrs['sourceId'] as string | null; return id ? byId.get(id) : undefined; });
};
const liveSequence = (session: EditorDocument, doc = session.doc): (ListLabel | undefined)[] => {
  const byNode = liveLabels(doc, session);
  return sequence(doc, (node) => byNode.get(node));
};

describe('the node walk agrees with the tree walk', () => {
  test.each(names)('%s untouched', async (name) => {
    const session = await EditorDocument.open(await fixture(name));
    expect(liveSequence(session)).toEqual(treeSequence(session));
  });

  test('numbering-stories.docx: the body items count in one story; the text boxes are opaque and start nothing', async () => {
    const session = await EditorDocument.open(await fixture('numbering-stories.docx'));
    const live = liveSequence(session).filter((l) => l !== undefined);
    expect(live.map((l) => l!.listString)).toEqual(['1.', '2.', '3.', '4.', '5.', '6.']);
    expect(live.every((l) => l!.numId === '70' && !l!.isBullet)).toBe(true);
  });

  test('a document with no numbering part has no labels', async () => {
    const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
    expect(liveLabels(session.doc, session).size).toBe(0);
  });
});

/** The live labels after a command against the tree walk over the document saved and reopened. */
async function expectLiveMatchesReopened(session: EditorDocument): Promise<void> {
  const saved = await session.save();
  const reopened = await EditorDocument.open(saved.bytes);
  expect(liveSequence(session)).toEqual(treeSequence(reopened));
}

/** The nth paragraph matching `pred`, with the position of its content. */
function paragraph(doc: PMNode, pred: (attrs: ParagraphAttrs) => boolean, n = 0): { pos: number; from: number } {
  let found: { pos: number; from: number } | undefined;
  let seen = 0;
  doc.descendants((node, pos) => {
    if (found || node.type.name !== 'paragraph') return !found;
    if (pred(node.attrs as ParagraphAttrs) && seen++ === n) found = { pos, from: pos + 1 };
    return false;
  });
  if (!found) throw new Error('no such paragraph');
  return found;
}

describe('live renumbering after each command (the reopened document as the oracle)', () => {
  test('attach, level, continue and detach in headers-footers.docx', async () => {
    const session = await EditorDocument.open(await fixture('headers-footers.docx'));
    const plain = (a: ParagraphAttrs): boolean => a.numPr === null && a.pStyle === null && a.sourceId !== null;
    const a = paragraph(session.doc, plain, 0);
    const b = paragraph(session.doc, plain, 1);
    const c = paragraph(session.doc, plain, 2);
    runCommand(session, 'paragraph.attachToList', { listId: 21 }, { from: a.from });
    runCommand(session, 'paragraph.attachToList', { listId: 22, level: 1 }, { from: b.from });
    runCommand(session, 'paragraph.attachToList', { listId: 22 }, { from: c.from });
    await expectLiveMatchesReopened(session);
    runCommand(session, 'list.continue', undefined, { from: b.from });
    await expectLiveMatchesReopened(session);
    runCommand(session, 'listItem.level', { value: 3 }, { from: c.from });
    await expectLiveMatchesReopened(session);
    const styled = paragraph(session.doc, (x) => x.pStyle === 'ListBullet' && x.numPr === null);
    runCommand(session, 'paragraph.detachFromList', undefined, { from: styled.from });
    await expectLiveMatchesReopened(session);
    const direct = paragraph(session.doc, (x) => x.numPr?.numId === '26');
    runCommand(session, 'list.outdent', undefined, { from: direct.from });
    await expectLiveMatchesReopened(session);
  });

  test('a paragraph typed into a list is numbered, and its followers renumber, before any save', async () => {
    const session = await EditorDocument.open(await fixture('numbering-stories.docx'));
    const first = paragraph(session.doc, (a) => a.numPr?.numId === '70');
    const end = first.from + session.doc.nodeAt(first.pos)!.content.size;
    runCommand(session, 'insertParagraph', undefined, { from: end });
    runCommand(session, 'insertText', { text: 'a new item' }, { from: end + 2 });
    const live = liveSequence(session).filter((l) => l !== undefined).map((l) => l!.listString);
    expect(live).toEqual(['1.', '2.', '3.', '4.', '5.', '6.', '7.']);
    await expectLiveMatchesReopened(session);
  });
});

describe('the level\'s indent (5.4)', () => {
  test('the level\'s w:ind with the paragraph\'s own over it; the first-line setting replaced as a whole', async () => {
    const session = await EditorDocument.open(await fixture('headers-footers.docx'));
    const numbering = numberingParts(session)!;
    const base: ParagraphAttrs = { sourceId: null, pStyle: 'ListParagraph', jc: null, ind: null, spacing: null, bidi: false, numPr: { numId: '26', ilvl: 0 }, tabs: null, sectPr: null, markRPr: null, markRevision: null, pPrChange: null, copiedFrom: null };
    const level = levelIndentOf(base, numbering)!;
    expect(level.hanging).toBeGreaterThan(0);
    expect(level.start).toBeGreaterThan(0);
    expect(levelIndentOf({ ...base, ind: { start: 2000, end: null, firstLine: null, hanging: null } }, numbering)).toEqual({ ...level, start: 2000 });
    expect(levelIndentOf({ ...base, ind: { start: null, end: null, firstLine: 100, hanging: null } }, numbering)).toEqual({ ...level, firstLine: 100, hanging: null });
    // Not a list item, or no numbering: null, and the paragraph's own attribute is all there is.
    expect(levelIndentOf({ ...base, numPr: null }, numbering)).toBeNull();
    expect(levelIndentOf(base, undefined)).toBeNull();
    // A style-numbered paragraph takes its style's level.
    expect(levelIndentOf({ ...base, pStyle: 'ListBullet', numPr: null }, numbering)).not.toBeNull();
    const labels = liveLabels(session.doc, session);
    const some = [...labels.values()].find((l) => l.numId === '26')!;
    expect(some.indent).toEqual(level);
    expect(some.suff).toBe('tab');
  });
});
