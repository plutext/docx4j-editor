// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 4.9 step B4: split and join with Word's paragraph-mark rules, breaks and tabs.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage, textOf as textOfElement, typeNameOf, runItemsOf, type Element } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { EditorDocument, runCommand, commandNamed, stateFor } from '../src/index.mjs';
import { fixture, comparePackages, plain, textOf } from './helpers.mjs';

async function reloadBody(bytes: Uint8Array): Promise<wml.Body> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  return (await pkg.getMainDocumentPart().readContents()).body!;
}

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
const pPrOf = (p: Element): wml.PPr | undefined => (p.value as wml.P).pPr;

describe('insertParagraph (Enter)', () => {
  test('in the middle: the text after keeps the paragraph identity; the text before is a new paragraph with the same properties', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    const i = 4;   // "Some text."
    const { from } = block(session, i);
    expect(runCommand(session, 'insertParagraph', undefined, { from: from + 4 }).ok).toBe(true);
    const doc = session.doc;
    expect(doc.childCount).toBe(before.length + 1);
    expect(doc.child(i).attrs['sourceId']).toBeNull();
    expect(doc.child(i + 1).attrs['sourceId']).toBe(session.imported.child(i).attrs['sourceId']);
    const after = (await saveAndReload(session, original)).content!;
    expect(after.length).toBe(before.length + 1);
    for (let k = 0; k < i; k++) expect(plain(after[k])).toEqual(plain(before[k]));
    for (let k = i + 1; k < before.length; k++) expect(plain(after[k + 1])).toEqual(plain(before[k]));
    expect(textOfElement(after[i]!)).toBe('Some');
    expect(textOfElement(after[i + 1]!)).toBe(' text.');
    expect(plain(pPrOf(after[i]!))).toEqual(plain(pPrOf(before[i]!)));
    expect(plain(pPrOf(after[i + 1]!))).toEqual(plain(pPrOf(before[i]!)));
    // Both halves' runs come from the original run.
    expect(plain(runsOf(after[i]!)[0]!.value.rPr)).toEqual(plain(runsOf(before[i]!)[0]!.value.rPr));
    expect(plain(runsOf(after[i + 1]!)[0]!.value.rPr)).toEqual(plain(runsOf(before[i]!)[0]!.value.rPr));
  });

  test("in a paragraph whose mark has an empty w:rPr (pandoc's): the first half is written with none, and the save's self-check holds (Jason, 2026-10-17)", async () => {
    const pkg = await WordprocessingMLPackage.load(await fixture('tracked-changes.docx'));
    const part = pkg.getMainDocumentPart();
    const xml = await part.getXml();
    const patched = xml.replace('<w:r><w:t>Some text.</w:t></w:r>', '<w:pPr><w:rPr/></w:pPr><w:r><w:t>Some text.</w:t></w:r>');
    expect(patched).not.toBe(xml);
    part.setXml(patched);
    const original = await pkg.save();
    const session = await EditorDocument.open(original);
    const i = 4;
    expect(session.doc.child(i).attrs['markRPr']).toBeNull();
    const { from } = block(session, i);
    expect(runCommand(session, 'insertParagraph', undefined, { from: from + 4 }).ok).toBe(true);
    // `save` runs the self-check: the written file re-projects to the live document.
    const after = (await saveAndReload(session, original)).content!;
    expect(textOfElement(after[i]!)).toBe('Some');
    expect(textOfElement(after[i + 1]!)).toBe(' text.');
    expect(pPrOf(after[i]!)).toBeUndefined();
    expect(plain(pPrOf(after[i + 1]!)?.rPr)).toEqual({ TYPE_NAME: 'org_docx4j_wml.ParaRPr' });
  });

  test("in a paragraph with properties outside the editable subset (a keep, a border, a bold mark): each half made by Enter keeps the whole w:pPr, as Word's new mark does (ED-005 section 12.102)", async () => {
    const pkg = await WordprocessingMLPackage.load(await fixture('tracked-changes.docx'));
    const part = pkg.getMainDocumentPart();
    const xml = await part.getXml();
    const patched = xml.replace('<w:r><w:t>Some text.</w:t></w:r>', '<w:pPr><w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>Some text.</w:t></w:r>');
    expect(patched).not.toBe(xml);
    part.setXml(patched);
    const original = await pkg.save();
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    const i = 4;
    const { from } = block(session, i);
    // "Some| text." then "So|me": the second split is in a paragraph that is itself a copy, so `copiedFrom` names the root.
    expect(runCommand(session, 'insertParagraph', undefined, { from: from + 4 }).ok).toBe(true);
    expect(runCommand(session, 'insertParagraph', undefined, { from: from + 2 }).ok).toBe(true);
    expect(session.doc.child(i).attrs['copiedFrom']).toBe(session.imported.child(i).attrs['sourceId']);
    expect(session.doc.child(i + 1).attrs['copiedFrom']).toBe(session.imported.child(i).attrs['sourceId']);
    expect(session.doc.child(i + 2).attrs['sourceId']).toBe(session.imported.child(i).attrs['sourceId']);
    const after = (await saveAndReload(session, original)).content!;
    expect(after.length).toBe(before.length + 2);
    expect([after[i], after[i + 1], after[i + 2]].map((p) => textOfElement(p!))).toEqual(['So', 'me', ' text.']);
    for (const k of [i, i + 1, i + 2]) expect(plain(pPrOf(after[k]!))).toEqual(plain(pPrOf(before[i]!)));
    // Opened again, the halves project as the original did, bold mark included.
    const again = await EditorDocument.open((await session.save()).bytes);
    for (const k of [i, i + 1, i + 2]) expect(plain(again.doc.child(k).attrs['markRPr'])).toEqual(plain(session.imported.child(i).attrs['markRPr']));
  });

  test('at the start: an empty new paragraph before, the original untouched (emitted by identity)', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    const i = 4;
    runCommand(session, 'insertParagraph', undefined, { from: block(session, i).from });
    expect(session.exportTree().content[i + 1]).toBe(before[i]);   // the very element: the identity fast path
    const after = (await saveAndReload(session, original)).content!;
    expect(textOfElement(after[i]!)).toBe('');
    expect(plain(after[i + 1])).toEqual(plain(before[i]));
  });

  test("at the end of a heading: the new empty paragraph takes the style's w:next", async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const i = 2;   // Heading1
    expect(session.doc.child(i).attrs['pStyle']).toBe('Heading1');
    const resolver = session.pkg.propertyResolver;
    const next = resolver.getStyle('Heading1')?.next?.val;
    runCommand(session, 'insertParagraph', undefined, { from: block(session, i).to });
    const after = (await saveAndReload(session, original)).content!;
    expect(pPrOf(after[i]!)?.pStyle?.val).toBe('Heading1');
    expect(textOfElement(after[i + 1]!)).toBe('');
    const newStyle = pPrOf(after[i + 1]!)?.pStyle?.val;
    if (next === undefined || next === resolver.defaultParagraphStyleId) expect(newStyle).toBeUndefined();
    else expect(newStyle).toBe(next);
    expect(newStyle).not.toBe('Heading1');
  });

  test('a paragraph that ends a section keeps its section break on the second half', async () => {
    const original = await fixture('sections-multiple.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    const i = before.findIndex((e) => (e.value as wml.P).pPr?.sectPr !== undefined);
    expect(i).toBeGreaterThanOrEqual(0);
    runCommand(session, 'insertText', { text: 'end of section' }, { from: block(session, i).from });
    runCommand(session, 'insertParagraph', undefined, { from: block(session, i).from + 3 });
    const after = (await saveAndReload(session, original)).content!;
    expect(textOfElement(after[i]!)).toBe('end');
    expect(pPrOf(after[i]!)?.sectPr).toBeUndefined();
    expect(textOfElement(after[i + 1]!)).toBe(' of section');
    expect(plain(pPrOf(after[i + 1]!)?.sectPr)).toEqual(plain(pPrOf(before[i]!)?.sectPr));
  });

  test('a selection is deleted first; a caret outside a paragraph does nothing', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const { from } = block(session, 4);
    runCommand(session, 'insertParagraph', undefined, { from: from + 4, to: from + 5 });
    expect(textOf(session.doc.child(4))).toBe('Some');
    expect(textOf(session.doc.child(5))).toBe('text.');
    expect(commandNamed('insertParagraph')!.isEnabled!(stateFor(session, { from: block(session, 4).from, to: block(session, 5).to }))).toBe(false);
  });
});

describe('joinBackward and joinForward', () => {
  test('an empty paragraph before is removed; the joined paragraph keeps its element', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    expect(textOf(session.doc.child(3))).toBe('');
    expect(runCommand(session, 'joinBackward', undefined, { from: block(session, 4).from }).ok).toBe(true);
    const exported = session.exportTree().content;
    expect(exported[3]).toBe(before[4]);
    for (let k = 5; k < before.length; k++) expect(exported[k - 1]).toBe(before[k]);
    const after = (await saveAndReload(session, original)).content!;
    expect(after.length).toBe(before.length - 1);
    expect(plain(after[3])).toEqual(plain(before[4]));
  });

  test('joining two paragraphs keeps the first\'s properties and identity, and takes the second\'s text', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    runCommand(session, 'joinBackward', undefined, { from: block(session, 1).from });
    expect(session.doc.child(0).attrs['sourceId']).toBe(session.imported.child(0).attrs['sourceId']);
    const after = (await saveAndReload(session, original)).content!;
    expect(after.length).toBe(before.length - 1);
    expect(textOfElement(after[0]!)).toBe(textOfElement(before[0]!) + textOfElement(before[1]!));
    expect(plain(pPrOf(after[0]!))).toEqual(plain(pPrOf(before[0]!)));
    // The second paragraph's runs come over with their own properties.
    const runsAfter = runsOf(after[0]!);
    const runsBefore = [...runsOf(before[0]!), ...runsOf(before[1]!)];
    expect(runsAfter.length).toBe(runsBefore.length);
    runsAfter.forEach((r, k) => expect(plain(r.value.rPr)).toEqual(plain(runsBefore[k]!.value.rPr)));
  });

  test('joinForward at the end of a paragraph is the same join from the other side', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    runCommand(session, 'joinForward', undefined, { from: block(session, 0).to });
    const after = (await saveAndReload(session, original)).content!;
    expect(textOfElement(after[0]!)).toBe(textOfElement(before[0]!) + textOfElement(before[1]!));
    expect(plain(pPrOf(after[0]!))).toEqual(plain(pPrOf(before[0]!)));
  });

  test('deleting a section-break mark joins the text into the following section', async () => {
    const original = await fixture('sections-multiple.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    const i = before.findIndex((e) => (e.value as wml.P).pPr?.sectPr !== undefined);
    runCommand(session, 'insertText', { text: 'last of section one' }, { from: block(session, i).from });
    runCommand(session, 'insertText', { text: 'first of section two' }, { from: block(session, i + 1).from });
    runCommand(session, 'joinBackward', undefined, { from: block(session, i + 1).from });
    expect(session.doc.child(i).attrs['sectPr']).toBeNull();
    const after = (await saveAndReload(session, original)).content!;
    expect(after.length).toBe(before.length - 1);
    expect(textOfElement(after[i]!)).toBe('last of section onefirst of section two');
    expect(pPrOf(after[i]!)?.sectPr).toBeUndefined();
    expect(after.some((e) => (e.value as wml.P).pPr?.sectPr !== undefined)).toBe(before.filter((e) => (e.value as wml.P).pPr?.sectPr !== undefined).length > 1);
  });

  test('nothing to join: false at the start of the document and inside a paragraph', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    expect(runCommand(session, 'joinBackward', undefined, { from: block(session, 0).from }).ok).toBe(false);
    expect(runCommand(session, 'joinBackward', undefined, { from: block(session, 4).from + 2 }).ok).toBe(false);
  });
});

describe('insertBreak and insertTab', () => {
  test('a page break inside a run joins the run: one w:r with t, br and t', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    const i = 4;
    expect(runCommand(session, 'insertBreak', { type: 'Page' }, { from: block(session, i).from + 4 }).ok).toBe(true);
    const node = session.doc.child(i).child(1);
    expect(node.type.name).toBe('hardBreak');
    expect(node.attrs['runSource']).toBe(session.doc.child(i).child(0).marks.find((m) => m.type.name === 'runSource')!.attrs['sourceId']);
    const after = (await saveAndReload(session, original)).content!;
    const runs = runsOf(after[i]!);
    expect(runs.length).toBe(1);
    expect((runs[0]!.value.content ?? []).map((c) => c.name.localPart)).toEqual(['t', 'br', 't']);
    expect((runs[0]!.value.content![1]!.value as wml.Br).type).toBe('page');
    expect(plain(runs[0]!.value.rPr)).toEqual(plain(runsOf(before[i]!)[0]!.value.rPr));
    expect(textOfElement(after[i]!)).toBe('Some\n text.');
  });

  test('a line break and a tab at the end of a paragraph, after formatting was turned on, carry that formatting', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const i = 4;
    const { to } = block(session, i);
    // A bold caret, then a tab: the tab's fmt says bold, its run is a copy of the source run with bold added.
    const bold = runCommand(session, 'font.bold', { value: true }, { from: to });
    const state = bold.state!;
    let next: typeof state | undefined;
    commandNamed('insertTab')!.run(undefined)(state, (tr) => { next = state.apply(tr); });
    session.replace(next!.doc);
    const tab = session.doc.child(i).lastChild!;
    expect(tab.type.name).toBe('tab');
    expect(tab.attrs['fmt']).toContain('bold');
    const after = (await saveAndReload(session, original)).content!;
    const runs = runsOf(after[i]!);
    expect(runs.length).toBe(2);
    expect((runs[1]!.value.content ?? []).map((c) => c.name.localPart)).toEqual(['tab']);
    expect(runs[1]!.value.rPr?.b).toBeDefined();
    expect(runs[0]!.value.rPr?.b).toBeUndefined();
  });
});
