// SPDX-License-Identifier: GPL-3.0-or-later
//
// core-ts check 35 (part 1 run by Jason in current Word, 2026-10-03; `fixtures/revisions/check35/`, copied from
// docx4j-core-ts's `test/fixtures/revisions/check35/`): rich text content controls and custom XML mapping.
// `35-input.docx` is the engine's; `35a-word15.docx` is Word's save after mapping them in its XML Mapping Pane. Word
// maps rich text with `w15:dataBinding`, the node's value winning; an edit writes the content into the node as escaped
// Flat OPC; a rich text control with `w:dataBinding` is plain text to Word; the pane nests a plain text control inside
// a mapped one, with an alert when it is the same node. Part 2, in Word 2010: `35c-word2010.docx` from Word's save and
// `35e-word2010.docx` from the editor's `35d-editor.docx` (`fixtures/build/check35-for-word.mjs`). ED-005 section 12.58.
import { describe, expect, test } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, stateFor, controlToBind, dataBindRefusal, applyBoundRefresh, isFlatOpc, mappedKind, readControl, sdtBinding } from '../src/index.mjs';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { fixture, mainXml, comparePackages } from './helpers.mjs';

async function open(name: '35-input.docx' | '35a-word15.docx' | '35c-word2010.docx' | '35e-word2010.docx'): Promise<EditorDocument> {
  const session = await EditorDocument.open(await fixture(`revisions/check35/${name}`));
  await session.template.load();
  return session;
}

/** The controls in a block, outermost first, with their positions. */
function controlsIn(session: EditorDocument, block: number): { node: PMNode; pos: number }[] {
  let start = 0;
  for (let k = 0; k < block; k++) start += session.doc.child(k).nodeSize;
  const out: { node: PMNode; pos: number }[] = [];
  const b = session.doc.child(block);
  if (b.type.name === 'sdtBlock') out.push({ node: b, pos: start });
  b.descendants((n, p) => { if (n.type.name === 'sdtInline' || n.type.name === 'sdtBlock') out.push({ node: n, pos: start + 1 + p }); return true; });
  return out;
}

/** The index of the block whose text starts with a case's label ("1a.", "1f."). */
function blockOf(session: EditorDocument, label: string): number {
  let found = -1;
  session.doc.forEach((b, _o, i) => { if (found < 0 && b.textContent.startsWith(label)) found = i; });
  expect(found).toBeGreaterThanOrEqual(0);
  return found;
}

/** The styles part of saved bytes as text. */
async function stylesXml(bytes: Uint8Array): Promise<string> {
  return (await WordprocessingMLPackage.load(bytes)).getMainDocumentPart().styleDefinitionsPart!.getXml();
}

const value = (session: EditorDocument, xpath: string): string => session.template.dataPart.part!.selectSingleNode(xpath)!.nodeValue;

describe('core-ts check 35: rich text content controls mapped as Word maps them', () => {
  test("Word's save: a node holding a rich text control's Flat OPC is neither shown as text nor written over", async () => {
    const session = await open('35a-word15.docx');
    // 1a: rich text, `w15:dataBinding`; its node holds the control's content as a Flat OPC package.
    const [a, nested] = controlsIn(session, 1);
    expect(readControl(a!.node.attrs['sdtPr']).kind).toBe('RichText');
    expect(sdtBinding(a!.node.attrs['sdtPr'])).toMatchObject({ xpath: '/data[1]/name[1]', w15: true });
    expect(isFlatOpc(value(session, '/data/name'))).toBe(true);
    // 1e: the plain text control Word nested in it, bound to the same node, which Word keeps inert.
    expect(sdtBinding(nested!.node.attrs['sdtPr'])).toMatchObject({ xpath: '/data[1]/name[1]', w15: false });
    // A refresh of every bound control changes nothing: the Flat OPC is not put into 1a or 1e as text.
    const before = session.doc;
    expect(applyBoundRefresh(session, stateFor(session))).toBe(false);
    expect(session.doc.eq(before)).toBe(true);
    // Typing in 1e, its placeholder flag cleared so that it has a value, does not write over the node.
    expect(runCommand(session, 'contentControl.properties', { showingPlaceholder: false }, { from: nested!.pos + 1 }).ok).toBe(true);
    expect(runCommand(session, 'insertText', { text: 'x' }, { from: nested!.pos + 1 }).ok).toBe(true);
    expect(controlsIn(session, 1)[1]!.node.textContent).toMatch(/^x/);
    expect(isFlatOpc(value(session, '/data/name'))).toBe(true);
    // 1f: Word made it plain text (`<w:text/>`) and wrote "Bea more"; typing there writes the node, as in Word.
    const [f] = controlsIn(session, 7);
    expect(readControl(f!.node.attrs['sdtPr']).kind).toBe('PlainText');
    expect(value(session, '/data/bound')).toBe('Bea more');
    expect(runCommand(session, 'insertText', { text: '!' }, { from: f!.pos + f!.node.nodeSize - 1 }).ok).toBe(true);
    expect(value(session, '/data/bound')).toBe('Bea more!');
    // Saved after an edit, Word's `w15:dataBinding`s and its `storeItemChecksum` stay.
    const xml = await mainXml((await session.save()).bytes);
    expect(xml.match(/<w15:dataBinding /g)).toHaveLength(4);
    expect(xml).toContain('storeItemChecksum="JAms/A=="');
  });

  test("the editor's binds over the input: rich text mapped with w15:dataBinding (1a, 1b), nesting inside it (1e), w:dataBinding as plain text (1f)", async () => {
    const session = await open('35-input.docx');
    const word = await open('35a-word15.docx');
    const part = session.template.dataPart.part!;
    const map = (xpath: string): { xpath: string; storeItemID: string } => ({ xpath, storeItemID: part.id });
    // 1a: Map to Selected over "quick brown": `w15:dataBinding`, as Word's, and the node's "Ann" in it.
    const [a] = controlsIn(session, 1);
    expect(runCommand(session, 'contentControl.bind', map('/data[1]/name[1]'), { from: a!.pos + 2 }).ok).toBe(true);
    const bound = controlsIn(session, 1)[0]!.node;
    expect(sdtBinding(bound.attrs['sdtPr'])).toMatchObject({ xpath: '/data[1]/name[1]', w15: true });
    expect(sdtBinding(controlsIn(word, 1)[0]!.node.attrs['sdtPr'])).toMatchObject({ xpath: '/data[1]/name[1]', w15: true });
    expect(bound.textContent).toBe('Ann');
    // 1e: inside it, the Data tab nests a control, but not on its own node: Word's pane nested one with an alert ("The
    // custom XML node is already mapped to a rich text content control, so it can't be mapped to a plain text content
    // control") and saved it with an inert binding; the editor refuses instead (docs/developer/word-divergences.md).
    const inside = { from: controlsIn(session, 1)[0]!.pos + 3 };
    expect(controlToBind(stateFor(session, inside))).toBeUndefined();
    expect(dataBindRefusal(stateFor(session, inside), map('/data[1]/name[1]'))).toBe('mapped');
    expect(runCommand(session, 'contentControl.insertBound', map('/data[1]/name[1]'), inside).ok).toBe(false);
    expect(runCommand(session, 'contentControl.insertBound', map('/data[1]/rich2[1]'), inside).ok).toBe(true);
    expect(controlsIn(session, 1).map((c) => readControl(c.node.attrs['sdtPr']).kind)).toEqual(['RichText', 'PlainText']);
    // 1b: the block-level control mapped to the empty rich1: the placeholder, as Word showed. Word also made it a run-level
    // control at the start of the next paragraph; the editor keeps it a block.
    const [b] = controlsIn(session, 3);
    expect(runCommand(session, 'contentControl.bind', map('/data[1]/rich1[1]'), { from: b!.pos + 2 }).ok).toBe(true);
    expect(readControl(session.doc.child(3).attrs['sdtPr']).showingPlaceholder).toBe(true);
    expect(sdtBinding(session.doc.child(3).attrs['sdtPr'])).toMatchObject({ w15: true });
    expect(readControl(controlsIn(word, 3)[0]!.node.attrs['sdtPr']).showingPlaceholder).toBe(true);
    expect(controlsIn(word, 3)[0]!.node.type.name).toBe('sdtInline');
    // 1f: rich text with `w:dataBinding` is plain text, as Word takes it: bound in place, and written back. The editor
    // shows the file's "Old text" where Word refreshed it to "Bea" on open (docs/developer/TODOs_OpenDoPE.md, "On open").
    const [f] = controlsIn(session, 8);
    expect(mappedKind(f!.node.attrs['sdtPr'])).toBe('PlainText');
    expect(controlToBind(stateFor(session, { from: f!.pos + 2 }))).toMatchObject({ pos: f!.pos });
    expect(runCommand(session, 'insertText', { text: ' more' }, { from: f!.pos + f!.node.nodeSize - 1 }).ok).toBe(true);
    expect(value(session, '/data/bound')).toBe('Old text more');
    // Saved: the `w15:dataBinding`s written, with `w15` ignorable for Word 2010, and the placeholder's style defined, which
    // the engine-made document lacked (UPSTREAM-19, ED-005 section 12.59).
    const bytes = (await session.save()).bytes;
    const xml = await mainXml(bytes);
    expect(xml.match(/<w15:dataBinding /g)).toHaveLength(2);
    expect(xml).toMatch(/<w:document [^>]*mc:Ignorable="[^"]*\bw15\b/);
    expect(await stylesXml(await fixture('revisions/check35/35-input.docx'))).not.toContain('PlaceholderText');
    expect(await stylesXml(bytes)).toMatch(/<w:style [^>]*w:styleId="PlaceholderText"[^>]*>.*<w:name w:val="Placeholder Text"\/>.*<w:color w:val="666666"\/>/);
  });

  test('the placeholder style is written only when a save writes the body and a run has it', async () => {
    const input = await fixture('revisions/check35/35-input.docx');
    // An edit with no placeholder: the styles part is the source's, byte for byte.
    const plain = await open('35-input.docx');
    expect(runCommand(plain, 'insertText', { text: 'x' }, { from: 2 }).ok).toBe(true);
    const saved = (await plain.save()).bytes;
    expect(comparePackages(input, saved, ['word/document.xml', 'docProps/app.xml'])).toEqual([]);
    // A placeholder made at a caret (a rich text control): the style is added, and kept when it is there already.
    const made = await open('35-input.docx');
    expect(runCommand(made, 'contentControl.insert', { kind: 'RichText' }, { from: 2 }).ok).toBe(true);
    const once = (await made.save()).bytes;
    expect((await stylesXml(once)).match(/w:styleId="PlaceholderText"/g)).toHaveLength(1);
    const again = await EditorDocument.open(once);
    expect(runCommand(again, 'insertText', { text: 'y' }, { from: 2 }).ok).toBe(true);
    expect((await stylesXml((await again.save()).bytes)).match(/w:styleId="PlaceholderText"/g)).toHaveLength(1);
  });

  test('Word 2010: the w15 bindings dropped, w:dataBinding rich text taken as plain text, the inert nested control showing the Flat OPC', async () => {
    // Word's save through Word 2010: no `w15:dataBinding` left; 1e's binding, inert in Word 15, shows the node's package as
    // text. The editor leaves that text as it is, and its typing there writes nothing over the node.
    const word = await open('35c-word2010.docx');
    const bindings: { w15: boolean }[] = [];
    word.doc.descendants((n) => { const b = n.type.name.startsWith('sdt') ? sdtBinding(n.attrs['sdtPr']) : undefined; if (b) bindings.push(b); return true; });
    expect(bindings.map((b) => b.w15)).toEqual([false, false]);   // 1e and 1f
    const [, nested] = controlsIn(word, blockOf(word, '1a.'));
    expect(nested!.node.textContent.startsWith('<?xml version="1.0" standalone="yes"?><?mso-application')).toBe(true);
    expect(applyBoundRefresh(word, stateFor(word))).toBe(false);
    expect(runCommand(word, 'insertText', { text: 'x' }, { from: nested!.pos + 1 }).ok).toBe(true);
    expect(isFlatOpc(value(word, '/data/name'))).toBe(true);
    // The editor's file through Word 2010: opened, its `w15` elements dropped (the bindings, the repeating section), 1f
    // shown from its node and made plain text, the placeholder runs without the style the styles part lacked
    // (docs/developer/upstream-limitations.md item 19).
    const editor = await open('35e-word2010.docx');
    const xml = await mainXml(await fixture('revisions/check35/35e-word2010.docx'));
    expect(xml).not.toMatch(/<w15:/);
    const f = controlsIn(editor, blockOf(editor, '1f.'))[0]!.node;
    expect(f.textContent).toBe('Bea');
    expect(readControl(f.attrs['sdtPr']).kind).toBe('PlainText');
    expect(xml).toContain('<w:r><w:t>Click here to enter text.</w:t></w:r>');
    expect(xml).not.toContain('PlaceholderText');
  });
});
