// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 proposal 46 (Word fragments): the editor's parts that need nothing new from the engine. A Flat OPC value is no
// XHTML (the defect: a bare pkg:package read as inline markup); a fragment's facts; the two forms of a fragment control;
// the Data tab's bind of a Flat OPC node making Word's form (decision (e)); the Control tab's switch to OpenDoPE's form.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorDocument, runCommand, xhtmlKindOf, fragmentFactsOf, refreshFragments, staleFragmentControls, liveLabels, fragmentBindingOf, isFragmentControl, dataBindRefusal, bindsFragment, sdtBinding, sdtTag, readTag, sessionPlugin, controlAt, sdtPrOf, sdtId } from '../src/index.mjs';
import { fixture, fixturesDir } from './helpers.mjs';

/** The fragment as check 46a's chunk part holds it: a Flat OPC package, with the XML and mso-application declarations. */
async function fragmentText(): Promise<string> {
  const zip = unzipSync(new Uint8Array(await readFile(join(fixturesDir, 'fragments', '46a-altchunk.docx'))));
  return strFromU8(zip['word/chunk1.xml']!);
}

function stateAt(session: EditorDocument, pos: number): EditorState {
  const state = EditorState.create({ doc: session.doc, plugins: [sessionPlugin(session)] });
  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));
}

/** The position inside the paragraph whose text is `text`, at its end. */
function endOf(session: EditorDocument, text: string): number {
  let at = -1;
  session.doc.descendants((n, pos) => { if (at < 0 && n.type.name === 'paragraph' && n.textContent === text) at = pos + 1 + n.content.size; return at < 0; });
  return at;
}

/** The position inside the control whose alias is `alias`. */
function inControl(session: EditorDocument, alias: string): number {
  let at = -1;
  session.doc.descendants((n, pos) => {
    if (at >= 0) return false;
    if ((n.type.name === 'sdtBlock' || n.type.name === 'sdtInline') && JSON.stringify(n.attrs['sdtPr']).includes(`"${alias}"`)) { at = pos + (n.type.name === 'sdtBlock' ? 2 : 1); return false; }
    return true;
  });
  return at;
}

const FRAGMENT = { xpath: '/data[1]/fragment[1]' };

describe('Word fragments (ED-005 proposal 46): the editor\'s parts that need no new engine', () => {
  test('a Flat OPC value is no XHTML, with or without its declarations (the defect)', async () => {
    const value = await fragmentText();
    expect(xhtmlKindOf(value)).toBe('none');
    const bare = value.replace(/^(<\?[^>]*\?>\s*)*/, '');
    expect(bare.startsWith('<pkg:package')).toBe(true);
    expect(xhtmlKindOf(bare)).toBe('none');
    expect(xhtmlKindOf('<p>Jane <b>Doe</b></p>')).toBe('block');
  });

  test('fragmentFactsOf: the body\'s own paragraphs and tables, its pictures, notes and controls, nothing to drop', async () => {
    const value = await fragmentText();
    const facts = await fragmentFactsOf(value);
    expect(facts).toMatchObject({ bytes: new TextEncoder().encode(value).length, paragraphs: 7, tables: 1, pictures: 1, notes: 1, controls: 1, comments: 0, sectionBreaks: 0 });
    expect(facts!.text.startsWith("Fragment heading (the fragment's Heading 1: red, 20 pt)A paragraph in FragmentNote")).toBe(true);
    expect(await fragmentFactsOf('plain text')).toBeUndefined();
    expect(await fragmentFactsOf('<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">broken')).toBeUndefined();
  });

  test('check 46c: OpenDoPE\'s control and Word\'s w15 mapping read as fragment bindings; the plain text control mapped to the same node does not', async () => {
    const session = await EditorDocument.open(await fixture('fragments/46c-template.docx'));
    await session.template.load();
    const sdtPrAt = (alias: string): unknown => sdtPrOf(controlAt(stateAt(session, inControl(session, alias)))!, session.controlEntry);
    expect(isFragmentControl(sdtPrAt('Word fragment'))).toBe(true);
    expect(fragmentBindingOf(sdtPrAt('Word fragment'), session.template)).toMatchObject({ shape: 'progid', binding: { xpath: '/data/fragment' } });
    expect(fragmentBindingOf(sdtPrAt('Rich text mapped'), session.template)).toMatchObject({ shape: 'w15', binding: { xpath: '/data[1]/fragment[1]' } });
    expect(fragmentBindingOf(sdtPrAt('Plain text mapped'), session.template)).toBeUndefined();
  });

  test('the Data tab\'s bind of a Flat OPC node makes Word\'s form around a blank paragraph, and is refused in a paragraph with content', async () => {
    const session = await EditorDocument.open(await fixture('fragments/46c-template.docx'));
    await session.template.load();
    const after = endOf(session, 'After the control.');
    expect(bindsFragment(stateAt(session, after), FRAGMENT)).toBe(true);
    expect(dataBindRefusal(stateAt(session, after), FRAGMENT)).toBe('blockFragment');
    expect(runCommand(session, 'contentControl.insertBound', FRAGMENT, { from: after }).ok).toBe(false);
    // A blank paragraph after it: the bind makes a block-level rich text control mapped by w15:dataBinding around it.
    expect(runCommand(session, 'insertParagraph', undefined, { from: after }).ok).toBe(true);
    const blank = after + 2;
    expect(session.doc.resolve(blank).parent.textContent).toBe('');
    expect(dataBindRefusal(stateAt(session, blank), FRAGMENT)).toBeUndefined();
    expect(runCommand(session, 'contentControl.insertBound', FRAGMENT, { from: blank }).ok).toBe(true);
    const control = session.doc.resolve(blank).node(1);
    expect(control.type.name).toBe('sdtBlock');
    const binding = sdtBinding(control.attrs['sdtPr']);
    expect(binding).toMatchObject({ xpath: '/data[1]/fragment[1]', w15: true });
    expect(isFragmentControl(control.attrs['sdtPr'])).toBe(false);
  });

  test('the Control tab\'s switch: Word\'s w15 form to OpenDoPE\'s od:progid form through the XPaths entry, and back', async () => {
    const session = await EditorDocument.open(await fixture('fragments/46c-template.docx'));
    await session.template.load();
    const at = inControl(session, 'Rich text mapped');
    const sdtPr = (): unknown => sdtPrOf(controlAt(stateAt(session, at))!, session.controlEntry);
    expect(runCommand(session, 'contentControl.wordFragment', { on: true }, { from: at }).ok).toBe(true);
    const tag = readTag(sdtTag(sdtPr()));
    expect(tag.progid).toBe('Word.Document');
    expect(tag.xpath).toBeDefined();
    expect(sdtBinding(sdtPr())).toBeUndefined();
    expect(session.template.xpaths?.entries.find((e) => e.id === tag.xpath)?.xpath).toBe('/data[1]/fragment[1]');
    expect(runCommand(session, 'contentControl.wordFragment', { on: false }, { from: at }).ok).toBe(true);
    expect(readTag(sdtTag(sdtPr())).progid).toBeUndefined();
    expect(sdtBinding(sdtPr())).toMatchObject({ xpath: '/data[1]/fragment[1]', w15: true });
    // The XHTML switch and the fragment switch exclude each other.
    expect(runCommand(session, 'contentControl.wordFragment', { on: true }, { from: at }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.xhtml', { on: true }, { from: at }).ok).toBe(true);
    expect(readTag(sdtTag(sdtPr())).progid).toBeUndefined();
  });

  // Part 3 (section 12.119): the editor's refresh, the engine's pass (core-ts CR-005 section 11) in a tree run over the controls.
  test('the refresh fills both forms as Word 365 does: the document\'s Heading 1, the style copied, a new list each, the document\'s list going on; out of the undo history; an unchanged value fills nothing; the save holds it', async () => {
    const session = await EditorDocument.open(await fixture('fragments/46c-template.docx'));
    await session.template.load();
    const undo = (): number => (session as unknown as { undoManager: { undoStack: unknown[] } }).undoManager.undoStack.length;
    const before = undo();
    expect(await refreshFragments(session)).toEqual({ changed: 2, notes: [] });
    const doc = session.doc;
    const labels = liveLabels(doc, session);
    const listed: string[] = [];
    const styled: string[] = [];
    doc.descendants((n) => {
      if (n.type.name !== 'paragraph') return true;
      const label = labels.get(n);
      if (label) listed.push(label.listString);
      if (n.textContent.startsWith('Fragment heading')) styled.push(n.attrs['pStyle'] as string);
      return false;
    });
    expect(listed).toEqual(['1.', '2.', 'I.', 'II.', 'I.', 'II.', '3.']);
    expect(styled).toEqual(['Heading1', 'Heading1']);
    expect(session.pkg.propertyResolverOrUndefined!.getEffectiveRPr('FragmentNote')).toMatchObject({ i: {}, color: { val: '00B050' } });
    expect(undo()).toBe(before);
    expect(await refreshFragments(session)).toEqual({ changed: 0, notes: [] });
    const saved = await session.save();
    expect(saved.selfCheck?.equal ?? true).toBe(true);
    const zip = unzipSync(saved.bytes);
    expect((strFromU8(zip['word/document.xml']!).match(/Fragment heading/g) ?? []).length).toBe(2);
    expect(zip['word/footnotes.xml']).toBeDefined();
  });

  test('a changed value fills again, and the previous fill\'s lists and picture that nothing references go', async () => {
    const session = await EditorDocument.open(await fixture('fragments/46c-template.docx'));
    await session.template.load();
    await refreshFragments(session, { shape: 'w15' });
    const node = session.template.dataPart.part!.selectSingleNode('/data/fragment')!;
    node.text = node.text.replace('Last paragraph of the fragment.', 'The last paragraph, changed.');
    session.markPartChanged(session.template.dataPart.part!.part.partName.toString());
    expect(await refreshFragments(session, { shape: 'w15' })).toEqual({ changed: 1, notes: [] });
    expect(session.doc.textContent).toContain('The last paragraph, changed.');
    const zip = unzipSync((await session.save()).bytes);
    // The document's own list and the second fill's: the first fill's went.
    expect(strFromU8(zip['word/numbering.xml']!).match(/<w:num w:numId=/g)).toHaveLength(2);
    expect(strFromU8(zip['word/numbering.xml']!).match(/<w:abstractNum /g)).toHaveLength(2);
    expect(Object.keys(zip).filter((k) => k.startsWith('word/media/'))).toHaveLength(1);
  });

  test('the stale count: Word\'s w15 control in 46c\'s template, none in Word 365\'s save of it, never the od:progid control', async () => {
    for (const [name, expected] of [['fragments/46c-template.docx', 1], ['fragments/46c-word365.docx', 0]] as const) {
      const session = await EditorDocument.open(await fixture(name));
      await session.template.load();
      expect(await staleFragmentControls(session, session.doc)).toBe(expected);
    }
  });
});


// Section 12.124 (Jason, 2026-11-12): a bind of an existing control to a Flat OPC node shows the Word document, as a bind of
// XHTML shows its markup: the control is made Word's form, and the refresh of that control fills it.
describe('a Word document bound to a control already there (ED-005 section 12.124)', () => {
  /** A session over 46c's template with a blank paragraph after "After the control.", the caret in it. */
  async function blankAfter(): Promise<{ session: EditorDocument; blank: number }> {
    const session = await EditorDocument.open(await fixture('fragments/46c-template.docx'));
    await session.template.load();
    expect(runCommand(session, 'insertParagraph', undefined, { from: endOf(session, 'After the control.') }).ok).toBe(true);
    return { session, blank: endOf(session, 'After the control.') + 2 };
  }

  test('a run-level plain text control alone in its paragraph becomes a block-level rich text control around it, mapped by w15, its id kept; the refresh of it fills it, and only it', async () => {
    const { session, blank } = await blankAfter();
    expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText', title: 'Was plain' }, { from: blank }).ok).toBe(true);
    const inside = blank + 1;
    const before = sdtPrOf(controlAt(stateAt(session, inside))!, session.controlEntry);
    expect(controlAt(stateAt(session, inside))!.form).toBe('run');
    expect(dataBindRefusal(stateAt(session, inside), FRAGMENT)).toBeUndefined();
    expect(runCommand(session, 'contentControl.bind', FRAGMENT, { from: inside }).ok).toBe(true);
    const target = controlAt(stateAt(session, blank + 1))!;
    expect(target.form).toBe('block');
    const sdtPr = sdtPrOf(target, session.controlEntry);
    expect(fragmentBindingOf(sdtPr, session.template)).toMatchObject({ shape: 'w15', binding: { xpath: '/data[1]/fragment[1]' } });
    expect(sdtBinding(sdtPr)).toMatchObject({ w15: true });
    // Its id and title kept; the template's XPaths part gives it an od:xpath entry too, as a new control's bind does.
    const id = sdtId(sdtPr)!;
    expect(id).toBe(sdtId(before));
    expect(JSON.stringify(sdtPr)).toContain('Was plain');
    expect(await refreshFragments(session, { ids: new Set([id]) })).toEqual({ changed: 1, notes: [] });
    // That control alone: Word's own w15 control in the template keeps its placeholder until its own Refresh.
    let mapped = '';
    session.doc.descendants((n) => { if (n.type.name === 'sdtBlock' && JSON.stringify(n.attrs['sdtPr']).includes('"Rich text mapped"')) mapped = n.textContent; return true; });
    expect(mapped).toContain('rich text control placeholder');
    let filled = '';
    session.doc.descendants((n) => { if (n.type.name === 'sdtBlock' && JSON.stringify(n.attrs['sdtPr']).includes('"Was plain"')) filled = n.textContent; return true; });
    expect(filled).toContain('Fragment heading');
    const saved = await session.save();
    expect(saved.selfCheck?.equal ?? true).toBe(true);
  });

  test('a block-level rich text control keeps its form and is mapped by w15', async () => {
    const { session, blank } = await blankAfter();
    const paragraph = blank - 1;
    expect(runCommand(session, 'contentControl.insert', { kind: 'RichText', title: 'A block' }, { from: paragraph, to: paragraph + session.doc.nodeAt(paragraph)!.nodeSize }).ok).toBe(true);
    const inside = paragraph + 2;
    expect(controlAt(stateAt(session, inside))!.form).toBe('block');
    expect(runCommand(session, 'contentControl.bind', FRAGMENT, { from: inside }).ok).toBe(true);
    const sdtPr = sdtPrOf(controlAt(stateAt(session, inside))!, session.controlEntry);
    expect(fragmentBindingOf(sdtPr, session.template)).toMatchObject({ shape: 'w15' });
  });

  test('refused, saying why: a run-level control beside text in its paragraph, and a control of another kind', async () => {
    const session = await EditorDocument.open(await fixture('fragments/46c-template.docx'));
    await session.template.load();
    const end = endOf(session, 'After the control.');
    expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: end - 8, to: end - 1 }).ok).toBe(true);
    const inText = end - 6;
    expect(dataBindRefusal(stateAt(session, inText), FRAGMENT)).toBe('blockFragment');
    expect(runCommand(session, 'contentControl.bind', FRAGMENT, { from: inText }).ok).toBe(false);
    const { session: other, blank } = await blankAfter();
    expect(runCommand(other, 'contentControl.insert', { kind: 'DatePicker' }, { from: blank }).ok).toBe(true);
    expect(dataBindRefusal(stateAt(other, blank + 1), FRAGMENT)).toBe('fragmentKind');
    expect(runCommand(other, 'contentControl.bind', FRAGMENT, { from: blank + 1 }).ok).toBe(false);
  });
});
