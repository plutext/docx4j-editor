// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.116: a comment from the file keeps its range when its paragraph is split or joined. Enter inside a
// comment makes the first half the new paragraph, and the export wrote the comment's start only in the paragraph that
// had held it, the second half: on reload the first half's text was outside the comment and every save was refused.
// A paste of paragraphs and Backspace at a commented paragraph's start did the same. Each case saves through the
// self-check, which is what refused them.
import { describe, test, expect } from 'vitest';
import { Fragment, Slice, type Node as PMNode } from 'prosemirror-model';
import { type EditorState, type Transaction } from 'prosemirror-state';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, LOCAL_ORIGIN, recordTransaction, runCommand, runCommandAsync, stateFor } from '../src/index.mjs';
import { fixture, mainXml } from './helpers.mjs';

/** The saved main part as its text and comment markers in order: `[0` a start, `0]` an end, `^0` a reference, `|` a paragraph's end. */
function shapeOf(xml: string): string {
  return [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>|<w:commentRangeStart w:id="(\d+)"\/>|<w:commentRangeEnd w:id="(\d+)"\/>|<w:commentReference w:id="(\d+)"\/>|<\/w:p>/g)]
    .map((m) => (m[1] !== undefined ? `"${m[1]}"` : m[2] !== undefined ? `[${m[2]}` : m[3] !== undefined ? `${m[3]}]` : m[4] !== undefined ? `^${m[4]}` : '|')).join(' ');
}

/** Saved through the self-check, which throws where the reload would differ. */
const saved = async (session: EditorDocument): Promise<string> => shapeOf(await mainXml((await session.save()).bytes));

/** The position `offset` characters into block `i` (whose children before it are text). */
function at(doc: PMNode, i: number, offset: number): number {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += doc.child(k).nodeSize;
  return pos + 1 + offset;
}

/** A local transaction recorded as the view records it and written to the store, one undo step. */
function edit(session: EditorDocument, make: (state: EditorState) => Transaction): void {
  const state = stateFor(session);
  const { tr } = recordTransaction(make(state), state);
  session.undoManager.stopCapturing();
  session.replace(state.apply(tr).doc, LOCAL_ORIGIN, [], undefined, state.doc);
}

const enter = (session: EditorDocument, i: number, offset: number): void => {
  session.undoManager.stopCapturing();
  expect(runCommand(session, 'insertParagraph', undefined, { from: at(session.doc, i, offset) }).ok).toBe(true);
};

/** Two paragraphs pasted at a place, as a paste of several opens its slice at both ends. */
const pasteTwo = (session: EditorDocument, i: number, offset: number): void => edit(session, (state) => {
  const para = (text: string): PMNode => state.schema.nodes['paragraph']!.create(null, state.schema.text(text));
  const pos = at(state.doc, i, offset);
  return state.tr.replace(pos, pos, new Slice(Fragment.from([para('first'), para('second')]), 1, 1));
});

/** A fixture with its main part's XML changed, and the main part as that file holds it. */
async function patchedFile(name: string, change: (xml: string) => string): Promise<{ session: EditorDocument; xml: string }> {
  const pkg = await WordprocessingMLPackage.load(await fixture(name));
  const part = pkg.getMainDocumentPart();
  const xml = await part.getXml();
  const next = change(xml);
  expect(next).not.toBe(xml);
  part.setXml(next);
  const bytes = new Uint8Array(await pkg.save());
  return { session: await EditorDocument.open(bytes), xml: await mainXml(bytes) };
}
const patched = async (name: string, change: (xml: string) => string): Promise<EditorDocument> => (await patchedFile(name, change)).session;

// comments-two.docx: "One [0 paragraph 0] ^0 |" then "[1 blsgh  1] ^1 |".
const TWO = '"One " [0 "paragraph" 0] ^0 | [1 "blsgh " 1] ^1 |';

describe('a comment from the file keeps its range across a split or a join (ED-005 section 12.116)', () => {
  test('the file as it is', async () => {
    expect(await saved(await EditorDocument.open(await fixture('comments-two.docx')))).toBe(TWO);
  });

  test('Enter inside the comment: its start goes into the first half, the new paragraph', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    enter(session, 0, 'One par'.length);
    expect(await saved(session)).toBe('"One " [0 "par" | "agraph" 0] ^0 | [1 "blsgh " 1] ^1 |');
  });

  test('Enter inside a comment that opens its paragraph', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    enter(session, 1, 'bl'.length);
    expect(await saved(session)).toBe('"One " [0 "paragraph" 0] ^0 | [1 "bl" | "sgh " 1] ^1 |');
  });

  test('Enter after the comment\'s text, before its reference', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    enter(session, 0, 'One paragraph'.length);
    expect(await saved(session)).toBe('"One " [0 "paragraph" | 0] ^0 | [1 "blsgh " 1] ^1 |');
  });

  test('Enter twice inside the comment: the start in the first piece, the end in the last', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    enter(session, 0, 'One para'.length);
    enter(session, 0, 'One p'.length);
    expect(await saved(session)).toBe('"One " [0 "p" | "ara" | "graph" 0] ^0 | [1 "blsgh " 1] ^1 |');
  });

  test('Enter inside a comment that spans paragraphs, in its first paragraph', async () => {
    const session = await EditorDocument.open(await fixture('comments-span.docx'));
    const before = await saved(session);
    expect(before.startsWith('"One " [0 "paragraph" |')).toBe(true);
    enter(session, 0, 'One para'.length);
    expect((await saved(session)).startsWith('"One " [0 "para" | "graph" | "The middle of the range" | "Its end" 0] ^0 |')).toBe(true);
  });

  test('two paragraphs pasted before the comment: the comment goes on in the last of them', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    pasteTwo(session, 0, 'One'.length);
    expect(await saved(session)).toBe('"One" "first" | "second" " " [0 "paragraph" 0] ^0 | [1 "blsgh " 1] ^1 |');
  });

  test('two paragraphs pasted inside the comment', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    pasteTwo(session, 0, 'One par'.length);
    expect(await saved(session)).toMatch(/^"One " \[0 "par" "first" \| "second" "agraph" 0\] \^0 \|/);
  });

  test('plain text put inside a comment, as a paste puts it, is in the comment', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    edit(session, (state) => state.tr.insert(at(state.doc, 0, 'One par'.length), state.schema.text('x')));
    expect(await saved(session)).toBe('"One " [0 "par" "x" "agraph" 0] ^0 | [1 "blsgh " 1] ^1 |');
  });

  test('plain text at a comment\'s edges stays outside it', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    edit(session, (state) => state.tr.insert(at(state.doc, 0, 'One paragraph'.length), state.schema.text('y')).insert(at(state.doc, 0, 'One '.length), state.schema.text('x')));
    const marked = (text: string): boolean => {
      let found = false;
      session.doc.descendants((n) => { if (n.isText && n.text === text) found = n.marks.some((m) => m.type.name === 'comment'); return true; });
      return found;
    };
    expect(marked('x')).toBe(false);
    expect(marked('y')).toBe(false);
    await saved(session);
  });

  test('text put at the start of a paragraph a comment runs on into is in the comment', async () => {
    const session = await EditorDocument.open(await fixture('comments-span.docx'));
    edit(session, (state) => state.tr.insert(at(state.doc, 1, 0), state.schema.text('Pasted ')));
    expect((await saved(session)).startsWith('"One " [0 "paragraph" | "Pasted " "The middle of the range" | "Its end" 0] ^0 |')).toBe(true);
    let marked = false;
    session.doc.descendants((n) => { if (n.isText && n.text === 'Pasted ') marked = n.marks.some((m) => m.type.name === 'comment'); return true; });
    expect(marked).toBe(true);
  });

  test('with tracking on, text pasted inside the comment is tracked and in the comment', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    await runCommandAsync(session, 'document.setChangeTrackingMode', { value: 'TrackAll' });
    pasteTwo(session, 0, 'One par'.length);
    const xml = await mainXml((await session.save()).bytes);
    expect(shapeOf(xml)).toMatch(/^"One " \[0 "par" "first" \| "second" "agraph" 0\] \^0 \|/);
    expect(xml).toMatch(/<w:ins [^>]*><w:r><w:t>first<\/w:t>/);
  });

  test('text pasted just after a bookmark inside the comment is in it: the search steps over invisible atoms', async () => {
    const session = await patched('comments-two.docx', (xml) => xml.replace('<w:r><w:t>paragraph</w:t></w:r>', '<w:r><w:t>par</w:t></w:r><w:bookmarkStart w:id="9" w:name="inside"/><w:bookmarkEnd w:id="9"/><w:r><w:t>agraph</w:t></w:r>'));
    // After "One ", "par" and the two markers.
    edit(session, (state) => state.tr.insert(at(state.doc, 0, 'One par'.length + 2), state.schema.text('x')));
    expect(await saved(session)).toMatch(/^"One " \[0 "par" "x" "agraph" 0\] \^0 \|/);
  });

  test('a start the file puts after its paragraph\'s text, the comment\'s text in the next: an edit elsewhere leaves both as the file has them', async () => {
    const { session, xml: file } = await patchedFile('comments-span.docx', (xml) => xml.replace('<w:commentRangeStart w:id="0"/><w:r><w:t>paragraph</w:t></w:r>', '<w:r><w:t>paragraph</w:t></w:r><w:commentRangeStart w:id="0"/>'));
    expect(await mainXml((await session.save()).bytes)).toBe(file);
    const firstTwo = (x: string): string => x.slice(x.indexOf('<w:body>'), x.split('</w:p>', 2).join('</w:p>').length);
    edit(session, (state) => state.tr.insertText('!', at(state.doc, 3, 'blsgh'.length)));
    expect(firstTwo(await mainXml((await session.save()).bytes))).toBe(firstTwo(file));
  });

  test('a commented paragraph joined into the one before it: its markers come with its text', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    edit(session, (state) => state.tr.join(state.doc.child(0).nodeSize));
    expect(await saved(session)).toBe('"One " [0 "paragraph" 0] ^0 [1 "blsgh " 1] ^1 |');
  });

  test('the paragraph holding a comment\'s start deleted whole: the start goes to the text left, in a paragraph written as it was', async () => {
    const session = await EditorDocument.open(await fixture('comments-span.docx'));
    edit(session, (state) => state.tr.delete(0, state.doc.child(0).nodeSize));
    expect((await saved(session)).startsWith('[0 "The middle of the range" | "Its end" 0] ^0 |')).toBe(true);
  });

  test('the start moves into a paragraph inside a content control, which is written again around it', async () => {
    const session = await patched('comments-span.docx', (xml) => xml.replace('<w:p><w:r><w:t>The middle of the range</w:t></w:r></w:p>', '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>The middle of the range</w:t></w:r></w:p></w:sdtContent></w:sdt>'));
    edit(session, (state) => state.tr.delete(0, state.doc.child(0).nodeSize));
    expect((await saved(session)).startsWith('[0 "The middle of the range" | "Its end" 0] ^0 |')).toBe(true);
  });

  test('an edit elsewhere leaves the comment\'s paragraphs as the file has them', async () => {
    const original = await mainXml(await fixture('comments-span.docx'));
    const session = await EditorDocument.open(await fixture('comments-span.docx'));
    edit(session, (state) => state.tr.insertText('!', at(state.doc, 3, 'blsgh'.length)));
    const xml = await mainXml((await session.save()).bytes);
    const firstThree = (x: string): string => x.slice(x.indexOf('<w:body>'), x.split('</w:p>', 3).join('</w:p>').length);
    expect(firstThree(xml)).toBe(firstThree(original));
  });

  test('undo after Enter gives the file\'s own part back', async () => {
    const original = await mainXml(await fixture('comments-two.docx'));
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    enter(session, 0, 'One par'.length);
    session.undoManager.undo();
    expect(await mainXml((await session.save()).bytes)).toBe(original);
  });

  test('with tracking on, Enter inside the comment saves, its new mark tracked', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    await runCommandAsync(session, 'document.setChangeTrackingMode', { value: 'TrackAll' });
    enter(session, 0, 'One par'.length);
    const xml = await mainXml((await session.save()).bytes);
    expect(shapeOf(xml)).toBe('"One " [0 "par" | "agraph" 0] ^0 | [1 "blsgh " 1] ^1 |');
    expect(xml).toMatch(/<w:rPr><w:ins /);
  });
});
