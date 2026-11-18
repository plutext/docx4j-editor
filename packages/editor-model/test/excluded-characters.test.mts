// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.114: no character XML 1.0 excludes stays in the edit layer, or the document can no longer be saved.
// `api-cases.mts` holds the breaks and the dropped characters to the engine's writers; these are what it cannot reach.
import { describe, test, expect } from 'vitest';
import { Fragment, Slice, type Node as PMNode } from 'prosemirror-model';
import { ReplaceAroundStep } from 'prosemirror-transform';
import { ySyncPluginKey } from 'y-prosemirror';
import { EditorDocument, UNTRACKED_META, commentParagraphs, heldText, markdownDocument, recordTransaction, runCommand, runCommandAsync, stateFor } from '../src/index.mjs';
import { fixture, mainXml } from './helpers.mjs';

const of = String.fromCodePoint;
const open = async (): Promise<EditorDocument> => EditorDocument.open(await fixture('comments-two.docx'));
/** The first paragraph's position after `offset` characters; it opens with plain text. */
const at = (offset: number): number => 1 + offset;
const insert = (session: EditorDocument, text: string, from = at(3)): void => { runCommand(session, 'insertText', { text }, { from }); };
/** A paragraph's children: text as it stands, a break as `<kind>`. */
const shape = (node: PMNode): string[] => {
  const out: string[] = [];
  node.forEach((child) => out.push(child.isText ? child.text! : child.type.name === 'hardBreak' ? `<${child.attrs['kind'] as string}>` : child.type.name));
  return out;
};
const first = (session: EditorDocument): PMNode => session.doc.child(0);
/** A paragraph's text and breaks in order, whatever runs and `w:t`s hold them. */
const stream = (xml: string): string => [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>|<w:br(?: w:type="(\w+)")?\/>|<\/w:p>/g)]
  .map((m) => (m[0] === '</w:p>' ? '|' : m[1] ?? `<${m[2] ?? 'line'}>`)).join('');
const NOT_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/;

describe('a transaction', () => {
  test.each([['\u000B', '<line>'], ['\u000C', '<page>'], ['\u000E', '<column>']])('U+%s is a break', async (character, drawn) => {
    const session = await open();
    insert(session, `a${character}b`);
    expect(shape(first(session)).slice(0, 3)).toEqual(['Onea', drawn, expect.stringMatching(/^b/)]);
    expect(first(session).textContent).not.toMatch(NOT_XML);
    expect(await mainXml((await session.save()).bytes)).not.toMatch(NOT_XML);
  });

  test.each([['a\u000Bb'], ['a\u000Cb'], ['a\u000Eb'], ['a\u0002b\u0007'], ['\u000Bab\u000C\u000E']])('%j reads as the engine writes it', async (text) => {
    const command = await open();
    insert(command, text);
    const line = await open();
    await line.runOnTree(async ({ body }) => {
      const paragraph = body.paragraphs[0]!;
      const word = paragraph.search('One')[0]!;
      word.insertText(text, 'End');
    });
    expect(stream(await mainXml((await command.save()).bytes))).toBe(stream(await mainXml((await line.save()).bytes)));
  });

  test('the two hyphens are the characters Word shows them as, in the marks of the text', async () => {
    const session = await open();
    runCommand(session, 'font.bold', { value: true }, { from: at(0), to: at(3) });
    insert(session, `a\u001Eb\u001Fc`, at(1));
    expect(first(session).textContent.slice(0, 7)).toBe(`Oa${of(0x2011)}b${of(0xAD)}cn`);
    expect(first(session).child(0).marks.map((m) => m.type.name)).toContain('bold');
    expect(first(session).childCount).toBe(shape(first(session)).length);
    expect(shape(first(session))[0]).toContain(of(0x2011));
  });

  test.each([0x00, 0x01, 0x02, 0x08, 0x0F, 0x10, 0x1D, 0xFFFE, 0xFFFF, 0xD83D, 0xDE00])('U+%s is dropped', async (code) => {
    const session = await open();
    insert(session, `a${of(code)}b`);
    expect(first(session).textContent.slice(0, 5)).toBe('Oneab');
    await session.save();
  });

  test('what XML 1.0 allows is left: its edges, the controls it keeps, and a surrogate pair', async () => {
    const text = [0x09, 0x20, 0x7F, 0x85, 0xA0, 0xD7FF, 0xE000, 0xFFFD, 0x10000, 0x1F600, 0x10FFFF].map((c) => of(c)).join('');
    const session = await open();
    insert(session, text);
    expect(first(session).textContent).toContain(text);
  });

  test('half a surrogate pair is read with the half beside it: a step that swaps one half keeps the pair', async () => {
    const session = await open();
    insert(session, of(0x1F600));
    const state = stateFor(session, { from: at(3) });
    const { tr } = recordTransaction(state.tr.insertText(of(0xDE01), at(4), at(5)), state);
    expect(state.apply(tr).doc.child(0).textContent).toContain(of(0x1F601));
  });

  test('tracked, the text around it is left alone: the insertion, its break among it, is the only change', async () => {
    const session = await open();
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author C';
    insert(session, 'a\u000Bb');
    const children: PMNode[] = [];
    first(session).forEach((child) => children.push(child));
    const inserted = (node: PMNode): boolean => node.marks.some((m) => m.type.name === 'insertion') || (node.attrs['revision'] as { insertion?: unknown } | null | undefined)?.insertion !== undefined;
    expect(children.slice(0, 5).map((n) => [n.isText ? n.text : n.type.name, inserted(n)])).toEqual([
      ['One', false], ['a', true], ['hardBreak', true], ['b', true], [expect.any(String), false],
    ]);
    expect(children.some((n) => n.marks.some((m) => m.type.name === 'deletion'))).toBe(false);
  });

  test('it is the transaction\'s own: one undo step takes the text and its break', async () => {
    const session = await open();
    const before = session.doc;
    session.undoManager.stopCapturing();
    insert(session, 'a\u000Cb');
    expect(session.doc.eq(before)).toBe(false);
    session.undoManager.undo();
    expect(session.doc.eq(before)).toBe(true);
  });

  test('positions hold across steps: a later step before it, and several in one step', async () => {
    const session = await open();
    const state = stateFor(session, { from: at(3) });
    const tr = state.tr.insertText('a\u0007b\u000Bc\u0002', at(3)).insertText('earlier ', at(0));
    const { tr: written } = recordTransaction(tr, state);
    expect(shape(state.apply(written).doc.child(0)).slice(0, 3)).toEqual(['earlier Oneab', '<line>', expect.stringMatching(/^c/)]);
  });

  test('a later step that puts a paragraph before it: the text is found where it now stands', async () => {
    const session = await open();
    const state = stateFor(session, { from: at(3) });
    const before = state.schema.nodes['paragraph']!.create(null, state.schema.text('a new paragraph, longer than what was inserted'));
    const { tr } = recordTransaction(state.tr.insertText('a\u000Bb', at(3)).insert(0, before), state);
    const doc = state.apply(tr).doc;
    expect(shape(doc.child(1)).slice(0, 2)).toEqual(['Onea', '<line>']);
    expect(doc.textContent).not.toMatch(NOT_XML);
  });

  test('text inside the paragraphs a step inserts, as a paste of several has it', async () => {
    const session = await open();
    const state = stateFor(session, { from: at(3) });
    const paragraph = (text: string): PMNode => state.schema.nodes['paragraph']!.create(null, state.schema.text(text));
    const { tr } = recordTransaction(state.tr.insert(0, [paragraph('one\u000Btwo'), paragraph('three\u0002')]), state);
    const doc = state.apply(tr).doc;
    expect(shape(doc.child(0))).toEqual(['one', '<line>', 'two']);
    expect(shape(doc.child(1))).toEqual(['three']);
  });

  test('a step that wraps content in what it inserts', async () => {
    const session = await open();
    const state = stateFor(session, { from: at(3) });
    const slice = new Slice(Fragment.from(state.schema.text('x\u0002y')), 0, 0);
    const tr = state.tr.step(new ReplaceAroundStep(at(0), at(3), at(0), at(3), slice, 1));
    const { tr: written } = recordTransaction(tr, state);
    expect(state.apply(written).doc.child(0).textContent.slice(0, 5)).toBe('xOney');
  });

  test('a later step that takes the text away leaves nothing to do', async () => {
    const session = await open();
    const state = stateFor(session, { from: at(3) });
    const tr = state.tr.insertText('a\u000Bb', at(3)).delete(at(3), at(6));
    const { tr: written } = recordTransaction(tr, state);
    expect(state.apply(written).doc.eq(state.doc)).toBe(true);
  });

  test('an untracked transaction is written too; a peer\'s is left as it came, or the two stores would part', async () => {
    const session = await open();
    const state = stateFor(session, { from: at(3) });
    const local = recordTransaction(state.tr.insertText('a\u000Bb', at(3)).setMeta(UNTRACKED_META, true), state).tr;
    expect(shape(state.apply(local).doc.child(0))[1]).toBe('<line>');
    const remote = recordTransaction(state.tr.insertText('a\u000Bb', at(3)).setMeta(ySyncPluginKey, { isChangeOrigin: true }), state).tr;
    expect(state.apply(remote).doc.child(0).textContent).toContain('a\u000Bb');
  });
});

describe('a text compared with what the edit layer holds', () => {
  test('a break is a line feed, a hyphen its character, the rest gone, and what XML allows left', () => {
    expect(heldText('a\u000Bb\u000Cc\u000Ed\u0002e\u001Ef\u0007')).toBe(`a\nb\nc\nd` + `e${of(0x2011)}f`);
    expect(heldText('plain\ttext')).toBe('plain\ttext');
  });
});

describe('text that passes through no transaction', () => {
  test('a Markdown file: in a paragraph, a heading, a list, a code block and a table cell, and it saves', async () => {
    const { bytes } = await markdownDocument([
      '# Head\u000Bline', '', 'first\u000Bsecond', '', 'before\u000Cafter', '', 'left\u000Eright', '', 'no\u001Ebreak soft\u001Fhyphen bell\u0007end',
      '', '- item\u000Bmore', '', '```', 'code\u000Cnext', '```', '', '| a\u000Bb | c |', '| - | - |', '| d | e\u0002f |', '',
    ].join('\n'));
    const xml = await mainXml(bytes);
    expect(xml).not.toMatch(NOT_XML);
    expect(xml.match(/<w:br w:type="page"\/>/g)).toHaveLength(2);
    expect(xml.match(/<w:br w:type="column"\/>/g)).toHaveLength(1);
    expect(xml).toContain(`no${of(0x2011)}break soft${of(0xAD)}hyphen bellend`);
    expect(xml).toContain('ef');
    const session = await EditorDocument.open(bytes);
    expect(shape(session.doc.child(1))).toEqual(['first', '<line>', 'second']);
    expect(shape(session.doc.child(2))).toEqual(['before', '<page>', 'after']);
  });

  test('without such a character a run is what `r` makes of it, a tab and a line break included', async () => {
    const { bytes } = await markdownDocument('plain words  \nnext\tline\n');
    const xml = await mainXml(bytes);
    expect(xml.match(/<w:br\/>/g)).toHaveLength(1);
    expect(xml).toMatch(/next<\/w:t><w:tab\/><w:t[^>]*>line<\/w:t>/);
  });

  test('a comment\'s text, and the save that writes it', async () => {
    const [paragraph] = commentParagraphs('a\u000Bb\u0002');
    const runs = paragraph!.value.content as { value: { content: { name: { localPart: string } }[] } }[];
    expect(runs[1]!.value.content.map((item) => item.name.localPart)).toEqual(['t', 'br', 't']);
    const session = await open();
    await runCommandAsync(session, 'range.insertComment', { text: 'seen\u000Bhere\u0007' }, { from: at(0), to: at(3) });
    await session.save();
  });
});
