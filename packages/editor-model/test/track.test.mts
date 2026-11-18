// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 5.4, step C3: the tracker. Each case runs a command, or a transaction as the
// view's input makes one (typing, a character deleted), with tracking on, through the one
// function the view and the headless runner both record transactions with, and reads the marks
// and the saved part: typed text inserted, deleted text kept and marked, this author's insertion
// deleted for real, a paragraph split and joined as marks, formatting and paragraph properties
// recorded once and dropped when changed back, rows, what runs untracked, tracking off, the mode
// written to the settings part, and a tree run tracked by the engine as the same author.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage, ZipPartStore } from '@docx4j/core-ts';
import type { EditorState, Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, runCommandAsync, stateFor, recordTransaction, formattingKey, type ParagraphAttrs, type Recorded } from '../src/index.mjs';
import { fixture, dated } from './helpers.mjs';
import { room } from './room-harness.mjs';

const DATE = '2026-09-28T10:00:00Z';
const C = { author: 'Author C', date: DATE, move: null };

async function tracking(name = 'tracked-changes.docx'): Promise<EditorDocument> {
  const session = await EditorDocument.open(await fixture(name));
  session.trackingMode = 'TrackAll';
  session.trackingAuthor = 'Author C';
  session.trackedChangeDate = new Date(DATE);
  return session;
}

/** A transaction as the view's input makes one (typing, a key the browser handles), recorded as the view records it, and applied. */
function input(session: EditorDocument, make: (state: EditorState) => Transaction, selection?: { from: number; to?: number }): Recorded & { state: EditorState } {
  const state = stateFor(session, selection);
  const recorded = recordTransaction(make(state), state);
  const next = state.apply(recorded.tr);
  session.replace(next.doc);
  return { ...recorded, state: next };
}

/** The first top-level paragraph whose text contains `text`, with its position and index. */
function paragraphWith(doc: PMNode, text: string): { node: PMNode; pos: number; index: number } {
  let found: { node: PMNode; pos: number; index: number } | undefined;
  doc.forEach((node, pos, index) => { if (!found && node.type.name === 'paragraph' && node.textContent.includes(text)) found = { node, pos, index }; });
  if (!found) throw new Error(`no paragraph with "${text}"`);
  return found;
}

/** Where `text` starts inside the paragraph holding it. */
function at(doc: PMNode, text: string, within = text): number {
  const p = paragraphWith(doc, within);
  let pos = -1;
  p.node.descendants((n, rel) => { if (pos < 0 && n.isText && n.text!.includes(text)) pos = p.pos + 1 + rel + n.text!.indexOf(text); return pos < 0; });
  if (pos < 0) throw new Error(`"${text}" is not in one text node`);
  return pos;
}

/** A paragraph's text pieces with their revisions, as `text|ins:author|del:author|fc`. */
function pieces(node: PMNode): string[] {
  const out: string[] = [];
  node.forEach((n) => {
    if (!n.isText) return;
    const tags = n.marks.flatMap((m) => m.type.name === 'insertion' ? [`ins:${String(m.attrs['author'])}`] : m.type.name === 'deletion' ? [`del:${String(m.attrs['author'])}`] : m.type.name === 'formatChange' ? ['fc'] : []);
    out.push([n.text, ...tags].join('|'));
  });
  return out;
}

async function savedMain(session: EditorDocument): Promise<string> {
  const saved = await session.save();
  expect(saved.selfCheck?.equal).toBe(true);
  return (await (await WordprocessingMLPackage.load(saved.bytes)).getMainDocumentPart().getXml()).replace(/ w:rsid\w*="[^"]*"/g, '');
}

describe('text (ED-005 section 5.4)', () => {
  test('typing: the text is inserted by the author at the minute, and saved in a w:ins', async () => {
    const session = await tracking();
    const r = runCommand(session, 'insertText', { text: 'new ' }, { from: at(session.doc, 'text.') });
    expect(r.ok).toBe(true);
    expect(pieces(paragraphWith(session.doc, 'Some').node)).toEqual(['Some ', 'new |ins:Author C', 'text.']);
    expect(await savedMain(session)).toMatch(dated(/<w:ins w:id="\d+" @C><w:r><w:t xml:space="preserve">new <\/w:t><\/w:r><\/w:ins>/));
  });

  test('typing over a selection: what was selected is kept, deleted, before what replaced it; the caret after the new text', async () => {
    const session = await tracking();
    const from = at(session.doc, 'text');
    const r = runCommand(session, 'insertText', { text: 'words' }, { from, to: from + 4 });
    expect(pieces(paragraphWith(session.doc, 'Some').node)).toEqual(['Some ', 'text|del:Author C', 'words|ins:Author C', '.']);
    expect(r.state!.selection.from).toBe(from + 4 + 5);
    expect(await savedMain(session)).toMatch(/<w:del w:id="\d+"(?: w16du:dateUtc="[^"]*")? w:author="Author C"[^>]*><w:r><w:delText>text<\/w:delText><\/w:r><\/w:del><w:ins w:id="\d+"(?: w16du:dateUtc="[^"]*")? w:author="Author C"[^>]*><w:r><w:t>words<\/w:t><\/w:r><\/w:ins>/);
  });

  test('Backspace keeps the character, deleted, and the caret goes before it; Delete puts the caret after it', async () => {
    const session = await tracking();
    const x = at(session.doc, 'text');
    const back = input(session, (s) => s.tr.delete(x + 3, x + 4), { from: x + 4 });
    expect(pieces(paragraphWith(session.doc, 'Some').node)).toEqual(['Some tex', 't|del:Author C', '.']);
    expect(back.state.selection.from).toBe(x + 3);
    const forward = input(session, (s) => s.tr.delete(x, x + 1), { from: x });
    expect(pieces(paragraphWith(session.doc, 'Some').node)).toEqual(['Some ', 't|del:Author C', 'ex', 't|del:Author C', '.']);
    expect(forward.state.selection.from).toBe(x + 1);
  });

  test("deleting this author's own insertion removes it; another author's is marked deleted as well", async () => {
    const session = await tracking();
    runCommand(session, 'insertText', { text: 'mine ' }, { from: at(session.doc, 'text.') });
    const mine = at(session.doc, 'mine');
    input(session, (s) => s.tr.delete(mine, mine + 5), { from: mine + 5 });
    expect(pieces(paragraphWith(session.doc, 'Some').node)).toEqual(['Some text.']);
    const theirs = at(session.doc, 'insertion', 'An insertion');
    input(session, (s) => s.tr.delete(theirs, theirs + 9), { from: theirs + 9 });
    expect(pieces(paragraphWith(session.doc, 'An insertion').node)).toContain('insertion|ins:Jason Harrop|del:Author C');
  });

  test('deleting text already deleted leaves it as it was', async () => {
    const session = await tracking();
    const p = paragraphWith(session.doc, 'A deletion');
    const d = at(session.doc, ' A deletion');
    input(session, (s) => s.tr.delete(d, d + 11), { from: d + 11 });
    expect(pieces(session.doc.nodeAt(p.pos)!)).toEqual(pieces(p.node));
  });

  test('replace all with tracking on: each hit deleted whole and its replacement inserted', async () => {
    const session = await tracking();
    const r = runCommand(session, 'find.replaceAll', { text: 'Some', replacement: 'Any' }, { from: 1 });
    expect(r.ok).toBe(true);
    expect(pieces(paragraphWith(session.doc, 'text.').node).slice(0, 2)).toEqual(['Some|del:Author C', 'Any|ins:Author C']);
  });
});

describe('paragraph marks (ED-005 section 5.4)', () => {
  test('Enter: the paragraph whose mark it made has it inserted, and Word\'s w:pPr/w:rPr/w:ins is saved', async () => {
    const session = await tracking();
    runCommand(session, 'insertParagraph', undefined, { from: at(session.doc, 'text.') });
    const first = paragraphWith(session.doc, 'Some ');
    expect((first.node.attrs as ParagraphAttrs).markRevision).toEqual({ insertion: C });
    expect((session.doc.child(first.index + 1).attrs as ParagraphAttrs).markRevision).toBeNull();
    expect(await savedMain(session)).toMatch(dated(/<w:p><w:pPr><w:rPr><w:ins w:id="\d+" @C\/><\/w:rPr><\/w:pPr><w:r><w:t xml:space="preserve">Some <\/w:t><\/w:r><\/w:p>/));
  });

  test("Enter in a paragraph whose mark is another author's insertion and carries a keep and a bold: the first half keeps the keep and the bold with its own new w:ins, never the copied one (ED-005 section 12.102)", async () => {
    const pkg = await WordprocessingMLPackage.load(await fixture('tracked-changes.docx'));
    const part = pkg.getMainDocumentPart();
    const xml = await part.getXml();
    const patched = xml.replace('<w:r><w:t>Some text.</w:t></w:r>', '<w:pPr><w:keepNext/><w:rPr><w:ins w:id="900" w:author="Author Z" w:date="2026-01-01T00:00:00Z"/><w:b/></w:rPr></w:pPr><w:r><w:t>Some text.</w:t></w:r>');
    expect(patched).not.toBe(xml);
    part.setXml(patched);
    const session = await EditorDocument.open(await pkg.save());
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author C';
    session.trackedChangeDate = new Date(DATE);
    runCommand(session, 'insertParagraph', undefined, { from: at(session.doc, 'text.') });
    const first = paragraphWith(session.doc, 'Some ');
    expect((first.node.attrs as ParagraphAttrs).markRevision).toEqual({ insertion: C });
    const main = await savedMain(session);
    const paragraphs = main.match(/<w:p>.*?<\/w:p>/g)!.filter((p) => p.includes('>Some <') || p.includes('>text.<'));
    expect(paragraphs).toHaveLength(2);
    const [a, b] = paragraphs as [string, string];
    expect(a).toContain('<w:keepNext/>');
    expect(a).toContain('<w:b/>');
    expect(a).toMatch(dated(/<w:ins w:id="(?!900")\d+" @C\/>/));
    expect(a).not.toContain('Author Z');
    expect(b).toContain('<w:keepNext/>');
    expect(b).toContain('<w:ins w:id="900" w:author="Author Z"');
  });

  test('Backspace at a paragraph\'s start marks the mark before it deleted and joins nothing; the caret goes before that mark, into the paragraph before (an empty one here)', async () => {
    const session = await tracking();
    const second = paragraphWith(session.doc, 'Some text.');
    const count = session.doc.childCount;
    expect(session.doc.child(second.index - 1).content.size).toBe(0);
    const r = runCommand(session, 'joinBackward', undefined, { from: second.pos + 1 });
    expect(r.ok).toBe(true);
    expect(session.doc.childCount).toBe(count);
    const before = session.doc.child(second.index - 1);
    expect((before.attrs as ParagraphAttrs).markRevision).toEqual({ deletion: C });
    expect(r.state!.selection.from).toBe(second.pos - 1);
    expect(await savedMain(session)).toMatch(dated(/<w:rPr><w:del w:id="\d+" @C\/><\/w:rPr>/));
  });

  test('Backspace between two paragraphs with text: the first paragraph\'s mark deleted, the caret at the end of its text, each keeping its section', async () => {
    const session = await tracking();
    const second = paragraphWith(session.doc, 'Tables');
    const first = session.doc.child(second.index - 1);
    expect(first.content.size).toBeGreaterThan(0);
    const r = runCommand(session, 'joinBackward', undefined, { from: second.pos + 1 });
    expect(r.ok).toBe(true);
    expect((session.doc.child(second.index - 1).attrs as ParagraphAttrs).markRevision).toEqual({ deletion: C });
    expect(session.doc.child(second.index - 1).attrs['sectPr']).toEqual(first.attrs['sectPr']);
    expect(r.state!.selection.from).toBe(second.pos - 1);
  });

  test("Enter in a section's last paragraph, tracked: the section stays with the old mark, and the file holds no section more (ED-005 section 12.17)", async () => {
    for (const where of ['middle', 'end'] as const) {
      const session = await tracking('notes-numbering.docx');
      let target: { pos: number; node: PMNode; index: number } | undefined;
      session.doc.forEach((node, pos, index) => { if (!target && node.type.name === 'paragraph' && node.attrs['sectPr'] && node.textContent.length > 4) target = { pos, node, index }; });
      const sections = (xml: string): number => xml.match(/<w:sectPr\b/g)!.length;
      const before = sections(await savedMain(await tracking('notes-numbering.docx')));
      runCommand(session, 'insertParagraph', undefined, { from: target!.pos + 1 + (where === 'middle' ? 2 : target!.node.content.size) });
      expect(session.doc.child(target!.index).attrs['sectPr']).toBeNull();
      expect(session.doc.child(target!.index + 1).attrs['sectPr']).toEqual(target!.node.attrs['sectPr']);
      expect(sections(await savedMain(session))).toBe(before);
    }
  });

  test("a paragraph mark this author inserted, deleted: the paragraphs join for real", async () => {
    const session = await tracking();
    const count = session.doc.childCount;
    runCommand(session, 'insertParagraph', undefined, { from: at(session.doc, 'text.') });
    const second = paragraphWith(session.doc, 'text.');
    runCommand(session, 'joinBackward', undefined, { from: second.pos + 1 });
    expect(session.doc.childCount).toBe(count);
    expect(pieces(paragraphWith(session.doc, 'Some').node)).toEqual(['Some text.']);
  });
});

describe("Enter as Word does it: the new mark at the caret, at a paragraph's end too (core-ts check 29; ED-005 section 12.17)", () => {
  /** Each top-level paragraph as the comparison reads it: text, style, its mark's revision, a recorded properties change, its text's revisions. */
  const summary = (doc: PMNode): string[] => {
    const out: string[] = [];
    doc.forEach((n) => {
      if (n.type.name !== 'paragraph') return;
      const a = n.attrs as ParagraphAttrs;
      const mark = a.markRevision ? Object.entries(a.markRevision).map(([side, info]) => `${side}:${(info as { author: string }).author}`).join('+') : '-';
      const change = a.pPrChange ? `${a.pPrChange.author}:${String((a.pPrChange.before as { pStyle?: { val?: string } } | undefined)?.pStyle?.val ?? 'Normal')}` : '-';
      const text = new Set<string>();
      n.forEach((c) => { for (const m of c.marks) if (m.type.name === 'insertion' || m.type.name === 'deletion') text.add(`${m.type.name}:${String(m.attrs['author'])}`); });
      out.push([n.textContent, a.pStyle ?? 'Normal', mark, change, [...text].sort().join(',') || '-'].join(' | '));
    });
    return out;
  };

  /** Enter where Word's case pressed it, "New" typed after it unless `typed` is false; tracked as Author B or not. */
  const edited = async (tracked: boolean, where: (doc: PMNode) => number, typed = true): Promise<string[]> => {
    const session = await tracking('revisions/check29/29-input.docx');
    session.trackingMode = tracked ? 'TrackAll' : 'Off';
    session.trackingAuthor = 'Author B';
    const r = runCommand(session, 'insertParagraph', undefined, { from: where(session.doc) });
    expect(r.ok).toBe(true);
    if (typed) runCommand(session, 'insertText', { text: 'New' }, { from: r.state!.selection.from });
    const saved = await session.save();
    expect(saved.selfCheck?.equal).toBe(true);
    return summary((await EditorDocument.open(saved.bytes)).doc);
  };
  const word = async (name: string): Promise<string[]> => summary((await EditorDocument.open(await fixture(`revisions/check29/${name}.docx`))).doc);
  const endOf = (text: string) => (doc: PMNode): number => { const p = paragraphWith(doc, text); return p.pos + p.node.nodeSize - 1; };

  test('29a: at the end of a Heading 1, tracked: the heading\'s mark inserted, the new line Normal with the heading\'s style recorded', async () => {
    expect(await edited(true, endOf('Case a heading'))).toEqual(await word('29a-heading-end-tracked'));
  });

  test('29b: at the end of a Normal paragraph, tracked: its mark inserted, the new line no change of properties', async () => {
    expect(await edited(true, endOf('Case b: the end'))).toEqual(await word('29b-normal-end-tracked'));
  });

  test("29c: just before a mark Author A inserted, untracked: a plain mark before it, the new line keeps A's", async () => {
    expect(await edited(false, endOf('Cases c and d'))).toEqual(await word('29c-inserted-mark-end-untracked'));
  });

  test("29d: just before a mark Author A inserted, tracked: B's mark before it, the new line keeps A's", async () => {
    expect(await edited(true, endOf('Cases c and d'))).toEqual(await word('29d-inserted-mark-end-tracked'));
  });

  test("29e: in the middle of a paragraph whose mark Author A inserted, untracked: the first half's mark plain, the second keeps A's", async () => {
    expect(await edited(false, (doc) => at(doc, ' after the split', 'Case e: before the split'), false)).toEqual(await word('29e-inserted-mark-middle-untracked'));
  });
});

describe('formatting and paragraph properties (ED-005 section 5.4)', () => {
  test('bold: the formatting before is recorded once, and saved as w:rPrChange; bold taken off again, the record goes', async () => {
    const session = await tracking();
    const x = at(session.doc, 'text');
    runCommand(session, 'font.bold', { value: true }, { from: x, to: x + 4 });
    runCommand(session, 'font.italic', { value: true }, { from: x, to: x + 4 });
    const node = session.doc.nodeAt(x)!;
    const change = node.marks.find((m) => m.type.name === 'formatChange')!;
    expect(change.attrs).toMatchObject({ author: 'Author C', date: DATE });
    expect(JSON.parse(change.attrs['before'] as string)).toEqual({ TYPE_NAME: 'org_docx4j_wml.CTRPrChange.RPr' });
    expect(await savedMain(session)).toMatch(dated(/<w:rPr><w:b\/><w:bCs\/><w:i\/><w:iCs\/><w:rPrChange w:id="\d+" @C><w:rPr\/><\/w:rPrChange><\/w:rPr><w:t>text<\/w:t>/));
    runCommand(session, 'font.bold', { value: false }, { from: x, to: x + 4 });
    runCommand(session, 'font.italic', { value: false }, { from: x, to: x + 4 });
    expect(session.doc.nodeAt(x)!.marks.some((m) => m.type.name === 'formatChange')).toBe(false);
  });

  // UPSTREAM-15: y-prosemirror's diff took a text whose bold went as its formatChange came for unchanged,
  // so the store (and so the save and every peer) kept the bold. Found 2026-09-30 by proposal 43's tests.
  for (const name of ['font.bold', 'font.italic', 'font.rtl'] as const) {
    test(`${name} stated, then taken off with tracking on: the store, the save and a peer have it`, async () => {
      const base = await fixture('tracked-changes.docx');
      const { creator, joiner } = await room(base);
      const x = at(creator.doc, 'text');
      runCommand(creator, name, { value: true }, { from: x, to: x + 4 });
      creator.trackingMode = 'TrackAll';
      creator.trackingAuthor = 'Author C';
      creator.trackedChangeDate = new Date(DATE);
      runCommand(creator, name, { value: false }, { from: x, to: x + 4 });
      const mark = name === 'font.bold' ? 'bold' : name === 'font.italic' ? 'italic' : 'rtl';
      for (const session of [creator, joiner]) {
        const marks = session.doc.nodeAt(x)!.marks.map((m) => m.type.name);
        expect(marks).toContain('formatChange');
        expect(marks).not.toContain(mark);
      }
      expect(await savedMain(creator)).toMatch(/<w:rPrChange [^>]*><w:rPr>.*<\/w:rPr><\/w:rPrChange><\/w:rPr><w:t>text<\/w:t>/);
    });
  }

  // The save's self-check compared a formatting change's `before` as a string: the tracker's names hold a
  // namespace and a local name, a reloaded file's its prefix too, so any tracked formatting of text that had
  // direct formatting refused the save (found with UPSTREAM-15, 2026-09-30; ED-005 section 12.40).
  test('italic made with tracking on over text bold before: the save passes its self-check, the rPrChange holds the bold', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const x = at(session.doc, 'text');
    runCommand(session, 'font.bold', { value: true }, { from: x, to: x + 4 });
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author C';
    session.trackedChangeDate = new Date(DATE);
    runCommand(session, 'font.italic', { value: true }, { from: x, to: x + 4 });
    expect(await savedMain(session)).toMatch(/<w:rPr><w:b\/><w:bCs\/><w:i\/><w:iCs\/><w:rPrChange [^>]*><w:rPr><w:b\/><w:bCs\/><\/w:rPr><\/w:rPrChange><\/w:rPr><w:t>text<\/w:t>/);
  });

  // Word's check 24 (core-ts 03b4d9f): bold made and taken off again by its author (24b) or by another (24c), and a
  // character style applied and cleared with Ctrl+Space (24e), each leave no w:rPrChange: the record goes whoever
  // changes the formatting back, as the tracker does by value.
  for (const [from, author, name, args, word, text] of [
    ['24a-bold', 'Author A', 'font.bold', { value: false }, '24b-unbold', 'beta'],
    ['24a-bold', 'Author B', 'font.bold', { value: false }, '24c-unbold-by-b', 'beta'],
    ['24d-style', 'Author A', 'font.clear', undefined, '24e-style-cleared', 'gamma'],
  ] as const) {
    test(`formatting changed back by ${author} (${name}): no record, as Word's ${word}.docx`, async () => {
      const session = await tracking(`revisions/check24/${from}.docx`);
      session.trackingAuthor = author;
      const x = at(session.doc, text);
      expect(session.doc.nodeAt(x)!.marks.some((m) => m.type.name === 'formatChange')).toBe(true);
      expect(runCommand(session, name, args, { from: x, to: x + text.length }).ok).toBe(true);
      const runs = (doc: PMNode): string[] => { const out: string[] = []; paragraphWith(doc, text).node.forEach((n) => { if (n.isText) out.push(`${n.text}|${formattingKey(n.marks)}|${pieces(doc.type.schema.node('paragraph', null, [n]))[0]}`); }); return out; };
      const theirs = await EditorDocument.open(await fixture(`revisions/check24/${word}.docx`));
      expect(runs(session.doc).join('')).toBe(runs(theirs.doc).join(''));
      expect(await savedMain(session)).not.toContain('rPrChange');
    });
  }

  // Word's check 26 (core-ts bce8206): a paragraph centred and set left again by its author (26b) or by another (26c), and
  // indented and indented back (26d-indented-back), each leave no w:pPrChange and no w:pPr: Ctrl+L and Decrease Indent
  // remove the value the style gives, and the record goes.
  for (const [from, author, name, args, word, text] of [
    ['26a-centred', 'Author A', 'paragraph.alignment', { value: 'Left' }, '26b-left-again', 'The first paragraph.'],
    ['26a-centred', 'Author B', 'paragraph.alignment', { value: 'Left' }, '26c-left-by-b', 'The first paragraph.'],
    ['26d-indented', 'Author A', 'paragraph.decreaseIndent', undefined, '26d-indented-back', 'The second paragraph.'],
  ] as const) {
    test(`paragraph properties changed back by ${author} (${name}): no record, as Word's ${word}.docx`, async () => {
      const session = await tracking(`revisions/check26/${from}.docx`);
      session.trackingAuthor = author;
      const p = paragraphWith(session.doc, text);
      expect((p.node.attrs as ParagraphAttrs).pPrChange).not.toBeNull();
      expect(runCommand(session, name, args, { from: p.pos + 1 }).ok).toBe(true);
      const after = paragraphWith(session.doc, text).node.attrs as ParagraphAttrs;
      const theirs = paragraphWith((await EditorDocument.open(await fixture(`revisions/check26/${word}.docx`))).doc, text).node.attrs as ParagraphAttrs;
      expect({ jc: after.jc, ind: after.ind, pPrChange: after.pPrChange }).toEqual({ jc: theirs.jc, ind: theirs.ind, pPrChange: theirs.pPrChange });
      expect(await savedMain(session)).not.toContain('pPrChange');
    });
  }

  test("text this author inserted takes formatting with no record", async () => {
    const session = await tracking();
    runCommand(session, 'insertText', { text: 'new' }, { from: at(session.doc, 'text.') });
    const x = at(session.doc, 'new');
    runCommand(session, 'font.bold', { value: true }, { from: x, to: x + 3 });
    expect(session.doc.nodeAt(x)!.marks.some((m) => m.type.name === 'formatChange')).toBe(false);
  });

  test('alignment: the paragraph\'s properties before in pPrChange, saved as w:pPrChange; changed back, the record goes', async () => {
    const session = await tracking();
    const p = paragraphWith(session.doc, 'Some text.');
    runCommand(session, 'paragraph.alignment', { value: 'Centered' }, { from: p.pos + 1 });
    const attrs = session.doc.nodeAt(p.pos)!.attrs as ParagraphAttrs;
    expect(attrs.pPrChange).toMatchObject({ author: 'Author C', date: DATE });
    expect(await savedMain(session)).toMatch(dated(/<w:pPr><w:jc w:val="center"\/><w:pPrChange w:id="\d+" @C><w:pPr\/><\/w:pPrChange><\/w:pPr><w:r><w:t>Some text\.<\/w:t>/));
    // A paragraph justified in the file, centred and justified again: its record goes.
    const j = paragraphWith(session.doc, 'change tracking');
    expect(j.node.attrs['jc']).toBe('both');
    runCommand(session, 'paragraph.alignment', { value: 'Centered' }, { from: j.pos + 1 });
    expect((session.doc.nodeAt(j.pos)!.attrs as ParagraphAttrs).pPrChange).not.toBeNull();
    runCommand(session, 'paragraph.alignment', { value: 'Justified' }, { from: j.pos + 1 });
    expect((session.doc.nodeAt(j.pos)!.attrs as ParagraphAttrs).pPrChange).toBeNull();
  });
  test("bold over a tab and a line break: the atoms are not replaced, and their formatting goes unrecorded (ED-005 section 12.5 item 7, section 12.16)", async () => {
    const session = await tracking('revisions/revisions-word15.docx');
    session.trackingMode = 'Off';
    const p = paragraphWith(session.doc, 'Split this paragraph');
    const tab = at(session.doc, 'Split', 'Split this paragraph') + 5;
    runCommand(session, 'insertTab', undefined, { from: tab });
    runCommand(session, 'insertBreak', { breakType: 'Line' }, { from: tab + 1 + ' this'.length });
    session.trackingMode = 'TrackAll';
    const count = paragraphWith(session.doc, 'Split').node.childCount;
    runCommand(session, 'font.bold', { value: true }, { from: p.pos + 1, to: p.pos + paragraphWith(session.doc, 'Split').node.nodeSize - 1 });
    const after = paragraphWith(session.doc, 'Split').node;
    expect(after.childCount).toBe(count);
    const atoms: string[] = [];
    after.forEach((n) => {
      if (n.isText) expect(n.marks.some((m) => m.type.name === 'formatChange')).toBe(true);
      else atoms.push(`${n.type.name} ${JSON.stringify(n.attrs['revision'])} ${String(n.attrs['fmt'])}`);
    });
    expect(atoms).toEqual(['tab null [["bold",{}]]', 'hardBreak null [["bold",{}]]']);
  });
});

describe('rows, and what runs untracked (ED-005 section 5.4, proposal 5)', () => {
  /** The position inside the first cell of the table's row `r`. */
  function inRow(doc: PMNode, r: number): number {
    let table = -1;
    doc.forEach((n, pos) => { if (table < 0 && n.type.name === 'table') table = pos; });
    let pos = table + 1;
    const t = doc.nodeAt(table)!;
    for (let k = 0; k < r; k++) pos += t.child(k).nodeSize;
    return pos + 3;
  }
  const tableOf = (doc: PMNode): PMNode => { let t: PMNode | undefined; doc.forEach((n) => { if (!t && n.type.name === 'table') t = n; }); return t!; };
  const rowOf = (doc: PMNode, r: number): PMNode => tableOf(doc).child(r);

  test('a row deleted is kept, marked deleted, its text deleted; saved as w:trPr/w:del', async () => {
    const session = await tracking();
    const rows = tableOf(session.doc).childCount;
    const r = runCommand(session, 'table.deleteRows', undefined, { from: inRow(session.doc, 1) });
    expect(r.ok).toBe(true);
    expect(tableOf(session.doc).childCount).toBe(rows);
    expect(rowOf(session.doc, 1).attrs['revision']).toEqual({ deletion: C });
    rowOf(session.doc, 1).descendants((n) => { if (n.isText) expect(n.marks.some((m) => m.type.name === 'deletion')).toBe(true); return true; });
    expect(await savedMain(session)).toMatch(dated(/<w:trPr><w:del w:id="\d+" @C\/><\/w:trPr>/));
  });

  test('the row a vertical merge starts in, deleted: that row alone is marked, and the merge stays whole', async () => {
    const session = await tracking();
    expect(rowOf(session.doc, 1).child(0).attrs['rowspan']).toBe(2);
    runCommand(session, 'table.deleteRows', undefined, { from: inRow(session.doc, 1) });
    expect(tableOf(session.doc).childCount).toBe(4);
    expect(rowOf(session.doc, 1).attrs['revision']).toEqual({ deletion: C });
    expect(rowOf(session.doc, 1).child(0).attrs['rowspan']).toBe(2);
    expect(rowOf(session.doc, 2).attrs['revision']).toBeNull();
    rowOf(session.doc, 2).descendants((n) => { if (n.isText) expect(n.marks.some((m) => m.type.name === 'deletion')).toBe(false); return true; });
    await savedMain(session);
  });

  test("a row this author added, deleted: it goes for real", async () => {
    const session = await tracking();
    const rows = tableOf(session.doc).childCount;
    runCommand(session, 'table.addRows', { location: 'After' }, { from: inRow(session.doc, 0) });
    expect(tableOf(session.doc).childCount).toBe(rows + 1);
    runCommand(session, 'table.deleteRows', undefined, { from: inRow(session.doc, 1) });
    expect(tableOf(session.doc).childCount).toBe(rows);
  });

  test('a row added is marked inserted', async () => {
    const session = await tracking();
    runCommand(session, 'table.addRows', { location: 'After' }, { from: inRow(session.doc, 0) });
    expect(rowOf(session.doc, 1).attrs['revision']).toEqual({ insertion: C });
  });

  test('an untracked command (a column added) runs untracked and says so; a story runs untracked', async () => {
    const session = await tracking();
    const r = runCommand(session, 'table.addColumns', undefined, { from: inRow(session.doc, 0) });
    expect(r.ok).toBe(true);
    expect(r.untracked).toBe('table.addColumns');
    let marked = false;
    session.doc.descendants((n) => { if (n.marks.some((m) => m.type.name === 'insertion' && m.attrs['author'] === 'Author C') || (n.attrs['revision'] as { insertion?: unknown } | null)?.insertion) marked = true; return !marked; });
    expect(marked).toBe(false);
    const withNotes = await tracking('footnote.docx');
    const note = (await withNotes.stories.footnotes()).find((n) => n.id > 0)!;
    const s = runCommand(withNotes, 'insertText', { text: 'x' }, { from: 3 }, note);
    expect(s.untracked).toBe('story');
  });
});

describe('tracking off, the mode, and a tree run (ED-005 section 5.4)', () => {
  test('with tracking off, text typed inside a deletion is not deleted', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    expect(session.trackingMode).toBe('Off');
    const d = at(session.doc, 'deletion');
    input(session, (s) => s.tr.insertText('X', d + 2), { from: d + 2 });
    const x = paragraphWith(session.doc, 'A deX');
    expect(pieces(x.node).find((p) => p.startsWith('X'))).toBe('X');
  });

  test('document.setChangeTrackingMode writes w:trackRevisions, and the mode is read back at open', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    expect((await runCommandAsync(session, 'document.setChangeTrackingMode', { value: 'toggle' })).ok).toBe(true);
    expect(session.trackingMode).toBe('TrackAll');
    runCommand(session, 'insertText', { text: 'x' }, { from: at(session.doc, 'text.') });
    const saved = await session.save();
    const settings = await (await WordprocessingMLPackage.load(saved.bytes)).getMainDocumentPart().documentSettingsPart!.getXml();
    expect(settings).toContain('<w:trackRevisions/>');
    expect((await EditorDocument.open(saved.bytes)).trackingMode).toBe('TrackAll');
    await runCommandAsync(session, 'document.setChangeTrackingMode', { value: 'Off' });
    expect(session.trackingMode).toBe('Off');
  });

  test("a tree run with tracking on is tracked by the engine as the session's author", async () => {
    const session = await tracking();
    const r = await session.runOnTree(({ body }) => { body.insertParagraph('From a script.', 'End'); });
    expect(r.error).toBeUndefined();
    const xml = await savedMain(session);
    expect(xml).toMatch(dated(/<w:ins w:id="\d+" @C><w:r><w:t>From a script\.<\/w:t><\/w:r><\/w:ins>/));
  });

  /** The saved paragraph that holds `text`, as XML. */
  const savedParagraph = (xml: string, text: string): string => {
    const at = xml.indexOf(text);
    return xml.slice(xml.lastIndexOf('<w:p>', at) >= xml.lastIndexOf('<w:p ', at) ? xml.lastIndexOf('<w:p>', at) : xml.lastIndexOf('<w:p ', at), xml.indexOf('</w:p>', at) + 6);
  };

  /** An optional run property block, as Word may write one. */
  const RPR = '(?:<w:rPr>(?:(?!<\\/w:r>).)*?<\\/w:rPr>)?';

  test("with tracking off, Enter inside a paragraph whose mark is inserted: the new mark is plain and the old one keeps the insertion, at the end too, where the new empty paragraph keeps it (Jason's findings, ED-005 sections 12.16 and 12.17)", async () => {
    const session = await tracking('revisions/revisions-word15.docx');
    session.trackingMode = 'Off';
    const mark = (index: number): unknown => (session.doc.child(index).attrs as ParagraphAttrs).markRevision;
    const p = paragraphWith(session.doc, 'Split this paragraph');
    const inserted = mark(p.index);
    expect(inserted).toMatchObject({ insertion: { author: 'Author A' } });
    // In the middle: Word's mark at the caret ends the first half, as the tracker marks it with tracking on.
    runCommand(session, 'insertParagraph', undefined, { from: at(session.doc, 'Split', 'Split this paragraph') + 5 });
    expect(session.doc.child(p.index).textContent).toBe('Split');
    expect(mark(p.index)).toBeNull();
    expect(mark(p.index + 1)).toEqual(inserted);
    // At the end: the new mark is at the caret too, before the old one, which the new empty paragraph keeps (core-ts check 29c).
    const q = paragraphWith(session.doc, ' this paragraph after');
    runCommand(session, 'insertParagraph', undefined, { from: q.pos + q.node.nodeSize - 1 });
    expect(mark(q.index)).toBeNull();
    expect(session.doc.child(q.index + 1).content.size).toBe(0);
    expect(mark(q.index + 1)).toEqual(inserted);
    // Saved: the paragraphs by their text, which Word's runs split.
    const saved = [...(await savedMain(session)).matchAll(/<w:p(?: [^>]*)?\/>|<w:p(?: [^>]*)?>(?:(?!<\/w:p>).)*<\/w:p>/g)].map((m) => m[0]);
    const textOf = (xml: string): string => [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('');
    const index = saved.findIndex((x) => textOf(x) === 'Split');
    expect(index).toBeGreaterThan(0);
    const markOf = (xml: string): string => /<w:pPr>(?:(?!<\/w:pPr>).)*<\/w:pPr>/.exec(xml)?.[0] ?? '';
    expect(markOf(saved[index]!)).not.toMatch(/<w:ins /);
    expect(textOf(saved[index + 1]!)).toBe(' this paragraph after the colon:');
    expect(markOf(saved[index + 1]!)).not.toMatch(/<w:ins /);
    expect(textOf(saved[index + 2]!)).toBe('');
    expect(markOf(saved[index + 2]!)).toMatch(/<w:rPr><w:ins w:id="\d+"(?: w16du:dateUtc="[^"]*")? w:author="Author A"/);
  });

  test("with tracking off, a paragraph's properties and an atom's formatting set keep the revisions they hold (ED-005 section 12.16)", async () => {
    const session = await tracking('revisions/revisions-word15.docx');
    session.trackingMode = 'Off';
    // A deleted mark and an inserted one, their paragraphs centred.
    for (const text of ['Join this paragraph', 'Split this paragraph']) {
      const p = paragraphWith(session.doc, text);
      const before = (p.node.attrs as ParagraphAttrs).markRevision;
      expect(before).not.toBeNull();
      expect(runCommand(session, 'paragraph.alignment', { value: 'Centered' }, { from: p.pos + 2 }).ok).toBe(true);
      const after = paragraphWith(session.doc, text).node.attrs as ParagraphAttrs;
      expect(after.jc).toBe('center');
      expect(after.markRevision).toEqual(before);
    }
    // Section 10's deleted tab, inserted break and inserted note reference, made bold.
    const atoms = (doc: PMNode): [string, unknown][] => { const out: [string, unknown][] = []; doc.descendants((n) => { if (n.isInline && !n.isText && n.attrs['revision']) out.push([n.type.name, n.attrs['revision']]); return true; }); return out; };
    const before = atoms(session.doc);
    expect(before.map(([name]) => name)).toEqual(['tab', 'hardBreak', 'footnoteRef']);
    let para: { pos: number; node: PMNode } | undefined;
    session.doc.forEach((node, pos) => { if (!para && atoms(node).length) para = { pos, node }; });
    expect(runCommand(session, 'font.bold', { value: true }, { from: para!.pos + 1, to: para!.pos + para!.node.nodeSize - 1 }).ok).toBe(true);
    expect(atoms(session.doc)).toEqual(before);
  });

  test("with tracking off, text typed inside or at the end of a revision is plain and splits it, as Word writes it (core-ts check 21 re-run, a to d)", async () => {
    const typed = async (author: string, text: string, where: (doc: PMNode) => number): Promise<string> => {
      const session = await tracking('revisions/revisions-word15.docx');
      session.trackingMode = 'Off';
      session.trackingAuthor = author;
      runCommand(session, 'insertText', { text }, { from: where(session.doc) });
      return savedParagraph(await savedMain(session), 'The quick brown');
    };
    const ins = (author: string, text: string): string => `<w:ins w:id="\\d+"(?: w16du:dateUtc="[^"]*")? w:author="${author}"[^>]*><w:r>${RPR}<w:t(?: xml:space="preserve")?>${text}<\\/w:t><\\/w:r><\\/w:ins>`;
    const del = (author: string, text: string): string => `<w:del w:id="\\d+"(?: w16du:dateUtc="[^"]*")? w:author="${author}"[^>]*><w:r>${RPR}<w:delText(?: xml:space="preserve")?>${text}<\\/w:delText><\\/w:r><\\/w:del>`;
    const plain = (text: string): string => `<w:r>${RPR}<w:t(?: xml:space="preserve")?>${text}<\\/w:t><\\/w:r>`;
    // 21a and 21d: inside A's inserted "very", by B or by A: A's w:ins split around a plain run.
    for (const author of ['Author B', 'Author A']) {
      expect(await typed(author, 'X', (doc) => at(doc, 'very', 'The quick brown') + 2)).toMatch(new RegExp(ins('Author A', 've') + plain('X') + ins('Author A', 'ry ')));
    }
    // 21b: at the end of A's insertion, a plain run after it.
    expect(await typed('Author B', 'Y', (doc) => at(doc, 'lazy', 'The quick brown'))).toMatch(new RegExp(ins('Author A', 'very ') + plain('Y')));
    // 21c: inside A's deleted "old", A's w:del split around a plain run.
    expect(await typed('Author B', 'Z', (doc) => at(doc, 'old', 'The quick brown') + 2)).toMatch(new RegExp(del('Author A', 'ol') + plain('Z') + del('Author A', 'd ')));
  });

  test("typing inside another author's revision splits it around the typist's own w:ins, as Word writes it with tracking on (core-ts check 21, a to c)", async () => {
    const typed = async (text: string, where: (doc: PMNode) => number): Promise<string> => {
      const session = await tracking('revisions/revisions-word15.docx');
      session.trackingAuthor = 'Author B';
      runCommand(session, 'insertText', { text }, { from: where(session.doc) });
      return savedParagraph(await savedMain(session), 'The quick brown');
    };
    const ins = (author: string, text: string): string => `<w:ins w:id="\\d+"(?: w16du:dateUtc="[^"]*")? w:author="${author}"[^>]*><w:r>${RPR}<w:t(?: xml:space="preserve")?>${text}<\\/w:t><\\/w:r><\\/w:ins>`;
    const del = (author: string, text: string): string => `<w:del w:id="\\d+"(?: w16du:dateUtc="[^"]*")? w:author="${author}"[^>]*><w:r>${RPR}<w:delText(?: xml:space="preserve")?>${text}<\\/w:delText><\\/w:r><\\/w:del>`;
    // 21a: inside A's inserted "very", A's w:ins split around B's.
    expect(await typed('X', (doc) => at(doc, 'very', 'The quick brown') + 2)).toMatch(new RegExp(ins('Author A', 've') + ins('Author B', 'X') + ins('Author A', 'ry ')));
    // 21b: at the end of A's insertion, B's right after it.
    expect(await typed('Y', (doc) => at(doc, 'lazy', 'The quick brown'))).toMatch(new RegExp(ins('Author A', 'very ') + ins('Author B', 'Y')));
    // 21c: inside A's deleted "old", A's w:del split around B's w:ins.
    expect(await typed('Z', (doc) => at(doc, 'old', 'The quick brown') + 1)).toMatch(new RegExp(del('Author A', 'o') + ins('Author B', 'Z') + del('Author A', 'ld ')));
  });

  test("typing inside moved-to text, as the move's author: the w:moveTo splits around a plain w:ins, as Word writes it (core-ts check 20, case 01)", async () => {
    const session = await tracking('revisions/revisions-word15.docx');
    session.trackingAuthor = 'Author A';
    runCommand(session, 'insertText', { text: 'quickly ' }, { from: at(session.doc, 'before the first', 'The second paragraph of the move') });
    const p = savedParagraph(await savedMain(session), 'quickly ');
    expect(p).toMatch(/<w:moveTo w:id="(\d+)"[^>]*>.*it goes <\/w:t><\/w:r><\/w:moveTo><w:ins w:id="\d+"(?: w16du:dateUtc="[^"]*")? w:author="Author A"[^>]*><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:t xml:space="preserve">quickly <\/w:t><\/w:r><\/w:ins><w:moveTo w:id="(\d+)"[^>]*><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:t>before the first\.<\/w:t>/);
    const ids = [...p.matchAll(/<w:moveTo w:id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
  });

  test("deleting inside moved-to text, as the move's author: a w:del nested in the w:moveTo, not split and not removed (core-ts check 20, case 06)", async () => {
    const session = await tracking('revisions/revisions-word15.docx');
    session.trackingAuthor = 'Author A';
    const x = at(session.doc, 'it goes ', 'The second paragraph of the move');
    input(session, (s) => s.tr.delete(x, x + 8), { from: x + 8 });
    expect(session.doc.nodeAt(x)!.marks.map((m) => m.type.name).sort()).toEqual(expect.arrayContaining(['deletion', 'insertion']));
    const p = savedParagraph(await savedMain(session), 'it goes ');
    // One run holder (the paragraph's mark is a self-closing w:moveTo of its own).
    expect(p.match(/<w:moveTo [^>]*[^/]>/g)).toHaveLength(1);
    expect(p).toMatch(/<w:moveTo [^>]*>(?:(?!<\/w:moveTo>).)*<w:del w:id="\d+"(?: w16du:dateUtc="[^"]*")? w:author="Author A"[^>]*><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:delText xml:space="preserve">it goes <\/w:delText><\/w:r><\/w:del>(?:(?!<\/w:moveTo>).)*<\/w:moveTo>/);
  });

  test('typing inside moved text makes a plain insertion, never a move of its own (section 12.4 item 4)', async () => {
    const session = await tracking('revisions/revisions-word15.docx');
    const x = at(session.doc, 'goes before', 'The second paragraph of the move');
    runCommand(session, 'insertText', { text: 'really ' }, { from: x });
    const node = session.doc.nodeAt(x)!;
    expect(node.marks.find((m) => m.type.name === 'insertion')?.attrs).toEqual(C);
  });
});

describe("a tree run tracked as the agent, whatever the document's mode (ED-005 section 8.5)", () => {
  const settingsOf = async (bytes: Uint8Array): Promise<string | undefined> => {
    try { return new TextDecoder().decode(new ZipPartStore(bytes).loadSync('word/settings.xml')); } catch { return undefined; }
  };
  for (const name of ['tracked-changes.docx', 'comments-two.docx']) {
    test(`${name}: the run's insertion is the agent's tracked change; the mode, the settings part and the room are as they were`, async () => {
      const base = await fixture(name);
      const r = await room(base);
      const session = r.creator;
      expect(session.trackingMode).toBe('Off');
      const settingsBefore = await settingsOf(base);
      expect(session.store.parts.has('part:/word/settings.xml')).toBe(false);
      const run = await session.runOnTree(({ body }) => { body.insertParagraph('Written by the agent.', 'End'); }, 'agent', { author: 'Claude', mode: 'TrackAll', date: new Date('2026-09-29T10:00:00Z') });
      expect(run.error).toBeUndefined();
      expect(session.trackingMode).toBe('Off');
      expect(session.store.parts.has('part:/word/settings.xml')).toBe(false);
      const saved = await session.save();
      expect(saved.selfCheck?.equal).toBe(true);
      const main = new TextDecoder().decode(new ZipPartStore(saved.bytes).loadSync('word/document.xml'));
      expect(main).toMatch(/<w:ins w:id="\d+"(?: w16du:dateUtc="[^"]*")? w:author="Claude"[^>]*><w:r><w:t>Written by the agent\.<\/w:t><\/w:r><\/w:ins>/);
      const settingsAfter = await settingsOf(saved.bytes);
      expect(settingsAfter).toBe(settingsBefore);
      expect(settingsAfter ?? '').not.toMatch(/trackRevisions/);
      // The other peer sees the agent's change as the agent's.
      let found = false;
      r.joiner.doc.descendants((n) => { if (n.isText && n.text!.includes('Written by the agent') && n.marks.some((m) => m.type.name === 'insertion' && m.attrs['author'] === 'Claude')) found = true; return !found; });
      expect(found).toBe(true);
    });
  }
});
