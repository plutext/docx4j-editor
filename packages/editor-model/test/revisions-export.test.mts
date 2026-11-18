// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 5.10, step C2: the export of revisions the edit layer holds and the file has no
// element for (sections 5.3 and 5.7): what the tracker of step C3, a peer or a tree run makes.
// Each case gives content the marks or attributes that tracker would (section 5.4), saves with the
// self-check, and reads the part: a `w:ins` or `w:del` made for the content, an insertion around
// a deletion, another author's text out of an insertion, a formatting change's `w:rPrChange`, an
// atom's run, a new paragraph's mark, a deleted field code as `w:delInstrText`; every id written
// afresh above every annotation id in every part, unread parts included, and never one the
// engine's tracker gives out in a tree run; the floor seeded at load from the editor's own tree.
import { describe, test, expect, vi } from 'vitest';
import { WordprocessingMLPackage, ZipPartStore } from '@docx4j/core-ts';
import type { Node as PMNode, Mark } from 'prosemirror-model';
import { EditorDocument, runCommand, stateFor, formatChangeBeforeRPr, type Revision, type RevisionInfo } from '../src/index.mjs';
import { fixture, dated } from './helpers.mjs';

const C: RevisionInfo = { author: 'Author C', date: '2026-09-28T10:00:00Z', move: null };

/** The first top-level paragraph whose text contains `text`, with its position. */
function paragraphWith(doc: PMNode, text: string): { node: PMNode; pos: number } {
  let found: { node: PMNode; pos: number } | undefined;
  doc.forEach((node, pos) => { if (!found && node.type.name === 'paragraph' && node.textContent.includes(text)) found = { node, pos }; });
  if (!found) throw new Error(`no paragraph with "${text}"`);
  return found;
}

/** The document positions of `text` inside the paragraph that holds it (or the first holding `within`). */
function rangeOf(doc: PMNode, text: string, within = text): { from: number; to: number } {
  const p = paragraphWith(doc, within);
  let from = -1;
  p.node.descendants((n, rel) => {
    if (from < 0 && n.isText && n.text!.includes(text)) from = p.pos + 1 + rel + n.text!.indexOf(text);
    return from < 0;
  });
  if (from < 0) throw new Error(`"${text}" is not in one text node`);
  return { from, to: from + text.length };
}

/** Put a mark on a range, as a transaction the session applies. */
function mark(session: EditorDocument, range: { from: number; to: number }, m: Mark): void {
  session.replace(stateFor(session).tr.addMark(range.from, range.to, m).doc);
}

/** Save with the self-check and return the main part's XML, the rsids left out. */
async function savedMain(session: EditorDocument): Promise<{ xml: string; bytes: Uint8Array }> {
  const saved = await session.save();
  expect(saved.selfCheck?.equal).toBe(true);
  const xml = await (await WordprocessingMLPackage.load(saved.bytes)).getMainDocumentPart().getXml();
  return { xml: xml.replace(/ w:rsid\w*="[^"]*"/g, ''), bytes: saved.bytes };
}

/**
 * Every annotation `w:id` in every WordprocessingML part of a package, once per annotation: a
 * range's end shares its start's id and is left out, as are the ids of other spaces (a note, a
 * comment and their references, a list definition).
 */
function annotationIds(bytes: Uint8Array): number[] {
  const store = new ZipPartStore(bytes);
  const out: number[] = [];
  const other = /^(footnote|endnote|comment|footnoteReference|endnoteReference|commentReference|abstractNum|num)$|End$/;
  for (const name of store.partNames()) {
    if (!/^word\/.*\.xml$/.test(name) || name.endsWith('comments.xml')) continue;
    const xml = new TextDecoder().decode(store.loadSync(name));
    for (const m of xml.matchAll(/<w:([A-Za-z]+)\b[^>]*\bw:id="(-?\d+)"/g)) if (!other.test(m[1]!)) out.push(Number(m[2]));
  }
  return out;
}

const idsOf = (xml: string, element: string): number[] => [...xml.matchAll(new RegExp(`<w:${element} w:id="(\\d+)"`, 'g'))].map((m) => Number(m[1]));

describe('revisions the file has no element for, exported (ED-005 section 5.3)', () => {
  test('text marked inserted goes out in a w:ins of its own, with an id above every annotation id in the package', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    mark(session, rangeOf(session.doc, 'text.'), session.schema.marks['insertion']!.create(C));
    const { xml, bytes } = await savedMain(session);
    const m = dated(/<w:ins w:id="(\d+)" @C><w:r><w:t>text\.<\/w:t><\/w:r><\/w:ins>/).exec(xml);
    expect(m, xml.slice(xml.indexOf('Some'), xml.indexOf('Some') + 400)).not.toBeNull();
    expect(Number(m![1])).toBeGreaterThan(Math.max(...annotationIds(original)));
    const ids = annotationIds(bytes);
    expect(new Set(ids).size).toBe(ids.length);
    // The reload carries the mark.
    const reloaded = await EditorDocument.open(bytes);
    const r = rangeOf(reloaded.doc, 'text.');
    expect(reloaded.doc.nodeAt(r.from)!.marks.find((x) => x.type.name === 'insertion')?.attrs).toEqual(C);
  });

  test('text marked deleted goes out in a w:del, its text w:delText, and the run split around it', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    mark(session, rangeOf(session.doc, 'text'), session.schema.marks['deletion']!.create(C));
    const { xml } = await savedMain(session);
    expect(xml).toMatch(dated(/<w:r><w:t xml:space="preserve">Some <\/w:t><\/w:r><w:del w:id="\d+" @C><w:r><w:delText>text<\/w:delText><\/w:r><\/w:del><w:r><w:t>\.<\/w:t><\/w:r>/));
  });

  test("another author's insertion deleted: the w:ins keeps its id and holds a w:del", async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    mark(session, rangeOf(session.doc, 'sertion'), session.schema.marks['deletion']!.create(C));
    const { xml } = await savedMain(session);
    expect(xml).toMatch(dated(/<w:ins w:id="1" w:author="Jason Harrop" w:date="2007-12-09T10:14:00Z"><w:r><w:t>An in<\/w:t><\/w:r><w:del w:id="\d+" @C><w:r><w:delText>sertion<\/w:delText><\/w:r><\/w:del><\/w:ins>/));
  });

  test("text another author types into an insertion leaves that author's w:ins, which goes on after it with an id of its own", async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const at = rangeOf(session.doc, 'insertion').from;
    const state = stateFor(session, { from: at });
    // What the tracker makes of typing there: the run's marks, with Author C's insertion for Jason Harrop's.
    const marks = state.doc.resolve(at).marks().filter((m) => m.type.name !== 'insertion').concat(session.schema.marks['insertion']!.create(C));
    session.replace(state.tr.insert(at, session.schema.text('new ', marks)).doc);
    const { xml } = await savedMain(session);
    expect(xml).toMatch(/<w:ins w:id="1" w:author="Jason Harrop"[^>]*><w:r><w:t xml:space="preserve">An <\/w:t><\/w:r><\/w:ins><w:ins w:id="\d+"(?: w16du:dateUtc="[^"]*")? w:author="Author C"[^>]*><w:r><w:t xml:space="preserve">new <\/w:t><\/w:r><\/w:ins><w:ins w:id="\d+"(?: w16du:dateUtc="[^"]*")? w:author="Jason Harrop"[^>]*><w:r><w:t>insertion<\/w:t><\/w:r><\/w:ins>/);
    const ins = idsOf(xml, 'ins');
    expect(new Set(ins).size).toBe(ins.length);
  });

  test('a paragraph split inside an insertion: each half has the w:ins, the second with an id of its own', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    runCommand(session, 'insertParagraph', undefined, { from: rangeOf(session.doc, 'sertion').from });
    const { xml } = await savedMain(session);
    const ins = idsOf(xml, 'ins');
    expect(ins.length).toBeGreaterThanOrEqual(2);
    expect(ins[0]).toBe(1);
    expect(new Set(ins).size).toBe(ins.length);
  });

  test("a formatting change the file has none for writes w:rPrChange with the old properties; the reload's mark is the same", async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const r = rangeOf(session.doc, 'text');
    // What the tracker makes of Bold: the mark, and the run's properties before it.
    mark(session, r, session.schema.marks['bold']!.create());
    const change = session.schema.marks['formatChange']!.create({ author: C.author, date: C.date, before: formatChangeBeforeRPr(undefined) });
    mark(session, r, change);
    const { xml, bytes } = await savedMain(session);
    expect(xml).toMatch(dated(/<w:r><w:rPr><w:b\/><w:bCs\/><w:rPrChange w:id="\d+" @C><w:rPr\/><\/w:rPrChange><\/w:rPr><w:t>text<\/w:t><\/w:r>/));
    const reloaded = await EditorDocument.open(bytes);
    expect(reloaded.doc.nodeAt(rangeOf(reloaded.doc, 'text').from)!.marks.find((x) => x.type.name === 'formatChange')?.attrs).toEqual(change.attrs);
  });

  test("a formatting change taken off (an accept) removes the run's w:rPrChange; the ones left keep their ids", async () => {
    const session = await EditorDocument.open(await fixture('revisions/revisions-word2010.docx'));
    const before = await savedMain(session);
    const ids = idsOf(before.xml, 'rPrChange');
    expect(ids).toHaveLength(3);
    const r = rangeOf(session.doc, 'alpha', 'The word alpha becomes');
    session.replace(stateFor(session).tr.removeMark(r.from, r.to, session.schema.marks['formatChange']!).doc);
    const { xml } = await savedMain(session);
    expect(idsOf(xml, 'rPrChange')).toEqual(ids.slice(1));
  });

  test('an atom the file has no run for, inserted, goes out in a w:ins', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const at = rangeOf(session.doc, 'text').from;
    const revision: Revision = { insertion: C };
    session.replace(stateFor(session).tr.insert(at, session.schema.nodes['tab']!.create({ revision })).doc);
    const { xml } = await savedMain(session);
    expect(xml).toMatch(dated(/<w:t xml:space="preserve">Some <\/w:t><\/w:r><w:ins w:id="\d+" @C><w:r><w:tab\/><\/w:r><\/w:ins><w:r><w:t>text\.<\/w:t><\/w:r>/));
  });

  test("a new paragraph whose mark and text are inserted: w:pPr/w:rPr/w:ins and a w:ins around its run, each id its own", async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const p = paragraphWith(session.doc, 'Some text.');
    const schema = session.schema;
    const node = schema.nodes['paragraph']!.create({ markRevision: { insertion: C } }, schema.text('Typed.', [schema.marks['insertion']!.create(C)]));
    session.replace(stateFor(session).tr.insert(p.pos + p.node.nodeSize, node).doc);
    const { xml } = await savedMain(session);
    const m = dated(/<w:p><w:pPr><w:rPr><w:ins w:id="(\d+)" @C\/><\/w:rPr><\/w:pPr><w:ins w:id="(\d+)" @C><w:r><w:t>Typed\.<\/w:t><\/w:r><\/w:ins><\/w:p>/).exec(xml);
    expect(m, xml.slice(xml.indexOf('Typed') - 300, xml.indexOf('Typed') + 100)).not.toBeNull();
    expect(m![1]).not.toBe(m![2]);
  });

  test('a deleted field: its runs in a w:del, its code w:delInstrText; the deletion taken off, w:instrText again', async () => {
    const session = await EditorDocument.open(await fixture('mergefield.docx'));
    let target: { node: PMNode; pos: number } | undefined;
    session.doc.forEach((node, pos) => { if (!target && node.type.name === 'paragraph' && node.childCount > 0 && [...Array(node.childCount).keys()].some((i) => node.child(i).attrs['kind'] === 'instrText')) target = { node, pos }; });
    expect(target).toBeDefined();
    const withRevision = (revision: Revision | null): void => {
      const tr = stateFor(session).tr;
      target!.node.forEach((child, offset) => { tr.setNodeMarkup(target!.pos + 1 + offset, undefined, { ...child.attrs, revision }); });
      session.replace(tr.doc);
      target = { node: session.doc.nodeAt(target!.pos)!, pos: target!.pos };
    };
    withRevision({ deletion: C });
    let { xml } = await savedMain(session);
    const field = /<w:del w:id="\d+"(?: w16du:dateUtc="[^"]*")? w:author="Author C"[^>]*>(.*?)<\/w:del>/.exec(xml)?.[1] ?? '';
    expect(field).toContain('<w:delInstrText xml:space="preserve"> MERGEFIELD  Kundenstrasse \\* MERGEFORMAT </w:delInstrText>');
    expect(field.match(/<w:fldChar /g)).toHaveLength(3);
    expect(xml).not.toContain('<w:instrText');
    withRevision(null);
    ({ xml } = await savedMain(session));
    expect(xml).toContain('<w:instrText xml:space="preserve"> MERGEFIELD  Kundenstrasse \\* MERGEFORMAT </w:instrText>');
    expect(xml).not.toContain('w:author="Author C"');
  });
});

describe('ids at export (ED-005 section 5.7)', () => {
  test("at load the floor covers the main part from the editor's own tree: the engine does not read that part again", async () => {
    const bytes = await fixture('revisions/revisions-word15.docx');
    const session = await EditorDocument.open(bytes);
    const highest = Math.max(...annotationIds(bytes));
    expect(session.pkg.annotationIdFloor).toBeGreaterThanOrEqual(highest);
    const read = vi.spyOn(session.main, 'readContents');
    await session.pkg.seedAnnotationIds();
    expect(read).not.toHaveBeenCalled();
  });

  test("above an annotation id in a part nobody has read: a bookmark with w:id 500 in the footnotes", async () => {
    const pkg = await WordprocessingMLPackage.load(await fixture('footnote.docx'));
    const part = pkg.getMainDocumentPart().footnotesPart!;
    const xml = await part.getXml();
    const patched = xml.replace('<w:t xml:space="preserve"> THIRD</w:t></w:r>', '<w:t xml:space="preserve"> THIRD</w:t></w:r><w:bookmarkStart w:id="500" w:name="far"/><w:bookmarkEnd w:id="500"/>');
    expect(patched).not.toBe(xml);
    part.setXml(patched);
    const session = await EditorDocument.open(await pkg.save());
    const p = paragraphWith(session.doc, 'footnotes');
    mark(session, { from: p.pos + 1, to: p.pos + 5 }, session.schema.marks['insertion']!.create(C));
    const { xml: main } = await savedMain(session);
    const ins = idsOf(main, 'ins');
    expect(ins).toHaveLength(1);
    expect(ins[0]).toBeGreaterThan(500);
  });

  test("one counter with the engine's tracker: a tree run's tracked edit after a save, then a save, and no id twice", async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    mark(session, rangeOf(session.doc, 'text.'), session.schema.marks['insertion']!.create(C));
    const first = await savedMain(session);
    const made = idsOf(first.xml, 'ins').filter((id) => id !== 1);
    expect(made).toHaveLength(1);
    const run = await session.runOnTree(async ({ pkg, body }) => {
      await pkg.withTracking({ author: { name: 'Engine' } }, () => { body.insertParagraph('From a script.', 'End'); });
    });
    expect(run.error).toBeUndefined();
    const { xml, bytes } = await savedMain(session);
    expect(xml).toContain('w:author="Engine"');
    const engineIds = [...xml.matchAll(/w:id="(\d+)"(?: w16du:dateUtc="[^"]*")? w:author="Engine"/g)].map((m) => Number(m[1]));
    expect(engineIds.length).toBeGreaterThan(0);
    for (const id of engineIds) expect(id).toBeGreaterThan(made[0]!);
    const ids = annotationIds(bytes);
    expect(new Set(ids).size).toBe(ids.length);
  });
});
