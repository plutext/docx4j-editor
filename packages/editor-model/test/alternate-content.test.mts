// SPDX-License-Identifier: GPL-3.0-or-later
//
// A kept `mc:AlternateContent` (core-ts CR-010 phase 4; ED-005 section 12.106; ED-009): its selected
// branch's items are projected as the run's own (`selectedBranchOf`), text, tabs, breaks, images and
// atoms, each naming the element as `alt`; a paragraph-level element holds its branch's runs as a
// hyperlink holds its runs. An unedited paragraph writes the element back by identity, every branch
// with it; an edited one writes the element with the selected branch rebuilt from the editor's content
// and every other branch the original, unchanged (ED-009 section 3).
import { describe, expect, test } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { WordprocessingMLPackage, type Element } from '@docx4j/core-ts';
import { EditorDocument, Finder, runCommand, runCommandAsync, blipEmbedOf, inlineDrawingOf } from '../src/index.mjs';
import { fixture, mainXml, comparePackages } from './helpers.mjs';

/** The inline nodes of a type, with their positions. */
function nodesOf(doc: PMNode, type: string): { node: PMNode; pos: number }[] {
  const out: { node: PMNode; pos: number }[] = [];
  doc.descendants((node, pos) => { if (node.type.name === type) out.push({ node, pos }); return true; });
  return out;
}

/** The `alt` of an inline node: the atom's attribute, or the text's `runSource` mark's. */
function altOf(node: PMNode): string | null {
  if (node.isText) return (node.marks.find((m) => m.type.name === 'runSource')?.attrs['alt'] as string | null | undefined) ?? null;
  return (node.attrs['alt'] as string | null | undefined) ?? null;
}

/** Where `text` starts in the document, and the range over it. */
function rangeOf(doc: PMNode, text: string): { from: number; to: number } {
  let at = -1;
  doc.descendants((n, pos) => { if (at < 0 && n.isText && n.text!.includes(text)) at = pos + n.text!.indexOf(text); return at < 0; });
  if (at < 0) throw new Error(`no "${text}"`);
  return { from: at, to: at + text.length };
}

/** Types before the node at `pos`, so its paragraph is rebuilt at the save, and gives the saved main part. */
async function savedAfterTypingBefore(session: EditorDocument, pos: number): Promise<string> {
  expect(runCommand(session, 'insertText', { text: 'typed ' }, { from: pos }).ok).toBe(true);
  return saved(session);
}

/** Saves, holding the self-check, and gives the main part. */
async function saved(session: EditorDocument): Promise<string> {
  const result = await session.save();
  expect(result.selfCheck?.equal, JSON.stringify(result.selfCheck?.difference)).toBe(true);
  return mainXml(result.bytes);
}

const count = (xml: string, s: string): number => xml.split(s).length - 1;
/** The body's paragraphs as XML, in order (the fixture nests none). */
const paragraphs = (xml: string): string[] => xml.replace(/.*<w:body>/s, '').replace(/<w:sectPr.*/s, '').match(/<w:p(?: [^>]*)?>.*?<\/w:p>|<w:p\/>/gs) ?? [];
/** The `mc:AlternateContent` elements of a piece of XML. */
const elements = (xml: string): string[] => xml.match(/<mc:AlternateContent>.*?<\/mc:AlternateContent>/gs) ?? [];

describe('a kept mc:AlternateContent holding a text box (ED-005 section 12.106)', () => {
  test('loadAndSave.docx: an anchored text box is an opaque atom of the drawing\'s kind naming the element, and a rebuilt paragraph writes it back with both branches', async () => {
    const session = await EditorDocument.open(await fixture('loadAndSave.docx'));
    const atoms = nodesOf(session.doc, 'opaqueInline').filter((a) => altOf(a.node) !== null);
    expect(atoms.length).toBe(1);
    expect(atoms[0]!.node.attrs['kind']).toBe('drawing');
    const element = session.ctx.ids.get(altOf(atoms[0]!.node)!) as Element;
    expect(element.name.localPart).toBe('AlternateContent');
    expect((session.ctx.ids.get(atoms[0]!.node.attrs['sourceId'] as string) as Element).name.localPart).toBe('drawing');
    expect(session.diagnostics.opaque.some((e) => e.kind === 'AlternateContent')).toBe(false);
    expect(session.diagnostics.opaque.filter((e) => e.kind === 'drawing').length).toBe(1);
    const before = await mainXml(await fixture('loadAndSave.docx'));
    const after = await savedAfterTypingBefore(session, atoms[0]!.pos);
    expect(after).toContain('typed ');
    for (const s of ['<mc:AlternateContent', '<mc:Choice Requires="wps"', '<mc:Fallback', '<wps:wsp', '<v:shape', 'o:gfxdata=']) expect(count(after, s), s).toBe(count(before, s));
  });

  test('numbering-stories.docx: an inline text box is an image node, as a plain inline drawing is, its source the drawing and its alt the element', async () => {
    const session = await EditorDocument.open(await fixture('numbering-stories.docx'));
    const images = nodesOf(session.doc, 'image');
    expect(images.length).toBe(1);
    const source = session.ctx.ids.get(images[0]!.node.attrs['sourceId'] as string) as Element;
    expect(source.name.localPart).toBe('drawing');
    expect(inlineDrawingOf(source)?.name.localPart).toBe('drawing');
    expect(blipEmbedOf(source)).toBeUndefined();   // a text box, not a picture
    expect((session.ctx.ids.get(altOf(images[0]!.node)!) as Element).name.localPart).toBe('AlternateContent');
    expect(nodesOf(session.doc, 'opaqueInline').some((a) => a.node.attrs['kind'] === 'AlternateContent')).toBe(false);
    const before = await mainXml(await fixture('numbering-stories.docx'));
    const after = await savedAfterTypingBefore(session, images[0]!.pos);
    for (const s of ['<mc:AlternateContent', '<mc:Choice', '<mc:Fallback', '<wps:wsp', '<v:shape']) expect(count(after, s), s).toBe(count(before, s));
  });

  test('a picture inside a Choice: an image node showing the picture, the element written back whole', async () => {
    // headers.docx's inline picture wrapped as Word wraps content for a newer Word, the Fallback a VML stand-in.
    const pkg = await WordprocessingMLPackage.load(await fixture('headers.docx'));
    const part = pkg.getMainDocumentPart();
    const xml = await part.getXml();
    const drawing = /<w:drawing>.*?<\/w:drawing>/s.exec(xml)![0];
    const embed = /r:embed="([^"]+)"/.exec(drawing)![1]!;
    const wrapped = `<mc:AlternateContent><mc:Choice Requires="wps">${drawing}</mc:Choice><mc:Fallback><w:pict><v:shape id="x" style="width:10pt;height:10pt" o:gfxdata="UEsDBBQABgAIAAAAIQA="><v:imagedata r:id="${embed}"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent>`;
    part.setXml(xml.replace(drawing, wrapped));
    const bytes = new Uint8Array(await pkg.save());
    const session = await EditorDocument.open(bytes);
    const image = nodesOf(session.doc, 'image').find((i) => altOf(i.node) !== null);
    expect(image).toBeDefined();
    const source = session.ctx.ids.get(image!.node.attrs['sourceId'] as string) as Element;
    expect(source.name.localPart).toBe('drawing');
    expect(blipEmbedOf(source)).toBe(embed);
    expect(image!.node.attrs['width']).toBeGreaterThan(0);
    expect(session.diagnostics.opaque.some((e) => e.kind === 'AlternateContent')).toBe(false);
    const after = await savedAfterTypingBefore(session, image!.pos);
    expect(after).toContain('typed ');
    expect(count(after, '<mc:AlternateContent')).toBe(1);
    expect(count(after, '<mc:Choice Requires="wps"')).toBe(1);
    expect(count(after, '<v:imagedata')).toBe(1);
    expect(count(after, `r:embed="${embed}"`)).toBe(1);
  });
});

describe('a kept mc:AlternateContent holding run content (ED-009)', () => {
  // mc/mc-branches.docx, built by fixtures/build/mc-branches.mjs: the case paragraphs are at these indices.
  const CASE = { text: 2, tabAndBreak: 4, x14ac: 6, x14acNoFallback: 8, picture: 10, ten: 12, tenReordered: 14, paragraphLevel: 16, sdtdhFirst: 18, sdtflFirst: 20, fallback: 22, noFallback: 24 };
  const FALLBACK = '<mc:Fallback><w:t xml:space="preserve">The Earth is round</w:t></mc:Fallback>';

  async function open(): Promise<{ session: EditorDocument; bytes: Uint8Array; before: string }> {
    const bytes = await fixture('mc/mc-branches.docx');
    return { session: await EditorDocument.open(bytes), bytes, before: await mainXml(bytes) };
  }
  const items = (p: PMNode): [string, string | null][] => { const out: [string, string | null][] = []; p.forEach((n) => out.push([n.isText ? n.text! : n.type.name, altOf(n)])); return out; };

  test('the projection: the selected branch\'s items are the run\'s, naming the element; a paragraph-level element holds its runs', async () => {
    const { session } = await open();
    const doc = session.doc;
    // The one element with nothing selected (no Choice the engine understands, no Fallback: core-ts CR-029) is an opaque atom.
    expect(session.diagnostics.opaque).toEqual([expect.objectContaining({ kind: 'AlternateContent', inline: true, address: `body/${CASE.noFallback}/0/0` })]);
    const alt = (i: number): string => altOf(doc.child(i).firstChild!)!;
    expect(items(doc.child(CASE.text))).toEqual([['The Earth is flat', alt(CASE.text)]]);
    expect((session.ctx.ids.get(alt(CASE.text)) as Element).name.localPart).toBe('AlternateContent');
    expect(items(doc.child(CASE.tabAndBreak)).map(([n]) => n)).toEqual(['Before', 'tab', 'after the tab', 'hardBreak', 'after the break']);
    expect(new Set(items(doc.child(CASE.tabAndBreak)).map(([, a]) => a)).size).toBe(1);
    // An unknown (private) Choice: the Fallback; with no Fallback nothing is selected (core-ts CR-029, objects CR-021), and the
    // element is an opaque atom kept whole, its text neither shown nor found, as Word reads it (ED-009 section 6).
    expect(items(doc.child(CASE.fallback)).map(([n]) => n)).toEqual(['private fallback']);
    expect(items(doc.child(CASE.noFallback))).toEqual([['opaqueInline', null]]);
    // Excel's x14ac: understood by Requires alone since core-ts CR-010 section 15, as Word 365 understands it (ED-009 section 9.2).
    expect(items(doc.child(CASE.x14ac)).map(([n]) => n)).toEqual(['unknown choice']);
    expect(items(doc.child(CASE.x14acNoFallback)).map(([n]) => n)).toEqual(['nothing to select']);
    const picture = items(doc.child(CASE.picture));
    expect(picture.map(([n]) => n)).toEqual(['A picture: ', 'image', ' after it']);
    expect(new Set(picture.map(([, a]) => a)).size).toBe(1);
    const image = doc.child(CASE.picture).child(1);
    expect(blipEmbedOf(session.ctx.ids.get(image.attrs['sourceId'] as string) as Element)).toBe('rId17');
    // Ten Choices: the first in the paragraph's order, since CR-010 section 15 counts w16du, w16sdtdh and w16sdtfl as
    // understood by Requires alone, as Word 365 does (ED-009 section 9.2 items 3 and 5).
    expect(items(doc.child(CASE.ten)).map(([n]) => n)).toEqual(['I prefer ', 'w14']);
    expect(items(doc.child(CASE.ten)).map(([, a]) => a !== null)).toEqual([false, true]);
    expect(items(doc.child(CASE.tenReordered)).map(([n]) => n)).toEqual(['I prefer ', 'w16du']);
    expect(items(doc.child(CASE.sdtdhFirst)).map(([n]) => n)).toEqual(['I prefer ', 'w16sdtdh']);
    expect(items(doc.child(CASE.sdtflFirst)).map(([n]) => n)).toEqual(['I prefer ', 'w16sdtfl']);
    // The paragraph-level shape: its runs with their own formatting, held by the element (no alt).
    const level = doc.child(CASE.paragraphLevel);
    expect(items(level)).toEqual([['bold ying', null], [' plain ying', null]]);
    expect(level.child(0).marks.some((m) => m.type.name === 'bold')).toBe(true);
  });

  test('untouched, the file round-trips part by part', async () => {
    const { session, bytes } = await open();
    const result = await session.save();
    expect(result.selfCheck?.equal).toBe(true);
    expect(comparePackages(bytes, result.bytes)).toEqual([]);
  });

  test('typing in the branch writes the Choice changed and the Fallback as it was', async () => {
    const { session, before } = await open();
    const { from } = rangeOf(session.doc, 'The Earth is flat');
    expect(runCommand(session, 'insertText', { text: 'XX' }, { from: from + 4 }).ok).toBe(true);
    const after = await saved(session);
    const p = paragraphs(after)[CASE.text]!;
    expect(elements(p).length).toBe(1);
    expect(p).toMatch(/<mc:Choice Requires="wps"><w:t[^>]*>The XXEarth is flat<\/w:t><\/mc:Choice>/);
    expect(p).toContain(FALLBACK);
    // Every other paragraph as it was (the part is marshalled again, which orders a drawing's attributes, and declares a
    // source-only prefix on the mc:Choice that requires it, objects 0.5.0: those paragraphs are compared without them).
    const was = paragraphs(before);
    const plain = (q: string): string => q.replace(/<wp:inline [^>]*>/, '<wp:inline>').replace(/<mc:Choice xmlns:zz9="[^"]*" Requires="zz9">/g, '<mc:Choice Requires="zz9">');
    paragraphs(after).forEach((q, i) => { if (i !== CASE.text) expect(plain(q), `paragraph ${i}`).toBe(plain(was[i]!)); });
  });

  test('the other branches, nine Choices and the Fallback, are written back unchanged and in order', async () => {
    const { session, before } = await open();
    const { from } = rangeOf(session.doc, 'w14');
    expect(runCommand(session, 'insertText', { text: 'XX' }, { from: from + 1 }).ok).toBe(true);
    const after = await saved(session);
    const [was] = elements(paragraphs(before)[CASE.ten]!);
    const [now] = elements(paragraphs(after)[CASE.ten]!);
    const others = (e: string): string => e.replace(/<mc:Choice Requires="w14">.*?<\/mc:Choice>/s, '');
    expect(now).toMatch(/<mc:Choice Requires="w14"><w:t[^>]*>wXX14<\/w:t><\/mc:Choice>/);
    expect(others(now!)).toBe(others(was!));
    expect(count(now!, '<mc:Choice ')).toBe(10);
  });

  test('a formatting split: the first piece keeps the other branches, each later piece carries them empty, in order', async () => {
    const { session } = await open();
    const { from } = rangeOf(session.doc, 'The Earth is flat');
    expect(runCommand(session, 'font.bold', { value: true }, { from: from + 4, to: from + 9 }).ok).toBe(true);
    const ten = rangeOf(session.doc, 'w14');
    expect(runCommand(session, 'font.bold', { value: true }, { from: ten.from + 1, to: ten.from + 2 }).ok).toBe(true);
    const after = await saved(session);
    const pieces = elements(paragraphs(after)[CASE.text]!);
    expect(pieces.length).toBe(3);
    expect(pieces[0]).toMatch(/<mc:Choice Requires="wps"><w:t[^>]*>The <\/w:t><\/mc:Choice>/);
    expect(pieces[0]).toContain(FALLBACK);
    expect(pieces[1]).toBe('<mc:AlternateContent><mc:Choice Requires="wps"><w:t>Earth</w:t></mc:Choice><mc:Fallback/></mc:AlternateContent>');
    expect(pieces[2]).toMatch(/<mc:Choice Requires="wps"><w:t[^>]*> is flat<\/w:t><\/mc:Choice><mc:Fallback\/><\/mc:AlternateContent>$/);
    expect(paragraphs(after)[CASE.text]).toMatch(/<w:r><w:rPr><w:b\/>.*?<\/w:rPr><mc:AlternateContent>/);
    const tenPieces = elements(paragraphs(after)[CASE.ten]!);
    expect(tenPieces.length).toBe(3);
    expect(tenPieces[0]).toContain(FALLBACK);
    for (const later of tenPieces.slice(1)) {
      expect(count(later, '<mc:Choice ')).toBe(10);
      expect(later).toMatch(/<mc:Choice Requires="w15"\/><mc:Choice Requires="w16se"\/>/);
      expect(later).toMatch(/<mc:Fallback\/><\/mc:AlternateContent>$/);
    }
  });

  test('part of the branch deleted keeps the element; all of it deleted removes the element with every branch', async () => {
    const { session } = await open();
    const part = rangeOf(session.doc, 'Earth ');
    expect(runCommand(session, 'insertText', { text: '' }, part).ok).toBe(true);
    const all = rangeOf(session.doc, 'private fallback');
    expect(runCommand(session, 'insertText', { text: '' }, all).ok).toBe(true);
    const after = await saved(session);
    const p = paragraphs(after);
    expect(p[CASE.text]).toMatch(/<mc:Choice Requires="wps"><w:t[^>]*>The is flat<\/w:t><\/mc:Choice>/);
    expect(p[CASE.text]).toContain(FALLBACK);
    expect(p[CASE.fallback]).toBe('<w:p/>');
  });

  test('a paragraph-level element: an edit lands in its Choice, the Fallback as it was', async () => {
    const { session, before } = await open();
    const { from } = rangeOf(session.doc, ' plain ying');
    expect(runCommand(session, 'insertText', { text: 'XX' }, { from: from + 1 }).ok).toBe(true);
    const after = await saved(session);
    const p = paragraphs(after)[CASE.paragraphLevel]!;
    expect(elements(p).length).toBe(1);
    expect(p).toMatch(/^<w:p><mc:AlternateContent><mc:Choice Requires="wps"><w:r><w:rPr><w:b\/>.*?<\/w:rPr><w:t[^>]*>bold ying<\/w:t><\/w:r><w:r><w:t[^>]*> XXplain ying<\/w:t><\/w:r><\/mc:Choice>/);
    const fallback = /<mc:Fallback>.*?<\/mc:Fallback>/s;
    expect(fallback.exec(p)![0]).toBe(fallback.exec(paragraphs(before)[CASE.paragraphLevel]!)![0]);
  });

  test('tracked: a deletion writes w:delText in the Choice inside w:del, the Fallback kept; accepted it removes the piece, rejected it is one element again', async () => {
    const { session } = await open();
    session.trackingAuthor = 'Tracker';
    expect((await runCommandAsync(session, 'document.setChangeTrackingMode', { value: 'TrackAll' })).ok).toBe(true);
    expect(runCommand(session, 'insertText', { text: '' }, rangeOf(session.doc, 'flat')).ok).toBe(true);
    // Typed inside the branch's text (at its start the marks before the element win, ED-009 section 3.3).
    expect(runCommand(session, 'insertText', { text: 'very ' }, { from: rangeOf(session.doc, 'w14').from + 1 }).ok).toBe(true);
    const tracked = await saved(session);
    const p = paragraphs(tracked)[CASE.text]!;
    expect(p).toMatch(/<w:del [^>]*w:author="Tracker"[^>]*><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:delText>flat<\/w:delText><\/mc:Choice><mc:Fallback\/><\/mc:AlternateContent><\/w:r><\/w:del>/);
    expect(elements(p)[0]).toContain(FALLBACK);
    const q = paragraphs(tracked)[CASE.ten]!;
    expect(q).toMatch(/<w:ins [^>]*><w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:t[^>]*>very <\/w:t><\/mc:Choice><mc:Choice Requires="w15"\/>/);
    expect(count(q, '<mc:AlternateContent')).toBe(3);
    expect(elements(q)[0]).toMatch(/<mc:Choice Requires="w14"><w:t[^>]*>w<\/w:t><\/mc:Choice><mc:Choice Requires="w15"><w:t/);   // the branches travel with the first piece in document order
    expect(elements(q)[0]).toContain(FALLBACK);
    // Rejected in the session: the text is one piece again, so one element. (A file saved with the
    // split holds two elements, and reopened stays two: the second's branches are its own, empty.)
    expect((await runCommandAsync(session, 'trackedChanges.rejectAll', {})).ok).toBe(true);
    const r = paragraphs(await saved(session));
    expect(elements(r[CASE.text]!).length).toBe(1);
    expect(r[CASE.text]).toMatch(/<w:t[^>]*>The Earth is flat<\/w:t>/);
    expect(r[CASE.text]).toContain(FALLBACK);
    expect(elements(r[CASE.ten]!).length).toBe(1);
    const accepted = (await open()).session;
    accepted.trackingAuthor = 'Tracker';
    expect((await runCommandAsync(accepted, 'document.setChangeTrackingMode', { value: 'TrackAll' })).ok).toBe(true);
    expect(runCommand(accepted, 'insertText', { text: '' }, rangeOf(accepted.doc, 'flat')).ok).toBe(true);
    expect(runCommand(accepted, 'insertText', { text: 'very ' }, { from: rangeOf(accepted.doc, 'w14').from + 1 }).ok).toBe(true);
    expect((await runCommandAsync(accepted, 'trackedChanges.acceptAll', {})).ok).toBe(true);
    const a = paragraphs(await saved(accepted));
    expect(elements(a[CASE.text]!).length).toBe(1);
    expect(a[CASE.text]).toMatch(/<w:t[^>]*>The Earth is <\/w:t>/);
    expect(a[CASE.text]).toContain(FALLBACK);
    expect(a[CASE.ten]).toMatch(/<w:t[^>]*>wvery 14<\/w:t>/);
    expect(elements(a[CASE.ten]!).length).toBe(1);
    expect(a[CASE.ten]).toContain(FALLBACK);
  });

  test('find and the engine both see the branch\'s text (ED-009 section 6\'s known difference, closed by core-ts 0.6.0: CR-012 section 12 reads the selected branch)', async () => {
    const { session, bytes } = await open();
    const body = await (await WordprocessingMLPackage.load(bytes)).getBody();
    expect(new Finder('flat', {}, session.ctx.ids).all(session.doc).map((h) => h.text)).toEqual(['flat']);
    expect(body.search('flat').length).toBe(1);
    expect(body.paragraphs[CASE.text]!.text).toBe('The Earth is flat');
    expect(session.doc.child(CASE.text).textContent).toBe('The Earth is flat');
  });
});

describe('Word 365\'s save of the editor\'s shapes (ED-009 section 4, step S6; section 9.2)', () => {
  // mc/mc-for-word-word365.docx: Jason's Word 365 save (2026-10-20) of test-results/mc-for-word.docx, built by
  // fixtures/build/mc-for-word.mjs from mc-branches.docx (item 1 "XX" typed into case 1's Choice; item 2 "w1" of
  // case 6 made bold, a split; item 3 "after the tab" of case 2 deleted with tracking on, accepted in Word;
  // item 4 "XX" typed into case 8's second run). Word resolved every element, edited or not, to plain runs.
  const CASE = { text: 2, tabAndBreak: 4, x14ac: 6, x14acNoFallback: 8, picture: 10, ten: 12, tenReordered: 14, paragraphLevel: 16 };
  const texts = (p: PMNode): string[] => { const out: string[] = []; p.forEach((n) => out.push(n.isText ? n.text! : n.type.name)); return out; };

  /** The editor's own edits of the Word file, made again over the fixture, the deletion accepted. */
  async function editedHere(): Promise<EditorDocument> {
    const session = await EditorDocument.open(await fixture('mc/mc-branches.docx'));
    expect(runCommand(session, 'insertText', { text: 'XX' }, { from: rangeOf(session.doc, 'The Earth is flat').from + 4 }).ok).toBe(true);
    const w14 = rangeOf(session.doc, 'w14');
    expect(runCommand(session, 'font.bold', { value: true }, { from: w14.from, to: w14.from + 2 }).ok).toBe(true);
    expect(runCommand(session, 'insertText', { text: 'XX' }, { from: rangeOf(session.doc, ' plain ying').from + 1 }).ok).toBe(true);
    session.trackingAuthor = 'Editor';
    expect((await runCommandAsync(session, 'document.setChangeTrackingMode', { value: 'TrackAll' })).ok).toBe(true);
    expect(runCommand(session, 'insertText', { text: '' }, rangeOf(session.doc, 'after the tab')).ok).toBe(true);
    expect((await runCommandAsync(session, 'trackedChanges.acceptAll', {})).ok).toBe(true);
    return session;
  }

  test('Word resolved every element to plain runs, every Fallback dropped, the four edited shapes as the editor meant them', async () => {
    const bytes = await fixture('mc/mc-for-word-word365.docx');
    const xml = await mainXml(bytes);
    expect(count(xml, '<mc:AlternateContent')).toBe(0);
    expect(count(xml, '<mc:Fallback')).toBe(0);
    expect(count(xml, '<w:del ')).toBe(0);
    const session = await EditorDocument.open(bytes);
    expect(session.diagnostics.opaque).toEqual([]);
    const doc = session.doc;
    expect(texts(doc.child(CASE.text))).toEqual(['The XXEarth is flat']);
    expect(texts(doc.child(CASE.tabAndBreak))).toEqual(['Before', 'tab', 'hardBreak', 'after the break']);
    expect(texts(doc.child(CASE.picture))).toEqual(['A picture: ', 'image', ' after it']);
    expect(blipEmbedOf(session.ctx.ids.get(doc.child(CASE.picture).child(1).attrs['sourceId'] as string) as Element)).toBe('rId8');
    expect(texts(doc.child(CASE.ten))).toEqual(['I prefer ', 'w1', '4']);
    expect(doc.child(CASE.ten).child(1).marks.some((m) => m.type.name === 'bold')).toBe(true);
    expect(texts(doc.child(CASE.paragraphLevel))).toEqual(['bold ying', ' XXplain ying']);
    expect(doc.child(CASE.paragraphLevel).child(0).marks.some((m) => m.type.name === 'bold')).toBe(true);
    // Untouched, Word's save round-trips part by part.
    const result = await session.save();
    expect(result.selfCheck?.equal).toBe(true);
    expect(comparePackages(bytes, result.bytes)).toEqual([]);
  });

  test('the editor reads its own file as Word read it, in every case since core-ts CR-010 section 15 (x14ac and w16du understood)', async () => {
    const word = (await EditorDocument.open(await fixture('mc/mc-for-word-word365.docx'))).doc;
    const here = (await editedHere()).doc;
    // Word's save predates cases 9 to 12 of the fixture; the eight it holds, by index, as text (Word merges the runs
    // a resolved element leaves, so case 7 is one run there and two here).
    for (const i of [CASE.text, CASE.tabAndBreak, CASE.x14ac, CASE.x14acNoFallback, CASE.picture, CASE.ten, CASE.tenReordered, CASE.paragraphLevel]) {
      expect(here.child(i).textContent, `paragraph ${i}`).toBe(word.child(i).textContent);
      expect(texts(here.child(i)).filter((t) => t === 'image' || t === 'tab' || t === 'hardBreak')).toEqual(texts(word.child(i)).filter((t) => t === 'image' || t === 'tab' || t === 'hardBreak'));
    }
    expect(texts(word.child(CASE.x14ac))).toEqual(['unknown choice']);
    expect(texts(word.child(CASE.tenReordered))).toEqual(['I prefer w16du']);
  });
});

describe('check 49: what Office JS reads over a kept element\'s branch (Word 365, ED-009 section 9.2 item 4)', () => {
  // mc/mc-branches-check49-word365.docx: Jason's Word 365 save (2026-10-20) of mc-branches.docx after the Script Lab
  // snippet on the share (Office 2016/mc-fun/check49-script-lab.js) read every paragraph's text, searched for words
  // that live only in a Choice or only in a Fallback, inserted "YY" after the first "flat" hit through its Range,
  // and appended its report to the document (each line inserted after the heading, so the report reads backwards).
  test('paragraph.text is the selected branch\'s text, search finds a Choice\'s words and never a Fallback\'s, a Range from a hit edits the branch', async () => {
    const bytes = await fixture('mc/mc-branches-check49-word365.docx');
    const session = await EditorDocument.open(bytes);
    expect(session.diagnostics.opaque).toEqual([]);
    expect(count(await mainXml(bytes), '<mc:AlternateContent')).toBe(0);   // Word's save resolved every element
    const doc = session.doc;
    const report = doc.child(17).textContent;
    expect(report).toBe('CHECK 49 REPORT');
    const lines: string[] = [];
    for (let i = 18; i < doc.childCount; i++) lines.push(doc.child(i).textContent);
    const line = (prefix: string): string | undefined => lines.find((l) => l.startsWith(prefix));
    // The readers, over the file as opened (before the edit).
    expect(line('2: ')).toBe('2: "The Earth is flat"');
    expect(line('4: ')).toBe('4: "Before\\tafter the tab\\u000bafter the break"');
    expect(line('6: ')).toBe('6: "unknown choice"');               // x14ac understood by Word 365 (and the engine since CR-010 section 15)
    expect(line('8: ')).toBe('8: "nothing to select"');
    expect(line('10: ')).toBe('10: "A picture:  after it"');
    expect(line('12: ')).toBe('12: "I prefer w14"');
    expect(line('14: ')).toBe('14: "I prefer w16du"');              // w16du understood by Word 365
    expect(line('16: ')).toBe('16: "bold ying plain ying"');
    for (const t of ['flat', 'after the tab', 'unknown choice', 'nothing to select', 'after it', 'w14', 'bold ying', 'plain ying']) expect(line(`search "${t}"`), t).toBe(`search "${t}": 1 hit(s)`);
    expect(line('search "w16du"')).toBe('search "w16du": 2 hit(s)');   // case 7's branch and its label
    for (const t of ['round', 'Fallback text', 'fallback shown', 'yang']) expect(line(`search "${t}"`), t).toBe(`search "${t}": 0 hit(s)`);
    expect(line('first \'flat\' hit')).toBe('first \'flat\' hit: range text "flat", its paragraph "The Earth is flat"');
    expect(line('after insertText')).toBe('after insertText(\'YY\', \'After\'): paragraph "The Earth is flatYY"');
    // The save: the edit in the paragraph, the element resolved.
    expect(doc.child(2).textContent).toBe('The Earth is flatYY');
    // The editor's find agrees with Office JS over the unedited fixture: a Choice's words found, a Fallback's not.
    const base = await EditorDocument.open(await fixture('mc/mc-branches.docx'));
    expect(new Finder('after the tab', {}, base.ctx.ids).all(base.doc).length).toBe(1);
    expect(new Finder('round', {}, base.ctx.ids).all(base.doc).length).toBe(0);
    expect(new Finder('yang', {}, base.ctx.ids).all(base.doc).length).toBe(0);
  });
});
