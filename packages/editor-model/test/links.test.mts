// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 5.8 item 4 (step C4): insert on a selection, at a caret with text, to an
// anchor; edit the target; remove; the reloaded document's w:hyperlink carries the relationship
// and the anchor; the edge rule of 5.5 in four cases.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage, runItemsOf, textOf as textOfElement, type Element } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, runCommandAsync, stateFor, linkAt, documentBookmarks, hyperlinkTarget, HYPERLINK_RELATIONSHIP, type ParagraphAttrs } from '../src/index.mjs';
import { fixture, comparePackages, plain } from './helpers.mjs';

function paragraph(doc: PMNode, pred: (attrs: ParagraphAttrs, node: PMNode) => boolean, n = 0): { node: PMNode; pos: number; from: number; to: number } {
  let found: { node: PMNode; pos: number; from: number; to: number } | undefined;
  let seen = 0;
  // Top-level paragraphs only: the reloaded body's content array is indexed the same way.
  doc.forEach((node, pos) => {
    if (found || node.type.name !== 'paragraph') return;
    if (pred(node.attrs as ParagraphAttrs, node) && seen++ === n) found = { node, pos, from: pos + 1, to: pos + 1 + node.content.size };
  });
  if (!found) throw new Error('no such paragraph');
  return found;
}

const plainText = (a: ParagraphAttrs, node: PMNode): boolean => a.sourceId !== null && node.childCount === 1 && node.firstChild!.isText && !node.firstChild!.marks.some((m) => m.type.name === 'link') && node.textContent.length > 10;

/** The hyperlinks of a reloaded paragraph: their attributes and text. */
async function linksOf(bytes: Uint8Array, index: number): Promise<{ id?: string; anchor?: string; history?: boolean; text: string }[]> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  const body = (await pkg.getMainDocumentPart().readContents()).body!;
  const p = body.content![index]!;
  return (runItemsOf(p.value as object) ?? []).filter((e) => e.name.localPart === 'hyperlink').map((h) => {
    const v = plain(h.value) as wml.P.Hyperlink;
    return { ...(v.id !== undefined ? { id: v.id } : {}), ...(v.anchor !== undefined ? { anchor: v.anchor } : {}), ...(v.history !== undefined ? { history: v.history } : {}), text: textOfElement(h as Element) };
  });
}

async function relationshipTarget(bytes: Uint8Array, id: string): Promise<{ target: string; type: string; mode: string | undefined } | undefined> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  const rel = pkg.getMainDocumentPart().relationshipsPart?.getRelationshipById(id);
  return rel ? { target: rel.target, type: rel.type, mode: rel.targetMode } : undefined;
}

const linkMarks = (node: PMNode): { rId: string | null; anchor: string | null; sourceId: string | null }[] => {
  const out: { rId: string | null; anchor: string | null; sourceId: string | null }[] = [];
  node.descendants((n) => { for (const m of n.marks) if (m.type.name === 'link') out.push({ rId: m.attrs['rId'] as string | null, anchor: m.attrs['anchor'] as string | null, sourceId: m.attrs['sourceId'] as string | null }); return true; });
  return out;
};

describe('range.hyperlink (5.5)', () => {
  test('insert on a selection: the relationship is made in the command, the mark carries it, the reloaded paragraph has the w:hyperlink with r:id and w:history; only the main part and the relationships change', async () => {
    const original = await fixture('loadAndSave.docx');
    const session = await EditorDocument.open(original);
    const p = paragraph(session.doc, plainText);
    const index = session.doc.resolve(p.pos).index(0);
    const word = { from: p.from, to: p.from + 5 };
    expect(runCommand.bind(null, session, 'range.hyperlink', { address: 'https://docx4j.org/' }, word)).toThrow(/runCommandAsync/);
    const result = await runCommandAsync(session, 'range.hyperlink', { address: 'https://docx4j.org/' }, word);
    expect(result.ok).toBe(true);
    const marks = linkMarks(session.doc.nodeAt(p.pos)!);
    expect(marks).toHaveLength(1);
    expect(marks[0]!.sourceId).toBeNull();
    expect(marks[0]!.rId).toMatch(/^rId\d+$/);
    expect(linkAt(stateFor(session, { from: word.from + 1 }))).toMatchObject({ from: word.from, to: word.to, address: 'https://docx4j.org/', anchor: null });
    const saved = await session.save();
    expect(comparePackages(original, saved.bytes, [session.main.partName.storeName, 'docProps/app.xml'])).toEqual([]);
    const links = await linksOf(saved.bytes, index);
    expect(links).toEqual([{ id: marks[0]!.rId, history: true, text: p.node.textContent.slice(0, 5) }]);
    expect(await relationshipTarget(saved.bytes, marks[0]!.rId!)).toEqual({ target: 'https://docx4j.org/', type: HYPERLINK_RELATIONSHIP, mode: 'External' });
    // Undo takes the mark off; the relationship stays, and the document is as imported.
    runCommand(session, 'history.undo');
    expect(session.doc.eq(session.imported)).toBe(true);
    expect(session.main.relationshipsPart!.getRelationshipById(marks[0]!.rId!)).toBeDefined();
  });

  test('at a caret with text: inserted and linked; to an anchor: w:anchor and no relationship; both at once', async () => {
    const original = await fixture('loadAndSave.docx');
    const session = await EditorDocument.open(original);
    const bookmarks = documentBookmarks(session.doc, session).map((b) => b.name);
    expect(bookmarks).toContain('_Toc230018410');
    const p = paragraph(session.doc, plainText);
    const index = session.doc.resolve(p.pos).index(0);
    // A bare caret with no text: nothing to link.
    expect((await runCommandAsync(session, 'range.hyperlink', { address: 'https://x.example/' }, { from: p.to })).ok).toBe(false);
    expect((await runCommandAsync(session, 'range.hyperlink', { subAddress: '_Toc230018410', text: ' see section 2' }, { from: p.to })).ok).toBe(true);
    expect(session.doc.nodeAt(p.pos)!.textContent).toBe(`${p.node.textContent} see section 2`);
    expect(linkMarks(session.doc.nodeAt(p.pos)!)).toEqual([{ rId: null, anchor: '_Toc230018410', sourceId: null }]);
    expect((await runCommandAsync(session, 'range.hyperlink', { address: 'https://x.example/', subAddress: 'top', text: 'both' }, { from: p.from })).ok).toBe(true);
    const saved = await session.save();
    const links = await linksOf(saved.bytes, index);
    expect(links).toHaveLength(2);
    expect(links[0]).toMatchObject({ anchor: 'top', history: true, text: 'both' });
    expect(links[0]!.id).toMatch(/^rId/);
    expect(links[1]).toEqual({ anchor: '_Toc230018410', history: true, text: ' see section 2' });
  });

  test('edit the target of an existing link (replaced with a new holder), then remove it keeping the text', async () => {
    const original = await fixture('loadAndSave.docx');
    const session = await EditorDocument.open(original);
    const linked = paragraph(session.doc, (_a, node) => node.textContent === 'Hyperlink');
    const index = session.doc.resolve(linked.pos).index(0);
    const before = linkAt(stateFor(session, { from: linked.from + 2 }))!;
    expect(before).toMatchObject({ from: linked.from, to: linked.to, rId: 'rId7', address: 'https://example.com/', text: 'Hyperlink' });
    expect((await runCommandAsync(session, 'range.hyperlink', { address: 'https://docx4j.org/' }, { from: linked.from + 2 })).ok).toBe(true);
    const after = linkAt(stateFor(session, { from: linked.from + 2 }))!;
    expect(after.address).toBe('https://docx4j.org/');
    expect(after.rId).not.toBe('rId7');
    expect(session.doc.nodeAt(linked.pos)!.textContent).toBe('Hyperlink');
    let saved = await session.save();
    let links = await linksOf(saved.bytes, index);
    expect(links).toHaveLength(1);
    expect(links[0]).toMatchObject({ id: after.rId, history: true, text: 'Hyperlink' });
    // Remove: the text stays, the holder goes.
    expect((await runCommandAsync(session, 'range.hyperlink', {}, { from: linked.from + 2 })).ok).toBe(true);
    expect(linkAt(stateFor(session, { from: linked.from + 2 }))).toBeUndefined();
    expect(session.doc.nodeAt(linked.pos)!.textContent).toBe('Hyperlink');
    saved = await session.save();
    links = await linksOf(saved.bytes, index);
    expect(links).toEqual([]);
    expect(comparePackages(original, saved.bytes, [session.main.partName.storeName, 'docProps/app.xml'])).toEqual([]);
  });

  test("a link's look, as core-ts 0.2.1's setter gives it (CR-002 section 25): the Hyperlink style over another character style, the direct colour gone, bold kept; the style's definition added where the document lacks it; the style taken off with the link", async () => {
    const original = await fixture('sections-multiple.docx');
    const stylesXml = async (bytes: Uint8Array): Promise<string> => (await WordprocessingMLPackage.load(bytes)).getMainDocumentPart().styleDefinitionsPart!.getXml();
    expect(await stylesXml(original)).not.toContain('w:styleId="Hyperlink"');
    const session = await EditorDocument.open(original);
    const p = paragraph(session.doc, (_a, node) => node.textContent === 'This is a test.');
    let at = -1;   // "test", after the paragraph's leading atoms
    p.node.forEach((child, offset) => { if (at < 0 && child.isText && child.text!.includes('test')) at = p.from + offset + child.text!.indexOf('test'); });
    const word = { from: at, to: at + 4 };
    for (const [name, value] of [['font.style', 'Strong'], ['font.color', '#FF0000'], ['font.bold', true]] as const) expect(runCommand(session, name, { value }, word).ok).toBe(true);
    expect((await runCommandAsync(session, 'range.hyperlink', { address: 'https://docx4j.org/' }, word)).ok).toBe(true);
    const marksOf = (): Record<string, unknown> => Object.fromEntries(session.doc.nodeAt(word.from)!.marks.map((m) => [m.type.name, m.attrs]));
    expect(marksOf()).toMatchObject({ runStyle: { id: 'Hyperlink' }, bold: {}, link: { sourceId: null } });
    expect(marksOf()).not.toHaveProperty('font');
    let saved = await session.save();
    expect(await stylesXml(saved.bytes)).toContain('w:styleId="Hyperlink"');
    const main = async (): Promise<string> => (await WordprocessingMLPackage.load(saved.bytes)).getMainDocumentPart().getXml();
    const link = (await main()).match(/<w:hyperlink [^>]*>.*?<\/w:hyperlink>/)![0];
    expect(link).toMatch(/<w:rPr><w:rStyle w:val="Hyperlink"\/><w:b\/>.*<\/w:rPr><w:t>test<\/w:t>/);
    expect(link).not.toMatch(/Strong|w:color/);
    // Removed: the style goes with the link; the colour does not come back.
    expect((await runCommandAsync(session, 'range.hyperlink', {}, { from: word.from + 1 })).ok).toBe(true);
    expect(Object.keys(marksOf()).filter((k) => k !== 'runSource')).toEqual(['bold']);
    saved = await session.save();
    const xml = await main();
    expect(xml).not.toMatch(/<w:hyperlink /);
    expect(xml).toMatch(/<w:r><w:rPr><w:b\/><w:bCs\/><w:noProof\/><\/w:rPr><w:t>test<\/w:t><\/w:r>/);
  });

  test("a range over part of a link, as Word does it (core-ts 0.2.2, check 16; ED-005 section 1.6): F takes the whole link in, G leaves the rest of it plain, H removes all of it", async () => {
    const original = await fixture('tracked-changes.docx');
    /** The paragraph "Some text." as (text, link address or '', style) pieces. */
    const pieces = (session: EditorDocument, p: { pos: number }): string[] => {
      const out: string[] = [];
      session.doc.nodeAt(p.pos)!.forEach((n) => {
        const link = n.marks.find((m) => m.type.name === 'link');
        const style = n.marks.find((m) => m.type.name === 'runStyle')?.attrs['id'] as string | undefined;
        out.push(`${n.text ?? ''}|${link ? hyperlinkTarget(session, link.attrs['rId'] as string) : ''}|${style ?? ''}`);
      });
      return out;
    };
    const setUp = async (range: [number, number], address: string): Promise<{ session: EditorDocument; p: { pos: number; from: number } }> => {
      const session = await EditorDocument.open(original);
      const p = paragraph(session.doc, (_a, node) => node.textContent === 'Some text.');
      expect((await runCommandAsync(session, 'range.hyperlink', { address }, { from: p.from + range[0], to: p.from + range[1] })).ok).toBe(true);
      return { session, p };
    };
    // F: "text" linked to f1, then "Some te" linked to f2: one link to f2 over "Some text".
    let { session, p } = await setUp([5, 9], 'https://f1.example/');
    expect((await runCommandAsync(session, 'range.hyperlink', { address: 'https://f2.example/' }, { from: p.from, to: p.from + 7 })).ok).toBe(true);
    expect(pieces(session, p)).toEqual(['Some text|https://f2.example/|Hyperlink', '.||']);
    // G: "Some text." linked to g1, then "me te" linked to g2: "So" and "xt." plain and unstyled, "me te" to g2.
    ({ session, p } = await setUp([0, 10], 'https://g1.example/'));
    expect((await runCommandAsync(session, 'range.hyperlink', { address: 'https://g2.example/' }, { from: p.from + 2, to: p.from + 7 })).ok).toBe(true);
    expect(pieces(session, p)).toEqual(['So||', 'me te|https://g2.example/|Hyperlink', 'xt.||']);
    // H: "Some text." linked to h1, then the link removed over "me te": no link and no style left.
    ({ session, p } = await setUp([0, 10], 'https://h1.example/'));
    expect((await runCommandAsync(session, 'range.hyperlink', {}, { from: p.from + 2, to: p.from + 7 })).ok).toBe(true);
    expect(pieces(session, p)).toEqual(['Some text.||']);
  });

  test('the edge rule (section 9 item 19): typing at the end leaves the link, inside extends it, deleting the whole text removes it, Backspace at the end deletes the last character and stays in it', async () => {
    const session = await EditorDocument.open(await fixture('loadAndSave.docx'));
    const linked = paragraph(session.doc, (_a, node) => node.textContent === 'Hyperlink');
    const at = (offset: number): number => linked.from + offset;
    runCommand(session, 'insertText', { text: '!' }, { from: at(9) });
    expect(linkAt(stateFor(session, { from: at(9) }))).toMatchObject({ from: at(0), to: at(9) });
    expect(session.doc.nodeAt(linked.pos)!.lastChild!.marks.some((m) => m.type.name === 'link')).toBe(false);
    runCommand(session, 'insertText', { text: 'X' }, { from: at(5) });
    expect(linkAt(stateFor(session, { from: at(5) }))).toMatchObject({ from: at(0), to: at(10), text: 'HyperXlink' });
    // Backspace at the end of the link: the last character goes, the link stays.
    const state = stateFor(session, { from: at(10) });
    session.replace(state.tr.delete(at(9), at(10)).doc);
    expect(linkAt(stateFor(session, { from: at(9) }))).toMatchObject({ from: at(0), to: at(9), text: 'HyperXlin' });
    // The whole text deleted: no link.
    session.replace(stateFor(session).tr.delete(at(0), at(9)).doc);
    expect(linkAt(stateFor(session, { from: at(0) }))).toBeUndefined();
    expect(session.doc.nodeAt(linked.pos)!.textContent).toBe('!');
    const saved = await session.save();
    const index = session.doc.resolve(linked.pos).index(0);
    expect(await linksOf(saved.bytes, index)).toEqual([]);
  });
});
