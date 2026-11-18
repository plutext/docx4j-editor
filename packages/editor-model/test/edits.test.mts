// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 3.7 test 4: scripted edits, deep-equal. Each edit is applied to the ProseMirror
// document with prosemirror-transform (E1.b's commands come later), pushed through the Yjs store,
// saved, and the saved main part is reloaded and compared with what the edit should have done:
// the edited element changed only where the edit says, every other element equal to the original.
import { describe, test, expect } from 'vitest';
import { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { WordprocessingMLPackage, textOf as textOfElement, typeNameOf, runItemsOf, deepCopy, applyRunOptions, type Element } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { EditorDocument, revive } from '../src/index.mjs';
import { fixture, comparePackages, plain, textOf } from './helpers.mjs';

/** A transaction over a bare document (insertText and the marks-at-position rules live on Transaction, not Transform). */
function tr(doc: PMNode) {
  return EditorState.create({ doc }).tr;
}

async function reloadBody(bytes: Uint8Array): Promise<wml.Body> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  return (await pkg.getMainDocumentPart().readContents()).body!;
}

/** The position just inside the end of block `i`'s content. */
function endOfBlock(doc: PMNode, i: number): number {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += doc.child(k).nodeSize;
  return pos + 1 + doc.child(i).content.size;
}

function startOfBlock(doc: PMNode, i: number): number {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += doc.child(k).nodeSize;
  return pos;
}

/** Save through the store and check that only the main part (and app.xml) changed. */
async function saveEdited(session: EditorDocument, original: Uint8Array, edited: PMNode) {
  session.replace(edited);
  const saved = await session.save();
  expect(saved.changed).toBe(true);
  const mainName = session.main.partName.storeName;   // 'word/document.xml' usually; one fixture says document22.xml
  const diffs = comparePackages(original, saved.bytes, [mainName, 'docProps/app.xml']).filter((d) => d.name !== 'docProps/app.xml');
  expect(diffs).toEqual([]);
  const app = await (await WordprocessingMLPackage.load(saved.bytes)).docPropsExtendedPart!.getXml();
  expect(app).toMatch(/Application>docx4j-editor</);
  expect(app).toMatch(/AppVersion>00\.0100</);
  return saved.bytes;
}

/** Every element of `after` other than the given indexes equals the original at the same index. */
function expectOthersEqual(before: Element[], after: Element[], changed: number[], inserted = 0): void {
  expect(after.length).toBe(before.length + inserted);
  for (let i = 0, j = 0; i < before.length; i++, j++) {
    if (changed.includes(j)) { if (inserted > 0 && !changed.includes(i)) i--; continue; }
    expect(plain(after[j])).toEqual(plain(before[i]));
  }
}

const runsOf = (p: Element): Element<wml.R>[] => (runItemsOf(p.value as object) ?? []).filter((e) => typeNameOf(e) === 'org_docx4j_wml.R') as Element<wml.R>[];

describe('scripted edits', () => {
  test('typing at the end of a paragraph extends its last run and keeps everything else', async () => {
    const original = await fixture('HelloWordOnline.docx');
    const session = await EditorDocument.open(original);
    const before = session.body.content!.map((e) => e);
    const doc = session.doc;
    const edited = tr(doc).insertText(' (edited)', endOfBlock(doc, 0)).doc;
    const body = await reloadBody(await saveEdited(session, original, edited));
    const after = body.content!;
    expectOthersEqual(before, after, [0]);
    const p0 = after[0]!;
    expect(textOfElement(p0)).toBe('Hello World from Word Online (edited)');
    // Same paragraph properties, same number of runs, the last run's properties untouched.
    expect(plain((p0.value as wml.P).pPr)).toEqual(plain((before[0]!.value as wml.P).pPr));
    const runsBefore = runsOf(before[0]!);
    const runsAfter = runsOf(p0);
    expect(runsAfter.length).toBe(runsBefore.length);
    expect(plain(runsAfter.at(-1)!.value.rPr)).toEqual(plain(runsBefore.at(-1)!.value.rPr));
    // The bookmarks that opened the paragraph are still there, in place.
    const kinds = (runItemsOf(p0.value as object) ?? []).map((e) => e.name.localPart);
    expect(kinds.slice(0, 2)).toEqual((runItemsOf(before[0]!.value as object) ?? []).map((e) => e.name.localPart).slice(0, 2));
  });

  test('bolding a word splits the run in three; the copies keep the original run properties', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const before = session.body.content!.map((e) => e);
    const doc = session.doc;
    const i = 4;   // "Some text."
    expect(textOf(doc.child(i))).toBe('Some text.');
    const from = startOfBlock(doc, i) + 1 + 'Some '.length;
    const edited = tr(doc).addMark(from, from + 'text'.length, session.schema.marks['bold']!.create()).doc;
    const body = await reloadBody(await saveEdited(session, original, edited));
    const after = body.content!;
    expectOthersEqual(before, after, [i]);
    const runsBefore = runsOf(before[i]!);
    expect(runsBefore.length).toBe(1);
    const runsAfter = runsOf(after[i]!);
    expect(runsAfter.map((r) => textOfElement(r))).toEqual(['Some ', 'text', '.']);
    const origRPr = runsBefore[0]!.value.rPr;
    expect(plain(runsAfter[0]!.value.rPr)).toEqual(plain(origRPr));
    expect(plain(runsAfter[2]!.value.rPr)).toEqual(plain(origRPr));
    const expectedBold = applyRunOptions(origRPr ? deepCopy(origRPr) : { TYPE_NAME: 'org_docx4j_wml.RPr' }, { bold: true });
    // The reloaded tree carries TYPE_NAMEs the builder did not write; compare the shape that matters.
    const boldRPr = runsAfter[1]!.value.rPr!;
    expect(boldRPr.b).toBeDefined();
    expect(boldRPr.bCs).toBeDefined();
    const strip = (v: unknown): unknown => JSON.parse(JSON.stringify(v, (k, x) => (k === 'TYPE_NAME' || k === 'PARENT' ? undefined : x)));
    expect(strip(boldRPr)).toEqual(strip(expectedBold));
  });

  test('a new paragraph with a style is built from the builders', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const before = session.body.content!.map((e) => e);
    const doc = session.doc;
    const i = 4;
    const newP = session.schema.nodes['paragraph']!.create({ pStyle: 'Heading1' }, session.schema.text('New heading'));
    const edited = tr(doc).insert(startOfBlock(doc, i + 1), newP).doc;
    const body = await reloadBody(await saveEdited(session, original, edited));
    const after = body.content!;
    expect(after.length).toBe(before.length + 1);
    for (let k = 0; k <= i; k++) expect(plain(after[k])).toEqual(plain(before[k]));
    for (let k = i + 1; k < before.length; k++) expect(plain(after[k + 1])).toEqual(plain(before[k]));
    const p = after[i + 1]!.value as wml.P;
    expect(p.pPr?.pStyle?.val).toBe('Heading1');
    expect(textOfElement(after[i + 1]!)).toBe('New heading');
    expect(runsOf(after[i + 1]!).length).toBe(1);
  });

  test('changing the alignment writes w:jc and leaves the rest of w:pPr alone', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const before = session.body.content!.map((e) => e);
    const doc = session.doc;
    const i = 2;   // Heading1 paragraph
    const node = doc.child(i);
    const edited = tr(doc).setNodeMarkup(startOfBlock(doc, i), undefined, { ...node.attrs, jc: 'center' }).doc;
    const body = await reloadBody(await saveEdited(session, original, edited));
    const after = body.content!;
    expectOthersEqual(before, after, [i]);
    const pPrAfter = plain((after[i]!.value as wml.P).pPr) as wml.PPr;
    expect(pPrAfter.jc?.val).toBe('center');
    delete pPrAfter.jc;
    const pPrBefore = plain((before[i]!.value as wml.P).pPr) as wml.PPr;
    delete pPrBefore.jc;
    expect(pPrAfter).toEqual(pPrBefore);
    // Runs untouched.
    expect(plain(runsOf(after[i]!))).toEqual(plain(runsOf(before[i]!)));
  });

  test('deleting an image drops the drawing and nothing else', async () => {
    const original = await fixture('images.docx');
    const session = await EditorDocument.open(original);
    const before = session.body.content!.map((e) => e);
    const doc = session.doc;
    expect(doc.child(0).child(0).type.name).toBe('image');
    const edited = tr(doc).delete(1, 2).doc;
    expect(doc.child(0).childCount).toBe(1);
    expect(edited.child(0).childCount).toBe(0);
    const body = await reloadBody(await saveEdited(session, original, edited));
    const after = body.content!;
    expectOthersEqual(before, after, [0]);
    expect(JSON.stringify(plain(after[0]))).not.toContain('drawing');
    expect(plain((after[0]!.value as wml.P).pPr)).toEqual(plain((before[0]!.value as wml.P).pPr));
  });

  test('typing inside a hyperlink keeps the w:hyperlink around the run', async () => {
    const original = await fixture('hyperlink_dupe.docx');
    const session = await EditorDocument.open(original);
    const before = session.body.content!.map((e) => e);
    const doc = session.doc;
    const i = 1;   // the TOC entry, a hyperlink around "H1", a tab and a PAGEREF field
    const first = doc.child(i).child(0);
    expect(first.marks.some((m) => m.type.name === 'link')).toBe(true);
    const at = startOfBlock(doc, i) + 1 + 1;   // between "H" and "1": inside the link's text, so the link mark is inherited
    const edited = tr(doc).insertText('X', at).doc;
    expect(edited.child(i).child(0).marks.some((m) => m.type.name === 'link')).toBe(true);
    const body = await reloadBody(await saveEdited(session, original, edited));
    const after = body.content!;
    expectOthersEqual(before, after, [i]);
    const linksBefore = (runItemsOf(before[i]!.value as object) ?? []).filter((e) => e.name.localPart === 'hyperlink');
    const linksAfter = (runItemsOf(after[i]!.value as object) ?? []).filter((e) => e.name.localPart === 'hyperlink');
    expect(linksAfter.length).toBe(linksBefore.length);
    expect(linksAfter.length).toBeGreaterThan(0);
    const attrsOf = (h: Element): unknown => { const v = plain(h.value) as Record<string, unknown>; delete v['content']; return v; };
    expect(linksAfter.map(attrsOf)).toEqual(linksBefore.map(attrsOf));
    expect(textOfElement(linksAfter[0]!)).toBe(textOfElement(linksBefore[0]!).replace('H1', 'HX1'));
    // The field's runs are still inside the hyperlink, in order: opaque items were re-emitted in place.
    const kindsAfter = (runItemsOf(linksAfter[0]!.value as object) ?? []).map((e) => e.name.localPart);
    const kindsBefore = (runItemsOf(linksBefore[0]!.value as object) ?? []).map((e) => e.name.localPart);
    expect(kindsAfter).toEqual(kindsBefore);
  });

  test('typing at the end of linked text goes outside the hyperlink, as the screen shows it', async () => {
    // The link mark is not inclusive, so text typed after "H1" is unlinked on screen; the file
    // agrees: the run goes after the hyperlink, which splits because the link continues. E2's
    // link editing decides whether that should instead extend the link.
    const original = await fixture('hyperlink_dupe.docx');
    const session = await EditorDocument.open(original);
    const doc = session.doc;
    const i = 1;
    const first = doc.child(i).child(0);
    const edited = tr(doc).insertText('X', startOfBlock(doc, i) + 1 + first.nodeSize).doc;
    expect(edited.child(i).child(1).marks.some((m) => m.type.name === 'link')).toBe(false);
    const body = await reloadBody(await saveEdited(session, original, edited));
    const items = runItemsOf(body.content![i]!.value as object) ?? [];
    expect(items.map((e) => e.name.localPart)).toEqual(['hyperlink', 'r', 'hyperlink']);
    expect(textOfElement(items[1]!)).toBe('X');
    expect(textOfElement(items[0]!)).toBe('H1');
  });

  test('editing a paragraph with tracked changes keeps the insertion wrapper and the deletion as it was (its text w:delText; ED-005 section 12.3)', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const before = session.body.content!.map((e) => e);
    const i = before.findIndex((e) => JSON.stringify(plain(e)).includes('org_docx4j_wml.RunIns'));
    expect(i).toBeGreaterThanOrEqual(0);
    const doc = session.doc;
    const edited = tr(doc).insertText(' [reviewed]', endOfBlock(doc, i)).doc;
    const body = await reloadBody(await saveEdited(session, original, edited));
    const after = body.content!;
    expectOthersEqual(before, after, [i]);
    const revisions = (p: Element): Element[] => (runItemsOf(p.value as object) ?? []).filter((e) => e.name.localPart === 'ins' || e.name.localPart === 'del');
    const revBefore = revisions(before[i]!);
    const revAfter = revisions(after[i]!);
    expect(revAfter.map((e) => e.name.localPart)).toEqual(revBefore.map((e) => e.name.localPart));
    const meta = (e: Element): unknown => { const v = plain(e.value) as Record<string, unknown>; delete v['customXmlOrSmartTagOrSdt']; return v; };
    expect(revAfter.map(meta)).toEqual(revBefore.map(meta));
    // The deletion's content is exactly what it was.
    const delBefore = revBefore.find((e) => e.name.localPart === 'del')!;
    const delAfter = revAfter.find((e) => e.name.localPart === 'del')!;
    expect(plain(delAfter)).toEqual(plain(delBefore));
    expect(textOfElement(after[i]!)).toBe(textOfElement(before[i]!) + ' [reviewed]');
  });

  test('typing inside an inline content control rebuilds it with the same properties', async () => {
    const original = await fixture('invoice2013.docx');
    const session = await EditorDocument.open(original);
    const before = session.body.content!.map((e) => e);
    const doc = session.doc;
    let blockIndex = -1;
    let sdtPos = -1;
    doc.forEach((block, offset, idx) => {
      if (blockIndex >= 0) return;
      block.forEach((child, childOffset) => {
        if (sdtPos < 0 && child.type.name === 'sdtInline' && child.content.size > 0) { blockIndex = idx; sdtPos = offset + 1 + childOffset; }
      });
    });
    expect(blockIndex).toBeGreaterThanOrEqual(0);
    const sdt = doc.nodeAt(sdtPos)!;
    const edited = tr(doc).insertText('!', sdtPos + 1 + sdt.content.size).doc;
    const body = await reloadBody(await saveEdited(session, original, edited));
    const after = body.content!;
    expectOthersEqual(before, after, [blockIndex]);
    const sdtsBefore = (runItemsOf(before[blockIndex]!.value as object) ?? []).filter((e) => typeNameOf(e) === 'org_docx4j_wml.SdtRun') as Element<wml.SdtRun>[];
    const sdtsAfter = (runItemsOf(after[blockIndex]!.value as object) ?? []).filter((e) => typeNameOf(e) === 'org_docx4j_wml.SdtRun') as Element<wml.SdtRun>[];
    expect(sdtsAfter.length).toBe(sdtsBefore.length);
    expect(sdtsAfter.map((s) => plain(s.value.sdtPr))).toEqual(sdtsBefore.map((s) => plain(s.value.sdtPr)));
    expect(textOfElement(after[blockIndex]!)).toBe(textOfElement(before[blockIndex]!).replace(textOfElement(sdtsBefore[0]!), textOfElement(sdtsBefore[0]!) + '!'));
  });

  test("changing the final section's orientation rewrites the body's w:sectPr from the attribute", async () => {
    const original = await fixture('sections-multiple.docx');
    const session = await EditorDocument.open(original);
    const before = session.body.content!.map((e) => e);
    const sectBefore = plain(session.body.sectPr) as wml.SectPr;
    const doc = session.doc;
    const finalSectPr = revive(doc.attrs['finalSectPr']) as wml.SectPr;
    expect(finalSectPr.pgSz).toBeDefined();
    finalSectPr.pgSz = { ...finalSectPr.pgSz, orient: 'landscape', w: finalSectPr.pgSz!.h, h: finalSectPr.pgSz!.w };
    const edited = session.schema.topNodeType.create({ finalSectPr }, doc.content);
    const body = await reloadBody(await saveEdited(session, original, edited));
    expectOthersEqual(before, body.content!, []);
    const sectAfter = plain(body.sectPr) as wml.SectPr;
    expect(sectAfter.pgSz?.orient).toBe('landscape');
    expect(sectAfter.pgSz?.w).toBe(sectBefore.pgSz?.h);
    const rest = (s: wml.SectPr): unknown => { const v = { ...s } as Record<string, unknown>; delete v['pgSz']; return v; };
    expect(rest(sectAfter)).toEqual(rest(sectBefore));
  });
});
