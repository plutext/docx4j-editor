// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 4.9 step B5: section breaks and properties, on the paragraph attribute or the
// final section, with everything else in w:sectPr riding through.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage, type Element } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { EditorDocument, runCommand, commandNamed, stateFor, sectionAt } from '../src/index.mjs';
import { fixture, comparePackages, plain } from './helpers.mjs';

async function reloadBody(bytes: Uint8Array): Promise<wml.Body> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  return (await pkg.getMainDocumentPart().readContents()).body!;
}

function block(session: EditorDocument, i: number): { from: number; to: number } {
  const doc = session.doc;
  let pos = 0;
  for (let k = 0; k < i; k++) pos += doc.child(k).nodeSize;
  return { from: pos + 1, to: pos + 1 + doc.child(i).content.size };
}

async function saveAndReload(session: EditorDocument, original: Uint8Array): Promise<wml.Body> {
  const saved = await session.save();
  expect(saved.changed).toBe(true);
  const diffs = comparePackages(original, saved.bytes, [session.main.partName.storeName, 'docProps/app.xml']).filter((d) => d.name !== 'docProps/app.xml');
  expect(diffs).toEqual([]);
  return reloadBody(saved.bytes);
}

const sectPrOf = (p: Element): wml.SectPr | undefined => (p.value as wml.P).pPr?.sectPr;
const without = (s: wml.SectPr | undefined, key: keyof wml.SectPr): unknown => { const v = { ...(plain(s) ?? {}) } as Record<string, unknown>; delete v[key]; return v; };

describe('sections', () => {
  test('insertBreak gives the paragraph a copy of the section it is in; removeBreak takes it away', async () => {
    const original = await fixture('sections-multiple.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    const breakAt = before.findIndex((e) => sectPrOf(e) !== undefined);
    const i = 0;   // "This is a test.", in the first section
    expect(breakAt).toBeGreaterThan(i);
    const at = { from: block(session, i).from };
    expect(sectionAt(stateFor(session, at)).kind).toBe('paragraph');
    expect(commandNamed('section.insertBreak')!.isEnabled!(stateFor(session, at))).toBe(true);
    expect(runCommand(session, 'section.insertBreak', {}, at).ok).toBe(true);
    expect(commandNamed('section.insertBreak')!.isEnabled!(stateFor(session, at))).toBe(false);
    let after = (await saveAndReload(session, original)).content!;
    expect(without(sectPrOf(after[i]!), 'type')).toEqual(without(sectPrOf(before[breakAt]!), 'type'));
    expect(sectPrOf(after[i]!)?.type).toBeUndefined();
    expect(plain(after[breakAt])).toEqual(plain(before[breakAt]));
    // The final section is untouched by a paragraph-level edit.
    expect(plain((await reloadBody((await session.save()).bytes)).sectPr)).toEqual(plain(session.body.sectPr));

    expect(runCommand(session, 'section.removeBreak', undefined, at).ok).toBe(true);
    expect(runCommand(session, 'section.removeBreak', undefined, at).ok).toBe(false);
    after = (await session.save().then((s) => reloadBody(s.bytes))).content!;
    expect(sectPrOf(after[i]!)).toBeUndefined();
  });

  test("insertBreak with a type writes it on the section after the break, how that section starts (ECMA-376 17.6.22); the new section keeps the original's start", async () => {
    const original = await fixture('sections-multiple.docx');
    const session = await EditorDocument.open(original);
    const breakAt = session.body.content!.findIndex((e) => sectPrOf(e) !== undefined);
    runCommand(session, 'section.insertBreak', { type: 'continuous' }, { from: block(session, 0).from });
    const after = (await saveAndReload(session, original)).content!;
    expect(sectPrOf(after[0]!)?.type).toBeUndefined();
    expect(sectPrOf(after[breakAt]!)?.type?.val).toBe('continuous');
    // A next-page break after it, in the continuous section: the new part keeps "continuous", the rest starts on a new page.
    runCommand(session, 'section.insertBreak', { type: 'nextPage' }, { from: block(session, 1).from });
    const again = (await saveAndReload(session, original)).content!;
    expect(sectPrOf(again[1]!)?.type?.val).toBe('continuous');
    expect(sectPrOf(again[breakAt]!)?.type).toBeUndefined();
  });

  test('orientation on the final section swaps the page size and writes w:orient; the paragraphs are untouched', async () => {
    const original = await fixture('sections-multiple.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    const finalBefore = plain(session.body.sectPr) as wml.SectPr;
    const last = before.length - 1;
    const at = { from: block(session, last).from };
    expect(sectionAt(stateFor(session, at)).kind).toBe('final');
    expect(runCommand(session, 'section.orientation', { value: 'landscape' }, at).ok).toBe(true);
    expect(commandNamed('section.orientation')!.isActive!(stateFor(session, at), { value: 'landscape' })).toBe(true);
    const body = await saveAndReload(session, original);
    before.forEach((e, k) => expect(plain(body.content![k])).toEqual(plain(e)));
    expect(body.sectPr?.pgSz?.orient).toBe('landscape');
    expect(body.sectPr?.pgSz?.w).toBe(finalBefore.pgSz?.h);
    expect(body.sectPr?.pgSz?.h).toBe(finalBefore.pgSz?.w);
    expect(without(body.sectPr, 'pgSz')).toEqual(without(finalBefore, 'pgSz'));
    // Back to portrait restores the size.
    runCommand(session, 'section.orientation', { value: 'portrait' }, at);
    const again = await reloadBody((await session.save()).bytes);
    expect(plain(again.sectPr)).toEqual(finalBefore);
  });

  test('margins, page size, type and columns rewrite only their fields of the paragraph section', async () => {
    const original = await fixture('sections-multiple.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    const breakAt = before.findIndex((e) => sectPrOf(e) !== undefined);
    const at = { from: block(session, 0).from };
    runCommand(session, 'section.margins', { top: 72, left: 36 }, at);
    runCommand(session, 'section.pageSize', { width: 595.3, height: 841.9 }, at);
    runCommand(session, 'section.type', { value: 'oddPage' }, at);
    runCommand(session, 'section.columns', { count: 2 }, at);
    const after = (await saveAndReload(session, original)).content!;
    const s = sectPrOf(after[breakAt]!)!;
    const b = sectPrOf(before[breakAt]!)!;
    expect(s.pgMar?.top).toBe(1440);
    expect(s.pgMar?.left).toBe(720);
    expect(s.pgMar?.bottom).toBe(b.pgMar?.bottom);
    expect(s.pgSz?.w).toBe(11906);
    expect(s.pgSz?.h).toBe(16838);
    expect(s.pgSz?.orient).toBeUndefined();
    expect(s.type?.val).toBe('oddPage');
    expect(s.cols?.num).toBe(2);
    for (const k of Object.keys(plain(b)) as (keyof wml.SectPr)[]) {
      if (['pgMar', 'pgSz', 'type', 'cols', 'PARENT'].includes(k)) continue;
      expect(plain(s[k])).toEqual(plain(b[k]));
    }
    expect(plain((await reloadBody((await session.save()).bytes)).sectPr)).toEqual(plain(session.body.sectPr));
  });

  test('undo takes a final-section edit back with the rest', async () => {
    const original = await fixture('sections-multiple.docx');
    const session = await EditorDocument.open(original);
    const last = session.doc.childCount - 1;
    runCommand(session, 'section.columns', { count: 3 }, { from: block(session, last).from });
    expect((session.doc.attrs['finalSectPr'] as wml.SectPr).cols?.num).toBe(3);
    runCommand(session, 'history.undo');
    expect((session.doc.attrs['finalSectPr'] as wml.SectPr).cols?.num).not.toBe(3);
    const saved = await session.save();
    expect(saved.changed).toBe(false);
    expect(comparePackages(original, saved.bytes)).toEqual([]);
  });
});
