// SPDX-License-Identifier: GPL-3.0-or-later
//
// Page numbers in every section's footer (ED-005 section 12.95): a footer made where a section has
// none to inherit, a footer the file has extended unless it already numbers its pages, the
// command idle when every section is numbered; the saved file reloads with the footers in place.
import { describe, test, expect } from 'vitest';
import { ZipPartStore } from '@docx4j/core-ts';
import { EditorDocument, runCommandAsync, hasPageField, planPageNumbers, pageNumberParagraph } from '../src/index.mjs';
import { fixture, comparePackages } from './helpers.mjs';
import { room } from './room-harness.mjs';

const partText = (bytes: Uint8Array, name: string): string => new TextDecoder().decode(new ZipPartStore(bytes).loadSync(name.replace(/^\//, '')));

describe('document.pageNumbers', () => {
  test('a document with headers and no footers, odd and even pages: a default and an even footer made and referenced from the first section, inherited by the rest; a second run does nothing', async () => {
    const original = await fixture('headers.docx');   // five sections, six header parts, w:evenAndOddHeaders
    const session = await EditorDocument.open(original);
    const plan = await planPageNumbers(session);
    expect(plan.extend).toEqual([]);
    expect(plan.create).toEqual([{ section: 0, variant: 'default' }, { section: 0, variant: 'even' }]);
    expect((await runCommandAsync(session, 'document.pageNumbers')).ok).toBe(true);
    const footers = (await session.stories.headers()).filter((h) => h.kind === 'footer');
    expect(footers.length).toBe(2);
    for (const f of footers) expect(hasPageField(f.element)).toBe(true);
    // Every section's pages, odd and even, get a footer; the last section inherits the first's.
    const last = session.stories.sections().length - 1;
    expect((await session.stories.headerFooterFor(last, 7, false)).footer?.rId).toBe(footers[0]!.rId);
    expect((await session.stories.headerFooterFor(last, 8, false)).footer?.rId).toBe(footers[1]!.rId);
    expect((await planPageNumbers(session)).create).toEqual([]);
    const saved = await session.save();
    const diffs = comparePackages(original, saved.bytes, ['word/document.xml', 'word/styles.xml', 'docProps/app.xml']);
    expect(diffs.map((d) => `${d.name}:${d.reason}`).sort()).toEqual(['word/footer1.xml:added', 'word/footer2.xml:added']);
    const footer = partText(saved.bytes, '/word/footer1.xml');
    expect(footer).toContain('<w:pStyle w:val="Footer"/>');
    expect(footer).toContain('<w:jc w:val="center"/>');
    expect(footer).toMatch(/<w:fldSimple w:instr=" PAGE "><w:r><w:t>1<\/w:t><\/w:r><\/w:fldSimple>/);
    const main = partText(saved.bytes, '/word/document.xml');
    expect(main).toMatch(/<w:footerReference [^>]*w:type="default"[^>]*\/><w:footerReference [^>]*w:type="even"[^>]*\/>/);
    expect(main.match(/<w:footerReference /g)?.length).toBe(2);
    expect(partText(saved.bytes, '/word/styles.xml')).toContain('w:styleId="Footer"');
    // Reloaded: the footers are read from the file, and the command finds nothing to do.
    const again = await EditorDocument.open(saved.bytes);
    expect((await again.stories.headers()).filter((h) => h.kind === 'footer').length).toBe(2);
    expect((await again.stories.headerFooterFor(last, 7, false)).footer).toBeDefined();
    expect(await planPageNumbers(again)).toEqual({ extend: [], create: [] });
    expect((await runCommandAsync(again, 'document.pageNumbers')).ok).toBe(false);
    expect(comparePackages(saved.bytes, (await again.save()).bytes)).toEqual([]);
  });

  test('a document whose first section has footers without a number: the default footer gets the paragraph at its end, the second section inherits it, no part is made', async () => {
    const original = await fixture('sections-multiple.docx');   // section 0: default, even and first footers; section 1: a header alone
    const session = await EditorDocument.open(original);
    const plan = await planPageNumbers(session);
    expect(plan.create).toEqual([]);
    expect(plan.extend.map((f) => f.rId)).toEqual(['rId10']);   // the default footer; no odd and even pages, so the even one is left alone
    expect((await runCommandAsync(session, 'document.pageNumbers')).ok).toBe(true);
    const saved = await session.save();
    const diffs = comparePackages(original, saved.bytes, ['docProps/app.xml', 'word/styles.xml']);
    expect(diffs.map((d) => `${d.name}:${d.reason}`)).toEqual(['word/footer2.xml:bytes']);
    const footer = partText(saved.bytes, '/word/footer2.xml');
    expect(footer).toContain('Just a footer.');
    expect(footer.indexOf('Just a footer.')).toBeLessThan(footer.indexOf('<w:fldSimple w:instr=" PAGE ">'));
    const again = await EditorDocument.open(saved.bytes);
    expect(await planPageNumbers(again)).toEqual({ extend: [], create: [] });
  });

  test('a document whose footers number their pages with complex fields is left as it is', async () => {
    const original = await fixture('loadAndSave.docx');
    const session = await EditorDocument.open(original);
    expect(await planPageNumbers(session)).toEqual({ extend: [], create: [] });
    expect((await runCommandAsync(session, 'document.pageNumbers')).ok).toBe(false);
    expect(comparePackages(original, (await session.save()).bytes)).toEqual([]);
  });

  test('the paragraph and the field test', () => {
    const p = pageNumberParagraph();
    expect(hasPageField(p)).toBe(true);
    expect(hasPageField({ TYPE_NAME: 'org_docx4j_wml.CTSimpleField', instr: ' NUMPAGES ' })).toBe(false);
    expect(hasPageField({ TYPE_NAME: 'org_docx4j_wml.CTSimpleField', instr: ' PAGE \\* MERGEFORMAT ' })).toBe(true);
    expect(hasPageField({ name: { localPart: 'instrText' }, value: { value: ' PAGE ' } })).toBe(true);
  });
});

describe('document.pageNumbers in a room (ED-004 section 3.3)', () => {
  test('the footer a peer makes reaches the other with the same part name and relationship id, and both save the same file', async () => {
    const r = await room(await fixture('headers.docx'));
    expect((await runCommandAsync(r.creator, 'document.pageNumbers')).ok).toBe(true);
    await r.creator.partsReady();
    await r.joiner.partsReady();
    const mine = (await r.creator.stories.headers()).filter((h) => h.kind === 'footer');
    const theirs = (await r.joiner.stories.headers()).filter((h) => h.kind === 'footer');
    expect(theirs.map((h) => [h.part.partName.toString(), h.rId])).toEqual(mine.map((h) => [h.part.partName.toString(), h.rId]));
    expect(mine[0]!.rId).toMatch(/^rId[0-9a-z]{8}$/);   // a random id, as a hyperlink's in a room
    const last = r.joiner.stories.sections().length - 1;
    expect((await r.joiner.stories.headerFooterFor(last, 7, false)).footer?.rId).toBe(mine[0]!.rId);
    expect(r.joiner.doc.eq(r.creator.doc)).toBe(true);
    const [a, b] = [await r.creator.save(), await r.joiner.save()];
    expect(comparePackages(a.bytes, b.bytes)).toEqual([]);
    expect(await planPageNumbers(await EditorDocument.open(b.bytes))).toEqual({ extend: [], create: [] });
  });
});
