// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.113, from a constitution ONLYOFFICE had saved: 677 revision elements, 74 of them real. Its Contents refresh
// with tracking on leaves a `w:ins` of no author around every run of every line and around empty runs beside every heading's
// bookmark. The listing lists what a reader can decide on: a revision around an empty run is no change, a tab inside a link is
// in the link, and no author and an empty author are one. "All" still leaves no revision in the file.
import { describe, test, expect } from 'vitest';
import { Node as PMNode } from 'prosemirror-model';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, emptyChangesOf, isEmptyRun, preservedChangesOf, resolveChanges, runCommand, runCommandAsync, stateFor, trackedChangesOf, type Revision, type TrackedChangeInfo } from '../src/index.mjs';
import { fixture, mainXml } from './helpers.mjs';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function documentOf(markup: string): Promise<EditorDocument> {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.getMainDocumentPart().setXml(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${markup}<w:sectPr/></w:body></w:document>`);
  return EditorDocument.open(new Uint8Array(await pkg.save()));
}

const open = async (): Promise<EditorDocument> => EditorDocument.open(await fixture('authorless-contents.docx'));
const listed = (doc: PMNode): [string, string | null, string][] => trackedChangesOf(doc).map((c) => [c.type, c.author, c.text]);

/** Every revision left in a document: text's marks, atoms', paragraph marks'. */
function revisionsIn(doc: PMNode): number {
  let n = 0;
  doc.descendants((node) => {
    if (node.isText) n += node.marks.filter((m) => m.type.name === 'insertion' || m.type.name === 'deletion').length;
    else if (node.attrs['revision'] as Revision | null | undefined) n++;
    if (node.attrs['markRevision'] as Revision | null | undefined) n++;
    return true;
  });
  return n;
}

const ins = (inner: string, author = 'w:author="A"'): string => `<w:ins w:id="${next()}" ${author}>${inner}</w:ins>`;
let id = 1;
const next = (): number => id++;
const run = (s: string): string => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
const TAB = '<w:r><w:tab/></w:r>';

describe('what a Contents refresh leaves, listed as a reader needs it (ED-005 section 12.113)', () => {
  test('a Contents line is its link and its mark; a heading with three wrappers lists nothing', async () => {
    const session = await open();
    expect(listed(session.doc)).toEqual([
      ['Added', null, '2\tMembership generally\t4'],
      ['Added', null, '\r'],
      ['Added', null, '3\tMembership applications\t4'],
      ['Added', null, '\r'],
      ['Added', null, '4\tRegister of members\t5'],
      ['Added', null, '\r'],
      ['Added', 'Reviewer', 'the committee agreed in writing'],
      ['Deleted', 'Reviewer', 'signed'],
      ['Added', 'Reviewer', 'made in writing'],
    ]);
  });

  test('every wrapper is accounted for: joined to a change it touches, or set aside', async () => {
    const session = await open();
    const doc = session.doc;
    let wrappers = 0;
    doc.descendants((node) => { if (isEmptyRun(node) && node.attrs['revision']) wrappers++; return true; });
    const joined = trackedChangesOf(doc).flatMap((c) => c.pieces).filter((p) => p.empty).length;
    expect(wrappers).toBeGreaterThan(0);
    expect(joined).toBeGreaterThan(0);
    expect(joined + emptyChangesOf(doc).length).toBe(wrappers);
    expect(emptyChangesOf(doc).every((c) => c.empty === true && c.text === '')).toBe(true);
  });

  test('a wrapper of another author parts no change', async () => {
    const session = await open();
    const change = trackedChangesOf(session.doc).find((c) => c.text.startsWith('the committee'))!;
    expect(change.author).toBe('Reviewer');
    expect(change.pieces.filter((p) => p.empty)).toEqual([]);
  });

  test('a named wrapper around nothing is no change either: ONLYOFFICE\'s own save, with Claude\'s name on an empty run', async () => {
    const session = await EditorDocument.open(await fixture('numbering-clauses.docx'));
    expect(trackedChangesOf(session.doc)).toEqual([]);
    expect(emptyChangesOf(session.doc).map((c) => c.author)).toEqual(['Claude']);
    await runCommandAsync(session, 'trackedChanges.acceptAll', undefined);
    expect(revisionsIn(session.doc)).toBe(0);
  });

  test('no author and an empty author are one unknown author, a change of either joining the other', async () => {
    const session = await documentOf(`<w:p><w:pPr><w:rPr><w:ins w:id="${next()}" w:author=""/></w:rPr></w:pPr>${ins(run('text'), '')}</w:p>`);
    expect(listed(session.doc)).toEqual([['Added', null, 'text\r']]);
  });

  test('the incremental listing agrees with a fresh walk after an edit', async () => {
    const session = await open();
    const before = stateFor(session);
    trackedChangesOf(before.doc);
    const state = before.apply(before.tr.insertText('x', 2));
    const fresh = PMNode.fromJSON(state.schema, state.doc.toJSON());
    const shape = (c: TrackedChangeInfo): unknown[] => [c.type, c.author, c.text, c.from, c.to, c.pieces.length];
    expect(trackedChangesOf(state.doc).map(shape)).toEqual(trackedChangesOf(fresh).map(shape));
    expect(emptyChangesOf(state.doc).map(shape)).toEqual(emptyChangesOf(fresh).map(shape));
  });
});

describe('an atom inside a link is in the link (ED-005 section 12.113)', () => {
  const link = (inner: string): string => `<w:hyperlink w:anchor="x" w:history="1">${inner}</w:hyperlink>`;

  test('a tab between a link\'s text joins it, and the link\'s edges still part an insertion (core-ts check 30)', async () => {
    const session = await documentOf(`<w:p>${ins(run('before'))}${link(`${ins(run('in'))}${ins(TAB)}${ins(run('link'))}`)}${ins(run('after'))}</w:p>`);
    expect(listed(session.doc).map(([, , text]) => text)).toEqual(['before', 'in\tlink', 'after']);
  });

  test('an atom sharing a run with the link\'s text joins it at the link\'s edge', async () => {
    const session = await documentOf(`<w:p>${link(ins('<w:r><w:t>end</w:t><w:tab/></w:r>'))}</w:p>`);
    expect(listed(session.doc).map(([, , text]) => text)).toEqual(['end\t']);
  });

  test('an atom at a link\'s edge in a run of its own says nothing, and counts as outside', async () => {
    const session = await documentOf(`<w:p>${link(`${ins(run('end'))}${ins(TAB)}`)}</w:p>`);
    expect(listed(session.doc).map(([, , text]) => text)).toEqual(['end', '\t']);
  });

  test('two links with a tab between them stay two', async () => {
    const other = `<w:hyperlink w:anchor="y" w:history="1">${ins(run('second'))}</w:hyperlink>`;
    const session = await documentOf(`<w:p>${link(ins(run('first')))}${ins(TAB)}${other}</w:p>`);
    expect(listed(session.doc).map(([, , text]) => text)).toEqual(['first', '\t', 'second']);
  });
});

describe('resolving, with the wrappers in the file (ED-005 section 12.113)', () => {
  test('Accept all leaves no revision in the document or the saved file, the text kept', async () => {
    const session = await open();
    await runCommandAsync(session, 'trackedChanges.acceptAll', undefined);
    expect(revisionsIn(session.doc)).toBe(0);
    expect(emptyChangesOf(session.doc)).toEqual([]);
    const xml = await mainXml((await session.save()).bytes);
    expect(xml).not.toMatch(/<w:(ins|del)[ />]/);
    expect(xml).toContain('Membership generally');
    expect(xml).toContain('made in writing');
    expect(xml).not.toContain('signed');
  });

  test('Reject all leaves no revision either, and takes what was inserted', async () => {
    const session = await open();
    await runCommandAsync(session, 'trackedChanges.rejectAll', undefined);
    expect(revisionsIn(session.doc)).toBe(0);
    const xml = await mainXml((await session.save()).bytes);
    expect(xml).not.toMatch(/<w:(ins|del)[ />]/);
    expect(xml).not.toContain('the committee agreed');
    expect(xml).toContain('signed');
    // The headings were never inserted: only the wrappers around them go.
    expect(xml.match(/Membership generally/g)).toHaveLength(1);
  });

  test('accepting by the unknown author clears both forms of it and its wrappers, and a named author\'s changes stay', async () => {
    const session = await open();
    await runCommandAsync(session, 'trackedChanges.acceptAll', { author: '' });
    expect(new Set(trackedChangesOf(session.doc).map((c) => c.author))).toEqual(new Set(['Reviewer']));
    expect(trackedChangesOf(session.doc)).toHaveLength(3);
    expect(emptyChangesOf(session.doc)).toEqual([]);
  });

  test('accepting by a named author leaves the unknown author\'s', async () => {
    const session = await open();
    await runCommandAsync(session, 'trackedChanges.acceptAll', { author: 'Reviewer' });
    expect(new Set(trackedChangesOf(session.doc).map((c) => c.author))).toEqual(new Set([null]));
    expect(trackedChangesOf(session.doc)).toHaveLength(6);
  });

  test.each([['no author', ''], ['an empty author', ' w:author=""']])('accepting by the unknown author reaches a revision the engine resolves, with %s', async (_, author) => {
    const table = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblPrChange w:id="${next()}"${author}><w:tblPr/></w:tblPrChange></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>${run('x')}</w:p></w:tc></w:tr></w:tbl><w:p/>`;
    const session = await documentOf(table);
    const left = (): number => preservedChangesOf(session.doc, (id) => session.ctx.ids.get(id)?.value).length;
    expect(left()).toBe(1);
    await runCommandAsync(session, 'trackedChanges.acceptAll', { author: '' });
    expect(left()).toBe(0);
  });

  test('accepting one change takes the wrappers in it, and leaves the others', async () => {
    const session = await open();
    const line = trackedChangesOf(session.doc)[0]!;
    const inside = line.pieces.filter((p) => p.empty).length;
    expect(inside).toBeGreaterThan(0);
    const aside = emptyChangesOf(session.doc).length;
    expect(runCommand(session, 'trackedChange.accept', { from: line.from + 1, to: line.from + 2 }).ok).toBe(true);
    expect(trackedChangesOf(session.doc).map((c) => c.text)).not.toContain(line.text);
    expect(emptyChangesOf(session.doc)).toHaveLength(aside);
  });

  test('the number resolved counts changes, not wrappers', async () => {
    const session = await open();
    const state = stateFor(session);
    const changes = trackedChangesOf(state.doc);
    const count = resolveChanges(state.tr, [...changes, ...emptyChangesOf(state.doc)], 'accept', { element: () => undefined, listIndent: () => null });
    expect(count).toBe(changes.length);
  });

  test('an unedited file keeps its wrappers: nothing was asked of them', async () => {
    const session = await open();
    const xml = await mainXml((await session.save()).bytes);
    expect(xml.match(/<w:ins /g)!.length).toBe(session.diagnostics.tracked.shown - 1);
  });
});
