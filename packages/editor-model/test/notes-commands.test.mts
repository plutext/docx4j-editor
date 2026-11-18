// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 6.6 item 3 (step D3): footnotes and endnotes inserted and deleted through the
// runner, the part and the styles made when the document lacks them, the reloaded document
// carrying the note; a deleted reference removing its note at save; undo of an insert writing
// nothing; the numbering display over `w:footnotePr` as a pure function.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { EditorDocument, runCommandAsync, runCommand, noteLabels, formatNumber } from '../src/index.mjs';
import { fixture, comparePackages, textOf } from './helpers.mjs';

/** The first `footnoteRef` atom of a kind, with its position. */
function firstReference(session: EditorDocument, kind: 'footnote' | 'endnote'): { pos: number; id: number; sourceId: string; runSource: string } | undefined {
  let found: { pos: number; id: number; sourceId: string; runSource: string } | undefined;
  session.doc.descendants((node, pos) => {
    if (found) return false;
    if (node.type.name === 'footnoteRef' && node.attrs['kind'] === kind) found = { pos, id: Number(node.attrs['footnoteId']), sourceId: String(node.attrs['sourceId']), runSource: String(node.attrs['runSource']) };
    return !found;
  });
  return found;
}

describe('range.insertFootnote', () => {
  test('in a document with no notes part and no note styles: the part with its separators, the styles, the reference run, the note; the reload carries it', async () => {
    const original = await fixture('HelloWordOnline.docx');
    const session = await EditorDocument.open(original);
    expect(session.main.footnotesPart).toBeUndefined();
    expect((await session.stories.footnotes()).length).toBe(0);
    // Word Online's styles part already carries the note styles; a document without them gets them spliced in (the next test).
    const hadStyles = (await session.main.styleDefinitionsPart!.readContents()).style!.some((s) => s.styleId === 'FootnoteText');
    // The caret after "Hello" in the first paragraph.
    const result = await runCommandAsync(session, 'range.insertFootnote', { text: 'A note made here.' }, { from: 6 });
    expect(result.ok).toBe(true);
    const ref = firstReference(session, 'footnote')!;
    expect(ref).toBeDefined();
    expect(ref.id).toBe(1);
    expect(ref.pos).toBe(6);
    // The reference run copies the caret run's properties and takes the reference style.
    const run = session.ctx.ids.get(ref.runSource)!.value as wml.R;
    expect(run.rPr?.rStyle?.val).toBe('FootnoteReference');
    expect((session.ctx.ids.get(ref.sourceId)!.value as wml.CTFtnEdnRef).id).toBe(1);
    // The part exists and is marked; the styles too.
    expect(session.main.footnotesPart).toBeDefined();
    expect([...session.touchedParts].sort()).toEqual(hadStyles ? ['/word/footnotes.xml'] : ['/word/footnotes.xml', '/word/styles.xml']);
    const styles = await session.main.styleDefinitionsPart!.readContents();
    for (const id of ['FootnoteText', 'FootnoteReference', 'FootnoteTextChar']) expect(styles.style!.some((s) => s.styleId === id), id).toBe(true);
    expect(session.pkg.propertyResolver.getEffectiveRPr('FootnoteReference')?.vertAlign?.val).toBe('superscript');
    // The note is in the map and listed with the body's order.
    const note = session.stories.noteOf('footnote', 1)!;
    expect(note.isNew).toBe(true);
    expect(session.stories.hasFragment(note)).toBe(true);
    expect(textOf(session.stories.currentDoc(note).firstChild!)).toBe(' A note made here.');
    expect(session.stories.orderedNotes('footnote', session.doc).map((n) => n.id)).toEqual([1]);
    expect(session.undoManager.canUndo()).toBe(true);

    const saved = await session.save();
    expect(saved.selfCheck?.equal).toBe(true);
    const mainName = session.main.partName.storeName;
    const diffs = comparePackages(original, saved.bytes, [mainName, 'word/styles.xml', 'docProps/app.xml']).filter((d) => d.name !== 'docProps/app.xml');
    expect(diffs).toEqual([{ name: 'word/footnotes.xml', reason: 'added' }]);
    if (hadStyles) expect(comparePackages(original, saved.bytes, [mainName, 'word/footnotes.xml', 'docProps/app.xml']).filter((d) => d.name === 'word/styles.xml')).toEqual([]);
    const pkg = await WordprocessingMLPackage.load(saved.bytes);
    const part = pkg.getMainDocumentPart().footnotesPart!;
    const tree = await part.readContents();
    expect(tree.footnote!.map((n) => [n.id, n.type])).toEqual([[-1, 'separator'], [0, 'continuationSeparator'], [1, undefined]]);
    const xml = await part.getXml();
    expect(xml).toContain('<w:footnoteRef/>');
    expect(xml).toContain('FootnoteText');
    expect(await pkg.getMainDocumentPart().getXml()).toContain('<w:footnoteReference w:id="1"/>');
    // Word needs the content type registered for the new part.
    const types = new TextDecoder().decode((await import('@docx4j/core-ts')).ZipPartStore.prototype.loadSync.call(new (await import('@docx4j/core-ts')).ZipPartStore(saved.bytes), '[Content_Types].xml'));
    expect(types).toContain('footnotes+xml');
    // Reopened: the note is there, numbered 1, with its text; the reference resolves.
    const reopened = await EditorDocument.open(saved.bytes);
    const notes = await reopened.stories.footnotes();
    expect(notes.map((n) => [n.id, n.orphan])).toEqual([[1, false]]);
    expect(textOf(notes[0]!.doc.firstChild!)).toBe(' A note made here.');
    expect(reopened.stories.referenceOf(reopened.doc, 'footnote', 1)).toBeDefined();
    // A second save writes nothing new: the note is in the part now.
    expect(note.isNew).toBe(false);
  });

  test('the note styles are spliced from docx4j\'s defaults into a styles part that lacks them, with the linked character style', async () => {
    const original = await fixture('headers-footers.docx');
    const session = await EditorDocument.open(original);
    const styles = await session.main.styleDefinitionsPart!.readContents();
    const has = (id: string): boolean => (styles.style ?? []).some((s) => s.styleId === id);
    if (has('EndnoteText') && has('EndnoteReference')) return;   // nothing to prove on this fixture
    const added = await session.ensureStyles(['EndnoteText', 'EndnoteReference']);
    expect(added).toBe(true);
    const after = await session.main.styleDefinitionsPart!.readContents();
    for (const id of ['EndnoteText', 'EndnoteReference', 'EndnoteTextChar']) expect((after.style ?? []).some((s) => s.styleId === id), id).toBe(true);
    expect([...session.touchedParts]).toEqual(['/word/styles.xml']);
    expect(session.pkg.propertyResolver.getEffectiveRPr('EndnoteReference')?.vertAlign?.val).toBe('superscript');
    const saved = await session.save();
    expect(comparePackages(original, saved.bytes, ['word/styles.xml', 'docProps/app.xml'])).toEqual([]);
    const reopened = await EditorDocument.open(saved.bytes);
    expect(reopened.pkg.propertyResolver.getEffectiveRPr('EndnoteReference')?.vertAlign?.val).toBe('superscript');
  });

  test('an endnote the same way, alongside footnotes; a second note takes the next id; undo of an insert writes nothing', async () => {
    const original = await fixture('footnote.docx');
    const session = await EditorDocument.open(original);
    const before = (await session.stories.footnotes()).map((n) => n.id);
    const maxId = Math.max(...before);
    const r1 = await runCommandAsync(session, 'range.insertFootnote', { text: 'Second-generation note' }, { from: 3 });
    expect(r1.ok).toBe(true);
    const mine = firstReference(session, 'footnote')!;
    expect(mine.id).toBe(maxId + 1);
    const r2 = await runCommandAsync(session, 'range.insertEndnote', undefined, { from: 5 });
    expect(r2.ok).toBe(true);
    const endnotes = await session.stories.endnotes();
    const en = firstReference(session, 'endnote')!;
    expect(en.id).toBe(Math.max(...endnotes.filter((n) => !n.isNew).map((n) => n.id)) + 1);
    expect(textOf(session.stories.currentDoc(session.stories.noteOf('endnote', en.id)!).firstChild!)).toBe(' ');
    // No part was touched yet: the parts and the styles exist; the writes come at save.
    expect([...session.touchedParts]).toEqual([]);
    const saved = await session.save();
    expect(saved.selfCheck?.equal).toBe(true);
    expect(comparePackages(original, saved.bytes, [session.main.partName.storeName, 'word/footnotes.xml', 'word/endnotes.xml', 'docProps/app.xml'])).toEqual([]);
    const reopened = await EditorDocument.open(saved.bytes);
    expect((await reopened.stories.footnotes()).some((n) => n.id === mine.id && !n.orphan)).toBe(true);
    expect((await reopened.stories.endnotes()).some((n) => n.id === en.id && !n.orphan)).toBe(true);
    expect(reopened.stories.orderedNotes('footnote', reopened.doc)[0]!.id).toBe(mine.id);   // first in reference order

    // Undo in a fresh session: the reference goes, the note stays in the map unwritten.
    const fresh = await EditorDocument.open(original);
    await fresh.stories.footnotes();
    await runCommandAsync(fresh, 'range.insertFootnote', { text: 'gone' }, { from: 3 });
    fresh.undoManager.undo();
    expect(firstReference(fresh, 'footnote')!.id).not.toBe(maxId + 1);
    const unsaved = await fresh.save();
    expect(unsaved.changed).toBe(false);
    expect(comparePackages(original, unsaved.bytes)).toEqual([]);
    // Redo brings the reference and the note's content back.
    fresh.undoManager.redo();
    expect(firstReference(fresh, 'footnote')!.id).toBe(maxId + 1);
    expect(fresh.stories.orderedNotes('footnote', fresh.doc)[0]!.id).toBe(maxId + 1);
  });

  test('refused outside a body paragraph: in a note, and with a selection across paragraphs', async () => {
    const session = await EditorDocument.open(await fixture('footnote.docx'));
    const note = (await session.stories.footnotes())[0]!;
    const inNote = await runCommandAsync(session, 'range.insertFootnote', undefined, { from: 2 }, note);
    expect(inNote.ok).toBe(false);
    const across = await runCommandAsync(session, 'range.insertFootnote', undefined, { from: 1, to: session.doc.child(0).nodeSize + 2 });
    expect(across.ok).toBe(false);
  });
});

describe('note.delete', () => {
  test('by id from the panel, and at the caret; the note goes at save', async () => {
    const original = await fixture('footnote.docx');
    const session = await EditorDocument.open(original);
    const notes = await session.stories.footnotes();
    const first = notes.find((n) => !n.orphan)!;
    const r = runCommand(session, 'note.delete', { kind: 'footnote', id: first.id });
    expect(r.ok).toBe(true);
    expect(session.stories.referenceOf(session.doc, 'footnote', first.id)).toBeUndefined();
    // At the caret: the reference just before it.
    const next = firstReference(session, 'footnote')!;
    const r2 = runCommand(session, 'note.delete', undefined, { from: next.pos + 1 });
    expect(r2.ok).toBe(true);
    expect(session.stories.referenceOf(session.doc, 'footnote', next.id)).toBeUndefined();
    // Nothing to delete at a caret with no reference before it.
    expect(runCommand(session, 'note.delete', undefined, { from: 1 }).ok).toBe(false);
    const saved = await session.save();
    const reopened = await EditorDocument.open(saved.bytes);
    const ids = (await reopened.stories.footnotes()).map((n) => n.id);
    expect(ids).not.toContain(first.id);
    expect(ids).not.toContain(next.id);
    expect(ids).toEqual(notes.map((n) => n.id).filter((id) => id !== first.id && id !== next.id));
  });
});

describe('the numbering display (w:footnotePr, w:endnotePr)', () => {
  test('formats', () => {
    expect([1, 4, 9, 14, 40, 99].map((n) => formatNumber(n, 'lowerRoman'))).toEqual(['i', 'iv', 'ix', 'xiv', 'xl', 'xcix']);
    expect([1, 26, 27, 53].map((n) => formatNumber(n, 'upperLetter'))).toEqual(['A', 'Z', 'AA', 'AAA']);
    expect([1, 2, 3, 4, 5].map((n) => formatNumber(n, 'chicago'))).toEqual(['*', '†', '‡', '§', '**']);
    expect(formatNumber(7, undefined)).toBe('7');
    expect(formatNumber(7, 'decimalZero')).toBe('07');
  });

  test('notes-numbering.docx: lower roman restarting each section; endnotes upper letter from 3, continuous', async () => {
    const session = await EditorDocument.open(await fixture('notes-numbering.docx'));
    const labels = noteLabels(session.doc, session.stories.sections(), session.noteSettings);
    expect([1, 2, 3, 4, 5].map((id) => labels.get(`footnote:${id}`))).toEqual(['i', 'ii', 'i', 'ii', 'iii']);
    expect([1, 2].map((id) => labels.get(`endnote:${id}`))).toEqual(['C', 'D']);
  });

  test('footnote.docx: plain decimal in reference order; a document without settings counts from 1', async () => {
    const session = await EditorDocument.open(await fixture('footnote.docx'));
    const labels = noteLabels(session.doc, session.stories.sections(), session.noteSettings);
    const order = session.stories.orderedNotes('footnote', session.doc);
    await session.stories.footnotes();
    const ordered = session.stories.orderedNotes('footnote', session.doc).filter((n) => !n.orphan);
    expect(ordered.map((n) => labels.get(n.key))).toEqual(ordered.map((_, i) => String(i + 1)));
    void order;
  });
});

test("Word's default note formats: footnotes 1, 2, 3 and endnotes i, ii, iii, where the document states none (ED-006 section 9.6)", async () => {
  const { noteNumberingFor } = await import('../src/index.mjs');
  expect(noteNumberingFor('footnote', undefined, undefined).fmt).toBeUndefined();
  expect(noteNumberingFor('endnote', undefined, undefined).fmt).toBe('lowerRoman');
  expect(noteNumberingFor('endnote', undefined, { endnotePr: { numFmt: { val: 'decimal' } } }).fmt).toBe('decimal');
});
