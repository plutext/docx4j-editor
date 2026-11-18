// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 6.6 items 1 and 2 (step D1): notes and headers bound to fragments and edited
// headless. The untouched round trip holds after every story is bound; an edit writes the
// story's part and nothing else beyond app.xml; a note whose reference the body lost goes at
// save, the orphans stay; undo restores the baseline; the self-check names a story.
import { describe, test, expect } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { EditorDocument, runCommand, type Story } from '../src/index.mjs';
import { fixture, comparePackages, allFixtures, textOf, plain } from './helpers.mjs';

/** The position just inside the first paragraph's text of a story's current document, after its own note mark. */
function startOfText(doc: PMNode): number {
  let pos = 1;
  doc.firstChild!.forEach((child, offset) => { if (pos === 1 && child.isText) pos = 1 + offset; });
  return pos;
}

async function reloadNotes(bytes: Uint8Array, kind: 'footnote' | 'endnote'): Promise<wml.CTFtnEdn[]> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  const main = pkg.getMainDocumentPart();
  const part = kind === 'footnote' ? main.footnotesPart! : main.endnotesPart!;
  const tree = await part.readContents();
  return ((kind === 'footnote' ? (tree as wml.CTFootnotes).footnote : (tree as wml.CTEndnotes).endnote) ?? []) as wml.CTFtnEdn[];
}

const withoutParents = (notes: wml.CTFtnEdn[]) => notes.map((n) => plain(n));

/** A note's text in saved bytes, through the editor's own projection. */
async function noteTextIn(bytes: Uint8Array, kind: 'footnote' | 'endnote', id: number): Promise<string> {
  const session = await EditorDocument.open(bytes);
  const notes = kind === 'footnote' ? await session.stories.footnotes() : await session.stories.endnotes();
  const note = notes.find((n) => n.id === id);
  return note ? textOf(note.doc.firstChild!) : '';
}

describe('binding stories', () => {
  test('every note and header bound: the maps fill, no part is written, the round trip holds', async () => {
    for (const name of await allFixtures()) {
      const original = await fixture(name);
      const session = await EditorDocument.open(original);
      const stories: Story[] = [...await session.stories.footnotes(), ...await session.stories.endnotes(), ...await session.stories.headers()];
      const before = session.undoManager.undoStack.length;
      for (const s of stories) {
        const fragment = session.stories.fragmentOf(s);
        expect(session.stories.fragmentOf(s)).toBe(fragment);
        expect(session.stories.hasFragment(s)).toBe(true);
        expect(session.stories.currentDoc(s).eq(s.doc)).toBe(true);
      }
      expect(session.undoManager.undoStack.length).toBe(before);
      expect(session.undoManager.canUndo()).toBe(false);
      const saved = await session.save();
      expect(saved.changed, name).toBe(false);
      expect(comparePackages(original, saved.bytes), name).toEqual([]);
    }
  });
});

describe('a footnote edited', () => {
  test('the footnotes part is written and nothing else; the note reloads edited, the others as they were; undo restores the baseline', async () => {
    const original = await fixture('footnote.docx');
    const session = await EditorDocument.open(original);
    const notes = await session.stories.footnotes();
    const note = notes.find((n) => !n.orphan)!;
    const others = withoutParents((await reloadNotes(original, 'footnote')).filter((n) => n.id !== note.id));
    const pos = startOfText(note.doc);
    const result = runCommand(session, 'insertText', { text: 'Edited: ' }, { from: pos }, note);
    expect(result.ok).toBe(true);
    expect(textOf(session.stories.currentDoc(note).firstChild!)).toMatch(/^Edited: /);
    expect(session.undoManager.canUndo()).toBe(true);
    expect(session.stories.orderedNotes('footnote', session.doc).map((n) => n.id)).toEqual(notes.map((n) => n.id));

    const saved = await session.save();
    expect(saved.changed).toBe(true);
    expect(saved.selfCheck?.equal).toBe(true);
    expect(comparePackages(original, saved.bytes, ['word/footnotes.xml', 'docProps/app.xml'])).toEqual([]);
    expect([...session.touchedParts]).toEqual(['/word/footnotes.xml']);
    const reloaded = await reloadNotes(saved.bytes, 'footnote');
    const mine = reloaded.find((n) => n.id === note.id)!;
    expect(await noteTextIn(saved.bytes, 'footnote', note.id)).toMatch(/^Edited: /);
    expect(withoutParents(reloaded.filter((n) => n.id !== note.id))).toEqual(others);
    // The note's own mark and its style survive the rebuild.
    const p = mine.content![0]!.value as wml.P;
    expect(p.pPr?.pStyle?.val).toBe((note.element.content![0]!.value as wml.P).pPr?.pStyle?.val);
    const reopened = await EditorDocument.open(saved.bytes);
    const again = (await reopened.stories.footnotes()).find((n) => n.id === note.id)!;
    expect(textOf(again.doc.firstChild!)).toBe(textOf(session.stories.currentDoc(note).firstChild!));

    // Undo takes the edit back; a fresh session over the original with the same undo has nothing to write.
    const fresh = await EditorDocument.open(original);
    const n2 = (await fresh.stories.footnotes()).find((n) => n.id === note.id)!;
    runCommand(fresh, 'insertText', { text: 'Edited: ' }, { from: startOfText(n2.doc) }, n2);
    fresh.undoManager.undo();
    expect(fresh.stories.currentDoc(n2).eq(n2.doc)).toBe(true);
    const unsaved = await fresh.save();
    expect(unsaved.changed).toBe(false);
    expect(comparePackages(original, unsaved.bytes)).toEqual([]);
  });

  test('an orphan is editable too, and stays', async () => {
    const original = await fixture('footnote.docx');
    const session = await EditorDocument.open(original);
    const orphan = (await session.stories.footnotes()).find((n) => n.orphan)!;
    runCommand(session, 'insertText', { text: 'Orphan text' }, { from: startOfText(orphan.doc) }, orphan);
    const saved = await session.save();
    expect(comparePackages(original, saved.bytes, ['word/footnotes.xml', 'docProps/app.xml'])).toEqual([]);
    expect(await noteTextIn(saved.bytes, 'footnote', orphan.id)).toContain('Orphan text');
  });
});

describe('an endnote edited', () => {
  test('notes-numbering.docx: the endnotes part alone', async () => {
    const original = await fixture('notes-numbering.docx');
    const session = await EditorDocument.open(original);
    const notes = await session.stories.endnotes();
    expect(notes.map((n) => n.id)).toEqual([1, 2]);
    expect(notes.every((n) => !n.orphan)).toBe(true);
    const note = notes[1]!;
    runCommand(session, 'insertText', { text: ' (edited)' }, { from: session.stories.currentDoc(note).firstChild!.nodeSize - 1 }, note);
    const saved = await session.save();
    expect(comparePackages(original, saved.bytes, ['word/endnotes.xml', 'docProps/app.xml'])).toEqual([]);
    expect(await noteTextIn(saved.bytes, 'endnote', 2)).toBe(' Endnote 2. (edited)');
    expect(await noteTextIn(saved.bytes, 'endnote', 1)).toBe(' Endnote 1.');
    expect((await session.stories.footnotes()).length).toBe(5);
  });
});

describe('a reference deleted', () => {
  test('the note goes at save, the others and the orphans stay; undo brings it back', async () => {
    const original = await fixture('footnote.docx');
    const session = await EditorDocument.open(original);
    const notes = await session.stories.footnotes();
    const target = notes.find((n) => !n.orphan)!;
    const orphanIds = notes.filter((n) => n.orphan).map((n) => n.id);
    // Delete the reference atom from the body.
    let at = -1;
    session.doc.descendants((node, pos) => { if (at < 0 && node.type.name === 'footnoteRef' && Number(node.attrs['footnoteId']) === target.id) at = pos; return at < 0; });
    expect(at).toBeGreaterThanOrEqual(0);
    const state = EditorState.create({ doc: session.doc, selection: TextSelection.create(session.doc, at, at + 1) });
    session.undoManager.stopCapturing();
    session.replace(state.apply(state.tr.deleteSelection()).doc);
    expect(session.stories.referenceOf(session.doc, 'footnote', target.id)).toBeUndefined();
    const listed = session.stories.orderedNotes('footnote', session.doc).map((n) => n.id);
    expect(listed).not.toContain(target.id);
    for (const id of orphanIds) expect(listed).toContain(id);

    const saved = await session.save();
    expect(saved.selfCheck?.equal).toBe(true);
    const mainName = session.main.partName.storeName;
    expect(comparePackages(original, saved.bytes, [mainName, 'word/footnotes.xml', 'docProps/app.xml'])).toEqual([]);
    const reloaded = await reloadNotes(saved.bytes, 'footnote');
    expect(reloaded.find((n) => n.id === target.id)).toBeUndefined();
    const before = await reloadNotes(original, 'footnote');
    expect(reloaded.filter((n) => n.id <= 0).length).toBe(before.filter((n) => n.id <= 0).length);
    for (const id of orphanIds) expect(reloaded.find((n) => n.id === id)).toBeDefined();
    expect(reloaded.length).toBe(before.length - 1);

    // Undo in a fresh session: the reference is back and the part is not written.
    const fresh = await EditorDocument.open(original);
    await fresh.stories.footnotes();
    const st = EditorState.create({ doc: fresh.doc, selection: TextSelection.create(fresh.doc, at, at + 1) });
    fresh.undoManager.stopCapturing();
    fresh.replace(st.apply(st.tr.deleteSelection()).doc);
    fresh.undoManager.undo();
    expect(fresh.stories.referenceOf(fresh.doc, 'footnote', target.id)).toBeDefined();
    const unsaved = await fresh.save();
    expect(comparePackages(original, unsaved.bytes)).toEqual([]);
  });
});

describe('a header and a footer edited', () => {
  test('headers.docx: the header part alone is written and reloads edited; the other headers keep their bytes', async () => {
    const original = await fixture('headers.docx');
    const session = await EditorDocument.open(original);
    const headers = await session.stories.headers();
    const header = headers.find((h) => h.kind === 'header' && textOf(h.doc.firstChild!).length > 0)!;
    runCommand(session, 'insertText', { text: 'Edited header: ' }, { from: 1 }, header);
    const saved = await session.save();
    expect(saved.changed).toBe(true);
    expect(saved.selfCheck?.equal).toBe(true);
    const storeName = header.part.partName.storeName;
    expect(comparePackages(original, saved.bytes, [storeName, 'docProps/app.xml'])).toEqual([]);
    const reopened = await EditorDocument.open(saved.bytes);
    const again = (await reopened.stories.headers()).find((h) => h.part.partName.storeName === storeName)!;
    expect(textOf(again.doc.firstChild!)).toMatch(/^Edited header: /);
    expect(again.rId).toBe(header.rId);
  });

  test('odd-even-first-page.docx: a footer', async () => {
    const original = await fixture('odd-even-first-page.docx');
    const session = await EditorDocument.open(original);
    const footer = (await session.stories.headers()).find((h) => h.kind === 'footer')!;
    runCommand(session, 'insertText', { text: 'Edited footer: ' }, { from: 1 }, footer);
    const saved = await session.save();
    expect(comparePackages(original, saved.bytes, [footer.part.partName.storeName, 'docProps/app.xml'])).toEqual([]);
    const reopened = await EditorDocument.open(saved.bytes);
    const again = (await reopened.stories.headers()).find((h) => h.part.partName.storeName === footer.part.partName.storeName)!;
    expect(textOf(again.doc.firstChild!)).toMatch(/^Edited footer: /);
  });

  test('header-image.docx: a header with an image through its own relationships edits and keeps the image', async () => {
    const original = await fixture('header-image.docx');
    const session = await EditorDocument.open(original);
    const withImage = (await session.stories.headers()).find((h) => { let img = false; h.doc.descendants((n) => { if (n.type.name === 'image') img = true; return !img; }); return img; })!;
    expect(withImage).toBeDefined();
    expect(withImage.part.relationshipsPart).toBeDefined();
    runCommand(session, 'insertText', { text: 'Edited: ' }, { from: 1 }, withImage);
    const saved = await session.save();
    // The fixture has no app.xml: the write adds it.
    expect(comparePackages(original, saved.bytes, [withImage.part.partName.storeName, 'docProps/app.xml']).filter((d) => d.name !== 'docProps/app.xml')).toEqual([]);
    const reopened = await EditorDocument.open(saved.bytes);
    const again = (await reopened.stories.headers()).find((h) => h.part.partName.storeName === withImage.part.partName.storeName)!;
    let images = 0;
    again.doc.descendants((n) => { if (n.type.name === 'image') images++; return true; });
    expect(images).toBe(1);
    expect(textOf(again.doc.firstChild!)).toMatch(/^Edited: /);
  });
});

describe('the self-check over stories', () => {
  test('names the story whose fragment the saved bytes do not re-project to', async () => {
    const original = await fixture('footnote.docx');
    const session = await EditorDocument.open(original);
    const note = (await session.stories.footnotes()).find((n) => !n.orphan)!;
    runCommand(session, 'insertText', { text: 'Edited: ' }, { from: startOfText(note.doc) }, note);
    // The original bytes, against a session whose note fragment differs: the check must say so.
    const result = await session.selfCheck(original, session.doc);
    expect(result.equal).toBe(false);
    expect(result.difference?.story).toBe(note.key);
  });
});
