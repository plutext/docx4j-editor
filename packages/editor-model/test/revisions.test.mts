// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 5.10, step C1: tracked changes in the edit layer, over the revisions Word 2010
// and Word 15 made on the reference machine (fixtures/revisions/, section 13). Every kind of
// section 5.2 imported as content: insertion and deletion marks with their authors, another
// author's insertion deleted as both, a move's halves with its name, a run's formatting change
// with its old properties, a paragraph mark's and a row's revision and a paragraph's properties
// change as attributes, an atom's run inserted or deleted as its `revision`; the revisions E4
// leaves as they are counted; a note's revisions projected as E1 did. Unedited, the file saves
// byte for byte. Edited, a paragraph or a row keeps its revisions (section 12.3). Moves checked
// against the engine (section 5.6). A room or a kept session made before refused (section 5.9).
import { describe, test, expect } from 'vitest';
import * as Y from 'yjs';
import { WordprocessingMLPackage, runItemsOf, type Element } from '@docx4j/core-ts';
import type { Node as PMNode, Mark } from 'prosemirror-model';
import { EditorDocument, runCommand, findIn, newEditStore, RoomVersionError, ProjectionVersionError, PROJECTION, PROJECTION_KEY, ROOM_SCHEMA, type Revision, type ParagraphAttrs } from '../src/index.mjs';
import { fixture, comparePackages, wordDate } from './helpers.mjs';

const WORD = ['revisions/revisions-word2010.docx', 'revisions/revisions-word15.docx'];

/** The top-level block whose text starts with `prefix`, with its position. */
function block(doc: PMNode, prefix: string): { node: PMNode; pos: number; index: number } {
  let found: { node: PMNode; pos: number; index: number } | undefined;
  doc.forEach((node, pos, index) => { if (!found && node.textContent.startsWith(prefix)) found = { node, pos, index }; });
  if (!found) throw new Error(`no block starting "${prefix}"`);
  return found;
}

/** The text nodes of a node with their revision marks, as `text|kind:author[/move]`. */
function marked(node: PMNode): string[] {
  const out: string[] = [];
  node.descendants((n) => {
    if (!n.isText) return true;
    const m = n.marks.filter((x) => ['insertion', 'deletion', 'formatChange'].includes(x.type.name));
    if (m.length > 0) out.push(`${n.text}|${m.map((x: Mark) => `${x.type.name}:${String(x.attrs['author'])}${x.attrs['move'] !== null && x.attrs['move'] !== undefined ? `/${String(x.attrs['move'])}` : ''}`).join(',')}`);
    return false;
  });
  return out;
}

const mark = (node: PMNode, type: string): Mark | undefined => {
  let out: Mark | undefined;
  node.descendants((n) => { if (!out) out = n.marks.find((m) => m.type.name === type); return !out; });
  return out;
};

describe.each(WORD)('%s: every kind of section 5.2 imported as content', (name) => {
  test('text inserted and deleted by two authors; an insertion deleted by another author; dates to the minute', async () => {
    const session = await EditorDocument.open(await fixture(name));
    const doc = session.doc;
    expect(marked(block(doc, 'The quick brown').node)).toEqual([
      'red |insertion:Author B', 'very |insertion:Author A', 'old |deletion:Author A', ' at da|deletion:Author B', 'wn|deletion:Author B',
    ]);
    expect(marked(block(doc, 'Cows graze').node)).toEqual(['green |insertion:Author A,deletion:Author B']);
    const date = mark(block(doc, 'The quick brown').node, 'insertion')!.attrs['date'] as string;
    // The file's w:date, Word's wall clock, is read as local time (core-ts 0.3.0): shown again as Word's, it is the file's.
    expect(wordDate(date)).toMatch(/^2026-09-28T\d\d:\d\d:00Z$/);
  });

  test("a move: both halves carry the move's name, text and paragraph mark", async () => {
    const session = await EditorDocument.open(await fixture(name));
    const doc = session.doc;
    const halves: { index: number; marks: string[]; mark: Revision | null }[] = [];
    doc.forEach((node, _pos, index) => {
      if (node.textContent === 'The second paragraph of the move: it goes before the first.') halves.push({ index, marks: marked(node), mark: (node.attrs as ParagraphAttrs).markRevision });
    });
    expect(halves).toHaveLength(2);
    const [to, from] = halves as [typeof halves[number], typeof halves[number]];
    expect(to.index).toBeLessThan(from.index);
    const moveName = /\/(move\d+)$/.exec(to.marks[0]!)![1]!;
    expect(to.marks).toEqual([`The second paragraph of the move: it goes before the first.|insertion:Author A/${moveName}`]);
    expect(from.marks).toEqual([`The second paragraph of the move: it goes before the first.|deletion:Author A/${moveName}`]);
    expect(to.mark?.insertion?.move).toBe(moveName);
    expect(from.mark?.deletion?.move).toBe(moveName);
    // The range markers stay opaque atoms, whatever is in them.
    const kinds: string[] = [];
    doc.descendants((n) => { if (n.type.name === 'opaqueInline' || n.type.name === 'opaqueBlock') kinds.push(String(n.attrs['kind'])); return true; });
    for (const k of ['moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd']) expect(kinds).toContain(k);
  });

  test("a run's formatting changed: the mark holds the old properties", async () => {
    const session = await EditorDocument.open(await fixture(name));
    const node = block(session.doc, 'The word alpha').node;
    // Word 15 took the spaces around "bravo" with it: the words, as the person selected them.
    expect(marked(node).map((m) => m.replace(/^\s+|\s+(?=\|)/g, ''))).toEqual(['alpha|formatChange:Author A', 'bravo|formatChange:Author A', 'charlie|formatChange:Author A']);
    const before = JSON.parse(mark(node, 'formatChange')!.attrs['before'] as string) as Record<string, unknown>;
    expect(Object.keys(before).filter((k) => k !== 'TYPE_NAME')).toEqual([]);   // alpha was plain before
  });

  test("paragraph properties changed, and a list's numbering, as pPrChange with the old properties", async () => {
    const session = await EditorDocument.open(await fixture(name));
    const doc = session.doc;
    for (const text of ['Centre this paragraph.', 'Indent this paragraph by one step.', 'Give this paragraph the style Heading 2.']) {
      const attrs = block(doc, text).node.attrs as ParagraphAttrs;
      expect(attrs.pPrChange?.author, text).toBe('Author A');
    }
    const item = block(doc, 'The first numbered item.').node.attrs as ParagraphAttrs;
    expect(item.pPrChange?.author).toBe('Author A');
    expect(JSON.stringify(item.pPrChange?.before)).toContain('numId');
    expect(item.numPr?.numId).not.toBe((item.pPrChange!.before as { numPr: { numId: { val: number } } }).numPr.numId.val.toString());
  });

  test('a paragraph mark inserted, and one deleted', async () => {
    const session = await EditorDocument.open(await fixture(name));
    const doc = session.doc;
    expect((block(doc, 'Split this p').node.attrs as ParagraphAttrs).markRevision).toEqual({ insertion: expect.objectContaining({ author: 'Author A', move: null }) });
    expect((block(doc, 'Join this paragraph').node.attrs as ParagraphAttrs).markRevision).toEqual({ deletion: expect.objectContaining({ author: 'Author A', move: null }) });
    // The mark's own w:rPr, less the revision: nothing else here.
    expect((block(doc, 'Join this paragraph').node.attrs as ParagraphAttrs).markRPr).toBeNull();
  });

  test('a row inserted and one deleted, and their cells marked', async () => {
    const session = await EditorDocument.open(await fixture(name));
    const table = block(session.doc, 'RowActionNote').node;
    const rows: { text: string; revision: Revision | null; marks: string[] }[] = [];
    table.forEach((row) => rows.push({ text: row.textContent, revision: row.attrs['revision'] as Revision | null, marks: marked(row) }));
    expect(rows.map((r) => r.revision === null ? '-' : r.revision.insertion ? 'ins' : 'del')).toEqual(['-', '-', '-', 'ins', 'del']);
    // Word 15 capitalised what was typed in the new row's cells.
    expect(rows[3]!.marks.map((m) => m.toLowerCase())).toEqual(['new|insertion:author a', 'new|insertion:author a', 'new|insertion:author a']);
    expect(rows[4]!.marks.every((m) => m.includes('deletion:Author A'))).toBe(true);
    // The row's trPr attribute is the properties less the revision.
    expect(JSON.stringify(table.child(3).attrs['trPr'] ?? {})).not.toContain('"ins"');
  });

  test("an atom's run inserted or deleted: a deleted tab, an inserted break and note reference", async () => {
    const session = await EditorDocument.open(await fixture(name));
    const node = block(session.doc, 'Left').node;
    const atoms: string[] = [];
    node.forEach((n) => { const r = n.attrs['revision'] as Revision | null | undefined; if (r) atoms.push(`${n.type.name}:${r.insertion ? 'ins' : ''}${r.deletion ? 'del' : ''}`); });
    expect(atoms).toEqual(['tab:del', 'hardBreak:ins', 'footnoteRef:ins']);
  });

  test('the revisions E4 leaves as they are are counted; a note keeps E1\'s projection of its revisions', async () => {
    const session = await EditorDocument.open(await fixture(name));
    expect(session.diagnostics.tracked.shown).toBe(38);
    expect(session.diagnostics.tracked.preserved).toEqual({ sectPrChange: 1, tblGridChange: 1, tblPrChange: 1, tcPrChange: 4 });
    const [note] = await session.stories.footnotes();
    expect(note!.doc.textContent).toContain('A tracked footnote.');
    expect(marked(note!.doc)).toEqual([]);
  });

  test('unedited: every part byte for byte, the self-check equal', async () => {
    const bytes = await fixture(name);
    const session = await EditorDocument.open(bytes);
    const saved = await session.save();
    expect(saved.changed).toBe(false);
    expect(saved.selfCheck?.equal).toBe(true);
    expect(comparePackages(bytes, saved.bytes)).toEqual([]);
  });
});

/** The main part's paragraph elements after a save, reloaded. */
async function savedParagraphs(session: EditorDocument): Promise<Element[]> {
  const saved = await session.save();
  expect(saved.selfCheck?.equal).toBe(true);
  const pkg = await WordprocessingMLPackage.load(saved.bytes);
  const doc = await pkg.getMainDocumentPart().getContents();
  return (doc.body?.content ?? []) as Element[];
}

/** The revision elements of a paragraph element as `name:author`, and its text items' names. */
function revisionsOf(p: Element): { revisions: string[]; items: string[] } {
  const revisions: string[] = [];
  const items: string[] = [];
  const walk = (list: Element[] | undefined): void => {
    for (const el of list ?? []) {
      const local = el.name.localPart;
      if (['ins', 'del', 'moveFrom', 'moveTo'].includes(local)) { revisions.push(`${local}:${(el.value as { author: string }).author}`); walk(runItemsOf(el.value as object) ?? []); }
      else if (local === 'r') for (const item of (el.value as { content?: Element[] }).content ?? []) items.push(item.name.localPart);
    }
  };
  walk(runItemsOf(p.value as object) ?? []);
  return { revisions, items };
}

/** Every revision `w:id` in the saved main part: none is used twice. */
async function revisionIds(session: EditorDocument): Promise<number[]> {
  const saved = await session.save({ selfCheck: false });
  const xml = await (await WordprocessingMLPackage.load(saved.bytes)).getMainDocumentPart().getXml();
  return [...xml.matchAll(/<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|moveFromRangeStart|moveToRangeStart)\b[^>]*\bw:id="(\d+)"/g)].map((m) => Number(m[1]));
}

describe('an edited block keeps its revisions (ED-005 section 12.3)', () => {
  // The Word files were saved with tracking on; these edits are C1's untracked ones, and the
  // tracker's are track.test.mts's (section 12.5).
  const untracked = async (name: string): Promise<EditorDocument> => { const s = await EditorDocument.open(await fixture(name)); s.trackingMode = 'Off'; return s; };
  test('typing in a paragraph with insertions and deletions: the same elements go back, deleted text as w:delText', async () => {
    const session = await untracked(WORD[0]!);
    const b = block(session.doc, 'The quick brown');
    runCommand(session, 'insertText', { text: 'Indeed, ' }, { from: b.pos + 1 });
    const content = await savedParagraphs(session);
    const p = content.find((el) => JSON.stringify(el.value).includes('Indeed, '))!;
    const { revisions, items } = revisionsOf(p);
    expect(revisions).toEqual(['ins:Author B', 'ins:Author A', 'del:Author A', 'del:Author B']);
    expect(items.filter((i) => i === 'delText')).toHaveLength(3);
  });

  test('text typed right after a deletion goes after it, not into it', async () => {
    const session = await untracked(WORD[0]!);
    const b = block(session.doc, 'The quick brown');
    // The caret after "…at dawn" (deleted), before the full stop.
    let at = -1;
    b.node.descendants((n, rel) => { if (n.isText && n.text === 'wn') at = b.pos + 1 + rel + n.nodeSize; return true; });
    runCommand(session, 'insertText', { text: ' today' }, { from: at });
    const node = block(session.doc, 'The quick brown').node;
    expect(marked(node)).toContain('wn|deletion:Author B');
    expect(marked(node).some((m) => m.startsWith(' today'))).toBe(false);
    const content = await savedParagraphs(session);
    const p = content.find((el) => JSON.stringify(el.value).includes(' today'))!;
    const dels = (runItemsOf(p.value as object) ?? []).filter((el) => el.name.localPart === 'del');
    expect(JSON.stringify(dels)).not.toContain(' today');
  });

  test("a paragraph split whose mark is inserted, with tracking off: the new mark is plain and the old one keeps the insertion (Jason's finding, section 12.16); every id unique", async () => {
    const session = await untracked(WORD[0]!);
    const b = block(session.doc, 'Split this p');
    runCommand(session, 'insertParagraph', undefined, { from: b.pos + 1 + 5 });
    expect((block(session.doc, 'Split').node.attrs as ParagraphAttrs).markRevision).toBeNull();
    expect((block(session.doc, ' this p').node.attrs as ParagraphAttrs).markRevision?.insertion?.author).toBe('Author A');
    const ids = await revisionIds(session);
    expect(new Set(ids).size).toBe(ids.length);
    const content = await savedParagraphs(session);
    const marks = content.filter((el) => JSON.stringify(el.value).includes('"ins"') && /Split| this p/.test(JSON.stringify(el.value)));
    expect(marks).toHaveLength(1);
    expect(JSON.stringify(marks[0]!.value)).toContain(' this p');
  });

  test('a cell of the inserted row edited: the row stays inserted', async () => {
    const session = await untracked(WORD[1]!);
    const table = block(session.doc, 'RowActionNote');
    let at = -1;
    table.node.child(3).descendants((n, rel) => { if (at < 0 && n.isText) at = rel; return at < 0; });
    let rowPos = table.pos + 1;
    for (let r = 0; r < 3; r++) rowPos += table.node.child(r).nodeSize;
    runCommand(session, 'insertText', { text: 'very ' }, { from: rowPos + 1 + at });
    const content = await savedParagraphs(session);
    // With tracking off the typed "very " is plain, the inserted row's own text still inserted (core-ts check 21).
    const tbl = content.find((el) => el.name.localPart === 'tbl' && /"very "/.test(JSON.stringify(el.value)))!;
    expect(JSON.stringify(tbl.value)).toMatch(/"trPr":\{[^}]*"ins":/);
  });

  test("an opaque item inside a revision (a bookmark in a w:ins and in a w:del) keeps its holder: an edit elsewhere in the paragraph leaves one insertion and one deletion (the C1 review's finding, section 12.3 item 11)", async () => {
    // tracked-changes.docx with a `_GoBack` pair inside its insertion, as Word leaves one where the
    // last edit was, and a bookmark pair inside its deletion; none of the Word fixtures has one.
    const pkg = await WordprocessingMLPackage.load(await fixture('tracked-changes.docx'));
    const part = pkg.getMainDocumentPart();
    const xml = await part.getXml();
    const patched = xml
      .replace('<w:r><w:t>An insertion</w:t></w:r></w:ins>', '<w:r><w:t>An ins</w:t></w:r><w:bookmarkStart w:id="90" w:name="_GoBack"/><w:bookmarkEnd w:id="90"/><w:r><w:t>ertion</w:t></w:r></w:ins>')
      .replace('<w:r><w:delText xml:space="preserve"> A deletion</w:delText></w:r></w:del>', '<w:r><w:delText xml:space="preserve"> A del</w:delText></w:r><w:bookmarkStart w:id="91" w:name="here"/><w:bookmarkEnd w:id="91"/><w:r><w:delText>etion</w:delText></w:r></w:del>');
    expect(patched.match(/w:bookmarkStart/g)?.length).toBe((xml.match(/w:bookmarkStart/g)?.length ?? 0) + 2);
    part.setXml(patched);
    const session = await EditorDocument.open(await pkg.save());
    let b: { node: PMNode; pos: number } | undefined;
    session.doc.forEach((node, pos) => { if (!b && node.textContent.includes('An ins')) b = { node, pos }; });
    if (!b) throw new Error('no paragraph with the insertion');
    const atoms: string[] = [];
    b.node.forEach((n) => { if (n.type.name === 'opaqueInline' && (n.attrs['kind'] as string).startsWith('bookmark')) { const r = n.attrs['revision'] as Revision | null; atoms.push(`${String(n.attrs['kind'])}:${r?.insertion ? 'ins' : ''}${r?.deletion ? 'del' : ''}`); } });
    expect(atoms).toEqual(['bookmarkStart:ins', 'bookmarkEnd:ins', 'bookmarkStart:del', 'bookmarkEnd:del']);
    runCommand(session, 'insertText', { text: '!' }, { from: b.pos + 1 + b.node.content.size });
    const content = await savedParagraphs(session);
    const p = content.find((el) => JSON.stringify(el.value).includes('Followed by'))!;
    expect(revisionsOf(p).revisions).toEqual(['ins:Jason Harrop', 'del:Jason Harrop']);
    const saved = await (await WordprocessingMLPackage.load((await session.save()).bytes)).getMainDocumentPart().getXml();
    expect(saved).toContain('<w:ins w:id="1" w:author="Jason Harrop" w:date="2007-12-09T10:14:00Z"><w:r><w:t>An ins</w:t></w:r><w:bookmarkStart w:id="90" w:name="_GoBack"/><w:bookmarkEnd w:id="90"/><w:r><w:t>ertion</w:t></w:r></w:ins>');
    expect(saved).toContain('<w:del w:id="2" w:author="Jason Harrop" w:date="2007-12-09T10:14:00Z"><w:r><w:delText xml:space="preserve"> A del</w:delText></w:r><w:bookmarkStart w:id="91" w:name="here"/><w:bookmarkEnd w:id="91"/><w:r><w:delText>etion</w:delText></w:r></w:del>');
  });

  test('find never covers deleted text, and reads across a deletion as the accepted view does', async () => {
    const session = await untracked(WORD[0]!);
    const doc = session.doc;
    expect(findIn(doc, 'old stone', {}, session.ctx.ids)).toEqual([]);
    expect(findIn(doc, 'stone bridge.', {}, session.ctx.ids).map((h) => h.text)).toEqual(['stone bridge.']);
    expect(findIn(doc, 'the green field', {}, session.ctx.ids)).toEqual([]);
    expect(findIn(doc, 'in the field', {}, session.ctx.ids).map((h) => h.text)).toEqual(['in the field']);
  });
});

describe('moves against the engine (ED-005 section 5.6)', () => {
  // Word, measured by core-ts (check 18): accepting or rejecting either half resolves the whole
  // move, all four range markers go, and nothing else changes; core-ts does the same since 0.2.2.
  const MOVED = 'The second paragraph of the move: it goes before the first.';
  const FIRST = 'The first paragraph of the move: it stays where it is.';
  test.each(WORD.flatMap((name) => (['moveFrom', 'moveTo'] as const).flatMap((half) => (['accept', 'reject'] as const).map((action) => [name, half, action] as const))))(
    '%s: %s half, %s: the whole move resolved, as Word does it',
    async (name, half, action) => {
      const pkg = await WordprocessingMLPackage.load(await fixture(name));
      const body = await pkg.getBody();
      const moves = (): string[] => body.getTrackedChanges().filter((c) => c.target.kind === 'run' && (c.target.revision === 'moveFrom' || c.target.revision === 'moveTo')).map((c) => (c.target as { revision: string }).revision);
      expect(moves().sort()).toEqual(['moveFrom', 'moveTo']);
      const count = body.getTrackedChanges().length;
      const change = body.getTrackedChanges().find((c) => c.target.kind === 'run' && c.target.revision === half)!;
      if (action === 'accept') change.accept(); else change.reject();
      expect(moves()).toEqual([]);
      expect(body.getTrackedChanges()).toHaveLength(count - 2);
      const lines = body.text.split('\r');   // a mark is \r since core-ts 0.6.0 (CR-012 phase 4, Office JS's form)
      expect(lines.filter((l) => l === MOVED)).toHaveLength(1);
      // Accepted, the paragraph is at its destination, before the first; rejected, back at its source, after it.
      const moved = lines.indexOf(MOVED);
      const first = lines.indexOf(FIRST);
      if (action === 'accept') expect(moved).toBeLessThan(first); else expect(moved).toBeGreaterThan(first);
      const xml = await pkg.getMainDocumentPart().getXml();
      expect(xml).not.toMatch(/<w:move(From|To)(RangeStart|RangeEnd)?\b/);
    },
  );
});

describe('a room or a kept session from before (ED-005 section 5.9)', () => {
  // Schema 1 is E3's room, with a deletion as an atom; 2 is E4.c's, without the room repair's two maps (section 12.15);
  // 3 is the room repair's, without E4.e's comment threads (section 7.2).
  for (const older of [1, 2, 3]) {
    test(`a room of schema ${older} is refused by a joiner and by its creator; this editor makes schema ${ROOM_SCHEMA}`, async () => {
      const base = await fixture('tracked-changes.docx');
      const creator = await EditorDocument.open(base);
      await creator.share(base, 'fixture.docx');
      expect(creator.room?.schema).toBe(ROOM_SCHEMA);
      const old = newEditStore();
      Y.applyUpdate(old.ydoc, Y.encodeStateAsUpdate(creator.store.ydoc));
      old.room.set('schema', older);
      await expect(EditorDocument.open(base, { store: old, mode: 'join' })).rejects.toBeInstanceOf(RoomVersionError);
      await expect(EditorDocument.open(base, { store: old })).rejects.toThrow(/another version of the editor/);
    });
  }

  test('a kept session from before this projection: refused for a document with revisions, stamped for one without', async () => {
    const withRevisions = await fixture(WORD[0]!);
    const kept = await EditorDocument.open(withRevisions);
    kept.store.package.delete(PROJECTION_KEY);
    await expect(EditorDocument.open(withRevisions, { store: kept.store })).rejects.toBeInstanceOf(ProjectionVersionError);
    // Projection 2, before core-ts 0.3.0's dates and overlapping deletions (ED-005 section 12.26), the same.
    const kept2 = await EditorDocument.open(withRevisions);
    kept2.store.package.set(PROJECTION_KEY, 2);
    await expect(EditorDocument.open(withRevisions, { store: kept2.store })).rejects.toBeInstanceOf(ProjectionVersionError);
    const plain = await fixture('revisions/revisions-base.docx');
    const keptPlain = await EditorDocument.open(plain);
    keptPlain.store.package.set(PROJECTION_KEY, 2);
    const recovered = await EditorDocument.open(plain, { store: keptPlain.store });
    expect(recovered.store.package.get(PROJECTION_KEY)).toBe(PROJECTION);
  });

  test('a kept session from projection 3 (core-ts 0.6.0): refused where 0.7.0 reads the file otherwise, stamped elsewhere', async () => {
    // Measured over every fixture at the install (ED-008 section 8.8 install item 18 (c)): an element with nothing to select
    // is now an atom, and objects 0.8.0's palette reads a lightGray highlight as #C0C0C0 where the kept mark holds #D3D3D3.
    for (const name of ['mc/mc-branches.docx', 'headers-footers.docx']) {
      const bytes = await fixture(name);
      const kept = await EditorDocument.open(bytes);
      kept.store.package.set(PROJECTION_KEY, 3);
      await expect(EditorDocument.open(bytes, { store: kept.store }), name).rejects.toBeInstanceOf(ProjectionVersionError);
    }
    // A document with revisions and none of those: its projection is the same, so the session is kept.
    const withRevisions = await fixture(WORD[0]!);
    const kept = await EditorDocument.open(withRevisions);
    kept.store.package.set(PROJECTION_KEY, 3);
    const recovered = await EditorDocument.open(withRevisions, { store: kept.store });
    expect(recovered.store.package.get(PROJECTION_KEY)).toBe(PROJECTION);
  });
});
