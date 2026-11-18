// SPDX-License-Identifier: GPL-3.0-or-later
//
// E3.a, the room in the model (ED-004 section 3.10): two stores linked in memory, a creator and
// a joiner over the same base file. A failing case here is a finding for ED-004 section 10 and a
// fix to the model, never a changed expectation (section 2.1).
import { describe, test, expect } from 'vitest';
import { Fragment, type Node as PMNode } from 'prosemirror-model';
import * as Y from 'yjs';
import { EditorDocument, newEditStore, RoomNotReadyError, RoomVersionError, ROOM_SCHEMA, isBaseId, runCommand } from '../src/index.mjs';
import { fixture, comparePackages } from './helpers.mjs';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { room, nodeIds, RELEASE_ORDERS } from './room-harness.mjs';

/** A paragraph with properties the editor does not model: what a rebuild from the builders would lose. */
const RICH = (text: string): string => `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:suppressLineNumbers/><w:suppressAutoHyphens/></w:pPr><w:r><w:rPr><w:emboss/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

/** A paragraph holding a simple field, which the projection keeps as an opaque inline atom. */
const OPAQUE = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:fldSimple w:instr=" DATE "><w:r><w:t>today</w:t></w:r></w:fldSimple></w:p>';

async function sameSave(a: EditorDocument, b: EditorDocument): Promise<void> {
  const [x, y] = [await a.save({ selfCheck: false }), await b.save({ selfCheck: false })];
  if (process.env['DUMP_SAVES']) { const { writeFileSync } = await import('node:fs'); writeFileSync(`${process.env['DUMP_SAVES']}/a.docx`, x.bytes); writeFileSync(`${process.env['DUMP_SAVES']}/b.docx`, y.bytes); }
  expect(comparePackages(x.bytes, y.bytes)).toEqual([]);
}

describe('create and join (ED-004 section 3.1)', () => {
  test('the joiner has the creator\'s document, body ids and story ids, whatever order each reads the stories in', async () => {
    const r = await room(await fixture('numbering-stories.docx'));
    expect(r.joiner.inRoom).toBe(true);
    expect(r.joiner.room?.schema).toBe(ROOM_SCHEMA);
    expect(r.joiner.doc.eq(r.creator.doc)).toBe(true);
    expect(nodeIds(r.joiner.doc)).toEqual(nodeIds(r.creator.doc));
    const headersC = await r.creator.stories.headers();
    const notesJ = await r.joiner.stories.footnotes();
    const notesC = await r.creator.stories.footnotes();
    const headersJ = await r.joiner.stories.headers();
    expect(headersJ.map((h) => nodeIds(h.doc))).toEqual(headersC.map((h) => nodeIds(h.doc)));
    expect(notesJ.map((n) => nodeIds(n.doc))).toEqual(notesC.map((n) => nodeIds(n.doc)));
    // The creator made every note's and header's fragment at Share: the joiner binds them, it never makes one.
    for (const s of [...notesJ, ...headersJ]) expect(r.joiner.stories.hasFragment(s)).toBe(true);
    expect(r.joiner.stories.fragmentOf(notesJ[0]!)).toBe(r.joiner.store.footnotes.get(String(notesJ[0]!.id)));
    await sameSave(r.creator, r.joiner);
  });

  test('a join refuses an unsynced store, a store with no room and a newer schema; a story without its fragment is not ready', async () => {
    const base = await fixture('footnote.docx');
    await expect(EditorDocument.open(base, { store: newEditStore(), mode: 'join' })).rejects.toBeInstanceOf(RoomNotReadyError);
    const plain = await EditorDocument.open(base);
    const unshared = newEditStore();
    Y.applyUpdate(unshared.ydoc, Y.encodeStateAsUpdate(plain.store.ydoc));
    await expect(EditorDocument.open(base, { store: unshared, mode: 'join' })).rejects.toBeInstanceOf(RoomNotReadyError);
    const r = await room(base);
    const newer = newEditStore();
    Y.applyUpdate(newer.ydoc, Y.encodeStateAsUpdate(r.creator.store.ydoc));
    newer.room.set('schema', ROOM_SCHEMA + 1);
    await expect(EditorDocument.open(base, { store: newer, mode: 'join' })).rejects.toBeInstanceOf(RoomVersionError);
    // A fragment gone from the room (as if it had not arrived): the joiner's binding is not ready, and no fragment is made.
    const note = (await r.joiner.stories.footnotes())[0]!;
    r.joiner.store.footnotes.delete(String(note.id));
    expect(() => r.joiner.stories.fragmentOf(note)).toThrow(RoomNotReadyError);
    expect(r.joiner.store.footnotes.has(String(note.id))).toBe(false);
  });
});

describe('tree runs in a room: the sources map (ED-004 section 3.7)', () => {
  test('a script\'s paragraph with unmodelled properties and an opaque field reaches the joiner, whose save is the creator\'s', async () => {
    const r = await room(await fixture('tracked-changes.docx'));
    const run = await r.creator.runOnTree(async ({ body }) => {
      const p = body.paragraphs[4]!;
      p.insertText(' via a script', 'End');
      await body.insertXml(RICH('Inserted by a script'), 'After', p);
      await body.insertXml(OPAQUE, 'End');
    });
    expect(run.error).toBeUndefined();
    await r.joiner.sourcesReady();
    expect(r.joiner.sourceProblems.size).toBe(0);
    expect(r.joiner.doc.eq(r.creator.doc)).toBe(true);
    // Every node's id resolves on the joiner, to the element the creator had.
    for (const id of nodeIds(r.joiner.doc)) expect(r.joiner.ctx.ids.get(id), id).toBeDefined();
    await sameSave(r.creator, r.joiner);
  });

  test('what a single user ran before Share is already in the map: the joiner resolves it with no sweep', async () => {
    const base = await fixture('tracked-changes.docx');
    const r = await room(base, 'fixture.docx', async (creator) => {
      await creator.runOnTree(async ({ body }) => { await body.insertXml(RICH('Before the room'), 'Start'); });
    });
    await r.joiner.sourcesReady();
    expect(r.joiner.doc.child(0).textContent).toBe('Before the room');
    for (const id of nodeIds(r.joiner.doc)) expect(r.joiner.ctx.ids.get(id), id).toBeDefined();
    await sameSave(r.creator, r.joiner);
  });

  for (const order of RELEASE_ORDERS) {
    test(`two peers running scripts at once mint distinct ids and converge (released ${order})`, async () => {
      const r = await room(await fixture('tracked-changes.docx'));
      r.link.hold();
      await r.creator.runOnTree(async ({ body }) => { await body.insertXml(RICH('From the creator'), 'Start'); });
      await r.joiner.runOnTree(async ({ body }) => { await body.insertXml(RICH('From the joiner'), 'End'); });
      r.link.release(order);
      await Promise.all([r.creator.sourcesReady(), r.joiner.sourcesReady()]);
      expect(r.creator.doc.eq(r.joiner.doc)).toBe(true);
      const minted = nodeIds(r.creator.doc).filter((id) => !isBaseId(id));
      expect(new Set(minted).size).toBe(minted.length);
      expect(minted.some((id) => id.startsWith(`c${r.creator.store.ydoc.clientID}:`))).toBe(true);
      expect(minted.some((id) => id.startsWith(`c${r.joiner.store.ydoc.clientID}:`))).toBe(true);
      expect(r.creator.doc.textContent).toContain('From the creator');
      expect(r.creator.doc.textContent).toContain('From the joiner');
      await sameSave(r.creator, r.joiner);
    });
  }

  test('a run through the middle of a comment range over three paragraphs: the joiner projects the entry inside the range', async () => {
    const r = await room(await fixture('comments-span.docx'));
    const commented = (s: EditorDocument, i: number): boolean => {
      let found = false;
      s.doc.child(i).descendants((n) => { if (n.marks.some((m) => m.type.name === 'comment')) found = true; });
      return found;
    };
    expect(commented(r.creator, 1)).toBe(true);
    await r.creator.runOnTree(({ body }) => { body.paragraphs[1]!.insertText(' (edited)', 'End'); });
    await r.joiner.sourcesReady();
    expect(r.joiner.sourceProblems.size, [...r.joiner.sourceProblems.values()].join('; ')).toBe(0);
    expect(commented(r.creator, 1)).toBe(true);
    expect(commented(r.joiner, 1)).toBe(true);
    expect(r.joiner.doc.eq(r.creator.doc)).toBe(true);
    await sameSave(r.creator, r.joiner);
    await r.joiner.save();
  });

  test('ids stay stable across a run: untouched blocks keep id and element; a block edited before the run gets an entry', async () => {
    const r = await room(await fixture('tracked-changes.docx'));
    const s = r.creator;
    // An edit in the store before the run: the run's export rebuilds this paragraph.
    let pos = 0;
    for (let k = 0; k < 4; k++) pos += s.doc.child(k).nodeSize;
    expect(runCommand(s, 'font.bold', { value: true }, { from: pos + 1, to: pos + 5 }).ok).toBe(true);
    const beforeIds = nodeIds(s.doc);
    const elements = new Map(beforeIds.map((id) => [id, s.ctx.ids.get(id)]));
    const edited = s.doc.child(4).attrs['sourceId'] as string;
    const touched = s.doc.child(1).attrs['sourceId'] as string;   // body.paragraphs[1]
    await s.runOnTree(({ body }) => { body.paragraphs[1]!.insertText(' touched', 'End'); });
    const after = new Set(nodeIds(s.doc));
    const blocks: string[] = [];
    s.doc.forEach((n) => { blocks.push(n.attrs['sourceId'] as string); });
    for (const id of blocks) {
      if (after.has(id) && elements.has(id)) expect(s.ctx.ids.get(id), id).toBe(elements.get(id));
    }
    expect(blocks).not.toContain(edited);
    expect(blocks).not.toContain(touched);
    const entries = [...s.store.sources.keys()];
    expect(entries.length).toBe(2);
    expect(entries.every((id) => !isBaseId(id))).toBe(true);
    await r.joiner.sourcesReady();
    await sameSave(r.creator, r.joiner);
  });
});

// ---------------------------------------------------------------- part edits and identifiers (ED-004 sections 3.3, 3.4)

import { runCommandAsync, readDocumentMacros, writeDocumentMacros, liveLabels } from '../src/index.mjs';

/** The position of block `i`'s first character. */
function blockFrom(s: EditorDocument, i: number): number {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += s.doc.child(k).nodeSize;
  return pos + 1;
}

async function settled(...sessions: EditorDocument[]): Promise<void> {
  for (let round = 0; round < 3; round++) for (const s of sessions) { await s.sourcesReady(); await s.partsReady(); }
}

describe('part edits and identifiers in a room (ED-004 sections 3.3 and 3.4)', () => {
  for (const order of RELEASE_ORDERS) {
    test(`a list, a link, a note and a control made on each side at once: distinct ids, both definitions everywhere, the same save (released ${order})`, async () => {
      const r = await room(await fixture('tracked-changes.docx'));
      // The fixture's own hyperlinks, which the count below leaves out.
      const baseLinks = new Set<string>();
      r.creator.doc.descendants((n) => { for (const m of n.marks) if (m.type.name === 'link' && m.attrs['rId']) baseLinks.add(m.attrs['rId'] as string); return true; });
      r.link.hold();
      const c = r.creator;
      const j = r.joiner;
      expect((await runCommandAsync(c, 'list.toggle', { bullet: true }, { from: blockFrom(c, 4) })).ok).toBe(true);
      expect((await runCommandAsync(j, 'list.toggle', { bullet: false }, { from: blockFrom(j, 5) })).ok).toBe(true);
      expect((await runCommandAsync(c, 'range.hyperlink', { address: 'https://creator.example/' }, { from: blockFrom(c, 1), to: blockFrom(c, 1) + 4 })).ok).toBe(true);
      expect((await runCommandAsync(j, 'range.hyperlink', { address: 'https://joiner.example/' }, { from: blockFrom(j, 2), to: blockFrom(j, 2) + 4 })).ok).toBe(true);
      expect((await runCommandAsync(c, 'range.insertFootnote', { text: 'The creator\'s note.' }, { from: blockFrom(c, 1) + 2 })).ok).toBe(true);
      expect((await runCommandAsync(j, 'range.insertFootnote', { text: 'The joiner\'s note.' }, { from: blockFrom(j, 2) + 2 })).ok).toBe(true);
      await settled(c, j);
      r.link.release(order);
      await settled(c, j);
      expect(c.partProblems.size, [...c.partProblems.values()].join('; ')).toBe(0);
      expect(j.partProblems.size, [...j.partProblems.values()].join('; ')).toBe(0);
      expect(c.doc.eq(j.doc)).toBe(true);
      // Every identifier distinct, and random in the room's range.
      const numIds = new Set<string>();
      const links = new Set<string>();
      const notes = new Set<number>();
      // The two lists toggled on blocks 4 and 5 (the fixture has lists 1 and 2 of its own).
      for (const i of [4, 5]) numIds.add((c.doc.child(i).attrs['numPr'] as { numId: string }).numId);
      c.doc.descendants((n) => {
        for (const m of n.marks) if (m.type.name === 'link' && m.attrs['rId']) links.add(m.attrs['rId'] as string);
        if (n.type.name === 'footnoteRef') notes.add(Number(n.attrs['footnoteId']));
        return true;
      });
      expect(numIds.size).toBe(2);
      for (const id of numIds) expect(Number(id)).toBeGreaterThanOrEqual(2 ** 20);
      for (const id of baseLinks) links.delete(id);
      expect([...links].every((id) => /^rId[0-9a-z]{8}$/.test(id))).toBe(true);
      expect(links.size).toBe(2);
      expect([...notes].filter((id) => id >= 2 ** 20).length).toBe(2);
      // Each side's numbering part holds both definitions, and both relationships.
      for (const s of [c, j]) {
        const nums = new Set(((await s.main.numberingDefinitionsPart!.readContents()).num ?? []).map((n) => String(n.numId)));
        for (const id of numIds) expect(nums.has(id), `${id} in ${s === c ? 'creator' : 'joiner'}`).toBe(true);
        for (const id of links) expect(s.main.relationshipsPart!.getRelationshipById(id), id).toBeDefined();
      }
      await sameSave(c, j);
    });
  }

  test('a style picked in on one side and another edited on the other reach both; the default tab stop and a data-part value arrive', async () => {
    const r = await room(await fixture('opendope/binding-simple.docx'));
    const c = r.creator;
    const j = r.joiner;
    expect(await c.ensureStyles(['FootnoteText'])).toBe(true);
    const styles = await j.main.styleDefinitionsPart!.readContents();
    const normal = styles.style!.find((s) => s.styleId === 'Normal')!;
    normal.rPr = { ...(normal.rPr ?? {}), i: {} };
    j.main.styleDefinitionsPart!.setContents(styles);
    j.markPartChanged(j.main.styleDefinitionsPart!.partName.toString());
    await c.setDefaultTabStop(1440);
    await c.template.load({ engine: false });
    const data = c.template.dataPart.part!;
    data.documentElement.childElements[0]!.text = 'Changed in a room';
    c.markPartChanged(data.part.partName.toString());
    await settled(c, j);
    expect(j.partProblems.size, [...j.partProblems.values()].join('; ')).toBe(0);
    expect((await j.main.styleDefinitionsPart!.readContents()).style!.some((s) => s.styleId === 'FootnoteText')).toBe(true);
    expect((await c.main.styleDefinitionsPart!.readContents()).style!.find((s) => s.styleId === 'Normal')!.rPr?.i).toBeDefined();
    expect(j.defaultTabStop).toBe(1440);
    await j.template.load({ engine: false });
    expect(j.template.dataPart.part!.documentElement.childElements[0]!.text).toBe('Changed in a room');
    await sameSave(c, j);
  });

  test('after a remote list start the joiner\'s labels are the creator\'s; after a remote style edit the resolver gives the new properties (section 3.10 item 4)', async () => {
    const r = await room(await fixture('tracked-changes.docx'));
    const c = r.creator;
    const j = r.joiner;
    expect((await runCommandAsync(c, 'list.toggle', { bullet: false }, { from: blockFrom(c, 4), to: blockFrom(c, 6) })).ok).toBe(true);
    const styles = await j.main.styleDefinitionsPart!.readContents();
    const normal = styles.style!.find((s) => s.styleId === 'Normal')!;
    normal.rPr = { ...(normal.rPr ?? {}), i: {} };
    j.main.styleDefinitionsPart!.setContents(styles);
    j.markPartChanged(j.main.styleDefinitionsPart!.partName.toString());
    await settled(c, j);
    const labels = (s: EditorDocument): string[] => {
      const out: string[] = [];
      const doc = s.doc;   // one reading: the map is keyed by its nodes
      const map = liveLabels(doc, s);
      doc.forEach((b) => { out.push(map.get(b)?.listString ?? ''); });
      return out;
    };
    const joinerLabels = labels(j);
    expect(joinerLabels.filter((l) => l !== '').length).toBeGreaterThan(0);
    expect(joinerLabels).toEqual(labels(c));
    expect(c.pkg.propertyResolver.getEffectiveRPr('Normal')?.i).toBeDefined();
    expect(j.pkg.propertyResolver.getEffectiveRPr('Normal')?.i).toBeDefined();
  });

  test('a session opened from a package publishes a style edit at Share, and only that style', async () => {
    const base = await fixture('tracked-changes.docx');
    const s = await EditorDocument.open(await WordprocessingMLPackage.load(base));
    const styles = await s.main.styleDefinitionsPart!.readContents();
    const normal = styles.style!.find((x) => x.styleId === 'Normal')!;
    normal.rPr = { ...(normal.rPr ?? {}), i: {} };
    s.main.styleDefinitionsPart!.setContents(styles);
    s.markPartChanged(s.main.styleDefinitionsPart!.partName.toString());
    await s.partsReady();
    expect(s.store.styles.size).toBe(0);   // no base yet: nothing to compare with
    await s.share(base, 'fixture.docx');
    expect([...s.store.styles.keys()]).toEqual(['style:Normal']);
  });

  test('a join replays the part edits written before it', async () => {
    const base = await fixture('tracked-changes.docx');
    const r = await room(base, 'fixture.docx', async (creator) => {
      expect(await creator.ensureStyles(['FootnoteText'])).toBe(true);
      await runCommandAsync(creator, 'list.toggle', { bullet: true }, { from: blockFrom(creator, 4) });
      await creator.partsReady();
    });
    expect((await r.joiner.main.styleDefinitionsPart!.readContents()).style!.some((s) => s.styleId === 'FootnoteText')).toBe(true);
    const numId = (r.joiner.doc.child(4).attrs['numPr'] as { numId: string }).numId;
    expect(((await r.joiner.main.numberingDefinitionsPart!.readContents()).num ?? []).some((n) => String(n.numId) === numId)).toBe(true);
    await sameSave(r.creator, r.joiner);
  });

  test('a new custom XML part made in a room reaches the other peer with the same itemID and part name; outside a room it is made as before', async () => {
    const r = await room(await fixture('tracked-changes.docx'));
    await writeDocumentMacros(r.creator, [{ name: 'm', script: 'return 1;' }]);
    await settled(r.creator, r.joiner);
    expect(r.creator.partProblems.size, [...r.creator.partProblems.values()].join('; ')).toBe(0);
    expect(r.joiner.partProblems.size, [...r.joiner.partProblems.values()].join('; ')).toBe(0);
    expect((await readDocumentMacros(r.joiner)).map((m) => m.name)).toEqual(['m']);
    const parts = (s: EditorDocument) => s.pkg.customXmlParts.items.map((p) => `${p.id.toUpperCase()} ${p.part.partName.toString()}`).sort();
    await r.joiner.pkg.customXmlParts.load();
    expect(parts(r.joiner)).toEqual(parts(r.creator));
    // In a room the part name is random, so two peers making a part at once cannot take the same one.
    const made = r.creator.pkg.customXmlParts.items.map((p) => p.part.partName.toString()).find((n) => /item\d{7,}\.xml$/.test(n));
    expect(made).toBeDefined();
    await sameSave(r.creator, r.joiner);
    const alone = await EditorDocument.open(await fixture('tracked-changes.docx'));
    await writeDocumentMacros(alone, [{ name: 'm', script: 'return 1;' }]);
    expect((await readDocumentMacros(alone)).map((m) => m.name)).toEqual(['m']);
    expect(alone.pkg.customXmlParts.items.some((p) => /item\d{1,2}\.xml$/.test(p.part.partName.toString()))).toBe(true);
  });
});

describe('outside a room (ED-004 sections 3.3, 3.4)', () => {
  test('identifiers stay sequential, and a recovered session replays its own part edits', async () => {
    const base = await fixture('tracked-changes.docx');
    const s = await EditorDocument.open(base);
    await runCommandAsync(s, 'list.toggle', { bullet: true }, { from: blockFrom(s, 4) });
    const numId = Number((s.doc.child(4).attrs['numPr'] as { numId: string }).numId);
    expect(numId).toBeLessThan(2 ** 20);
    expect(await s.ensureStyles(['FootnoteText'])).toBe(true);
    await s.partsReady();
    const before = await s.save({ selfCheck: false });
    const store = newEditStore();
    Y.applyUpdate(store.ydoc, Y.encodeStateAsUpdate(s.store.ydoc));
    const recovered = await EditorDocument.open(base, { store });
    expect((await recovered.main.styleDefinitionsPart!.readContents()).style!.some((st) => st.styleId === 'FootnoteText')).toBe(true);
    const after = await recovered.save({ selfCheck: false });
    expect(comparePackages(before.bytes, after.bytes)).toEqual([]);
  });
});

// ---------------------------------------------------------------- edits during a tree run (ED-004 section 10.2 item 1)

/** Block `i`'s text. */
const textOf = (s: EditorDocument, i: number): string => s.doc.child(i).textContent;

/** Insert text at the start of block `i` through the command a keystroke runs. */
function type(s: EditorDocument, i: number, text: string): void {
  expect(runCommand(s, 'insertText', { text }, { from: blockFrom(s, i) }).ok).toBe(true);
}

describe('edits arriving while a tree run awaits (ED-004 section 10.2 item 1)', () => {
  test('a peer\'s edit of another block during the run survives it; the saves agree and pass the default self-check', async () => {
    const r = await room(await fixture('tracked-changes.docx'));
    const target = textOf(r.creator, 3);
    await r.creator.runOnTree(async ({ body }) => {
      type(r.joiner, 3, 'REMOTE ');
      await Promise.resolve();
      body.paragraphs[1]!.insertText(' by the script', 'End');
    });
    await settled(r.creator, r.joiner);
    for (const s of [r.creator, r.joiner]) {
      expect(textOf(s, 3)).toBe(`REMOTE ${target}`);
      expect(textOf(s, 1)).toMatch(/ by the script$/);
    }
    expect(r.joiner.doc.eq(r.creator.doc)).toBe(true);
    await sameSave(r.creator, r.joiner);
    // The default save runs the self-check, in a room after a tree run as anywhere.
    await r.creator.save();
    await r.joiner.save();
  });

  test('a paragraph a peer inserts or deletes during the run stays inserted or deleted, around the script\'s own', async () => {
    const r = await room(await fixture('tracked-changes.docx'));
    const count = r.creator.doc.childCount;
    const gone = textOf(r.creator, 5);
    await r.creator.runOnTree(async ({ body }) => {
      // The joiner: a new paragraph before block 3, and block 5 removed.
      const j = r.joiner.doc;
      const blocks: PMNode[] = [];
      j.forEach((b, _o, i) => {
        if (i === 3) blocks.push(j.type.schema.nodes['paragraph']!.create(null, j.type.schema.text('From the peer')));
        if (i !== 5) blocks.push(b);
      });
      r.joiner.replace(j.copy(Fragment.from(blocks)));
      await body.insertXml(RICH('From the script'), 'After', body.paragraphs[1]!);
    });
    await settled(r.creator, r.joiner);
    for (const s of [r.creator, r.joiner]) {
      expect(s.doc.childCount).toBe(count + 1);
      const texts: string[] = [];
      s.doc.forEach((b) => { texts.push(b.textContent); });
      expect(texts).not.toContain(gone);
      expect(texts[2]).toBe('From the script');
      expect(texts[4]).toBe('From the peer');
    }
    await sameSave(r.creator, r.joiner);
  });

  test('a block both change during the run goes to the script', async () => {
    const r = await room(await fixture('tracked-changes.docx'));
    await r.creator.runOnTree(({ body }) => {
      type(r.joiner, 1, 'REMOTE ');
      body.paragraphs[1]!.insertText(' by the script', 'End');
    });
    await settled(r.creator, r.joiner);
    for (const s of [r.creator, r.joiner]) {
      expect(textOf(s, 1)).toMatch(/ by the script$/);
      expect(textOf(s, 1)).not.toContain('REMOTE');
    }
    await sameSave(r.creator, r.joiner);
  });

  test('a block the run\'s export rebuilt and a peer edits during the run keeps the peer\'s edit and still exports', async () => {
    const r = await room(await fixture('tracked-changes.docx'));
    const s = r.creator;
    expect(runCommand(s, 'font.bold', { value: true }, { from: blockFrom(s, 4), to: blockFrom(s, 4) + 4 }).ok).toBe(true);
    await settled(s, r.joiner);
    await s.runOnTree(({ body }) => {
      type(r.joiner, 4, 'REMOTE ');
      body.paragraphs[1]!.insertText(' by the script', 'End');
    });
    await settled(s, r.joiner);
    for (const x of [s, r.joiner]) expect(textOf(x, 4)).toMatch(/^REMOTE /);
    for (const id of nodeIds(s.doc)) expect(s.ctx.ids.get(id), id).toBeDefined();
    await sameSave(s, r.joiner);
    await s.save();
  });
});

describe('a single user typing during a macro\'s await (ED-004 section 10.2 item 1)', () => {
  test('the keystrokes made while the run awaited survive it, and the save passes the self-check', async () => {
    const s = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const target = textOf(s, 3);
    await s.runOnTree(async ({ body }) => {
      await Promise.resolve();
      type(s, 3, 'TYPED ');
      body.paragraphs[1]!.insertText(' by the macro', 'End');
    });
    expect(textOf(s, 3)).toBe(`TYPED ${target}`);
    expect(textOf(s, 1)).toMatch(/ by the macro$/);
    await s.save();
  });
});

// ---------------------------------------------------------------- Re-paginate in a room (ED-004 section 3.8)

import { paginateDouble } from './fixtures/build/companion-double.mjs';
import { PaginationMismatchError, isMarker } from '../src/index.mjs';

/** The indexes of the top-level blocks holding a page-break marker. */
function markerBlocks(s: EditorDocument): number[] {
  const out: number[] = [];
  s.doc.forEach((block, _o, index) => { block.descendants((n) => { if (isMarker(n) && !out.includes(index)) out.push(index); return true; }); });
  return out;
}

describe('Re-paginate in a room (ED-004 section 3.8)', () => {
  test('the markers move on both peers and nothing else: no sources entry, one undo step, a marker built from its kind saves as <w:lastRenderedPageBreak/>', async () => {
    const r = await room(await fixture('notes-numbering.docx'));
    const c = r.creator;
    expect(markerBlocks(c)).toEqual([4]);
    const texts = (s: EditorDocument): string[] => { const out: string[] = []; s.doc.forEach((b) => { out.push(b.textContent); }); return out; };
    const before = texts(c);
    const sources = c.store.sources.size;
    const answer = await paginateDouble((await c.save({ selfCheck: false })).bytes);
    const result = await c.adoptPaginated(answer);
    expect(result.error).toBeUndefined();
    expect(result.changed).toBe(true);
    expect(c.paginationSource).toBe('companion');
    await settled(c, r.joiner);
    for (const s of [c, r.joiner]) {
      expect(markerBlocks(s)).toEqual([5]);
      expect(texts(s)).toEqual(before);
    }
    expect(c.store.sources.size).toBe(sources);
    await sameSave(c, r.joiner);
    // The default save's self-check reloads the built marker and finds it where the screen has it.
    const saved = await r.joiner.save();
    expect(saved.selfCheck?.equal).toBe(true);
    expect(markerBlocks(await EditorDocument.open(saved.bytes))).toEqual([5]);
    c.undoManager.undo();
    expect(markerBlocks(c)).toEqual([4]);
    expect(markerBlocks(r.joiner)).toEqual([4]);
  });

  test('a paragraph edited since the save keeps its markers; an answer with a different number of paragraphs is refused and changes nothing', async () => {
    const r = await room(await fixture('notes-numbering.docx'));
    const c = r.creator;
    const answer = await paginateDouble((await c.save({ selfCheck: false })).bytes);
    type(r.joiner, 5, 'Edited meanwhile. ');
    await settled(c, r.joiner);
    expect((await c.adoptPaginated(answer)).error).toBeUndefined();
    // Block 4's text agrees, so its marker goes; block 5's does not, so it gains none.
    expect(markerBlocks(c)).toEqual([]);
    // An answer with a paragraph more.
    const { WordprocessingMLPackage: Pkg } = await import('@docx4j/core-ts');
    const pkg = await Pkg.load(answer);
    const doc = await pkg.getMainDocumentPart().getContents();
    const content = doc.body!.content!;
    content.splice(1, 0, structuredClone({ ...content[0]!, value: { ...(content[0]!.value as object), PARENT: undefined } }) as typeof content[0]);
    pkg.getMainDocumentPart().setContents(doc);
    const longer = await pkg.save();
    const snapshot = c.doc;
    const refused = await c.adoptPaginated(longer);
    expect(refused.error).toBeInstanceOf(PaginationMismatchError);
    expect(refused.changed).toBe(false);
    expect(c.doc.eq(snapshot)).toBe(true);
  });
});

describe('problems shown (ED-004 section 4.4)', () => {
  test('an entry that cannot be applied is a room problem with its key and reason, and the listeners hear of it', async () => {
    const r = await room(await fixture('tracked-changes.docx'));
    let heard = 0;
    r.joiner.onSourcesReceived(() => { heard++; });
    let parts = 0;
    r.joiner.onPartsReceived(() => { parts++; });
    r.creator.store.ydoc.transact(() => {
      r.creator.store.sources.set('c1:e999', { xml: '<not-wordprocessingml/>', ids: { '': 'c1:e999' } });
      r.creator.store.parts.set('part:/word/nowhere.xml', { contentType: 'application/xml', xml: '<x/>' });
    });
    await settled(r.creator, r.joiner);
    expect(heard).toBeGreaterThan(0);
    expect(parts).toBeGreaterThan(0);
    const problems = r.joiner.roomProblems;
    expect(problems.find((p) => p.kind === 'source')?.key).toBe('c1:e999');
    expect(problems.find((p) => p.kind === 'part')?.key).toBe('parts|part:/word/nowhere.xml');
    expect(problems.every((p) => p.reason.length > 0)).toBe(true);
  });
});

// ---------------------------------------------------------------- note ids in the file (ED-004 section 10.8 item 6)

import { ZipPartStore } from '@docx4j/core-ts';

describe('note ids in the file (ED-004 section 10.8 item 6)', () => {
  test('notes made in a room are written with the smallest free ids in document order; the room keeps its random ids; every peer writes the same', async () => {
    const r = await room(await fixture('notes-numbering.docx'));
    r.link.hold();
    // The joiner's note (block 5) goes later in the document than the creator's (block 2).
    expect((await runCommandAsync(r.joiner, 'range.insertFootnote', { text: 'The joiner\'s note.' }, { from: blockFrom(r.joiner, 5) + 2 })).ok).toBe(true);
    expect((await runCommandAsync(r.creator, 'range.insertFootnote', { text: 'The creator\'s note.' }, { from: blockFrom(r.creator, 2) + 2 })).ok).toBe(true);
    await settled(r.creator, r.joiner);
    r.link.release('interleaved');
    await settled(r.creator, r.joiner);
    const random = new Set<number>();
    r.creator.doc.descendants((n) => { if (n.type.name === 'footnoteRef' && Number(n.attrs['footnoteId']) >= 2 ** 20) random.add(Number(n.attrs['footnoteId'])); return true; });
    expect(random.size).toBe(2);
    // The default save runs the self-check under the mapping.
    const [a, b] = [await r.creator.save(), await r.joiner.save()];
    expect(comparePackages(a.bytes, b.bytes)).toEqual([]);
    const text = (bytes: Uint8Array, part: string): string => new TextDecoder().decode(new ZipPartStore(bytes).loadSync(part));
    const notes = [...text(a.bytes, 'word/footnotes.xml').matchAll(/<w:footnote\b[^>]*w:id="(-?\d+)"/g)].map((m) => Number(m[1]));
    const refs = [...text(a.bytes, 'word/document.xml').matchAll(/<w:footnoteReference w:id="(\d+)"/g)].map((m) => Number(m[1]));
    expect(notes).toEqual([-1, 0, 1, 2, 3, 4, 5, 6, 7]);
    // In document order: the creator's (block 2) is 6 and the joiner's (block 5) 7, whatever their random ids.
    expect([...refs].sort((x, y) => x - y)).toEqual([1, 2, 3, 4, 5, 6, 7]);
    const reopened = await EditorDocument.open(a.bytes);
    const reread = await reopened.stories.footnotes();
    expect(reread.find((n) => n.id === 6)?.doc.textContent).toContain('The creator\'s note.');
    expect(reread.find((n) => n.id === 7)?.doc.textContent).toContain('The joiner\'s note.');
    // The room still speaks the random ids, and a second save writes the same file.
    const after = new Set<number>();
    r.creator.doc.descendants((n) => { if (n.type.name === 'footnoteRef' && Number(n.attrs['footnoteId']) >= 2 ** 20) after.add(Number(n.attrs['footnoteId'])); return true; });
    expect(after).toEqual(random);
    expect(comparePackages((await r.creator.save()).bytes, a.bytes)).toEqual([]);
  });

  test('outside a room nothing is renumbered: a note made alone keeps its sequential id', async () => {
    const s = await EditorDocument.open(await fixture('notes-numbering.docx'));
    expect((await runCommandAsync(s, 'range.insertFootnote', { text: 'Alone.' }, { from: blockFrom(s, 1) + 2 })).ok).toBe(true);
    const saved = await s.save();
    const xml = new TextDecoder().decode(new ZipPartStore(saved.bytes).loadSync('word/footnotes.xml'));
    expect([...xml.matchAll(/<w:footnote\b[^>]*w:id="(-?\d+)"/g)].map((m) => Number(m[1]))).toEqual([-1, 0, 1, 2, 3, 4, 5, 6]);
  });
});

describe('a save and a tree run started together (ED-004 section 10.9 item 1)', () => {
  test('they run one after the other: the run never sees the file\'s note ids, and the room keeps its own', async () => {
    const r = await room(await fixture('notes-numbering.docx'));
    const c = r.creator;
    expect((await runCommandAsync(c, 'range.insertFootnote', { text: 'A note made in the room.' }, { from: blockFrom(c, 2) + 2 })).ok).toBe(true);
    await settled(c, r.joiner);
    const roomIds = (s: EditorDocument): number[] => {
      const out: number[] = [];
      s.doc.descendants((n) => { if (n.type.name === 'footnoteRef') out.push(Number(n.attrs['footnoteId'])); return true; });
      return out;
    };
    const before = roomIds(c);
    expect(before.some((id) => id >= 2 ** 20)).toBe(true);
    // The run starts inside the save's zip, while the trees carry the file's ids: the one window
    // that matters, entered for certain by starting the run from the zip itself.
    const zip = c.pkg.save.bind(c.pkg);
    let running: ReturnType<typeof c.runOnTree> | undefined;
    c.pkg.save = (async (...args: Parameters<typeof zip>) => {
      running ??= c.runOnTree(({ body }) => { body.paragraphs[0]!.insertText(' (a script)', 'End'); });
      await new Promise((resolve) => setTimeout(resolve, 20));
      return zip(...args);
    }) as typeof c.pkg.save;
    await c.save();
    c.pkg.save = zip;
    const run = await running!;
    expect(run.error).toBeUndefined();
    await settled(c, r.joiner);
    expect(roomIds(c)).toEqual(before);
    expect(roomIds(r.joiner)).toEqual(before);
    const [a, b] = [await c.save(), await r.joiner.save()];
    expect(comparePackages(a.bytes, b.bytes)).toEqual([]);
  });
});
