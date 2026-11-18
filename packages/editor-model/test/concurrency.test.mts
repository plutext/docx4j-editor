// SPDX-License-Identifier: GPL-3.0-or-later
//
// E3.d, concurrent structure and undo across peers in the model (ED-004 sections 6.1 and 6.2):
// two peers linked in memory make conflicting structural edits with their updates held back,
// released in each of the four orders. Every case ends with both documents equal, both saves
// byte-identical part by part and passing the default self-check, and no table written with rows
// of different widths. A failing case is a finding for ED-004 section 10 and a fix to the model,
// never a changed expectation (section 2.1).
import { describe, test, expect } from 'vitest';
import { Fragment, type Node as PMNode } from 'prosemirror-model';
import { CellSelection, TableMap } from 'prosemirror-tables';
import { EditorDocument, runCommand, runCommandAsync, commandNamed, stateFor, trackedChangesOf, repairWrites, applyPartXml, checkPartXml } from '../src/index.mjs';
import { ZipPartStore } from '@docx4j/core-ts';
import { fixture, comparePackages, mainXml } from './helpers.mjs';
import { room, RELEASE_ORDERS, type Room, type ReleaseOrder } from './room-harness.mjs';

/** The position of block `i`'s first character. */
function blockFrom(s: EditorDocument, i: number): number {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += s.doc.child(k).nodeSize;
  return pos + 1;
}

/** Inside the first paragraph of cell (r, c) of the table at block `block`. */
function cellStart(s: EditorDocument, block: number, r: number, c: number): number {
  const doc = s.doc;
  let pos = 0;
  for (let k = 0; k < block; k++) pos += doc.child(k).nodeSize;
  const map = TableMap.get(doc.child(block));
  return pos + 1 + map.map[r * map.width + c]! + 2;
}

/** The position of cell (r, c) itself, for a cell selection. */
function cellPos(s: EditorDocument, block: number, r: number, c: number): number {
  let pos = 0;
  for (let k = 0; k < block; k++) pos += s.doc.child(k).nodeSize;
  return pos + 1 + TableMap.get(s.doc.child(block)).map[r * TableMap.get(s.doc.child(block)).width + c]!;
}

async function settled(...sessions: EditorDocument[]): Promise<void> {
  for (let round = 0; round < 3; round++) for (const s of sessions) { await s.sourcesReady(); await s.partsReady(); }
}

/**
 * Each table's structural problems, in order: rows of different widths, spans that collide or
 * overrun. Column-width mismatches are left out (a reload shows them the same way), and so is
 * what the base file already had: `tables.docx`'s block 29 has rows Word wrote short on purpose.
 */
/** A document's top-level blocks as their text, a table as its cells'. */
function blockTexts(doc: PMNode): string[] {
  const out: string[] = [];
  doc.forEach((b) => out.push(b.textContent));
  return out;
}

function structural(doc: PMNode): number[] {
  const out: number[] = [];
  doc.descendants((n) => {
    if (n.type.name !== 'table') return true;
    out.push((TableMap.get(n).problems ?? []).filter((p) => p.type !== 'colwidth mismatch').length);
    return true;
  });
  return out;
}

/** Both peers equal, both saves the same bytes and passing the self-check, no table written broken by the edits. */
async function converged(r: Room): Promise<void> {
  await settled(r.creator, r.joiner);
  expect(r.joiner.doc.eq(r.creator.doc)).toBe(true);
  const [a, b] = [await r.creator.save(), await r.joiner.save()];
  expect(comparePackages(a.bytes, b.bytes)).toEqual([]);
  const reopened = await EditorDocument.open(a.bytes);
  const base = await EditorDocument.open(r.base);
  const after = structural(reopened.doc);
  const irregular = structural(base.doc);
  // A table the base file had irregular may still be; every other table is regular.
  after.forEach((n, i) => { if (!(irregular[i]! > 0)) expect(n, `table ${i}`).toBe(0); });
}

/** Run a command over a cell selection (a merge), as the view would. */
function onCells(s: EditorDocument, name: string, block: number, from: [number, number], to: [number, number]): void {
  const doc = s.doc;
  const state = stateFor(s, CellSelection.create(doc, cellPos(s, block, ...from), cellPos(s, block, ...to)));
  let next = state;
  expect(commandNamed(name)!.run(undefined)(state, (tr) => { next = state.apply(tr); })).toBe(true);
  s.replace(next.doc);
}

/** The document without block `i` (a peer deleting a paragraph). */
function without(s: EditorDocument, i: number): PMNode {
  const blocks: PMNode[] = [];
  s.doc.forEach((b, _o, k) => { if (k !== i) blocks.push(b); });
  return s.doc.copy(Fragment.from(blocks));
}

describe('concurrent structure (ED-004 section 6.1)', () => {
  for (const order of RELEASE_ORDERS) {
    describe(`released ${order}`, () => {
      test('1. two peers add a column to the same table; a row added and a row deleted at the same place', async () => {
        const r = await room(await fixture('tables.docx'));
        r.link.hold();
        expect(runCommand(r.creator, 'table.addColumns', { location: 'After' }, { from: cellStart(r.creator, 1, 0, 0) }).ok).toBe(true);
        expect(runCommand(r.joiner, 'table.addColumns', { location: 'After' }, { from: cellStart(r.joiner, 1, 0, 0) }).ok).toBe(true);
        expect(runCommand(r.creator, 'table.addRows', { location: 'After' }, { from: cellStart(r.creator, 7, 1, 0) }).ok).toBe(true);
        expect(runCommand(r.joiner, 'table.deleteRows', undefined, { from: cellStart(r.joiner, 7, 1, 0) }).ok).toBe(true);
        r.link.release(order);
        await converged(r);
      });

      test('1. a column added on one side and a row on the other: the new row lacks the new column\'s cell, and the save writes it whole', async () => {
        const r = await room(await fixture('tables.docx'));
        r.link.hold();
        expect(runCommand(r.creator, 'table.addColumns', { location: 'After' }, { from: cellStart(r.creator, 1, 0, 0) }).ok).toBe(true);
        expect(runCommand(r.joiner, 'table.addRows', { location: 'After' }, { from: cellStart(r.joiner, 1, 1, 0) }).ok).toBe(true);
        r.link.release(order);
        await settled(r.creator, r.joiner);
        // The merged table is irregular in the store, on both peers alike: the view repairs it on the next local change.
        expect(structural(r.creator.doc)[0]).toBeGreaterThan(0);
        await converged(r);
      });

      test('1. a cell merge on one side against a cell split on the other, in the same table', async () => {
        const r = await room(await fixture('tables.docx'));
        r.link.hold();
        // Block 3: the second column's cell spans both rows. The creator splits it; the joiner merges the first column's two cells.
        expect(runCommand(r.creator, 'table.splitCell', undefined, { from: cellStart(r.creator, 3, 0, 1) }).ok).toBe(true);
        onCells(r.joiner, 'table.mergeCells', 3, [0, 0], [1, 0]);
        r.link.release(order);
        await converged(r);
      });

      test('2. a paragraph split on one side while the other types at the split point; Enter at the end of the same paragraph on both', async () => {
        const r = await room(await fixture('tracked-changes.docx'));
        r.link.hold();
        expect(runCommand(r.creator, 'insertParagraph', undefined, { from: blockFrom(r.creator, 4) + 4 }).ok).toBe(true);
        expect(runCommand(r.joiner, 'insertText', { text: 'typed' }, { from: blockFrom(r.joiner, 4) + 4 }).ok).toBe(true);
        const end = (s: EditorDocument): number => blockFrom(s, 1) + s.doc.child(1).content.size;
        expect(runCommand(r.creator, 'insertParagraph', undefined, { from: end(r.creator) }).ok).toBe(true);
        expect(runCommand(r.joiner, 'insertParagraph', undefined, { from: end(r.joiner) }).ok).toBe(true);
        r.link.release(order);
        await converged(r);
        expect(r.creator.doc.textContent).toContain('typed');
      });

      test('3. a list started on the same paragraph by both peers', async () => {
        const r = await room(await fixture('tracked-changes.docx'));
        r.link.hold();
        expect((await runCommandAsync(r.creator, 'list.toggle', { bullet: true }, { from: blockFrom(r.creator, 4) })).ok).toBe(true);
        expect((await runCommandAsync(r.joiner, 'list.toggle', { bullet: false }, { from: blockFrom(r.joiner, 4) })).ok).toBe(true);
        await settled(r.creator, r.joiner);
        r.link.release(order);
        await converged(r);
        expect(r.creator.doc.child(4).attrs['numPr']).not.toBeNull();
      });

      test('4. a control inserted around text the other peer is deleting; a row group wrapped while the other deletes one of its rows', async () => {
        const r = await room(await fixture('tables.docx'));
        r.link.hold();
        const p = 0;   // the first paragraph
        expect(runCommand(r.creator, 'contentControl.insert', { kind: 'RichText', tag: 'held' }, { from: blockFrom(r.creator, p), to: blockFrom(r.creator, p) + 4 }).ok).toBe(true);
        expect(runCommand(r.joiner, 'insertText', { text: '' }, { from: blockFrom(r.joiner, p), to: blockFrom(r.joiner, p) + 6 }).ok).toBe(true);
        expect(runCommand(r.creator, 'contentControl.wrapRows', { kind: 'RepeatingSection', tag: 'rows' }, { from: cellStart(r.creator, 7, 1, 0), to: cellStart(r.creator, 7, 2, 0) }).ok).toBe(true);
        expect(runCommand(r.joiner, 'table.deleteRows', undefined, { from: cellStart(r.joiner, 7, 2, 0) }).ok).toBe(true);
        r.link.release(order);
        await converged(r);
      });

      test('5. a section break inserted in a paragraph the other peer deletes', async () => {
        const r = await room(await fixture('tracked-changes.docx'));
        r.link.hold();
        expect(runCommand(r.creator, 'section.insertBreak', { type: 'nextPage' }, { from: blockFrom(r.creator, 4) + 2 }).ok).toBe(true);
        r.joiner.replace(without(r.joiner, 4));
        r.link.release(order);
        await converged(r);
      });

      test('6. a note inserted in a paragraph the other peer deletes: the note is not written', async () => {
        const r = await room(await fixture('tracked-changes.docx'));
        r.link.hold();
        expect((await runCommandAsync(r.creator, 'range.insertFootnote', { text: 'An orphan to be.' }, { from: blockFrom(r.creator, 4) + 2 })).ok).toBe(true);
        r.joiner.replace(without(r.joiner, 4));
        await settled(r.creator, r.joiner);
        r.link.release(order);
        await converged(r);
        const saved = await r.creator.save();
        const reopened = await EditorDocument.open(saved.bytes);
        const notes = await reopened.stories.footnotes();
        expect(notes.some((n) => n.doc.textContent.includes('An orphan to be.'))).toBe(false);
      });
    });
  }
});

/** Block `i`'s text as pieces: the text, 'ins', 'del' or '' and the revision's author; neighbours alike joined. */
function revisionsIn(s: EditorDocument, i: number): [string, string, string | null][] {
  const out: [string, string, string | null][] = [];
  s.doc.child(i).forEach((n) => {
    if (!n.isText) return;
    const m = n.marks.find((k) => k.type.name === 'insertion' || k.type.name === 'deletion');
    const kind = m ? (m.type.name === 'insertion' ? 'ins' : 'del') : '';
    const author = (m?.attrs['author'] as string | undefined) ?? null;
    const last = out[out.length - 1];
    if (last && last[1] === kind && last[2] === author) last[0] += n.text!; else out.push([n.text!, kind, author]);
  });
  return out;
}

/** A room with tracking on, turned on by the creator and replicated; the creator tracks as Author A, the joiner as Author B. */
async function trackingRoom(): Promise<Room> {
  const r = await room(await fixture('tracked-changes.docx'));
  expect((await runCommandAsync(r.creator, 'document.setChangeTrackingMode', { value: 'TrackAll' })).ok).toBe(true);
  await settled(r.creator, r.joiner);
  expect(r.joiner.trackingMode).toBe('TrackAll');
  r.creator.trackingAuthor = 'Author A';
  r.joiner.trackingAuthor = 'Author B';
  return r;
}

/** Block 4 is "Some text."; A deletes "Some te" while B deletes "text.", held and released in `order`. */
async function overlappingDeletions(order: ReleaseOrder): Promise<Room> {
  const r = await trackingRoom();
  const b = blockFrom(r.creator, 4);
  r.link.hold();
  expect(runCommand(r.creator, 'insertText', { text: '' }, { from: b, to: b + 7 }).ok).toBe(true);
  expect(runCommand(r.joiner, 'insertText', { text: '' }, { from: b + 5, to: b + 10 }).ok).toBe(true);
  r.link.release(order);
  await settled(r.creator, r.joiner);
  return r;
}

/** A's insertion "new words " in block 4, seen by both; then A types "more " inside it while B accepts or rejects it by its key. */
async function resolvingWhileTyping(order: ReleaseOrder, action: 'accept' | 'reject'): Promise<Room> {
  const r = await trackingRoom();
  const b = blockFrom(r.creator, 4);
  expect(runCommand(r.creator, 'insertText', { text: 'new words ' }, { from: b + 5 }).ok).toBe(true);
  await settled(r.creator, r.joiner);
  const [change] = trackedChangesOf(r.joiner.doc).filter((c) => c.author === 'Author A');
  const key = r.joiner.changeKey(change!);
  r.link.hold();
  expect(runCommand(r.creator, 'insertText', { text: 'more ' }, { from: b + 9 }).ok).toBe(true);
  expect(runCommand(r.joiner, `trackedChange.${action}`, { key }, { from: b }).ok).toBe(true);
  r.link.release(order);
  await settled(r.creator, r.joiner);
  return r;
}

// Step C5 (ED-005 section 5.10). Every case converges. Three do not end sensibly, for one reason: Yjs
// keeps formatting as markers at positions, so text inside a formatted range carries no marker
// of its own, and a concurrent change to the range's markers reaches text its author never saw
// (docs/developer/upstream-limitations.md item 10; ED-005 section 12.7 item 5). They were marked
// `test.fails` until the fixes Jason chose landed: a repair any peer runs after E4.d (ED-005 section
// 12.15: an accept or a reject while the other types), and overlapping deletion marks with the
// core-ts 0.3.0 upgrade (section 12.26). The fixes removed `.fails`, never the expectations.
describe('tracking concurrently (ED-005 section 5.10, step C5)', () => {
  for (const order of RELEASE_ORDERS) {
    describe(`released ${order}`, () => {
      test('two authors delete overlapping text: the peers converge, and nothing is lost', async () => {
        const r = await overlappingDeletions(order);
        await converged(r);
        expect(r.creator.doc.child(4).textContent).toBe('Some text.');
        // Accepted on one peer, what is deleted goes on both.
        expect(runCommand(r.joiner, 'trackedChanges.acceptAll', undefined, { from: blockFrom(r.joiner, 4) }).ok).toBe(true);
        await converged(r);
      });

      test('two authors delete overlapping text: every character is deleted, once, by one of them (item 5)', async () => {
        const r = await overlappingDeletions(order);
        const pieces = revisionsIn(r.creator, 4);
        expect(pieces.every((p) => p[1] === 'del')).toBe(true);
        // The overlap, "te", is one author's: which one is Yjs's order, the same on both peers.
        expect(pieces[0]![2]).toBe('Author A');
        expect(pieces[pieces.length - 1]![2]).toBe('Author B');
      });

      // Overlapping keys for the deletion mark (section 12.7 item 5 (a), at the core-ts 0.3.0 upgrade, section 12.26).
      test('two authors delete overlapping text: the overlap carries both deletions, listed apart; one rejected leaves the other; the save writes the earlier', async () => {
        const r = await overlappingDeletions(order);
        await converged(r);
        const deleted = (s: EditorDocument): [string, string[]][] => {
          const out: [string, string[]][] = [];
          s.doc.child(4).forEach((n) => { if (n.isText) out.push([n.text!, n.marks.filter((m) => m.type.name === 'deletion').map((m) => m.attrs['author'] as string).sort()]); });
          return out;
        };
        expect(deleted(r.creator)).toEqual([['Some ', ['Author A']], ['te', ['Author A', 'Author B']], ['xt.', ['Author B']]]);
        const changes = trackedChangesOf(r.joiner.doc).filter((c) => c.type === 'Deleted' && (c.text.includes('Some') || c.text.includes('xt.')));
        expect(changes.map((c) => [c.author, c.text]).sort()).toEqual([['Author A', 'Some te'], ['Author B', 'text.']]);
        // A's rejected: "te" stays deleted, B's.
        const a = changes.find((c) => c.author === 'Author A')!;
        expect(runCommand(r.joiner, 'trackedChange.reject', { key: r.joiner.changeKey(a) }).ok).toBe(true);
        await converged(r);
        expect(deleted(r.creator)).toEqual([['Some ', []], ['text.', ['Author B']]].map(([t, w]) => [t, w]) as [string, string[]][]);
      });

      test('the save of two overlapping deletions writes the earlier over the overlap, the same bytes on both peers', async () => {
        const r = await trackingRoom();
        r.creator.trackedChangeDate = new Date('2026-09-29T01:00:00Z');
        r.joiner.trackedChangeDate = new Date('2026-09-29T02:00:00Z');
        const b = blockFrom(r.creator, 4);
        r.link.hold();
        expect(runCommand(r.creator, 'insertText', { text: '' }, { from: b, to: b + 7 }).ok).toBe(true);
        expect(runCommand(r.joiner, 'insertText', { text: '' }, { from: b + 5, to: b + 10 }).ok).toBe(true);
        r.link.release(order);
        await converged(r);
        const saved = await r.joiner.save();
        const xml = new TextDecoder().decode(new ZipPartStore(saved.bytes).loadSync('word/document.xml'));
        // "Some te" A's (the earlier), "xt." B's.
        expect(xml).toMatch(/<w:del w:id="\d+"[^>]* w:author="Author A"[^>]*><w:r><w:delText>Some te<\/w:delText><\/w:r><\/w:del><w:del w:id="\d+"[^>]* w:author="Author B"[^>]*><w:r><w:delText>xt\.<\/w:delText><\/w:r><\/w:del>/);
      });

      test('two authors type at the same point: both insertions kept, each its author\'s', async () => {
        const r = await trackingRoom();
        const b = blockFrom(r.creator, 4);
        r.link.hold();
        expect(runCommand(r.creator, 'insertText', { text: 'alpha ' }, { from: b + 5 }).ok).toBe(true);
        expect(runCommand(r.joiner, 'insertText', { text: 'beta ' }, { from: b + 5 }).ok).toBe(true);
        r.link.release(order);
        await converged(r);
        const pieces = revisionsIn(r.creator, 4).filter((p) => p[1] === 'ins');
        expect(pieces.map((p) => [p[0], p[2]]).sort()).toEqual([['alpha ', 'Author A'], ['beta ', 'Author B']]);
      });

      for (const action of ['accept', 'reject'] as const) {
        test(`one author ${action}s an insertion the other is typing into: the peers converge, and the typing is not lost`, async () => {
          const r = await resolvingWhileTyping(order, action);
          await converged(r);
          expect(r.creator.doc.child(4).textContent).toContain('more ');
        });

        // What the resolver saw is resolved; what it did not see, typed at the same time, stays the typist's, pending (the room repair, ED-005 section 12.15).
        test(`one author ${action}s an insertion the other is typing into: the typing stays pending, its author's (item 5)`, async () => {
          const r = await resolvingWhileTyping(order, action);
          const pieces = revisionsIn(r.creator, 4);
          expect(pieces.filter((p) => p[1] === 'ins')).toEqual([['more ', 'ins', 'Author A']]);
          expect(pieces.map((p) => p[0]).join('')).toBe(action === 'accept' ? 'Some new more words text.' : 'Some more text.');
        });
      }
    });
  }
});

describe('a tree run over a table merged irregular (ED-004 section 6.1)', () => {
  test('the run exports the table repaired and its change carries the repair into the store, on both peers', async () => {
    const r = await room(await fixture('tables.docx'));
    r.link.hold();
    expect(runCommand(r.creator, 'table.addColumns', { location: 'After' }, { from: cellStart(r.creator, 1, 0, 0) }).ok).toBe(true);
    expect(runCommand(r.joiner, 'table.addRows', { location: 'After' }, { from: cellStart(r.joiner, 1, 1, 0) }).ok).toBe(true);
    r.link.release('aFirst');
    await settled(r.creator, r.joiner);
    expect(structural(r.creator.doc)[0]).toBeGreaterThan(0);
    const run = await r.creator.runOnTree(({ body }) => { body.paragraphs[0]!.insertText(' (a script)', 'End'); });
    expect(run.error).toBeUndefined();
    await settled(r.creator, r.joiner);
    for (const s of [r.creator, r.joiner]) expect(structural(s.doc)[0]).toBe(0);
    await converged(r);
  });
});

/**
 * The agent's call, as the companion's `mutate` makes it (ED-005 section 9.3): a tree run tracked as
 * Claude that replaces `from` with `to` in the first paragraph holding it, guarded or not. `during`
 * runs inside the callback, after the edit: what arrives while a call runs (the drain lets it in).
 */
async function agentReplaces(s: EditorDocument, from: string, to: string, guard: boolean, during?: () => void): Promise<Awaited<ReturnType<EditorDocument['runOnTree']>>> {
  return s.runOnTree(({ body }) => {
    const found = body.search(from) as unknown as { items?: { insertText(t: string, where: string): unknown }[] } | { insertText(t: string, where: string): unknown }[];
    const hits = Array.isArray(found) ? found : found.items!;
    hits[0]!.insertText(to, 'Replace');
    during?.();
  }, 'agent', { author: 'Claude', mode: 'TrackAll' }, { guard });
}

/** Block `i`'s text with its revisions, as `revisionsIn` pieces joined: "text[ins:Author]" and "text[del:Author]". */
function marked(s: EditorDocument, i: number): string {
  return revisionsIn(s, i).map(([t, k, a]) => (k ? `${t}[${k}:${a}]` : t)).join('');
}

// ED-005 proposal 42, section 9.3 item 1 (ii): an agent's call that meets the person's typing in a block it replaces is
// refused, and writes nothing; the agent reads again and redoes it (the companion's side is its own suite). A script's
// run keeps E3's rule. What stays is typing still in flight when the call's change is sent: (iii), a stated limitation.
describe('an agent\'s call and a person typing in its paragraph (ED-005 proposal 42)', () => {
  for (const mode of ['TrackAll', 'Off'] as const) {
    describe(`the document's tracking ${mode === 'TrackAll' ? 'on' : 'off'}`, () => {
      async function roomWith(): Promise<Room> {
        const r = await room(await fixture('tracked-changes.docx'));
        if (mode === 'TrackAll') expect((await runCommandAsync(r.creator, 'document.setChangeTrackingMode', { value: 'TrackAll' })).ok).toBe(true);
        await settled(r.creator, r.joiner);
        r.joiner.trackingAuthor = 'Author B';
        return r;
      }
      const typed = mode === 'TrackAll' ? 'Some new [ins:Author B]text.' : 'Some new text.';

      test('the typing reaches the call\'s paragraph while it runs: the call is refused, nothing is written, the words stay the person\'s, and the agent\'s session is spent', async () => {
        const r = await roomWith();
        const b = blockFrom(r.joiner, 4);
        r.link.hold();
        expect(runCommand(r.joiner, 'insertText', { text: 'new ' }, { from: b + 5 }).ok).toBe(true);
        const run = await agentReplaces(r.creator, 'Some text.', 'Other words.', true, () => r.link.release('aFirst'));
        expect(run.refused?.blocks).toEqual([{ before: 4, text: 'Some new text.' }]);
        await settled(r.creator, r.joiner);
        expect(marked(r.joiner, 4)).toBe(typed);
        expect(r.creator.store.ydoc.getXmlFragment('body').toArray().length).toBe(r.joiner.doc.childCount);
        expect(trackedChangesOf(r.joiner.doc).filter((c) => c.author === 'Claude')).toEqual([]);
        await expect(r.creator.runOnTree(() => undefined)).rejects.toThrow(/spent/);
        await expect(r.creator.save()).rejects.toThrow(/spent/);
      });

      test('typing elsewhere while the call runs is merged: not refused, both kept', async () => {
        const r = await roomWith();
        r.link.hold();
        expect(runCommand(r.joiner, 'insertText', { text: 'Typed ' }, { from: blockFrom(r.joiner, 1) }).ok).toBe(true);
        const run = await agentReplaces(r.creator, 'Some text.', 'Other words.', true, () => r.link.release('aFirst'));
        expect(run.refused).toBeUndefined();
        await converged(r);
        expect(r.creator.doc.child(1).textContent.startsWith('Typed ')).toBe(true);
        expect(trackedChangesOf(r.creator.doc).filter((c) => c.author === 'Claude').map((c) => [c.type, c.text])).toEqual([['Added', 'Other words.'], ['Deleted', 'Some text.']]);   // the engine's tracked Replace writes the insertion first (core-ts 0.6.0, CR-014 phase 3; ED-008 section 8.8 item 9)
      });

      test('a script\'s run meets the same typing and keeps E3\'s rule: the run\'s block in place of the person\'s (ED-004 section 10.2)', async () => {
        const r = await roomWith();
        const b = blockFrom(r.joiner, 4);
        r.link.hold();
        expect(runCommand(r.joiner, 'insertText', { text: 'new ' }, { from: b + 5 }).ok).toBe(true);
        const run = await agentReplaces(r.creator, 'Some text.', 'Other words.', false, () => r.link.release('aFirst'));
        expect(run.refused).toBeUndefined();
        await converged(r);
        expect(r.creator.doc.child(4).textContent).not.toContain('new ');
      });

      // Found writing the cases above (ED-005 section 12.36): the person's save of the agent's changes, the agent's
      // blocks adopted through `sources` with their revisions as the agent's engine wrote them.
      test('the person\'s save of the agent\'s changes declares w16du, and its own revisions take ids above the agent\'s', async () => {
        const r = await roomWith();
        await agentReplaces(r.creator, 'Some text.', 'Other words.', true);
        await settled(r.creator, r.joiner);
        if (mode === 'TrackAll') expect(runCommand(r.joiner, 'insertText', { text: 'Typed ' }, { from: blockFrom(r.joiner, 1) }).ok).toBe(true);
        // Both peers' first saves the same bytes; then the person's, checked on its own (a session numbers what it writes afresh at each save).
        await converged(r);
        const xml = new TextDecoder().decode(new ZipPartStore((await r.joiner.save()).bytes).loadSync('word/document.xml'));
        expect(xml).toMatch(/w16du:dateUtc="/);
        expect(/mc:Ignorable="([^"]*)"/.exec(xml)![1]!.split(' ')).toContain('w16du');
        const ids = [...xml.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => m[1]);
        expect(ids.length).toBeGreaterThanOrEqual(mode === 'TrackAll' ? 3 : 2);
        expect(new Set(ids).size).toBe(ids.length);
      });

      for (const order of RELEASE_ORDERS) {
        // (iii): typing in flight when the call's change is sent is concurrent in Yjs's sense, which the guard cannot see
        // (UPSTREAM-10). What it should do is keep the typing the person's, undeleted; what it does is give it Claude's
        // deletion. A stated limitation (docs/user/review.md, upstream item 10), kept here as the outcome it should have.
        test.fails(`released ${order}: typing in flight when the call's change is sent stays the person's, undeleted (a stated limitation)`, async () => {
          const r = await roomWith();
          const b = blockFrom(r.joiner, 4);
          r.link.hold();
          const run = await agentReplaces(r.creator, 'Some text.', 'Other words.', true);
          expect(run.refused).toBeUndefined();
          expect(runCommand(r.joiner, 'insertText', { text: 'new ' }, { from: b + 5 }).ok).toBe(true);
          r.link.release(order);
          await converged(r);
          // The typist's words: no deletion on them (with tracking on they are the typist's insertion, and should stay only that).
          const deleted: boolean[] = [];
          r.creator.doc.child(4).forEach((n) => { if (n.isText && n.text!.includes('new')) deleted.push(n.marks.some((m) => m.type.name === 'deletion')); });
          expect(deleted.length).toBeGreaterThan(0);
          expect(deleted.some((d) => d)).toBe(false);
        });
      }
    });
  }
});

describe('undo across peers (ED-004 section 6.2)', () => {
  test('A types, B types, A undoes: only A\'s typing goes', async () => {
    const r = await room(await fixture('tracked-changes.docx'));
    const a = r.creator;
    const b = r.joiner;
    expect(runCommand(a, 'insertText', { text: 'Mine. ' }, { from: blockFrom(a, 1) }).ok).toBe(true);
    a.undoManager.stopCapturing();
    expect(runCommand(b, 'insertText', { text: 'Theirs. ' }, { from: blockFrom(b, 3) }).ok).toBe(true);
    a.undoManager.undo();
    for (const s of [a, b]) {
      expect(s.doc.child(1).textContent.startsWith('Mine. ')).toBe(false);
      expect(s.doc.child(3).textContent.startsWith('Theirs. ')).toBe(true);
    }
    await converged(r);
  });

  test('A starts a list, B types in it, A undoes the start: the paragraph leaves the list, B\'s text stays, the definition stays', async () => {
    const r = await room(await fixture('tracked-changes.docx'));
    const a = r.creator;
    const b = r.joiner;
    expect((await runCommandAsync(a, 'list.toggle', { bullet: true }, { from: blockFrom(a, 4) })).ok).toBe(true);
    await settled(a, b);
    const numId = (a.doc.child(4).attrs['numPr'] as { numId: string }).numId;
    a.undoManager.stopCapturing();
    expect(runCommand(b, 'insertText', { text: 'In the list. ' }, { from: blockFrom(b, 4) }).ok).toBe(true);
    a.undoManager.undo();
    await settled(a, b);
    for (const s of [a, b]) {
      expect(s.doc.child(4).attrs['numPr']).toBeNull();
      expect(s.doc.child(4).textContent.startsWith('In the list. ')).toBe(true);
      const nums = ((await s.main.numberingDefinitionsPart!.readContents()).num ?? []).map((n) => String(n.numId));
      expect(nums).toContain(numId);
    }
    await converged(r);
  });

  test('an undo never touches another peer\'s part edit', async () => {
    const r = await room(await fixture('opendope/binding-simple.docx'));
    const a = r.creator;
    const b = r.joiner;
    expect(runCommand(a, 'insertText', { text: 'Mine. ' }, { from: blockFrom(a, 0) }).ok).toBe(true);
    expect(await b.ensureStyles(['FootnoteText'])).toBe(true);
    await settled(a, b);
    a.undoManager.undo();
    await settled(a, b);
    expect((await a.main.styleDefinitionsPart!.readContents()).style!.some((s) => s.styleId === 'FootnoteText')).toBe(true);
    expect(a.doc.child(0).textContent.startsWith('Mine. ')).toBe(false);
    await converged(r);
  });
});

// The room repair (ED-005 section 12.15): its other paths, in every release order. Each ends converged,
// and with nothing left for a repair to write on either peer (its termination).
describe('the room repair (ED-005 section 12.15)', () => {
  const nothingLeft = (r: Room): void => {
    expect(repairWrites(r.creator.store)).toEqual([]);
    expect(repairWrites(r.joiner.store)).toEqual([]);
  };

  for (const order of RELEASE_ORDERS) {
    describe(`released ${order}`, () => {
      test('the typing a minute later has markers of its own: accepted, the text after it stays accepted', async () => {
        const r = await trackingRoom();
        const b = blockFrom(r.creator, 4);
        r.creator.trackedChangeDate = new Date('2026-09-28T10:00:00Z');
        expect(runCommand(r.creator, 'insertText', { text: 'new words ' }, { from: b + 5 }).ok).toBe(true);
        await settled(r.creator, r.joiner);
        const [change] = trackedChangesOf(r.joiner.doc).filter((c) => c.author === 'Author A');
        const key = r.joiner.changeKey(change!);
        // A later minute: the typing is an insertion of its own, with markers of its own, whose restoring marker would put the accepted one back after it.
        r.creator.trackedChangeDate = new Date('2026-09-28T10:05:00Z');
        r.link.hold();
        expect(runCommand(r.creator, 'insertText', { text: 'more ' }, { from: b + 9 }).ok).toBe(true);
        expect(runCommand(r.joiner, 'trackedChange.accept', { key }, { from: b }).ok).toBe(true);
        r.link.release(order);
        await settled(r.creator, r.joiner);
        await converged(r);
        const pieces = revisionsIn(r.creator, 4);
        expect(pieces.filter((p) => p[1] === 'ins')).toEqual([['more ', 'ins', 'Author A']]);
        expect(pieces.map((p) => p[0]).join('')).toBe('Some new more words text.');
        nothingLeft(r);
      });

      test('typing after the resolution reached the typist is its own: untracked, it stays plain', async () => {
        const r = await resolvingWhileTyping(order, 'accept');
        await converged(r);
        // Tracking off on the typist's side only, and text typed inside the accepted range once the accept has arrived.
        r.creator.trackingMode = 'Off';
        const b = blockFrom(r.creator, 4);
        expect(runCommand(r.creator, 'insertText', { text: 'late ' }, { from: b + 5 }).ok).toBe(true);
        await settled(r.creator, r.joiner);
        await converged(r);
        const pieces = revisionsIn(r.creator, 4);
        expect(pieces.filter((p) => p[1] === 'ins')).toEqual([['more ', 'ins', 'Author A']]);
        expect(pieces.map((p) => p[0]).join('')).toBe('Some late new more words text.');
        nothingLeft(r);
      });

      // A limitation (ED-005 section 12.15 item 5): the repair's marker after the typing it marks again restores
      // the value there at the time, accepted, and it is a peer's write, which the resolver's undo does not take
      // back. So after the undo, the text after the typing stays accepted. The record goes (asserted first), and
      // nothing is accepted again; the expectation is kept.
      test.fails('the resolver undoes the accept: its record goes with it, and the insertion is back whole, not accepted again', async () => {
        const r = await resolvingWhileTyping(order, 'accept');
        await converged(r);
        expect(r.joiner.store.resolutions.size).toBe(1);
        expect(runCommand(r.joiner, 'history.undo').ok).toBe(true);
        await settled(r.creator, r.joiner);
        await converged(r);
        expect(r.creator.store.resolutions.size).toBe(0);
        const pieces = revisionsIn(r.creator, 4);
        expect(pieces.filter((p) => p[1] === 'ins')).toEqual([['new more words ', 'ins', 'Author A']]);
        nothingLeft(r);
      });

      test('a deletion rejected while another author types inside it: the typing stays theirs, and the rest is not deleted again', async () => {
        const r = await trackingRoom();
        const b = blockFrom(r.creator, 4);
        // B deletes "Some ", seen by both.
        expect(runCommand(r.joiner, 'insertText', { text: '' }, { from: b, to: b + 5 }).ok).toBe(true);
        await settled(r.creator, r.joiner);
        const [change] = trackedChangesOf(r.joiner.doc).filter((c) => c.type === 'Deleted' && c.author === 'Author B');
        const key = r.joiner.changeKey(change!);
        r.link.hold();
        // A types inside B's deletion, "So|me ": A's insertion splits it, with markers of its own; B rejects the deletion.
        expect(runCommand(r.creator, 'insertText', { text: 'X' }, { from: b + 2 }).ok).toBe(true);
        expect(runCommand(r.joiner, 'trackedChange.reject', { key }, { from: b }).ok).toBe(true);
        r.link.release(order);
        await settled(r.creator, r.joiner);
        await converged(r);
        const pieces = revisionsIn(r.creator, 4);
        expect(pieces.filter((p) => p[1] === 'del')).toEqual([]);
        expect(pieces.filter((p) => p[1] === 'ins')).toEqual([['X', 'ins', 'Author A']]);
        expect(pieces.map((p) => p[0]).join('')).toBe('SoXme text.');
        nothingLeft(r);
      });
    });
  }
});

// ED-005 section 9.3 items 2 to 6 (E4.g step G2): the agent beside a person's typing, two people resolving one change,
// "all" racing an edit, Replace all racing typing and a resolution, and two pages applying one part. Each case in the
// four release orders, ending converged (the same document, the same bytes, the self-check passing) and, where a
// resolution is involved, with nothing left for the room repair to write.
describe('concurrency in a room with tracking (ED-005 section 9.3 items 2 to 6)', () => {
  const nothingLeft = (r: Room): void => {
    expect(repairWrites(r.creator.store)).toEqual([]);
    expect(repairWrites(r.joiner.store)).toEqual([]);
  };
  /** The first change by `author` of `type` whose text holds `text`, as `s` lists it. */
  const changeOf = (s: EditorDocument, author: string, type: string, text: string) => {
    const c = trackedChangesOf(s.doc).find((x) => x.author === author && x.type === type && x.text.includes(text));
    if (!c) throw new Error(`no ${type} by ${author} holding "${text}"`);
    return c;
  };
  /** Whether the text nodes holding `text` in block `i` carry a deletion. */
  const deletedIn = (s: EditorDocument, i: number, text: string): boolean => {
    let found = false;
    s.doc.child(i).forEach((n) => { if (n.isText && n.text!.includes(text) && n.marks.some((m) => m.type.name === 'deletion')) found = true; });
    return found;
  };

  for (const order of RELEASE_ORDERS) {
    describe(`released ${order}`, () => {
      // Item 2: the agent's call beside the person's typing in flight (concurrent in Yjs's sense). The typing arrives
      // after the call ended, so the guard (item 1) does not see it; neither edit reaches the other's text.
      test('item 2: the agent inserts a paragraph after the one the person is typing in: both kept, each its author\'s', async () => {
        const r = await trackingRoom();
        r.link.hold();
        const run = await r.creator.runOnTree(({ body }) => {
          const found = body.search('Some text.') as unknown as { items?: { insertParagraph(t: string, w: string): unknown }[] } | { insertParagraph(t: string, w: string): unknown }[];
          (Array.isArray(found) ? found : found.items!)[0]!.insertParagraph('The agent\'s paragraph.', 'After');
        }, 'agent', { author: 'Claude', mode: 'TrackAll' }, { guard: true });
        expect(run.error).toBeUndefined();
        expect(runCommand(r.joiner, 'insertText', { text: 'new ' }, { from: blockFrom(r.joiner, 4) + 5 }).ok).toBe(true);
        r.link.release(order);
        await converged(r);
        expect(r.creator.doc.child(4).textContent).toBe('Some new text.');
        expect(deletedIn(r.creator, 4, 'new')).toBe(false);
        expect(changeOf(r.creator, 'Claude', 'Added', 'The agent\'s paragraph.')).toBeDefined();
      });

      test('item 2: the agent replaces words in one part of a paragraph while the person types in another: both kept', async () => {
        const r = await trackingRoom();
        r.link.hold();
        const run = await agentReplaces(r.creator, 'basic', 'simple', true);
        expect(run.refused).toBeUndefined();
        expect(runCommand(r.joiner, 'insertText', { text: 'really ' }, { from: blockFrom(r.joiner, 1) + 8 }).ok).toBe(true);
        r.link.release(order);
        await converged(r);
        expect(r.creator.doc.child(1).textContent).toContain('This is really a document');
        expect(deletedIn(r.creator, 1, 'really')).toBe(false);
        expect(changeOf(r.creator, 'Claude', 'Deleted', 'basic')).toBeDefined();
        expect(changeOf(r.creator, 'Claude', 'Added', 'simple')).toBeDefined();
      });

      test('item 2: the agent comments on words while the person types beside them: both kept, the comment on each peer', async () => {
        const r = await trackingRoom();
        await r.creator.stories.comments();
        r.link.hold();
        const run = await r.creator.runOnTree(({ body }) => {
          const found = body.search('Some') as unknown as { items?: { insertComment(t: string): unknown }[] } | { insertComment(t: string): unknown }[];
          (Array.isArray(found) ? found : found.items!)[0]!.insertComment('The agent\'s comment.');
        }, 'agent', { author: 'Claude', mode: 'Off' }, { guard: true });
        expect(run.error).toBeUndefined();
        expect(runCommand(r.joiner, 'insertText', { text: 'new ' }, { from: blockFrom(r.joiner, 4) + 5 }).ok).toBe(true);
        r.link.release(order);
        await settled(r.creator, r.joiner);
        expect(r.joiner.doc.eq(r.creator.doc)).toBe(true);
        // The person's words kept, theirs and undeleted; where they land is the case below (section 12.37).
        expect(r.creator.doc.child(4).textContent.replace(/\s+/g, ' ')).toMatch(/^Some ?new text\.$/);
        expect(deletedIn(r.creator, 4, 'new')).toBe(false);
        // Comments compare by content, as E4.e's two peers do: a comment made in a room takes a new paragraph id and durable
        // id at each peer's save. Its date must agree (ED-005 section 12.37: the other peer's was the time zone's offset out).
        await r.joiner.stories.comments();
        const [a, b] = [(await r.creator.save()).bytes, (await r.joiner.save()).bytes];
        expect(comparePackages(a, b, ['word/comments.xml', 'word/commentsExtended.xml', 'word/commentsIds.xml'])).toEqual([]);
        const listed = async (bytes: Uint8Array) => {
          const s = await EditorDocument.open(bytes);
          return (await s.stories.comments()).map((t) => [t.author, s.stories.currentDoc(t).textContent, t.date?.toISOString()]);
        };
        const [la, lb] = [await listed(a), await listed(b)];
        expect(la).toEqual(lb);
        expect(la.map((c) => c[1])).toContain('The agent\'s comment.');
      });

      // Found writing item 2 (ED-005 section 12.37): y-prosemirror keeps a paragraph's text as runs of `Y.XmlText` split around
      // each inline node, so a peer that puts an atom (the agent's comment markers, a tab, a note reference) inside a run the
      // other is typing into splits that run, and the concurrent typing stays with its old anchor, at the first half's end:
      // "Somenew  text." for "Some new text.". The words are kept, the typist's, undeleted. Kept here as where they should land.
      test.fails('item 2: the typing lands where it was typed when the agent\'s comment splits its run (a stated limitation)', async () => {
        const r = await trackingRoom();
        await r.creator.stories.comments();
        r.link.hold();
        await r.creator.runOnTree(({ body }) => {
          const found = body.search('Some') as unknown as { items?: { insertComment(t: string): unknown }[] } | { insertComment(t: string): unknown }[];
          (Array.isArray(found) ? found : found.items!)[0]!.insertComment('The agent\'s comment.');
        }, 'agent', { author: 'Claude', mode: 'Off' }, { guard: true });
        expect(runCommand(r.joiner, 'insertText', { text: 'new ' }, { from: blockFrom(r.joiner, 4) + 5 }).ok).toBe(true);
        r.link.release(order);
        await settled(r.creator, r.joiner);
        expect(r.creator.doc.child(4).textContent).toBe('Some new text.');
      });

      // Item 3: two people resolve one change at once.
      for (const [kind, make, accepted, rejected] of [
        ['an insertion', (s: EditorDocument) => runCommand(s, 'insertText', { text: 'new words ' }, { from: blockFrom(s, 4) + 5 }), 'Some new words text.', 'Some text.'],
        ['a deletion', (s: EditorDocument) => runCommand(s, 'insertText', { text: '' }, { from: blockFrom(s, 4), to: blockFrom(s, 4) + 5 }), 'text.', 'Some text.'],
      ] as const) {
        for (const [first, second] of [['accept', 'reject'], ['accept', 'accept'], ['reject', 'reject']] as const) {
          test(`item 3: two people ${first} and ${second} ${kind} at once: the same outcome on both, one of the two, and the repair settles`, async () => {
            const r = await trackingRoom();
            expect(make(r.creator).ok).toBe(true);
            await settled(r.creator, r.joiner);
            const type = kind === 'an insertion' ? 'Added' : 'Deleted';
            const inA = changeOf(r.creator, 'Author A', type, kind === 'an insertion' ? 'new words' : 'Some');
            const inB = changeOf(r.joiner, 'Author A', type, kind === 'an insertion' ? 'new words' : 'Some');
            r.link.hold();
            expect(runCommand(r.creator, `trackedChange.${first}`, { key: r.creator.changeKey(inA) }).ok).toBe(true);
            expect(runCommand(r.joiner, `trackedChange.${second}`, { key: r.joiner.changeKey(inB) }).ok).toBe(true);
            r.link.release(order);
            await converged(r);
            const text = r.creator.doc.child(4).textContent;
            const allowed = first === second ? [first === 'accept' ? accepted : rejected] : [accepted, rejected];
            expect(allowed).toContain(text);
            expect(trackedChangesOf(r.creator.doc).filter((c) => c.author === 'Author A')).toEqual([]);
            nothingLeft(r);
          });
        }
      }

      for (const second of ['reject', 'accept'] as const) {
        test(`item 3: two people accept and ${second} a formatting change at once: the same formatting on both, no record left`, async () => {
          const r = await trackingRoom();
          const b = blockFrom(r.creator, 4);
          expect(runCommand(r.creator, 'font.bold', { value: true }, { from: b, to: b + 4 }).ok).toBe(true);
          await settled(r.creator, r.joiner);
          const inA = changeOf(r.creator, 'Author A', 'Formatted', 'Some');
          const inB = changeOf(r.joiner, 'Author A', 'Formatted', 'Some');
          r.link.hold();
          expect(runCommand(r.creator, 'trackedChange.accept', { key: r.creator.changeKey(inA) }).ok).toBe(true);
          expect(runCommand(r.joiner, `trackedChange.${second}`, { key: r.joiner.changeKey(inB) }).ok).toBe(true);
          r.link.release(order);
          await converged(r);
          expect(trackedChangesOf(r.creator.doc).filter((c) => c.author === 'Author A')).toEqual([]);
          // Accepted twice, the bold stays (ED-005 section 12.47, the case 12.37 left out).
          if (second === 'accept') {
            const marks = r.creator.doc.nodeAt(b)!.marks.map((m) => m.type.name);
            expect(marks).toContain('bold');
          }
          nothingLeft(r);
        });
      }

      // A move resolved the same way by both: y-prosemirror's diff had put a copy of a kept paragraph in the place of the
      // move's range marker beside it, and a copy of a paragraph's text in the place of its inline marker, once per peer,
      // so the copies survived the merge: the paragraph twice (two accepts), words run together (two rejects). The store's
      // children are now matched to the new ones before the diff (UPSTREAM-16, ED-005 section 12.47). Resolved opposite
      // ways, each goes its deleting way, as section 12.43 states for any two changes: the moved text is in neither place,
      // and each person's undo restores their own half (with the move's range markers twice: UPSTREAM-17).
      const moveOf = (s: EditorDocument): ReturnType<typeof trackedChangesOf>[number] => trackedChangesOf(s.doc).find((c) => c.move !== undefined)!;
      const movedAlone = async (bytes: Uint8Array, action: 'accept' | 'reject'): Promise<string[]> => {
        const s = await EditorDocument.open(bytes);
        expect(runCommand(s, `trackedChange.${action}`, { key: s.changeKey(moveOf(s)) }).ok).toBe(true);
        return blockTexts(s.doc);
      };
      for (const [first, second] of [['accept', 'accept'], ['reject', 'reject'], ['accept', 'reject']] as const) {
        test(`item 3: two people ${first} and ${second} a move at once, over Word's own move: ${first === second ? 'the document either left alone' : 'the moved text in neither place, and each undo restores its half'}`, async () => {
          const bytes = await fixture('revisions/revisions-word15.docx');
          const r = await room(bytes);
          await settled(r.creator, r.joiner);
          const move = moveOf(r.creator);
          const moveB = trackedChangesOf(r.joiner.doc).find((c) => c.move === move.move && c.type === move.type)!;
          const moved = move.text.replace(/\r$/, '');
          r.link.hold();
          expect(runCommand(r.creator, `trackedChange.${first}`, { key: r.creator.changeKey(move) }).ok).toBe(true);
          expect(runCommand(r.joiner, `trackedChange.${second}`, { key: r.joiner.changeKey(moveB) }).ok).toBe(true);
          r.link.release(order);
          await converged(r);
          expect(trackedChangesOf(r.creator.doc).filter((c) => c.move === move.move)).toEqual([]);
          if (first === second) {
            expect(blockTexts(r.creator.doc)).toEqual(await movedAlone(bytes, first));
            nothingLeft(r);
            return;
          }
          expect(blockTexts(r.creator.doc)).toEqual((await movedAlone(bytes, 'accept')).filter((t) => t !== moved));
          nothingLeft(r);
          const halves = (): string[] => trackedChangesOf(r.creator.doc).filter((c) => c.move === move.move).map((c) => c.type).sort();
          r.creator.undoManager.undo();
          await settled(r.creator, r.joiner);
          r.joiner.undoManager.undo();
          await settled(r.creator, r.joiner);
          expect(halves()).toEqual(['Added', 'Deleted']);
          expect(r.joiner.doc.eq(r.creator.doc)).toBe(true);
          expect(runCommand(r.creator, 'trackedChange.accept', { key: r.creator.changeKey(moveOf(r.creator)) }).ok).toBe(true);
          await converged(r);
          // The text as one accept leaves it. Both resolutions deleted the move's two range markers, and each undo put
          // them back, so they were there twice; the accept removes one of each and leaves the others and an empty
          // paragraph (UPSTREAM-17, stated in ED-005 section 12.47).
          const accepted = await movedAlone(bytes, 'accept');
          expect(blockTexts(r.joiner.doc).filter((t) => t !== '')).toEqual(accepted.filter((t) => t !== ''));
          expect(r.joiner.doc.childCount - accepted.length).toBe(3);
          nothingLeft(r);
        });
      }

      // Item 4: "all" racing an edit.
      // Proposal 41 (section 9.7): a replacement is one entry, resolved by the halves' keys in one transaction and one
      // room record. Two people resolve the same pair; or one resolves the pair while the other, having split it,
      // resolves one half, so the two records disagree on the extent. Converged either way, and the repair settles.
      const replaced = async (): Promise<Room> => {
        const r = await trackingRoom();
        const b = blockFrom(r.creator, 4);
        expect(runCommand(r.creator, 'insertText', { text: 'Many' }, { from: b, to: b + 4 }).ok).toBe(true);
        await settled(r.creator, r.joiner);
        expect(r.joiner.doc.child(4).textContent).toBe('SomeMany text.');
        return r;
      };
      const pairKeys = (s: EditorDocument): string[] => s.changeKeys([changeOf(s, 'Author A', 'Deleted', 'Some'), changeOf(s, 'Author A', 'Added', 'Many')]);

      for (const [first, second] of [['accept', 'reject'], ['reject', 'accept']] as const) {
        test(`proposal 41: two people ${first} and ${second} a replacement by its halves' keys at once: converged, neither text left, the repair settles`, async () => {
          const r = await replaced();
          r.link.hold();
          expect(runCommand(r.creator, `trackedChange.${first}`, { keys: pairKeys(r.creator) }).ok).toBe(true);
          expect(runCommand(r.joiner, `trackedChange.${second}`, { keys: pairKeys(r.joiner) }).ok).toBe(true);
          r.link.release(order);
          await converged(r);
          // Neither text: each change goes the way that deletes, as a single change resolved both ways does in item 3
          // (Yjs keeps a deletion over a formatting change), so the deletion's accept and the insertion's reject both
          // stand. The same in every order, and not the pair's doing (the next case). Stated (ED-005 section 12.43).
          expect(r.creator.doc.child(4).textContent).toBe(' text.');
          expect(trackedChangesOf(r.creator.doc).filter((c) => c.author === 'Author A')).toEqual([]);
          nothingLeft(r);
        });

        // The outcome wanted (ED-005 section 12.43, the guard specified there): one of the two a person chose, the
        // same on both peers, the winner by a symmetric rule (the lower Yjs client id). Pinned until the guard lands.
        test.fails(`proposal 41: two people ${first} and ${second} a replacement at once: one of their outcomes, the lower client id's (wanted, the guard of section 12.43)`, async () => {
          const r = await replaced();
          r.link.hold();
          expect(runCommand(r.creator, `trackedChange.${first}`, { keys: pairKeys(r.creator) }).ok).toBe(true);
          expect(runCommand(r.joiner, `trackedChange.${second}`, { keys: pairKeys(r.joiner) }).ok).toBe(true);
          r.link.release(order);
          await converged(r);
          const winner = r.creator.store.ydoc.clientID < r.joiner.store.ydoc.clientID ? first : second;
          expect(r.creator.doc.child(4).textContent).toBe(winner === 'accept' ? 'Many text.' : 'Some text.');
          nothingLeft(r);
        });
      }

      test('proposal 41: not the pair\'s doing: two adjacent changes resolved opposite ways, one command each, lose both texts too', async () => {
        const r = await replaced();
        r.link.hold();
        for (const type of ['Deleted', 'Added'] as const) {
          const text = type === 'Deleted' ? 'Some' : 'Many';
          expect(runCommand(r.creator, 'trackedChange.accept', { key: r.creator.changeKey(changeOf(r.creator, 'Author A', type, text)) }).ok).toBe(true);
          expect(runCommand(r.joiner, 'trackedChange.reject', { key: r.joiner.changeKey(changeOf(r.joiner, 'Author A', type, text)) }).ok).toBe(true);
        }
        r.link.release(order);
        await converged(r);
        expect(r.creator.doc.child(4).textContent).toBe(' text.');
        nothingLeft(r);
      });

      // The recovery the guide gives (ED-005 section 12.43): each person's undo restores their own half, with its mark;
      // both undone, the replacement is pending again on both peers, and resolving it again works.
      for (const [first, second] of [['accept', 'reject'], ['reject', 'accept']] as const) {
        test(`proposal 41: after the race (${first} against ${second}), each person's undo restores their own half, and the pair resolves again`, async () => {
          const r = await replaced();
          r.link.hold();
          expect(runCommand(r.creator, `trackedChange.${first}`, { keys: pairKeys(r.creator) }).ok).toBe(true);
          expect(runCommand(r.joiner, `trackedChange.${second}`, { keys: pairKeys(r.joiner) }).ok).toBe(true);
          r.link.release(order);
          await converged(r);
          const pending = (): string[] => trackedChangesOf(r.creator.doc).filter((c) => c.author === 'Author A').map((c) => `${c.type}:${c.text}`);
          r.creator.undoManager.undo();
          await settled(r.creator, r.joiner);
          // The creator's own resolution removed "Some" (an accept) or "Many" (a reject): that one comes back, marked.
          expect(pending()).toEqual(first === 'accept' ? ['Deleted:Some'] : ['Added:Many']);
          r.joiner.undoManager.undo();
          await settled(r.creator, r.joiner);
          expect(pending()).toEqual(['Deleted:Some', 'Added:Many']);
          expect(r.joiner.doc.eq(r.creator.doc)).toBe(true);
          expect(runCommand(r.creator, 'trackedChange.accept', { keys: pairKeys(r.creator) }).ok).toBe(true);
          await converged(r);
          expect(r.joiner.doc.child(4).textContent).toBe('Many text.');
          nothingLeft(r);
        });
      }

      for (const [pair, half] of [['accept', 'reject'], ['reject', 'accept']] as const) {
        test(`proposal 41: one person ${pair}s the replacement while the other ${half}s its deletion alone: converged, the repair settles`, async () => {
          const r = await replaced();
          r.link.hold();
          expect(runCommand(r.creator, `trackedChange.${pair}`, { keys: pairKeys(r.creator) }).ok).toBe(true);
          expect(runCommand(r.joiner, `trackedChange.${half}`, { key: r.joiner.changeKey(changeOf(r.joiner, 'Author A', 'Deleted', 'Some')) }).ok).toBe(true);
          r.link.release(order);
          await converged(r);
          // Whatever the order, no deletion is left pending: both resolvers settled it, one way or the other.
          expect(trackedChangesOf(r.creator.doc).filter((c) => c.author === 'Author A' && c.type === 'Deleted')).toEqual([]);
          nothingLeft(r);
        });
      }

      test('item 4: Accept all races an insertion elsewhere: the insertion stays pending, the rest accepted', async () => {
        const r = await trackingRoom();
        expect(runCommand(r.creator, 'insertText', { text: 'one ' }, { from: blockFrom(r.creator, 4) + 5 }).ok).toBe(true);
        await settled(r.creator, r.joiner);
        r.link.hold();
        expect((await runCommandAsync(r.creator, 'trackedChanges.acceptAll', undefined, { from: 1 })).ok).toBe(true);
        expect(runCommand(r.joiner, 'insertText', { text: 'two ' }, { from: blockFrom(r.joiner, 1) }).ok).toBe(true);
        r.link.release(order);
        await converged(r);
        expect(trackedChangesOf(r.creator.doc).map((c) => [c.author, c.type, c.text])).toEqual([['Author B', 'Added', 'two ']]);
        nothingLeft(r);
      });

      // Item 5: Replace all racing typing inside a hit, and racing a resolution of a change that holds a hit.
      test('item 5: Replace all races typing inside a hit: nothing lost, converged', async () => {
        const r = await trackingRoom();
        r.link.hold();
        expect(runCommand(r.creator, 'find.replaceAll', { text: 'text', replacement: 'words', matchWholeWord: true }).ok).toBe(true);
        expect(runCommand(r.joiner, 'insertText', { text: 'X' }, { from: blockFrom(r.joiner, 4) + 7 }).ok).toBe(true);
        r.link.release(order);
        await converged(r);
        expect(r.creator.doc.child(4).textContent).toContain('X');
        expect(r.creator.doc.child(4).textContent).toContain('words');
      });

      test('item 5: Replace all races the accept of an insertion holding a hit: converged, the repair settles', async () => {
        const r = await trackingRoom();
        expect(runCommand(r.creator, 'insertText', { text: 'plain text ' }, { from: blockFrom(r.creator, 4) }).ok).toBe(true);
        await settled(r.creator, r.joiner);
        const inB = changeOf(r.joiner, 'Author A', 'Added', 'plain text');
        r.link.hold();
        expect(runCommand(r.joiner, 'trackedChange.accept', { key: r.joiner.changeKey(inB) }).ok).toBe(true);
        expect(runCommand(r.creator, 'find.replaceAll', { text: 'text', replacement: 'words', matchWholeWord: true }).ok).toBe(true);
        r.link.release(order);
        await converged(r);
        expect(r.creator.doc.child(4).textContent).toContain('words');
        nothingLeft(r);
      });

      // Item 6: two pages apply the same part.
      test('item 6: two pages change the same style at once: one colour wins, the same on both, the saves the same bytes', async () => {
        const r = await trackingRoom();
        const styles = (s: EditorDocument) => s.pkg.parts.get('/word/styles.xml')!;
        const xmlOf = async (s: EditorDocument) => (styles(s) as unknown as { getXml(): Promise<string> }).getXml();
        const [xa, xb] = [withColor(await xmlOf(r.creator), 'Heading1', 'FF0000'), withColor(await xmlOf(r.joiner), 'Heading1', '0000FF')];
        expect(await checkPartXml(styles(r.creator), xa)).toEqual([]);
        r.link.hold();
        await applyPartXml(r.creator, styles(r.creator), xa);
        await applyPartXml(r.joiner, styles(r.joiner), xb);
        r.link.release(order);
        await settled(r.creator, r.joiner);
        const colour = async (s: EditorDocument) => (await s.main.styleDefinitionsPart!.readContents()).style!.find((x) => x.styleId === 'Heading1')!.rPr?.color?.val;
        const [ca, cb] = [await colour(r.creator), await colour(r.joiner)];
        expect(ca).toBe(cb);
        expect(['FF0000', '0000FF']).toContain(ca);
        await converged(r);
      });

      test('item 6: two pages change the same numbering format at once: one wins, the same on both', async () => {
        const r = await trackingRoom();
        const numbering = (s: EditorDocument) => s.pkg.parts.get('/word/numbering.xml')!;
        const xmlOf = async (s: EditorDocument) => (numbering(s) as unknown as { getXml(): Promise<string> }).getXml();
        const [xa, xb] = [(await xmlOf(r.creator)).replace(/w:numFmt w:val="decimal"/g, 'w:numFmt w:val="upperRoman"'), (await xmlOf(r.joiner)).replace(/w:numFmt w:val="decimal"/g, 'w:numFmt w:val="lowerRoman"')];
        r.link.hold();
        await applyPartXml(r.creator, numbering(r.creator), xa);
        await applyPartXml(r.joiner, numbering(r.joiner), xb);
        r.link.release(order);
        await settled(r.creator, r.joiner);
        const [na, nb] = [await xmlOf(r.creator), await xmlOf(r.joiner)];
        expect(na).toBe(nb);
        expect(/upperRoman|lowerRoman/.test(na)).toBe(true);
        await converged(r);
      });
    });
  }
});

/** A style's colour set: its own `w:color` replaced, or one added at the start of its run properties. */
function withColor(xml: string, id: string, color: string): string {
  const style = new RegExp(`<w:style [^>]*w:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`).exec(xml)![0];
  const own = /<w:color [^>]*\/>/;
  const changed = own.test(style) ? style.replace(own, `<w:color w:val="${color}"/>`) : style.replace('<w:rPr>', `<w:rPr><w:color w:val="${color}"/>`);
  return xml.replace(style, changed);
}

describe('a kept mc:AlternateContent in a room (ED-009 section 3.6)', () => {
  test('a peer\'s typing in the selected branch reaches the other peer, and both write it into the Choice with the Fallback as it was', async () => {
    const r = await room(await fixture('mc/mc-branches.docx'));
    const from = blockFrom(r.creator, 2) + 4;
    expect(runCommand(r.creator, 'insertText', { text: 'XX' }, { from }).ok).toBe(true);
    await settled(r.creator, r.joiner);
    expect(r.joiner.doc.child(2).textContent).toBe('The XXEarth is flat');
    for (const s of [r.creator, r.joiner]) {
      const result = await s.save();
      expect(result.selfCheck?.equal, JSON.stringify(result.selfCheck?.difference)).toBe(true);
      const xml = await mainXml(result.bytes);
      expect(xml).toMatch(/<mc:Choice Requires="wps"><w:t[^>]*>The XXEarth is flat<\/w:t><\/mc:Choice><mc:Fallback><w:t xml:space="preserve">The Earth is round<\/w:t><\/mc:Fallback>/);
      expect(xml.split('<mc:AlternateContent').length - 1).toBe(12);
    }
  });
});
