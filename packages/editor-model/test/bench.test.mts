// SPDX-License-Identifier: GPL-3.0-or-later
//
// The benchmark in Node (ED-002 section 7.5; ED-001 decision 10): the 255-page document under
// `fixtures/large`, opened, saved with the self-check, its headers read; the footnote panel's
// model work timed over `footnote.docx`. Printed, never asserted: the tests pass whatever the
// numbers say. SuperDoc's two seconds for 400 pages (ED-001 Appendix C) is the aspiration.
import { beforeAll, describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { EditorDocument } from '../src/index.mjs';
import { fixture, fixturesDir } from './helpers.mjs';

const LARGE = join(fixturesDir, 'large', 'anon-250.docx');

async function timed<T>(label: string, fn: () => Promise<T>): Promise<{ result: T; ms: number }> {
  const start = performance.now();
  const result = await fn();
  const ms = Math.round(performance.now() - start);
  console.log(`benchmark ${label}: ${ms} ms`);
  return { result, ms };
}

describe('benchmark (reported, never gating)', () => {
  test('anon-250.docx: open, save with the self-check, the headers', async () => {
    const bytes = new Uint8Array(await readFile(LARGE));
    const open = await timed('open (255 pages)', () => EditorDocument.open(bytes));
    const session = open.result;
    let blocks = 0;
    let paragraphs = 0;
    session.doc.descendants((n) => { if (n.type.name === 'paragraph') paragraphs++; return true; });
    blocks = session.doc.childCount;
    console.log(`benchmark document: ${blocks} top-level blocks, ${paragraphs} paragraphs, ${session.diagnostics.pageBreakMarkers} page-break markers, ${session.diagnostics.opaque.length} opaque elements`);
    const headers = await timed('headers and footers read and projected', () => session.stories.headers());
    expect(headers.result.length).toBe(3);
    const save = await timed('save with the self-check', () => session.save());
    expect(save.result.changed).toBe(false);
    const aspiration = 2000;
    console.log(`benchmark aspiration (ED-001 Appendix C): ${aspiration} ms to open 400 pages; open took ${open.ms} ms for 255`);
  }, 120_000);

  test('anon-250.docx: the text compared with its own bytes, as the disk-offer banner does (ED-011 section 3.4)', async () => {
    const { paragraphTextsOfSession, paragraphTextsOfBytes, compareText } = await import('../src/index.mjs');
    const bytes = new Uint8Array(await readFile(LARGE));
    const session = await EditorDocument.open(bytes);
    const live = await timed('compare: the live side read through readTree over a copy (255 pages)', () => paragraphTextsOfSession(session));
    const disk = await timed('compare: the disk side loaded and read (255 pages)', () => paragraphTextsOfBytes(bytes));
    const diff = await timed('compare: compareText over the two lists', async () => compareText(live.result, disk.result));
    console.log(`benchmark compare: ${live.result.length} paragraphs, ${diff.result.changed} changed, ${diff.result.added} added, ${diff.result.removed} removed; the live side ${live.ms} ms, the disk side ${disk.ms} ms, the diff ${diff.ms} ms`);
    expect(diff.result.changed + diff.result.added + diff.result.removed).toBe(0);
  }, 120_000);

  test('anon-250.docx: find "um" over every block, then after a keystroke (ED-005 section 3.6)', async () => {
    const { Finder, stateFor } = await import('../src/index.mjs');
    const session = await EditorDocument.open(new Uint8Array(await readFile(LARGE)));
    const state = stateFor(session);
    // "um" is as common in the anonymised document (3,636 times, case aside) as "the" was in the specification it stands in for (3,772).
    const finder = new Finder('um', {}, session.ctx.ids);
    const first = await timed('find "um" (255 pages)', async () => finder.all(state.doc));
    const blocks = finder.misses;
    const next = state.apply(state.tr.insertText('x', 3));
    const again = await timed('find "um" again after a keystroke', async () => finder.all(next.doc));
    console.log(`benchmark find: ${first.result.length} hits in ${blocks} blocks; after the keystroke ${finder.misses - blocks} block searched again, ${again.result.length} hits`);
    expect(again.result.length).toBe(first.result.length);
  }, 120_000);

  test('anon-250.docx: find across paragraph marks (a wildcard, then ^p) over every group, then after a keystroke (ED-008 section 8.8 item 11)', async () => {
    const { Finder, stateFor } = await import('../src/index.mjs');
    const session = await EditorDocument.open(new Uint8Array(await readFile(LARGE)));
    const state = stateFor(session);
    for (const [query, options] of [['?^13?', { matchWildcards: true }], ['t*e', { matchWildcards: true }], ['.^p', {}]] as const) {
      const finder = new Finder(query, options, session.ctx.ids);
      const first = await timed(`find ${JSON.stringify(query)} across marks (255 pages)`, async () => finder.all(state.doc));
      const groups = finder.misses;
      const next = state.apply(state.tr.insertText('x', 3));
      const again = await timed(`find ${JSON.stringify(query)} again after a keystroke`, async () => finder.all(next.doc));
      console.log(`benchmark find across: ${JSON.stringify(query)} ${first.result.length} hits in ${groups} groups; after the keystroke ${finder.misses - groups} group searched again, ${again.result.length} hits`);
      expect(Math.abs(again.result.length - first.result.length)).toBeLessThanOrEqual(1);
    }
  }, 120_000);

  test('anon-250.docx: the tracked changes listed, then after each of thirty tracked keystrokes (ED-005 section 6.5)', async () => {
    const { stateFor, recordTransaction, trackedChangesOf, listingBlockWalks } = await import('../src/index.mjs');
    const session = await EditorDocument.open(new Uint8Array(await readFile(LARGE)));
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author C';
    let state = stateFor(session);
    const first = await timed('tracked changes listed (255 pages, the first time)', async () => trackedChangesOf(state.doc));
    // A plain paragraph near the middle, as the browser benchmark's typing uses.
    let at = -1;
    state.doc.forEach((node, pos, index) => { if (at < 0 && index >= state.doc.childCount / 2 && node.type.name === 'paragraph' && node.content.size > 20) at = pos + 5; });
    const walks = listingBlockWalks();
    const times: number[] = [];
    for (let i = 0; i < 30; i++) {
      state = state.apply(recordTransaction(state.tr.insertText('x', at + i), state).tr);
      const start = performance.now();
      trackedChangesOf(state.doc);
      times.push(performance.now() - start);
    }
    times.sort((a, b) => a - b);
    const median = times[15]!;
    console.log(`benchmark listing after a tracked keystroke (255 pages, ${state.doc.childCount} top-level blocks): median ${median.toFixed(2)} ms, max ${times[29]!.toFixed(2)} ms, ${listingBlockWalks() - walks} blocks walked in thirty keystrokes; ${first.result.length} changes before, ${trackedChangesOf(state.doc).length} after`);
    expect(trackedChangesOf(state.doc).length).toBe(first.result.length + 1);
    // A change's key through a map rebuilt from the store (headless; the view passes its live map, ED-005 section 6.5).
    const change = trackedChangesOf(state.doc)[0]!;
    await timed('a change key through a map rebuilt from the store (255 pages)', async () => session.changeKey(change));
  }, 120_000);

  test('anon-250.docx: Select by XPath, //w:p[w:pPr/w:pStyle], the export and the marshal included (ED-005 section 3.5)', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(LARGE)));
    const xpath = '//w:p[w:pPr/w:pStyle]';
    const plain = await timed('select by XPath, the default engine (XPath 1.0)', () => session.selectXPath(xpath));
    await session.template.load();
    const fonto = await timed('select by XPath, FontoXPath (XPath 3.1)', () => session.selectXPath(xpath));
    console.log(`benchmark select: ${plain.result.length} hits (default engine), ${fonto.result.length} (FontoXPath)`);
    expect(fonto.result.length).toBe(plain.result.length);
  }, 120_000);

  test('anon-250.docx: one agent call, a tree run tracked as the agent, the first and then a second (ED-005 section 8.9, step F1)', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(LARGE)));
    const call = (text: string) => session.runOnTree(({ body }) => { body.insertParagraph(text, 'End'); }, 'agent', { author: 'Claude', mode: 'TrackAll' });
    const first = await timed('an agent call, the first (the tree handed to the part)', () => call('The first paragraph the agent writes.'));
    expect(first.result.error).toBeUndefined();
    const second = await timed('an agent call, the second', () => call('The second paragraph the agent writes.'));
    expect(second.result.changed).toBe(true);
  }, 180_000);

  test('footnote.docx: the notes read and projected', async () => {
    const session = await EditorDocument.open(await fixture('footnote.docx'));
    const notes = await timed('footnotes read and projected', () => session.stories.footnotes());
    expect(notes.result.length).toBeGreaterThan(0);
  });
});

// ED-005 section 9.4 items 1 and 2 (E4.g step G3): the revision-scale documents, built by the engine at test time
// (`fixtures/build/revisions-scale.mjs`), at 200 revisions (a long document people have reviewed) and 1,000 (Word's
// Compare Documents, or an agent's pass). The two sizes side by side show what grows with the count.
describe('benchmark: the revision-scale documents (ED-005 section 9.4; reported, never gating)', () => {
  for (const count of [200, 1000]) {
    describe(`${count} revisions`, () => {
      // Built once for the size; each figure a test of its own, so that no one test holds the worker for long.
      let bytes: Uint8Array;
      beforeAll(async () => {
        const { buildRevisionScale } = await import('./fixtures/build/revisions-scale.mjs');
        bytes = await buildRevisionScale(count);
      }, 300_000);

      test('open, and the changes listed the first time', async () => {
        const { trackedChangesOf } = await import('../src/index.mjs');
        const open = await timed(`${count} revisions: open (255 pages)`, () => EditorDocument.open(bytes));
        const listed = await timed(`${count} revisions: the changes listed, the first time`, async () => trackedChangesOf(open.result.doc));
        console.log(`benchmark ${count} revisions: ${listed.result.length} changes listed`);
      }, 300_000);

      test('the listing after each of thirty tracked keystrokes', async () => {
        const { stateFor, recordTransaction, trackedChangesOf } = await import('../src/index.mjs');
        const session = await EditorDocument.open(bytes);
        session.trackingMode = 'TrackAll';
        session.trackingAuthor = 'Author D';
        let state = stateFor(session);
        trackedChangesOf(state.doc);
        let at = -1;
        state.doc.forEach((node, pos, index) => { if (at < 0 && index >= state.doc.childCount / 2 && node.type.name === 'paragraph' && node.content.size > 20) at = pos + 5; });
        const times: number[] = [];
        for (let i = 0; i < 30; i++) {
          state = state.apply(recordTransaction(state.tr.insertText('x', at + i), state).tr);
          const start = performance.now();
          trackedChangesOf(state.doc);
          times.push(performance.now() - start);
        }
        times.sort((a, b) => a - b);
        console.log(`benchmark ${count} revisions: the listing after a tracked keystroke, median ${times[15]!.toFixed(2)} ms, max ${times[29]!.toFixed(2)} ms`);
      }, 300_000);

      test('one accept by its key', async () => {
        const { trackedChangesOf, runCommand } = await import('../src/index.mjs');
        const one = await EditorDocument.open(bytes);
        const changes = trackedChangesOf(one.doc);
        const key = await timed(`${count} revisions: a change's key`, async () => one.changeKey(changes[Math.floor(changes.length / 2)]!));
        const accepted = await timed(`${count} revisions: one accept by its key`, async () => runCommand(one, 'trackedChange.accept', { key: key.result }));
        expect(accepted.result.ok).toBe(true);
      }, 300_000);

      // No revision here is shown in Word only, so neither "all" has a settle half (a tree run) to time apart.
      for (const action of ['accept', 'reject'] as const) {
        test(`${action === 'accept' ? 'Accept' : 'Reject'} all, and the save after it with the self-check`, async () => {
          const { trackedChangesOf, runCommand } = await import('../src/index.mjs');
          const session = await EditorDocument.open(bytes);
          const all = await timed(`${count} revisions: ${action === 'accept' ? 'Accept' : 'Reject'} all`, async () => runCommand(session, `trackedChanges.${action}All`, undefined, { from: 1 }));
          expect(all.result.ok).toBe(true);
          expect(trackedChangesOf(session.doc)).toEqual([]);
          await new Promise((r) => setTimeout(r, 0));
          await timed(`${count} revisions: the save after ${action === 'accept' ? 'Accept' : 'Reject'} all, with the self-check`, () => session.save());
        }, 300_000);
      }

      // The room repair (section 12.15 item 7): after an Accept all in a room, the pass each peer makes over the records
      // after each remote update, timed over thirty of the other peer's keystrokes.
      test('the room repair\'s pass after Accept all, per remote keystroke', async () => {
        const { runCommand, repairWrites } = await import('../src/index.mjs');
        const { room } = await import('./room-harness.mjs');
        const phase = async <T,>(label: string, fn: () => Promise<T>): Promise<T> => (await timed(`${count} revisions, the room: ${label}`, fn)).result;
        const r = await phase('shared and joined', () => room(bytes, 'revisions.docx'));
        const settle = async (): Promise<void> => { for (let round = 0; round < 3; round++) for (const s of [r.creator, r.joiner]) { await s.sourcesReady(); await s.partsReady(); } };
        await settle();
        await phase('Accept all, reaching the other peer', async () => { expect(runCommand(r.creator, 'trackedChanges.acceptAll', undefined, { from: 1 }).ok).toBe(true); await settle(); });
        const records = r.creator.store.resolutions.size;
        // The other peer's keystrokes straight into its Yjs text, one update each: a command would rebuild that
        // peer's 255-page document from the store for each one, which is not what is measured here.
        const Y = await import('yjs');
        const body = r.joiner.store.body;
        let text: InstanceType<typeof Y.XmlText> | undefined;
        for (let b = Math.floor(body.length / 2); b < body.length && !text; b++) {
          const block = body.get(b);
          if (block instanceof Y.XmlElement) text = block.toArray().find((c): c is InstanceType<typeof Y.XmlText> => c instanceof Y.XmlText && c.length > 20);
        }
        expect(text).toBeDefined();
        const passes: number[] = [];
        for (let i = 0; i < 30; i++) {
          r.joiner.store.ydoc.transact(() => text!.insert(5 + i, 'y'));
          const start = performance.now();
          repairWrites(r.creator.store);
          passes.push(performance.now() - start);
          await new Promise((resolve) => setTimeout(resolve, 0));
        }
        passes.sort((a, b) => a - b);
        console.log(`benchmark ${count} revisions: the room repair's pass after Accept all (${records} record${records === 1 ? '' : 's'} holding every change), per remote keystroke: median ${passes[15]!.toFixed(2)} ms, max ${passes[29]!.toFixed(2)} ms`);
        expect(repairWrites(r.creator.store)).toEqual([]);
      }, 300_000);
    });
  }
});

describe('the room\'s stored size after typing (ED-004 section 6.3; reported, never gating)', () => {
  test('5,000 keystrokes in a shared room, one in ten a deletion: the state the relay stores, per keystroke and for a day', async () => {
    const { runCommand } = await import('../src/index.mjs');
    const Y = await import('yjs');
    const bytes = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(bytes);
    await session.share(bytes, 'tracked-changes.docx');
    const size = (): number => Y.encodeStateAsUpdate(session.store.ydoc).length;
    const start = size();
    const keystrokes = 5000;
    // Typing at the end of block 4, as a person types: a character, and every tenth keystroke a Backspace.
    for (let i = 0; i < keystrokes; i++) {
      let pos = 0;
      for (let k = 0; k < 4; k++) pos += session.doc.child(k).nodeSize;
      const end = pos + 1 + session.doc.child(4).content.size;
      if (i % 10 === 9) runCommand(session, 'insertText', { text: '' }, { from: end - 1, to: end });
      else runCommand(session, 'insertText', { text: String.fromCharCode(97 + (i % 26)) }, { from: end });
    }
    const after = size();
    const perKeystroke = (after - start) / keystrokes;
    const day = 43_200;   // eight hours at 1.5 keystrokes a second
    console.log(`benchmark room state: ${start} bytes shared, ${after} after ${keystrokes} keystrokes (${perKeystroke.toFixed(1)} bytes each); a day of typing (${day} keystrokes) about ${Math.round((start + perKeystroke * day) / 1024)} kB`);
    expect(after).toBeGreaterThan(start);
  }, 300_000);
});
