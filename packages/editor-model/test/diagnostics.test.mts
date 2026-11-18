// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 3.7 test 7: the opaque report per fixture is pinned, so a change that starts
// dropping content (or starts treating editable content as opaque) is caught. Regenerate with
// `UPDATE_DIAGNOSTICS=1 npx vitest run test/diagnostics.test.mts` after a deliberate change,
// and say in the commit message why the numbers moved.
import { describe, test, expect, afterAll } from 'vitest';
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { EditorDocument } from '../src/index.mjs';
import { allFixtures, fixture, fixturesDir } from './helpers.mjs';

interface Snapshot {
  blocks: number;
  opaque: Record<string, number>;
  /** Tracked changes the edit layer holds, and those kept as they are by element (ED-005 section 5.2). */
  tracked: number;
  preserved: Record<string, number>;
  pageBreakMarkers: number;
  comments: number;
}

const expectedPath = join(fixturesDir, 'diagnostics.expected.json');
const update = process.env['UPDATE_DIAGNOSTICS'] === '1';
const expected: Record<string, Snapshot> = update ? {} : JSON.parse(await readFile(expectedPath, 'utf8'));
const actual: Record<string, Snapshot> = {};

describe.each(await allFixtures())('%s', (name) => {
  test('diagnostics match the recorded snapshot', async () => {
    const doc = await EditorDocument.open(await fixture(name));
    const d = doc.diagnostics;
    const opaque: Record<string, number> = {};
    for (const e of d.opaque) opaque[e.kind] = (opaque[e.kind] ?? 0) + 1;
    const snapshot: Snapshot = {
      blocks: doc.imported.childCount,
      opaque: Object.fromEntries(Object.entries(opaque).sort()),
      tracked: d.tracked.shown,
      preserved: Object.fromEntries(Object.entries(d.tracked.preserved).sort()),
      pageBreakMarkers: d.pageBreakMarkers,
      comments: d.comments,
    };
    actual[name] = snapshot;
    if (update) return;
    expect(expected[name], `${name} has no recorded snapshot; run with UPDATE_DIAGNOSTICS=1`).toBeDefined();
    expect(snapshot).toEqual(expected[name]);
  });
});

afterAll(async () => {
  if (update) await writeFile(expectedPath, JSON.stringify(Object.fromEntries(Object.entries(actual).sort()), null, 2) + '\n');
});
