// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 3.7, tests 1 to 3 and 7: every fixture, untouched, round-trips part by part
// (the main document part included), re-projects to the same document, and survives Yjs.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, compareProjections, readDocument, summarize, listLabels } from '../src/index.mjs';
import { allFixtures, fixture, comparePackages } from './helpers.mjs';

const names = await allFixtures();

describe.each(names)('%s', (name) => {
  test('untouched: byte-identical part by part, main part included, app.xml untouched', async () => {
    const original = await fixture(name);
    const doc = await EditorDocument.open(original);
    // Asking for the list labels (the Emulator over the private tree) must not cost the numbering part its round trip.
    listLabels(doc);
    const saved = await doc.save();
    expect(saved.changed).toBe(false);
    expect(comparePackages(original, saved.bytes)).toEqual([]);
    // Saving did not unmarshal the main part: the part still answers with its source bytes.
    const reloaded = await WordprocessingMLPackage.load(saved.bytes);
    expect(reloaded.getMainDocumentPart().partName.toString()).toBe(doc.main.partName.toString());
  });

  test('projection is stable across a reload', async () => {
    const original = await fixture(name);
    const a = await EditorDocument.open(original);
    const b = await EditorDocument.open((await a.save({ selfCheck: false })).bytes);
    const result = compareProjections(a.imported, b.imported);
    expect(result.difference).toBeUndefined();
    expect(result.equal).toBe(true);
  });

  test('through Yjs: the store gives back an equal document, and the export is unchanged', async () => {
    const doc = await EditorDocument.open(await fixture(name));
    const fromStore = readDocument(doc.store, doc.schema);
    expect(fromStore.eq(doc.imported)).toBe(true);
    expect(doc.exportTree(fromStore).changed).toBe(false);
  });

  test('diagnostics: opaque content is listed with a kind and an address', async () => {
    const doc = await EditorDocument.open(await fixture(name));
    for (const entry of doc.diagnostics.opaque) {
      expect(entry.kind).not.toBe('');
      expect(entry.address.startsWith('body/')).toBe(true);
      expect(doc.ctx.ids.get(entry.sourceId)).toBeDefined();
    }
    // Printed, not asserted: the counts are recorded per fixture in diagnostics.test.mts.
    if (doc.diagnostics.opaque.length > 0 || doc.diagnostics.tracked.shown > 0) console.log(`${name}: ${summarize(doc.diagnostics)}`);
  });
});
