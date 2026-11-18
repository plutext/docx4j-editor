// SPDX-License-Identifier: GPL-3.0-or-later
//
// A part edited in place (the Styles tab, 2026-09-24): the session records it, the save reports
// a change and writes app.xml, the main part keeps its source bytes, and the store saw an
// update under the parts origin, not the local one the undo manager tracks.
import { describe, test, expect } from 'vitest';
import { EditorDocument, PARTS_ORIGIN, PARTS_REVISION } from '../src/index.mjs';
import { fixture, comparePackages } from './helpers.mjs';

describe('a part edited in place', () => {
  test('save reports the change, writes app.xml, and leaves the main part alone; undo has nothing', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const origins: unknown[] = [];
    session.store.ydoc.on('update', (_u: Uint8Array, origin: unknown) => origins.push(origin));
    const part = session.main.styleDefinitionsPart!;
    const styles = await part.readContents();
    const normal = styles.style!.find((s) => s.styleId === 'Normal')!;
    normal.rPr = { ...(normal.rPr ?? {}), i: {} } as typeof normal.rPr;
    part.setContents(styles);
    session.pkg.refreshPropertyResolver();
    session.markPartChanged(part.partName.toString());
    expect([...session.touchedParts]).toEqual(['/word/styles.xml']);
    expect(origins).toEqual([PARTS_ORIGIN]);
    expect(session.store.package.get(PARTS_REVISION)).toBe(1);
    expect(session.undoManager.canUndo()).toBe(false);
    expect(session.pkg.propertyResolver.getEffectiveRPr('Normal')?.i).toBeDefined();
    const saved = await session.save();
    expect(saved.changed).toBe(true);
    expect(comparePackages(original, saved.bytes, ['word/styles.xml', 'docProps/app.xml'])).toEqual([]);
    const reopened = await EditorDocument.open(saved.bytes);
    expect(reopened.pkg.propertyResolver.getEffectiveRPr('Normal')?.i).toBeDefined();
  });
});
