// SPDX-License-Identifier: GPL-3.0-or-later
// @vitest-environment node
// (A package saved under jsdom cannot be reloaded there, ED-003 section 11.2 item 18, and the report reloads an edited one.)
//
// ED-008 section 8.8: the normalisation report. Read-only (Jason, 2026-11-05): what normal form would change in the
// document as it is now, grouped so a file with a thousand values stays readable, and what the engine noted on reading.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { ZipPartStore, type ValueNormalised } from '@docx4j/core-ts';
import { EditorDocument } from '@docx4j/editor-model';
import { groupNormalised, kindOf, normalisationReport } from '../src/app/normalisation.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');
const bytesOf = async (name: string) => new Uint8Array(await readFile(join(fixtures, name)));
const open = async (name: string) => EditorDocument.open(await bytesOf(name));

const entry = (partName: string, path: string, before: string | undefined, after: string | null): ValueNormalised => ({ code: 'value.normalised', partName, path, before, after, message: '' });

describe('grouping', () => {
  test('a value\'s kind is its element and attribute, positions left out; an element left out is its name', () => {
    expect(kindOf('/w:document/w:body[1]/w:p[3]/w:r[1]/w:rPr[1]/w:color[1]/@w:val')).toBe('w:color/@w:val');
    expect(kindOf('/w:document/w:body[1]/w:p[3]/w:r[1]/w:rPr[1]/w:highlight')).toBe('w:highlight');
    expect(kindOf('/w:document/w:body[1]/w:p[3]/w:r[1]/w:rPr[1]/w:highlight[2]')).toBe('w:highlight');
  });

  test('by part, by kind, by value, the most frequent first, with their counts', () => {
    const color = '/w:styles/w:style[1]/w:rPr[1]/w:color[1]/@w:val';
    const fill = '/w:styles/w:style[2]/w:pPr[1]/w:shd[1]/@w:fill';
    expect(groupNormalised([
      entry('/word/header1.xml', '/w:hdr/w:p[1]/w:r[1]/w:rPr[1]/w:highlight', 'purple', null),
      entry('/word/styles.xml', color, 'ffffff', 'FFFFFF'),
      entry('/word/styles.xml', fill, 'aaaaaa', 'AAAAAA'),
      entry('/word/styles.xml', fill, 'bbbbbb', 'BBBBBB'),
      entry('/word/styles.xml', fill, 'bbbbbb', 'BBBBBB'),
    ])).toEqual([
      { partName: '/word/styles.xml', count: 4, kinds: [
        { kind: 'w:shd/@w:fill', count: 3, changes: [{ before: 'bbbbbb', after: 'BBBBBB', count: 2 }, { before: 'aaaaaa', after: 'AAAAAA', count: 1 }] },
        { kind: 'w:color/@w:val', count: 1, changes: [{ before: 'ffffff', after: 'FFFFFF', count: 1 }] },
      ] },
      { partName: '/word/header1.xml', count: 1, kinds: [{ kind: 'w:highlight', count: 1, changes: [{ before: 'purple', after: null, count: 1 }] }] },
    ]);
  });
});

describe('the report', () => {
  test('a Word file with lower-case values throughout its styles: a thousand entries in a few groups', async () => {
    const report = await normalisationReport(await open('numbering-clauses.docx'));
    expect(report.values).toBe(1053);
    expect(report.parts.map((p) => [p.partName, p.count])).toEqual([['/word/styles.xml', 1053]]);
    expect(report.parts[0]!.kinds.map((k) => k.kind)).toEqual(['w:shd/@w:color', 'w:shd/@w:fill', 'w:color/@w:val', 'w:bottom/@w:color', 'w:top/@w:color']);
    expect(report.parts[0]!.kinds.flatMap((k) => k.changes).every((c) => c.after === c.before.toUpperCase())).toBe(true);
  });

  test('reaches every part it reads, the body and the headers, a colour named and a highlight outside Word\'s list', async () => {
    const report = await normalisationReport(await open('normal-form-values.docx'));
    const main = report.parts.find((p) => p.partName === '/word/document.xml')!;
    expect(main.kinds.find((k) => k.kind === 'w:color/@w:val')!.changes).toEqual([
      { before: '00ff00', after: '00FF00', count: 1 },
      { before: 'ff0000', after: 'FF0000', count: 1 },
      { before: 'red', after: 'FF0000', count: 1 },
    ]);
    expect(main.kinds.find((k) => k.kind === 'w:highlight')!.changes).toEqual([{ before: 'purple', after: null, count: 1 }]);
    expect(report.parts.map((p) => p.partName).sort()).toEqual(['/word/document.xml', '/word/header1.xml', '/word/header2.xml', '/word/header3.xml']);
  });

  test('reads the document as it is now: a value the edit took away has gone from the report', async () => {
    const session = await open('normal-form-values.docx');
    await session.runOnTree(({ body }) => { body.paragraphs[3]!.delete(); });
    const report = await normalisationReport(session);
    const changes = report.parts.find((p) => p.partName === '/word/document.xml')!.kinds.find((k) => k.kind === 'w:color/@w:val')!.changes;
    expect(changes.map((c) => c.before)).toEqual(['ff0000', 'red']);
  });

  test('a file in normal form reports nothing; what the engine noted on reading is there all the same', async () => {
    const clean = await normalisationReport(await open('tracked-changes.docx'));
    expect(clean).toEqual({ values: 0, parts: [], notes: [] });
    const noted = await normalisationReport(await open('mc/mc-branches.docx'));
    expect(noted.values).toBe(0);
    // The element with nothing to select is noted by the engine (core-ts CR-029), the only notice besides its opaque atom.
    expect(noted.notes.map((n) => [n.code, n.messages.length])).toEqual([['mce.branches_differ', 10], ['mce.nothing_selected', 1]]);
  });

  test('changes nothing: the session keeps writing values as read, and an unedited document still saves part by part as it was', async () => {
    const source = await bytesOf('numbering-clauses.docx');
    const session = await EditorDocument.open(source);
    await normalisationReport(session);
    expect(session.pkg.normaliseOnSave).toBe(false);
    const { bytes, changed } = await session.save();
    expect(changed).toBe(false);
    const before = new ZipPartStore(source);
    const after = new ZipPartStore(bytes);
    for (const name of before.partNames()) {
      if (name === '[Content_Types].xml' || name.endsWith('.rels')) continue;
      expect(after.loadSync(name), name).toEqual(before.loadSync(name));
    }
  });
});
