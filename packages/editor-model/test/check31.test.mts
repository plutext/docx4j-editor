// SPDX-License-Identifier: GPL-3.0-or-later
//
// core-ts check 31 (run by Jason in Word 16.0.20326.20158, 2026-10-01; `fixtures/revisions/check31/result.json`, copied
// from docx4j-core-ts's `test/fixtures/revisions/check31/`, its commit 2414220): rows inserted with tracking on, in
// Word's own markup (a table inserted between paragraphs, alone, two tables; rows added at a table's end, its start,
// apart, with text typed before or after) and in the engine's, with what Office JS's `body.getTrackedChanges()` listed.
// A row groups with what touches it, as any piece does: neither a table's edge nor a row's is a break (Jason's decision,
// ED-005 section 12.45). The editor's listing is held to it.
import { describe, expect, test } from 'vitest';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommandAsync, trackedChangesOf } from '../src/index.mjs';
import { fixturesDir } from './helpers.mjs';

interface Listed { type: string; author: string; text: string }
interface Case { markup: string[]; listed: Listed[]; acceptAll: { markup: string[] }; rejectAll: { markup: string[] } }

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A document whose body is the case's markup, as Word or the engine left it. */
async function documentOf(markup: readonly string[]): Promise<Uint8Array> {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.getMainDocumentPart().setXml(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${markup.join('')}<w:sectPr/></w:body></w:document>`);
  return new Uint8Array(await pkg.save());
}

/**
 * Where the editor does not follow Office JS, with core-ts (its `docs/word-known-issues.md`): the middle row of three by
 * Author B, Office JS lists nothing for (entry 7), though its accept and reject resolve it. The editor lists it.
 */
const EXPECTED: Record<string, Listed[]> = {
  'engine-two-authors': [
    { type: 'Added', author: 'Author A', text: 'before the table\rrow one\r\n' },
    { type: 'Added', author: 'Author B', text: 'row two\r\n' },
    { type: 'Added', author: 'Author A', text: 'row three\r\n(end of insertion)' },
  ],
};

/** A document as its blocks' text, a table as its rows' cells: what Word's resolution is compared by. */
function shape(doc: PMNode): string[] {
  const out: string[] = [];
  doc.forEach((b) => {
    if (b.type.name !== 'table') { out.push(b.textContent); return; }
    const rows: string[] = [];
    b.forEach((r) => { const cells: string[] = []; r.forEach((c) => cells.push(c.textContent)); rows.push(cells.join('|')); });
    out.push(`[${rows.join(' / ')}]`);
  });
  return out;
}

describe('core-ts check 31: rows group with what touches them, as Office JS lists them', async () => {
  const result = JSON.parse(await readFile(join(fixturesDir, 'revisions', 'check31', 'result.json'), 'utf8')) as Record<'wordInserts' | 'gestures' | 'engine', Record<string, Case>>;
  for (const group of ['wordInserts', 'gestures', 'engine'] as const) {
    for (const [name, c] of Object.entries(result[group])) {
      test(`${group}: ${name}`, async () => {
        const session = await EditorDocument.open(await documentOf(c.markup));
        const listed = trackedChangesOf(session.doc).map((x) => ({ type: x.type, author: x.author, text: x.text }));
        expect(listed).toEqual((EXPECTED[name] ?? c.listed).map((x) => ({ type: x.type, author: x.author, text: x.text })));
      });
      // Word's Accept All and Reject All over the same markup: the editor's leaves the same blocks, rows and text.
      for (const action of ['accept', 'reject'] as const) {
        test(`${group}: ${name}, ${action} all as Word's`, async () => {
          const session = await EditorDocument.open(await documentOf(c.markup));
          expect((await runCommandAsync(session, `trackedChanges.${action}All`, {})).ok).toBe(true);
          const word = await EditorDocument.open(await documentOf(c[`${action}All`].markup));
          expect(trackedChangesOf(session.doc)).toEqual([]);
          expect(shape(session.doc)).toEqual(shape(word.doc));
        });
      }
      // A case Office JS lists as one change: resolved by its key alone, the panel's path, it leaves what "all" leaves.
      if (c.listed.length === 1) {
        test(`${group}: ${name}, its one change accepted by its key`, async () => {
          const session = await EditorDocument.open(await documentOf(c.markup));
          const [change] = trackedChangesOf(session.doc);
          expect((await runCommandAsync(session, 'trackedChange.accept', { key: session.changeKey(change!) })).ok).toBe(true);
          const word = await EditorDocument.open(await documentOf(c.acceptAll.markup));
          expect(shape(session.doc)).toEqual(shape(word.doc));
        });
      }
    }
  }
});
