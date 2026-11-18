// SPDX-License-Identifier: GPL-3.0-or-later
//
// core-ts check 30 (run by Jason in Word 16.0.20326.20158, 2026-09-30; `fixtures/revisions/check30/result.json`, copied from
// docx4j-core-ts's `test/fixtures/revisions/check30/`, its commit 426e7b5): the markup a tracked insertion ended as in Word for each kind of run
// holder, and what Office JS's `body.getTrackedChanges()` listed over it. The editor's listing is held to it (ED-005 section
// 9.7, "the listing split at a run holder"): a holder's edge parts an insertion, as Office JS parts it.
import { describe, expect, test } from 'vitest';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, trackedChangesOf } from '../src/index.mjs';
import { fixturesDir } from './helpers.mjs';

interface Case { markup: string[]; listed: { type: string; author: string; text: string }[] }

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A document whose body is the case's markup, as Word left it. */
async function documentOf(markup: readonly string[]): Promise<Uint8Array> {
  const pkg = await WordprocessingMLPackage.createPackage();
  const main = pkg.getMainDocumentPart();
  main.setXml(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${markup.join('')}<w:sectPr/></w:body></w:document>`);
  return new Uint8Array(await pkg.save());
}

describe('core-ts check 30: the listing parts an insertion at a run holder, as Office JS does', async () => {
  const result = JSON.parse(await readFile(join(fixturesDir, 'revisions', 'check30', 'result.json'), 'utf8')) as { engine: Record<string, Case> };
  // An inserted table is one change with the insertions that touch it, as Office JS lists it: the editor followed at
  // Jason's decision (core-ts check 31, ED-005 section 12.45).
  for (const [name, c] of Object.entries(result.engine)) {
    test(name, async () => {
      const session = await EditorDocument.open(await documentOf(c.markup));
      const listed = trackedChangesOf(session.doc).map((x) => ({ type: x.type, author: x.author, text: x.text }));
      expect(listed).toEqual(c.listed.map((x) => ({ type: x.type, author: x.author, text: x.text })));
    });
  }
});
