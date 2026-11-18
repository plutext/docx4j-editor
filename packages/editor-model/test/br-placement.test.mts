// SPDX-License-Identifier: GPL-3.0-or-later
//
// A w:br outside a w:r, read as Word reads it (ED-008 section 8.8 install item 17; docx4j CR-073, core-ts CR-021): the
// oracle's thirteen one-case files and Word 365's saves of them (`fixtures/br-placement/`, from docx4j-core-tests, Jason's
// run of 3 November 2026; its README says what each holds). Before core-ts 0.7.0 five of them showed the break as an opaque
// atom, the two halves run together, and four did not open at all (`Element text content may not contain START_ELEMENT`).
import { describe, test, expect } from 'vitest';
import { readdir } from 'node:fs/promises';
import { join } from 'node:path';
import { ZipPartStore } from '@docx4j/core-ts';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, paragraphTextsOfBytes } from '../src/index.mjs';
import { bytesEqual, fixture, fixturesDir, regenerated } from './helpers.mjs';

const CASES = (await readdir(join(fixturesDir, 'br-placement'))).filter((n) => n.endsWith('.docx') && !n.endsWith('-word365.docx')).map((n) => n.slice(0, -'.docx'.length)).sort();

/** Every paragraph's text as the projection holds it, a break as the character Office JS reads it as. */
function projectedTexts(doc: PMNode): string[] {
  const out: string[] = [];
  doc.descendants((node) => {
    if (node.type.name !== 'paragraph') return true;
    let text = '';
    node.forEach((child) => {
      if (child.isText) text += child.text!;
      else if (child.type.name === 'hardBreak') text += child.attrs['kind'] === 'page' ? '\u000C' : child.attrs['kind'] === 'column' ? '\u000E' : '\u000B';
    });
    out.push(text);
    return false;
  });
  return out;
}

describe('a w:br outside a w:r (install item 17)', () => {
  test('thirteen cases', () => {
    expect(CASES).toHaveLength(13);
  });

  test.each(CASES)('%s: opens, no break shown as an opaque atom, and an untouched save is the file as it was', async (name) => {
    const bytes = await fixture(`br-placement/${name}.docx`);
    const session = await EditorDocument.open(bytes);
    expect(session.diagnostics.opaque.filter((o) => o.kind === 'br')).toEqual([]);
    const { bytes: saved } = await session.save();
    const before = new ZipPartStore(bytes);
    const after = new ZipPartStore(saved);
    for (const part of before.partNames()) {
      if (regenerated(part)) continue;
      expect(bytesEqual(after.loadSync(part), before.loadSync(part)), part).toBe(true);
    }
  });

  test.each(CASES)('%s: the body reads as Word 365 saved it', async (name) => {
    const session = await EditorDocument.open(await fixture(`br-placement/${name}.docx`));
    const word = await paragraphTextsOfBytes(await fixture(`br-placement/${name}-word365.docx`));
    expect(projectedTexts(session.doc)).toEqual(word);
  });
});
