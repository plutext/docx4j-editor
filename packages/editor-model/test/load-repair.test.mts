// SPDX-License-Identifier: GPL-3.0-or-later
//
// A file Word opens only through a repair (core-ts 0.7.0's `LoadOptions.repair`, CR-031; ED-008 section 8.8 install item 18 (h)):
// the application and the companion open with `repair: 'word'`, Jason's decision of 2026-11-09, so such a file opens as Word
// reads it, the repair said on `pkg.warnings`, and its save writes the repaired package. The engine's default refuses it.
import { describe, test, expect } from 'vitest';
import { unzipSync, zipSync } from 'fflate';
import { EditorDocument } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

/** A fixture whose package relationships are in a namespace Word reads only by its repair. */
async function wrongNamespace(): Promise<Uint8Array> {
  const entries = unzipSync(await fixture('revisions/revisions-base.docx'));
  const rels = new TextDecoder().decode(entries['_rels/.rels']!);
  entries['_rels/.rels'] = new TextEncoder().encode(rels.replace('http://schemas.openxmlformats.org/package/2006/relationships', 'http://example.com/not-relationships'));
  return zipSync(entries);
}

describe('repair as Word does', () => {
  test('refused by default; with repair: word it opens, the repair noted, and the save writes the repaired package', async () => {
    const bytes = await wrongNamespace();
    await expect(EditorDocument.open(bytes)).rejects.toThrow();
    const session = await EditorDocument.open(bytes, { repair: 'word' });
    expect(session.pkg.repaired).toBe('word');
    expect(session.pkg.warnings.map((w) => w.code)).toContain('package.relationships_namespace');
    expect(session.doc.childCount).toBeGreaterThan(0);
    const { bytes: saved } = await session.save();
    const again = await EditorDocument.open(saved);
    expect(again.pkg.repaired).toBe('none');
    expect(again.doc.textContent).toBe(session.doc.textContent);
  });

  test('an intact file is read the same either way', async () => {
    const bytes = await fixture('revisions/revisions-base.docx');
    const session = await EditorDocument.open(bytes, { repair: 'word' });
    expect(session.pkg.repaired).toBe('none');
    expect(session.doc.eq((await EditorDocument.open(bytes)).doc)).toBe(true);
  });
});
