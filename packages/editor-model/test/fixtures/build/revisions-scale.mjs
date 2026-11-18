// SPDX-License-Identifier: GPL-3.0-or-later
//
// The revision-scale documents of ED-005 section 9.4 item 1 (E4.g step G3): `anon-250.docx` with
// `count` tracked edits made by the engine, spread evenly over its paragraphs, by three authors in
// turn and of four kinds in turn (text inserted at a paragraph's start, a word deleted, a word made
// bold or not, a paragraph's alignment changed), all dated at one instant, so a build is the same
// every time. 200 is the everyday figure, a long document people have reviewed; 1,000 the stress
// end, which Word's Compare Documents or an agent's pass over a long document reaches in minutes.
// Built at test time, never committed (ED-005 proposal 37).
//
//   node packages/editor-model/test/fixtures/build/revisions-scale.mjs [out-dir] [count ...]
//
// writes revisions-<count>.docx into out-dir (default packages/editor/test-results/revision-scale/,
// where the browser benchmark reads them), for 200 and 1,000 unless counts are given.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { WordprocessingMLPackage } from '@docx4j/core-ts';

const here = dirname(fileURLToPath(import.meta.url));

/** The source: the 255-page document. */
export const LARGE = join(here, '..', 'large', 'anon-250.docx');

/** The sizes the figures are taken at. */
export const SIZES = [200, 1000];

export const AUTHORS = ['Author A', 'Author B', 'Author C'];
const KINDS = ['insertion', 'deletion', 'formatting', 'properties'];
/** One instant for every revision, so that a build is the same bytes every time. */
const DATE = new Date('2026-09-30T00:00:00Z');
/** A word to act on: four letters or more. */
const WORD = /\b[A-Za-z]{4,}\b/;

/** The document with `count` tracked edits: its bytes. */
export async function buildRevisionScale(count, source) {
  const pkg = await WordprocessingMLPackage.load(source ?? new Uint8Array(await readFile(LARGE)));
  // The body with the property resolver the effective reads (alignment, bold) need.
  const body = await pkg.getBody();
  const eligible = body.paragraphs.filter((p) => p.text.length > 20 && WORD.test(p.text));
  if (eligible.length < count) throw new Error(`only ${eligible.length} paragraphs to revise, for ${count} revisions`);
  const step = eligible.length / count;
  for (let i = 0; i < count; i++) {
    const paragraph = eligible[Math.floor(i * step)];
    const kind = KINDS[i % KINDS.length];
    await pkg.withTracking({ author: { name: AUTHORS[i % AUTHORS.length] }, mode: 'TrackAll', date: DATE }, () => {
      const word = WORD.exec(paragraph.text)?.[0];
      if (kind === 'insertion' || word === undefined) { paragraph.insertText('inserted ', 'Start'); return; }
      const range = paragraph.search(word)[0];
      if (!range) { paragraph.insertText('inserted ', 'Start'); return; }
      if (kind === 'deletion') range.delete();
      else if (kind === 'formatting') range.font.bold = range.font.bold !== true;
      else paragraph.alignment = paragraph.alignment === 'Centered' ? 'Left' : 'Centered';
    });
  }
  return new Uint8Array(await pkg.save());
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const [dir, ...counts] = process.argv.slice(2);
  const out = dir ?? join(here, '..', '..', '..', '..', 'editor', 'test-results', 'revision-scale');
  await mkdir(out, { recursive: true });
  for (const count of counts.length ? counts.map(Number) : SIZES) {
    await writeFile(join(out, `revisions-${count}.docx`), await buildRevisionScale(count));
    console.log(`wrote ${join(out, `revisions-${count}.docx`)}`);
  }
}
