// SPDX-License-Identifier: GPL-3.0-or-later
//
// The files for bisecting Word 2010's limit on a note's or a comment's `w:id` (ED-005 section 9.2
// item 3, E4.g step G1; ED-004 section 7). Word 2010 refused a room's random note id, 1,050,702,260,
// just under 2^30 (ED-004 section 10.8 item 6); a save now writes small ids, so the limit gates
// nothing, and this records a number. Twelve small files, each with one footnote or one comment made
// by the editor, its id then set to N (in the part and at its references), for N in 32,767, 32,768,
// 65,535, 65,536, 2^24 and 2^29: the known refusal bounds the search from above. Word 2010 opens
// each; the limit lies between the largest that opens and the smallest refused. Run from the
// repository root after `npm run build`:
//
//   node packages/editor-model/test/fixtures/build/ids-for-word.mjs [out-dir] [N ...]
//
// It writes ids/footnote-N.docx and ids/comment-N.docx into out-dir, by default
// packages/editor-mcp/test-results/word-session/, which git ignores. Give N values to make a second
// round between two of them.
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, runCommandAsync } from '../../../dist/index.mjs';

export const IDS = [32_767, 32_768, 65_535, 65_536, 2 ** 24, 2 ** 29];

/** A small document with one footnote or one comment, made by the editor. */
async function made(kind) {
  const pkg = await WordprocessingMLPackage.createPackage();
  pkg.body.insertParagraph(`One ${kind} whose id is set to the value in the file's name, for Word 2010's limit (ED-005 section 9.2).`, 'End');
  const session = await EditorDocument.open(new Uint8Array(await pkg.save()));
  session.trackingAuthor = 'Author A';
  if (kind === 'comment') await session.stories.comments();
  const end = session.doc.child(0).nodeSize - 1;
  const result = kind === 'footnote'
    ? await runCommandAsync(session, 'range.insertFootnote', { text: 'The footnote.' }, { from: end })
    : await runCommandAsync(session, 'range.insertComment', { text: 'The comment.' }, { from: 1, to: 4 });
  if (!result.ok) throw new Error(`the ${kind} was not made`);
  return (await session.save()).bytes;
}

/** Sets the one note's or comment's id to `n`, in its part and at its references and markers. */
async function withId(bytes, kind, n) {
  const pkg = await WordprocessingMLPackage.load(bytes);
  const main = pkg.getMainDocumentPart();
  const partName = kind === 'footnote' ? '/word/footnotes.xml' : '/word/comments.xml';
  const part = pkg.parts.get(partName);
  if (!part) throw new Error(`no ${partName}`);
  let doc = await main.getXml();
  const ref = kind === 'footnote' ? 'w:footnoteReference' : 'w:commentReference';
  const id = new RegExp(`<${ref} [^>]*w:id="(\\d+)"`).exec(doc)?.[1];
  if (id === undefined) throw new Error(`no ${ref}`);
  const elements = kind === 'footnote' ? ['w:footnoteReference'] : ['w:commentRangeStart', 'w:commentRangeEnd', 'w:commentReference'];
  for (const element of elements) doc = doc.replace(new RegExp(`(<${element} [^>]*w:id=")${id}(")`, 'g'), `$1${n}$2`);
  let xml = await part.getXml();
  const own = kind === 'footnote' ? 'w:footnote' : 'w:comment';
  const before = xml;
  xml = xml.replace(new RegExp(`(<${own} [^>]*w:id=")${id}(")`), `$1${n}$2`);
  if (xml === before) throw new Error(`no ${own} with id ${id}`);
  main.setXml(doc);
  part.setXml(xml);
  return new Uint8Array(await pkg.save());
}

/** The files: `{ name, bytes, kind, id }` for each kind and each N. */
export async function buildIdsForWord(ids = IDS) {
  const out = [];
  for (const kind of ['footnote', 'comment']) {
    const base = await made(kind);
    for (const n of ids) out.push({ name: `${kind}-${n}.docx`, bytes: await withId(base, kind, n), kind, id: n });
  }
  return out;
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const here = dirname(fileURLToPath(import.meta.url));
  const [dir, ...values] = process.argv.slice(2);
  const out = join(dir ?? join(here, '..', '..', '..', '..', 'editor-mcp', 'test-results', 'word-session'), 'ids');
  await mkdir(out, { recursive: true });
  const files = await buildIdsForWord(values.length ? values.map(Number) : IDS);
  for (const f of files) await writeFile(join(out, f.name), f.bytes);
  console.log(`wrote ${files.length} files into ${out}: ${files.map((f) => f.name).join(', ')}`);
}
