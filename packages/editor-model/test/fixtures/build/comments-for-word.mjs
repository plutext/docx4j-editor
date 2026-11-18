// SPDX-License-Identifier: GPL-3.0-or-later
//
// The file for ED-005 section 7.8's check in Word 2010 and Word 15 on the reference machine (step
// E5): comments made, replied to, resolved and deleted in the editor, in a room, so each comment
// made there is written with a small id in place of its random one (section 7.4). Made from
// `tracked-changes.docx`, which has no comments, so the editor makes the comments part, its side
// parts and the two comment styles itself. Run from packages/editor-model after `npm run build`:
//
//   node test/fixtures/build/comments-for-word.mjs [out.docx]
//
// (the default output is test-results/comments-for-word.docx, which git ignores)
//
// What to see in Word 15: Author A's comment on "Docx sample" with Author B's reply under it, the
// thread resolved; Author A's comment over four paragraphs, from "exhibiting" to "Some"; Author B's
// comment on "Heading" and its reply by Author A; no comment on "another document" (made by Author
// B and deleted).
// Word 2010 shows each comment and reply as a comment of its own (it has no threads) and nothing
// resolved. Then save in each Word and reopen the save in the editor.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { EditorDocument, runCommandAsync } from '../../../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const out = process.argv[2] ?? join(here, '..', '..', '..', 'test-results', 'comments-for-word.docx');
await mkdir(dirname(out), { recursive: true });
const bytes = new Uint8Array(await readFile(join(here, '..', 'tracked-changes.docx')));
const session = await EditorDocument.open(bytes);
await session.share(bytes, 'tracked-changes.docx');
await session.stories.comments();

/** Where `text` starts in the document, and a selection over it. */
const find = (text) => {
  let at = -1;
  session.doc.descendants((n, pos) => { if (at < 0 && n.isText && n.text.includes(text)) at = pos + n.text.indexOf(text); return at < 0; });
  if (at < 0) throw new Error(`no "${text}"`);
  return { from: at, to: at + text.length };
};
const as = (author) => { session.trackingAuthor = author; };
const run = async (name, args, selection) => {
  const result = await runCommandAsync(session, name, args, selection ?? { from: 1 });
  if (!result.ok) throw new Error(`${name} did not apply`);
};
const idOf = (text) => session.stories.commentThreads(session.doc).flatMap((t) => [t, ...t.replies]).find((t) => session.stories.currentDoc(t).textContent === text)?.id;

as('Author A');
await run('range.insertComment', { text: 'A comment by Author A, replied to and resolved.' }, find('Docx sample'));
await run('range.insertComment', { text: 'A comment by Author A over four paragraphs.' }, { from: find('exhibiting').from, to: find('Some').to });
as('Author B');
await run('comment.reply', { id: idOf('A comment by Author A, replied to and resolved.'), text: 'A reply by Author B.' });
await run('comment.setResolved', { id: idOf('A comment by Author A, replied to and resolved.'), value: true });
await run('range.insertComment', { text: 'A comment by Author B, with a reply by Author A.' }, find('Heading'));
await run('range.insertComment', { text: 'A comment by Author B, deleted.' }, find('another document'));
await run('comment.delete', { id: idOf('A comment by Author B, deleted.') });
as('Author A');
await run('comment.reply', { id: idOf('A comment by Author B, with a reply by Author A.'), text: 'A reply by Author A.' });

const saved = await session.save();
await writeFile(out, saved.bytes);
const listed = session.stories.commentThreads(session.doc).flatMap((t) => [`${t.author}${t.done ? ' (resolved)' : ''}: ${session.stories.currentDoc(t).textContent}`, ...t.replies.map((r) => `  ${r.author}: ${session.stories.currentDoc(r).textContent}`)]);
console.log(`wrote ${out} (self-check ${saved.selfCheck?.equal ? 'passed' : 'FAILED'}):\n${listed.join('\n')}`);
