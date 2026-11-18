// SPDX-License-Identifier: GPL-3.0-or-later
//
// The editor's file for the Word pass of ED-005 section 9.2 item 1 (E4.g step G1): a document made
// in a room by two people and the agent, carrying what the editor and the companion write, for
// Jason to open, resolve and save in Word 2010 and Word 15 on the reference machine. Nothing in it
// is hand-written markup: the engine builds the base, a person's session makes its revisions
// through the editor's commands with tracking on (as Author A and Author B), the agent makes its own
// through the companion's tools, and the parts are changed through E4.b's Apply. The person's session
// saves it, so the room's ids are made small as any save in a room makes them.
//
// Run from the repository root after `npm run build`:
//
//   node packages/editor-mcp/scripts/review-for-word.mjs [out-dir]
//
// It writes review-for-word.docx and review-for-word.txt (the changes as the editor lists them, in
// Office JS's form, to compare with Word's Reviewing Pane) into out-dir, by default
// packages/editor-mcp/test-results/word-session/, which git ignores. `fixtures/review/README.md` in
// the model's fixtures says what to do with them in Word.
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import * as Y from 'yjs';
import WebSocket from 'ws';
import { HocuspocusProvider, HocuspocusProviderWebsocket } from '@hocuspocus/provider';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { p, r } from '@docx4j/generated-objects-ts/builders/wml';
import {
  EditorDocument, applyPartXml, checkPartXml, newEditStore, preservedChangesOf, recordTransaction, runCommandAsync, stateFor, trackedChangesOf,
} from '@docx4j/editor-model';
import { createCompanion } from '../dist/index.mjs';

const AUTHOR_A = 'Author A';
const AUTHOR_B = 'Author B';
const AGENT = 'Claude';

/** The base document, built with the engine: a heading per section naming what the session does there. */
async function baseDocument() {
  const pkg = await WordprocessingMLPackage.createPackage();
  const body = pkg.body;
  const heading = (text, level = 1) => { const h = body.insertParagraph(text, 'End'); h.styleBuiltIn = `Heading${level}`; return h; };
  const plain = (text) => body.insertParagraph(text, 'End');
  body.insertParagraph('Review file for Word (ED-005 section 9.2)', 'End').styleBuiltIn = 'Title';
  plain('Made by the docx4j editor in a room: its revisions by Author A and Author B through the editor, and by Claude through the companion\'s tools. Compare Word\'s Reviewing Pane with review-for-word.txt, save this file untouched, then Accept All and save, then Reject All and save.');

  heading('1. Text inserted and deleted by two authors');
  plain('The quick brown fox jumps over the lazy dog near the old stone bridge.');
  heading('2. An insertion deleted by another author');
  plain('Cows graze in the field beside the river.');
  heading('3. Typed over a selection');
  plain('The notice period is 7 days from delivery.');
  heading('4. Run formatting changed');
  plain('The word alpha becomes bold and the word bravo becomes larger.');
  heading('5. Paragraph properties changed');
  plain('Centre this paragraph.');
  plain('Indent this paragraph.');
  plain('Give this paragraph the style Heading 2.');
  heading('A Heading 2 already in the file, for comparison', 2);
  heading('6. Paragraph marks inserted and deleted');
  plain('Split this paragraph after the colon: the rest becomes a paragraph of its own.');
  plain('Join this paragraph with the next one.');
  plain('This is the paragraph it joins.');
  heading('7. Table rows inserted and deleted');
  const rows = body.insertTable(4, 3, 'End', [['Row', 'Action', 'Note'], ['1', 'Keep', 'stays'], ['2', 'Keep', 'a row goes below this one'], ['3', 'Delete', 'this row is deleted']]);
  rows.styleBuiltIn = 'TableGrid';
  plain('The paragraph after the table.');
  heading('8. Inline items inserted and deleted');
  // The builders' `r` makes a `w:tab` of the `\t`.
  body.insertElement(p([r('Left\tright: the tab between them is deleted. A tab goes after this colon: and a footnote at the end of this sentence.')]), 'End');
  plain('The footnote on this sentence is deleted.');
  heading('9. A list started, and a link');
  plain('This paragraph becomes a numbered list item.');
  plain('The docx4j site is linked here.');
  heading('10. Replace all');
  plain('The colour of the sky and the colour of the sea.');
  heading('11. The parts, changed through the parts panel (untracked)');
  plain('The headings\' colour, the numbering format of the list below and the default tab stop were changed by an Apply, which Word does not track.');
  const first = plain('The first item of a list the engine made.');
  const list = await first.startNewList();
  plain('The second item.').attachToList(list.id, 0);
  heading('12. The agent\'s edits');
  plain('The tenant pays rent monthly in arrears.');
  plain('Notice must be given within 30 days.');
  heading('13. Comments');
  body.insertElement(p([r('Author A comments on this sentence, and Author B answers. '), r('Author B comments on this sentence alone. '), r('This sentence carries a comment that is deleted.')]), 'End');
  plain('A comment by Author A runs from here');
  plain('to here, over two paragraphs.');
  return pkg.save();
}

/** Resolves when `check` holds, polling. */
async function until(check, what, ms = 15_000) {
  const end = Date.now() + ms;
  while (!check()) {
    if (Date.now() > end) throw new Error(`timed out waiting for ${what}`);
    await new Promise((res) => setTimeout(res, 25));
  }
}

/**
 * Builds the file. Returns its bytes, the self-check's verdict, and the changes and comments as the
 * editor lists them after reopening the saved file.
 */
export async function buildReviewForWord() {
  const cleanups = [];
  try {
    const root = await mkdtemp(join(tmpdir(), 'review-for-word-root-'));
    const app = await mkdtemp(join(tmpdir(), 'review-for-word-app-'));
    const data = await mkdtemp(join(tmpdir(), 'review-for-word-rooms-'));
    cleanups.push(() => rm(root, { recursive: true, force: true }), () => rm(app, { recursive: true, force: true }), () => rm(data, { recursive: true, force: true }));
    await writeFile(join(app, 'index.html'), '<!doctype html><title>docx4j editor</title>');
    await writeFile(join(root, 'review-for-word.docx'), await baseDocument());

    const companion = createCompanion({ roots: [root], port: 0, app, data, name: AGENT, log: () => undefined });
    await companion.listen();
    const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
    await companion.mcp().connect(serverSide);
    const client = new Client({ name: 'review-for-word', version: '1' });
    await client.connect(clientSide);
    cleanups.push(async () => { await client.close(); await companion.close(); });
    const tool = async (name, args) => {
      const result = await client.callTool({ name, arguments: args });
      if (result.isError) throw new Error(`${name}: ${result.content?.[0]?.text}`);
      return result.structuredContent ?? {};
    };

    // The agent opens the file; the person joins its room through the link, as a page does.
    const opened = await tool('open', { path: 'review-for-word.docx' });
    const document = opened.document;
    const params = new URLSearchParams(opened.link.split('#')[1]);
    const pageDoc = new Y.Doc();
    const socket = new HocuspocusProviderWebsocket({ url: params.get('relay'), WebSocketPolyfill: WebSocket, maxAttempts: 1 });
    let synced;
    const ready = new Promise((res) => { synced = res; });
    const provider = new HocuspocusProvider({ websocketProvider: socket, name: params.get('room'), document: pageDoc, token: params.get('token'), onSynced: () => synced() });
    provider.attach();
    cleanups.push(() => { provider.destroy(); socket.destroy(); socket.connect = async () => undefined; });
    await ready;
    const store = newEditStore(pageDoc);
    const person = await EditorDocument.open(store.room.get('base'), { store, mode: 'join' });
    await person.stories.comments();
    const agentRoom = companion.rooms.get(document);

    /** Where `text` starts in the person's document (inside one text node). */
    const at = (text) => {
      let found = -1;
      person.doc.descendants((n, pos) => { if (found < 0 && n.isText && n.text.includes(text)) found = pos + n.text.indexOf(text); return found < 0; });
      if (found < 0) throw new Error(`no "${text}"`);
      return found;
    };
    const over = (text) => ({ from: at(text), to: at(text) + text.length });
    /** The end of the paragraph holding `text`. */
    const endOf = (text) => {
      let found = -1;
      person.doc.descendants((n, pos) => { if (found < 0 && n.type.name === 'paragraph' && n.textContent.includes(text)) found = pos + n.nodeSize - 1; return found < 0; });
      if (found < 0) throw new Error(`no paragraph with "${text}"`);
      return found;
    };
    const as = (author) => { person.trackingAuthor = author; };
    const run = async (name, args, selection) => {
      const result = await runCommandAsync(person, name, args, selection ?? { from: 1 });
      if (!result.ok) throw new Error(`${name} did not apply`);
    };
    /** A deletion typed as a person types it (Delete over a selection), tracked by the session's mode. */
    const erase = (from, to) => {
      const state = stateFor(person, { from, to });
      const recorded = recordTransaction(state.tr.delete(from, to), state);
      person.replace(state.apply(recorded.tr).doc);
    };
    /** The position inside the first paragraph of the first table's cell at `row`, `column`. */
    const cell = (row, column) => {
      let table = -1;
      person.doc.descendants((n, pos) => { if (table < 0 && n.type.name === 'table') table = pos; return table < 0; });
      const node = person.doc.nodeAt(table);
      let pos = table + 1;
      for (let i = 0; i < row; i++) pos += node.child(i).nodeSize;
      pos += 1;
      for (let j = 0; j < column; j++) pos += node.child(row).child(j).nodeSize;
      return pos + 2;
    };

    // Before tracking: a footnote the session then deletes, so that a deleted note reference is in the file.
    as(AUTHOR_A);
    await run('range.insertFootnote', { text: 'A footnote whose reference is deleted.' }, { from: at('The footnote on this sentence is deleted') + 'The footnote on this sentence is deleted'.length });
    await run('document.setChangeTrackingMode', { value: 'TrackAll' });

    // 1 and 2: text inserted and deleted; an insertion deleted by another author.
    await run('insertText', { text: 'very ' }, { from: at('lazy') });
    await run('insertText', { text: 'green ' }, { from: at('field') });
    as(AUTHOR_B);
    erase(at('old '), at('old ') + 4);
    erase(at('green '), at('green ') + 6);
    // 3: typed over a selection, a deletion and an insertion (ED-005 section 9.7's replacement).
    as(AUTHOR_A);
    await run('insertText', { text: '14 days' }, over('7 days'));
    // 4: run formatting.
    await run('font.bold', { value: true }, over('alpha'));
    as(AUTHOR_B);
    await run('font.size', { value: 16 }, over('bravo'));
    // 5: paragraph properties.
    as(AUTHOR_A);
    await run('paragraph.alignment', { value: 'Centered' }, { from: at('Centre this') });
    as(AUTHOR_B);
    await run('paragraph.increaseIndent', {}, { from: at('Indent this') });
    as(AUTHOR_A);
    await run('paragraph.style', { value: 'Heading2' }, { from: at('Give this paragraph') });
    // 6: a paragraph split, and two joined.
    await run('insertParagraph', {}, { from: at('colon:') + 'colon:'.length });
    as(AUTHOR_B);
    await run('joinForward', {}, { from: endOf('Join this paragraph') });
    // 7: a row inserted and typed in, and one deleted.
    as(AUTHOR_A);
    await run('table.addRows', { location: 'After', count: 1 }, { from: at('a row goes below') });
    for (let c = 0; c < 3; c++) await run('insertText', { text: 'new' }, { from: cell(3, c) });
    as(AUTHOR_B);
    await run('table.deleteRows', { count: 1 }, { from: at('this row is deleted') });
    // 8: a tab deleted and one inserted; a footnote inserted, and the earlier one's reference deleted.
    as(AUTHOR_A);
    {
      let tab = -1;
      person.doc.nodesBetween(at('Left'), endOf('Left'), (n, pos) => { if (tab < 0 && n.type.name === 'tab') tab = pos; return tab < 0; });
      if (tab < 0) throw new Error('no tab to delete');
      erase(tab, tab + 1);
    }
    await run('insertTab', {}, { from: at('this colon:') + 'this colon:'.length });
    as(AUTHOR_B);
    await run('range.insertFootnote', { text: 'A tracked footnote.' }, { from: at('of this sentence.') + 'of this sentence'.length });
    {
      let ref = -1;
      const para = at('The footnote on this sentence');
      person.doc.nodesBetween(para, endOf('The footnote on this sentence'), (n, pos) => { if (ref < 0 && n.type.name === 'footnoteRef') ref = pos; return ref < 0; });
      if (ref < 0) throw new Error('no footnote reference to delete');
      erase(ref, ref + 1);
    }
    // 9: a list started; a link made.
    as(AUTHOR_A);
    await run('list.start', {}, { from: at('This paragraph becomes') });
    as(AUTHOR_B);
    await run('range.hyperlink', { address: 'https://www.docx4java.org/' }, over('docx4j site'));
    // 10: Replace all.
    await run('find.replaceAll', { text: 'colour', replacement: 'color', matchWholeWord: true });

    // 12: the agent, once the person's edits have reached it.
    await until(() => agentRoom.session.doc.textContent.includes('color of the sea'), 'the person\'s edits at the companion');
    const blocks = (await tool('outline', { document, budget: 1000 })).blocks;
    const blockWith = (text) => {
      const b = blocks.find((x) => x.kind === 'paragraph' && x.text.includes(text));
      if (!b) throw new Error(`no block with "${text}"`);
      return b;
    };
    const rent = blockWith('The tenant pays rent');
    const edited = await tool('edit', { document, address: rent.address, if_text: rent.text, markdown: '## Rent\n\nThe tenant pays rent **monthly** in advance, as the [lease](https://www.docx4java.org/) says.' });
    const created = edited.created;
    const advance = created.find((c) => c.text.includes('in advance'));
    await tool('insert', { document, address: advance.address, location: 'After', markdown: '- The first obligation.\n- The second obligation.' });
    await tool('comment', { document, address: advance.address, anchor: 'in advance', text: 'Changed from arrears, as agreed.' });
    await tool('replace_text', { document, find: '30 days', replace: '60 days' });
    await until(() => person.doc.textContent.includes('60 days') && person.doc.textContent.includes('The second obligation.'), 'the agent\'s edits at the person');
    // A peer's comment is adopted when the comments are read, as the margin reads them on each change.
    for (let i = 0; !(await person.stories.comments()).some((t) => t.author === AGENT); i++) {
      if (i > 300) throw new Error('timed out waiting for the agent\'s comment at the person');
      await new Promise((res) => setTimeout(res, 50));
    }

    // 11: the parts, through E4.b's Apply, untracked.
    const partNamed = (name) => {
      const part = person.pkg.parts.get(name);
      if (!part) throw new Error(`no part ${name}`);
      return part;
    };
    const apply = async (name, change) => {
      const part = partNamed(name);
      const before = await part.getXml();
      const xml = change(before);
      if (xml === before) throw new Error(`${name}: the change found nothing to change`);
      const problems = await checkPartXml(part, xml);
      if (problems.length) throw new Error(`${name}: ${JSON.stringify(problems)}`);
      await applyPartXml(person, part, xml);
    };
    await apply('/word/styles.xml', (xml) => {
      const style = /<w:style [^>]*w:styleId="Heading1"[^>]*>[\s\S]*?<\/w:style>/.exec(xml)[0];
      const own = /<w:color [^>]*\/>/;
      const changed = own.test(style) ? style.replace(own, '<w:color w:val="C00000"/>') : style.replace('<w:rPr>', '<w:rPr><w:color w:val="C00000"/>');
      return xml.replace(style, changed);
    });
    await apply('/word/numbering.xml', (xml) => xml.replace(/w:numFmt w:val="decimal"/g, 'w:numFmt w:val="upperRoman"'));
    await apply('/word/settings.xml', (xml) => (/<w:defaultTabStop /.test(xml)
      ? xml.replace(/<w:defaultTabStop w:val="\d+"\/>/, '<w:defaultTabStop w:val="1440"/>')
      : xml.replace(/(<w:settings[^>]*>)/, '$1<w:defaultTabStop w:val="1440"/>')));

    // 13: comments, untracked.
    const idOf = (text) => person.stories.commentThreads(person.doc).flatMap((t) => [t, ...t.replies]).find((t) => person.stories.currentDoc(t).textContent === text)?.id;
    as(AUTHOR_A);
    await run('range.insertComment', { text: 'A comment by Author A, replied to and resolved.' }, over('Author A comments on this sentence, and Author B answers.'));
    await run('range.insertComment', { text: 'A comment by Author A over two paragraphs.' }, { from: at('A comment by Author A runs'), to: at('over two paragraphs.') + 'over two paragraphs.'.length });
    await run('range.insertComment', { text: 'A comment by Author A, deleted.' }, over('This sentence carries a comment that is deleted.'));
    await run('comment.delete', { id: idOf('A comment by Author A, deleted.') });
    as(AUTHOR_B);
    await run('comment.reply', { id: idOf('A comment by Author A, replied to and resolved.'), text: 'A reply by Author B.' });
    await run('comment.setResolved', { id: idOf('A comment by Author A, replied to and resolved.'), value: true });
    await run('range.insertComment', { text: 'A comment by Author B.' }, over('Author B comments on this sentence alone.'));

    const saved = await person.save();
    const reopened = await EditorDocument.open(saved.bytes);
    await reopened.stories.comments();
    const changes = trackedChangesOf(reopened.doc).map((c) => ({ type: c.type, author: c.author, text: c.text, ...(c.row ? { row: true } : {}), ...(c.move ? { move: c.move } : {}) }));
    const preserved = preservedChangesOf(reopened.doc, (id) => reopened.ctx.ids.get(id)?.value).map((c) => ({ kind: c.kind, author: c.author }));
    const comments = reopened.stories.commentThreads(reopened.doc).flatMap((t) => [
      { author: t.author, text: reopened.stories.currentDoc(t).textContent, done: !!t.done },
      ...t.replies.map((rep) => ({ author: rep.author, text: reopened.stories.currentDoc(rep).textContent, reply: true })),
    ]);
    return { bytes: saved.bytes, selfCheck: saved.selfCheck?.equal ?? false, changes, preserved, comments };
  } finally {
    for (const c of cleanups.splice(0).reverse()) await c();
  }
}

/** The listing as text, for review-for-word.txt and the console. */
export function describe(result) {
  const counts = new Map();
  for (const c of result.changes) counts.set(c.type, (counts.get(c.type) ?? 0) + 1);
  const lines = [
    `review-for-word.docx: ${result.changes.length} tracked changes as the editor lists them (Office JS's form: touching pieces of one type and author are one change; a paragraph mark is "\\r").`,
    `By type: ${[...counts].map(([t, n]) => `${t} ${n}`).join(', ')}. Authors: ${[...new Set(result.changes.map((c) => c.author))].join(', ')}.`,
    '',
    ...result.changes.map((c, i) => `${String(i + 1).padStart(2)}. ${c.type}${c.row ? ' (a row)' : ''} by ${c.author}: ${JSON.stringify(c.text)}`),
    '',
    `Shown in Word only (the editor lists but does not draw them): ${result.preserved.length ? result.preserved.map((c) => `${c.kind} by ${c.author}`).join('; ') : 'none'}.`,
    '',
    `Comments (${result.comments.length}):`,
    ...result.comments.map((c) => `${c.reply ? '    reply ' : '  '}${c.author}${c.done ? ' (resolved)' : ''}: ${c.text}`),
    '',
    `The editor's self-check at save: ${result.selfCheck ? 'passed' : 'FAILED'}.`,
  ];
  return lines.join('\n');
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const here = dirname(fileURLToPath(import.meta.url));
  const out = process.argv[2] ?? join(here, '..', 'test-results', 'word-session');
  await mkdir(out, { recursive: true });
  const result = await buildReviewForWord();
  await writeFile(join(out, 'review-for-word.docx'), result.bytes);
  const text = describe(result);
  await writeFile(join(out, 'review-for-word.txt'), `${text}\n`);
  console.log(`wrote ${join(out, 'review-for-word.docx')}\n\n${text}`);
  process.exit(0);
}
