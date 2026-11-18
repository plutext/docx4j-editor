// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 7, E4.e: comments edited. Step E1, the model: the file's comments listed with
// the `commentThreads` entries laid over them, a comment's fragment, the anchor rules (a comment
// whose anchor the body lost is left out, an orphan at open stays, a reference run alone is an
// anchor), a comment made with the next id or, in a room, a random one that the other peer adopts
// with the maker's ids.
import { describe, test, expect } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { ZipPartStore } from '@docx4j/core-ts';
import { EditorDocument, stateFor, runCommand, runCommandAsync, trackedChangesOf, writeThreadFields, readThreadEntry, commentSourceKey, referenceIdOf, LOCAL_ORIGIN, type CommentThread } from '../src/index.mjs';
import { fixture } from './helpers.mjs';
import { room, RELEASE_ORDERS } from './room-harness.mjs';

/** Each listed comment as `id author: text`, replies indented under their parent. */
function listing(threads: CommentThread[], session: EditorDocument, depth = 0): string[] {
  return threads.flatMap((t) => [`${'  '.repeat(depth)}${t.id} ${t.author}${t.done ? ' (done)' : ''}: ${session.stories.currentDoc(t).textContent}`, ...listing(t.replies, session, depth + 1)]);
}

/** The document with a `comment` mark for `id` over `[from, to)`, applied as a local edit. */
function anchor(session: EditorDocument, id: string, from: number, to: number): void {
  const state = stateFor(session);
  session.replace(state.apply(state.tr.addMark(from, to, session.schema.marks['comment']!.create({ id }))).doc);
}

/** Where a paragraph's text starts, by its text. */
function paragraphStart(doc: PMNode, text: string): number {
  let at = -1;
  doc.forEach((node, pos) => { if (at < 0 && node.type.name === 'paragraph' && node.textContent.includes(text)) at = pos + 1; });
  if (at < 0) throw new Error(`no paragraph with "${text}"`);
  return at;
}

describe('E4.e step E1: comments in the model (ED-005 section 7)', () => {
  test("the file's comments listed as before; Word 15's reply threaded under its parent, marked done", async () => {
    const two = await EditorDocument.open(await fixture('comments-two.docx'));
    expect(listing(await two.stories.comments(), two)).toEqual(['0 jharrop2: One comment', '1 jharrop2: blagh']);
    const word = await EditorDocument.open(await fixture('revisions/revisions-word15.docx'));
    const threads = await word.stories.comments();
    const replied = threads.find((t) => t.replies.length > 0);
    expect(replied).toBeDefined();
    expect(replied!.replies[0]!.author).toBe('Author B');
    expect(threads.some((t) => t.done) || replied!.replies.some((r) => r.done)).toBe(true);
  });

  test("a comment's fragment is made on its first binding and reads back its projection; Share makes every comment's", async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    const [first] = await session.stories.comments();
    expect(session.stories.hasFragment(first!)).toBe(false);
    session.stories.fragmentOf(first!);
    expect(session.stories.hasFragment(first!)).toBe(true);
    expect(session.stories.currentDoc(first!).eq(first!.doc)).toBe(true);
    const shared = await EditorDocument.open(await fixture('comments-two.docx'));
    await shared.share(await fixture('comments-two.docx'), 'comments-two.docx');
    expect([...shared.store.comments.keys()].sort()).toEqual(['0', '1']);
  });

  test("a thread entry lays its fields over the file's: resolved, then deleted and left out; undo takes each back", async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    await session.stories.comments();
    session.undoManager.stopCapturing();
    session.store.ydoc.transact(() => writeThreadFields(session.store, '0', { done: true }), LOCAL_ORIGIN);
    expect(listing(session.stories.commentThreads(session.doc), session)).toEqual(['0 jharrop2 (done): One comment', '1 jharrop2: blagh']);
    session.undoManager.stopCapturing();
    session.store.ydoc.transact(() => writeThreadFields(session.store, '1', { deleted: true }), LOCAL_ORIGIN);
    expect(listing(session.stories.commentThreads(session.doc), session)).toEqual(['0 jharrop2 (done): One comment']);
    session.undoManager.undo();
    expect(listing(session.stories.commentThreads(session.doc), session)).toHaveLength(2);
    session.undoManager.undo();
    expect(readThreadEntry(session.store, '0')?.done ?? false).toBe(false);
  });

  test('a comment whose anchor the body lost is left out, as a save drops it; a comment anchored by its reference run alone is no orphan', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    await session.stories.comments();
    // Comment 1 covers "blsgh " in the second paragraph with its reference after: the paragraph's content deleted untracked.
    const at = paragraphStart(session.doc, 'blsgh');
    const state = stateFor(session);
    const para = state.doc.resolve(at).parent;
    session.replace(state.apply(state.tr.delete(at, at + para.content.size)).doc);
    expect(session.stories.commentThreads(session.doc).map((t) => t.id)).toEqual(['0']);

    // The range markers of comment 0 taken out of the file, its reference run kept: still anchored when opened.
    const source = await EditorDocument.open(await fixture('comments-two.docx'));
    await source.runOnTree(({ body }) => {
      for (const block of body.container.content ?? []) {
        const content = (block.value as { content?: { name?: { localPart?: string }; value?: { id?: unknown } }[] }).content;
        if (!content) continue;
        for (let i = content.length - 1; i >= 0; i--) {
          const item = content[i]!;
          if ((item.name?.localPart === 'commentRangeStart' || item.name?.localPart === 'commentRangeEnd') && String(item.value?.id) === '0') content.splice(i, 1);
        }
      }
    });
    const reopened = await EditorDocument.open((await source.save()).bytes);
    const threads = await reopened.stories.comments();
    expect(threads.map((t) => [t.id, t.orphan])).toEqual([['0', false], ['1', false]]);
  });

  test('a comment made outside a room: the next id, its entry and its element in the store, listed once the body anchors it and not after an undo', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    const made = await session.stories.newComment({ author: 'Reviewer One', text: 'A new comment\nits second line', date: new Date('2026-09-29T10:00:00Z') });
    expect(made.id).toBe('2');
    expect(made.base.initials).toBe('RO');
    expect(readThreadEntry(session.store, '2')).toEqual({ author: 'Reviewer One', initials: 'RO', date: '2026-09-29T10:00:00.000Z', parent: null, done: false });
    expect(session.store.sources.has(commentSourceKey('2'))).toBe(true);
    expect(session.stories.currentDoc(made).childCount).toBe(2);
    // Not anchored yet: not listed.
    expect(session.stories.commentThreads(session.doc).map((t) => t.id)).toEqual(['0', '1']);
    session.undoManager.stopCapturing();
    const at = paragraphStart(session.doc, 'One');
    anchor(session, '2', at, at + 3);
    const listed = session.stories.commentThreads(session.doc);
    expect(listing(listed, session)).toEqual(['0 jharrop2: One comment', '1 jharrop2: blagh', '2 Reviewer One: A new commentits second line']);
    expect(listed.find((t) => t.id === '2')!.isNew).toBe(true);
    session.undoManager.undo();
    expect(session.stories.commentThreads(session.doc).map((t) => t.id)).toEqual(['0', '1']);
  });

  describe('in a room: a comment made has a random id, and the other peer adopts it with the maker\'s ids', () => {
    for (const order of RELEASE_ORDERS) {
      test(order, async () => {
        const r = await room(await fixture('comments-two.docx'));
        await Promise.all([r.creator.stories.comments(), r.joiner.stories.comments()]);
        r.link.hold();
        const mine = await r.creator.stories.newComment({ author: 'Author A', text: 'from A' });
        anchor(r.creator, mine.id, paragraphStart(r.creator.doc, 'One'), paragraphStart(r.creator.doc, 'One') + 3);
        const theirs = await r.joiner.stories.newComment({ author: 'Author B', text: 'from B' });
        anchor(r.joiner, theirs.id, paragraphStart(r.joiner.doc, 'blsgh'), paragraphStart(r.joiner.doc, 'blsgh') + 4);
        expect(Number(mine.id)).toBeGreaterThanOrEqual(2 ** 20);
        expect(mine.id).not.toBe(theirs.id);
        r.link.release(order);
        const [a, b] = await Promise.all([r.creator.stories.comments(), r.joiner.stories.comments()]);
        expect(listing(a, r.creator)).toEqual(listing(b, r.joiner));
        expect(listing(a, r.creator)).toEqual(expect.arrayContaining([`${mine.id} Author A: from A`, `${theirs.id} Author B: from B`]));
        // The adopted record projects with the maker's ids, so the two stories' documents are the same.
        expect(r.joiner.stories.commentOf(mine.id)!.doc.eq(mine.doc)).toBe(true);
        expect(r.creator.stories.commentOf(theirs.id)!.doc.eq(theirs.doc)).toBe(true);
      });
    }
  });
});

// ---------------------------------------------------------------- step E2: the commands

/** The comment ids each text node carries and each reference atom stands for, in document order, as `text[ids]` and `@id`. */
function anchorsIn(session: EditorDocument, node: PMNode): string[] {
  const out: string[] = [];
  node.descendants((n) => {
    const ref = referenceIdOf(n, session.ctx.ids);
    if (ref !== undefined) out.push(`@${ref}`);
    else if (n.isText) {
      const ids = n.marks.filter((m) => m.type.name === 'comment').map((m) => String(m.attrs['id']));
      out.push(ids.length > 0 ? `${n.text}[${ids.join(',')}]` : n.text!);
    }
    return true;
  });
  return out;
}

/** A session over `tracked-changes.docx` whose "Some text." paragraph holds `text` instead, tracking off. */
async function withParagraph(text: string): Promise<{ session: EditorDocument; index: number; start: number }> {
  const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
  session.trackingAuthor = 'Reviewer One';
  session.trackedChangeDate = new Date('2026-09-29T10:00:00Z');
  let index = -1;
  let start = -1;
  session.doc.forEach((node, pos, i) => { if (index < 0 && node.textContent === 'Some text.') { index = i; start = pos + 1; } });
  runCommand(session, 'insertText', { text }, { from: start, to: start + 'Some text.'.length });
  return { session, index, start };
}

const TEXT_WITH_SPACES = '          Hello, the cat                    sat on the mat';

describe('E4.e step E2: the commands (ED-005 section 7.3)', () => {
  const TEXT = TEXT_WITH_SPACES;

  test("New Comment at a caret covers a word as Word's does (Jason's measurements, proposal 32)", async () => {
    const cases: [string, number, string][] = [
      ['inside a word', TEXT.indexOf('cat') + 1, 'cat'],
      ['touching the word after', TEXT.indexOf('sat'), 'sat'],
      ['just after a word', TEXT.indexOf('cat') + 3, 'cat'],
      ['among the spaces between words, not touching either', TEXT.indexOf('sat') - 5, 'cat'],
      ['one space before the word after, touching it', TEXT.indexOf('sat') - 0, 'sat'],
      ['after a comma, before a space', TEXT.indexOf(',') + 1, 'Hello'],
    ];
    for (const [what, offset, word] of cases) {
      const { session, index, start } = await withParagraph(TEXT);
      const r = await runCommandAsync(session, 'range.insertComment', { text: 'x' }, { from: start + offset });
      expect(r.ok, what).toBe(true);
      const marked = anchorsIn(session, session.doc.child(index)).filter((s) => s.includes('['));
      expect(marked.map((s) => s.slice(0, s.indexOf('['))), what).toEqual([word]);
    }
  });

  test("with no word before the caret, an empty range at the paragraph's start; a second one there goes before the first (to be nested at save, as Word writes them)", async () => {
    const { session, index, start } = await withParagraph(TEXT);
    await runCommandAsync(session, 'range.insertComment', { text: 'first' }, { from: start + 3 });
    await runCommandAsync(session, 'range.insertComment', { text: 'second' }, { from: start + 5 + 1 });
    const [first, second] = session.stories.commentThreads(session.doc).filter((t) => t.isNew).map((t) => t.id);
    expect(anchorsIn(session, session.doc.child(index)).slice(0, 3)).toEqual([`@${second}`, `@${first}`, TEXT]);
  });

  test('a selection within a paragraph and across two: the marks over it, the reference after its end, the comment listed, one undo step', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    session.trackingAuthor = 'Reviewer One';
    const one = paragraphStart(session.doc, 'One');
    const r = await runCommandAsync(session, 'range.insertComment', { text: 'On "One"' }, { from: one, to: one + 3 });
    expect(r.ok).toBe(true);
    expect(anchorsIn(session, session.doc.child(0))).toEqual(['One[2]', '@2', ' ', 'paragraph[0]', '@0']);
    // The comment's own `sources` entry is the stories' to adopt, not a block the session registers: it reported
    // "the projection minted 1 ids the entry did not name" on every comment made (ED-005 section 12.96).
    await session.sourcesReady();
    expect(session.roomProblems).toEqual([]);
    const blsgh = paragraphStart(session.doc, 'blsgh');
    // "paragraph" starts after "One", the new reference and the space.
    await runCommandAsync(session, 'range.insertComment', { text: 'across' }, { from: one + 5, to: blsgh + 3 });
    expect(anchorsIn(session, session.doc.child(0))).toEqual(['One[2]', '@2', ' ', 'paragraph[0,3]', '@0']);
    expect(anchorsIn(session, session.doc.child(1))).toEqual(['bls[1,3]', '@3', 'gh [1]', '@1']);
    expect(session.stories.commentThreads(session.doc).map((t) => `${t.id} ${t.author}`)).toEqual(['0 jharrop2', '1 jharrop2', '2 Reviewer One', '3 Reviewer One']);
    session.undoManager.undo();
    expect(session.stories.commentThreads(session.doc).map((t) => t.id)).toEqual(['0', '1', '2']);
    expect(anchorsIn(session, session.doc.child(1))).toEqual(['blsgh [1]', '@1']);
  });

  test('with tracking on, commenting is untracked and says nothing: no revision on the reference, no untracked notice', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Reviewer One';
    const one = paragraphStart(session.doc, 'One');
    const r = await runCommandAsync(session, 'range.insertComment', { text: 'tracked?' }, { from: one, to: one + 3 });
    expect(r.ok).toBe(true);
    expect(r.untracked).toBeUndefined();
    let revisions = 0;
    session.doc.descendants((n) => { if (n.attrs['revision'] || n.marks.some((m) => m.type.name === 'insertion' || m.type.name === 'deletion')) revisions++; return true; });
    expect(revisions).toBe(0);
  });

  test("a reply covers its thread's text, its reference after the thread's; a reply to a reply joins the thread", async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    session.trackingAuthor = 'Reviewer One';
    expect((await runCommandAsync(session, 'comment.reply', { id: '0', text: 'a reply' })).ok).toBe(true);
    expect(anchorsIn(session, session.doc.child(0))).toEqual(['One ', 'paragraph[0,2]', '@0', '@2']);
    session.trackingAuthor = 'Reviewer Two';
    expect((await runCommandAsync(session, 'comment.reply', { id: '2', text: 'and another' })).ok).toBe(true);
    expect(anchorsIn(session, session.doc.child(0))).toEqual(['One ', 'paragraph[0,2,3]', '@0', '@3', '@2']);
    expect(listing(session.stories.commentThreads(session.doc), session)).toEqual(['0 jharrop2: One comment', '  2 Reviewer One: a reply', '  3 Reviewer Two: and another', '1 jharrop2: blagh']);
  });

  test('resolve toggles the thread (a reply resolves its thread); delete takes a thread with its replies, or a reply alone; each one undo step', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    await runCommandAsync(session, 'comment.reply', { id: '0', text: 'a reply' });
    expect((await runCommandAsync(session, 'comment.setResolved', { id: '2' })).ok).toBe(true);
    expect(session.stories.commentThreads(session.doc)[0]!.done).toBe(true);
    expect((await runCommandAsync(session, 'comment.setResolved', { id: '0', value: false })).ok).toBe(true);
    expect(session.stories.commentThreads(session.doc)[0]!.done).toBe(false);
    session.undoManager.undo();
    expect(session.stories.commentThreads(session.doc)[0]!.done).toBe(true);
    // A reply alone.
    await runCommandAsync(session, 'comment.delete', { id: '2' });
    expect(listing(session.stories.commentThreads(session.doc), session)).toEqual(['0 jharrop2 (done): One comment', '1 jharrop2: blagh']);
    expect(anchorsIn(session, session.doc.child(0))).toEqual(['One ', 'paragraph[0]', '@0']);
    session.undoManager.undo();
    expect(listing(session.stories.commentThreads(session.doc), session)).toHaveLength(3);
    // The thread.
    await runCommandAsync(session, 'comment.delete', { id: '0' });
    expect(listing(session.stories.commentThreads(session.doc), session)).toEqual(['1 jharrop2: blagh']);
    expect(anchorsIn(session, session.doc.child(0))).toEqual(['One ', 'paragraph']);
    session.undoManager.undo();
    expect(anchorsIn(session, session.doc.child(0))).toEqual(['One ', 'paragraph[0,2]', '@0', '@2']);
  });

  test('delete with tracking on removes the reference for real, as Word does', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    session.trackingMode = 'TrackAll';
    await session.stories.comments();
    expect((await runCommandAsync(session, 'comment.delete', { id: '1' })).ok).toBe(true);
    expect(anchorsIn(session, session.doc.child(1))).toEqual(['blsgh ']);
  });
});

// ---------------------------------------------------------------- step E3: the save

/** A saved package's part as XML, rsids dropped; '' when it has no such part. */
async function partXml(bytes: Uint8Array, name: string): Promise<string> {
  const store = new ZipPartStore(bytes);
  try { return new TextDecoder().decode(store.loadSync(name.replace(/^\//, ''))).replace(/ w:rsid\w*="[^"]*"/g, ''); } catch { return ''; }
}

/** The body's markers and reference runs in order, as `s0`, `e0`, `r0`, with the text between as `"text"`. */
function markerOrder(xml: string, paragraphText: string): string[] {
  const paragraph = [...xml.matchAll(/<w:p[ >](?:(?!<\/w:p>).)*<\/w:p>/g)].map((m) => m[0]).find((p) => [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('').includes(paragraphText));
  if (!paragraph) return [];
  const out: string[] = [];
  for (const m of paragraph.matchAll(/<w:commentRange(Start|End) w:id="(\d+)"\/>|<w:commentReference w:id="(\d+)"\/>|<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)) {
    if (m[1]) out.push(`${m[1] === 'Start' ? 's' : 'e'}${m[2]}`);
    else if (m[3]) out.push(`r${m[3]}`);
    else out.push(JSON.stringify(m[4]));
  }
  return out;
}

async function saved(session: EditorDocument): Promise<Uint8Array> {
  const result = await session.save();
  expect(result.selfCheck?.equal, JSON.stringify(result.selfCheck?.difference)).toBe(true);
  return result.bytes;
}

describe('E4.e step E3: the save (ED-005 section 7.4)', () => {
  test('a comment made on a selection: its markers around the text and its reference after, the comment with its author and text, its side-part entries; reopened, it is listed', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    session.trackingAuthor = 'Reviewer One';
    session.trackedChangeDate = new Date('2026-09-29T10:00:00Z');
    const one = paragraphStart(session.doc, 'One');
    await runCommandAsync(session, 'range.insertComment', { text: 'On "One"' }, { from: one, to: one + 3 });
    const bytes = await saved(session);
    const main = await partXml(bytes, '/word/document.xml');
    expect(markerOrder(main, 'One')).toEqual(['s2', '"One"', 'e2', 'r2', '" "', 's0', '"paragraph"', 'e0', 'r0']);
    expect(main).toMatch(/<w:r><w:rPr><w:rStyle w:val="CommentReference"\/><\/w:rPr><w:commentReference w:id="2"\/><\/w:r>/);
    const comments = await partXml(bytes, '/word/comments.xml');
    expect(comments).toMatch(/<w:comment [^>]*w:id="2"[^>]*>/);
    expect(comments).toMatch(/w:author="Reviewer One"/);
    expect(comments).toMatch(/w:initials="RO"/);
    expect(comments).toMatch(/<w:p [^>]*w14:paraId="([0-9A-F]{8})"[^>]*><w:pPr><w:pStyle w:val="CommentText"\/><\/w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"\/><\/w:rPr><w:annotationRef\/><\/w:r><w:r><w:t>On "One"<\/w:t><\/w:r><\/w:p>/);
    const paraId = /w14:paraId="([0-9A-F]{8})"[^>]*><w:pPr><w:pStyle w:val="CommentText"\/><\/w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"\/><\/w:rPr><w:annotationRef\/><\/w:r><w:r><w:t>On "One"/.exec(comments)![1];
    expect(await partXml(bytes, '/word/commentsExtended.xml')).toMatch(new RegExp(`<w15:commentEx w15:done="0" w15:paraId="${paraId}"/>`));
    expect(await partXml(bytes, '/word/people.xml')).toMatch(/w15:author="Reviewer One"/);
    expect(await partXml(bytes, '/word/styles.xml')).toMatch(/w:styleId="CommentText"/);
    const reopened = await EditorDocument.open(bytes);
    expect(listing(await reopened.stories.comments(), reopened)).toEqual(['0 jharrop2: One comment', '1 jharrop2: blagh', '2 Reviewer One: On "One"']);
  });

  test('a comment across two paragraphs: the start in the first, the end and the reference in the last', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    const one = paragraphStart(session.doc, 'One');
    const blsgh = paragraphStart(session.doc, 'blsgh');
    await runCommandAsync(session, 'range.insertComment', { text: 'across' }, { from: one + 4, to: blsgh + 3 });
    const main = await partXml(await saved(session), '/word/document.xml');
    expect(markerOrder(main, 'One')).toEqual(['"One "', 's0', 's2', '"paragraph"', 'e0', 'r0']);
    expect(markerOrder(main, 'gh')).toEqual(['s1', '"bls"', 'e2', 'r2', '"gh "', 'e1', 'r1']);
  });

  test("a reply: its markers follow its parent's, as the engine writes them, and its commentEx names its parent's paragraph", async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    await runCommandAsync(session, 'comment.reply', { id: '0', text: 'a reply' });
    const bytes = await saved(session);
    expect(markerOrder(await partXml(bytes, '/word/document.xml'), 'One')).toEqual(['"One "', 's0', 's2', '"paragraph"', 'e0', 'e2', 'r0', 'r2']);
    const comments = await partXml(bytes, '/word/comments.xml');
    const parentParaId = /<w:comment [^>]*w:id="0"[^>]*><w:p [^>]*w14:paraId="([0-9A-F]{8})"/.exec(comments)?.[1];
    expect(parentParaId).toBeDefined();
    expect(await partXml(bytes, '/word/commentsExtended.xml')).toMatch(new RegExp(`w15:paraIdParent="${parentParaId}"`));
    const reopened = await EditorDocument.open(bytes);
    expect(listing(await reopened.stories.comments(), reopened)).toEqual(['0 jharrop2: One comment', '  2 docx4j: a reply', '1 jharrop2: blagh']);
  });

  test('resolve, delete and an edit of a file comment\'s text, saved and reopened; the deleted one leaves no marker', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    await runCommandAsync(session, 'comment.setResolved', { id: '0', value: true });
    await runCommandAsync(session, 'comment.delete', { id: '1' });
    const [first] = session.stories.commentThreads(session.doc);
    const doc = session.stories.currentDoc(first!);
    const end = doc.child(0).nodeSize - 1;
    session.stories.replace(first!, EditorState.create({ doc }).apply(EditorState.create({ doc }).tr.insertText(', edited', end)).doc);
    const bytes = await saved(session);
    const main = await partXml(bytes, '/word/document.xml');
    expect(main).not.toMatch(/w:id="1"/);
    const reopened = await EditorDocument.open(bytes);
    expect(listing(await reopened.stories.comments(), reopened)).toEqual(['0 jharrop2 (done): One comment, edited']);
  });

  // ED-005 section 12.47 (found writing word-divergences.md's entry for section 12.33 item 2): with no text left carrying
  // a file comment, its start was put at its reference, after the end the export had just put before it, so the range
  // ran on to the paragraph's end; and with the reference the paragraph's first item, the end went to the paragraph's
  // end. Either way the save's self-check refused. Word keeps the markers, collapsed, before the reference.
  test("a file comment whose text is deleted keeps its markers, collapsed before its reference", async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    await session.stories.comments();
    const start = paragraphStart(session.doc, 'One paragraph');
    expect(runCommand(session, 'insertText', { text: '' }, { from: start + 4, to: start + 13 }).ok).toBe(true);
    expect(markerOrder(await partXml(await saved(session), '/word/document.xml'), 'One')).toEqual(['"One "', 's0', 'e0', 'r0']);
  });

  // Word check 32 (ED-005 section 12.49) overturned this case's first expectation (markers kept, collapsed): the
  // editor's comment has its start marker inside the insertion, as Word's own form does, and Word removes such a
  // comment, reference and all, when the insertion is rejected. The reference-first path of the fix above is held
  // by check32.test.mts, where a file put the markers outside the revision and the comment stays.
  test("a comment on a word of an insertion at a paragraph's start, saved, then the insertion rejected: the comment goes with it, as Word removes it", async () => {
    const { session, start } = await withParagraph('Some text.');
    session.trackingMode = 'TrackAll';
    expect(runCommand(session, 'insertText', { text: 'alpha beta gamma ' }, { from: start }).ok).toBe(true);
    await runCommandAsync(session, 'range.insertComment', { text: 'On beta.' }, { from: start + 6, to: start + 10 });
    const reopened = await EditorDocument.open(await saved(session));
    await reopened.stories.comments();
    // One change across the plain reference run, as Office JS lists it (check 32).
    const insertion = trackedChangesOf(reopened.doc).find((c) => c.type === 'Added' && c.text.startsWith('alpha'))!;
    expect(insertion.text).toBe('alpha beta gamma ');
    expect((await runCommandAsync(reopened, 'trackedChange.reject', { key: reopened.changeKey(insertion) })).ok).toBe(true);
    const bytes = await saved(reopened);
    expect(markerOrder(await partXml(bytes, '/word/document.xml'), 'Some text.')).toEqual(['"Some text."']);
    const again = await EditorDocument.open(bytes);
    expect(listing(await again.stories.comments(), again)).toEqual([]);
  });

  test("empty ranges at a paragraph's start nest as Word writes them (Jason's XML: s0 s1 e1 r1 e0 r0)", async () => {
    const { session, start } = await withParagraph(TEXT_WITH_SPACES);
    await runCommandAsync(session, 'range.insertComment', { text: 'first' }, { from: start + 3 });
    await runCommandAsync(session, 'range.insertComment', { text: 'second' }, { from: start + 3 });
    const [a, b] = session.stories.commentThreads(session.doc).filter((t) => t.isNew).map((t) => t.id);
    expect(markerOrder(await partXml(await saved(session), '/word/document.xml'), 'Hello')).toEqual([`s${a}`, `s${b}`, `e${b}`, `r${b}`, `e${a}`, `r${a}`, JSON.stringify(TEXT_WITH_SPACES)]);
  });

  test('a document with no comments gets its comment parts and styles; an unedited one with comments writes nothing', async () => {
    const { session, start } = await withParagraph('Some words here');
    await runCommandAsync(session, 'range.insertComment', { text: 'the first' }, { from: start + 5, to: start + 10 });
    const bytes = await saved(session);
    for (const part of ['/word/comments.xml', '/word/commentsExtended.xml', '/word/people.xml', '/word/commentsIds.xml']) expect(await partXml(bytes, part), part).not.toBe('');
    expect(await partXml(bytes, '/word/styles.xml')).toMatch(/w:styleId="CommentReference"/);
    const untouched = await EditorDocument.open(await fixture('comments-two.docx'));
    await untouched.stories.comments();
    const result = await untouched.save();
    expect(result.changed).toBe(false);
  });

  test('a second save after the new comment was typed in keeps its paragraph id, so its side entries stay its own', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    const one = paragraphStart(session.doc, 'One');
    await runCommandAsync(session, 'range.insertComment', { text: 'draft' }, { from: one, to: one + 3 });
    const first = await partXml(await saved(session), '/word/comments.xml');
    const record = session.stories.commentThreads(session.doc).find((t) => t.id === '2')!;
    const doc = session.stories.currentDoc(record);
    session.stories.replace(record, EditorState.create({ doc }).apply(EditorState.create({ doc }).tr.insertText(' two', doc.child(0).nodeSize - 1)).doc);
    const second = await partXml(await saved(session), '/word/comments.xml');
    const paraIdOf = (xml: string): string | undefined => /<w:comment [^>]*w:id="2"[^>]*><w:p [^>]*w14:paraId="([0-9A-F]{8})"/.exec(xml)?.[1];
    expect(paraIdOf(first)).toBeDefined();
    expect(paraIdOf(second)).toBe(paraIdOf(first));
    expect(second).toMatch(/draft two/);
  });

  test('in a room: the comments made there are written with small ids, the same on both peers, and the session keeps its random ones', async () => {
    const r = await room(await fixture('comments-two.docx'));
    await Promise.all([r.creator.stories.comments(), r.joiner.stories.comments()]);
    const one = paragraphStart(r.creator.doc, 'One');
    await runCommandAsync(r.creator, 'range.insertComment', { text: 'from A' }, { from: one, to: one + 3 });
    const blsgh = paragraphStart(r.joiner.doc, 'blsgh');
    await runCommandAsync(r.joiner, 'range.insertComment', { text: 'from B' }, { from: blsgh, to: blsgh + 3 });
    await Promise.all([r.creator.stories.comments(), r.joiner.stories.comments()]);
    // Each peer adopted the other's comment with nothing to report (section 12.96).
    await Promise.all([r.creator.sourcesReady(), r.joiner.sourcesReady()]);
    expect([r.creator.roomProblems, r.joiner.roomProblems]).toEqual([[], []]);
    const [a, b] = [await saved(r.creator), await saved(r.joiner)];
    const ids = async (bytes: Uint8Array): Promise<string[]> => [...(await partXml(bytes, '/word/comments.xml')).matchAll(/<w:comment [^>]*w:id="(\d+)"/g)].map((m) => m[1]!).sort();
    expect(await ids(a)).toEqual(['0', '1', '2', '3']);
    expect(await ids(b)).toEqual(['0', '1', '2', '3']);
    expect(markerOrder(await partXml(a, '/word/document.xml'), 'One')).toEqual(markerOrder(await partXml(b, '/word/document.xml'), 'One'));
    expect(r.creator.stories.commentThreads(r.creator.doc).some((t) => Number(t.id) >= 2 ** 20)).toBe(true);
  });
});

describe('E4.e step E3: comments in a tree run (ED-005 section 7.4)', () => {
  test('a script sees the comments on screen, resolves one, makes another and changes a third\'s text; the session follows, and its save keeps them all', async () => {
    const session = await EditorDocument.open(await fixture('comments-two.docx'));
    session.trackingAuthor = 'Reviewer One';
    const one = paragraphStart(session.doc, 'One');
    await runCommandAsync(session, 'range.insertComment', { text: 'made here' }, { from: one, to: one + 3 });
    const run = await session.runOnTree(async ({ body }) => {
      const all = (await body.getComments()).flatMap((c) => [c, ...c.replies]);
      const seen = all.map((c) => `${c.id} ${c.content}`);
      all.find((c) => c.id === 2)!.resolved = true;
      all.find((c) => c.id === 1)!.content = 'blagh, rewritten';
      const last = body.paragraphs[body.paragraphs.length - 1]!;
      await last.getRange().insertComment('made by a script', { author: { name: 'Script' } });
      return seen;
    });
    expect(run.error).toBeUndefined();
    // In document order, as the engine lists them: comment 2 is on "One".
    expect(run.value).toEqual(['2 made here', '0 One comment', '1 blagh']);
    const listed = listing(session.stories.commentThreads(session.doc), session);
    expect(listed).toEqual(['0 jharrop2: One comment', '1 jharrop2: blagh, rewritten', '2 Reviewer One (done): made here', '3 Script: made by a script']);
    const bytes = await saved(session);
    const reopened = await EditorDocument.open(bytes);
    expect(listing(await reopened.stories.comments(), reopened)).toEqual(listed);
  });
});

describe('E4.e step E5: two peers commenting at once (ED-005 section 7.8)', () => {
  for (const order of RELEASE_ORDERS) {
    test(`both reply to one thread, one resolves it while the other deletes another thread; both converge and save the same comments (${order})`, async () => {
      const r = await room(await fixture('comments-two.docx'));
      r.creator.trackingAuthor = 'Author A';
      r.joiner.trackingAuthor = 'Author B';
      await Promise.all([r.creator.stories.comments(), r.joiner.stories.comments()]);
      r.link.hold();
      expect((await runCommandAsync(r.creator, 'comment.reply', { id: '0', text: 'from A' })).ok).toBe(true);
      expect((await runCommandAsync(r.creator, 'comment.setResolved', { id: '0', value: true })).ok).toBe(true);
      expect((await runCommandAsync(r.joiner, 'comment.reply', { id: '0', text: 'from B' })).ok).toBe(true);
      expect((await runCommandAsync(r.joiner, 'comment.delete', { id: '1' })).ok).toBe(true);
      r.link.release(order);
      const [a, b] = await Promise.all([r.creator.stories.comments(), r.joiner.stories.comments()]);
      const shown = (threads: CommentThread[], s: EditorDocument): string[] => listing(threads, s).map((line) => line.replace(/^(\s*)\d+ /, '$1'));
      expect(shown(a, r.creator)).toEqual(shown(b, r.joiner));
      expect(shown(a, r.creator)[0]).toBe('jharrop2 (done): One comment');
      expect(shown(a, r.creator)).toEqual(expect.arrayContaining(['  Author A: from A', '  Author B: from B']));
      expect(shown(a, r.creator).some((line) => line.includes('blagh'))).toBe(false);
      const comments = async (s: EditorDocument): Promise<string[]> => {
        const xml = await partXml(await saved(s), '/word/comments.xml');
        return [...xml.matchAll(/<w:comment [^>]*w:author="([^"]*)"[^>]*w:id="(\d+)"|<w:comment [^>]*w:id="(\d+)"[^>]*w:author="([^"]*)"/g)].map((m) => `${m[2] ?? m[3]} ${m[1] ?? m[4]}`).sort();
      };
      expect(await comments(r.creator)).toEqual(await comments(r.joiner));
    });
  }
});

describe("Jason's check in Word (ED-005 section 12.23): comments-for-word.docx as the editor wrote it and as Word 15 and Word 2010 saved it", () => {
  /** Each listed comment as `author: text`, done marked, replies indented. */
  const shape = (threads: CommentThread[], s: EditorDocument, depth = 0): string[] => threads.flatMap((t) => [`${'  '.repeat(depth)}${t.author}${t.done ? ' (done)' : ''}: ${s.stories.currentDoc(t).textContent}`, ...shape(t.replies, s, depth + 1)]);
  const A1 = 'A comment by Author A, replied to and resolved.';
  const B1 = 'A comment by Author B, with a reply by Author A.';
  const FOUR = 'A comment by Author A over four paragraphs.';

  test('what the editor wrote: two threads, the first resolved, and the comment over four paragraphs', async () => {
    const s = await EditorDocument.open(await fixture('comments/comments-for-word.docx'));
    expect(shape(await s.stories.comments(), s).sort()).toEqual([`Author A (done): ${A1}`, `Author A: ${FOUR}`, `Author B: ${B1}`, '  Author A: A reply by Author A.', '  Author B: A reply by Author B.'].sort());
  });

  test("Word 15's save: every comment, each reply under its comment, the ids renumbered in document order; the resolved state not kept (section 12.23 item 3)", async () => {
    const s = await EditorDocument.open(await fixture('comments/comments-for-word-word15.docx'));
    expect(shape(await s.stories.comments(), s)).toEqual([`Author A: ${A1}`, '  Author B: A reply by Author B.', `Author A: ${FOUR}`, `Author B: ${B1}`, '  Author A: A reply by Author A.']);
    expect((await s.stories.comments()).map((t) => t.id)).toEqual(['0', '3', '4']);
  });

  test("Word 2010's save: every comment, each on its own (Word 2010 writes no paragraph ids, so the threads' links name nothing)", async () => {
    const s = await EditorDocument.open(await fixture('comments/comments-for-word-word2010.docx'));
    expect(shape(await s.stories.comments(), s)).toEqual([`Author A: ${A1}`, 'Author B: A reply by Author B.', `Author A: ${FOUR}`, `Author B: ${B1}`, 'Author A: A reply by Author A.']);
  });

  test("Word 15 keeps the resolved state saving in compatibility mode, and clears it converting the file to its own format (Jason's second opening, section 12.23 item 3)", async () => {
    for (const [name, done] of [['comments-for-word-word15-compat', true], ['comments-for-word-word15-upgraded', false]] as const) {
      const s = await EditorDocument.open(await fixture(`comments/${name}.docx`));
      expect(shape(await s.stories.comments(), s), name).toEqual([`Author A${done ? ' (done)' : ''}: ${A1}`, '  Author B: A reply by Author B.', `Author A: ${FOUR}`, `Author B: ${B1}`, '  Author A: A reply by Author A.']);
    }
  });

  test("current Word (Microsoft 365) shows the resolved state and does not save it, in compatibility mode or converted, whoever serialised it (Jason's check, section 12.23 item 7)", async () => {
    for (const name of ['comments-for-word-word365-compat', 'comments-for-word-word365-upgraded', 'comments-for-word-word15-compat-word365']) {
      const s = await EditorDocument.open(await fixture(`comments/${name}.docx`));
      expect(shape(await s.stories.comments(), s), name).toEqual([`Author A: ${A1}`, '  Author B: A reply by Author B.', `Author A: ${FOUR}`, `Author B: ${B1}`, '  Author A: A reply by Author A.']);
    }
  });

  for (const name of ['comments-for-word-word15', 'comments-for-word-word2010', 'comments-for-word-word15-compat', 'comments-for-word-word15-upgraded', 'comments-for-word-word365-compat', 'comments-for-word-word365-upgraded', 'comments-for-word-word15-compat-word365']) {
    test(`${name}: an unedited save writes nothing; a resolve and a reply save with the self-check and reopen as made`, async () => {
      const bytes = await fixture(`comments/${name}.docx`);
      const untouched = await EditorDocument.open(bytes);
      await untouched.stories.comments();
      expect((await untouched.save()).changed).toBe(false);
      const s = await EditorDocument.open(bytes);
      s.trackingAuthor = 'Reviewer';
      const [first, second] = await s.stories.comments();
      expect((await runCommandAsync(s, 'comment.setResolved', { id: first!.id, value: true })).ok).toBe(true);
      expect((await runCommandAsync(s, 'comment.reply', { id: second!.id, text: 'After Word' })).ok).toBe(true);
      const back = await EditorDocument.open(await saved(s));
      const threads = await back.stories.comments();
      expect(threads[0]!.done).toBe(true);
      expect(threads.find((t) => t.id === second!.id)!.replies.map((r) => `${r.author}: ${back.stories.currentDoc(r).textContent}`)).toContain('Reviewer: After Word');
    });
  }
});

describe('a comment a tree run makes, in a room (ED-005 section 12.29: found by E4.f step F5)', () => {
  test('published as the editor\'s own are: the other peer adopts it, with its author, anchor and text', async () => {
    const r = await room(await fixture('tracked-changes.docx'));
    const run = await r.creator.runOnTree(async ({ body }) => {
      const p = body.paragraphs.find((q) => q.text === 'Some text.')!;
      await p.getRange('Whole').insertComment('From a script.', { author: { name: 'Claude' }, date: new Date('2026-09-29T10:00:00Z'), id: 1234567 });
    }, undefined, { author: 'Claude', mode: 'Off' });
    expect(run.error).toBeUndefined();
    await r.joiner.sourcesReady();
    await r.joiner.stories.comments();
    const thread = r.joiner.stories.commentThreads(r.joiner.doc).find((t) => t.id === '1234567');
    expect(thread).toMatchObject({ author: 'Claude', parentId: null, done: false });
    expect(r.joiner.stories.currentDoc(thread!.record).textContent).toBe('From a script.');
    // Both peers' saves write it, once.
    const [a, b] = [await r.creator.save(), await r.joiner.save()];
    for (const saved of [a, b]) {
      const xml = new TextDecoder().decode(new ZipPartStore(saved.bytes).loadSync('word/comments.xml'));
      expect(xml.match(/<w:comment /g)?.length).toBe(1);
      expect(xml).toContain('From a script.');
    }
  });
});
