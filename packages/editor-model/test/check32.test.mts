// SPDX-License-Identifier: GPL-3.0-or-later
//
// core-ts check 32 (run by Jason in Word 16.0.20430.20118, 2026-10-02; `fixtures/revisions/check32/`, copied from
// docx4j-core-ts's `test/fixtures/revisions/check32/`, its commit 8cfdef9), asked by the editor (ED-005 section 12.49):
// - case A: a pending insertion holding a comment's word, in Word's form (A1) and the editor's (A2, which Word's
//   insertOoxml turned into A1's form on the way in), rejected alone, by Reject All, and accepted; a deletion in A1's
//   shape (A3) accepted. Word removes the comment whose text goes: its markers, its reference run (plain, outside the
//   revision) and its entry. Office JS lists the two insertions either side of the plain reference run as one change;
// - case B: typing into a link that was already there, with tracking on: "x" in the middle lands inside the link,
//   "y" at its end after it, "z" at its start before it; three changes. Word's save keeps the link as a HYPERLINK field.
// The editor is held to both. Office JS's `range.insertText` at the link's ends (inside, case B's API half) is a
// range-relative insertion the editor's caret-based typing does not make, and is not compared.
import { describe, expect, test } from 'vitest';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { WordprocessingMLPackage, CommentsPart, ZipPartStore } from '@docx4j/core-ts';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, runCommandAsync, trackedChangesOf } from '../src/index.mjs';
import { fixture, fixturesDir } from './helpers.mjs';

interface Listed { type: string; author: string; text: string }
interface WordComment { content: string; author: string; anchoredText: string }
interface State { markup: string[]; listed: Listed[] | null; comments: WordComment[] | null }
interface CaseA { before: State; thatOneWas: { type: string; action: 'accept' | 'reject' }; thatOne: State; rejectAll: State; acceptAll: State }

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A document whose body is the markup Word left, with the comments its markers name ("Reviewer", the check's texts). */
async function documentOf(markup: readonly string[], comments: Record<string, string>): Promise<Uint8Array> {
  const pkg = await WordprocessingMLPackage.createPackage();
  const main = pkg.getMainDocumentPart();
  main.setXml(`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${markup.join('')}<w:sectPr/></w:body></w:document>`);
  const ids = Object.keys(comments);
  if (ids.length > 0) {
    const part = new CommentsPart();
    main.addTargetPart(part);
    part.setXml(`<w:comments xmlns:w="${W}">${ids.map((id) => `<w:comment w:id="${id}" w:author="Reviewer" w:initials="R" w:date="2026-10-02T09:00:00Z"><w:p><w:r><w:t>${comments[id]}</w:t></w:r></w:p></w:comment>`).join('')}</w:comments>`);
  }
  return new Uint8Array(await pkg.save());
}

/** The comment ids a markup's references name, with the check's text for each. */
function commentsIn(markup: readonly string[], text: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of markup.join('').matchAll(/<w:commentReference w:id="(\d+)"\/>/g)) out[m[1]!] = text;
  return out;
}

/** A document as its blocks' text: what Word's result is compared by. */
function shape(doc: PMNode): string[] {
  const out: string[] = [];
  doc.forEach((b) => out.push(b.textContent));
  return out;
}

/** The comments as Word's `body.getComments()` gave them: content, author, the text the comment covers. */
async function commentsOf(session: EditorDocument): Promise<WordComment[]> {
  const threads = await session.stories.comments();
  return threads.filter((t) => !t.orphan).map((t) => {
    let anchoredText = '';
    session.doc.descendants((n) => { if (n.isText && n.marks.some((m) => m.type.name === 'comment' && String(m.attrs['id']) === t.id)) anchoredText += n.text; return true; });
    return { content: session.stories.currentDoc(t).textContent, author: t.author, anchoredText };
  });
}

/** Saved and opened again, as Word's result was read. */
async function reopened(session: EditorDocument): Promise<EditorDocument> {
  const result = await session.save();
  expect(result.selfCheck?.equal, JSON.stringify(result.selfCheck?.difference)).toBe(true);
  return EditorDocument.open(result.bytes);
}

describe('core-ts check 32, case A: a comment whose text a resolution takes goes with it, as Word removes it', async () => {
  const result = JSON.parse(await readFile(join(fixturesDir, 'revisions', 'check32', 'result.json'), 'utf8')) as { caseA: Record<string, CaseA> };
  for (const [name, c] of Object.entries(result.caseA)) {
    const comments = commentsIn(c.before.markup, c.before.comments![0]!.content);
    const open = async (): Promise<EditorDocument> => {
      const session = await EditorDocument.open(await documentOf(c.before.markup, comments));
      await session.stories.comments();
      return session;
    };

    test(`${name}: listed as Office JS lists it, across the plain reference run; the comment on its word`, async () => {
      const session = await open();
      expect(trackedChangesOf(session.doc).map((x) => ({ type: x.type, author: x.author, text: x.text }))).toEqual(c.before.listed!.map((x) => ({ type: x.type, author: x.author, text: x.text })));
      expect(await commentsOf(session)).toEqual(c.before.comments);
    });

    const cases: [string, (s: EditorDocument) => Promise<boolean>, State][] = [
      [`that one change ${c.thatOneWas.action}ed`, async (s) => {
        const [change] = trackedChangesOf(s.doc);
        expect(change!.type).toBe(c.thatOneWas.type);
        return (await runCommandAsync(s, `trackedChange.${c.thatOneWas.action}`, { key: s.changeKey(change!) })).ok;
      }, c.thatOne],
      ['Reject All', async (s) => (await runCommandAsync(s, 'trackedChanges.rejectAll', {})).ok, c.rejectAll],
      ['Accept All', async (s) => (await runCommandAsync(s, 'trackedChanges.acceptAll', {})).ok, c.acceptAll],
    ];
    for (const [what, act, word] of cases) {
      test(`${name}: ${what}, as Word left it`, async () => {
        const session = await open();
        expect(await act(session)).toBe(true);
        expect(trackedChangesOf(session.doc)).toEqual([]);
        const theirs = await EditorDocument.open(await documentOf(word.markup, commentsIn(word.markup, c.before.comments![0]!.content)));
        expect(shape(session.doc)).toEqual(shape(theirs.doc));
        const saved = await reopened(session);
        expect(await commentsOf(saved)).toEqual(word.comments);
      });
    }

    test(`${name}: an undo of the resolution brings the comment back`, async () => {
      const session = await open();
      const [change] = trackedChangesOf(session.doc);
      expect((await runCommandAsync(session, `trackedChange.${c.thatOneWas.action}`, { key: session.changeKey(change!) })).ok).toBe(true);
      session.undoManager.undo();
      expect(await commentsOf(await reopened(session))).toEqual(c.before.comments);
    });
  }
});

// The rule both measurements fit (ED-005 section 12.49): Word removes the comment when its start marker goes with the
// revision. Where the file put the markers outside it, as core-ts 0.3.0 wrote the agent's comment in
// `review-for-word.docx`, Word 2010's and Word 15's Reject All kept the comment as its reference alone
// (`word-review.test.mts` holds the editor to both saves). Here the same form in a small document: the comment stays,
// and its markers go just before its reference, which is the paragraph's first item now (the save had refused this
// before ED-005 section 12.47 item 10).
describe('a comment whose markers the file put outside the revision stays when the revision goes', () => {
  const markup = ['<w:p><w:r><w:t>Kept.</w:t></w:r></w:p>', '<w:p><w:commentRangeStart w:id="0"/><w:ins w:id="1" w:author="Claude" w:date="2026-09-30T13:34:00Z"><w:r><w:t xml:space="preserve">The tenant pays rent monthly</w:t></w:r></w:ins><w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r><w:r><w:t xml:space="preserve"> Kept after.</w:t></w:r></w:p>'];
  for (const how of ['the change rejected', 'Reject All'] as const) {
    test(`${how}: the comment kept, its markers collapsed before its reference`, async () => {
      const session = await EditorDocument.open(await documentOf(markup, { 0: 'On the whole insertion.' }));
      await session.stories.comments();
      const [change] = trackedChangesOf(session.doc);
      const ok = how === 'Reject All' ? (await runCommandAsync(session, 'trackedChanges.rejectAll', {})).ok : (await runCommandAsync(session, 'trackedChange.reject', { key: session.changeKey(change!) })).ok;
      expect(ok).toBe(true);
      const result = await session.save();
      expect(result.selfCheck?.equal, JSON.stringify(result.selfCheck?.difference)).toBe(true);
      const xml = new TextDecoder().decode(new ZipPartStore(result.bytes).loadSync('word/document.xml'));
      const paragraph = xml.slice(xml.indexOf('<w:p', xml.indexOf('Kept.</w:t>')), xml.indexOf('Kept after.'));
      expect([...paragraph.matchAll(/<w:comment(RangeStart|RangeEnd|Reference) w:id="0"\/>/g)].map((m) => m[1])).toEqual(['RangeStart', 'RangeEnd', 'Reference']);
      const again = await EditorDocument.open(result.bytes);
      expect((await commentsOf(again)).map((c) => c.content)).toEqual(['On the whole insertion.']);
    });
  }
});

describe('core-ts check 32, case B: typing into a link that was there, with tracking on, as Word places it', async () => {
  const typed = JSON.parse(await readFile(join(fixturesDir, 'revisions', 'check32', 'typed.json'), 'utf8')) as { typed: State };

  /** "Read the lease today." with an external link on "lease", untracked, saved and opened again: a link that was there. */
  async function withLink(): Promise<{ session: EditorDocument; link: number }> {
    const made = await EditorDocument.open(await documentOf(['<w:p><w:r><w:t xml:space="preserve">Read the lease today.</w:t></w:r></w:p>'], {}));
    expect((await runCommandAsync(made, 'range.hyperlink', { address: 'https://example.com/lease' }, { from: 1 + 9, to: 1 + 14 })).ok).toBe(true);
    const session = await reopened(made);
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author B';
    return { session, link: 1 + 9 };
  }

  /** Each inserted letter, and whether it carries the link. */
  function letters(doc: PMNode): [string, boolean][] {
    const out: [string, boolean][] = [];
    doc.descendants((n) => {
      if (n.isText && n.marks.some((m) => m.type.name === 'insertion')) for (const ch of n.text!) out.push([ch, n.marks.some((m) => m.type.name === 'link')]);
      return true;
    });
    return out;
  }

  test('x in the middle inside the link, y at its end after it, z at its start before it; three changes, as Word lists them', async () => {
    const { session, link } = await withLink();
    // As Jason typed them: "x" after "le", "y" at the link's end, "z" at its start.
    expect(runCommand(session, 'insertText', { text: 'x' }, { from: link + 2 }).ok).toBe(true);
    expect(runCommand(session, 'insertText', { text: 'y' }, { from: link + 6 }).ok).toBe(true);
    expect(runCommand(session, 'insertText', { text: 'z' }, { from: link }).ok).toBe(true);
    expect(session.doc.child(0).textContent).toBe('Read the zlexasey today.');
    expect(letters(session.doc)).toEqual([['z', false], ['x', true], ['y', false]]);
    expect(trackedChangesOf(session.doc).map((x) => ({ type: x.type, author: x.author, text: x.text }))).toEqual(typed.typed.listed!.map((x) => ({ type: x.type, author: x.author, text: x.text })));
    // The file: the middle's w:ins inside the w:hyperlink, the ends' outside it (Word writes the link as a field).
    const saved = await reopened(session);
    expect(letters(saved.doc)).toEqual([['z', false], ['x', true], ['y', false]]);
  });

  test("Word's save of its typing: the three changes listed as Word listed them, the text as Word left it", async () => {
    const session = await EditorDocument.open(await fixture('revisions/check32/check32-typed.docx'));
    expect(session.doc.textContent).toContain('Read the zlexasey today.');
    expect(trackedChangesOf(session.doc).map((x) => ({ type: x.type, author: x.author, text: x.text }))).toEqual(typed.typed.listed!.map((x) => ({ type: x.type, author: x.author, text: x.text })));
  });
});
