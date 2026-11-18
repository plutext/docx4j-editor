// SPDX-License-Identifier: GPL-3.0-or-later
//
// Comments made, replied to, resolved and deleted (ED-005 section 7.3), in Office JS's shapes:
// `range.insertComment(text)`, `comment.reply(text)`, `comment.resolved`, `comment.delete()`. A
// comment's anchor is its `comment` marks on the text it covers and its reference atom, the
// `w:commentReference` run built here and registered as a note's reference is (`newReference`),
// so export's identity path serves it. `prepare` makes the comment (its fragment, its thread
// entry, its element in `sources`) and the reference run's source; `run` anchors it. The thread
// fields a command changes travel on its transaction (`addCommentWrite`), so each command is one
// undo step. Commenting is never tracked, as in Word: the transactions carry `UNTRACKED_META`,
// which also keeps a deleted comment's reference atom from being kept as a deletion.
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { deepCopy, runOf, type Element } from '@docx4j/core-ts';
import * as el from '@docx4j/generated-objects-ts/modules/org_docx4j_wml.el';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { register, sessionOf, storyOf } from './spec.mjs';
import { addCommentWrite } from '../yjs.mjs';
import { UNTRACKED_META, trackingDate } from '../track.mjs';
import { COMMENT_STYLES, referenceIdOf } from '../comments.mjs';
import type { EditorDocument } from '../session.mjs';
import type { CommentThread } from '../stories.mjs';

/** The transaction meta naming the comment whose view should take the caret once it mounts (`comment:7`). */
export const COMMENT_FOCUS = 'docx4j-comment-focus';

export function commentFocusOf(tr: Transaction): string | undefined {
  return tr.getMeta(COMMENT_FOCUS) as string | undefined;
}

/** Who a comment made now is by: the name tracked changes carry, else the package's author (ED-005 section 7.3). */
export function commentAuthor(session: EditorDocument): string {
  return session.trackingAuthor ?? session.pkg.author.name;
}

/** When a comment made now was made: to the minute, as Word writes it, the session's fixed date in tests. */
function commentDate(session: EditorDocument): Date {
  return new Date(trackingDate(session.trackedChangeDate));
}

const inBody = (state: EditorState): boolean => storyOf(state) === undefined && sessionOf(state) !== undefined;

/**
 * What New Comment covers (ED-005 section 7.3, proposal 32, as Jason measured Word): a selection
 * as it is; at a caret, the word it is in, else the word after it when the caret touches that
 * word, else the word before it, however many spaces lie between; with no word before, an empty
 * range at the paragraph's start. Words are `Intl.Segmenter`'s, in logical order; an atom counts
 * as one character that is no word. Undefined where no comment can go (not in a paragraph).
 */
export function commentTarget(state: EditorState): { from: number; to: number } | undefined {
  const { $from, $to, empty } = state.selection;
  if (!empty) return $from.parent.isTextblock || $to.parent.isTextblock ? { from: $from.pos, to: $to.pos } : undefined;
  const para = $from.parent;
  if (!para.isTextblock) return undefined;
  const start = $from.start();
  // Each UTF-16 unit of the paragraph's text, and each inline leaf, with its position.
  const positions: number[] = [];
  let text = '';
  para.descendants((node, pos) => {
    if (node.isText) { for (let i = 0; i < node.text!.length; i++) positions.push(start + pos + i); text += node.text; }
    else if (node.isLeaf) { positions.push(start + pos); text += '￼'; }
    return true;
  });
  const caret = positions.filter((p) => p < $from.pos).length;
  const words = [...new Intl.Segmenter(undefined, { granularity: 'word' }).segment(text)].filter((s) => s.isWordLike);
  const word = words.find((s) => s.index <= caret && caret < s.index + s.segment.length)
    ?? [...words].reverse().find((s) => s.index + s.segment.length <= caret);
  if (!word) return { from: start, to: start };
  return { from: positions[word.index]!, to: positions[word.index + word.segment.length - 1]! + 1 };
}

/** The run properties at a position (the text or atom before it, else after), for a reference run to copy. */
function runPropertiesAt(state: EditorState, session: EditorDocument, pos: number): wml.RPr | undefined {
  const $pos = state.doc.resolve(pos);
  for (const n of [$pos.nodeBefore, $pos.nodeAfter]) {
    if (!n) continue;
    const id = n.isText ? n.marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'] : n.attrs['runSource'];
    if (typeof id !== 'string') continue;
    const run = session.ctx.ids.get(id) as Element<wml.R> | undefined;
    if (run?.value.rPr) return run.value.rPr;
  }
  return undefined;
}

/**
 * A comment's reference run, built and registered: a `w:commentReference` with the id, in a
 * `w:r` whose properties are a copy of `base` with `w:rStyle` `CommentReference`, as Word writes
 * it. The atom carries both ids; export emits them by identity, as a note's reference.
 */
export function newCommentReference(session: EditorDocument, id: number, base: wml.RPr | undefined): { sourceId: string; runSource: string; run: Element } {
  const rPr: wml.RPr = base ? deepCopy(base) : { TYPE_NAME: 'org_docx4j_wml.RPr' };
  rPr.rStyle = { val: COMMENT_STYLES.reference };
  // The run's own formatting change stays with the run it came from.
  delete rPr.rPrChange;
  const reference = el.commentReference({ id }) as Element;
  const run = runOf([reference], rPr);
  const runSource = session.ctx.ids.idOf(run);
  session.ctx.sources.set(runSource, { holders: [] });
  return { sourceId: session.ctx.ids.idOf(reference), runSource, run };
}

/** The `comment` mark for `id` added to every text node in `[from, to)`. */
function markRange(tr: Transaction, from: number, to: number, id: string): void {
  const mark = tr.doc.type.schema.marks['comment']!.create({ id });
  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (node.isText) tr.addMark(Math.max(pos, from), Math.min(pos + node.nodeSize, to), mark);
    return true;
  });
}

/** The ranges of text carrying a comment's mark, in document order. */
function markedRanges(doc: PMNode, id: string): { from: number; to: number }[] {
  const out: { from: number; to: number }[] = [];
  doc.descendants((node, pos) => {
    if (node.isText && node.marks.some((m) => m.type.name === 'comment' && String(m.attrs['id']) === id)) {
      const last = out[out.length - 1];
      if (last && last.to === pos) last.to = pos + node.nodeSize; else out.push({ from: pos, to: pos + node.nodeSize });
    }
    return true;
  });
  return out;
}

/** Where a comment's reference atom is, when the body has one. */
function referencePosition(session: EditorDocument, doc: PMNode, id: string): number | undefined {
  let found: number | undefined;
  doc.descendants((node, pos) => {
    if (found !== undefined) return false;
    if (referenceIdOf(node, session.ctx.ids) === id) found = pos;
    return !node.isLeaf;
  });
  return found;
}

/** The listed thread a comment belongs to (its first comment), and the comment itself. */
function threadOf(session: EditorDocument, doc: PMNode, id: string): { top: CommentThread; comment: CommentThread } | undefined {
  for (const top of session.stories.commentThreads(doc)) {
    if (top.id === id) return { top, comment: top };
    const reply = top.replies.find((r) => r.id === id);
    if (reply) return { top, comment: reply };
  }
  return undefined;
}

/** The comments are read before any command over them runs: `prepare`'s one job for the three that change a comment. */
async function loaded<A>(state: EditorState, args: A): Promise<A> {
  await sessionOf(state)?.stories.comments();
  return args;
}

export interface InsertCommentArgs {
  /** The comment's text; nothing for an empty comment with the caret in it. */
  text?: string;
  /** Filled by `prepare`: the comment made. */
  id?: string;
  /** Filled by `prepare`: the reference's element ids, its run already in `sources` (ED-004 section 3.7). */
  reference?: { sourceId: string; runSource: string };
}

register<InsertCommentArgs | undefined>({
  name: 'range.insertComment',
  label: 'New comment',
  args: { text: 'the comment text, or nothing for an empty comment to type into' },
  // Office JS's `insertComment` returns the comment; the engine's, a promise of it: awaited, as both allow.
  api: (args) => `await selection.range.insertComment(${JSON.stringify(args?.text ?? '')});`,
  prepare: async (state, args) => {
    const session = sessionOf(state);
    const target = inBody(state) ? commentTarget(state) : undefined;
    if (!session || !target) return args;
    await session.ensureCommentStyles();
    const record = await session.stories.newComment({ author: commentAuthor(session), date: commentDate(session), text: args?.text ?? '' });
    // The reference's run, which is not in the base file, into `sources` before the transaction that inserts it (ED-004 section 3.7).
    const { run, ...reference } = newCommentReference(session, Number(record.id), runPropertiesAt(state, session, target.to));
    await session.writeInlineSource(run);
    return { ...(args ?? {}), id: record.id, reference };
  },
  run: (args) => (state, dispatch) => {
    const session = sessionOf(state);
    const target = inBody(state) ? commentTarget(state) : undefined;
    if (!session || !target || args?.id === undefined || !args.reference) return false;   // the part half did not run
    const tr = state.tr;
    markRange(tr, target.from, target.to, args.id);
    tr.insert(target.to, state.schema.nodes['opaqueInline']!.create({ ...args.reference, kind: 'commentReference' }));
    tr.setSelection(TextSelection.create(tr.doc, target.to));
    tr.setMeta(COMMENT_FOCUS, `comment:${args.id}`).setMeta(UNTRACKED_META, true);
    dispatch?.(tr.scrollIntoView());
    return true;
  },
  isEnabled: (state) => inBody(state) && commentTarget(state) !== undefined,
});

export interface ReplyArgs {
  /** The comment replied to; a reply to a reply joins its thread, as in Word 15. */
  id: string;
  text?: string;
  /** Filled by `prepare`: the reply made. */
  replyId?: string;
  reference?: { sourceId: string; runSource: string };
}

register<ReplyArgs>({
  name: 'comment.reply',
  label: 'Reply to comment',
  args: { id: "the comment's id", text: 'the reply text' },
  api: (args) => `await (await body.getComments()).flatMap((c) => [c, ...c.replies]).find((c) => c.id === ${Number(args.id)}).reply(${JSON.stringify(args.text ?? '')});`,
  prepare: async (state, args) => {
    const session = sessionOf(state);
    if (!session || !inBody(state)) return args;
    await session.stories.comments();
    const thread = threadOf(session, state.doc, args.id);
    if (!thread) return args;
    const at = referencePosition(session, state.doc, thread.top.id);
    await session.ensureCommentStyles();
    const record = await session.stories.newComment({ author: commentAuthor(session), date: commentDate(session), parent: thread.top.id, text: args.text ?? '' });
    const { run, ...reference } = newCommentReference(session, Number(record.id), at === undefined ? undefined : (session.ctx.ids.get(state.doc.nodeAt(at)!.attrs['runSource'] as string) as Element<wml.R> | undefined)?.value.rPr);
    await session.writeInlineSource(run);
    return { ...args, replyId: record.id, reference };
  },
  run: (args) => (state, dispatch) => {
    const session = sessionOf(state);
    if (!session || !inBody(state) || args.replyId === undefined || !args.reference) return false;
    const thread = threadOf(session, state.doc, args.id);
    if (!thread) return false;
    const tr = state.tr;
    // The reply covers its thread's text and its reference follows the thread's, as Word and the engine nest them.
    for (const r of markedRanges(state.doc, thread.top.id)) markRange(tr, r.from, r.to, args.replyId);
    const at = referencePosition(session, state.doc, thread.top.id);
    const ranges = markedRanges(state.doc, thread.top.id);
    const place = at !== undefined ? at + 1 : ranges.length > 0 ? ranges[ranges.length - 1]!.to : undefined;
    if (place === undefined) return false;
    tr.insert(place, state.schema.nodes['opaqueInline']!.create({ ...args.reference, kind: 'commentReference' }));
    tr.setMeta(COMMENT_FOCUS, `comment:${args.replyId}`).setMeta(UNTRACKED_META, true);
    dispatch?.(tr);
    return true;
  },
  isEnabled: (state, args) => inBody(state) && args !== undefined,
});

export interface CommentIdArgs {
  id: string;
}

register<CommentIdArgs & { value?: boolean | 'toggle' }>({
  name: 'comment.setResolved',
  label: 'Resolve comment',
  args: { id: "the comment's id; a reply resolves its thread", value: "true, false or 'toggle' (the default)" },
  choices: { value: ['true', 'false', 'toggle'] },
  api: (args) => typeof args.value === 'boolean' ? `(await body.getComments()).flatMap((c) => [c, ...c.replies]).find((c) => c.id === ${Number(args.id)}).resolved = ${args.value};` : undefined,
  prepare: loaded,
  run: (args) => (state, dispatch) => {
    const session = sessionOf(state);
    const thread = session && inBody(state) ? threadOf(session, state.doc, args.id) : undefined;
    if (!thread) return false;
    const value = args.value === undefined || args.value === 'toggle' ? !thread.top.done : args.value;
    const tr = state.tr;
    addCommentWrite(tr, { id: thread.top.id, fields: { done: value } });
    dispatch?.(tr.setMeta(UNTRACKED_META, true));
    return true;
  },
  isEnabled: (state, args) => {
    const session = sessionOf(state);
    return args !== undefined && session !== undefined && inBody(state) && threadOf(session, state.doc, args.id) !== undefined;
  },
});

register<CommentIdArgs>({
  name: 'comment.delete',
  label: 'Delete comment',
  args: { id: "the comment's id; a thread's first comment deletes the thread" },
  api: (args) => `await (await body.getComments()).flatMap((c) => [c, ...c.replies]).find((c) => c.id === ${Number(args.id)}).delete();`,
  prepare: loaded,
  run: (args) => (state, dispatch) => {
    const session = sessionOf(state);
    const thread = session && inBody(state) ? threadOf(session, state.doc, args.id) : undefined;
    if (!session || !thread) return false;
    const ids = new Set(thread.comment === thread.top ? [thread.top.id, ...thread.top.replies.map((r) => r.id)] : [thread.comment.id]);
    const tr = state.tr;
    const atoms: number[] = [];
    state.doc.descendants((node, pos) => {
      if (node.isText) {
        for (const m of node.marks) if (m.type.name === 'comment' && ids.has(String(m.attrs['id']))) tr.removeMark(pos, pos + node.nodeSize, m);
      } else {
        const ref = referenceIdOf(node, session.ctx.ids);
        if (ref !== undefined && ids.has(ref)) atoms.push(pos);
      }
      return true;
    });
    for (const pos of atoms.reverse()) tr.delete(pos, pos + 1);
    for (const id of ids) addCommentWrite(tr, { id, fields: { deleted: true } });
    dispatch?.(tr.setMeta(UNTRACKED_META, true));
    return true;
  },
  isEnabled: (state, args) => {
    const session = sessionOf(state);
    return args !== undefined && session !== undefined && inBody(state) && threadOf(session, state.doc, args.id) !== undefined;
  },
});
