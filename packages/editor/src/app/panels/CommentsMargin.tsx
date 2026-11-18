// SPDX-License-Identifier: GPL-3.0-or-later
//
// Comments in the margin (ED-002 section 7.3; ED-005 section 7.5): a bubble per thread beside the
// vertical position of its anchor, the first node carrying its `comment` mark (or its reference,
// for an empty range), laid out again after every transaction and on resize and pushed down so
// bubbles never overlap; the author's colour on the anchored text and on the bubble; replies under
// their parent; resolved threads collapsed to a line. Since E4.e: the list follows the model
// (`CommentsStore`), one's own comments are edited in place (an `EditableStoryView` over the
// comment's fragment, the toolbar following it as for a note), others' are shown as their fragment
// has them, and a thread can be replied to, resolved, reopened and deleted.
import { memo, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import { commentFocusOf, referenceIdOf, type EditorDocument, type CommentThread, type Story as StoryModel } from '@docx4j/editor-model';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorSurface } from '../../view/surface.js';
import type { DocxEditorView } from '../../view/DocxEditorView.js';
import { EditableStoryView } from '../../view/StoryView.js';
import { useEditorState } from '../state.js';
import { lockRefusedStatus } from '../labels.js';
import { t as tr, useLanguage, formatDate as formatInLanguage } from '../i18n.js';
import { Story } from '../Story.js';
import { authorColor, authorSwatch } from '../../view/authors.js';
import { scaleOf } from '../../view/zoom.js';
import { CommentsStore } from '../comments.js';

export interface CommentsMarginProps {
  editor: DocxEditorView;
  session: EditorDocument;
  /** The name this page's comments carry: its own comments are the ones it edits (ED-005 section 7.3). */
  me?: string;
  onStatus?: (text: string) => void;
  onActivate?: (surface: EditorSurface) => void;
  onDestroy?: (surface: EditorSurface) => void;
  shortcuts?: Record<string, () => void>;
}

const GAP = 8;

type Range = { from: number; to: number };

/** The comment ranges inside one top-level block, relative to the block's position, kept while the block's node object lives (ED-003 section 8.2). */
const blockRanges = new WeakMap<PMNode, Map<string, Range>>();

function rangesWithin(block: PMNode): Map<string, Range> {
  let out = blockRanges.get(block);
  if (out) return out;
  out = new Map();
  const note = (node: PMNode, at: number): void => {
    for (const m of node.marks) {
      if (m.type.name !== 'comment') continue;
      const id = String(m.attrs['id']);
      const r = out!.get(id);
      if (r) r.to = at + node.nodeSize; else out!.set(id, { from: at, to: at + node.nodeSize });
    }
  };
  note(block, 0);
  block.descendants((node, rel) => { note(node, 1 + rel); return true; });
  blockRanges.set(block, out);
  return out;
}

/** Every comment's range in a document, the first to the last node carrying its mark, in one pass over the blocks (memoised per block, then per document). */
const docRanges = new WeakMap<PMNode, Map<string, Range>>();
export function commentRanges(doc: PMNode): Map<string, Range> {
  let out = docRanges.get(doc);
  if (out) return out;
  out = new Map();
  doc.forEach((block, pos) => {
    for (const [id, r] of rangesWithin(block)) {
      const have = out!.get(id);
      if (have) have.to = pos + r.to; else out!.set(id, { from: pos + r.from, to: pos + r.to });
    }
  });
  docRanges.set(doc, out);
  return out;
}

/** The range a comment id marks in the view's document. */
export function commentRange(editor: EditorSurface, id: string): Range | undefined {
  return commentRanges(editor.state.doc).get(id);
}

/** A comment's date in the application's language, cached per date and language. */
const dates = new WeakMap<Date, { lang: string; text: string }>();
function formatDate(d: Date | undefined, lang: string): string {
  if (!d) return '';
  let s = dates.get(d);
  if (s === undefined || s.lang !== lang) {
    s = { lang, text: formatInLanguage(d) };
    dates.set(d, s);
  }
  return s.text;
}

/** Where each comment with an empty range has its reference, the anchor it has instead of text (ED-005 section 7.3). */
const docReferences = new WeakMap<PMNode, Map<string, number>>();
function referencePositions(doc: PMNode, session: EditorDocument): Map<string, number> {
  let out = docReferences.get(doc);
  if (out) return out;
  out = new Map();
  doc.descendants((node, pos) => {
    const id = referenceIdOf(node, session.ctx.ids);
    if (id !== undefined && !out!.has(id)) out!.set(id, pos);
    return true;
  });
  docReferences.set(doc, out);
  return out;
}

export function CommentsMargin({ editor, session, me, onStatus, onActivate, onDestroy, shortcuts }: CommentsMarginProps) {
  const lang = useLanguage();
  const store = useMemo(() => new CommentsStore(editor, session), [editor, session]);
  const threads = useSyncExternalStore(store.subscribe, store.snapshot);
  const state = useEditorState(editor);
  const root = useRef<HTMLElement>(null);
  const bubbles = useRef(new Map<string, HTMLElement>());
  const bubbleRefs = useRef(new Map<string, (el: HTMLElement | null) => void>());
  const [tops, setTops] = useState<Map<string, number>>(new Map());
  const [tick, setTick] = useState(0);
  /** The comment a command asked to take the caret (a new comment or reply), until its body mounts. */
  const pendingFocus = useRef<string | null>(null);
  useEffect(() => editor.subscribe((_state, trs) => { for (const tr of trs) { const key = commentFocusOf(tr); if (key) pendingFocus.current = key; } }), [editor]);

  // Lay out again when the page's size changes (fonts, images, the window), and when a bubble's
  // does (a reply box opened, a comment typed into), so no bubble covers the one before it.
  const resized = useRef<ResizeObserver | null>(null);
  useEffect(() => {
    if (typeof ResizeObserver === 'undefined') return;
    const ro = new ResizeObserver(() => setTick((t) => t + 1));
    ro.observe(editor.container);
    for (const el of bubbles.current.values()) ro.observe(el);
    resized.current = ro;
    return () => { ro.disconnect(); resized.current = null; };
  }, [editor]);

  // The anchors follow the document, not the selection: a caret move lays nothing out again.
  const doc = state?.doc;
  const anchors = useMemo(() => {
    const out = new Map<string, number>();
    if (!doc) return out;
    const ranges = commentRanges(doc);
    for (const t of threads) {
      const r = ranges.get(t.id);
      const at = r?.from ?? referencePositions(doc, session).get(t.id);
      if (at !== undefined) out.set(t.id, at);
    }
    return out;
  }, [threads, doc, session]);

  // Anchor tops from the view, then each bubble no higher than the one before it ends.
  useLayoutEffect(() => {
    if (!root.current) return;
    const marginTop = root.current.getBoundingClientRect().top;
    // The margin is zoomed with the page: a bubble's `top` is layout pixels, the caret's coordinates client ones.
    const scale = scaleOf(editor.view.dom);
    const wanted: { id: string; top: number }[] = [];
    for (const t of threads) {
      const pos = anchors.get(t.id);
      let top = 0;
      if (pos !== undefined) {
        try { top = (editor.view.coordsAtPos(pos).top - marginTop) / scale; } catch { top = 0; }
      }
      wanted.push({ id: t.id, top: Math.max(0, top) });
    }
    wanted.sort((a, b) => a.top - b.top || Number(a.id) - Number(b.id));
    const next = new Map<string, number>();
    let floor = 0;
    for (const w of wanted) {
      const top = Math.max(w.top, floor);
      next.set(w.id, top);
      floor = top + (bubbles.current.get(w.id)?.offsetHeight ?? 0) + GAP;
    }
    let same = next.size === tops.size;
    if (same) for (const [k, v] of next) if (Math.abs((tops.get(k) ?? -1) - v) > 0.5) { same = false; break; }
    if (!same) setTops(next);
  }, [threads, anchors, editor, tick, tops]);

  /**
   * Each bubble's ref, one function per comment. React calls a new ref function on every render, with null and then
   * the element, and a fresh `observe` always reports: an inline ref re-observing here redrew the margin on every
   * frame, for ever, and starved any long render beside it (ED-005 section 12.113).
   */
  const bubbleRef = (id: string): ((el: HTMLElement | null) => void) => {
    let ref = bubbleRefs.current.get(id);
    if (!ref) {
      ref = (el) => {
        const was = bubbles.current.get(id);
        if (was && was !== el) resized.current?.unobserve(was);
        if (el) { bubbles.current.set(id, el); resized.current?.observe(el); } else bubbles.current.delete(id);
      };
      bubbleRefs.current.set(id, ref);
    }
    return ref;
  };

  const select = useCallback((id: string): void => {
    const r = commentRange(editor, id);
    if (r) { editor.reveal(r.from, r.to); editor.caretAt(r.from, r.to); }
  }, [editor]);

  const actions = useMemo<ThreadActions>(() => ({
    reply: (id, text) => { void editor.runAsync('comment.reply', { id, text }).then((ok) => { if (!ok) onStatus?.(tr('comments.notDone')); }); },
    resolve: (id, value) => { void editor.runAsync('comment.setResolved', { id, value }); },
    remove: (id) => { void editor.runAsync('comment.delete', { id }).then((ok) => { if (ok) onStatus?.(tr('comments.deleted')); }); },
    claimFocus: (id) => { if (pendingFocus.current !== `comment:${id}`) return false; pendingFocus.current = null; return true; },
    onStatus: (text) => onStatus?.(text),
    ...(onActivate ? { onActivate } : {}), ...(onDestroy ? { onDestroy } : {}), ...(shortcuts ? { shortcuts } : {}),
  }), [editor, onStatus, onActivate, onDestroy, shortcuts]);

  if (threads.length === 0) return null;

  // The anchored text takes its author's colour through a variable the mark's span reads.
  const css = threads.flatMap((t) => [t, ...t.replies]).map((t) => `.docx-editor .comment[data-comment="${t.id}"] { --comment-color: ${authorColor(t.author)}; }`).join('\n');
  const height = Math.max(0, ...[...tops.entries()].map(([id, top]) => top + (bubbles.current.get(id)?.offsetHeight ?? 0)));

  return (
    <aside ref={root} className="comments-margin" aria-label={tr('comments.margin')} data-testid="comments-margin" style={{ minBlockSize: height }}
      onBlur={(e) => { if (!e.currentTarget.contains(e.relatedTarget as Node | null)) onStatus?.(''); }}>
      <style>{css}</style>
      {threads.map((t) => (
        <div key={t.id} ref={bubbleRef(t.id)}
          className={`comment-bubble${t.done ? ' done' : ''}${t.orphan ? ' orphan' : ''}`} data-testid="comment-bubble" data-comment-id={t.id}
          role="group" aria-label={tr('comments.by', { author: t.author })}
          style={{ top: tops.get(t.id) ?? 0, '--comment-color': authorColor(t.author), '--comment-pattern': authorSwatch(t.author).backgroundImage } as React.CSSProperties}>
          {t.done
            ? <details><summary className="comment-head"><b>{t.author}</b> <span className="meta">{tr('comments.resolved')}</span></summary><Thread thread={t} session={session} editor={editor} me={me} onSelect={select} lang={lang} actions={actions} /></details>
            : <Thread thread={t} session={session} editor={editor} me={me} onSelect={select} lang={lang} actions={actions} />}
        </div>
      ))}
    </aside>
  );
}

/** What a thread's buttons do, and what its editable bodies need: stable for the margin's life, so the memoised threads are not drawn again for them. */
interface ThreadActions {
  reply: (id: string, text: string) => void;
  resolve: (id: string, value: boolean) => void;
  remove: (id: string) => void;
  /** Whether this comment's body should take the caret as it mounts (the command that made it asked). */
  claimFocus: (id: string) => boolean;
  onStatus: (text: string) => void;
  onActivate?: (surface: EditorSurface) => void;
  onDestroy?: (surface: EditorSurface) => void;
  shortcuts?: Record<string, () => void>;
}

interface ThreadProps {
  thread: CommentThread;
  session: EditorDocument;
  editor: DocxEditorView;
  me: string | undefined;
  onSelect: (id: string) => void;
  /** A prop so the memo draws again when the language changes. */
  lang: string;
  actions: ThreadActions;
  reply?: boolean;
}

/** One comment and its replies, with the thread's buttons: Reply, Resolve or Reopen and Delete for its first comment, Delete for a reply. */
const Thread = memo(function Thread({ thread, session, editor, me, onSelect, lang, actions, reply = false }: ThreadProps) {
  const [replying, setReplying] = useState(false);
  const own = me !== undefined && thread.author === me;
  const label = tr('comments.by', { author: thread.author });
  return (
    <div className={reply ? 'comment-reply' : 'comment-thread'}>
      <button type="button" className="comment-head" title={thread.orphan ? tr('comments.orphan.title') : tr('comments.select.title')} disabled={thread.orphan} onClick={() => onSelect(thread.id)}>
        <span className="initials" aria-hidden="true">{thread.initials || thread.author.slice(0, 2)}</span>
        <b className="author">{thread.author}</b>
        <span className="meta">{formatDate(thread.date, lang)}</span>
      </button>
      {own
        ? <CommentEditor editor={editor} session={session} story={thread.record} label={label} actions={actions} />
        : <CommentText session={session} story={thread.record} label={label} onFocus={() => actions.onStatus(tr('comments.notYours', { author: thread.author }))} />}
      <div className="comment-actions">
        {!reply && !thread.done && <button type="button" data-testid="comment-reply" onClick={() => setReplying((r) => !r)}>{tr('comments.reply')}</button>}
        {!reply && <button type="button" data-testid={thread.done ? 'comment-reopen' : 'comment-resolve'} onClick={() => actions.resolve(thread.id, !thread.done)}>{thread.done ? tr('comments.reopen') : tr('comments.resolve')}</button>}
        <button type="button" data-testid="comment-delete" title={tr('comments.delete.title')} onClick={() => actions.remove(thread.id)}>{tr('comments.delete')}</button>
      </div>
      {thread.replies.map((r) => <Thread key={r.id} thread={r} session={session} editor={editor} me={me} onSelect={onSelect} lang={lang} actions={actions} reply />)}
      {replying && <ReplyBox author={thread.author} onPost={(text) => { setReplying(false); actions.reply(thread.id, text); }} onCancel={() => setReplying(false)} />}
    </div>
  );
});

/** The reply field: a line of text, posted as a reply whose text is then edited like any comment's. */
function ReplyBox({ author, onPost, onCancel }: { author: string; onPost: (text: string) => void; onCancel: () => void }) {
  const [text, setText] = useState('');
  const post = (): void => { if (text.trim() !== '') onPost(text.trim()); };
  return (
    <div className="comment-reply-box">
      <textarea data-testid="comment-reply-text" aria-label={tr('comments.reply.label', { author })} placeholder={tr('comments.reply.placeholder')} value={text} autoFocus rows={2}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); post(); } else if (e.key === 'Escape') { e.preventDefault(); onCancel(); } }} />
      <button type="button" data-testid="comment-post" disabled={text.trim() === ''} onClick={post}>{tr('comments.post')}</button>
      <button type="button" onClick={onCancel}>{tr('comments.cancel')}</button>
    </div>
  );
}

/**
 * One's own comment, editable in place over its fragment, as a note is; it takes the caret when the command that made
 * it asked. It takes it in a microtask after the mount, not in the effect (ED-005 section 12.103): the effect runs
 * before the margin's layout pass has placed a new bubble, which is drawn at the margin's top until then, so a caret
 * taken there scrolled the page to the first comment; and under StrictMode (development) the mount is replayed, and
 * the first view, destroyed by then, had spent the claim, leaving the caret nowhere. The microtask runs after both:
 * the layout pass's re-render is synchronous with the commit, and a destroyed view is not the current one.
 */
function CommentEditor({ editor, session, story, label, actions }: { editor: DocxEditorView; session: EditorDocument; story: StoryModel; label: string; actions: ThreadActions }) {
  const ref = useRef<HTMLDivElement>(null);
  const current = useRef<EditableStoryView | null>(null);
  useEffect(() => {
    if (!ref.current) return;
    const view = new EditableStoryView(ref.current, session, story, {
      label, stylesheet: editor,
      ...(actions.onActivate ? { onActivate: actions.onActivate } : {}), ...(actions.onDestroy ? { onDestroy: actions.onDestroy } : {}), ...(actions.shortcuts ? { shortcuts: actions.shortcuts } : {}),
      onRefused: (reason) => actions.onStatus(lockRefusedStatus(reason)),
    });
    current.current = view;
    queueMicrotask(() => {
      if (current.current !== view || !actions.claimFocus(story.key.slice('comment:'.length))) return;
      view.caretAt(view.state.doc.content.size - 1);
    });
    return () => { if (current.current === view) current.current = null; view.destroy(); };
  }, [editor, session, story, label, actions]);
  return <div ref={ref} className="comment-body" data-testid="comment-body" data-editable="true" />;
}

/** Someone else's comment, read-only, as its fragment has it now: a peer's edit shows as it arrives. */
function CommentText({ session, story, label, onFocus }: { session: EditorDocument; story: StoryModel; label: string; onFocus: () => void }) {
  const [doc, setDoc] = useState(() => session.stories.currentDoc(story));
  useEffect(() => {
    const key = story.key.slice('comment:'.length);
    const store = session.store.comments;
    let fragment = store.get(key);
    const update = (): void => setDoc(session.stories.currentDoc(story));
    fragment?.observeDeep(update);
    // The fragment can arrive later (a peer's first binding); the map says when.
    const onMap = (): void => { const now = store.get(key); if (now !== fragment) { fragment?.unobserveDeep(update); fragment = now; fragment?.observeDeep(update); update(); } };
    store.observe(onMap);
    update();
    return () => { fragment?.unobserveDeep(update); store.unobserve(onMap); };
  }, [session, story]);
  const shown = useMemo(() => ({ ...story, doc }), [story, doc]);
  return <div onFocus={onFocus}><Story session={session} story={shown} className="comment-body" label={label} testId="comment-body" /></div>;
}

