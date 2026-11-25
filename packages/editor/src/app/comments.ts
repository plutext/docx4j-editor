// SPDX-License-Identifier: GPL-3.0-or-later
//
// The comments as the margin lists them (ED-005 section 7.5): the model's listing
// (`Stories.commentThreads`) over the body view's document, listed again when the document is
// another one, when a comment's thread fields change in the store (a resolve, a peer's comment
// arriving), and when the stories say something was read or made. The same shape as the review
// panel's store: subscribed through React's `useSyncExternalStore`, the editor and the store
// observed only while something listens (ED-005 section 12.14: StrictMode runs an effect's cleanup
// once at mount).
import { useMemo, useSyncExternalStore } from 'react';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorState } from 'prosemirror-state';
import { referenceIdOf, type EditorDocument, type CommentThread } from '@docx4j/editor-model';
import type { DocxEditorView } from '../view/DocxEditorView.js';

/** What a listing shows of a thread: when two listings say the same, the margin keeps the old one. */
function shapeOf(threads: CommentThread[]): string {
  const one = (t: CommentThread): unknown => [t.id, t.author, t.initials, t.date?.getTime() ?? null, t.done, t.orphan, t.replies.map(one)];
  return JSON.stringify(threads.map(one));
}

export class CommentsStore {
  threads: CommentThread[] = [];
  /** Whether the comments part has been read: until then the margin shows nothing. */
  loaded = false;
  private shape = '';
  private doc: PMNode | undefined;
  private readonly listeners = new Set<() => void>();
  private offs: (() => void)[] = [];

  constructor(private readonly editor: DocxEditorView, private readonly session: EditorDocument) {}

  /** The listing now; the same array while nothing it shows changed, as `useSyncExternalStore` needs. */
  snapshot = (): CommentThread[] => this.threads;

  subscribe = (fn: () => void): (() => void) => {
    this.listeners.add(fn);
    if (this.offs.length === 0) this.start();
    this.refresh(true);
    return () => {
      this.listeners.delete(fn);
      if (this.listeners.size === 0) this.stop();
    };
  };

  private start(): void {
    const store = this.session.store;
    // A peer's comment is adopted by reading again; a resolve changes a field, not the document.
    const onThreads = (): void => { void this.session.stories.comments().then(() => this.refresh(true)); };
    store.commentThreads.observeDeep(onThreads);
    this.offs.push(
      () => store.commentThreads.unobserveDeep(onThreads),
      this.editor.subscribe((state) => { if (state.doc !== this.doc) this.refresh(false); }),
      this.session.stories.subscribe(() => this.refresh(true)),
    );
    void this.session.stories.comments().then(() => { this.loaded = true; this.refresh(true); });
  }

  private stop(): void {
    for (const off of this.offs) off();
    this.offs = [];
  }

  /** List again for the editor's document; tell the listeners when what the margin shows changed. */
  private refresh(always: boolean): void {
    const doc = this.editor.state.doc;
    if (!always && doc === this.doc) return;
    this.doc = doc;
    const next = this.session.stories.commentThreads(doc);
    const shape = shapeOf(next);
    if (shape === this.shape && this.threads.length === next.length) return;
    this.shape = shape;
    this.threads = next;
    for (const fn of this.listeners) fn();
  }

  destroy(): void {
    this.stop();
    this.listeners.clear();
  }
}

const NO_THREADS: CommentThread[] = [];
const noSubscription = (): (() => void) => () => {};
const noThreads = (): CommentThread[] => NO_THREADS;

/**
 * The comments as the margin lists them, for a shell that draws no margin (the compact layout, ED-013 section 3.7): its
 * bar counts the threads at the caret and its sheet draws them. Nothing is listed while `editor` is null.
 */
export function useCommentThreads(editor: DocxEditorView | null, session: EditorDocument | null): CommentThread[] {
  const store = useMemo(() => (editor && session ? new CommentsStore(editor, session) : null), [editor, session]);
  return useSyncExternalStore(store ? store.subscribe : noSubscription, store ? store.snapshot : noThreads, noThreads);
}

/** How many nodes a selection's walk visits before it stops: a selection of the whole document asks no more. */
const SELECTION_WALK = 2000;

/**
 * The threads at the selection, in the listing's order (ED-013 section 3.7: "Comments (n)" on the bar): every top-level
 * thread one of whose comments (it or a reply, which marks the same text) marks the text at a caret, either side of it,
 * or text inside a selection; or whose reference is beside the caret or inside the selection, for a comment with no text.
 */
export function threadsAt(state: EditorState, session: EditorDocument, threads: readonly CommentThread[]): CommentThread[] {
  if (threads.length === 0) return [];
  const ids = new Set<string>();
  const note = (node: PMNode | null | undefined): void => {
    if (!node) return;
    if (node.isText) { for (const m of node.marks) if (m.type.name === 'comment') ids.add(String(m.attrs['id'])); return; }
    const ref = referenceIdOf(node, session.ctx.ids);
    if (ref !== undefined) ids.add(ref);
  };
  const { from, to, empty, $head } = state.selection;
  if (empty) { note($head.nodeBefore); note($head.nodeAfter); }
  else {
    let visited = 0;
    state.doc.nodesBetween(from, to, (node) => { note(node); return ++visited < SELECTION_WALK; });
  }
  if (ids.size === 0) return [];
  return threads.filter((t) => ids.has(t.id) || t.replies.some(function has(r): boolean { return ids.has(r.id) || r.replies.some(has); }));
}
