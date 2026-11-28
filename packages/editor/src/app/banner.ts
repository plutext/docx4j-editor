// SPDX-License-Identifier: GPL-3.0-or-later
//
// What the banner counts as the document now stands (ED-005 section 6.7; Jason, 2026-11-28: it had
// counted the file as opened, and stood unchanged through every change accepted and comment
// resolved). The tracked changes, and the revisions shown in Word only, as the review panel lists
// them (`reviewEntries`, with this view's split replacements); the comments not yet resolved, as the
// margin lists them (`Stories.commentThreads`). Counted again a moment after the document or a
// comment's thread fields last changed, never on the keystroke: the listings walk every top-level
// block, about 2 ms a keystroke on the 255-page document (section 12.134). Subscribed through
// React's `useSyncExternalStore`, the editor observed only while something listens (section 12.14).
import { useMemo, useSyncExternalStore } from 'react';
import type { Node as PMNode } from 'prosemirror-model';
import type { CommentThread } from '@docx4j/editor-model';
import type { DocxEditorView } from '../view/DocxEditorView.js';
import { reviewEntries, splitsOf } from './review.js';

/** The banner's counts that follow the document. */
export interface BannerCounts {
  /** The tracked changes, as the review panel lists them (a move, a replacement, one entry). */
  changes: number;
  /** The revisions shown in Word only, as the panel lists them (a table's records one entry). */
  preserved: number;
  /** The comments not resolved, replies included; undefined until the comments part is read. */
  comments: number | undefined;
}

/** How long after the last change the banner counts again. */
export const RECOUNT_MS = 250;

/** The comments of a listing not resolved: every comment of a thread whose top comment is open, its replies with it. */
export function unresolvedComments(threads: readonly CommentThread[]): number {
  const whole = (t: CommentThread): number => 1 + t.replies.reduce((n, r) => n + whole(r), 0);
  return threads.reduce((n, t) => n + (t.done ? 0 : whole(t)), 0);
}

const same = (a: BannerCounts, b: BannerCounts): boolean => a.changes === b.changes && a.preserved === b.preserved && a.comments === b.comments;

/** The counts over one editor, the same object while they stand, as `useSyncExternalStore` needs. */
export class BannerStore {
  private counts: BannerCounts;
  /** Whether the comments part has been read: until then nothing is said of the comments. */
  private loaded = false;
  /** The document last counted. */
  private doc: PMNode | undefined;
  private timer: ReturnType<typeof setTimeout> | undefined;
  private readonly listeners = new Set<() => void>();
  private offs: (() => void)[] = [];

  constructor(private readonly editor: DocxEditorView, private readonly delay = RECOUNT_MS) {
    this.counts = this.read();
  }

  private read(): BannerCounts {
    const { session, state } = this.editor;
    this.doc = state.doc;
    const entries = reviewEntries(state, session, splitsOf(this.editor));
    const changes = entries.filter((e) => e.kind === 'change').length;
    return { changes, preserved: entries.length - changes, comments: this.loaded ? unresolvedComments(session.stories.commentThreads(state.doc)) : undefined };
  }

  /** Count now, and tell the listeners if anything changed. */
  private recount = (): void => {
    this.timer = undefined;
    const next = this.read();
    if (same(next, this.counts)) return;
    this.counts = next;
    for (const fn of this.listeners) fn();
  };

  /** Count again once the changes pause. */
  private later = (): void => {
    if (this.timer !== undefined) clearTimeout(this.timer);
    this.timer = setTimeout(this.recount, this.delay);
  };

  snapshot = (): BannerCounts => this.counts;

  subscribe = (fn: () => void): (() => void) => {
    this.listeners.add(fn);
    if (this.offs.length === 0) this.start();
    // Whatever changed while nobody listened.
    if (this.editor.state.doc !== this.doc) this.recount();
    return () => {
      this.listeners.delete(fn);
      if (this.listeners.size === 0) this.stop();
    };
  };

  private start(): void {
    const { session } = this.editor;
    const stories = session.stories;
    const threads = session.store.commentThreads;
    // A resolve changes a thread's field, not the document; a peer's comment is adopted by reading the comments again.
    const onThreads = (): void => { void stories.comments().then(this.later); };
    threads.observeDeep(onThreads);
    this.offs.push(
      () => threads.unobserveDeep(onThreads),
      this.editor.subscribe((state) => { if (state.doc !== this.doc) this.later(); }),
      stories.subscribe(this.later),
      () => { if (this.timer !== undefined) clearTimeout(this.timer); this.timer = undefined; },
    );
    void stories.comments().then(() => { if (this.loaded) return; this.loaded = true; this.recount(); });
  }

  private stop(): void {
    for (const off of this.offs) off();
    this.offs = [];
  }
}

const noSubscription = (): (() => void) => () => {};
const noCounts = (): null => null;

/** The banner's counts over `editor`'s document; null while `editor` is null, when nothing is listed. */
export function useBannerCounts(editor: DocxEditorView | null): BannerCounts | null {
  const store = useMemo(() => (editor ? new BannerStore(editor) : null), [editor]);
  return useSyncExternalStore(store ? store.subscribe : noSubscription, store ? store.snapshot : noCounts, noCounts);
}
