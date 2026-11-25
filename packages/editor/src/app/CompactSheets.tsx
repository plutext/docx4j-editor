// SPDX-License-Identifier: GPL-3.0-or-later
//
// The compact layout's sheets for review and comments (ED-013 sections 3.6 and 3.7, P5 and P6), over the same pieces the
// full layout draws beside the page:
//
// - **Review:** the review panel (`ReviewPanel` over `ReviewStore`) in a sheet. An entry, Previous or Next puts the
//   caret in the change and closes the sheet, so the bar above the keyboard offers Accept, Reject and Next there;
//   Accept all and Reject all ask first. Every resolution still goes through the four commands (ED-005 section 12.15).
// - **A thread:** the threads at the caret, drawn by the margin's `Thread` without its positioning: each comment with
//   its author and date, Reply, Resolve or Reopen, Delete on one's own, and one field the replies are written in.
// - **A new comment:** the selection's text quoted and the same field. The comment is made with its text when Comment is
//   tapped, so cancelling leaves no empty comment behind.
//
// With no name set, the field asks it (the guest's name to start from), as the full layout's name dialog does: a dialog
// opened as a sheet closes lost the phone's keyboard (ED-013 section 9).
import { Suspense, lazy, useEffect, useMemo, useRef, useState, type CSSProperties } from 'react';
import type { CommentThread, EditorDocument } from '@docx4j/editor-model';
import type { DocxEditorView } from '../view/DocxEditorView.js';
import { authorColor } from '../view/authors.js';
import { Sheet } from './Sheet.js';
import { Thread, commentRange, type ThreadActions } from './panels/CommentsMargin.js';
import { useLanguage, useT } from './i18n.js';

// The same chunk as the full layout's Changes tab (presets.tsx).
const ReviewPanel = lazy(() => import('./panels/ReviewPanel.js'));

/** The changes as a sheet (ED-013 P5): the review panel, closing when it puts the caret in a change, asking before any "all". */
export function ReviewSheet({ editor, session, onStatus, onTrackChanges, onClose }: { editor: DocxEditorView; session: EditorDocument; onStatus: (text: string) => void; onTrackChanges: () => void; onClose: () => void }) {
  const t = useT();
  return (
    <Sheet title={t('review.panel')} onClose={onClose} testId="review-sheet" className="review-sheet">
      <Suspense fallback={null}>
        {/* The hint's "Track my changes" may ask a name in a dialog, which must not open under the sheet. */}
        <ReviewPanel editor={editor} session={session} onStatus={onStatus} onTrackChanges={() => { onClose(); onTrackChanges(); }} onSelect={onClose} confirmAll />
      </Suspense>
    </Sheet>
  );
}

export interface ThreadSheetProps {
  editor: DocxEditorView;
  session: EditorDocument;
  /** The listing now: a reply, a resolve or a peer's comment shows as it comes. */
  threads: readonly CommentThread[];
  /** The threads the sheet was opened for, the ones at the caret then. */
  ids: readonly string[];
  /** This page's name: its own comments are the ones it may delete. */
  me: string;
  /** With no name set: the name to start from, asked above the field. */
  askName?: string | undefined;
  onReply: (id: string, text: string, name?: string) => void;
  onStatus: (text: string) => void;
  onClose: () => void;
}

/**
 * The threads at the caret (ED-013 P6). The bodies are read-only (editing one's own comment is the full layout's), Reply
 * points the field at its thread, and the field replies to the first open thread until then. The sheet closes when the
 * threads it showed are all gone (deleted, here or by a peer).
 */
export function ThreadSheet({ editor, session, threads, ids, me, askName, onReply, onStatus, onClose }: ThreadSheetProps) {
  const t = useT();
  const lang = useLanguage();
  const shown = threads.filter((th) => ids.includes(th.id));
  const [target, setTarget] = useState<string | undefined>(undefined);
  const replyTo = shown.find((th) => th.id === target && !th.done) ?? shown.find((th) => !th.done);
  const field = useRef<HTMLTextAreaElement>(null);
  const closeRef = useRef(onClose);
  closeRef.current = onClose;
  // Closed when the threads it showed are gone, never before the listing has come (it is empty until the part is read).
  const showed = useRef(false);
  useEffect(() => {
    if (shown.length > 0) showed.current = true;
    else if (showed.current) closeRef.current();
  }, [shown.length]);
  const statusRef = useRef(onStatus);
  statusRef.current = onStatus;
  const actions = useMemo<ThreadActions>(() => ({
    reply: () => undefined,
    resolve: (id, value) => { void editor.runAsync('comment.setResolved', { id, value }); },
    remove: (id) => { void editor.runAsync('comment.delete', { id }).then((ok) => { if (ok) statusRef.current(t('comments.deleted')); }); },
    claimFocus: () => false,
    onStatus: (text) => statusRef.current(text),
    sheet: { onReply: (id) => { setTarget(id); field.current?.focus(); } },
  }), [editor, t]);
  /** A comment's head: its text selected in the document, the sheet closed (the caret taken first, so the focus is not given back over it). */
  const select = (id: string): void => {
    const r = commentRange(editor, id);
    if (!r) return;
    editor.reveal(r.from, r.to);
    editor.caretAt(r.from, r.to);
    onClose();
  };
  return (
    <Sheet title={t('compact.comments.title')} onClose={onClose} testId="thread-sheet" className="thread-sheet">
      {shown.map((th) => (
        <div key={th.id} className={`sheet-thread${th.done ? ' done' : ''}`} data-testid="sheet-thread" data-comment-id={th.id}
          role="group" aria-label={t('comments.by', { author: th.author })} style={{ '--comment-color': authorColor(th.author) } as CSSProperties}>
          {th.done && <p className="sheet-thread-state" data-testid="sheet-thread-resolved">{t('compact.comments.resolved')}</p>}
          <Thread thread={th} session={session} editor={editor} me={me} onSelect={select} lang={lang} actions={actions} />
        </div>
      ))}
      {replyTo && (
        <CommentField testPrefix="thread-sheet" fieldRef={field} label={t('comments.reply.label', { author: replyTo.author })} askName={askName}
          post={t('comments.reply')} onPost={(text, name) => onReply(replyTo.id, text, name)} />
      )}
    </Sheet>
  );
}

/** A new comment on the selection (ED-013 section 3.7): the text quoted, the field, Cancel and Comment. */
export function NewCommentSheet({ quote, askName, onPost, onClose }: { quote: string; askName?: string | undefined; onPost: (text: string, name?: string) => void; onClose: () => void }) {
  const t = useT();
  return (
    <Sheet title={t('comments.new')} onClose={onClose} testId="comment-sheet">
      {quote && <blockquote className="sheet-quote" dir="auto" data-testid="comment-sheet-quote">{quote}</blockquote>}
      {/* The field takes the focus as the sheet opens (a field's autoFocus wins, ED-005 section 12.41), so the keyboard stays.
          The sheet closes first, giving the focus back to the document, then the comment is made there. */}
      <CommentField testPrefix="comment-sheet" label={t('compact.comment.text')} askName={askName} post={t('compact.comment')} autoFocus
        onCancel={onClose} onPost={(text, name) => { onClose(); onPost(text, name); }} />
    </Sheet>
  );
}

interface CommentFieldProps {
  /** The test ids' prefix: `<prefix>-text`, `-name`, `-post`, `-cancel`. */
  testPrefix: string;
  label: string;
  askName?: string | undefined;
  /** The post button's words. */
  post: string;
  autoFocus?: boolean;
  fieldRef?: React.RefObject<HTMLTextAreaElement | null>;
  onPost: (text: string, name?: string) => void;
  onCancel?: () => void;
}

/** The comment sheets' one field: the name when it is asked, the text, and the buttons. Emptied after a post. */
function CommentField({ testPrefix, label, askName, post, autoFocus = false, fieldRef, onPost, onCancel }: CommentFieldProps) {
  const t = useT();
  const [text, setText] = useState('');
  const [name, setName] = useState(askName ?? '');
  const ready = text.trim() !== '' && (askName === undefined || name.trim() !== '');
  const actions = useRef<HTMLDivElement>(null);
  // While the field has the focus its buttons stay in view: the keyboard opens after the focus and shortens the sheet,
  // and on the phone Reply was under it until the sheet was scrolled (ED-013 section 9, S4's phone check). The sheet's
  // body is scrolled, never the page, as the viewport resizes (the keyboard) and as the field takes the focus.
  const reveal = (): void => {
    const el = actions.current;
    const body = el?.closest<HTMLElement>('.sheet-body');
    if (!el || !body || !el.parentElement?.contains(document.activeElement)) return;
    const bottom = Math.min(body.getBoundingClientRect().bottom, window.visualViewport?.height ?? window.innerHeight);
    const over = el.getBoundingClientRect().bottom - bottom;
    if (over > 0) body.scrollTop += over;
  };
  const revealRef = useRef(reveal);
  revealRef.current = reveal;
  useEffect(() => {
    const onResize = (): void => revealRef.current();
    const vv = window.visualViewport;
    vv?.addEventListener('resize', onResize);
    window.addEventListener('resize', onResize);
    return () => { vv?.removeEventListener('resize', onResize); window.removeEventListener('resize', onResize); };
  }, []);
  return (
    <div className="sheet-comment-field" onFocus={() => requestAnimationFrame(() => revealRef.current())}>
      {askName !== undefined && (
        <label className="sheet-field">
          <span>{t('options.name')}</span>
          <input type="text" data-testid={`${testPrefix}-name`} value={name} onChange={(e) => setName(e.target.value)} />
        </label>
      )}
      <textarea ref={fieldRef} className="sheet-text" data-testid={`${testPrefix}-text`} aria-label={label} rows={3} autoFocus={autoFocus}
        value={text} onChange={(e) => setText(e.target.value)} />
      <div ref={actions} className="sheet-actions">
        {onCancel && <button type="button" data-testid={`${testPrefix}-cancel`} onClick={onCancel}>{t('comments.cancel')}</button>}
        <button type="button" className="primary" data-testid={`${testPrefix}-post`} disabled={!ready}
          onClick={() => { const value = text.trim(); setText(''); onPost(value, askName === undefined ? undefined : name.trim()); }}>{post}</button>
      </div>
    </div>
  );
}
