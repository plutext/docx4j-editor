// SPDX-License-Identifier: GPL-3.0-or-later
//
// Footnotes and endnotes below the page (ED-002 section 7.2; ED-003 section 6.4): the notes in
// reference order, each with the number the body's decoration paints on its reference and an
// editable view over its fragment. The list follows the body: a transaction that touched a
// reference re-lists (a deleted reference drops its note, undo brings it back); a selected
// reference brings its note into view and flashes it; a note's number selects the reference. The
// parts are read on first display, never at load (decision 5).
import { useCallback, useEffect, useRef, useState, useSyncExternalStore } from 'react';
import { NodeSelection, type Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { noteFocusOf, type EditorDocument, type Note, type NoteKind, type LockRefusal } from '@docx4j/editor-model';
import type { DocxEditorView } from '../../view/DocxEditorView.js';
import type { PasteOptions } from '../../view/paste.js';
import type { EditorSurface } from '../../view/surface.js';
import { EditableStoryView } from '../../view/StoryView.js';
import { decorationsKey } from '../../view/plugins/decorations.js';
import { t, useT } from '../i18n.js';
import { untrackedStatus, lockRefusedStatus } from '../labels.js';
import { useEditorState, selectSource } from '../state.js';

export interface FootnotesPanelProps {
  /** The body view: the references live there, and it owns the stylesheet the notes share. */
  editor: DocxEditorView;
  session: EditorDocument;
  /** A note's view took focus: the application makes it the active surface. */
  onActivate?: (surface: EditorSurface) => void;
  onDestroy?: (surface: EditorSurface) => void;
  /** After a paste into a note (ED-003 section 8.3). */
  onPaste?: PasteOptions['onPaste'];
  onMarkdownOffer?: PasteOptions['onMarkdownOffer'];
  /** The status line: a caret in a note names the note. */
  onStatus?: (text: string) => void;
  /** Application shortcuts bound inside the notes' keymaps too (the palette). */
  shortcuts?: Record<string, () => void>;
}

export interface Notes {
  footnotes: Note[];
  endnotes: Note[];
}

/** The document's notes, read once per session. */
export function useNotes(session: EditorDocument): Notes | undefined {
  const [notes, setNotes] = useState<Notes | undefined>();
  useEffect(() => {
    let cancelled = false;
    setNotes(undefined);
    void Promise.all([session.stories.footnotes(), session.stories.endnotes()]).then(([footnotes, endnotes]) => {
      if (cancelled) return;
      setNotes({ footnotes, endnotes });
      if (footnotes.length + endnotes.length > 0) performance.mark('docx4j:footnotes');   // the benchmark reads it
    });
    return () => { cancelled = true; };
  }, [session]);
  return notes;
}

/** Whether a transaction's changed ranges hold a note reference, before or after: only then is the list computed again. */
export function touchesReference(tr: Transaction): boolean {
  let found = false;
  const has = (doc: PMNode, from: number, to: number): void => {
    doc.nodesBetween(Math.max(0, from), Math.min(to, doc.content.size), (node) => {
      if (found) return false;
      if (node.type.name === 'footnoteRef') found = true;
      return !found;
    });
  };
  tr.steps.forEach((step, i) => {
    if (found) return;
    const rest = tr.mapping.slice(i + 1);
    step.getMap().forEach((oldStart, oldEnd, newStart, newEnd) => {
      has(tr.docs[i]!, oldStart, oldEnd);
      has(tr.doc, rest.map(newStart, -1), rest.map(newEnd, 1));
    });
  });
  return found;
}

/** The notes of a kind for the body as it is now, recomputed after a transaction that touched a reference or a part read. */
function useOrderedNotes(editor: DocxEditorView, session: EditorDocument, kind: NoteKind, loaded: boolean): Note[] {
  const cache = useRef<{ version: number; notes: Note[] }>({ version: -1, notes: [] });
  const version = useRef(0);
  return useSyncExternalStore(
    (onChange) => {
      const off = editor.subscribe((_state, trs) => { if (trs.some(touchesReference)) { version.current++; onChange(); } });
      const off2 = session.stories.subscribe(() => { version.current++; onChange(); });
      return () => { off(); off2(); };
    },
    () => {
      const v = loaded ? version.current : -2;
      if (cache.current.version !== v) cache.current = { version: v, notes: loaded ? session.stories.orderedNotes(kind, editor.state.doc) : [] };
      return cache.current.notes;
    },
    () => [],
  );
}

/** The number a note shows: what the body's decoration paints on its reference (`w:footnotePr` honoured), none for an orphan. */
function numbered(notes: Note[], labels: Map<string, string> | undefined): { note: Note; n: string | null }[] {
  let rank = 0;
  return notes.map((note) => ({ note, n: note.orphan ? null : labels?.get(note.key) ?? String(++rank) }));
}

interface NoteViewProps {
  editor: DocxEditorView;
  session: EditorDocument;
  note: Note;
  label: string;
  /** Take the caret when mounted (a note just inserted, ED-003 section 6.3). */
  focusOnMount?: boolean;
  onActivate?: (surface: EditorSurface) => void;
  onDestroy?: (surface: EditorSurface) => void;
  /** After a paste into a note (ED-003 section 8.3). */
  onPaste?: PasteOptions['onPaste'];
  onMarkdownOffer?: PasteOptions['onMarkdownOffer'];
  /** With tracking on, an edit in a note ran untracked (ED-005 section 5.4). */
  onUntracked?: (what: string) => void;
  /** An edit a content control's lock refused (ED-005 section 12.60). */
  onRefused?: (reason: LockRefusal) => void;
  shortcuts?: Record<string, () => void>;
}

/** A note's editable view, alive as long as the element. */
function NoteView({ editor, session, note, label, focusOnMount, onActivate, onDestroy, onPaste, onMarkdownOffer, onUntracked, onRefused, shortcuts }: NoteViewProps) {
  const ref = useRef<HTMLDivElement>(null);
  const focus = useRef(focusOnMount);
  focus.current = focusOnMount;
  useEffect(() => {
    if (!ref.current) return;
    const view = new EditableStoryView(ref.current, session, note, {
      label, stylesheet: editor,
      ...(onActivate ? { onActivate } : {}), ...(onDestroy ? { onDestroy } : {}), ...(onPaste ? { onPaste } : {}), ...(onMarkdownOffer ? { onMarkdownOffer } : {}), ...(onUntracked ? { onUntracked } : {}), ...(onRefused ? { onRefused } : {}), ...(shortcuts ? { shortcuts } : {}),
    });
    if (focus.current) view.caretAt(view.state.doc.content.size - 1);   // the end of the new note, after its mark and space
    return () => view.destroy();
    // The callbacks are stable for the panel's life; a new label (the numbering moved) remounts, which is rare.
  }, [editor, session, note, label]);   // eslint-disable-line react-hooks/exhaustive-deps
  return <div ref={ref} className="note-body" />;
}

export function FootnotesPanel({ editor, session, onActivate, onDestroy, onPaste, onMarkdownOffer, onStatus, shortcuts }: FootnotesPanelProps) {
  useT();
  const notes = useNotes(session);
  const state = useEditorState(editor);
  const footnotes = useOrderedNotes(editor, session, 'footnote', notes !== undefined);
  const endnotes = useOrderedNotes(editor, session, 'endnote', notes !== undefined);
  const [flash, setFlash] = useState<string | null>(null);
  const root = useRef<HTMLElement>(null);
  // With tracking on, an edit in a note runs untracked and the status line says so (ED-005 section 5.4).
  const untracked = useCallback((what: string) => onStatus?.(untrackedStatus(what)), [onStatus]);
  const refused = useCallback((reason: LockRefusal) => onStatus?.(lockRefusedStatus(reason)), [onStatus]);
  /** The note an insert command asked to take the caret (its transaction meta), until it mounts. */
  const pendingFocus = useRef<string | null>(null);
  useEffect(() => editor.subscribe((_state, trs) => { for (const tr of trs) { const key = noteFocusOf(tr); if (key) pendingFocus.current = key; } }), [editor]);
  const labels = state ? decorationsKey.getState(state)?.noteLabels : undefined;

  // A selected reference (a click on it, or a note's number) brings the note into view.
  useEffect(() => {
    if (!state || !(state.selection instanceof NodeSelection)) return;
    const node = state.selection.node;
    if (node.type.name !== 'footnoteRef') return;
    const key = `${String(node.attrs['kind'])}:${String(node.attrs['footnoteId'])}`;
    const li = root.current?.querySelector<HTMLElement>(`[data-note="${key}"]`);
    if (!li) return;
    li.scrollIntoView({ block: 'nearest' });
    setFlash(key);
    const timer = setTimeout(() => setFlash((f) => (f === key ? null : f)), 1500);
    return () => clearTimeout(timer);
  }, [state]);

  if (!notes || (footnotes.length === 0 && endnotes.length === 0)) return null;

  const list = (kind: NoteKind, items: Note[]) => (
    <ol className="notes" data-testid={`${kind}s`}>
      {numbered(items, labels).map(({ note, n }) => {
        const key = note.key;
        const label = n !== undefined && n !== null ? t(`notes.${kind}`, { n }) : t(`notes.${kind}.orphan`, { id: String(note.id) });
        const focusOnMount = pendingFocus.current === key;
        if (focusOnMount) pendingFocus.current = null;
        return (
          <li key={note.id} data-note={key} data-n={n ?? ''} className={`note${note.orphan ? ' orphan' : ''}${flash === key ? ' flash' : ''}`}
            onFocus={() => onStatus?.(label)} onBlur={(e) => { if (!e.currentTarget.contains(e.relatedTarget as Node | null)) onStatus?.(''); }}>
            <button type="button" className="note-number" title={note.orphan ? t('notes.orphan.title') : t(`notes.goto.${kind}`, { n: n ?? '' })}
              disabled={note.orphan} onClick={() => { const ref = session.stories.referenceOf(editor.state.doc, kind, note.id); if (ref) selectSource(editor, ref); }}>{n ?? '?'}</button>
            <NoteView editor={editor} session={session} note={note} label={label} focusOnMount={focusOnMount} {...(onActivate ? { onActivate } : {})} {...(onDestroy ? { onDestroy } : {})} {...(shortcuts ? { shortcuts } : {})} {...(onPaste ? { onPaste } : {})} {...(onMarkdownOffer ? { onMarkdownOffer } : {})} {...(onStatus ? { onUntracked: untracked, onRefused: refused } : {})} />
            {!note.orphan && (
              <button type="button" className="note-delete" aria-label={t('notes.delete', { note: label })} title={t('notes.delete.title')}
                onMouseDown={(e) => e.preventDefault()} onClick={() => { editor.run('note.delete', { kind, id: note.id }); editor.focus(); }}>×</button>
            )}
          </li>
        );
      })}
    </ol>
  );

  return (
    <section ref={root} className="footnotes-panel" aria-label={t('notes.panel')} data-testid="footnotes-panel">
      {footnotes.length > 0 && <><h2>{t('notes.footnotes')}</h2>{list('footnote', footnotes)}</>}
      {endnotes.length > 0 && <><h2>{t('notes.endnotes')}</h2>{list('endnote', endnotes)}</>}
    </section>
  );
}
