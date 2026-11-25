// SPDX-License-Identifier: GPL-3.0-or-later
//
// The compact layout's chrome (ED-013 section 3.2): one top bar in place of the application bar, the perspective
// strip, the menus and the ribbon. In order: the menu, the file's name, Undo and Redo, the tracking switch, and
// Edit | Review, the two perspectives a phone has. With the keyboard open the bar drops the name alone, so that it does not
// change under a thumb as the focus comes and goes; in landscape with the keyboard open it is not drawn at all (P11). The menu is a sheet: the file actions
// (Send a copy among them, through the system's share sheet where the browser has it), the room's Share and Join,
// Find, Close, the command palette, and Options (the layout and the language among them). Nothing is removed: every
// command stays in the palette (ED-001 decision 9).
//
// The bar above the keyboard (section 3.4) is here too: one row while the keyboard is open and an editing surface has
// the focus, its buttons by where the caret is. The sheets it and the bar open (review, comments) are in
// `CompactSheets.tsx`.
import { useRef, useState, useSyncExternalStore, type ReactNode } from 'react';
import type { Node as PMNode } from 'prosemirror-model';
import { commandNamed, isInTable, type CommentThread, type EditorDocument } from '@docx4j/editor-model';
import type { DocxEditorView } from '../view/DocxEditorView.js';
import type { EditorSurface } from '../view/surface.js';
import type { PerspectiveId } from './perspectives.js';
import { barGroups, type BarContext, type KeyboardState } from './compact.js';
import { ToolButton } from './Toolbar.js';
import { TrackingSwitch } from './TrackingSwitch.js';
import { Sheet } from './Sheet.js';
import { useToolbarReading } from './state.js';
import { currentEntry, resolveEntry, reviewEntries, splitsOf, stepChange, stepTarget, type ReviewEntry } from './review.js';
import { threadsAt } from './comments.js';
import { useT } from './i18n.js';

/** What the menu offers; an entry the browser or the document cannot do now is absent or disabled. */
export interface CompactActions {
  onOpen: () => void;
  onNew: () => void;
  /** Save in place: only where the file has a handle and the option allows it. */
  onSave?: () => void;
  /** Save As: only where the browser has the File System Access pickers. */
  onSaveAs?: () => void;
  onDownload: () => void;
  /** Send a copy through the system's share sheet: only where the browser can hand it a file (section 3.8). */
  onSendCopy?: () => void;
  /** Share: only where a relay is configured. */
  onShare?: () => void;
  onJoin: () => void;
  onFind: () => void;
  onClose: () => void;
  onPalette: () => void;
}

export interface CompactShellProps {
  editor: DocxEditorView | null;
  /** The surface that last had the focus: Undo and Redo act on it, as the ribbon's do. */
  active: EditorSurface | null;
  /** The open file's name, with "•" when it has unsaved changes; undefined with nothing open. */
  name: string | undefined;
  perspective: PerspectiveId;
  /** Edit or Review chosen; Review opens the changes' sheet, chosen again too (section 3.7). */
  onPerspective: (id: PerspectiveId) => void;
  /** The perspectives' labels, in the application's words. */
  labels: { edit: string; review: string };
  tracking: { author: string; onToggle: () => void; onStatus: (text: string) => void; local: { current: boolean } };
  keyboard: KeyboardState;
  actions: CompactActions;
  /** The Options fields, the same as the full layout's Options menu. */
  options: ReactNode;
}

export function CompactShell({ editor, active, name, perspective, onPerspective, labels, tracking, keyboard, actions, options }: CompactShellProps) {
  const t = useT();
  const [menu, setMenu] = useState(false);
  const typing = keyboard.open;
  if (typing && keyboard.landscape) return menu ? <Menu actions={actions} options={options} hasDocument={name !== undefined} onClose={() => setMenu(false)} /> : null;
  const surface = active ?? editor;
  return (
    <>
      <div className="app-bar compact-bar" role="toolbar" aria-label={t('app.bar')} data-testid="compact-bar">
        <button type="button" className="compact-menu-button" aria-label={t('compact.menu')} title={t('compact.menu')} aria-expanded={menu} data-testid="compact-menu"
          onMouseDown={(e) => e.preventDefault()} onClick={() => setMenu(true)}>☰</button>
        {!typing && name !== undefined && <span className="name" data-testid="name">{name}</span>}
        {typing || name === undefined ? <span className="spacer" /> : null}
        {surface && <UndoRedo surface={surface} />}
        {editor && <TrackingSwitch editor={editor} author={tracking.author} onToggle={tracking.onToggle} onStatus={tracking.onStatus} local={tracking.local} short />}
        {/* Review takes the focus as it is tapped, unlike the bar's other buttons: left in the document, a tap raised the
            keyboard over the sheet Review opens (ED-013 section 9). Edit opens nothing, so it keeps the keyboard. */}
        {editor && (
          <div className="perspectives" role="tablist" aria-label={t('app.perspectives')}>
            {(['edit', 'review'] as const).map((id) => (
              <button key={id} type="button" role="tab" aria-selected={id === perspective} data-testid={`perspective-${id}`} className={id === perspective ? 'selected' : ''}
                onMouseDown={id === 'edit' ? (e) => e.preventDefault() : undefined} onClick={() => onPerspective(id)}>{labels[id]}</button>
            ))}
          </div>
        )}
      </div>
      {menu && <Menu actions={actions} options={options} hasDocument={name !== undefined} onClose={() => setMenu(false)} />}
    </>
  );
}

function UndoRedo({ surface }: { surface: EditorSurface }) {
  const t = useT();
  const reading = useToolbarReading(surface);
  const enabled = (n: string): boolean => reading.enabled[n] ?? true;
  return (
    <>
      <ToolButton editor={surface} command="history.undo" label={t('toolbar.undo')} disabled={!enabled('history.undo')}>↶</ToolButton>
      <ToolButton editor={surface} command="history.redo" label={t('toolbar.redo')} disabled={!enabled('history.redo')}>↷</ToolButton>
    </>
  );
}

function Menu({ actions, options, hasDocument, onClose }: { actions: CompactActions; options: ReactNode; hasDocument: boolean; onClose: () => void }) {
  const t = useT();
  /** An entry: the sheet closes, then the action runs, so a dialog it opens takes the focus the sheet gives back. */
  const item = (testId: string, label: string, action: (() => void) | undefined, disabled = false) => action && (
    <button type="button" className="sheet-item" data-testid={testId} disabled={disabled} onClick={() => { onClose(); action(); }}>{label}</button>
  );
  return (
    <Sheet title={t('compact.menu')} onClose={onClose} testId="compact-menu-sheet">
      <div className="sheet-group">
        {item('compact-open', t('app.open'), actions.onOpen)}
        {item('compact-new', t('app.new'), actions.onNew)}
        {item('compact-save', t('app.save'), actions.onSave, !hasDocument)}
        {item('compact-save-as', t('app.saveAs'), actions.onSaveAs, !hasDocument)}
        {item('compact-download', t('app.download'), actions.onDownload, !hasDocument)}
        {item('compact-send-copy', t('app.sendCopy'), actions.onSendCopy, !hasDocument)}
      </div>
      <div className="sheet-group">
        {item('compact-share', t('app.share'), actions.onShare, !hasDocument)}
        {item('compact-join', t('app.join'), actions.onJoin)}
      </div>
      <div className="sheet-group">
        {item('compact-find', t('compact.find'), actions.onFind, !hasDocument)}
        {item('compact-palette', t('compact.palette'), actions.onPalette, !hasDocument)}
        {item('compact-close', t('app.close'), actions.onClose, !hasDocument)}
      </div>
      <details className="sheet-options" data-testid="compact-options">
        <summary>{t('options.menu')}</summary>
        <div className="options-fields">{options}</div>
      </details>
    </Sheet>
  );
}

export interface KeyboardBarProps {
  /** The body's view: Review's buttons act on it alone, as the review panel does (notes and headers list no changes). */
  editor: DocxEditorView;
  /** The surface with the focus, which the rows and Done act on. */
  surface: EditorSurface;
  session: EditorDocument;
  /** The Review perspective. */
  review: boolean;
  /** Landscape with the keyboard open: a strip with Done and one group (P11). */
  landscape: boolean;
  /** The comments as the margin would list them: Comments (n) counts those at the caret. */
  threads: readonly CommentThread[];
  /** Comment on the selection: the new comment's sheet. */
  onComment: () => void;
  /** Comments (n): the sheet of the threads at the caret, by their ids. */
  onComments: (ids: string[]) => void;
  onStatus: (text: string) => void;
}

/**
 * The bar above the keyboard (ED-013 section 3.4, P3): Done, which blurs the surface and so closes the keyboard; in
 * Review, Previous, then Accept and Reject with the caret in a change, then Next, the two greyed at the first change and
 * the last (they do not go round, as the panel's do); Comment on a selection;
 * Comments (n) on commented text; in a table, Row above, Row below and Delete row. Every button keeps the focus where it
 * is (its mouse-down's default prevented), so the keyboard stays up, and every resolution goes through the four commands
 * (ED-005 section 12.15).
 */
export function KeyboardBar({ editor, surface, session, review, landscape, threads, onComment, onComments, onStatus }: KeyboardBarProps) {
  const t = useT();
  const ctx = useBarContext(editor, surface, session, review, threads);
  const groups = barGroups(ctx, landscape);
  const keep = (e: { preventDefault(): void }): void => e.preventDefault();
  /** Previous or Next: greyed by `aria-disabled`, not `disabled`, so that a tap on it still keeps the focus (and the keyboard). */
  const step = (dir: 1 | -1, enabled: boolean, testId: string, label: string) => (
    <button type="button" data-testid={testId} aria-disabled={!enabled} onMouseDown={keep}
      onClick={() => { if (enabled) stepChange(editor, session, dir, onStatus); }}>{label}</button>
  );
  const resolve = (action: 'accept' | 'reject'): void => {
    const entry = currentEntry(editor, reviewEntries(editor.state, session, splitsOf(editor)));
    if (entry) onStatus(t('status.review.resolved', { action, count: resolveEntry(editor, session, entry, action) }));
    editor.focus();
  };
  return (
    <div className="keyboard-bar" role="toolbar" aria-label={t('compact.bar')} data-testid="keyboard-bar">
      {groups.map((group) => (
        <div key={group} className="keyboard-group" role="group" data-testid={`keyboard-group-${group}`}>
          {group === 'review' && (
            <>
              {step(-1, ctx.hasPrevious, 'keyboard-previous', t('review.previous'))}
              {ctx.inChange && <button type="button" data-testid="keyboard-accept" onMouseDown={keep} onClick={() => resolve('accept')}>{t('review.accept')}</button>}
              {ctx.inChange && <button type="button" data-testid="keyboard-reject" onMouseDown={keep} onClick={() => resolve('reject')}>{t('review.reject')}</button>}
              {step(1, ctx.hasNext, 'keyboard-next', t('review.next'))}
            </>
          )}
          {group === 'comment' && <button type="button" data-testid="keyboard-comment" onMouseDown={keep} onClick={onComment}>{t('compact.comment')}</button>}
          {group === 'comments' && (
            <button type="button" data-testid="keyboard-comments" onMouseDown={keep}
              onClick={() => onComments(threadsAt(editor.state, session, threads).map((th) => th.id))}>{t('compact.comments', { count: ctx.comments })}</button>
          )}
          {group === 'table' && (
            <>
              <ToolButton editor={surface} command="table.addRows" args={{ location: 'Before' }} label={t('toolbar.table.addRowAbove')}>{t('compact.rowAbove')}</ToolButton>
              <ToolButton editor={surface} command="table.addRows" args={{ location: 'After' }} label={t('toolbar.table.addRowBelow')}>{t('compact.rowBelow')}</ToolButton>
              <ToolButton editor={surface} command="table.deleteRows" label={t('toolbar.table.deleteRow')}>{t('compact.deleteRow')}</ToolButton>
            </>
          )}
        </div>
      ))}
      <button type="button" className="keyboard-done" data-testid="keyboard-done" aria-label={t('compact.done')} title={t('compact.done')} onMouseDown={keep}
        onClick={() => { const el = document.activeElement; if (el instanceof HTMLElement) el.blur(); }}><HideKeyboardIcon /></button>
    </div>
  );
}

/**
 * What the bar reads, the same object while it is the same, so that a keystroke re-renders nothing (as the toolbar's
 * reading, ED-003 section 8.2). Review's entries are listed again only when the document is another one: a caret
 * moved keeps them (the listing is the model's, memoised per block). The comments at the caret are read from the nodes
 * beside it alone (`threadsAt`), and only in the body, which alone has comments.
 */
function useBarContext(editor: DocxEditorView, surface: EditorSurface, session: EditorDocument, review: boolean, threads: readonly CommentThread[]): BarContext {
  const last = useRef<{ key: string; ctx: BarContext } | null>(null);
  const listed = useRef<{ doc: PMNode; entries: ReviewEntry[] } | null>(null);
  const snapshot = (): BarContext => {
    const state = surface.state;
    let inReview = false;
    let inChange = false;
    let hasPrevious = false;
    let hasNext = false;
    if (review && surface === editor) {
      if (listed.current?.doc !== state.doc) listed.current = { doc: state.doc, entries: reviewEntries(state, session, splitsOf(editor)) };
      const entries = listed.current.entries;
      inReview = entries.length > 0;
      const entry = currentEntry(editor, entries, state);
      inChange = entry !== undefined && (entry.kind === 'change' || entry.target !== undefined);
      hasPrevious = stepTarget(entries, entry, state.selection.from, -1) !== undefined;
      hasNext = stepTarget(entries, entry, state.selection.from, 1) !== undefined;
    }
    const ctx: BarContext = {
      review: inReview,
      inChange,
      hasPrevious,
      hasNext,
      selection: !state.selection.empty && commandNamed('range.insertComment')?.isEnabled?.(state) === true,
      comments: surface === editor ? threadsAt(state, session, threads).length : 0,
      inTable: isInTable(state),
    };
    const key = JSON.stringify(ctx);
    if (last.current?.key === key) return last.current.ctx;
    last.current = { key, ctx };
    return ctx;
  };
  return useSyncExternalStore((onChange) => surface.subscribe(() => onChange()), snapshot, snapshot);
}

/**
 * The keyboard-hiding symbol Android itself shows while a keyboard is open: a keyboard over a downward chevron, drawn in the
 * button's colour. The button's name is in its label and title.
 */
function HideKeyboardIcon() {
  return (
    <svg className="keyboard-done-icon" viewBox="0 0 24 24" width="24" height="24" aria-hidden="true" focusable="false"
      fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round">
      <rect x="2.5" y="2.5" width="19" height="12" rx="2" />
      <path d="M6 6.5h.01M9.3 6.5h.01M12.6 6.5h.01M15.9 6.5h.01M18 6.5h.01M6 9.5h.01M9.3 9.5h.01M12.6 9.5h.01M15.9 9.5h.01M18 9.5h.01" strokeWidth="2.2" />
      <path d="M8 12h8" />
      <path d="M8.5 18l3.5 3.5 3.5-3.5" strokeWidth="1.8" />
    </svg>
  );
}
