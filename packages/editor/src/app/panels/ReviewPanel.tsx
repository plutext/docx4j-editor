// SPDX-License-Identifier: GPL-3.0-or-later
//
// The review panel (ED-005 section 6.5), the Review perspective's tab: the body's tracked changes
// in document order, or grouped by author, each with its type, author, date, text and what a
// formatting change changed, with its values (proposal 43); a move one entry; a replacement one,
// split on demand into its halves (proposal 41, section 9.7); the revisions shown in Word only among them,
// badged, resolved by the engine (core-ts 0.3.0, section 12.26), with no Accept or Reject for the
// two Word's API does not list (a cell merge, a list's numbering), which only "all" resolves. Next
// and previous, the three views, Accept all and Reject all,
// an author's all. Everything resolves through the four commands (app/review.ts); the list is the
// store's, listed again only when a transaction touches a revision. Notes and headers are not
// listed (section 6.9); the panel's foot says so.
//
// The compact layout shows the same panel in a sheet (ED-013 section 3.7, P5): `onSelect` closes the sheet when an
// entry, Previous or Next puts the selection in the document, and `confirmAll` asks before any "all".
import { memo, useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import type { EditorDocument } from '@docx4j/editor-model';
import type { DocxEditorView } from '../../view/DocxEditorView.js';
import { authorSwatch } from '../../view/authors.js';
import { useEditorState } from '../state.js';
import { t as tr, useLanguage, formatDate } from '../i18n.js';
import { AgentLog } from './AgentLog.js';
import { changeWords, noteReferenceWords } from '../review-words.js';
import { ReviewStore, currentEntry, selectEntry, resolveEntry, resolveAll, hasStories, snippetOf, replacementWords, moveHalves, chooseMarkupView, stepChange, type ReviewEntry } from '../review.js';

export interface ReviewPanelProps {
  editor: DocxEditorView;
  session: EditorDocument;
  onStatus: (text: string) => void;
  /** Turn tracking on through the switch's path (the hint's "Track my changes"). */
  onTrackChanges?: (() => void) | undefined;
  /** After the selection was put on a change (an entry, Previous, Next, a move's other half): the compact layout's sheet closes. */
  onSelect?: (() => void) | undefined;
  /** Accept all, Reject all and an author's all ask first (the compact layout, ED-013 P5). */
  confirmAll?: boolean | undefined;
}

/** Where a browser keeps that the tracking hint was answered (ED-005 section 6.7: shown once, never a switch flipped). */
const HINT_KEY = 'docx4j-editor.review.trackingHint';

function hintAnswered(): boolean {
  try { return localStorage.getItem(HINT_KEY) === 'answered'; } catch { return false; }
}

/** A change's type in words. */
function typeLabel(entry: ReviewEntry): string {
  if (entry.kind === 'preserved') return tr(`review.preserved.kind.${entry.preserved.kind}`);
  const c = entry.change;
  if (c.move !== undefined) return tr('review.type.moved');
  if (entry.inserted) return tr('review.type.replaced');
  // A change of rows alone, several rows one change when they touch (core-ts check 31).
  if (c.row) return tr(c.type === 'Added' ? 'review.type.rowAdded' : 'review.type.rowDeleted', { count: c.pieces.filter((p) => p.what === 'row').length });
  if (c.type === 'Formatted') return tr(c.pieces[0]!.what === 'properties' ? 'review.type.properties' : 'review.type.formatted');
  return tr(c.type === 'Added' ? 'review.type.added' : 'review.type.deleted');
}

/** One entry; `lang` is a prop so the memo re-renders when the language changes. */
const Entry = memo(function Entry({ entry, current, editor, session, store, onStatus, onSelect, lang }: { entry: ReviewEntry; current: boolean; editor: DocxEditorView; session: EditorDocument; store: ReviewStore; onStatus: (s: string) => void; onSelect: (() => void) | undefined; lang: string }) {
  void lang;
  const ref = useRef<HTMLLIElement>(null);
  useEffect(() => { if (current) ref.current?.scrollIntoView?.({ block: 'nearest' }); }, [current]);
  const author = entry.kind === 'change' ? entry.change.author : entry.preserved.author;
  const date = entry.kind === 'change' ? entry.change.date : entry.preserved.date;
  const halves = moveHalves(entry);
  const brokenMove = entry.kind === 'change' && entry.change.move === undefined && entry.change.pieces.some((p) => p.move !== null);
  // A note reference has no text: the note and its number, as Word's pane shows it (proposal 43).
  const text = entry.kind !== 'change' ? '' : entry.inserted ? replacementWords(entry.change, entry.inserted) : snippetOf((halves?.destination ?? entry.change).text) || (noteReferenceWords(editor.state, entry.change) ?? '');
  // A replacement shown split: the join goes on its deletion's entry, the first of the two.
  const join = entry.kind === 'change' && entry.split !== undefined && entry.change.type === 'Deleted' ? entry.split : undefined;
  const what = entry.kind === 'change' && entry.change.type === 'Formatted' ? changeWords(editor.state, session, entry.change) : '';
  // What Accept and Reject resolve, for their accessible names: "Accept: Replaced “7 days” with “14 days”".
  const named = [typeLabel(entry), text || what].filter((w) => w !== '').join(' ');
  const act = (action: 'accept' | 'reject'): void => {
    const n = resolveEntry(editor, session, entry, action);
    onStatus(tr('status.review.resolved', { action, count: n }));
  };
  return (
    <li ref={ref} className={`review-entry${current ? ' current' : ''}${entry.kind === 'preserved' ? ' preserved' : ''}`} data-testid="review-entry"
      data-kind={entry.kind} data-type={entry.kind === 'change' ? (entry.inserted ? 'Replaced' : entry.change.type) : entry.preserved.kind} data-author={author ?? ''} aria-current={current ? 'true' : undefined}
      {...(entry.kind === 'change' && entry.split !== undefined ? { 'data-split': 'true' } : {})}>
      <button type="button" className="review-entry-main" onClick={() => { selectEntry(editor, entry); onSelect?.(); }}>
        <span className="type">{typeLabel(entry)}</span>
        {entry.kind === 'preserved' && <span className="badge in-word" title={tr(entry.target ? 'review.preserved.resolvable' : 'review.preserved.allOnly')}>{tr('review.preserved')}</span>}
        <span className="who"><span className="chip" style={authorSwatch(author ?? '')} aria-hidden="true" />{author || tr('review.author.unknown')}</span>
        {date && <span className="meta">{formatDate(date)}</span>}
        {text && <span className="snippet" dir="auto">{text}{brokenMove ? ` (${tr('review.moved')})` : ''}</span>}
        {what && <span className="what">{what}</span>}
      </button>
      {(entry.kind === 'change' || entry.target) && (
        <span className="review-entry-actions">
          <button type="button" data-testid="review-accept" aria-label={tr('review.accept.named', { what: named })} onClick={() => act('accept')}>{tr('review.accept')}</button>
          <button type="button" data-testid="review-reject" aria-label={tr('review.reject.named', { what: named })} onClick={() => act('reject')}>{tr('review.reject')}</button>
          {halves?.source && <button type="button" className="link" data-testid="review-move-other" onClick={() => { selectEntry(editor, entry, 'source'); onSelect?.(); }}>{tr('review.move.other')}</button>}
          {entry.kind === 'change' && entry.inserted && <button type="button" className="link" data-testid="review-split" aria-expanded="false" title={tr('review.split.title')} onClick={() => store.split(entry)}>{tr('review.split')}</button>}
          {join !== undefined && <button type="button" className="link" data-testid="review-join" aria-expanded="true" title={tr('review.join.title')} onClick={() => store.join(join)}>{tr('review.join')}</button>}
        </span>
      )}
    </li>
  );
});

export default function ReviewPanel({ editor, session, onStatus, onTrackChanges, onSelect, confirmAll = false }: ReviewPanelProps) {
  const lang = useLanguage();
  // The store subscribes to the editor only while React subscribes to it (useSyncExternalStore), so StrictMode's extra mount and unmount leave it listening.
  const store = useMemo(() => new ReviewStore(editor, session), [editor, session]);
  const subscribe = useCallback((fn: () => void) => store.subscribe(fn), [store]);
  useSyncExternalStore(subscribe, () => store.version);
  const state = useEditorState(editor);
  const [byAuthor, setByAuthor] = useState(false);
  const [answered, setAnswered] = useState(hintAnswered);
  const answer = (): void => { setAnswered(true); try { localStorage.setItem(HINT_KEY, 'answered'); } catch { /* ignore */ } };
  const entries = store.entries;
  const current = state ? currentEntry(editor, entries, state) : undefined;
  const changes = entries.filter((e) => e.kind === 'change');
  const preserved = entries.length - changes.length;
  const people = new Set(entries.map((e) => (e.kind === 'change' ? e.change.author : e.preserved.author) ?? '')).size;
  const view = editor.markupView;
  /** An "all" waiting for its confirmation (`confirmAll`), with the number it would resolve. */
  const [asking, setAsking] = useState<{ action: 'accept' | 'reject'; author: string | undefined; count: number } | null>(null);
  const resolveAllNow = (action: 'accept' | 'reject', author?: string): void => {
    setAsking(null);
    onStatus(tr('status.review.resolved', { action, count: resolveAll(editor, session, action, author) }));
  };
  const all = (action: 'accept' | 'reject', author?: string): void => {
    if (!confirmAll) { resolveAllNow(action, author); return; }
    const count = author === undefined ? entries.length : entries.filter((e) => ((e.kind === 'change' ? e.change.author : e.preserved.author) ?? '') === author).length;
    setAsking({ action, author, count });
  };
  const step = (dir: 1 | -1): void => { stepChange(editor, session, dir, onStatus); if (entries.length > 0) onSelect?.(); };
  // Under the buttons that asked: the head's, or the author's group's, which may be far down the list.
  const confirmation = asking && (
    <p className="review-confirm" role="alert" data-testid="review-confirm">
      {asking.author === undefined
        ? tr('review.confirm.all', { action: asking.action, count: asking.count })
        : tr('review.confirm.author', { action: asking.action, count: asking.count, author: asking.author || tr('review.author.unknown') })}{' '}
      <button type="button" className="primary" data-testid="review-confirm-yes" onClick={() => resolveAllNow(asking.action, asking.author)}>
        {tr(asking.action === 'accept' ? 'review.acceptAll' : 'review.rejectAll')}
      </button>{' '}
      <button type="button" data-testid="review-confirm-no" onClick={() => setAsking(null)}>{tr('review.confirm.cancel')}</button>
    </p>
  );
  const groups = byAuthor
    ? [...new Set(entries.map((e) => (e.kind === 'change' ? e.change.author : e.preserved.author) ?? ''))].sort((a, b) => a.localeCompare(b, lang)).map((author) => ({ author, entries: entries.filter((e) => ((e.kind === 'change' ? e.change.author : e.preserved.author) ?? '') === author) }))
    : [{ author: undefined as string | undefined, entries }];
  const list = (items: readonly ReviewEntry[]) => (
    <ul className="review-list">
      {items.map((e) => <Entry key={e.id} entry={e} current={e === current} editor={editor} session={session} store={store} onStatus={onStatus} onSelect={onSelect} lang={lang} />)}
    </ul>
  );
  return (
    <section className="panel review-panel" aria-label={tr('review.panel')} data-testid="panel-review">
      <h2>{tr('review.panel')}</h2>
      {session.trackingMode === 'Off' && !answered && (
        <p className="review-hint" data-testid="review-hint">
          {tr('review.hint')}{' '}
          <button type="button" data-testid="review-hint-track" onClick={() => { answer(); onTrackChanges?.(); }}>{tr('review.hint.track')}</button>{' '}
          <button type="button" className="link" data-testid="review-hint-dismiss" onClick={answer}>{tr('review.hint.dismiss')}</button>
        </p>
      )}
      <p className="meta" data-testid="review-count">
        {changes.length === 0 ? tr('review.empty') : tr('review.count', { changes: changes.length, people })}
        {preserved > 0 && <> {tr('review.preservedCount', { count: preserved })}</>}
      </p>
      <div className="review-head" role="toolbar" aria-label={tr('review.panel')}>
        <button type="button" data-testid="review-previous" disabled={entries.length === 0} onClick={() => step(-1)}>{tr('review.previous')}</button>
        <button type="button" data-testid="review-next" disabled={entries.length === 0} onClick={() => step(1)}>{tr('review.next')}</button>
        <button type="button" data-testid="review-accept-all" disabled={entries.length === 0} onClick={() => all('accept')}>{tr('review.acceptAll')}</button>
        <button type="button" data-testid="review-reject-all" disabled={entries.length === 0} onClick={() => all('reject')}>{tr('review.rejectAll')}</button>
      </div>
      {asking && asking.author === undefined && confirmation}
      <fieldset className="review-view" data-testid="review-view">
        <legend>{tr('review.view')}</legend>
        {(['all', 'accepted', 'original'] as const).map((v) => (
          <label key={v}><input type="radio" name="review-view" value={v} data-testid={`review-view-${v}`} checked={view === v} onChange={() => chooseMarkupView(editor, v, onStatus)} /> {tr(`review.view.${v}`)}</label>
        ))}
      </fieldset>
      <label className="review-by-author"><input type="checkbox" data-testid="review-by-author" checked={byAuthor} onChange={(e) => setByAuthor(e.target.checked)} /> {tr('review.byAuthor')}</label>
      {groups.map((g) => g.author === undefined ? <div key="all">{list(g.entries)}</div> : (
        <section key={`author:${g.author}`} className="review-author" data-testid="review-author" data-author={g.author}>
          <h3><span className="chip" style={authorSwatch(g.author)} aria-hidden="true" />{g.author || tr('review.author.unknown')}</h3>
          <div className="review-author-actions">
            <button type="button" data-testid="review-author-accept" onClick={() => all('accept', g.author)}>{tr('review.author.acceptAll', { author: g.author || tr('review.author.unknown') })}</button>
            <button type="button" data-testid="review-author-reject" onClick={() => all('reject', g.author)}>{tr('review.author.rejectAll', { author: g.author || tr('review.author.unknown') })}</button>
          </div>
          {asking && asking.author === g.author && confirmation}
          {list(g.entries)}
        </section>
      ))}
      <AgentLog editor={editor} session={session} onStatus={onStatus} lang={lang} />
      {hasStories(session) && <p className="meta" data-testid="review-stories">{tr('review.stories')}</p>}
    </section>
  );
}
