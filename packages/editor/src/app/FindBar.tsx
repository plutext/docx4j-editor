// SPDX-License-Identifier: GPL-3.0-or-later
//
// The find bar (ED-005 sections 3.4 and 3.7): a strip docked above the editing surface, so the
// document stays live beside it. The query (pre-filled from a selection inside one paragraph),
// Word's options and regular expressions, "3 of 41", previous and next (Enter, Shift+Enter, F3,
// Shift+F3), and with Ctrl+H a replacement with Replace and Replace all. The body is searched first,
// then the footnotes, the endnotes and the headers and footers in Word's order; a hit in a note or
// a header is selected in that story's view, which becomes the active surface. Every surface of
// the session highlights the query (the find plugin); the matcher is the model's, the one the
// commands, the console and the agent use. The options are remembered per browser.
import { useCallback, useEffect, useRef, useState } from 'react';
import { Finder, hitSelection, replaceAllInStory, runCommand, type EditorDocument, type FindOptions, type Hit, type ReplaceArgs, type Story } from '@docx4j/editor-model';
import type { DocxEditorView } from '../view/DocxEditorView.js';
import { surfaces, type EditorSurface } from '../view/surface.js';
import { findKey, FIND_META, hitAt, indexAtOrAfter, setActiveFinder, type FindMeta } from '../view/plugins/find.js';
import { useEditorState } from './state.js';
import { useT } from './i18n.js';

const OPTIONS_KEY = 'docx4j-editor.find';

/** The options this browser last used. */
export function loadFindOptions(): FindOptions {
  try {
    const v = JSON.parse(localStorage.getItem(OPTIONS_KEY) ?? '{}') as FindOptions;
    return { matchCase: v.matchCase === true, matchWholeWord: v.matchWholeWord === true, matchWildcards: v.matchWildcards === true, matchRegExp: v.matchRegExp === true && v.matchWildcards !== true };
  } catch { return {}; }
}

function saveFindOptions(options: FindOptions): void {
  try { localStorage.setItem(OPTIONS_KEY, JSON.stringify(options)); } catch { /* ignore */ }
}

/** The navigation the shell's F3 and Shift+F3 reach while the bar is open. */
export interface FindNavigation {
  next(): void;
  previous(): void;
}

export interface FindBarProps {
  /** The body view. */
  editor: DocxEditorView;
  session: EditorDocument;
  /** The replacement row shows (Ctrl+H). */
  replace: boolean;
  onReplace: (show: boolean) => void;
  /** Each Ctrl+F or Ctrl+H: the query field takes the focus again, and the selection's text when there is one. */
  request: { seq: number; text?: string };
  /** Escape or ×: the surface holding the current hit, which takes the focus with the hit selected. */
  onClose: (surface: EditorSurface) => void;
  /** A hit in a note or a header: its view becomes the active surface. */
  onActivate: (surface: EditorSurface) => void;
  onStatus: (text: string) => void;
  /** The bar's navigation for the shell's keys, and null when it closes. */
  bind: (nav: FindNavigation | null) => void;
}

/** A story's hits, as the model finds them in the story as it is now. */
interface StoryHits {
  story: Story;
  hits: Hit[];
}

/** The surfaces of the session showing a story (a header shows at every boundary that has it), in the order they were made. */
function surfacesOf(session: EditorDocument, story: Story): EditorSurface[] {
  return [...surfaces.values()].filter((s) => s.session === session && s.story?.key === story.key);
}

/** Every surface of the session, the body first. */
function sessionSurfaces(session: EditorDocument): EditorSurface[] {
  return [...surfaces.values()].filter((s) => s.session === session);
}

export function FindBar({ editor, session, replace, onReplace, request, onClose, onActivate, onStatus, bind }: FindBarProps) {
  const t = useT();
  const [query, setQuery] = useState(request.text ?? '');
  const [replacement, setReplacement] = useState('');
  const [options, setOptions] = useState<FindOptions>(() => loadFindOptions());
  const [finder, setFinder] = useState<Finder | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [stories, setStories] = useState<StoryHits[]>([]);
  /** Where the current hit is: 0 the body (its plugin holds the index), k the k-th entry of `stories`. */
  const [place, setPlace] = useState(0);
  const [storyIndex, setStoryIndex] = useState(-1);
  const input = useRef<HTMLInputElement>(null);
  const state = useEditorState(editor);
  const body = state ? findKey.getState(state) : undefined;

  // A new Ctrl+F or Ctrl+H: the field takes the focus (and the selection's text).
  useEffect(() => {
    if (request.text !== undefined && request.text !== '') setQuery(request.text);
    input.current?.focus();
    input.current?.select();
  }, [request.seq]);   // eslint-disable-line react-hooks/exhaustive-deps

  const argsOf = useCallback((): ReplaceArgs => ({ text: query, replacement, ...options }), [query, replacement, options]);

  /** The stories' hits for the finder, over the stories as they are now. */
  const storyHits = useCallback(async (f: Finder | null): Promise<StoryHits[]> => {
    if (!f) return [];
    const list = await session.stories.inFindOrder(editor.state.doc);
    return list.map((story) => ({ story, hits: f.all(session.stories.currentDoc(story)) })).filter((s) => s.hits.length > 0);
  }, [session, editor]);

  /** The query and options the finder was compiled for, so the pause's compile skips what Enter compiled already. */
  const compiledFor = useRef<string>('');
  const compile = useCallback((q: string): void => {
    const key = JSON.stringify([q, options]);
    if (key === compiledFor.current) return;
    compiledFor.current = key;
    if (q === '') { setFinder(null); setError(null); return; }
    try {
      setFinder(new Finder(q, options, session.ctx.ids, { resolver: session.pkg.propertyResolverOrUndefined }));
      setError(null);
    } catch (e) {
      setFinder(null);
      setError((e as Error).message);
    }
  }, [options, session]);
  useEffect(() => { compiledFor.current = ''; }, [session]);

  // The query or an option changed: compiled after a pause, never per keystroke of the query.
  useEffect(() => {
    const timer = setTimeout(() => compile(query), 150);
    return () => clearTimeout(timer);
  }, [query, compile]);

  // A new finder: every surface of the session highlights it, and the body's first hit from the
  // selection is selected as the query is typed, as Word's navigation pane and VS Code do.
  useEffect(() => {
    setActiveFinder(session, finder);
    for (const s of sessionSurfaces(session)) {
      const meta: FindMeta = s === editor ? { finder } : { finder, current: -1 };
      s.view.dispatch(s.view.state.tr.setMeta(FIND_META, meta));
    }
    const st = findKey.getState(editor.state);
    const first = st && st.current >= 0 ? hitAt(st, st.current) : undefined;
    if (first) { editor.reveal(first.from, first.to); editor.view.dispatch(editor.state.tr.setSelection(hitSelection(editor.state.doc, first)).scrollIntoView()); }
    setPlace(0);
    setStoryIndex(-1);
    let cancelled = false;
    void storyHits(finder).then((found) => { if (!cancelled) setStories(found); });
    return () => { cancelled = true; };
  }, [finder, session, editor, storyHits]);

  // Closing: the highlights go.
  useEffect(() => () => {
    setActiveFinder(session, null);
    for (const s of sessionSurfaces(session)) s.view.dispatch(s.view.state.tr.setMeta(FIND_META, { finder: null } satisfies FindMeta));
  }, [session]);

  const bodyTotal = body?.total ?? 0;
  const counts = [bodyTotal, ...stories.map((s) => s.hits.length)];
  const total = counts.reduce((a, b) => a + b, 0);
  const current = place === 0 ? (body?.current ?? -1) : storyIndex;
  const ordinal = current < 0 ? 0 : counts.slice(0, place).reduce((a, b) => a + b, 0) + current + 1;

  /** Make hit `index` of place `p` the current one: selected in its surface and scrolled to, highlighted as current there and nowhere else. */
  const show = useCallback((p: number, index: number, list: StoryHits[] = stories) => {
    const others = (keep: EditorSurface | undefined): void => {
      for (const s of sessionSurfaces(session)) {
        if (s === keep) continue;
        const st = findKey.getState(s.state);
        if (st && st.current !== -1) s.view.dispatch(s.state.tr.setMeta(FIND_META, { current: -1 } satisfies FindMeta));
      }
    };
    if (p === 0) {
      const st = findKey.getState(editor.state);
      const hit = st ? hitAt(st, index) : undefined;
      if (!hit) return;
      others(editor);
      editor.reveal(hit.from, hit.to);
      editor.view.dispatch(editor.state.tr.setMeta(FIND_META, { current: index } satisfies FindMeta).setSelection(hitSelection(editor.state.doc, hit)).scrollIntoView());
      setPlace(0);
      setStoryIndex(-1);
      return;
    }
    const entry = list[p - 1];
    const hit = entry?.hits[index];
    if (!entry || !hit) return;
    const views = surfacesOf(session, entry.story);
    const target = views[0];
    others(target);
    setPlace(p);
    setStoryIndex(index);
    if (!target) { onStatus(t('find.notShown', { story: entry.story.key })); return; }
    target.view.dispatch(target.state.tr.setMeta(FIND_META, { current: index } satisfies FindMeta).setSelection(hitSelection(target.state.doc, hit)).scrollIntoView());
    onActivate(target);
  }, [editor, session, stories, onActivate, onStatus, t]);

  /** The next or the previous hit, from the body through the stories and round again. */
  const go = useCallback((delta: 1 | -1) => {
    if (!finder || total === 0) return;
    const count = (k: number): number => counts[k] ?? 0;
    const places = counts.length;
    let p = place;
    let i = current;
    if (i < 0) {
      // Nothing current yet: from the body's selection.
      p = 0;
      const st = findKey.getState(editor.state);
      i = st ? indexAtOrAfter(st.blocks, editor.state.selection.from) : -1;
      if (i >= 0 && delta === 1) { show(0, i); return; }
      i = i < 0 ? count(0) : i;
    }
    i += delta;
    for (let guard = 0; guard <= places && (i < 0 || i >= count(p)); guard++) {
      if (i >= count(p)) { p = (p + 1) % places; i = 0; } else { p = (p - 1 + places) % places; i = count(p) - 1; }
    }
    show(p, i);
  }, [finder, total, counts, place, current, editor, show]);

  useEffect(() => {
    bind({ next: () => go(1), previous: () => go(-1) });
  }, [bind, go]);
  useEffect(() => () => bind(null), [bind]);

  /** Replace: the current hit replaced in its story, then the next hit after the replacement shown. */
  const replaceOne = useCallback(async () => {
    if (!finder) return;
    if (current < 0) { go(1); return; }
    const args = argsOf();
    session.undoManager.stopCapturing();
    if (place === 0) {
      const st = findKey.getState(editor.state);
      const hit = st ? hitAt(st, current) : undefined;
      if (!hit) return;
      if (hit.mark) { onStatus(t('find.holdsMark')); return; }
      const after = hit.from + finder.replacementFor(hit, replacement).length;
      editor.view.dispatch(editor.state.tr.setSelection(hitSelection(editor.state.doc, hit)));
      editor.run('find.replace', args);
      session.undoManager.stopCapturing();
      onStatus(t('find.replaced', { count: 1 }));
      const next = indexAtOrAfter(findKey.getState(editor.state)?.blocks ?? [], after);
      if (next >= 0) { show(0, next); return; }
      const found = await storyHits(finder);
      setStories(found);
      if (found.length > 0) show(1, 0, found);
      else if ((findKey.getState(editor.state)?.total ?? 0) > 0) show(0, 0, found);
      return;
    }
    const entry = stories[place - 1];
    const hit = entry?.hits[current];
    if (!entry || !hit) return;
    if (hit.mark) { onStatus(t('find.holdsMark')); return; }
    const after = hit.from + finder.replacementFor(hit, replacement).length;
    const target = surfacesOf(session, entry.story)[0];
    if (target) {
      target.view.dispatch(target.state.tr.setSelection(hitSelection(target.state.doc, hit)));
      target.run('find.replace', args);
    } else {
      runCommand(session, 'find.replace', args, { from: hit.from, to: hit.to }, entry.story);
    }
    session.undoManager.stopCapturing();
    onStatus(t('find.replaced', { count: 1 }));
    const found = await storyHits(finder);
    setStories(found);
    const same = found.findIndex((s) => s.story.key === entry.story.key);
    const nextInStory = same >= 0 ? found[same]!.hits.findIndex((h) => h.from >= after) : -1;
    if (nextInStory >= 0) { show(same + 1, nextInStory, found); return; }
    const later = found.findIndex((s, k) => k > (same >= 0 ? same : place - 2) && s.hits.length > 0);
    if (later >= 0) show(later + 1, 0, found);
    else if ((findKey.getState(editor.state)?.total ?? 0) > 0) show(0, 0, found);
  }, [finder, current, place, stories, argsOf, session, editor, replacement, go, show, storyHits, onStatus, t]);

  /** Replace all: the body, then every story, as one undo step. */
  const replaceAll = useCallback(async () => {
    if (!finder) return;
    const args = argsOf();
    const list = await session.stories.inFindOrder(editor.state.doc);
    // A hit holding a paragraph mark refuses the whole replacement, as the engine's `replaceText` does, before anything changes.
    const bodyHits = findKey.getState(editor.state)?.blocks ?? [];
    if (bodyHits.some((b) => b.hits.some((h) => h.mark)) || stories.some((s) => s.hits.some((h) => h.mark))) { onStatus(t('find.holdsMark')); return; }
    session.undoManager.stopCapturing();
    let count = findKey.getState(editor.state)?.total ?? 0;
    if (count > 0) editor.run('find.replaceAll', args);
    for (const story of list) count += replaceAllInStory(session, story, args);
    session.undoManager.stopCapturing();
    onStatus(t('find.replacedAll', { count }));
    setStories(await storyHits(finder));
    setPlace(0);
    setStoryIndex(-1);
  }, [finder, argsOf, session, editor, stories, storyHits, onStatus, t]);

  const close = useCallback(() => {
    const entry = place === 0 ? undefined : stories[place - 1];
    const target = entry ? surfacesOf(session, entry.story)[0] ?? editor : editor;
    onClose(target);
  }, [place, stories, session, editor, onClose]);

  const setOption = (key: keyof FindOptions, value: boolean): void => {
    const next: FindOptions = { ...options, [key]: value };
    // Word's wildcards and a regular expression are two languages for the one field.
    if (value && key === 'matchWildcards') next.matchRegExp = false;
    if (value && key === 'matchRegExp') next.matchWildcards = false;
    setOptions(next);
    saveFindOptions(next);
  };

  const onQueryKey = (e: React.KeyboardEvent): void => {
    if (e.key === 'Enter' || e.key === 'F3') {
      e.preventDefault();
      // Enter within the pause: the query is compiled now, and its first hit from the selection is shown by the
      // finder's effect, rather than the old query's next hit.
      if (JSON.stringify([query, options]) !== compiledFor.current) { compile(query); return; }
      go(e.shiftKey ? -1 : 1);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      close();
    }
  };
  const onReplacementKey = (e: React.KeyboardEvent): void => {
    if (e.key === 'Enter') { e.preventDefault(); void (e.ctrlKey || e.metaKey ? replaceAll() : replaceOne()); }
    else if (e.key === 'Escape') { e.preventDefault(); close(); }
    else if (e.key === 'F3') { e.preventDefault(); go(e.shiftKey ? -1 : 1); }
  };

  const status = error !== null ? '' : query === '' ? '' : finder === null ? '…' : total === 0 ? t('find.none') : ordinal > 0 ? t('find.count', { current: ordinal, total }) : t('find.total', { total });

  return (
    <div className="find-bar" role="search" aria-label={t('find.bar')} data-testid="find-bar">
      <div className="find-row">
        <button type="button" className="find-toggle-replace" aria-expanded={replace} aria-label={t('find.showReplace')} title={t('find.showReplace')} data-testid="find-toggle-replace" onClick={() => onReplace(!replace)}>{replace ? '▾' : '▸'}</button>
        <input ref={input} type="text" dir="auto" className="find-query" data-testid="find-query" aria-label={t('find.query')} placeholder={t('find.query')} value={query}
          aria-invalid={error !== null} aria-describedby={error !== null ? 'find-error' : undefined}
          onChange={(e) => setQuery(e.target.value)} onKeyDown={onQueryKey} />
        <span className="find-count" data-testid="find-count" aria-live="polite">{status}</span>
        <button type="button" data-testid="find-previous" aria-label={t('find.previous')} title={t('find.previous.title')} disabled={total === 0} onClick={() => go(-1)}>↑</button>
        <button type="button" data-testid="find-next" aria-label={t('find.next')} title={t('find.next.title')} disabled={total === 0} onClick={() => go(1)}>↓</button>
        <label title={t('find.matchCase.title')}><input type="checkbox" data-testid="find-match-case" checked={options.matchCase === true} onChange={(e) => setOption('matchCase', e.target.checked)} /> {t('find.matchCase')}</label>
        <label title={t('find.wholeWords.title')}><input type="checkbox" data-testid="find-whole-words" checked={options.matchWholeWord === true} onChange={(e) => setOption('matchWholeWord', e.target.checked)} /> {t('find.wholeWords')}</label>
        <label title={t('find.wildcards.title')}><input type="checkbox" data-testid="find-wildcards" checked={options.matchWildcards === true} onChange={(e) => setOption('matchWildcards', e.target.checked)} /> {t('find.wildcards')}</label>
        <label title={t('find.regexp.title')}><input type="checkbox" data-testid="find-regexp" checked={options.matchRegExp === true} onChange={(e) => setOption('matchRegExp', e.target.checked)} /> {t('find.regexp')}</label>
        <button type="button" className="find-close" data-testid="find-close" aria-label={t('find.close')} title={t('find.close.title')} onClick={close}>×</button>
      </div>
      {error !== null && <p className="error" role="alert" id="find-error" data-testid="find-error">{error}</p>}
      {replace && (
        <div className="find-row">
          <span className="find-toggle-replace" aria-hidden="true" />
          <input type="text" dir="auto" className="find-query" data-testid="find-replacement" aria-label={t('find.replacement')} placeholder={t('find.replacement')} value={replacement}
            onChange={(e) => setReplacement(e.target.value)} onKeyDown={onReplacementKey} />
          <button type="button" data-testid="find-replace" disabled={total === 0} onClick={() => void replaceOne()} title={t('find.replace.title')}>{t('find.replace')}</button>
          <button type="button" data-testid="find-replace-all" disabled={total === 0} onClick={() => void replaceAll()} title={t('find.replaceAll.title')}>{t('find.replaceAll')}</button>
        </div>
      )}
    </div>
  );
}
