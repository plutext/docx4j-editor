// SPDX-License-Identifier: GPL-3.0-or-later
//
// The review panel's model (ED-005 section 6.5): its entries over the model's listing (a move's
// two halves one entry; the revisions shown in Word only among them, resolved by the engine through
// the same commands since core-ts 0.3.0, section 12.26), selecting one, stepping to
// the next or the previous, and resolving through the four commands, never a path of its own
// (section 6.8: the room repair after E4.d hooks their transaction). The listing is the model's,
// memoised per block, and the panel's store lists again on every transaction that changes the
// document: one source of truth. It once mapped its entries through an edit that touched no
// revision, and the changes inside them kept their old positions (the review of E4.d, section 12.13).
import { TextSelection, type EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { Namespaces } from '@docx4j/core-ts';
import { trackedChangesOf, preservedChangesOf, preservedGroupsOf, preservedTargetOf, replacementPairs, changeOfKey, liveMapping, stateFor, RESOLVED_META, type EditorDocument, type TrackedChangeInfo, type PreservedGroup, type PreservedTarget } from '@docx4j/editor-model';
import type { DocxEditorView } from '../view/DocxEditorView.js';
import { setCurrentChange } from '../view/plugins/review.js';
import type { MarkupView } from '../view/markup.js';
import { t } from './i18n.js';

/** One line of the panel: a change (a move's halves together, a replacement's too), or a revision shown in Word only. */
export type ReviewEntry =
  | {
    kind: 'change'; id: string; from: number; to: number; change: TrackedChangeInfo;
    /** A move's other half. */
    other?: TrackedChangeInfo;
    /** A replacement's insertion (proposal 41, section 9.7): the entry is the pair, `change` its deletion. */
    inserted?: TrackedChangeInfo;
    /** A replacement shown split: its deletion's key, on both halves' entries. */
    split?: string;
  }
  | {
    kind: 'preserved'; id: string; from: number; to: number;
    /** A table's records as one, as Office JS lists them (core-ts check 27). */
    preserved: PreservedGroup;
    /** Where the engine finds it, for Accept and Reject; none for what Word's API does not list (a cell merge, a list's numbering), which only "all" resolves. */
    target?: PreservedTarget;
  };

/**
 * The replacements each view shows split, by their deletion's key (section 9.7): a key is a Yjs
 * relative position, which survives edits before the pair, where an entry's id does not. This
 * viewer's preference, never the room's; a key that no longer starts a pair is dropped at the
 * listing after.
 */
const splitKeys = new WeakMap<DocxEditorView, Set<string>>();

/** The keys of the replacements a view shows split. */
export function splitsOf(editor: DocxEditorView): Set<string> {
  let keys = splitKeys.get(editor);
  if (!keys) splitKeys.set(editor, keys = new Set());
  return keys;
}

/**
 * The panel's entries in document order: a move's halves paired into one, at its first half; a
 * replacement's deletion and insertion one (`replacementPairs`), unless the view shows it split; the
 * revisions shown in Word only among them.
 */
export function reviewEntries(state: EditorState, session: EditorDocument, splits?: ReadonlySet<string>): ReviewEntry[] {
  const changes = trackedChangesOf(state.doc);
  const pairs = replacementPairs(state.doc, changes);
  const deletionOf = new Map([...pairs].map(([d, i]) => [i, d] as const));
  // Each remembered split mapped back to its deletion: one mapping per split, not one per entry.
  const splitAt = new Map<TrackedChangeInfo, string>();
  if (splits && splits.size > 0 && pairs.size > 0) {
    const live = liveMapping(state, session.store.body);
    for (const key of splits) {
      const c = changeOfKey(session.store.body, state.schema, state.doc, key, live);
      const d = c && [...pairs.keys()].find((p) => p.from === c.from && p.type === c.type && p.author === c.author);
      if (d) splitAt.set(d, key);
    }
  }
  const out: ReviewEntry[] = [];
  const moves = new Map<string, TrackedChangeInfo[]>();
  for (const c of changes) if (c.move !== undefined) (moves.get(c.move) ?? moves.set(c.move, []).get(c.move)!).push(c);
  const paired = new Set<string>();
  for (const c of changes) {
    const id = `${c.type}:${c.from}:${c.author ?? ''}:${c.move ?? ''}`;
    if (c.move !== undefined) {
      if (paired.has(c.move)) continue;
      paired.add(c.move);
      const halves = moves.get(c.move)!;
      const other = halves.find((h) => h !== c);
      out.push({ kind: 'change', id, from: c.from, to: c.to, change: c, ...(other ? { other } : {}) });
      continue;
    }
    // A replacement's two halves in the order the document holds them (a person's deletion first, a script's
    // insertion first; ED-008 section 8.8 item 9): one entry at the first half, spanning both, keyed by the
    // deletion; each half on its own when the pair is shown split.
    const deleted = deletionOf.get(c);
    if (deleted) {
      const key = splitAt.get(deleted);
      if (key) out.push({ kind: 'change', id, from: c.from, to: c.to, change: c, split: key });
      else if (c.from < deleted.from) out.push({ kind: 'change', id: `Deleted:${deleted.from}:${deleted.author ?? ''}:`, from: c.from, to: deleted.to, change: deleted, inserted: c });
      continue;
    }
    const inserted = pairs.get(c);
    if (inserted) {
      const key = splitAt.get(c);
      if (key) out.push({ kind: 'change', id, from: c.from, to: c.to, change: c, split: key });
      else if (c.from < inserted.from) out.push({ kind: 'change', id, from: c.from, to: inserted.to, change: c, inserted });
      continue;
    }
    out.push({ kind: 'change', id, from: c.from, to: c.to, change: c });
  }
  const groups = preservedGroupsOf(state.doc, preservedChangesOf(state.doc, (id) => session.ctx.ids.get(id)?.value));
  for (const p of groups) {
    const target = preservedTargetOf(state.doc, p, groups);
    out.push({ kind: 'preserved', id: `preserved:${p.kind}:${p.pos}`, from: p.pos, to: p.pos, preserved: p, ...(target ? { target } : {}) });
  }
  return out.sort((a, b) => a.from - b.from);
}

/** A document's count as it opened, per session. */
const listed = new WeakMap<EditorDocument, number>();

/**
 * The tracked changes of a document as it opened, counted as the review panel lists them, for the banner and the
 * Diagnostics tab. `diagnostics.tracked.shown` counts the file's revision elements, which can be many times that:
 * 677 against about 130 in a file from ONLYOFFICE (ED-005 section 12.113). Read once: it walks the store's document.
 */
export function listedChanges(session: EditorDocument): number {
  let count = listed.get(session);
  if (count === undefined) {
    count = session.diagnostics.tracked.shown === 0 ? 0 : reviewEntries(stateFor(session), session).filter((e) => e.kind === 'change').length;
    listed.set(session, count);
  }
  return count;
}

/** A move entry's destination and source halves (the destination the text as it now stands); undefined for any other entry. */
export function moveHalves(entry: ReviewEntry): { destination: TrackedChangeInfo; source: TrackedChangeInfo | undefined } | undefined {
  if (entry.kind !== 'change' || entry.change.move === undefined) return undefined;
  const halves = [entry.change, entry.other].filter((c): c is TrackedChangeInfo => !!c);
  const destination = halves.find((h) => h.type === 'Added') ?? halves[0]!;
  return { destination, source: halves.find((h) => h !== destination) };
}

/**
 * The panel's list, listed again on every transaction that changes the document. The listing is
 * memoised per block (0.6 ms after a keystroke on the 255-page document, section 12.8), so a
 * second path that mapped positions saved nothing worth its staleness.
 */
export class ReviewStore {
  entries: ReviewEntry[];
  /** Bumped on every change, for React. */
  version = 0;
  private readonly listeners = new Set<() => void>();
  /** The editor subscription, held while the store has a listener. */
  private off: (() => void) | undefined;

  /** The document the entries were listed from. */
  private doc: PMNode;

  constructor(private readonly editor: DocxEditorView, private readonly session: EditorDocument) {
    this.entries = this.list(editor.state);
    this.doc = editor.state.doc;
  }

  /** The entries, with this view's split replacements; a split whose pair is gone forgotten. */
  private list(state: EditorState): ReviewEntry[] {
    const splits = splitsOf(this.editor);
    const entries = reviewEntries(state, this.session, splits);
    for (const key of [...splits]) if (!entries.some((e) => e.kind === 'change' && e.split === key)) splits.delete(key);
    return entries;
  }

  /** Listed again when the document is another one, whichever transaction made it (a listener sees the root transaction, not those plugins append). */
  private refresh(state: EditorState): boolean {
    if (state.doc === this.doc) return false;
    this.doc = state.doc;
    this.entries = this.list(state);
    this.version++;
    return true;
  }

  /** Show a replacement's two halves apart, each with its own Accept and Reject (section 9.7). */
  split(entry: ReviewEntry): void {
    if (entry.kind !== 'change' || !entry.inserted) return;
    splitsOf(this.editor).add(this.session.changeKey(entry.change, this.editor.state));
    this.relist();
  }

  /** Show a split replacement as one entry again. */
  join(key: string): void {
    splitsOf(this.editor).delete(key);
    this.relist();
  }

  private relist(): void {
    this.entries = this.list(this.editor.state);
    this.doc = this.editor.state.doc;
    this.version++;
    for (const l of this.listeners) l();
  }

  /**
   * A listener, and the editor subscribed while there is one: through React's `useSyncExternalStore`,
   * which subscribes and unsubscribes as a component mounts and unmounts. Subscribing in the
   * constructor and unsubscribing in an effect's cleanup left the panel deaf under React's
   * StrictMode, which runs that cleanup once at mount (the dev server's; ED-005 section 12.14).
   */
  subscribe(fn: () => void): () => void {
    this.listeners.add(fn);
    if (!this.off) this.off = this.editor.subscribe((state) => { if (this.refresh(state)) for (const l of this.listeners) l(); });
    // Whatever changed while nobody listened: React reads the version again after subscribing.
    this.refresh(this.editor.state);
    return () => {
      this.listeners.delete(fn);
      if (this.listeners.size === 0) { this.off?.(); this.off = undefined; }
    };
  }

  destroy(): void {
    this.off?.();
    this.off = undefined;
    this.listeners.clear();
  }
}

/**
 * The entry last selected in each view and the selection it left: while the selection is still
 * that, it is the current entry. Two changes can start at one place (another author's insertion
 * deleted), and the selection alone names only the first of them.
 */
const navigated = new WeakMap<DocxEditorView, { id: string; from: number; to: number }>();

/** The current entry: the one last selected while the selection is still its, else the one the selection is in. */
export function currentEntry(editor: DocxEditorView, entries: readonly ReviewEntry[], state: EditorState = editor.state): ReviewEntry | undefined {
  const last = navigated.get(editor);
  if (last && state.selection.from === last.from && state.selection.to === last.to) {
    const found = entries.find((e) => e.id === last.id);
    if (found) return found;
  }
  return entryAtSelection(entries, state);
}

/** The entry the selection is in (a move's either half), if any. */
export function entryAtSelection(entries: readonly ReviewEntry[], state: EditorState): ReviewEntry | undefined {
  const { from, to } = state.selection;
  const inside = (a: number, b: number): boolean => (from === to ? a <= from && from <= b : a < to && b > from);
  return entries.find((e) => {
    if (e.kind === 'preserved') {
      const node = state.doc.nodeAt(e.from);
      return !!node && inside(e.from, e.from + node.nodeSize);
    }
    return inside(e.change.from, e.change.to) || (!!e.other && inside(e.other.from, e.other.to)) || (!!e.inserted && inside(e.inserted.from, e.inserted.to));
  });
}

/** Where selecting an entry puts the selection, and the ranges drawn as the current change. */
function targetOf(entry: ReviewEntry, doc: PMNode): { selection: [number, number]; ranges: [number, number][] } {
  const inner = (pos: number): [number, number] => {
    if (pos >= doc.content.size) { const end = TextSelection.near(doc.resolve(doc.content.size), -1); return [end.from, end.from]; }
    const node = doc.nodeAt(pos);
    if (node?.isTextblock) return [pos + 1, pos + node.nodeSize - 1];
    const near = TextSelection.near(doc.resolve(pos + 1));
    return [near.from, near.from];
  };
  if (entry.kind === 'preserved') {
    const node = doc.nodeAt(entry.from);
    const sel = inner(entry.from);
    return { selection: [sel[0], sel[0]], ranges: node ? [[entry.from, entry.from + node.nodeSize]] : [] };
  }
  const halves = moveHalves(entry);
  // A replacement selects its insertion, the text as it now stands, and draws both halves (section 9.7).
  const c = halves?.destination ?? entry.inserted ?? entry.change;
  const ranges: [number, number][] = [[c.from, c.to]];
  if (halves?.source) ranges.push([halves.source.from, halves.source.to]);
  if (entry.inserted) ranges.push([entry.change.from, entry.change.to]);
  const first = c.pieces[0]!;
  // A row or a paragraph's properties: the text inside; a paragraph mark alone: the caret before it.
  if (first.what === 'row' || first.what === 'properties') { const sel = inner(c.from); return { selection: sel, ranges }; }
  const $from = doc.resolve(c.from);
  if (!$from.parent.inlineContent) { const near = TextSelection.near($from); return { selection: [near.from, near.from], ranges }; }
  return { selection: [c.from, Math.max(c.from, Math.min(c.to, $from.end()))], ranges };
}

/**
 * Select an entry (a click in the panel, next or previous): the view switched to All markup when
 * it hides the entry (section 6.4), the selection on it (a move's destination), both halves of a
 * move drawn as the current change, scrolled to.
 */
export function selectEntry(editor: DocxEditorView, entry: ReviewEntry, half: 'destination' | 'source' = 'destination'): void {
  const doc = editor.state.doc;
  const target = targetOf(entry, doc);
  let [from, to] = target.selection;
  const halves = moveHalves(entry);
  if (half === 'source' && halves?.source) { from = halves.source.from; to = halves.source.to; }
  editor.reveal(from, to);
  const state = editor.state;
  const clamp = (p: number): number => Math.max(0, Math.min(p, state.doc.content.size));
  let sel;
  try { sel = TextSelection.create(state.doc, clamp(from), clamp(to)); } catch { sel = TextSelection.near(state.doc.resolve(clamp(from))); }
  editor.view.dispatch(setCurrentChange(state.tr.setSelection(sel), target.ranges).scrollIntoView());
  navigated.set(editor, { id: entry.id, from: editor.state.selection.from, to: editor.state.selection.to });
  editor.focus();
}

/**
 * The entry a step goes to, without going round: the next (`dir` 1) or previous (-1) after `current`, else after the
 * position `at` (the selection's start); none past the last or before the first. The phone's bar greys its Previous and
 * Next by it (ED-013 section 9), so a step and its button read one rule.
 */
export function stepTarget(entries: readonly ReviewEntry[], current: ReviewEntry | undefined, at: number, dir: 1 | -1): ReviewEntry | undefined {
  const i = current ? entries.indexOf(current) : -1;
  if (dir > 0) return i >= 0 ? entries[i + 1] : entries.find((e) => e.from > at);
  return i >= 0 ? entries[i - 1] : [...entries].reverse().find((e) => e.from < at);
}

/**
 * The next (`dir` 1) or previous (-1) entry after the selection, selected; at the end, round to the
 * first (or the last). What happened, for the status line.
 */
export function stepEntry(editor: DocxEditorView, session: EditorDocument, dir: 1 | -1, entries = reviewEntries(editor.state, session, splitsOf(editor))): 'moved' | 'wrapped' | 'none' {
  if (entries.length === 0) return 'none';
  const target = stepTarget(entries, currentEntry(editor, entries), editor.state.selection.from, dir);
  selectEntry(editor, target ?? (dir > 0 ? entries[0]! : entries[entries.length - 1]!));
  return target ? 'moved' : 'wrapped';
}

/** Resolve through a command, and how many changes it resolved (its `RESOLVED_META`). */
function runCounting(editor: DocxEditorView, name: string, args: unknown): number {
  let count = 0;
  const off = editor.subscribe((_state, trs) => {
    for (const tr of trs) { const n = tr.getMeta(RESOLVED_META) as number | undefined; if (typeof n === 'number') count = n; }
  });
  try { editor.run(name, args); } finally { off(); }
  return count;
}

/**
 * Accept or reject one entry, by its change's key through the view's live map: the four commands'
 * path. A revision shown in Word only goes by its target, and the command's tree run resolves it a
 * moment later. The number resolved.
 */
export function resolveEntry(editor: DocxEditorView, session: EditorDocument, entry: ReviewEntry, action: 'accept' | 'reject'): number {
  if (entry.kind === 'preserved') return entry.target && editor.run(`trackedChange.${action}`, { preserved: entry.target }) ? 1 : 0;
  // A replacement: both halves in one transaction, one undo step and one room record, and nothing else in their span.
  if (entry.inserted) return runCounting(editor, `trackedChange.${action}`, { keys: [session.changeKey(entry.change, editor.state), session.changeKey(entry.inserted, editor.state)] });
  const key = session.changeKey(entry.change, editor.state);
  return runCounting(editor, `trackedChange.${action}`, { key });
}

/**
 * Accept or reject all, or all of one author's: the number resolved, the revisions shown in Word
 * only among them (all of them, or the author's that Word's API lists), which the command's tree run
 * resolves after its transaction.
 */
export function resolveAll(editor: DocxEditorView, session: EditorDocument, action: 'accept' | 'reject', author?: string): number {
  const preserved = reviewEntries(editor.state, session).filter((e) => e.kind === 'preserved' && (author === undefined || (e.target !== undefined && (e.preserved.author ?? '') === author))).length;
  return runCounting(editor, action === 'accept' ? 'trackedChanges.acceptAll' : 'trackedChanges.rejectAll', author !== undefined ? { author } : undefined) + preserved;
}

/** Choose a view of the markup (the panel's head, the palette): Original says it is read-only. */
export function chooseMarkupView(editor: DocxEditorView, view: MarkupView, onStatus: (s: string) => void): void {
  editor.markupView = view;
  onStatus(view === 'original' ? t('status.review.originalReadOnly') : '');
}

/** Step to the next or previous change, saying so when it went round (the panel's buttons, Alt+PageDown and Alt+PageUp). */
export function stepChange(editor: DocxEditorView, session: EditorDocument, dir: 1 | -1, onStatus: (s: string) => void): void {
  const r = stepEntry(editor, session, dir);
  onStatus(r === 'none' ? t('review.empty') : r === 'wrapped' ? t('status.review.wrapped', { dir: dir > 0 ? 'next' : 'previous' }) : '');
}

/** Whether the package has notes, headers or footers, whose revisions the panel does not list (section 6.9). */
export function hasStories(session: EditorDocument): boolean {
  const main = session.main as { footnotesPart?: unknown; endnotesPart?: unknown; relationshipsPart?: { getRelationshipsByType(type: string): unknown[] } };
  if (main.footnotesPart || main.endnotesPart) return true;
  const rels = main.relationshipsPart;
  return !!rels && (rels.getRelationshipsByType(Namespaces.HEADER).length > 0 || rels.getRelationshipsByType(Namespaces.FOOTER).length > 0);
}

/** A change's snippet: its text, a paragraph mark or a row's end (`\r\n`) as ¶ and a tab as →, cut at `max` characters (80). */
export function snippetOf(text: string, max = 80): string {
  const s = text.replace(/\r\n?/g, '¶').replace(/\t/g, '→').replace(/[\v\f]/g, '↵');
  return s.length > max ? `${s.slice(0, max - 1)}…` : s;
}

/**
 * A replacement's words (section 9.7): whole paragraphs replaced by whole paragraphs counted
 * ("1 paragraph with 2"), anything else the two texts ("“7 days” with “14 days”").
 */
export function replacementWords(deleted: TrackedChangeInfo, inserted: TrackedChangeInfo): string {
  const marks = (text: string): number => text.split('\r').length - 1;
  if (deleted.text.endsWith('\r') && inserted.text.endsWith('\r')) return t('review.replaced.paragraphs', { old: marks(deleted.text), new: marks(inserted.text) });
  return t('review.replaced.text', { old: snippetOf(deleted.text, 36), new: snippetOf(inserted.text, 36) });
}
