// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 6.5, step D3: the review panel. Mounted over a document with no revision and over
// revisions-word15.docx (a hook after an early return once crashed a panel over one kind of
// document; CLAUDE.md). Its entries are the model's listing with a move's halves paired; the two
// entries of revisions shown in Word only (a table's records one, core-ts check 27) are badged and have no buttons; an Accept lowers the count and one
// undo restores it; an author's Accept all takes that author's alone; a formatting change says
// what changed. The store lists again only when a transaction touches a revision. Next and
// previous visit every entry once and go round. A change's key through the view's live map is the
// key a map rebuilt from the store gives.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { act, StrictMode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { EditorDocument, trackedChangesOf, runCommand, type TrackedChangeInfo } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import ReviewPanel from '../src/app/panels/ReviewPanel.js';
import { valueWords } from '../src/app/review-words.js';
import { ReviewStore, reviewEntries, stepEntry, stepTarget, currentEntry, selectEntry, resolveEntry, type ReviewEntry } from '../src/app/review.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

// jsdom lays nothing out: ProseMirror's scroll-into-view reads rectangles it has not got, so they answer empty (as in story.test.ts).
const zeroRect = (): DOMRect => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON: () => ({}) }) as DOMRect;
if (typeof Range !== 'undefined') {
  Range.prototype.getClientRects ??= () => ([] as unknown as DOMRectList);
  Range.prototype.getBoundingClientRect ??= zeroRect;
}
if (typeof Element !== 'undefined') Element.prototype.getClientRects ??= () => ([] as unknown as DOMRectList);

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

async function open(name: string): Promise<{ session: EditorDocument; editor: DocxEditorView }> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));
  const container = document.createElement('div');
  document.body.appendChild(container);
  return { session, editor: new DocxEditorView(container, session) };
}

async function mount(editor: DocxEditorView, session: EditorDocument): Promise<{ host: HTMLElement; root: Root; status: string[] }> {
  const host = document.createElement('div');
  document.body.appendChild(host);
  const root = createRoot(host);
  const status: string[] = [];
  await act(async () => { root.render(<ReviewPanel editor={editor} session={session} onStatus={(s) => status.push(s)} />); });
  return { host, root, status };
}

/** The number of moves with both halves in a listing: each is one entry. */
const pairedMoves = (changes: readonly TrackedChangeInfo[]): number => {
  const names = new Map<string, number>();
  for (const c of changes) if (c.move !== undefined) names.set(c.move, (names.get(c.move) ?? 0) + 1);
  return [...names.values()].filter((n) => n === 2).length;
};

describe('the review panel (ED-005 section 6.5)', () => {
  test('under React\'s StrictMode (the dev server\'s), the panel follows the document: a change typed after it mounts is listed (section 12.14)', async () => {
    const { session, editor } = await open('loadAndSave.docx');
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Jason';
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    await act(async () => { root.render(<StrictMode><ReviewPanel editor={editor} session={session} onStatus={() => undefined} /></StrictMode>); });
    expect(host.querySelectorAll('[data-testid="review-entry"]').length).toBe(0);
    await act(async () => { editor.caretAt(2); editor.view.dispatch(editor.state.tr.insertText('hhh')); });
    expect(host.querySelectorAll('[data-testid="review-entry"][data-kind="change"]').length).toBe(1);
    expect(host.querySelector('[data-testid="review-count"]')!.textContent).toContain('1 change by 1 person');
    root.unmount();
    editor.destroy();
  });

  test('mounts over a document without revisions', async () => {
    const { session, editor } = await open('loadAndSave.docx');
    const { host, root } = await mount(editor, session);
    expect(host.querySelector('[data-testid="review-count"]')!.textContent).toContain('No tracked changes');
    expect(host.querySelectorAll('[data-testid="review-entry"]').length).toBe(0);
    expect((host.querySelector('[data-testid="review-next"]') as HTMLButtonElement).disabled).toBe(true);
    root.unmount();
    editor.destroy();
  });

  test('revisions-word15.docx: the entries are the listing with a move\'s halves paired, the two in Word only badged, with buttons (the engine resolves them), a formatting change saying what changed', async () => {
    const { session, editor } = await open('revisions/revisions-word15.docx');
    const { host, root } = await mount(editor, session);
    const changes = trackedChangesOf(editor.state.doc);
    const entries = [...host.querySelectorAll<HTMLElement>('[data-testid="review-entry"]')];
    const changeEntries = entries.filter((e) => e.dataset['kind'] === 'change');
    const preserved = entries.filter((e) => e.dataset['kind'] === 'preserved');
    expect(pairedMoves(changes)).toBeGreaterThan(0);
    expect(changeEntries.length).toBe(changes.length - pairedMoves(changes));
    // A table's records are one entry, as Office JS lists them (core-ts check 27): the table's, its grid's and four cells'.
    expect(preserved.map((e) => e.dataset['type'])).toEqual(['tableProperties', 'sectPrChange']);
    for (const e of preserved) {
      expect(e.querySelector('.badge.in-word')!.getAttribute('title')).toContain('Accept and Reject resolve it here');
      expect(e.querySelector('[data-testid="review-accept"]')).not.toBeNull();
    }
    // A move is one entry, "Moved", with the way to its other half.
    const moved = changeEntries.filter((e) => e.querySelector('.type')!.textContent === 'Moved');
    expect(moved.length).toBe(pairedMoves(changes));
    expect(moved[0]!.querySelector('[data-testid="review-move-other"]')).not.toBeNull();
    // A run's formatting change and a paragraph's properties change say what changed.
    const formatted = changeEntries.filter((e) => e.dataset['type'] === 'Formatted');
    expect(formatted.length).toBeGreaterThan(0);
    expect(formatted.some((e) => (e.querySelector('.what')?.textContent ?? '').length > 0)).toBe(true);
    expect(host.querySelector('[data-testid="review-stories"]')).not.toBeNull();
    root.unmount();
    editor.destroy();
  });

  test('an entry shown in Word only accepted: the engine resolves it in a tree run, and it leaves the list (ED-005 section 12.26)', async () => {
    const { session, editor } = await open('revisions/revisions-word15.docx');
    const { host, root, status } = await mount(editor, session);
    const preserved = (): string[] => [...host.querySelectorAll<HTMLElement>('[data-testid="review-entry"][data-kind="preserved"]')].map((e) => e.dataset['type']!);
    expect(preserved()).toEqual(['tableProperties', 'sectPrChange']);
    const accept = host.querySelector('[data-testid="review-entry"][data-type="tableProperties"] [data-testid="review-accept"]') as HTMLButtonElement;
    await act(async () => { accept.click(); });
    expect(status.at(-1)).toBe('Accepted 1 change');
    for (let i = 0; i < 200 && preserved().length > 1; i++) await act(async () => { await new Promise((r) => setTimeout(r, 20)); });
    expect(preserved()).toEqual(['sectPrChange']);
    root.unmount();
    editor.destroy();
  });

  test('an Accept lowers the count and one undo restores it; an author\'s Accept all takes that author\'s alone', async () => {
    const { session, editor } = await open('revisions/revisions-word15.docx');
    const { host, root, status } = await mount(editor, session);
    const count = (): number => host.querySelectorAll('[data-testid="review-entry"][data-kind="change"]').length;
    const before = count();
    const first = host.querySelector('[data-testid="review-entry"][data-kind="change"]:not([data-type="Formatted"]) [data-testid="review-accept"]') as HTMLButtonElement;
    await act(async () => { first.click(); });
    expect(count()).toBe(before - 1);
    expect(status.at(-1)).toBe('Accepted 1 change');
    await act(async () => { editor.run('history.undo'); });
    expect(count()).toBe(before);
    // By author: Author B's all.
    await act(async () => { (host.querySelector('[data-testid="review-by-author"]') as HTMLInputElement).click(); });
    const groups = [...host.querySelectorAll<HTMLElement>('[data-testid="review-author"]')].map((g) => g.dataset['author']);
    expect(groups).toContain('Author A');
    expect(groups).toContain('Author B');
    const b = host.querySelector('[data-testid="review-author"][data-author="Author B"] [data-testid="review-author-accept"]') as HTMLButtonElement;
    await act(async () => { b.click(); });
    expect(trackedChangesOf(editor.state.doc).some((c) => c.author === 'Author B')).toBe(false);
    expect(trackedChangesOf(editor.state.doc).some((c) => c.author === 'Author A')).toBe(true);
    root.unmount();
    editor.destroy();
  });

  /** An entry's every position (its own, its change's, its pieces', a move's other half's), to compare a store with a fresh listing. */
  const positions = (entries: readonly ReviewEntry[]): unknown[] => entries.map((e) => (e.kind === 'change'
    ? [e.id, e.from, e.to, e.change.from, e.change.to, e.change.pieces.map((p) => [p.from, p.to, p.pos]), e.other ? [e.other.from, e.other.to] : null]
    : [e.id, e.from, e.to, e.preserved.pos, e.preserved.records.map((r) => r.pos)]));

  test('after an untracked edit the store\'s entries are a fresh listing\'s, their changes included: selecting and accepting act on the change shown (the review of E4.d, section 12.13)', async () => {
    const { session, editor } = await open('revisions/revisions-word15.docx');
    // Word saved the file with tracking on: off here, so that typing makes no change of its own.
    session.trackingMode = 'Off';
    const store = new ReviewStore(editor, session);
    let notified = 0;
    const unsubscribe = store.subscribe(() => { notified++; });
    // Twenty characters at the start of the first paragraph, untracked: every change after them moves on.
    editor.view.dispatch(editor.state.tr.insertText('x'.repeat(20), 1));
    expect(notified).toBe(1);
    expect(positions(store.entries)).toEqual(positions(reviewEntries(editor.state, session)));
    // The first plain deletion within a paragraph: selecting it selects its text, accepting it resolves it.
    const del = store.entries.find((e) => e.kind === 'change' && e.change.type === 'Deleted' && e.change.move === undefined && !e.change.text.includes('\r')) as Extract<ReviewEntry, { kind: 'change' }>;
    selectEntry(editor, del);
    expect(editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to)).toBe(del.change.text);
    const deletions = (): number => reviewEntries(editor.state, session).filter((e) => e.kind === 'change' && e.change.type === 'Deleted').length;
    const before = deletions();
    expect(resolveEntry(editor, session, del, 'accept')).toBe(1);
    expect(deletions()).toBe(before - 1);
    unsubscribe();
    store.destroy();
    editor.destroy();
  });

  test('an untracked edit that deletes a paragraph holding a change leaves no entry for it', async () => {
    const { session, editor } = await open('revisions/revisions-word15.docx');
    session.trackingMode = 'Off';
    const store = new ReviewStore(editor, session);
    const unsubscribe = store.subscribe(() => undefined);
    const count = store.entries.length;
    // A top-level paragraph holding an insertion and nothing that joins it to its neighbours.
    let target: { pos: number; size: number } | undefined;
    editor.state.doc.forEach((node, pos) => {
      if (target || node.type.name !== 'paragraph' || node.attrs['markRevision'] || node.attrs['pPrChange']) return;
      let ins = false;
      node.descendants((n) => { if (n.isText && n.marks.some((m) => m.type.name === 'insertion' && !m.attrs['move'])) ins = true; return !ins; });
      if (ins) target = { pos, size: node.nodeSize };
    });
    editor.view.dispatch(editor.state.tr.delete(target!.pos, target!.pos + target!.size));
    const fresh = reviewEntries(editor.state, session);
    expect(fresh.length).toBeLessThan(count);
    expect(positions(store.entries)).toEqual(positions(fresh));
    unsubscribe();
    store.destroy();
    editor.destroy();
  });

  test('next and previous visit every entry once and go round', async () => {
    const { session, editor } = await open('revisions/revisions-word15.docx');
    const entries = reviewEntries(editor.state, session);
    editor.caretAt(1);
    const seen: string[] = [];
    for (let i = 0; i < entries.length; i++) {
      expect(stepEntry(editor, session, 1)).toBe('moved');
      const at = currentEntry(editor, reviewEntries(editor.state, session));
      expect(at, `step ${i}`).toBeDefined();
      seen.push(at!.id);
    }
    expect(seen).toEqual(entries.map((e) => e.id));
    expect(stepEntry(editor, session, 1)).toBe('wrapped');
    expect(currentEntry(editor, entries)!.id).toBe(entries[0]!.id);
    expect(stepEntry(editor, session, -1)).toBe('wrapped');
    expect(currentEntry(editor, entries)!.id).toBe(entries[entries.length - 1]!.id);
    editor.destroy();
  });

  test("stepTarget, by which the phone's bar greys Previous and Next (ED-013 section 9): none exactly where a step goes round", async () => {
    const { session, editor } = await open('revisions/revisions-word15.docx');
    const entries = reviewEntries(editor.state, session);
    expect(entries.length).toBeGreaterThan(2);
    const targets = (): [string | undefined, string | undefined] => {
      const current = currentEntry(editor, entries);
      const from = editor.state.selection.from;
      return [stepTarget(entries, current, from, -1)?.id, stepTarget(entries, current, from, 1)?.id];
    };
    // Before the first change: nothing before, the first after.
    editor.caretAt(1);
    expect(currentEntry(editor, entries)).toBeUndefined();
    expect(targets()).toEqual([undefined, entries[0]!.id]);
    // On each change: its neighbours, none before the first or after the last, and a step agrees.
    for (let i = 0; i < entries.length; i++) {
      selectEntry(editor, entries[i]!);
      expect(targets(), `entry ${i}`).toEqual([entries[i - 1]?.id, entries[i + 1]?.id]);
    }
    expect(stepEntry(editor, session, 1, entries)).toBe('wrapped');
    selectEntry(editor, entries[1]!);
    expect(stepEntry(editor, session, -1, entries)).toBe('moved');
    expect(currentEntry(editor, entries)!.id).toBe(entries[0]!.id);
    editor.destroy();
  });

  test('a change\'s key through the view\'s live map is the key a map rebuilt from the store gives', async () => {
    const { session, editor } = await open('revisions/revisions-word15.docx');
    const changes = trackedChangesOf(editor.state.doc);
    for (const c of changes) expect(session.changeKey(c, editor.state)).toBe(session.changeKey(c));
    // And the command finds a change by it.
    const key = session.changeKey(changes[0]!, editor.state);
    expect(editor.run('trackedChange.accept', { key })).toBe(true);
    editor.destroy();
  });
});

describe('the Review perspective offered (ED-005 section 6.7)', () => {
  test('a document with comments and nothing else read-only or tracked has a banner, so it is offered Review', async () => {
    const { diagnosticSummary } = await import('../src/app/labels.js');
    expect(diagnosticSummary({ opaque: [], tracked: { shown: 0, preserved: {} }, pageBreakMarkers: 0, comments: 2 }, 0)).toBe('2 comments.');
    expect(diagnosticSummary({ opaque: [], tracked: { shown: 0, preserved: {} }, pageBreakMarkers: 0, comments: 0 }, 0)).toBe('');
  });

  test('the banner counts the changes it is given, not the file\'s revision elements (ED-005 section 12.113)', async () => {
    const { diagnosticSummary } = await import('../src/app/labels.js');
    expect(diagnosticSummary({ opaque: [], tracked: { shown: 677, preserved: {} }, pageBreakMarkers: 0, comments: 0 }, 131)).toBe('131 tracked changes.');
    expect(diagnosticSummary({ opaque: [], tracked: { shown: 3, preserved: {} }, pageBreakMarkers: 0, comments: 0 }, 0)).toBe('');
  });
});

// Proposal 43 (ED-005 section 9.7): what a change changed, with its values, as Word's Reviewing Pane words it
// (Jason's comparison, section 12.32); a note reference by its number. A style by the name the styles part
// gives it, as the style box shows it: Word's own UI capitalises its built-in names ("Heading 2").
describe('a change\'s values, as Word\'s pane words them (proposal 43)', () => {
  async function shown(name: string): Promise<[string, string, string][]> {
    const { session, editor } = await open(name);
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    await act(async () => { root.render(<StrictMode><ReviewPanel editor={editor} session={session} onStatus={() => undefined} /></StrictMode>); });
    const out = [...host.querySelectorAll<HTMLElement>('[data-testid="review-entry"]')].map((e) => [e.dataset['type']!, e.querySelector('.snippet')?.textContent ?? '', e.querySelector('.what')?.textContent ?? ''] as [string, string, string]);
    await act(async () => root.unmount());
    editor.destroy();
    return out;
  }

  test('review-for-word.docx', async () => {
    const entries = await shown('review/review-for-word.docx');
    expect(entries.filter(([type]) => type === 'Formatted').map(([, text, what]) => [text, what])).toEqual([
      ['alpha', 'Font: Bold'],
      ['bravo', 'Font: 16 pt'],
      ['', 'Centred'],
      ['', 'Left indent: 36 pt'],
      ['', 'heading 2'],
      ['', 'Numbered + Level: 1 + I.'],
      ['docx4j site', 'Hyperlink'],
    ]);
    // The note references, whose text is empty, by the number the view draws.
    expect(entries.filter(([, text]) => text.startsWith('Footnote'))).toEqual([['Added', 'Footnote 1', ''], ['Deleted', 'Footnote 2', '']]);
  });

  test('revisions-word15.docx: two font properties under one "Font:", a list\'s bullet, a colour', async () => {
    const formatted = (await shown('revisions/revisions-word15.docx')).filter(([type]) => type === 'Formatted').map(([, text, what]) => [text.trim(), what]);
    expect(formatted).toEqual([
      ['alpha', 'Font: Bold'],
      ['bravo', 'Font colour: #FF0000'],
      ['charlie', 'Font: Italic, Underline'],
      ['', 'Centred'],
      ['', 'Left indent: 36 pt'],
      ['', 'heading 2'],
      // Word took the paragraphs' own indents away as it made the list: the list's item describes them.
      ['', 'Bulleted + Level: 1 + •'], ['', 'Bulleted + Level: 1 + •'], ['', 'Bulleted + Level: 1 + •'],
    ]);
  });

  test('a property taken away, a hanging indent, line spacing, a highlight by its name', async () => {
    const { session, editor } = await open('tracked-changes.docx');
    const words = (property: string, value: unknown): string => valueWords({ property, value }, editor.state, session, null);
    expect(words('bold', null)).toBe('Font: Not Bold');
    expect(words('vertAlign', null)).toBe('Font: Not superscript or subscript');
    expect(words('strike', 'double')).toBe('Font: Double strikethrough');
    expect(words('underline', 'Wave')).toBe('Font: Underline (Wave)');
    expect(words('font.size', null)).toBe('Font size: (Default)');
    expect(words('font.color', 'auto')).toBe('Font colour: Automatic');
    expect(words('font.highlight', '#FFFF00')).toBe('Highlight: Yellow');
    expect(words('font.highlight', 'darkGray')).toBe('Highlight: Dark grey');
    expect(words('font.highlight', null)).toBe('No highlight');
    expect(words('jc', 'both')).toBe('Justified');
    expect(words('bidi', true)).toBe('Right-to-left');
    expect(words('ind.firstLine', -360)).toBe('Hanging indent: 18 pt');
    expect(words('ind.firstLine', 283)).toBe('First line indent: 14.2 pt');
    expect(words('spacing.after', 120)).toBe('Space after: 6 pt');
    expect(words('spacing.line', { line: 240, rule: 'exact' })).toBe('Line spacing: Exactly 12 pt');
    expect(words('spacing.line', { line: 360, rule: 'auto' })).toBe('Line spacing: 1.5 lines');
    expect(words('spacing.line', { line: 240, rule: null })).toBe('Line spacing: Single');
    expect(words('tabs', [{ pos: 720, val: 'left', leader: null }, { pos: 1440, val: 'clear', leader: null }])).toBe('Tab stops: 36 pt, Not at 72 pt');
    expect(words('numPr', null)).toBe('No bullets or numbering');
    editor.destroy();
  });
});

// Proposal 41 (ED-005 section 9.7): a deletion and the same author's insertion right after it are one entry, resolved
// whole; the split shows the halves apart, kept by the deletion's key across edits and relistings, dropped once one
// half is resolved.
describe('a replacement as one entry (proposal 41)', () => {
  async function mounted(): Promise<{ session: EditorDocument; editor: DocxEditorView; host: HTMLElement; root: Root }> {
    const { session, editor } = await open('review/review-for-word.docx');
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    await act(async () => { root.render(<StrictMode><ReviewPanel editor={editor} session={session} onStatus={() => undefined} /></StrictMode>); });
    return { session, editor, host, root };
  }
  const entryWith = (host: HTMLElement, text: string): HTMLElement[] => [...host.querySelectorAll<HTMLElement>('[data-testid="review-entry"]')].filter((e) => (e.querySelector('.snippet')?.textContent ?? '').includes(text));
  const listed = (editor: DocxEditorView, text: string): string[] => trackedChangesOf(editor.state.doc).filter((c) => c.text.includes(text)).map((c) => c.type);

  test('an insertion first and its deletion after, the order a script\'s Replace writes (core-ts 0.6.0; ED-008 section 8.8 item 9): one entry at the insertion, spanning both, keyed by the deletion', async () => {
    const { session, editor } = await open('review/review-for-word.docx');
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author Z';
    let b = 0;
    for (let i = 0; i < 4; i++) b += session.doc.child(i).nodeSize;
    b += 1;
    const word = session.doc.child(4).textContent.split(' ')[0]!;
    expect(runCommand(session, 'insertText', { text: 'Many' }, { from: b }).ok).toBe(true);
    expect(runCommand(session, 'insertText', { text: '' }, { from: b + 4, to: b + 4 + word.length }).ok).toBe(true);
    await act(async () => undefined);
    const entries = reviewEntries(editor.state, session).filter((e) => e.kind === 'change' && e.change.author === 'Author Z');
    expect(entries).toHaveLength(1);
    const entry = entries[0]!;
    expect(entry.kind === 'change' && entry.inserted?.text).toBe('Many');
    expect(entry.kind === 'change' && entry.change.text).toBe(word);
    expect(entry.kind === 'change' && entry.from < entry.change.from).toBe(true);
    expect(entry.to).toBe(entry.kind === 'change' ? entry.change.to : -1);
    editor.destroy();
  });

  test('one entry, "Replaced", its Accept resolving both halves and nothing else', async () => {
    const { editor, host, root } = await mounted();
    const [pair] = entryWith(host, '7 days');
    expect(pair!.dataset['type']).toBe('Replaced');
    expect(pair!.querySelector('.type')!.textContent).toBe('Replaced');
    expect(pair!.querySelector('.snippet')!.textContent).toBe('“7 days” with “14 days”');
    const before = trackedChangesOf(editor.state.doc).length;
    await act(async () => { pair!.querySelector<HTMLButtonElement>('[data-testid="review-accept"]')!.click(); });
    expect(listed(editor, 'days')).toEqual(['Deleted', 'Added']);   // the agent's "30 days" and "60 days" pair, still there
    expect(trackedChangesOf(editor.state.doc).length).toBe(before - 2);
    expect(entryWith(host, '7 days')).toEqual([]);
    await act(async () => root.unmount());
    editor.destroy();
  });

  test('split: the halves apart, kept across an edit before them and a relisting, dropped when one half is resolved', async () => {
    const { editor, host, root } = await mounted();
    await act(async () => { entryWith(host, '7 days')[0]!.querySelector<HTMLButtonElement>('[data-testid="review-split"]')!.click(); });
    const halves = (): string[] => [...host.querySelectorAll<HTMLElement>('[data-testid="review-entry"][data-split]')].map((e) => `${e.dataset['type']}:${e.querySelector('.snippet')?.textContent ?? ''}`);
    expect(halves()).toEqual(['Deleted:7 days', 'Added:14 days']);
    // An edit before the pair, and so a relisting: the key is a relative position, and the split holds.
    await act(async () => { editor.view.dispatch(editor.state.tr.insertText('Typed. ', 1)); });
    expect(halves()).toEqual(['Deleted:7 days', 'Added:14 days']);
    // Join, and split again: one entry, then two.
    await act(async () => { host.querySelector<HTMLButtonElement>('[data-testid="review-join"]')!.click(); });
    expect(halves()).toEqual([]);
    expect(entryWith(host, '7 days')[0]!.dataset['type']).toBe('Replaced');
    await act(async () => { entryWith(host, '7 days')[0]!.querySelector<HTMLButtonElement>('[data-testid="review-split"]')!.click(); });
    // One half resolved: the other lists as the single change it now is, and the split is gone.
    const deletion = host.querySelector<HTMLElement>('[data-testid="review-entry"][data-split][data-type="Deleted"]')!;
    await act(async () => { deletion.querySelector<HTMLButtonElement>('[data-testid="review-reject"]')!.click(); });
    expect(halves()).toEqual([]);
    expect(entryWith(host, '14 days').map((e) => e.dataset['type'])).toEqual(['Added']);
    await act(async () => root.unmount());
    editor.destroy();
  });

  test('previous and next step over a pair as one, selecting its insertion', async () => {
    const { session, editor, root } = await mounted();
    const entries = reviewEntries(editor.state, session);
    const pair = entries.find((e) => e.kind === 'change' && e.inserted?.text === '14 days')!;
    selectEntry(editor, pair);
    expect(editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to)).toBe('14 days');
    expect(currentEntry(editor, entries)).toBe(pair);
    stepEntry(editor, session, 1, entries);
    const next = currentEntry(editor, entries)!;
    expect(entries.indexOf(next)).toBe(entries.indexOf(pair) + 1);
    stepEntry(editor, session, -1, entries);
    expect(currentEntry(editor, entries)).toBe(pair);
    await act(async () => root.unmount());
    editor.destroy();
  });
});
