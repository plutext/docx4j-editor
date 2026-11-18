// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 sections 6.3 and 6.4, step D2: tracked changes drawn from the classes the mark and node
// views set, one stylesheet rule per author, and the three views. Over the Word-made fixtures:
// the change bar's class on exactly the blocks holding a revision; a rule per author, a name with
// a quote and a backslash included; the marks', the pilcrows', the rows' and the atoms' classes;
// a paragraph's own classes toggled, so a section's class from a decoration survives an edit.
// The views: the root's class, Original not editable, a caret in hidden text moved to the edge
// that is drawn (the start of a run, or the end of a paragraph's leading run where the anchor
// widget is), Backspace across hidden text deleting the visible character beyond it, a reveal.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { TextSelection } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, trackedChangesOf, type Revision } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { visibleCaret, clampCaret } from '../src/view/markup.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

// jsdom lays nothing out: ProseMirror's scroll-into-view reads rectangles it has not got, so they answer empty (as in story.test.ts).
const zeroRect = (): DOMRect => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON: () => ({}) }) as DOMRect;
if (typeof Range !== 'undefined') {
  Range.prototype.getClientRects ??= () => ([] as unknown as DOMRectList);
  Range.prototype.getBoundingClientRect ??= zeroRect;
}
if (typeof Element !== 'undefined') Element.prototype.getClientRects ??= () => ([] as unknown as DOMRectList);

async function open(name: string): Promise<{ session: EditorDocument; editor: DocxEditorView }> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));
  const container = document.createElement('div');
  document.body.appendChild(container);
  return { session, editor: new DocxEditorView(container, session) };
}

const REVISION_MARKS = new Set(['insertion', 'deletion', 'formatChange']);

/** Whether a paragraph holds a revision, read from the model: what the change bar's class must say. */
function holds(p: PMNode): boolean {
  const a = p.attrs as { markRevision: unknown; pPrChange: unknown; sectPr: { sectPrChange?: unknown } | null; markRPr: { rPrChange?: unknown } | null };
  if (a.markRevision || a.pPrChange || a.sectPr?.sectPrChange || a.markRPr?.rPrChange) return true;
  let found = false;
  p.descendants((n) => { if (n.isText ? n.marks.some((m) => REVISION_MARKS.has(m.type.name)) : n.attrs['revision']) found = true; return !found; });
  return found;
}

/** The first hidden run (deleted text) with visible text on both sides, in the view's document. */
function junction(doc: PMNode): { from: number; to: number } {
  let out: { from: number; to: number } | undefined;
  doc.descendants((node, pos) => {
    if (out || node.type.name !== 'paragraph') return !out;
    const kids: { n: PMNode; p: number }[] = [];
    node.forEach((n, off) => kids.push({ n, p: pos + 1 + off }));
    const del = (k?: { n: PMNode }): boolean => !!k && k.n.isText && k.n.marks.some((m) => m.type.name === 'deletion');
    const vis = (k?: { n: PMNode }): boolean => !!k && k.n.isText && !del(k);
    kids.forEach((k, i) => { if (!out && del(k) && vis(kids[i - 1]) && vis(kids[i + 1])) out = { from: k.p, to: k.p + k.n.nodeSize }; });
    return false;
  });
  if (!out) throw new Error('no junction');
  return out;
}

describe('tracked changes drawn (ED-005 section 6.3)', () => {
  test('revisions-word15.docx: the change bar on exactly the paragraphs and rows that hold a revision', async () => {
    const { editor } = await open('revisions/revisions-word15.docx');
    const doc = editor.state.doc;
    let paragraphs = 0;
    let barred = 0;
    doc.descendants((node, pos) => {
      if (node.type.name === 'paragraph') {
        const dom = editor.view.nodeDOM(pos) as HTMLElement;
        expect(dom.classList.contains('has-change'), `paragraph at ${pos}: "${node.textContent.slice(0, 30)}"`).toBe(holds(node));
        paragraphs++;
        if (holds(node)) barred++;
        return false;
      }
      if (node.type.name === 'table_row') {
        const dom = editor.view.nodeDOM(pos) as HTMLElement;
        let inner = !!node.attrs['revision'];
        node.descendants((n) => { if (n.type.name === 'paragraph' && holds(n)) inner = true; if (n.type.name === 'table_cell' && (n.attrs['tcPr'] as { tcPrChange?: unknown } | null)?.tcPrChange) inner = true; return !inner; });
        expect(dom.classList.contains('has-change'), `row at ${pos}`).toBe(inner);
      }
      return true;
    });
    expect(barred).toBeGreaterThan(10);
    expect(barred).toBeLessThan(paragraphs);
    // The table with its properties changed (shown in Word only) carries the bar along it.
    expect(editor.container.querySelectorAll('.tableWrapper.has-change').length).toBe(1);
    editor.destroy();
  });

  test('revisions-word15.docx: the marks\', the pilcrows\', the rows\' and the atoms\' classes, a move\'s markers as a rule', async () => {
    const { editor } = await open('revisions/revisions-word15.docx');
    const root = editor.container;
    expect(root.querySelectorAll('ins.insertion[data-author="Author A"]').length).toBeGreaterThan(0);
    expect(root.querySelectorAll('del.deletion[data-author="Author B"]').length).toBeGreaterThan(0);
    expect(root.querySelectorAll('ins.insertion.move').length).toBeGreaterThan(0);
    expect(root.querySelectorAll('del.deletion.move').length).toBeGreaterThan(0);
    expect(root.querySelectorAll('span.format-change[data-author]').length).toBeGreaterThan(0);
    expect(root.querySelectorAll('p.rev-mark-ins').length).toBeGreaterThan(0);
    expect(root.querySelectorAll('p.rev-mark-del').length).toBeGreaterThan(0);
    expect(root.querySelectorAll('p.rev-props').length).toBeGreaterThan(0);
    expect(root.querySelectorAll('tr.rev-row-ins').length).toBe(1);
    expect(root.querySelectorAll('tr.rev-row-del').length).toBe(1);
    // Section 10: a deleted tab, an inserted break and note reference.
    expect(root.querySelectorAll('span.tab.rev-del').length).toBe(1);
    expect(root.querySelectorAll('sup.footnote-ref.rev-ins').length).toBe(1);
    expect(root.querySelectorAll('.opaque-block.marker.move-marker').length).toBeGreaterThan(0);
    expect(root.querySelectorAll('.opaque-block.move-marker .opaque-label').length).toBe(0);
    // A mark's tooltip: the author and the date, no words of its own.
    const ins = root.querySelector('ins.insertion[data-author="Author A"]') as HTMLElement;
    expect(ins.title).toMatch(/^Author A, .*2026/);
    editor.destroy();
  });

  test('one rule per author, a name with a quote and a backslash included; the rule removed with the view', async () => {
    const { session, editor } = await open('revisions/revisions-word15.docx');
    const authors = new Set<string>();
    editor.state.doc.descendants((n) => {
      for (const m of n.marks) if (REVISION_MARKS.has(m.type.name)) authors.add(String(m.attrs['author']));
      for (const attr of ['markRevision', 'revision']) {
        const r = n.attrs[attr] as Revision | null | undefined;
        if (r?.insertion) authors.add(String(r.insertion.author));
        if (r?.deletion) authors.add(String(r.deletion.author));
      }
      return true;
    });
    expect(editor.authors.size).toBe(authors.size);
    // Typing with tracking on as an author whose name needs escaping.
    const name = 'O"Brien \\ Co';
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = name;
    let at = -1;
    editor.state.doc.forEach((node, pos) => { if (at < 0 && node.type.name === 'paragraph' && node.textContent.length > 5 && !holds(node)) at = pos + 2; });
    editor.caretAt(at);
    editor.view.dispatch(editor.state.tr.insertText('Z'));
    expect(editor.authors.size).toBe(authors.size + 1);
    const sheet = editor.authors.element.sheet as CSSStyleSheet;
    expect(sheet.cssRules.length).toBe(authors.size + 1);
    expect(editor.container.querySelector(`ins.insertion[data-author="O\\"Brien \\\\ Co"]`)).not.toBeNull();
    const element = editor.authors.element;
    editor.destroy();
    expect(element.isConnected).toBe(false);
  });

  test('a paragraph\'s own classes are toggled: a class a node decoration puts on it (a section\'s) survives an edit of it', async () => {
    const { editor } = await open('tracked-changes.docx');
    let at = -1;
    editor.state.doc.forEach((node, pos) => { if (at < 0 && node.type.name === 'paragraph' && node.textContent.length > 5) at = pos; });
    // A node decoration as decorations.ts makes a section's (`sect-N`), kept the same across the edit.
    const deco = (doc: typeof editor.state.doc): DecorationSet => DecorationSet.create(doc, [Decoration.node(at, at + doc.nodeAt(at)!.nodeSize, { class: 'sect-2' })]);
    editor.view.setProps({ decorations: (state) => deco(state.doc) });
    const dom = editor.view.nodeDOM(at) as HTMLElement;
    expect(dom.classList.contains('sect-2')).toBe(true);
    editor.view.dispatch(editor.state.tr.insertText('Q', at + 2));
    expect(editor.view.nodeDOM(at)).toBe(dom);
    expect(dom.classList.contains('sect-2')).toBe(true);
    editor.destroy();
  });
});

describe('the three views (ED-005 section 6.4)', () => {
  test('the root\'s class; Original not editable; All markup again', async () => {
    const { editor } = await open('tracked-changes.docx');
    const root = editor.view.dom;
    expect(editor.markupView).toBe('all');
    editor.markupView = 'accepted';
    expect(root.classList.contains('markup-accepted')).toBe(true);
    expect(editor.view.editable).toBe(true);
    editor.markupView = 'original';
    expect(root.classList.contains('markup-original')).toBe(true);
    expect(root.classList.contains('markup-accepted')).toBe(false);
    expect(editor.view.editable).toBe(false);
    expect(root.getAttribute('contenteditable')).toBe('false');
    editor.markupView = 'all';
    expect(root.classList.contains('markup-original')).toBe(false);
    expect(editor.view.editable).toBe(true);
    editor.destroy();
  });

  test('As accepted: a caret put inside hidden text moves to the run\'s start; at its end, too; in All markup it stays', async () => {
    const { editor } = await open('tracked-changes.docx');
    const h = junction(editor.state.doc);
    editor.caretAt(h.from + 3);
    expect(editor.state.selection.head).toBe(h.from + 3);
    editor.markupView = 'accepted';
    expect(editor.state.selection.head).toBe(h.from);
    editor.caretAt(h.to);
    expect(editor.state.selection.head).toBe(h.from);
    // A caret before the run, where the DOM point is the visible text's, stays.
    editor.caretAt(h.from - 1);
    expect(editor.state.selection.head).toBe(h.from - 1);
    // A peer's caret, a widget that takes the marks of what follows it, goes to the run's end, before visible text.
    expect(clampCaret(editor.state, h.from + 2)).toBe(h.to);
    expect(clampCaret(editor.state, h.from)).toBe(h.to);
    expect(clampCaret(editor.state, h.from - 1)).toBe(h.from - 1);
    editor.destroy();
  });

  test('As accepted: a paragraph that starts with hidden text has an anchor at the run\'s end, where its caret goes', async () => {
    const { editor } = await open('revisions/check22/22-deleted.docx');
    // Block 5 starts with deleted text ("Seven "), then visible text.
    let start = -1;
    let runEnd = -1;
    editor.state.doc.forEach((node, pos) => {
      if (start >= 0 || node.type.name !== 'paragraph') return;
      const first = node.firstChild;
      if (first?.isText && first.marks.some((m) => m.type.name === 'deletion') && node.childCount > 1) { start = pos + 1; runEnd = pos + 1 + first.nodeSize; }
    });
    expect(start).toBeGreaterThan(0);
    expect(editor.container.querySelectorAll('.rev-anchor').length).toBe(0);
    editor.markupView = 'accepted';
    expect(editor.container.querySelectorAll('.rev-anchor').length).toBeGreaterThan(0);
    expect(visibleCaret(editor.state.doc, start, 'accepted')).toBe(runEnd);
    editor.caretAt(start + 2);
    expect(editor.state.selection.head).toBe(runEnd);
    editor.markupView = 'all';
    expect(editor.container.querySelectorAll('.rev-anchor').length).toBe(0);
    editor.destroy();
  });

  test('As accepted, tracking on: Backspace after hidden text deletes the visible character before it; Delete before it the one after', async () => {
    const { session, editor } = await open('tracked-changes.docx');
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author C';
    editor.markupView = 'accepted';
    const key = (k: string): boolean => {
      const event = new KeyboardEvent('keydown', { key: k });
      return editor.view.someProp('handleKeyDown', (f) => f(editor.view, event)) ?? false;
    };
    const acceptedText = (pos: number): string => {
      const $p = editor.state.doc.resolve(pos);
      let s = '';
      $p.parent.forEach((n) => { if (n.isText && !n.marks.some((m) => m.type.name === 'deletion')) s += n.text; });
      return s;
    };
    const h = junction(editor.state.doc);
    const before = acceptedText(h.from);
    expect(before.endsWith('by.')).toBe(true);
    editor.caretAt(h.to);
    expect(key('Backspace')).toBe(true);
    expect(acceptedText(h.from)).toBe(before.slice(0, -3) + 'b.');
    // The caret before what it deleted, which is drawn text.
    expect(editor.state.selection.head).toBe(h.from - 1);
    // Delete from before the hidden runs: the "." after them is deleted, tracked.
    editor.caretAt(h.from - 1);
    expect(key('Delete')).toBe(true);
    expect(acceptedText(h.from)).toBe(before.slice(0, -3) + 'b');
    // Away from hidden text, the browser's own deletion stands: the handler declines.
    editor.caretAt(h.from - 10);
    expect(key('Backspace')).toBe(false);
    editor.destroy();
  });

  test('As accepted, tracking on: Backspace at a join crosses the hidden text and the deleted mark to the visible character before them; Delete the other way (check 22)', async () => {
    // "One." / "[Two.]¶" / "Three." / "Four.¶" / "Five [six.]¶" / "[Seven ]eight." / "Nine.", the marks marked ¶ deleted: As accepted reads "Four.Five eight." on one line.
    const { session, editor } = await open('revisions/check22/22-deleted.docx');
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author C';
    editor.markupView = 'accepted';
    const key = (k: string): boolean => editor.view.someProp('handleKeyDown', (f) => f(editor.view, new KeyboardEvent('keydown', { key: k }))) ?? false;
    const accepted = (): string[] => {
      const out: string[] = [];
      editor.state.doc.forEach((p) => { let s = ''; p.forEach((n) => { if (n.isText && !n.marks.some((m) => m.type.name === 'deletion')) s += n.text; }); out.push(s); });
      return out;
    };
    const start = (i: number): number => { let pos = 0; for (let k = 0; k < i; k++) pos += editor.state.doc.child(k).nodeSize; return pos + 1; };
    // The caret at "eight"'s start: after the hidden "Seven ", the anchor.
    editor.caretAt(start(5) + 'Seven '.length);
    expect(editor.state.selection.head).toBe(start(5) + 'Seven '.length);
    expect(key('Backspace')).toBe(true);
    expect(accepted()).toEqual(['One.', '', 'Three.', 'Four.', 'Five', 'eight.', 'Nine.', '']);
    // Delete at "Four."'s end: across its deleted mark to "Five"'s F.
    editor.caretAt(start(3) + 'Four.'.length);
    expect(key('Delete')).toBe(true);
    expect(accepted()).toEqual(['One.', '', 'Three.', 'Four.', 'ive', 'eight.', 'Nine.', '']);
    editor.destroy();
  });

  test('a navigation into hidden text switches to All markup; one to visible text does not', async () => {
    const { editor } = await open('tracked-changes.docx');
    let revealed = 0;
    (editor as unknown as { options: { onRevealed?: () => void } }).options.onRevealed = () => { revealed++; };
    const h = junction(editor.state.doc);
    editor.markupView = 'accepted';
    expect(editor.reveal(h.from - 5, h.from - 1)).toBe(false);
    expect(editor.markupView).toBe('accepted');
    expect(editor.reveal(h.from, h.to)).toBe(true);
    expect(editor.markupView).toBe('all');
    expect(revealed).toBe(1);
    editor.destroy();
  });
});

// UPSTREAM-15 (ED-005 section 12.40): y-prosemirror's diff took a text whose bold went as its formatChange came
// for unchanged, so the view showed the change and the store (the save, every peer) kept the bold.
describe('a formatting change through the view reaches the store (UPSTREAM-15)', () => {
  for (const [command, mark] of [['font.bold', 'bold'], ['font.italic', 'italic'], ['font.rtl', 'rtl']] as const) {
    test(`${command} taken off with tracking on`, async () => {
      const { session, editor } = await open('tracked-changes.docx');
      let from = -1;
      editor.state.doc.descendants((n, pos) => { if (from < 0 && n.isText && /\btext\b/.test(n.text!)) from = pos + n.text!.search(/\btext\b/); return from < 0; });
      const select = (): void => { editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, from + 4))); };
      select();
      expect(editor.run(command, { value: true })).toBe(true);
      session.trackingMode = 'TrackAll';
      session.trackingAuthor = 'Author C';
      // Two undo steps, as two commands a person gives are.
      session.undoManager.stopCapturing();
      select();
      expect(editor.run(command, { value: false })).toBe(true);
      for (const doc of [editor.state.doc, session.doc]) {
        const names = doc.nodeAt(from)!.marks.map((m) => m.type.name);
        expect(names).toContain('formatChange');
        expect(names).not.toContain(mark);
      }
      // One undo step takes the change back whole, in the store as on screen.
      session.undoManager.undo();
      expect(session.doc.nodeAt(from)!.marks.map((m) => m.type.name).sort()).toEqual([mark, 'runSource'].sort());
      editor.destroy();
    });
  }
});

// ED-005 section 12.41: a command through the view is an undo step of its own, as the headless runner makes it. Two
// Accepts in quick succession (within Yjs's 500 ms capture window) were one step, and Ctrl+Z took both back.
describe('a command through the view is an undo step of its own', () => {
  // The third review point (ED-005 section 12.47): keys keep Yjs's grouping, as before 12.41, so Enter among typed
  // text joins the typing's step; a command given from the surface does not.
  test('text typed with an Enter among it: one undo takes it all back', async () => {
    const { session, editor } = await open('tracked-changes.docx');
    const before = editor.state.doc.toJSON();
    let at = -1;
    editor.state.doc.descendants((n, pos) => { if (at < 0 && n.isText && n.text!.includes('Some text.')) at = pos + n.text!.indexOf('Some text.') + 4; return at < 0; });
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at)));
    session.undoManager.stopCapturing();
    editor.view.dispatch(editor.state.tr.insertText('ab'));
    const enter = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true });
    expect(editor.view.someProp('handleKeyDown', (f) => f(editor.view, enter))).toBe(true);
    editor.view.dispatch(editor.state.tr.insertText('c'));
    expect(editor.state.doc.toJSON()).not.toEqual(before);
    session.undoManager.undo();
    expect(editor.state.doc.toJSON()).toEqual(before);
    editor.destroy();
  });

  test('two accepts at once: one undo takes back the second only', async () => {
    const { session, editor } = await open('review/review-for-word.docx');
    const pending = (): number => trackedChangesOf(editor.state.doc).length;
    const [first, second] = trackedChangesOf(editor.state.doc);
    const before = pending();
    expect(editor.run('trackedChange.accept', { key: session.changeKey(first!, editor.state) })).toBe(true);
    const afterOne = pending();
    expect(editor.run('trackedChange.accept', { key: session.changeKey(trackedChangesOf(editor.state.doc).find((c) => c.text === second!.text && c.type === second!.type)!, editor.state) })).toBe(true);
    expect(pending()).toBeLessThan(afterOne);
    session.undoManager.undo();
    expect(pending()).toBe(afterOne);
    session.undoManager.undo();
    expect(pending()).toBe(before);
    editor.destroy();
  });
});
