// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 8.2 (section 11.16 item 4): the comments margin over a document with comments,
// its ranges from one pass, and its anchors following the document rather than the selection.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { act, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { TextSelection } from 'prosemirror-state';
import { EditorDocument } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { CommentsMargin, commentRanges, commentRange } from '../src/app/panels/CommentsMargin.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

describe('the comments margin', () => {
  test('mounts over a document with comments (a hook after an early return once crashed it), and finds every range in one memoised pass', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'comments-two.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    await act(async () => { root.render(<CommentsMargin editor={editor} session={session} />); });
    await act(async () => { await new Promise((r) => setTimeout(r, 50)); });
    expect(host.querySelectorAll('[data-testid="comment-bubble"]').length).toBeGreaterThan(0);
    const ranges = commentRanges(editor.state.doc);
    expect(ranges.size).toBeGreaterThan(0);
    for (const [id, r] of ranges) {
      expect(r.from).toBeLessThan(r.to);
      expect(commentRange(editor, id)).toEqual(r);
    }
    // The same document object gives the same map; a selection change makes no new document.
    expect(commentRanges(editor.state.doc)).toBe(ranges);
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1)));
    expect(commentRanges(editor.state.doc)).toBe(ranges);
    // An edit outside the commented text (at the document's end) keeps every range's text: the ranges are re-read over the new blocks.
    const before = editor.state.doc;
    const end = before.content.size - 1;
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(before, end)));
    editor.run('insertText', { text: 'x' });
    const after = commentRanges(editor.state.doc);
    expect(after.size).toBe(ranges.size);
    for (const [id, r] of after) expect(editor.state.doc.textBetween(r.from, r.to)).toBe(before.textBetween(ranges.get(id)!.from, ranges.get(id)!.to));
    root.unmount();
  });
});

describe('the comments margin edits comments (ED-005 section 7.5)', () => {
  const zeroRect = (): DOMRect => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON: () => ({}) }) as DOMRect;
  if (typeof Range !== 'undefined') {
    Range.prototype.getClientRects ??= () => ([] as unknown as DOMRectList);
    Range.prototype.getBoundingClientRect ??= zeroRect;
  }
  if (typeof Element !== 'undefined') Element.prototype.getClientRects ??= () => ([] as unknown as DOMRectList);

  const settle = async (): Promise<void> => { await act(async () => { await new Promise((r) => setTimeout(r, 30)); }); };

  test("under React's StrictMode, a comment made after the margin mounts is listed, one's own editable and another's not; Resolve, Reopen and Delete through its buttons, each undone", async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'comments-two.docx'))));
    session.trackingAuthor = 'Reviewer One';
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    await act(async () => { root.render(<StrictMode><CommentsMargin editor={editor} session={session} me="Reviewer One" /></StrictMode>); });
    await settle();
    const bubbles = (): HTMLElement[] => [...host.querySelectorAll<HTMLElement>('[data-testid="comment-bubble"]')];
    expect(bubbles().map((b) => b.dataset['commentId'])).toEqual(['0', '1']);
    // Made after the margin mounted: listed, and the author's own body is editable; the file's are not.
    editor.caretAt(3, 6);
    await act(async () => { expect(await editor.runAsync('range.insertComment', { text: 'mine' })).toBe(true); });
    await settle();
    expect(bubbles().map((b) => b.dataset['commentId']).sort()).toEqual(['0', '1', '2']);
    const bubble = (id: string): HTMLElement => bubbles().find((b) => b.dataset['commentId'] === id)!;
    expect(bubble('2').querySelector('[data-testid="comment-body"]')!.getAttribute('data-editable')).toBe('true');
    expect(bubble('0').querySelector('[data-testid="comment-body"]')!.getAttribute('data-editable')).toBeNull();
    // Resolve, then Reopen, then Delete; undo brings each back.
    await act(async () => { bubble('0').querySelector<HTMLButtonElement>('[data-testid="comment-resolve"]')!.click(); });
    await settle();
    expect(bubble('0').classList.contains('done')).toBe(true);
    await act(async () => { bubble('0').querySelector<HTMLButtonElement>('[data-testid="comment-reopen"]')!.click(); });
    await settle();
    expect(bubble('0').classList.contains('done')).toBe(false);
    await act(async () => { bubble('1').querySelector<HTMLButtonElement>('[data-testid="comment-delete"]')!.click(); });
    await settle();
    expect(bubbles().map((b) => b.dataset['commentId']).sort()).toEqual(['0', '2']);
    await act(async () => { editor.run('history.undo'); });
    await settle();
    expect(bubbles().map((b) => b.dataset['commentId']).sort()).toEqual(['0', '1', '2']);
    root.unmount();
  });

  test('Reply posts a reply under the thread, which its author edits', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'comments-two.docx'))));
    session.trackingAuthor = 'Reviewer Two';
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    await act(async () => { root.render(<StrictMode><CommentsMargin editor={editor} session={session} me="Reviewer Two" /></StrictMode>); });
    await settle();
    const bubble = host.querySelector<HTMLElement>('[data-comment-id="0"]')!;
    await act(async () => { bubble.querySelector<HTMLButtonElement>('[data-testid="comment-reply"]')!.click(); });
    const box = bubble.querySelector<HTMLTextAreaElement>('[data-testid="comment-reply-text"]')!;
    await act(async () => {
      const set = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!;
      set.call(box, 'a reply');
      box.dispatchEvent(new Event('input', { bubbles: true }));
    });
    await act(async () => { bubble.querySelector<HTMLButtonElement>('[data-testid="comment-post"]')!.click(); });
    await settle();
    const reply = host.querySelector<HTMLElement>('[data-comment-id="0"] .comment-reply');
    expect(reply).not.toBeNull();
    expect(reply!.textContent).toContain('Reviewer Two');
    expect(reply!.textContent).toContain('a reply');
    expect(reply!.querySelector('[data-testid="comment-body"]')!.getAttribute('data-editable')).toBe('true');
    root.unmount();
  });
});
