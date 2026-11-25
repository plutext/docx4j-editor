// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-013 S3 (sections 3.6 and 3.7, P5 and P6): the compact layout's sheets, mounted under React's StrictMode (the dev
// server's, CLAUDE.md). The changes' sheet is the review panel: any "all" asks first, and an entry, Next or Previous
// closes it with the selection in the change. The thread sheet draws the threads at the caret with the margin's
// `Thread`: bodies read-only, Delete on one's own alone, a reply written in the sheet's one field.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { act, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { EditorDocument, trackedChangesOf } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { ReviewSheet, ThreadSheet } from '../src/app/CompactSheets.js';
import { threadsAt, useCommentThreads } from '../src/app/comments.js';
import { reviewEntries, entryAtSelection } from '../src/app/review.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

// jsdom lays nothing out: ProseMirror's scroll-into-view reads rectangles it has not got (as in review-panel.test.tsx).
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

const settle = async (): Promise<void> => { await act(async () => { await new Promise((r) => setTimeout(r, 30)); }); };

function mountHost(): { host: HTMLElement; render: (node: React.ReactNode) => Promise<void>; unmount: () => void } {
  const host = document.createElement('div');
  document.body.appendChild(host);
  const root = createRoot(host);
  return { host, render: async (node) => { await act(async () => { root.render(<StrictMode>{node}</StrictMode>); }); await settle(); }, unmount: () => act(() => root.unmount()) };
}

const textChanges = (editor: DocxEditorView): number => trackedChangesOf(editor.state.doc).length;

/** The changes' sheet once its panel, a chunk of its own (lazy), has come and listed its entries. */
async function reviewSheet(): Promise<Element> {
  for (let i = 0; i < 200; i++) {
    const sheet = document.querySelector('[data-testid="review-sheet"]');
    if (sheet?.querySelector('[data-testid="review-entry"]')) return sheet;
    await settle();
  }
  throw new Error('the review panel never came');
}

describe("the changes' sheet (P5)", () => {
  test('Accept all asks first: Cancel keeps every change, the confirmation resolves them all', async () => {
    const { session, editor } = await open('revisions/revisions-word15.docx');
    const { render, unmount } = mountHost();
    const status: string[] = [];
    let closed = 0;
    await render(<ReviewSheet editor={editor} session={session} onStatus={(s) => status.push(s)} onTrackChanges={() => undefined} onClose={() => { closed++; }} />);
    const sheet = await reviewSheet();
    expect(sheet.querySelector('[data-testid="panel-review"]')).not.toBeNull();
    const before = textChanges(editor);
    expect(before).toBeGreaterThan(1);
    const click = async (id: string): Promise<void> => { await act(async () => { sheet.querySelector<HTMLButtonElement>(`[data-testid="${id}"]`)!.click(); }); };
    await click('review-accept-all');
    expect(sheet.querySelector('[data-testid="review-confirm"]')!.textContent).toMatch(/^Accept all \d+ changes\?/);
    expect(textChanges(editor)).toBe(before);
    await click('review-confirm-no');
    expect(sheet.querySelector('[data-testid="review-confirm"]')).toBeNull();
    expect(textChanges(editor)).toBe(before);
    await click('review-reject-all');
    expect(sheet.querySelector('[data-testid="review-confirm"]')!.textContent).toMatch(/^Reject all/);
    await click('review-confirm-yes');
    await settle();
    expect(textChanges(editor)).toBe(0);
    expect(status.at(-1)).toMatch(/^Rejected \d+ changes$/);
    expect(closed).toBe(0);
    unmount();
    editor.destroy();
  });

  test("an author's all asks in that author's group, and resolves that author's alone", async () => {
    const { session, editor } = await open('revisions/revisions-word15.docx');
    const { render, unmount } = mountHost();
    await render(<ReviewSheet editor={editor} session={session} onStatus={() => undefined} onTrackChanges={() => undefined} onClose={() => undefined} />);
    const sheet = await reviewSheet();
    await act(async () => { sheet.querySelector<HTMLInputElement>('[data-testid="review-by-author"]')!.click(); });
    const group = sheet.querySelector('[data-testid="review-author"][data-author="Author B"]')!;
    await act(async () => { group.querySelector<HTMLButtonElement>('[data-testid="review-author-accept"]')!.click(); });
    expect(sheet.querySelectorAll('[data-testid="review-confirm"]')).toHaveLength(1);
    expect(group.querySelector('[data-testid="review-confirm"]')!.textContent).toMatch(/by Author B\?/);
    expect(trackedChangesOf(editor.state.doc).some((c) => c.author === 'Author B')).toBe(true);
    await act(async () => { group.querySelector<HTMLButtonElement>('[data-testid="review-confirm-yes"]')!.click(); });
    await settle();
    expect(trackedChangesOf(editor.state.doc).some((c) => c.author === 'Author B')).toBe(false);
    expect(trackedChangesOf(editor.state.doc).some((c) => c.author === 'Author A')).toBe(true);
    unmount();
    editor.destroy();
  });

  test('an entry, Next and Previous put the selection in a change and close the sheet; an entry\'s own Accept does not', async () => {
    const { session, editor } = await open('revisions/revisions-word15.docx');
    const { render, unmount } = mountHost();
    let closed = 0;
    await render(<ReviewSheet editor={editor} session={session} onStatus={() => undefined} onTrackChanges={() => undefined} onClose={() => { closed++; }} />);
    const sheet = await reviewSheet();
    const entries = reviewEntries(editor.state, session);
    const second = sheet.querySelectorAll<HTMLElement>('[data-testid="review-entry"]')[1]!;
    await act(async () => { second.querySelector<HTMLButtonElement>('.review-entry-main')!.click(); });
    expect(closed).toBe(1);
    expect(entryAtSelection(entries, editor.state)?.id).toBe(entries[1]!.id);
    await act(async () => { sheet.querySelector<HTMLButtonElement>('[data-testid="review-next"]')!.click(); });
    expect(closed).toBe(2);
    expect(entryAtSelection(entries, editor.state)?.id).toBe(entries[2]!.id);
    await act(async () => { sheet.querySelector<HTMLButtonElement>('[data-testid="review-previous"]')!.click(); });
    expect(closed).toBe(3);
    const accept = sheet.querySelector<HTMLButtonElement>('[data-testid="review-entry"][data-kind="change"]:not([data-type="Formatted"]) [data-testid="review-accept"]')!;
    await act(async () => { accept.click(); });
    expect(closed).toBe(3);
    unmount();
    editor.destroy();
  });
});

describe('the thread sheet (P6)', () => {
  /** comments-two.docx: "One paragraph" with comment 0 on "paragraph" (positions 5 to 14), "blsgh " with comment 1. */
  test('threadsAt: the threads whose text is beside a caret or inside a selection, a reply counted as its thread', async () => {
    const { session, editor } = await open('comments-two.docx');
    const threads = await session.stories.comments(editor.state.doc);
    expect(threads.map((t) => t.id)).toEqual(['0', '1']);
    editor.caretAt(8);
    expect(threadsAt(editor.state, session, threads).map((t) => t.id)).toEqual(['0']);
    editor.caretAt(14);
    expect(threadsAt(editor.state, session, threads).map((t) => t.id)).toEqual(['0']);
    editor.caretAt(2);
    expect(threadsAt(editor.state, session, threads)).toEqual([]);
    editor.caretAt(2, editor.state.doc.content.size - 1);
    expect(threadsAt(editor.state, session, threads).map((t) => t.id)).toEqual(['0', '1']);
    session.trackingAuthor = 'Reviewer One';
    expect(await editor.runAsync('comment.reply', { id: '0', text: 'a reply' })).toBe(true);
    const after = await session.stories.comments(editor.state.doc);
    expect(after[0]!.replies).toHaveLength(1);
    editor.caretAt(8);
    expect(threadsAt(editor.state, session, after).map((t) => t.id)).toEqual(['0']);
    editor.destroy();
  });

  function Threads({ editor, session, ids, me, onClose }: { editor: DocxEditorView; session: EditorDocument; ids: string[]; me: string; onClose: () => void }) {
    const threads = useCommentThreads(editor, session);
    return <ThreadSheet editor={editor} session={session} threads={threads} ids={ids} me={me} onStatus={() => undefined} onClose={onClose}
      onReply={(id, text) => { session.trackingAuthor = me; void editor.runAsync('comment.reply', { id, text }); }} />;
  }

  test("the thread at the caret, read-only; a reply from the field, Delete on one's own alone; Resolve marks it; deleting the last closes the sheet", async () => {
    const { session, editor } = await open('comments-two.docx');
    const { render, unmount } = mountHost();
    let closed = 0;
    await render(<Threads editor={editor} session={session} ids={['0']} me="Reviewer One" onClose={() => { closed++; }} />);
    const sheet = document.querySelector<HTMLElement>('[data-testid="thread-sheet"]')!;
    const threads = (): HTMLElement[] => [...sheet.querySelectorAll<HTMLElement>('[data-testid="sheet-thread"]')];
    expect(threads().map((t) => t.dataset['commentId'])).toEqual(['0']);
    expect(threads()[0]!.textContent).toContain('jharrop2');
    expect(threads()[0]!.textContent).toContain('One comment');
    // Read-only, and not this page's: no Delete.
    expect(sheet.querySelector('[data-editable="true"]')).toBeNull();
    expect(sheet.querySelector('[data-testid="comment-delete"]')).toBeNull();
    const field = sheet.querySelector<HTMLTextAreaElement>('[data-testid="thread-sheet-text"]')!;
    expect(field.getAttribute('aria-label')).toBe('Reply to jharrop2');
    const post = sheet.querySelector<HTMLButtonElement>('[data-testid="thread-sheet-post"]')!;
    expect(post.disabled).toBe(true);
    // Reply points the field at the thread.
    await act(async () => { sheet.querySelector<HTMLButtonElement>('[data-testid="comment-reply"]')!.click(); });
    expect(document.activeElement).toBe(field);
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(field, 'from the sheet');
      field.dispatchEvent(new Event('input', { bubbles: true }));
    });
    await act(async () => { post.click(); });
    await settle();
    expect(field.value).toBe('');
    const reply = threads()[0]!.querySelector<HTMLElement>('.comment-reply')!;
    expect(reply.textContent).toContain('Reviewer One');
    expect(reply.textContent).toContain('from the sheet');
    // The reply is this page's: it has a Delete, still read-only.
    expect(reply.querySelectorAll('[data-testid="comment-delete"]')).toHaveLength(1);
    expect(sheet.querySelectorAll('[data-testid="comment-delete"]')).toHaveLength(1);
    expect(sheet.querySelector('[data-editable="true"]')).toBeNull();
    // Resolve: marked, its field gone (a resolved thread takes no reply), Reopen offered.
    await act(async () => { sheet.querySelector<HTMLButtonElement>('[data-testid="comment-resolve"]')!.click(); });
    await settle();
    expect(sheet.querySelector('[data-testid="sheet-thread-resolved"]')).not.toBeNull();
    expect(sheet.querySelector('[data-testid="thread-sheet-text"]')).toBeNull();
    expect(sheet.querySelector('[data-testid="comment-reopen"]')).not.toBeNull();
    // Delete the reply: the thread stays, the sheet open.
    await act(async () => { sheet.querySelector<HTMLButtonElement>('[data-testid="comment-delete"]')!.click(); });
    await settle();
    expect(threads()[0]!.querySelector('.comment-reply')).toBeNull();
    expect(closed).toBe(0);
    // The thread deleted another way (a peer, the palette): nothing is left to show, so the sheet closes.
    await act(async () => { await editor.runAsync('comment.delete', { id: '0' }); });
    await settle();
    expect(closed).toBeGreaterThan(0);
    unmount();
    editor.destroy();
  });
});
