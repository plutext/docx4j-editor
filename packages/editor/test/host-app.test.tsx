// SPDX-License-Identifier: GPL-3.0-or-later
//
// The application as VS Code's page (ED-014 section 4.1), mounted whole under React's StrictMode (CLAUDE.md: a store a
// panel subscribes to, and here the host's attachment, must survive the replayed mount) over a fake `acquireVsCodeApi`:
// `ready` once, an open, the chrome VS Code takes over gone, typing posting `changed`, a backup taken and resumed dirty,
// and a backup an earlier version of the editor kept opening the bytes alone, not dirty, said. A save is the browser
// suite's (`e2e/host-vscode.spec.ts`): under jsdom a saved package cannot be reloaded, so the self-check refuses it.
import { afterEach, describe, expect, test } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { StrictMode, act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import * as Y from 'yjs';
import { EditorDocument, newEditStore, PROJECTION_KEY } from '@docx4j/editor-model';
import { App } from '../src/app/App.js';
import { initHost, resetHost } from '../src/app/host.js';
import type { FromPage } from '../src/app/host-protocol.js';
import { t } from '../src/app/i18n.js';
import type { DocxEditorView } from '../src/view/DocxEditorView.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');
const fixture = async (name: string): Promise<Uint8Array> => new Uint8Array(await readFile(join(fixtures, name)));

let root: Root | undefined;
let container: HTMLDivElement | undefined;

afterEach(async () => {
  await act(async () => { root?.unmount(); });
  container?.remove();
  root = undefined;
  resetHost();
  localStorage.clear();
  delete (window as { docx4jEditor?: unknown }).docx4jEditor;
});

/** The application mounted under StrictMode as VS Code's page; what it posts is recorded. */
async function mount(): Promise<FromPage[]> {
  const posted: FromPage[] = [];
  initHost({ postMessage: (m) => { posted.push(m as FromPage); } });
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  await act(async () => { root!.render(<StrictMode><App /></StrictMode>); });
  return posted;
}

const send = async (data: unknown): Promise<void> => { await act(async () => { window.dispatchEvent(new MessageEvent('message', { data })); }); };

/** Until `check` holds, acting between tries so effects and state settle. */
async function until(check: () => boolean, what: string, ms = 20_000): Promise<void> {
  const end = Date.now() + ms;
  while (!check()) {
    if (Date.now() > end) throw new Error(`timed out waiting for ${what}`);
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 20)); });
  }
}

const editor = (): DocxEditorView | undefined => (window as { docx4jEditor?: DocxEditorView }).docx4jEditor;
const testId = (id: string): Element | null => container!.querySelector(`[data-testid="${id}"]`);
const nameShown = (): string => testId('name')?.textContent ?? '';

describe('the application as VS Code\'s page', () => {
  test('ready once under StrictMode; the page waits, opens, hides VS Code\'s chrome, and posts changed as it is typed in', async () => {
    const posted = await mount();
    expect(posted.filter((m) => m.type === 'ready')).toEqual([{ type: 'ready', origin: location.origin }]);
    expect(testId('opening')?.textContent).toBe(t('app.opening'));
    expect(testId('open-sample')).toBeNull();
    await send({ type: 'open', name: 'hello.docx', bytes: await fixture('HelloWordOnline.docx') });
    await until(() => editor() !== undefined && nameShown() === 'hello.docx', 'the document');
    for (const id of ['new', 'open', 'save', 'save-as', 'download', 'print', 'join', 'close', 'option-layout', 'option-save-in-place']) expect(testId(id), id).toBeNull();
    expect(posted.some((m) => m.type === 'changed')).toBe(false);
    const view = editor()!;
    await act(async () => { view.view.dispatch(view.state.tr.insertText('Typed ', 1)); });
    await until(() => posted.some((m) => m.type === 'changed'), 'changed');
    expect(nameShown()).toBe('hello.docx •');
  });

  test('a backup taken is resumed in a new page, the edit there and the document dirty', async () => {
    let posted = await mount();
    const bytes = await fixture('HelloWordOnline.docx');
    await send({ type: 'open', name: 'hello.docx', bytes });
    await until(() => editor() !== undefined, 'the document');
    const view = editor()!;
    await act(async () => { view.view.dispatch(view.state.tr.insertText('Resumed ', 1)); });
    await send({ type: 'backup', id: 1 });
    const backup = posted.find((m): m is Extract<FromPage, { type: 'backup'; update: Uint8Array }> => m.type === 'backup' && 'update' in m);
    expect(backup?.meta).toMatchObject({ name: 'hello.docx' });
    await act(async () => { root!.unmount(); });
    container!.remove();
    resetHost();
    delete (window as { docx4jEditor?: unknown }).docx4jEditor;
    posted = await mount();
    await send({ type: 'open', name: 'hello.docx', bytes, resume: { update: backup!.update, meta: backup!.meta } });
    await until(() => editor() !== undefined && nameShown() !== '', 'the resumed document');
    expect(editor()!.state.doc.textContent).toContain('Resumed ');
    expect(nameShown()).toBe('hello.docx •');
    expect(testId('status')?.textContent).toContain(t('status.recovered'));
  });

  test('a backup an earlier version of the editor kept opens the bytes alone, not dirty, and says so to the page and the host', async () => {
    // A store made by projection 2 over a document with revisions, which this projection reads otherwise (ED-005 section 12.3).
    const bytes = await fixture('revisions/revisions-word15.docx');
    const store = newEditStore(new Y.Doc());
    await EditorDocument.open(bytes, { store });
    store.ydoc.transact(() => store.package.set(PROJECTION_KEY, 2));
    const update = Y.encodeStateAsUpdate(store.ydoc);
    const posted = await mount();
    await send({ type: 'open', name: 'revisions.docx', bytes, resume: { update, meta: { name: 'revisions.docx' } } });
    await until(() => editor() !== undefined && nameShown() !== '', 'the document');
    const said = t('status.resume.versionRefused', { name: 'revisions.docx' });
    expect(nameShown()).toBe('revisions.docx');
    expect(testId('status')?.textContent).toContain(said);
    expect(posted).toContainEqual({ type: 'status', text: said });
  });

  test('a backup whose room does not answer opens out of the room with its edits, said; a backup asked while it waits keeps them', async () => {
    let posted = await mount();
    const bytes = await fixture('HelloWordOnline.docx');
    await send({ type: 'open', name: 'hello.docx', bytes });
    await until(() => editor() !== undefined, 'the document');
    const view = editor()!;
    await act(async () => { view.view.dispatch(view.state.tr.insertText('Kept ', 1)); });
    await send({ type: 'backup', id: 1 });
    const backup = posted.find((m): m is Extract<FromPage, { type: 'backup'; update: Uint8Array }> => m.type === 'backup' && 'update' in m)!;
    await act(async () => { root!.unmount(); });
    container!.remove();
    resetHost();
    delete (window as { docx4jEditor?: unknown }).docx4jEditor;
    posted = await mount();
    // A relay on a port nothing listens on: the companion's relay after a restart, at a new port (ED-014 section 3.6).
    const room = { room: 'R'.repeat(22), token: 'T'.repeat(43), relay: 'ws://127.0.0.1:9' };
    await send({ type: 'open', name: 'hello.docx', bytes, resume: { update: backup.update, meta: { ...backup.meta, room } } });
    // VS Code asks for a backup while the page waits for the relay: the backup being resumed is the answer.
    await send({ type: 'backup', id: 2 });
    const meanwhile = posted.find((m): m is Extract<FromPage, { type: 'backup'; update: Uint8Array }> => m.type === 'backup' && m.id === 2 && 'update' in m);
    expect(meanwhile?.update).toEqual(backup.update);
    await until(() => editor() !== undefined && nameShown() !== '', 'the document out of the room', 15_000);
    const note = t('status.resume.roomUnreachable');
    expect(editor()!.state.doc.textContent).toContain('Kept ');
    expect(nameShown()).toBe('hello.docx •');
    expect(testId('status')?.textContent).toContain(note);
    expect(testId('room-status')).toBeNull();
    expect(posted).toContainEqual({ type: 'status', text: note });
    // Now open, the document is what a backup holds.
    await send({ type: 'backup', id: 3 });
    const after = posted.find((m): m is Extract<FromPage, { type: 'backup'; meta: { room?: unknown } }> => m.type === 'backup' && m.id === 3 && 'meta' in m);
    expect(after?.meta.room).toBeUndefined();
  }, 30_000);

  test('an open that fails says so on the waiting page and to the host', async () => {
    const posted = await mount();
    await send({ type: 'open', name: 'broken.docx', bytes: new Uint8Array([1, 2, 3, 4]) });
    await until(() => testId('opening')?.textContent === t('app.openFailed'), 'the failure');
    expect(posted.some((m) => m.type === 'status' && m.text.includes('broken.docx'))).toBe(true);
  });
});
