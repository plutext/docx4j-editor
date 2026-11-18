// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-004 section 5 in jsdom: a peer's caret drawn in the surface whose story it is in and no
// other, installed when the session joins a room and removed when it leaves; where a peer is; the
// colour, the initials and the guest name.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import * as Y from 'yjs';
import { TextSelection } from 'prosemirror-state';
import { Awareness } from 'y-protocols/awareness';
import { EditorDocument } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { EditableStoryView } from '../src/view/StoryView.js';
import { setPresence, whereIs, peerColor } from '../src/view/presence.js';
import { initials, localUser, peersOf, whereText } from '../src/app/presence.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

/** A caret at `index` of the first text of `fragment`'s block `block`, as the JSON a peer's awareness carries. */
function caretIn(fragment: Y.XmlFragment, block: number, index = 1): { anchor: unknown; head: unknown } {
  let text: Y.XmlText | undefined;
  const walk = (node: Y.XmlElement | Y.XmlText | Y.XmlHook): void => {
    if (text) return;
    if (node instanceof Y.XmlText) { text = node; return; }
    if (node instanceof Y.XmlElement) for (const child of node.toArray()) walk(child as Y.XmlElement | Y.XmlText);
  };
  walk(fragment.get(block) as Y.XmlElement);
  if (!text) throw new Error('no text in that block');
  const position = Y.relativePositionToJSON(Y.createRelativePositionFromTypeIndex(text, index));
  return { anchor: position, head: position };
}

/** A peer's state as the provider would have received it; the cursor plugin redraws on the next tick (y-prosemirror's `setMeta` batches). */
async function remote(awareness: Awareness, clientId: number, state: Record<string, unknown>): Promise<void> {
  awareness.states.set(clientId, state);
  awareness.meta.set(clientId, { clock: 1, lastUpdated: Date.now() });
  awareness.emit('change', [{ added: [clientId], updated: [], removed: [] }, 'remote']);
  await new Promise((r) => setTimeout(r, 5));
}

/** The presence plugin redraws on the next tick. */
const tick = (): Promise<void> => new Promise((r) => setTimeout(r, 5));

async function shared(): Promise<{ session: EditorDocument; body: DocxEditorView; awareness: Awareness }> {
  const bytes = new Uint8Array(await readFile(join(fixtures, 'footnote.docx')));
  const session = await EditorDocument.open(bytes);
  await session.share(bytes, 'footnote.docx');
  const container = document.createElement('div');
  document.body.appendChild(container);
  const body = new DocxEditorView(container, session);
  return { session, body, awareness: new Awareness(session.store.ydoc) };
}

describe('presence (ED-004 section 5)', () => {
  test('a peer\'s caret in the body is drawn in the body, in its colour, once the session is in a room, and goes when it leaves', async () => {
    const { session, body, awareness } = await shared();
    await remote(awareness, 999, { user: { name: 'Ada Lovelace', color: '#aa3355' }, cursor: caretIn(session.store.body, 0) });
    expect(body.container.querySelector('.peer-caret')).toBeNull();   // not in a room yet
    setPresence(session, awareness);
    await tick();
    const caret = body.container.querySelector<HTMLElement>('.peer-caret[data-peer="999"]');
    expect(caret).not.toBeNull();
    expect(caret!.textContent).toContain('Ada Lovelace');
    expect(caret!.style.borderColor).not.toBe('');
    // The caret follows the peer.
    await remote(awareness, 999, { user: { name: 'Ada Lovelace', color: '#aa3355' }, cursor: caretIn(session.store.body, 1) });
    expect(body.container.querySelectorAll('.peer-caret').length).toBe(1);
    setPresence(session, null);
    await tick();
    expect(body.container.querySelector('.peer-caret')).toBeNull();
    body.destroy();
  });

  test('a caret in a footnote is drawn in that footnote\'s view and not in the body; where it is reads as the footnote', async () => {
    const { session, body, awareness } = await shared();
    setPresence(session, awareness);
    const note = (await session.stories.footnotes())[0]!;
    const fragment = session.stories.fragmentOf(note);
    const container = document.createElement('div');
    document.body.appendChild(container);
    const story = new EditableStoryView(container, session, note, { stylesheet: body });
    await remote(awareness, 7, { user: { name: 'Grace', color: '#3355aa' }, cursor: caretIn(fragment, 0) });
    expect(container.querySelector('.peer-caret[data-peer="7"]')).not.toBeNull();
    expect(body.container.querySelector('.peer-caret[data-peer="7"]')).toBeNull();
    expect(whereIs(session.store, caretIn(fragment, 0))).toBe(`footnote:${note.id}`);
    expect(whereIs(session.store, caretIn(session.store.body, 0))).toBe('body');
    expect(whereIs(session.store, null)).toBeUndefined();
    const peers = peersOf(awareness, session.store);
    expect(peers.map((p) => [p.name, p.where])).toEqual([['Grace', `footnote:${note.id}`]]);
    expect(whereText(peers[0]!.where)).toBe(`in footnote ${note.id}`);
    story.destroy();
    setPresence(session, null);
    body.destroy();
  });

  test('this peer\'s own caret is written to the one cursor field and not drawn for itself', async () => {
    const { session, body, awareness } = await shared();
    setPresence(session, awareness);
    body.focus();
    // Not `caretAt`, which scrolls (jsdom has no layout).
    body.view.dispatch(body.state.tr.setSelection(TextSelection.create(body.state.doc, 3)));
    const own = awareness.getLocalState()?.['cursor'];
    expect(own).toBeTruthy();
    expect(whereIs(session.store, own)).toBe('body');
    expect(body.container.querySelector('.peer-caret')).toBeNull();
    setPresence(session, null);
    body.destroy();
  });

  test('joining and leaving a room leave undo working: the plugin list never changes (ED-004 section 10.8 item 1)', async () => {
    const { session, body, awareness } = await shared();
    const plugins = body.state.plugins.length;
    setPresence(session, awareness);
    expect(body.state.plugins.length).toBe(plugins);
    body.view.dispatch(body.state.tr.insertText('Mine. ', 2));
    session.undoManager.stopCapturing();
    expect(session.undoManager.canUndo()).toBe(true);
    expect(body.run('history.undo')).toBe(true);
    expect(body.state.doc.textContent).not.toContain('Mine. ');
    setPresence(session, null);
    body.view.dispatch(body.state.tr.insertText('Again. ', 2));
    expect(session.undoManager.canUndo()).toBe(true);
    setPresence(session, awareness);
    expect(body.run('history.undo')).toBe(true);
    expect(body.state.doc.textContent).not.toContain('Again. ');
    setPresence(session, null);
    body.destroy();
  });

  test('colours, initials, the guest name', () => {
    expect(peerColor(0)).toMatch(/^#[0-9a-f]{6}$/);
    expect(peerColor(200)).toMatch(/^#[0-9a-f]{6}$/);
    expect(peerColor(0)).not.toBe(peerColor(200));
    expect(initials('Ada Lovelace')).toBe('AL');
    expect(initials('grace brewster hopper')).toBe('GH');
    expect(initials('Jason')).toBe('JA');
    expect(initials('李小龍')).toBe('李小');
    expect(initials('𝒜da')).toBe('𝒜D');
    expect(localUser('').name).toMatch(/^Guest \d\d$/);
    expect(localUser('').name).toBe(localUser('').name);
    expect(localUser('Ada').color).toMatch(/^#[0-9a-f]{6}$/);
  });
});
