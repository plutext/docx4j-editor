// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 8.9, step F4: an agent in the room as the page shows it. The People list badges a
// peer whose awareness marks it an agent and gives it a Stop, which writes the Stop into the room
// (`agents`), where the agent's companion hears it; the review panel's log lists the agent's calls
// newest first, a refused one marked, and a click goes to the paragraph a call acted on. Mounted
// under StrictMode (ED-005 section 12.14).
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { act, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { Awareness } from 'y-protocols/awareness';
import { EditorDocument, appendAgentLog, isAgentStopped, type AgentLogEntry } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { PeopleList } from '../src/app/PeopleList.js';
import { AgentLog } from '../src/app/panels/AgentLog.js';
import { peersOf } from '../src/app/presence.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
const zeroRect = (): DOMRect => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON: () => ({}) }) as DOMRect;
if (typeof Range !== 'undefined') {
  Range.prototype.getClientRects ??= () => ([] as unknown as DOMRectList);
  Range.prototype.getBoundingClientRect ??= zeroRect;
}
if (typeof Element !== 'undefined') Element.prototype.getClientRects ??= () => ([] as unknown as DOMRectList);

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

async function shared(): Promise<{ session: EditorDocument; awareness: Awareness }> {
  const bytes = new Uint8Array(await readFile(join(fixtures, 'tracked-changes.docx')));
  const session = await EditorDocument.open(bytes);
  await session.share(bytes, 'lease.docx');
  return { session, awareness: new Awareness(session.store.ydoc) };
}

/** A peer's state as the provider would have received it. */
function remote(awareness: Awareness, clientId: number, state: Record<string, unknown>): void {
  awareness.states.set(clientId, state);
  awareness.meta.set(clientId, { clock: 1, lastUpdated: Date.now() });
  awareness.emit('change', [{ added: [clientId], updated: [], removed: [] }, 'remote']);
}

describe('an agent in the room, as the page shows it (ED-005 section 8.9, step F4)', () => {
  test('People: an agent badged, with a Stop that writes the Stop into the room; a person neither', async () => {
    const { session, awareness } = await shared();
    remote(awareness, 101, { user: { name: 'Claude', color: '#6d28d9', agent: true } });
    remote(awareness, 102, { user: { name: 'Ada Lovelace', color: '#0a7' } });
    expect(peersOf(awareness, session.store).map((p) => [p.name, p.agent])).toEqual([['Claude', true], ['Ada Lovelace', false]]);
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    const status: string[] = [];
    await act(async () => { root.render(<StrictMode><PeopleList awareness={awareness} store={session.store} onStatus={(s) => status.push(s)} /></StrictMode>); });
    expect(host.querySelectorAll('[data-testid="agent-badge"]').length).toBe(1);
    expect(host.querySelector('[data-testid="person"][data-agent="true"]')!.textContent).toBe('CL');
    const stop = host.querySelector<HTMLButtonElement>('[data-testid="agent-stop"]')!;
    expect(stop.title).toContain('Claude');
    await act(async () => { stop.click(); });
    expect(isAgentStopped(session.store, 'Claude')).toBe(true);
    expect(isAgentStopped(session.store, 'Ada Lovelace')).toBe(false);
    expect(status.at(-1)).toBe('Claude was stopped and has left the document');
    // A Stop is not an edit: nothing for the person's Ctrl+Z to take back.
    expect(session.undoManager.undoStack.length).toBe(0);
    root.unmount();
  });

  test('the log: nothing without calls; the calls newest first, a refused one marked; a click goes to the paragraph a call acted on', async () => {
    const { session } = await shared();
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    await act(async () => { root.render(<StrictMode><AgentLog editor={editor} session={session} onStatus={() => undefined} lang="en" /></StrictMode>); });
    expect(host.querySelector('[data-testid="agent-log"]')).toBeNull();
    const target = (await session.readTree(async ({ body }) => (await body.outline()).paragraphs.find((p) => p.text === 'Some text.')))!;
    const entry = (over: Partial<AgentLogEntry>): AgentLogEntry => ({ time: '2026-09-29T10:00:00Z', agent: 'Claude', tool: 'insert', summary: 'Inserted a paragraph', addresses: [], paraIds: [], ok: true, ...over });
    await act(async () => {
      appendAgentLog(session.store, entry({ tool: 'set_format', summary: 'Set bold on "text"', addresses: [target.address] }));
      appendAgentLog(session.store, entry({ tool: 'edit', summary: 'target.stale: the paragraph changed since it was read', ok: false }));
    });
    const items = [...host.querySelectorAll<HTMLElement>('[data-testid="agent-log-entry"]')];
    expect(items.map((i) => [i.dataset['tool'], i.dataset['ok']])).toEqual([['edit', 'false'], ['set_format', 'true']]);
    expect(items[0]!.textContent).toContain('Refused');
    expect(items[0]!.querySelector('button')!.disabled).toBe(true);
    await act(async () => { items[1]!.querySelector('button')!.click(); await new Promise((r) => setTimeout(r, 20)); });
    expect(editor.state.selection.$from.parent.textContent).toBe('Some text.');
    root.unmount();
    editor.destroy();
  });
});
