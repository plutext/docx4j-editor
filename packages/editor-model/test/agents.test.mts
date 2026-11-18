// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.28, E4.f step F4: an agent's Stop and log in the room (`agents.mts`). A Stop is
// heard, cleared by an invitation, and reaches a peer; the log keeps its last entries; neither is an
// edit the person's undo takes back.
import { describe, test, expect } from 'vitest';
import * as Y from 'yjs';
import { EditorDocument, newEditStore, stopAgent, resumeAgent, isAgentStopped, appendAgentLog, agentLog, onAgents, AGENT_LOG_CAP, runCommand, type AgentLogEntry } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

const entry = (n: number): AgentLogEntry => ({ time: `2026-09-29T10:00:${String(n % 60).padStart(2, '0')}Z`, agent: 'Claude', tool: 'insert', summary: `call ${n}`, addresses: ['body/1'], paraIds: [], ok: true });

describe('an agent in the room (ED-005 section 12.28)', () => {
  test('a Stop by name, heard, cleared by an invitation, and reaching a peer', async () => {
    const store = newEditStore();
    let heard = 0;
    const off = onAgents(store, () => { heard++; });
    expect(isAgentStopped(store, 'Claude')).toBe(false);
    stopAgent(store, 'Claude');
    expect(isAgentStopped(store, 'Claude')).toBe(true);
    expect(isAgentStopped(store, 'Other')).toBe(false);
    const peer = newEditStore();
    Y.applyUpdate(peer.ydoc, Y.encodeStateAsUpdate(store.ydoc));
    expect(isAgentStopped(peer, 'Claude')).toBe(true);
    resumeAgent(store, 'Claude');
    expect(isAgentStopped(store, 'Claude')).toBe(false);
    expect(heard).toBe(2);
    off();
  });

  test('the log keeps its last entries, oldest first', () => {
    const store = newEditStore();
    for (let n = 0; n < AGENT_LOG_CAP + 7; n++) appendAgentLog(store, entry(n));
    const log = agentLog(store);
    expect(log.length).toBe(AGENT_LOG_CAP);
    expect(log[0]!.summary).toBe('call 7');
    expect(log.at(-1)!.summary).toBe(`call ${AGENT_LOG_CAP + 6}`);
  });

  test('neither a Stop nor a log entry is an edit the person\'s undo takes back', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    runCommand(session, 'insertText', { text: 'Typed ' }, { from: 1 });
    stopAgent(session.store, 'Claude');
    appendAgentLog(session.store, entry(1));
    expect(runCommand(session, 'history.undo').ok).toBe(true);
    expect(session.doc.textContent).not.toContain('Typed ');
    expect(isAgentStopped(session.store, 'Claude')).toBe(true);
    expect(agentLog(session.store).length).toBe(1);
  });
});
