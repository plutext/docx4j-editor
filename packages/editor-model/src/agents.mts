// SPDX-License-Identifier: GPL-3.0-or-later
//
// An agent in a room (ED-005 sections 8.2 and 8.6, E4.f step F4): whether a person stopped it, and
// the log of what it did. Both live in the room, so every page sees them and the agent's companion,
// being in the room, sees a Stop at once whichever relay the room is on. Neither is undone by a
// person's Ctrl+Z (their writes carry `AGENTS_ORIGIN`, which the undo manager does not track), and
// neither changes what a room's document is, so an older peer ignores them rather than misreading
// them (no room schema change).
import type { EditStore } from './yjs.mjs';

/** The origin of an agent's log entries and a Stop: not an edit, never undone. */
export const AGENTS_ORIGIN = 'agents';

/** The log keeps this many entries, the oldest dropped first. */
export const AGENT_LOG_CAP = 500;

/** One call an agent made on the room (section 8.6). */
export interface AgentLogEntry {
  /** When, as an ISO date. */
  time: string;
  /** The agent's name. */
  agent: string;
  /** The tool, by the contract's name. */
  tool: string;
  /** The change report's first line, a reading's gist, or the refusal. */
  summary: string;
  /** Where it acted (the blocks touched, made or removed), ordinals as of the call. */
  addresses: string[];
  /** The same blocks' paragraph ids where they have them, which the person's edits do not shift. */
  paraIds: string[];
  /** False for a refused call: the summary is its error. */
  ok: boolean;
}

/** An agent's entry in the `agents` map. */
interface AgentState {
  stopped?: boolean;
}

/** Whether a person stopped the agent of this name in the room. */
export function isAgentStopped(store: EditStore, name: string): boolean {
  return (store.agents.get(name) as AgentState | undefined)?.stopped === true;
}

/** A person stops the agent (the People list's Stop): its companion leaves the room and refuses its calls there. */
export function stopAgent(store: EditStore, name: string): void {
  store.ydoc.transact(() => store.agents.set(name, { stopped: true } satisfies AgentState), AGENTS_ORIGIN);
}

/** The agent invited again: the Stop cleared. */
export function resumeAgent(store: EditStore, name: string): void {
  if (!store.agents.has(name)) return;
  store.ydoc.transact(() => store.agents.delete(name), AGENTS_ORIGIN);
}

/** Every call an agent made on the room, oldest first. */
export function agentLog(store: EditStore): AgentLogEntry[] {
  return store.agentLog.toArray() as AgentLogEntry[];
}

/** A call appended to the log, the oldest entries dropped past the cap. */
export function appendAgentLog(store: EditStore, entry: AgentLogEntry): void {
  store.ydoc.transact(() => {
    store.agentLog.push([entry]);
    const over = store.agentLog.length - AGENT_LOG_CAP;
    if (over > 0) store.agentLog.delete(0, over);
  }, AGENTS_ORIGIN);
}

/** Told of every change to the log or to the agents' Stop flags. */
export function onAgents(store: EditStore, listener: () => void): () => void {
  const heard = (): void => listener();
  store.agentLog.observe(heard);
  store.agents.observe(heard);
  return () => { store.agentLog.unobserve(heard); store.agents.unobserve(heard); };
}

