// SPDX-License-Identifier: GPL-3.0-or-later
//
// The companion process (ED-001 section 11 item 4; ED-003 section 6.5): optional, local, and the
// only route to Re-paginate, since the static site never pretends to lay a page out (section 9
// item 6). The contract the editor speaks, which E4's `@docx4j/editor-mcp` serves over
// docx4j-mcp's `paginate` tool and the test double serves for the suites:
//
//   GET  <companion>/           -> { "tools": [...], "relay"?, "agent"? }   the probe; "paginate" among the tools shows the button
//   POST <companion>/paginate   -> the docx bytes in, the docx bytes out, markers rewritten
//   POST <companion>/rooms      -> { room, token } in: the agent invited into a room on the companion's relay (E4.f step F4)
//
// The address is an option; empty, a page the companion itself serves finds it at its own
// `/companion` (`@docx4j/editor-mcp`, ED-005 section 12.28): the companion marks the index.html it
// serves with `<meta name="docx4j-companion" content="/companion">`, and a page without the mark (a
// static site, the demo on GitHub Pages) probes nothing, where a probe would be a 404 on every open
// (release plan D8, 2026-11-18). Counted only when its probe names an
// agent (the dev and preview servers mount Re-paginate's test double there). It is probed when a
// document opens and when the option changes. A probe that names an `agent` and a `relay` offers
// "Invite an agent".

export const OPTION_COMPANION = 'docx4j-editor.companion';

export const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

export function companionUrl(): string {
  try { return (localStorage.getItem(OPTION_COMPANION) ?? '').trim(); } catch { return ''; }
}

/** The mark a page the companion serves carries: `@docx4j/editor-mcp`'s server writes it into index.html (`server.mts`). */
export const COMPANION_META = 'docx4j-companion';

/** Where the page's own companion answers, from the mark the companion put in the page; '' on a page it did not serve. */
export function servedCompanion(doc: Document | undefined = globalThis.document): string {
  try { return doc?.querySelector(`meta[name="${COMPANION_META}"]`)?.getAttribute('content')?.trim() ?? ''; } catch { return ''; }
}

/** The companion to probe: the option, else the one that served this page (its mark resolved against the page), else none. */
export function effectiveCompanion(option: string, page: string = globalThis.location?.href ?? '', served: string = servedCompanion()): string {
  if (option.trim() !== '') return option.trim();
  if (served === '') return '';
  try {
    const url = new URL(page);
    if (url.protocol !== 'http:' && url.protocol !== 'https:') return '';
    return new URL(served, url).href;
  } catch {
    return '';
  }
}

/** What a companion's probe answered. */
export interface CompanionInfo {
  tools: string[];
  /** Its relay (a path on the companion, or an address), where it takes invitations to rooms. */
  relay?: string;
  /** The name of its agent, when it has one (`@docx4j/editor-mcp`). */
  agent?: string;
}

/** The companion at `url`, or null when nothing answers (no companion, not running, refused, not JSON). */
export async function probeCompanionInfo(url: string): Promise<CompanionInfo | null> {
  const base = url.trim().replace(/\/+$/, '');
  if (base === '') return null;
  try {
    const res = await fetch(`${base}/`, { method: 'GET', headers: { Accept: 'application/json' } });
    if (!res.ok) return null;
    const json = (await res.json()) as { tools?: unknown; relay?: unknown; agent?: unknown };
    const tools = Array.isArray(json.tools) ? json.tools.filter((t): t is string => typeof t === 'string') : [];
    return { tools, ...(typeof json.relay === 'string' ? { relay: json.relay } : {}), ...(typeof json.agent === 'string' ? { agent: json.agent } : {}) };
  } catch {
    return null;
  }
}

/** The companion's relay as a WebSocket address: its probe's `relay` against the companion's own address. */
export function companionRelay(url: string, info: CompanionInfo): string | undefined {
  if (!info.relay) return undefined;
  try {
    const address = new URL(info.relay, `${url.trim().replace(/\/+$/, '')}/`);
    if (address.protocol === 'http:') address.protocol = 'ws:';
    else if (address.protocol === 'https:') address.protocol = 'wss:';
    return address.toString().replace(/\/$/, '');
  } catch {
    return undefined;
  }
}

/** The agent invited into a room on the companion's relay: `POST <companion>/rooms`. */
export async function inviteAgent(url: string, room: { room: string; token: string }): Promise<void> {
  const base = url.trim().replace(/\/+$/, '');
  const res = await fetch(`${base}/rooms`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ room: room.room, token: room.token }) });
  if (!res.ok) {
    let reason = `${res.status} ${res.statusText}`;
    try { const json = (await res.json()) as { error?: unknown }; if (typeof json.error === 'string') reason = json.error; } catch { /* keep the status */ }
    throw new Error(reason);
  }
}

/** The tools the companion at `url` offers, or null when nothing answers (no companion, not running, refused). */
export async function probeCompanion(url: string): Promise<string[] | null> {
  return (await probeCompanionInfo(url))?.tools ?? null;
}

/** The document paginated by the companion: the bytes of a save in, the bytes with the markers rewritten out. */
export async function paginateWith(url: string, bytes: Uint8Array): Promise<Uint8Array> {
  const base = url.trim().replace(/\/+$/, '');
  const res = await fetch(`${base}/paginate`, { method: 'POST', headers: { 'Content-Type': DOCX_TYPE }, body: bytes as BodyInit });
  if (!res.ok) throw new Error(`the companion answered ${res.status} ${res.statusText}`);
  return new Uint8Array(await res.arrayBuffer());
}
