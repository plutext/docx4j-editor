// SPDX-License-Identifier: GPL-3.0-or-later
//
// A shared room in the application (ED-004 sections 4.3 and 4.4): the link that names it, and
// the connection that keeps the edit store in step with the relay. The provider binds the store's
// own `Y.Doc`, so the edit layer and the provider share one document and nothing is copied; its
// transactions carry the provider as their origin, which the session's undo manager does not
// track (it tracks the local origin alone), so a colleague's or an agent's change is never undone
// by this user's Ctrl+Z.
//
// The link is `<site>#room=<id>&token=<token>&relay=<address>`: the fragment, which a browser
// never sends to the site's server. The room id is 128 random bits, the token 256, both
// base64url; the relay's address is absolute (`ws:` or `wss:`).

/** The relay's address (an option; empty: no Share). Defaults to the deployment's `VITE_RELAY_URL`. */
export const OPTION_RELAY = 'docx4j-editor.relay';

export function relayOption(): string {
  try {
    const stored = localStorage.getItem(OPTION_RELAY);
    if (stored !== null) return stored.trim();
  } catch { /* no storage */ }
  const env = (import.meta as { env?: Record<string, string | undefined> }).env;
  return (env?.['VITE_RELAY_URL'] ?? '').trim();
}

export function setRelayOption(value: string): void {
  try { localStorage.setItem(OPTION_RELAY, value); } catch { /* ignore */ }
}

export interface RoomLink {
  room: string;
  token: string;
  relay: string;
}

function randomBase64Url(bytes: number): string {
  const buffer = new Uint8Array(bytes);
  globalThis.crypto.getRandomValues(buffer);
  let binary = '';
  for (const b of buffer) binary += String.fromCharCode(b);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

/** A new room on a relay: a fresh id and token. */
export function newRoomLink(relay: string): RoomLink {
  return { room: randomBase64Url(16), token: randomBase64Url(32), relay };
}

/**
 * The relay option as a WebSocket address: `ws:` and `wss:` as they are, `http:` and `https:`
 * turned into them, and a path (`/relay`, what the dev and preview servers mount) against the
 * page's own origin. Empty for an empty option.
 */
export function relayAddress(option: string, page: string = globalThis.location?.href ?? 'http://localhost/'): string {
  const value = option.trim();
  if (value === '') return '';
  const url = new URL(value, page);
  if (url.protocol === 'http:') url.protocol = 'ws:';
  else if (url.protocol === 'https:') url.protocol = 'wss:';
  if (url.protocol !== 'ws:' && url.protocol !== 'wss:') throw new Error(`not a relay address: ${option}`);
  return url.toString().replace(/\/$/, '');
}

/** The link to a room, on the page `site` (the application's own address, without its fragment). */
export function roomLinkUrl(site: string, link: RoomLink): string {
  const base = site.replace(/#.*$/, '');
  const fragment = new URLSearchParams({ room: link.room, token: link.token, relay: link.relay });
  return `${base}#${fragment.toString()}`;
}

/** The room a link names (`host-protocol.ts`, where the extension shares it, ED-014 section 4.5). */
export { parseRoomLink } from './host-protocol.js';

/** Whether a relay is on this computer (`127.0.0.1`, `localhost`, `[::1]`): a companion's, in VS Code. */
export function isLoopbackRelay(relay: string): boolean {
  try { const { hostname } = new URL(relay); return hostname === '127.0.0.1' || hostname === 'localhost' || hostname === '[::1]'; } catch { return false; }
}

/** The relay's host, for the join question and the Share dialog ("through relay.example.org"). */
export function relayHost(relay: string): string {
  try { return new URL(relay).host; } catch { return relay; }
}

// The connection to a room lives beside the relay (`@docx4j/editor-relay/client`), where the MCP
// companion's sessions use it too (ED-005 section 9.3 item 1 (i)).
export { RoomConnection, RoomRefusedError, type RoomStatus, type RoomConnectionOptions } from '@docx4j/editor-relay/client';
