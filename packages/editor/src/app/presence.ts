// SPDX-License-Identifier: GPL-3.0-or-later
//
// Who you are in a shared room (ED-004 section 5): the "Your name" option (`docx4j-editor.name`;
// empty: "Guest" and a number kept per browser), and a colour from the name with the comments
// margin's hue. E4's tracked changes will take the same name as the author of this user's edits.
import { authorHue } from '../view/authors.js';
import { peerColor, whereIs, type PeerUser } from '../view/presence.js';
import { t } from './i18n.js';
import type { EditStore } from '@docx4j/editor-model';
import type { Awareness } from 'y-protocols/awareness';

export const OPTION_NAME = 'docx4j-editor.name';
const GUEST = 'docx4j-editor.guest';

export function nameOption(): string {
  try { return (localStorage.getItem(OPTION_NAME) ?? '').trim(); } catch { return ''; }
}

export function setNameOption(value: string): void {
  try { localStorage.setItem(OPTION_NAME, value); } catch { /* ignore */ }
}

/** This browser's guest number, made once. */
function guestNumber(): number {
  try {
    const kept = Number(localStorage.getItem(GUEST));
    if (Number.isInteger(kept) && kept >= 10 && kept <= 99) return kept;
    const made = 10 + Math.floor(Math.random() * 90);
    localStorage.setItem(GUEST, String(made));
    return made;
  } catch {
    return 10 + Math.floor(Math.random() * 90);
  }
}

/** This user as the others see them. */
export function localUser(name: string = nameOption()): PeerUser {
  const shown = name.trim() !== '' ? name.trim() : t('presence.guest', { number: String(guestNumber()) });
  return { name: shown, color: peerColor(authorHue(shown)) };
}

/** Up to two initials of a name, whole characters (not UTF-16 halves). */
export function initials(name: string): string {
  const words = name.trim().split(/\s+/).filter((w) => w !== '');
  const first = (w: string): string => Array.from(w)[0] ?? '';
  const letters = words.length >= 2 ? first(words[0]!) + first(words[words.length - 1]!) : Array.from(words[0] ?? '?').slice(0, 2).join('');
  return letters.toLocaleUpperCase();
}

export interface Peer {
  clientId: number;
  name: string;
  color: string;
  /** The story the peer's caret is in (`whereIs`), or undefined when it is in none. */
  where: string | undefined;
  /** An AI agent (`@docx4j/editor-mcp` marks its awareness so): badged, and a person can stop it. */
  agent: boolean;
}

/** The other people in the room: every awareness state with a user but this peer's own, by client id. */
export function peersOf(awareness: Awareness, store: EditStore): Peer[] {
  const out: Peer[] = [];
  awareness.getStates().forEach((state, clientId) => {
    if (clientId === awareness.clientID) return;
    const user = state['user'] as (PeerUser & { agent?: boolean }) | undefined;
    if (!user || typeof user.name !== 'string') return;
    out.push({ clientId, name: user.name, color: user.color, where: whereIs(store, state['cursor']), agent: user.agent === true });
  });
  return out.sort((a, b) => a.clientId - b.clientId);
}

/**
 * Where a peer is, in words: "in the body", "in footnote 3", "in a header". A note is named by
 * the number its reference shows (`label`, the body's `noteLabels` by `kind:id`), else its id.
 */
export function whereText(where: string | undefined, label: (key: string) => string | undefined = () => undefined): string {
  if (!where) return t('presence.where.none');
  if (where === 'body') return t('presence.where.body');
  const colon = where.indexOf(':');
  const kind = where.slice(0, colon);
  const rest = where.slice(colon + 1);
  if (kind === 'footnote') return t('presence.where.footnote', { id: label(where) ?? rest });
  if (kind === 'endnote') return t('presence.where.endnote', { id: label(where) ?? rest });
  if (kind === 'footer') return t('presence.where.footer');
  return t('presence.where.header');
}
