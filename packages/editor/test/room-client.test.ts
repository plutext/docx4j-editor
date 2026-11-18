// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-004 section 4.6's client tests: the application's room connection over a relay in the same
// process. A creator shares, a joiner opens from the synced store, typing converges both ways, a
// colleague's change is not undone by this user's undo, and a refusal or a full room stops the
// connection instead of retrying. The link is parsed and made as the application does. In the
// node project (`vitest.config.ts`): nothing here needs a DOM, and a save is reloaded.
// @vitest-environment node
import { afterEach, describe, expect, test } from 'vitest';
import { readFile, mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import WebSocket from 'ws';
import { EditorDocument, newEditStore, runCommand } from '@docx4j/editor-model';
import { ZipPartStore } from '@docx4j/core-ts';
import { createRelay } from '@docx4j/editor-relay';
import { RoomConnection, newRoomLink, parseRoomLink, relayAddress, relayHost, roomLinkUrl, type RoomStatus } from '../src/app/room.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');
const cleanups: (() => Promise<void> | void)[] = [];
afterEach(async () => { for (const c of cleanups.splice(0).reverse()) await c(); });

async function relay(options: { maxRoom?: number } = {}): Promise<string> {
  const data = await mkdtemp(join(tmpdir(), 'room-client-'));
  cleanups.push(() => rm(data, { recursive: true, force: true }));
  const r = createRelay({ port: 0, debounce: 10, data, ...options });
  const url = await r.listen();
  cleanups.push(() => r.destroy());
  return url;
}

function connect(store: EditorDocument['store'], link: ReturnType<typeof newRoomLink>, statuses: RoomStatus[] = [], create = false): RoomConnection {
  const c = new RoomConnection(store, link, { WebSocketPolyfill: WebSocket, onStatus: (s) => statuses.push(s), create });
  cleanups.push(() => c.close());
  return c;
}

/** The parts whose bytes differ, `[Content_Types].xml` and the `.rels` parts left out (the engine regenerates them). */
function differingParts(x: Uint8Array, y: Uint8Array): string[] {
  const a = new ZipPartStore(x);
  const b = new ZipPartStore(y);
  const an = new Set(a.partNames());
  const bn = new Set(b.partNames());
  const out: string[] = [];
  for (const n of new Set([...an, ...bn])) {
    if (n === '[Content_Types].xml' || n.endsWith('.rels')) continue;
    const p = an.has(n) ? a.loadSync(n) : undefined;
    const q = bn.has(n) ? b.loadSync(n) : undefined;
    if (!p || !q || p.length !== q.length || p.some((v, i) => v !== q[i])) out.push(n);
  }
  return out;
}

async function until(check: () => boolean, ms = 5000): Promise<void> {
  const end = Date.now() + ms;
  while (!check()) {
    if (Date.now() > end) throw new Error('timed out');
    await new Promise((r) => setTimeout(r, 10));
  }
}

/** Insert text at the start of block `i`, as a keystroke's command does. */
function type(s: EditorDocument, i: number, text: string): void {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += s.doc.child(k).nodeSize;
  expect(runCommand(s, 'insertText', { text }, { from: pos + 1 }).ok).toBe(true);
}

async function share(): Promise<{ creator: EditorDocument; joiner: EditorDocument; link: ReturnType<typeof newRoomLink>; base: Uint8Array }> {
  const url = await relay();
  const base = new Uint8Array(await readFile(join(fixtures, 'tracked-changes.docx')));
  const creator = await EditorDocument.open(base);
  const link = newRoomLink(url);
  // Share: connect over the open store (claiming the room), then write the room map.
  await connect(creator.store, link, [], true).synced;
  await creator.share(base, 'tracked-changes.docx');
  // Join: a fresh store, synced first, then opened from the base file the room carries.
  const store = newEditStore();
  await connect(store, link).synced;
  await until(() => store.room.get('base') !== undefined);
  const joiner = await EditorDocument.open(store.room.get('base') as Uint8Array, { store, mode: 'join' });
  return { creator, joiner, link, base };
}

describe('the room link (ED-004 section 4.4)', () => {
  test('a link made for a room parses back; what is not one does not', () => {
    const link = newRoomLink('ws://127.0.0.1:4175');
    expect(link.room).toMatch(/^[A-Za-z0-9_-]{22}$/);
    expect(link.token).toMatch(/^[A-Za-z0-9_-]{43}$/);
    const url = roomLinkUrl('https://editor.example/app/#old', link);
    expect(url.startsWith('https://editor.example/app/#room=')).toBe(true);
    expect(parseRoomLink(url)).toEqual(link);
    expect(parseRoomLink('https://editor.example/#room=short&token=x&relay=ws://a')).toBeUndefined();
    expect(parseRoomLink(`#room=${link.room}&token=x&relay=http://a`)).toBeUndefined();
    expect(parseRoomLink('https://editor.example/')).toBeUndefined();
    expect(relayHost('wss://relay.example.org/rooms')).toBe('relay.example.org');
  });

  test('the relay option becomes a WebSocket address', () => {
    expect(relayAddress('/relay', 'https://editor.example/app/')).toBe('wss://editor.example/relay');
    expect(relayAddress('http://127.0.0.1:4175', 'https://x/')).toBe('ws://127.0.0.1:4175');
    expect(relayAddress('wss://relay.example.org/', 'https://x/')).toBe('wss://relay.example.org');
    expect(relayAddress('  ', 'https://x/')).toBe('');
    expect(() => relayAddress('ftp://a', 'https://x/')).toThrow();
  });
});

describe('the room connection (ED-004 section 4.3)', () => {
  test('a creator shares and a joiner opens the same document; typing converges both ways; the saves agree', async () => {
    const { creator, joiner } = await share();
    expect(joiner.doc.eq(creator.doc)).toBe(true);
    type(creator, 1, 'Creator. ');
    type(joiner, 3, 'Joiner. ');
    await until(() => creator.doc.eq(joiner.doc) && creator.doc.child(3).textContent.startsWith('Joiner. ') && joiner.doc.child(1).textContent.startsWith('Creator. '));
    const [a, b] = [await creator.save(), await joiner.save()];
    expect(a.bytes.length).toBeGreaterThan(0);
    expect(differingParts(a.bytes, b.bytes)).toEqual([]);
  });

  test('a colleague\'s change is not undone by this user\'s undo', async () => {
    const { creator, joiner } = await share();
    type(creator, 1, 'Mine. ');
    creator.undoManager.stopCapturing();
    type(joiner, 3, 'Theirs. ');
    await until(() => creator.doc.child(3).textContent.startsWith('Theirs. '));
    creator.undoManager.undo();
    expect(creator.doc.child(1).textContent.startsWith('Mine. ')).toBe(false);
    expect(creator.doc.child(3).textContent.startsWith('Theirs. ')).toBe(true);
    // Nothing left to undo: the colleague's change was never on this user's stack.
    expect(creator.undoManager.canUndo()).toBe(false);
    await until(() => joiner.doc.child(1).textContent.startsWith('Mine. ') === false);
  });

  test('a wrong token is refused and not retried; the joiner opens nothing', async () => {
    const { link } = await share();
    const statuses: RoomStatus[] = [];
    const store = newEditStore();
    const c = connect(store, { ...link, token: newRoomLink(link.relay).token }, statuses);
    await expect(c.synced).rejects.toMatchObject({ status: 'refused' });
    expect(c.status).toBe('refused');
    await new Promise((r) => setTimeout(r, 1500));   // longer than the provider's reconnect delay
    expect(statuses.filter((s) => s === 'connecting').length).toBeLessThanOrEqual(1);
    expect(store.room.get('base')).toBeUndefined();
  });

  test('a room at the relay\'s limit stops the connection as full', async () => {
    const url = await relay({ maxRoom: 64 * 1024 });
    const base = new Uint8Array(await readFile(join(fixtures, 'tracked-changes.docx')));
    const creator = await EditorDocument.open(base);
    const statuses: RoomStatus[] = [];
    const c = connect(creator.store, newRoomLink(url), statuses, true);
    await c.synced;
    // Sharing writes the base file into the room: past 64 kB for this fixture with its store.
    await creator.share(new Uint8Array(200 * 1024), 'too big.docx');
    await until(() => c.status === 'full');
    expect(statuses).toContain('full');
  });

  test('a join to a room the relay does not hold is said as unknown, not retried, and makes no room', async () => {
    const url = await relay();
    const statuses: RoomStatus[] = [];
    const link = newRoomLink(url);
    const c = connect(newEditStore(), link, statuses);
    await expect(c.synced).rejects.toMatchObject({ status: 'unknown' });
    expect(c.status).toBe('unknown');
    // The joiner made nothing: a sharer can still make the room with its own token.
    const store = newEditStore();
    await connect(store, { ...link, token: newRoomLink(url).token }, [], true).synced;
  });

  test('closing leaves the room for good: no reconnect', async () => {
    const { link } = await share();
    const statuses: RoomStatus[] = [];
    const c = connect(newEditStore(), link, statuses);
    await c.synced;
    c.close();
    await new Promise((r) => setTimeout(r, 1500));
    expect(statuses[statuses.length - 1]).toBe('closed');
    expect(c.status).toBe('closed');
  });
});
