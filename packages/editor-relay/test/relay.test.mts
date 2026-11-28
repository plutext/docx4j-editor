// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-004 section 4.6's relay tests, in Node: a room claimed by its first token and refused to
// another; two providers converging; an in-process peer's direct connection reaching a provider,
// which applies it under its own origin (the agent's path, E4); a read-only connection's update
// refused; persistence across a restart; the size limit; the idle rule; origins.
import { afterEach, describe, expect, test } from 'vitest';
import { mkdtemp, readdir, rm, stat, symlink } from 'node:fs/promises';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { randomBytes } from 'node:crypto';
import * as Y from 'yjs';
import WebSocket from 'ws';
import { HocuspocusProvider, HocuspocusProviderWebsocket } from '@hocuspocus/provider';
import { createRelay, type Relay, type RelayOptions } from '../src/index.mjs';

const cleanups: (() => Promise<void>)[] = [];
afterEach(async () => { for (const c of cleanups.splice(0).reverse()) await c(); });

async function start(options: RelayOptions = {}): Promise<{ relay: Relay; url: string; data: string }> {
  const data = options.data ?? await mkdtemp(join(tmpdir(), 'relay-'));
  const relay = createRelay({ port: 0, debounce: 10, ...options, data });
  const url = await relay.listen();
  // Cleanups run in reverse: the relay stops before its folder goes.
  if (!options.data) cleanups.push(() => rm(data, { recursive: true, force: true }));
  cleanups.push(async () => { await relay.destroy(); });
  return { relay, url, data };
}

const roomId = (): string => randomBytes(16).toString('base64url');
const token = (): string => randomBytes(32).toString('base64url');

interface Peer { doc: Y.Doc; provider: HocuspocusProvider; failed: Promise<string>; synced: Promise<void>; closed: Promise<{ code: number; reason: string }>; close(): void }

function peer(url: string, name: string, secret: string, headers?: Record<string, string>): Peer {
  const doc = new Y.Doc();
  // Node 20 has no global WebSocket; `ws` is the provider's polyfill (with the headers a browser would send).
  const Polyfill = headers ? class extends WebSocket { constructor(address: string) { super(address, { headers }); } } : WebSocket;
  // A short delay: the provider's `onClose` schedules a reconnect on a timer its `destroy` does not
  // cancel, so a refused peer's last attempt must land while its relay still runs (see `close`).
  const socket = new HocuspocusProviderWebsocket({ url, WebSocketPolyfill: Polyfill, maxAttempts: 1 });
  let failed!: (reason: string) => void;
  let synced!: () => void;
  let closed!: (e: { code: number; reason: string }) => void;
  const out: Peer = {
    closed: new Promise((resolve) => { closed = resolve; }),
    doc,
    provider: undefined as never,
    close: () => undefined,
    failed: new Promise((resolve) => { failed = resolve; }),
    synced: new Promise((resolve) => { synced = resolve; }),
  };
  out.provider = new HocuspocusProvider({
    websocketProvider: socket, name, document: doc, token: secret,
    onAuthenticationFailed: ({ reason }) => failed(reason),
    onSynced: () => synced(),
    onClose: ({ event }) => closed({ code: event.code, reason: event.reason }),
  });
  out.provider.attach();
  // The provider and its socket both: a socket left alone goes on reconnecting.
  out.close = () => { out.provider.destroy(); closeSocket(socket); };
  cleanups.push(async () => out.close());
  return out;
}

/** Destroy a socket for good. */
function closeSocket(socket: HocuspocusProviderWebsocket): void {
  socket.destroy();
  // UPSTREAM-6: `destroy` leaves `onClose`'s scheduled reconnect in place, and `connect` revives the socket.
  socket.connect = async () => undefined;
}

/** Wait until `check` holds (a remote update is asynchronous). */
async function until(check: () => boolean, ms = 3000): Promise<void> {
  const end = Date.now() + ms;
  while (!check()) {
    if (Date.now() > end) throw new Error('timed out');
    await new Promise((r) => setTimeout(r, 10));
  }
}

const text = (doc: Y.Doc): string => doc.getText('t').toString();

describe('the relay (ED-004 section 4.2)', () => {
  test('a room is claimed by its first token and refused to another; two providers with the token converge', async () => {
    const { url, data } = await start();
    const room = roomId();
    const secret = token();
    const a = peer(url, room, `create:${secret}`);
    await a.synced;
    const b = peer(url, room, secret);
    await b.synced;
    a.doc.getText('t').insert(0, 'from a ');
    b.doc.getText('t').insert(0, 'from b ');
    await until(() => text(a.doc) === text(b.doc) && text(a.doc).length === 14);
    const intruder = peer(url, room, token());
    expect(await intruder.failed).toBe('Forbidden');
    expect(text(intruder.doc)).toBe('');
    // Only the token's hash is kept.
    const files = await readdir(data);
    expect(files).toContain(`${room}.json`);
    const { readFile } = await import('node:fs/promises');
    expect(await readFile(join(data, `${room}.json`), 'utf8')).not.toContain(secret);
  });

  test('a join to a room the relay does not hold is refused as unknown, and makes nothing; only create: makes a room', async () => {
    const { relay, url, data } = await start();
    const room = roomId();
    const secret = token();
    expect(await peer(url, room, secret).failed).toBe('Unknown Room');
    expect(await peer(url, room, `readonly:${secret}`).failed).toBe('Unknown Room');
    await expect(relay.openRoom(room, secret)).rejects.toThrow(/no such room/);
    expect((await readdir(data).catch(() => [])).filter((f) => f.startsWith(room))).toEqual([]);
    // The maker, then a plain join with the same token, then the in-process peer.
    await peer(url, room, `create:${secret}`).synced;
    await peer(url, room, secret).synced;
    const agent = await relay.openRoom(room, secret);
    await agent.disconnect();
    // create: on a room that exists with another token is refused like any wrong token.
    expect(await peer(url, room, `create:${token()}`).failed).toBe('Forbidden');
    const made = await relay.openRoom(roomId(), token(), { create: true });
    await made.disconnect();
  });

  test('a room id that is not one, and an empty token, are refused', async () => {
    const { url } = await start();
    expect(await peer(url, '../../etc/passwd', token()).failed).toBeTruthy();
    expect(await peer(url, roomId(), 'create:').failed).toBeTruthy();
  });

  test('an in-process peer\'s direct connection reaches a provider, which applies it under its own origin', async () => {
    const { relay, url } = await start();
    const room = roomId();
    const secret = token();
    const browser = peer(url, room, `create:${secret}`);
    await browser.synced;
    const origins: unknown[] = [];
    browser.doc.on('update', (_u: Uint8Array, origin: unknown) => { origins.push(origin); });
    const agent = await relay.openRoom(room, secret);
    await agent.transact((doc) => { doc.getText('t').insert(0, 'written by the agent'); });
    await until(() => text(browser.doc) === 'written by the agent');
    // Not the agent's origin: the browser's undo manager, tracking local origins only, never undoes it.
    expect(origins.every((o) => o === browser.provider)).toBe(true);
    await agent.disconnect();
    await expect(relay.openRoom(room, token())).rejects.toThrow(/wrong token/);
  });

  test('a read-only connection receives the room and its updates are refused', async () => {
    const { url } = await start();
    const room = roomId();
    const secret = token();
    const writer = peer(url, room, `create:${secret}`);
    await writer.synced;
    writer.doc.getText('t').insert(0, 'shared');
    const viewer = peer(url, room, `readonly:${secret}`);
    await viewer.synced;
    await until(() => text(viewer.doc) === 'shared');
    viewer.doc.getText('t').insert(0, 'not allowed ');
    writer.doc.getText('t').insert(6, '.');
    await until(() => text(viewer.doc).endsWith('shared.'));
    await new Promise((r) => setTimeout(r, 100));
    expect(text(writer.doc)).toBe('shared.');
  });

  test('a room persists across a restart of the relay', async () => {
    const data = await mkdtemp(join(tmpdir(), 'relay-'));
    cleanups.push(() => rm(data, { recursive: true, force: true }));
    const room = roomId();
    const secret = token();
    const first = createRelay({ port: 0, debounce: 10, data });
    const url1 = await first.listen();
    const a = peer(url1, room, `create:${secret}`);
    await a.synced;
    a.doc.getText('t').insert(0, 'kept');
    await until(() => first.hocuspocus.documents.get(room)?.getText('t').toString() === 'kept');
    a.close();
    await first.destroy();
    const second = await start({ data });
    const b = peer(second.url, room, secret);
    await b.synced;
    await until(() => text(b.doc) === 'kept');
    expect(await peer(second.url, room, token()).failed).toBeTruthy();
  });

  test('an update past the room limit is refused', async () => {
    const { relay, url } = await start({ maxRoom: 4096 });
    const room = roomId();
    const secret = token();
    const a = peer(url, room, `create:${secret}`);
    await a.synced;
    a.doc.getText('t').insert(0, 'small');
    await until(() => relay.hocuspocus.documents.get(room)?.getText('t').toString() === 'small');
    a.doc.getText('t').insert(5, 'x'.repeat(10_000));
    await new Promise((r) => setTimeout(r, 200));
    expect(relay.hocuspocus.documents.get(room)?.getText('t').toString() ?? 'small').toBe('small');
  });

  test('small updates that together pass the room limit stop at it', async () => {
    const { relay, url } = await start({ maxRoom: 4096 });
    const room = roomId();
    const a = peer(url, room, `create:${token()}`);
    void relay;
    await a.synced;
    const held: number[] = [];
    relay.hocuspocus.configuration.extensions.push({ async onChange({ document }) { held.push(document.getText('t').length); } });
    for (let i = 0; i < 8; i++) {
      a.doc.getText('t').insert(a.doc.getText('t').length, String(i).repeat(1000));
      await new Promise((r) => setTimeout(r, 30));
    }
    // The connection is closed as a full room, which the application reads to stop reconnecting.
    const close = await a.closed;
    // The reason, not the code: the close reaches the provider as a normal close (1000) with the relay's reason (ED-004 section 10.4).
    expect(close.reason).toBe('Room Full');
    a.close();
    expect(Math.max(...held)).toBeGreaterThan(0);
    expect(Math.max(...held)).toBeLessThan(4096);
  });

  test('a room untouched for longer than the idle days is deleted when nobody has it open', async () => {
    let clock = Date.parse('2026-09-26T00:00:00Z');
    const { relay, url, data } = await start({ idleDays: 30, now: () => clock });
    const room = roomId();
    const a = peer(url, room, `create:${token()}`);
    await a.synced;
    a.doc.getText('t').insert(0, 'x');
    await until(() => relay.hocuspocus.documents.get(room)?.getText('t').toString() === 'x');
    clock += 31 * 24 * 60 * 60 * 1000;
    expect(await relay.sweep()).toEqual([]);   // still open
    a.close();
    await until(() => !relay.hocuspocus.documents.has(room));
    expect(await relay.sweep()).toEqual([room]);
    expect((await readdir(data)).filter((f) => f.startsWith(room))).toEqual([]);
  });

  test('a page from another site is refused unless its origin is named; a local page and a client without an origin connect', async () => {
    const { url } = await start();
    const room = roomId();
    const secret = token();
    await peer(url, room, `create:${secret}`, { origin: 'http://localhost:5173' }).synced;
    await peer(url, room, secret).synced;
    const stranger = peer(url, room, secret, { origin: 'https://elsewhere.example' });
    const outcome = await Promise.race([stranger.synced.then(() => 'synced'), new Promise((r) => setTimeout(() => r('refused'), 500))]);
    expect(outcome).toBe('refused');
    const named = await start({ origins: ['https://editor.example'] });
    await peer(named.url, roomId(), `create:${token()}`, { origin: 'https://editor.example' }).synced;
  });

  test('allowOrigins adds to the local rule and to origins, and allowOrigin adds one while it runs; any other origin is still refused (ED-014 section 4.5)', async () => {
    const refused = async (p: { synced: Promise<void> }): Promise<boolean> =>
      (await Promise.race([p.synced.then(() => 'synced'), new Promise((r) => setTimeout(() => r('refused'), 500))])) === 'refused';
    const webview = 'vscode-webview://1a2b3c4d5e6f';
    const local = await start({ allowOrigins: [webview] });
    await peer(local.url, roomId(), `create:${token()}`, { origin: webview }).synced;
    await peer(local.url, roomId(), `create:${token()}`, { origin: 'http://127.0.0.1:4174' }).synced;
    expect(await refused(peer(local.url, roomId(), `create:${token()}`, { origin: 'https://elsewhere.example' }))).toBe(true);
    expect(await refused(peer(local.url, roomId(), `create:${token()}`, { origin: 'vscode-webview://another' }))).toBe(true);

    const named = await start({ origins: ['https://editor.example'], allowOrigins: [webview] });
    await peer(named.url, roomId(), `create:${token()}`, { origin: 'https://editor.example' }).synced;
    await peer(named.url, roomId(), `create:${token()}`, { origin: webview }).synced;
    // `origins` still replaces the local rule.
    expect(await refused(peer(named.url, roomId(), `create:${token()}`, { origin: 'http://localhost:5173' }))).toBe(true);

    const later = 'vscode-webview://later';
    expect(await refused(peer(named.url, roomId(), `create:${token()}`, { origin: later }))).toBe(true);
    named.relay.allowOrigin(later);
    await peer(named.url, roomId(), `create:${token()}`, { origin: later }).synced;
  });
});

describe('the command line', () => {
  test('started through a link of any name, as npm\'s bin starts it, the command runs', async () => {
    // 0.1.0 compared process.argv[1] with its own file, and let a link through only by the bin's name.
    const cli = join(dirname(fileURLToPath(import.meta.url)), '..', 'dist', 'cli.mjs');
    try { await stat(cli); } catch { throw new Error(`build the relay first (npm run build -w @docx4j/editor-relay): no ${cli}`); }
    const bin = await mkdtemp(join(tmpdir(), 'relay-bin-'));
    cleanups.push(() => rm(bin, { recursive: true, force: true }));
    const link = join(bin, 'relay-under-another-name');
    await symlink(cli, link);
    const { stdout } = await promisify(execFile)(process.execPath, [link, '--help']);
    expect(stdout).toContain('usage: docx4j-editor-relay');
  });

  test('the options, and refusals of what is not one', async () => {
    const { parseArguments } = await import('../src/cli.mjs');
    expect(parseArguments(['--port', '5000', '--host', '0.0.0.0', '--data', '/tmp/r', '--max-room', '1024', '--idle-days', '7', '--origin', 'https://a.example', '--origin', 'https://b.example']))
      .toEqual({ port: 5000, host: '0.0.0.0', data: '/tmp/r', maxRoom: 1024, idleDays: 7, origins: ['https://a.example', 'https://b.example'] });
    expect(parseArguments([])).toEqual({});
    expect(() => parseArguments(['--port'])).toThrow(/needs a value/);
    expect(() => parseArguments(['--port', 'x'])).toThrow(/needs a number/);
    expect(() => parseArguments(['--nope'])).toThrow(/unknown option/);
  });
});
