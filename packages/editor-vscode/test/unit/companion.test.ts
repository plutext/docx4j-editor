// SPDX-License-Identifier: GPL-3.0-or-later
//
// The companion as the extension runs it (ED-014 section 4.5), over fakes: the ports kept, picked in range and replaced when
// taken; the spawn's arguments and environment; the rooms' folder; a page's policy; where it may run; the offer rule (P40);
// a room named by a message; and the process, with its single start, its invitations, its restarts and its close.
import { describe, expect, test, vi } from 'vitest';
import { join } from 'node:path';
import { EventEmitter } from 'node:events';
import {
  CompanionRunner, EXITS_PER_MINUTE, KILL_AFTER_MS, START_MS, PORT_MAX, PORT_MIN, TOKEN_ENV, canRunCompanion, choosePorts, claudeCodeLine, companionArgs, companionEnv,
  connectFor, isPorts, isWebviewOrigin, joinOf, pageReaches, nodeMajor, offered, randomPorts, relayAddressOf, roomsFolder, withOrigin, ORIGINS_KEPT,
  type ChildLike, type Launch,
} from '../../src/companion.js';
import { parseRoomLink, type FromCompanion, type ToCompanion } from '../../src/shared.js';

const launch = (over: Partial<Launch> = {}): Launch => ({
  cli: '/ext/companion/cli.mjs',
  ports: { port: 24402, relay: 22642 },
  data: '/storage/rooms',
  app: '/ext/app',
  roots: ['/work/a', '/work/b'],
  name: 'Agent',
  origins: ['vscode-webview://abc'],
  token: 'secret-token',
  ...over,
});

describe('the ports (P37)', () => {
  test('the kept pair reused while both are free; a fresh pair in range otherwise, its two free and different', async () => {
    const kept = { port: 24402, relay: 22642 };
    expect(await choosePorts(kept, async () => true)).toEqual({ ports: kept, fresh: false });
    const taken = new Set([22642]);
    const free = async (p: number) => !taken.has(p);
    const out = await choosePorts(kept, free);
    expect(out.fresh).toBe(true);
    expect(out.ports).not.toEqual(kept);
    for (const p of [out.ports.port, out.ports.relay]) { expect(p).toBeGreaterThanOrEqual(PORT_MIN); expect(p).toBeLessThanOrEqual(PORT_MAX); expect(taken.has(p)).toBe(false); }
    expect(out.ports.port).not.toBe(out.ports.relay);
    // None kept: a fresh pair.
    expect((await choosePorts(undefined, async () => true)).fresh).toBe(true);
    // A pair whose first pick is taken: picked again.
    let calls = 0;
    const firstTaken = async (): Promise<boolean> => { calls++; return calls > 2; };
    expect((await choosePorts(undefined, firstTaken)).fresh).toBe(true);
  });

  test('a fresh pair at the range\'s ends, never one port twice; a kept value is read only when it is a pair in range', () => {
    expect(randomPorts(() => 0).port).toBe(PORT_MIN);
    const top = randomPorts(() => 0.99999999);
    expect(top.port).toBe(PORT_MAX);
    expect(top.relay).not.toBe(top.port);
    expect(isPorts({ port: 24402, relay: 22642 })).toBe(true);
    expect(isPorts({ port: 4174, relay: 22642 })).toBe(false);
    expect(isPorts({ port: 24402, relay: 24402 })).toBe(false);
    expect(isPorts({ port: '24402', relay: 22642 })).toBe(false);
    expect(isPorts(undefined)).toBe(false);
    expect(relayAddressOf({ port: 24402, relay: 22642 })).toBe('ws://127.0.0.1:22642');
  });
});

describe('the spawn (P36, P39)', () => {
  test('the arguments: HTTP alone and the channel, the ports, rooms, application, file roots, name and origins; never the token', () => {
    const args = companionArgs(launch());
    expect(args).toEqual([
      '/ext/companion/cli.mjs', '--http', '--ipc', '--port', '24402', '--relay-port', '22642', '--data', '/storage/rooms', '--app', '/ext/app',
      '--root', '/work/a', '--root', '/work/b', '--name', 'Agent', '--allow-origin', 'vscode-webview://abc', '--mcp-token-env', TOKEN_ENV,
    ]);
    expect(args.join(' ')).not.toContain('secret-token');
    // A window with no folder: no root, so open and save refuse every path while Invite works.
    expect(companionArgs(launch({ roots: [], origins: [] }))).not.toContain('--root');
  });

  test('the environment: VS Code\'s own, Electron run as Node, and the token', () => {
    const env = companionEnv({ PATH: '/bin', HOME: '/home/a', UNSET: undefined }, 'secret-token');
    expect(env).toEqual({ PATH: '/bin', HOME: '/home/a', ELECTRON_RUN_AS_NODE: '1', [TOKEN_ENV]: 'secret-token' });
  });

  test('the rooms: in the workspace\'s storage, else one folder per relay port under the global storage', () => {
    expect(roomsFolder('/ws/storage', '/global', { port: 24402, relay: 22642 }, join)).toBe(join('/ws/storage', 'rooms'));
    expect(roomsFolder(undefined, '/global', { port: 24402, relay: 22642 }, join)).toBe(join('/global', 'rooms-22642'));
  });
});

describe('where it may run (P15, P30)', () => {
  const window = { trusted: true, remoteName: undefined, folderSchemes: ['file'], node: '24.22.1' };
  test('a trusted local window on disk, with Node 22 or later', () => {
    expect(canRunCompanion(window)).toEqual({ ok: true });
    expect(canRunCompanion({ ...window, folderSchemes: [] })).toEqual({ ok: true });
    expect(canRunCompanion({ ...window, node: '22.19.0' })).toEqual({ ok: true });
    expect(canRunCompanion({ ...window, trusted: false })).toEqual({ ok: false, why: 'untrusted' });
    expect(canRunCompanion({ ...window, remoteName: 'ssh-remote' })).toEqual({ ok: false, why: 'remote' });
    expect(canRunCompanion({ ...window, folderSchemes: ['file', 'vscode-vfs'] })).toEqual({ ok: false, why: 'virtual' });
    expect(canRunCompanion({ ...window, node: '20.18.0' })).toEqual({ ok: false, why: 'node' });
    expect(nodeMajor('v22.12.0')).toBe(22);
    expect(nodeMajor('garbage')).toBe(0);
  });

  test('the offer (P40): only once the person has admitted the agent here, where it may run', () => {
    expect(offered(false, { ok: true })).toBe(false);
    expect(offered(true, { ok: true })).toBe(true);
    expect(offered(true, { ok: false, why: 'untrusted' })).toBe(false);
  });
});

describe('a page\'s policy and its rooms (section 4.5)', () => {
  test('connect names the workspace\'s relay, and a room\'s on another relay; nothing else', () => {
    expect(connectFor('ws://127.0.0.1:22642', undefined)).toEqual(['ws://127.0.0.1:22642']);
    expect(connectFor('ws://127.0.0.1:22642', 'ws://127.0.0.1:22642')).toEqual(['ws://127.0.0.1:22642']);
    expect(connectFor('ws://127.0.0.1:22642', 'wss://relay.example.org/rooms')).toEqual(['ws://127.0.0.1:22642', 'wss://relay.example.org']);
    expect(connectFor(undefined, 'ws://127.0.0.1:30000')).toEqual(['ws://127.0.0.1:30000']);
    expect(connectFor(undefined, undefined)).toEqual([]);
    // Not a WebSocket address: nothing added.
    expect(connectFor('http://127.0.0.1:24402', 'javascript:alert(1)')).toEqual([]);
  });

  test('a room named by a message is checked as a link is; parseRoomLink in its new place (host-protocol.ts)', () => {
    const room = { room: 'AAAAAAAAAAAAAAAAAAAAAA', token: 'a-token', relay: 'ws://127.0.0.1:22642' };
    expect(joinOf(room)).toEqual(room);
    expect(joinOf({ ...room, room: 'short' })).toBeUndefined();
    expect(joinOf({ ...room, relay: 'http://127.0.0.1:22642' })).toBeUndefined();
    expect(joinOf({ room: room.room })).toBeUndefined();
    expect(joinOf(null)).toBeUndefined();
    expect(parseRoomLink(`http://127.0.0.1:24402/#room=${room.room}&token=a-token&relay=ws%3A%2F%2F127.0.0.1%3A22642`)).toEqual(room);
  });

  test('only a webview\'s origin is kept for the relay; a page reaches the relay its policy was written with, and no other', () => {
    expect(isWebviewOrigin('vscode-webview://14c3a720puos8rhqc9d1afhahcbc5jul0477i3psnc5s58sc770k')).toBe(true);
    expect(isWebviewOrigin('http://localhost:5173')).toBe(false);
    expect(isWebviewOrigin('vscode-webview://abc/path')).toBe(false);
    expect(isWebviewOrigin('null')).toBe(false);
    expect(pageReaches('ws://127.0.0.1:22642', 'ws://127.0.0.1:22642')).toBe(true);
    expect(pageReaches('ws://127.0.0.1:22642', 'ws://127.0.0.1:30000')).toBe(false);
    expect(pageReaches(undefined, 'ws://127.0.0.1:22642')).toBe(false);
  });

  test('the origins kept: the newest last, each once, at most a few', () => {
    expect(withOrigin(['a'], 'a')).toEqual(['a']);
    expect(withOrigin(['a'], 'b')).toEqual(['a', 'b']);
    const many = Array.from({ length: ORIGINS_KEPT }, (_, i) => `o${i}`);
    expect(withOrigin(many, 'new')).toEqual([...many.slice(1), 'new']);
  });

  test('Claude Code\'s line: HTTP, the address, the token as a header', () => {
    expect(claudeCodeLine('http://127.0.0.1:24402/mcp', 'tok')).toBe('claude mcp add --transport http docx4j-editor http://127.0.0.1:24402/mcp --header "Authorization: Bearer tok"');
  });
});

/** A child that the test drives: what it was sent, and its messages, errors and exit. */
class FakeChild extends EventEmitter implements ChildLike {
  readonly sent: ToCompanion[] = [];
  readonly killed: string[] = [];
  connected = true;
  /** A send fails, as one does while the channel closes. */
  failSends = false;
  constructor(readonly pid: number) { super(); }
  send(message: ToCompanion, callback: (error: Error | null) => void): boolean {
    this.sent.push(message);
    callback(this.failSends ? new Error('channel closed') : null);
    return !this.failSends;
  }
  kill(signal: NodeJS.Signals = 'SIGTERM'): boolean { this.killed.push(signal); return true; }
  say(message: FromCompanion): void { this.emit('message', message); }
  exit(code: number | null = 1): void { this.connected = false; this.emit('exit', code, null); }
}

function runner(over: { roomsOn?: (relay: string) => boolean; now?: () => number } = {}) {
  const children: FakeChild[] = [];
  const log: string[] = [];
  const started = vi.fn();
  const r = new CompanionRunner({
    spawn: () => { const c = new FakeChild(children.length + 1); children.push(c); return c; },
    launch: async () => launch(),
    log: (l) => log.push(l),
    started,
    ...over,
  });
  return { r, children, log, started };
}

const flush = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));

describe('the process (P37)', () => {
  test('started once though asked twice; listening resolves both, and every page is told', async () => {
    const { r, children, started } = runner();
    const a = r.start();
    const b = r.start();
    await flush();
    expect(children).toHaveLength(1);
    children[0]!.say({ type: 'listening', url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
    expect(await a).toMatchObject({ url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
    expect(await b).toBe(await a);
    expect(started).toHaveBeenCalledTimes(1);
    expect(await r.start()).toBe(await a);
    expect(children).toHaveLength(1);
  });

  test('an invitation goes over the channel and is answered by its id; a refusal rejects with the companion\'s words', async () => {
    const { r, children } = runner();
    const started = r.start();
    await flush();
    children[0]!.say({ type: 'listening', url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
    await started;
    const ok = r.invite('room-0123456789abcdef', 'tok');
    await flush();
    const sent = children[0]!.sent.at(-1) as Extract<ToCompanion, { type: 'invite' }>;
    expect(sent).toMatchObject({ type: 'invite', room: 'room-0123456789abcdef', token: 'tok' });
    children[0]!.say({ type: 'invited', id: sent.id });
    await expect(ok).resolves.toBeUndefined();
    const refused = r.invite('room-0123456789abcdef', 'wrong');
    await flush();
    const again = children[0]!.sent.at(-1) as Extract<ToCompanion, { type: 'invite' }>;
    children[0]!.say({ type: 'invited', id: again.id, status: 403, error: 'wrong token for this room' });
    await expect(refused).rejects.toThrow('wrong token for this room');
  });

  test('close posts close and kills nothing; its exit after is no restart', async () => {
    const { r, children } = runner({ roomsOn: () => true });
    const started = r.start();
    await flush();
    children[0]!.say({ type: 'listening', url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
    await started;
    r.close();
    expect(children[0]!.sent.at(-1)).toEqual({ type: 'close' });
    children[0]!.exit(0);
    await flush();
    expect(children).toHaveLength(1);
    expect(r.running).toBeUndefined();
  });

  test('an exit unasked: restarted at once while a page is in a room on its relay, else at the next need', async () => {
    let inRoom = true;
    const { r, children } = runner({ roomsOn: (relay) => inRoom && relay === 'ws://127.0.0.1:22642' });
    const listen = (c: FakeChild): void => c.say({ type: 'listening', url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
    const first = r.start();
    await flush();
    listen(children[0]!);
    await first;
    children[0]!.exit(1);
    await flush();
    expect(children).toHaveLength(2);
    listen(children[1]!);
    await flush();
    expect(r.running).toBeDefined();
    inRoom = false;
    children[1]!.exit(1);
    await flush();
    expect(children).toHaveLength(2);
    expect(r.running).toBeUndefined();
    const next = r.start();
    await flush();
    expect(children).toHaveLength(3);
    listen(children[2]!);
    await next;
  });

  test('five exits in a minute leave it stopped', async () => {
    let now = 0;
    const { r, children } = runner({ roomsOn: () => false, now: () => now });
    for (let i = 0; i < EXITS_PER_MINUTE; i++) {
      const s = r.start();
      await flush();
      children.at(-1)!.say({ type: 'listening', url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
      await s;
      now += 5_000;
      children.at(-1)!.exit(1);
      await flush();
    }
    expect(r.gaveUp).toBe(true);
    await expect(r.start()).rejects.toThrow('stopped');
    expect(children).toHaveLength(EXITS_PER_MINUTE);
  });

  test('an exit before listening fails the start, and is not counted toward stopping: a missing bundle never reads as five crashes', async () => {
    const { r, children, log } = runner({ roomsOn: () => true });
    for (let i = 0; i < EXITS_PER_MINUTE + 1; i++) {
      const s = r.start();
      await flush();
      children.at(-1)!.exit(1);
      await expect(s).rejects.toThrow(/exited before it listened/);
    }
    expect(r.gaveUp).toBe(false);
    // Nothing restarted it by itself: each start was asked for.
    expect(children).toHaveLength(EXITS_PER_MINUTE + 1);
    expect(log.some((l) => l.includes('five times'))).toBe(false);
  });

  test('an exit while invitations wait fails them', async () => {
    const { r, children } = runner({ roomsOn: () => false });
    const s = r.start();
    await flush();
    children[0]!.say({ type: 'listening', url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
    await s;
    const waiting = r.invite('room-0123456789abcdef', 'tok');
    await flush();
    children[0]!.exit(1);
    await expect(waiting).rejects.toThrow('the companion stopped');
  });

  test('exits more than a minute apart are not counted together', async () => {
    let now = 0;
    const { r, children } = runner({ roomsOn: () => false, now: () => now });
    for (let i = 0; i < EXITS_PER_MINUTE * 2; i++) {
      const s = r.start();
      await flush();
      children.at(-1)!.say({ type: 'listening', url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
      await s;
      now += 20_000;
      children.at(-1)!.exit(1);
      await flush();
    }
    expect(r.gaveUp).toBe(false);
  });

  test('a start that does not listen in time: rejected, the child asked to close and killed if it stays, a second start on a new child; the first one\'s late listening counts for nothing (section 9.19 finding 2)', async () => {
    vi.useFakeTimers();
    try {
      const { r, children, started } = runner();
      const first = r.start();
      first.catch(() => undefined);
      await vi.advanceTimersByTimeAsync(0);
      expect(children).toHaveLength(1);
      await vi.advanceTimersByTimeAsync(START_MS);
      await expect(first).rejects.toThrow('did not start in time');
      expect(children[0]!.sent).toContainEqual({ type: 'close' });
      const second = r.start();
      await vi.advanceTimersByTimeAsync(0);
      expect(children).toHaveLength(2);
      // The first speaks at last: ignored.
      children[0]!.say({ type: 'listening', url: 'http://127.0.0.1:1', relay: 'ws://127.0.0.1:2' });
      expect(r.running).toBeUndefined();
      expect(started).not.toHaveBeenCalled();
      children[1]!.say({ type: 'listening', url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
      await expect(second).resolves.toMatchObject({ url: 'http://127.0.0.1:24402' });
      // Still running after the grace: killed. Its exit then is no restart and no stop of the second.
      await vi.advanceTimersByTimeAsync(KILL_AFTER_MS);
      expect(children[0]!.killed).toEqual(['SIGKILL']);
      children[0]!.exit(null);
      await vi.advanceTimersByTimeAsync(0);
      expect(r.running?.url).toBe('http://127.0.0.1:24402');
      expect(children).toHaveLength(2);
    } finally {
      vi.useRealTimers();
    }
  });

  test('every error the child reports is heard, a second one too, and a send that fails is logged, never thrown (section 9.19 finding 3)', async () => {
    const { r, children, log } = runner();
    const s = r.start();
    await flush();
    children[0]!.say({ type: 'listening', url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
    await s;
    children[0]!.emit('error', new Error('one'));
    children[0]!.emit('error', new Error('two'));
    children[0]!.failSends = true;
    r.post({ type: 'allowOrigin', origin: 'vscode-webview://abc' });
    r.post({ type: 'roots', roots: [] });
    expect(log.filter((l) => l.includes('could not tell it'))).toHaveLength(2);
    expect(log.some((l) => l.includes('two'))).toBe(true);
  });

  test('a restart of a companion that does not exit: killed after the grace, then started again', async () => {
    vi.useFakeTimers();
    try {
      const { r, children } = runner();
      const s = r.start();
      await vi.advanceTimersByTimeAsync(0);
      children[0]!.say({ type: 'listening', url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
      await s;
      const again = r.restart();
      await vi.advanceTimersByTimeAsync(KILL_AFTER_MS);
      expect(children[0]!.killed).toEqual(['SIGKILL']);
      children[0]!.exit(null);
      await vi.advanceTimersByTimeAsync(0);
      expect(children).toHaveLength(2);
      children[1]!.say({ type: 'listening', url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
      await expect(again).resolves.toMatchObject({ url: 'http://127.0.0.1:24402' });
    } finally {
      vi.useRealTimers();
    }
  });

  test('restart: closed, its exit awaited, started again on the launch as it is now', async () => {
    const { r, children } = runner();
    const s = r.start();
    await flush();
    children[0]!.say({ type: 'listening', url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
    await s;
    const again = r.restart();
    await flush();
    expect(children[0]!.sent.at(-1)).toEqual({ type: 'close' });
    children[0]!.exit(0);
    await flush();
    expect(children).toHaveLength(2);
    children[1]!.say({ type: 'listening', url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
    await expect(again).resolves.toMatchObject({ url: 'http://127.0.0.1:24402' });
  });
});
