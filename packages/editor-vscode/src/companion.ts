// SPDX-License-Identifier: GPL-3.0-or-later
//
// The companion as the extension runs it (ED-014 section 4.5, P36 to P40), with no `vscode` import: where it may run, its
// ports, its arguments and environment, the page's policy, the offer rule, and the process with its channel and restarts.
// `agent.ts` gives it VS Code's storage, settings and commands; the unit tests give it fakes.
//
// The process is the companion's bundle (`companion/cli.mjs`, P36) under VS Code's own Node: `process.execPath` with
// `ELECTRON_RUN_AS_NODE=1`, an IPC channel and stderr read, MCP over HTTP alone. Nothing kills it: `close` asks it to write
// its rooms and exit, and its channel's closing (the extension host gone, by a quit, a reload or a crash) does the same,
// so no companion outlives its window (measured, section 9.17).
import { parseRoomLink, type FromCompanion, type RoomLinkData, type ToCompanion } from './shared.js';

/** The environment variable that carries MCP's token to the companion: never an argument, which another user can read (P39). */
export const TOKEN_ENV = 'DOCX4J_EDITOR_MCP_TOKEN';

/** The ports' range (P37): outside the ephemeral ranges (Linux's from 32768, Windows' and macOS's from 49152). */
export const PORT_MIN = 20000;
export const PORT_MAX = 32767;

/** The companion's two ports, kept per workspace (P37): its own (MCP, the editor it serves) and its relay's. */
export interface CompanionPorts {
  port: number;
  relay: number;
}

/** The relay's WebSocket address on these ports: what a page's policy names and a room link carries. */
export function relayAddressOf(ports: CompanionPorts): string {
  return `ws://127.0.0.1:${ports.relay}`;
}

/** The companion's own address. */
export function companionAddressOf(ports: CompanionPorts): string {
  return `http://127.0.0.1:${ports.port}`;
}

/** MCP's address, which VS Code's definition and the address command give. */
export function mcpAddressOf(ports: CompanionPorts): string {
  return `${companionAddressOf(ports)}/mcp`;
}

/** Whether a value is a kept pair of ports. */
export function isPorts(value: unknown): value is CompanionPorts {
  const p = value as Partial<CompanionPorts> | null;
  const ok = (n: unknown): boolean => typeof n === 'number' && Number.isInteger(n) && n >= PORT_MIN && n <= PORT_MAX;
  return !!p && typeof p === 'object' && ok(p.port) && ok(p.relay) && p.port !== p.relay;
}

/** A fresh pair in range, the two different. */
export function randomPorts(random: () => number = Math.random): CompanionPorts {
  const pick = (): number => PORT_MIN + Math.floor(random() * (PORT_MAX - PORT_MIN + 1));
  const port = pick();
  const relay = pick();
  // Never one port twice: the next one along, inside the range.
  return { port, relay: relay !== port ? relay : port === PORT_MAX ? port - 1 : port + 1 };
}

/**
 * The ports to start on (P37): the kept pair if both are free, else a fresh pair whose two are free, which the caller keeps.
 * `fresh` says the kept pair was replaced, so pages written before name a relay port that is no longer the companion's.
 */
export async function choosePorts(kept: CompanionPorts | undefined, isFree: (port: number) => Promise<boolean>, random: () => number = Math.random): Promise<{ ports: CompanionPorts; fresh: boolean }> {
  if (kept && isPorts(kept) && await isFree(kept.port) && await isFree(kept.relay)) return { ports: kept, fresh: false };
  for (let i = 0; i < 50; i++) {
    const ports = randomPorts(random);
    if (await isFree(ports.port) && await isFree(ports.relay)) return { ports, fresh: true };
  }
  throw new Error('no free pair of ports between 20000 and 32767');
}

/** Why the companion cannot run for a window, or `ok` (ED-014 section 4.5, P15, P30). */
export type CanRun = { ok: true } | { ok: false; why: 'untrusted' | 'remote' | 'virtual' | 'node' };

/** The major version of a Node version string (`24.22.1`); 0 when it does not read. */
export function nodeMajor(version: string): number {
  const major = Number(/^v?(\d+)\./.exec(version)?.[1]);
  return Number.isInteger(major) ? major : 0;
}

/** Where the companion may run: a trusted workspace whose extension host is local, not a virtual one, on Node 22 or later. */
export function canRunCompanion(window: { trusted: boolean; remoteName: string | undefined; folderSchemes: readonly string[]; node: string }): CanRun {
  if (!window.trusted) return { ok: false, why: 'untrusted' };
  if (window.remoteName) return { ok: false, why: 'remote' };
  if (window.folderSchemes.some((s) => s !== 'file')) return { ok: false, why: 'virtual' };
  if (nodeMajor(window.node) < 22) return { ok: false, why: 'node' };
  return { ok: true };
}

/** What the extension starts the companion with. */
export interface Launch {
  /** The bundle's `cli.mjs`. */
  cli: string;
  ports: CompanionPorts;
  /** The rooms' folder (`--data`): the workspace's storage, else a folder per relay port under the global storage. */
  data: string;
  /** The extension's `app/`, which the companion serves (`--app`), so the application ships once. */
  app: string;
  /** The workspace's folders on the `file:` scheme (`--root`); none in a window with no folder. */
  roots: readonly string[];
  /** The agent's name (`--name`, P43). */
  name: string;
  /** The page origins the extension has seen (`--allow-origin`, P38). */
  origins: readonly string[];
  /** MCP's token, which goes in the environment (P39). */
  token: string;
}

/** The companion's arguments: MCP over HTTP alone, the channel, its ports, rooms, application, roots, name and origins. No token. */
export function companionArgs(launch: Launch): string[] {
  return [
    launch.cli,
    '--http', '--ipc',
    '--port', String(launch.ports.port),
    '--relay-port', String(launch.ports.relay),
    '--data', launch.data,
    '--app', launch.app,
    ...launch.roots.flatMap((r) => ['--root', r]),
    '--name', launch.name,
    ...launch.origins.flatMap((o) => ['--allow-origin', o]),
    '--mcp-token-env', TOKEN_ENV,
  ];
}

/** The companion's environment: VS Code's own, its Electron run as Node, and the token. */
export function companionEnv(base: Readonly<Record<string, string | undefined>>, token: string): Record<string, string> {
  const env: Record<string, string> = {};
  for (const [key, value] of Object.entries(base)) if (value !== undefined) env[key] = value;
  env['ELECTRON_RUN_AS_NODE'] = '1';
  env[TOKEN_ENV] = token;
  return env;
}

/** The rooms' folder: in the workspace's storage, else one per relay port under the global storage, which every window with no folder shares. */
export function roomsFolder(storage: string | undefined, globalStorage: string, ports: CompanionPorts, join: (...parts: string[]) => string): string {
  return storage !== undefined ? join(storage, 'rooms') : join(globalStorage, `rooms-${ports.relay}`);
}

/**
 * A page's `connect-src` additions (ED-014 section 4.5): the workspace's relay, and a room's relay where the document opens
 * into a room on another (a Join, a backup's room, a hand-over's). Nothing else.
 */
export function connectFor(workspaceRelay: string | undefined, roomRelay: string | undefined): string[] {
  const out: string[] = [];
  for (const relay of [workspaceRelay, roomRelay]) {
    if (relay === undefined || out.includes(relay)) continue;
    try {
      const url = new URL(relay);
      if (url.protocol !== 'ws:' && url.protocol !== 'wss:') continue;
      out.push(`${url.protocol}//${url.host}`);
    } catch { /* not an address */ }
  }
  return [...new Set(out)];
}

/** A room as a message names it (`untitled { join }`), checked as a link is (`parseRoomLink`); undefined if it is not one. */
export function joinOf(value: unknown): RoomLinkData | undefined {
  const v = value as Partial<RoomLinkData> | null;
  if (!v || typeof v !== 'object' || typeof v.room !== 'string' || typeof v.token !== 'string' || typeof v.relay !== 'string') return undefined;
  return parseRoomLink(new URLSearchParams({ room: v.room, token: v.token, relay: v.relay }).toString());
}

/** Whether VS Code is offered the MCP server (P40): once the person has admitted the agent in this workspace, where it can run. */
export function offered(admitted: boolean, canRun: CanRun): boolean {
  return admitted && canRun.ok;
}

/** The line for Claude Code (ED-014 section 4.5): its default local scope keeps the token out of the project's files. */
export function claudeCodeLine(address: string, token: string): string {
  return `claude mcp add --transport http docx4j-editor ${address} --header "Authorization: Bearer ${token}"`;
}

/**
 * Whether an origin is a VS Code webview's (`vscode-webview://<id>`), the only kind a page of this extension posts with
 * `ready`. Nothing else is kept or allowed at the relay (Fable's review, section 9.19).
 */
export function isWebviewOrigin(origin: string): boolean {
  return /^vscode-webview:\/\/[A-Za-z0-9-]{1,128}$/.test(origin);
}

/**
 * Whether a page reaches the running companion's relay: its policy names the relay it was written with, so a page written
 * before the ports moved does not, and is told to reopen (section 4.5). A page written with none (the companion could not
 * run then) names none either.
 */
export function pageReaches(writtenRelay: string | undefined, runningRelay: string): boolean {
  return writtenRelay === runningRelay;
}

/** The origins kept, the newest last, at most this many: one per view type and VS Code release, in practice. */
export const ORIGINS_KEPT = 8;

/** An origin added to the kept ones; the same list if it was there. */
export function withOrigin(kept: readonly string[], origin: string): string[] {
  if (kept.includes(origin)) return [...kept];
  return [...kept, origin].slice(-ORIGINS_KEPT);
}

// ---- The process.

/** The child as the runner needs it: Node's `ChildProcess`, or a fake. */
export interface ChildLike {
  readonly pid?: number | undefined;
  readonly connected: boolean;
  /** With a callback, a send that fails is reported there, never as an `'error'` event. */
  send(message: ToCompanion, callback: (error: Error | null) => void): boolean;
  on(event: 'message', listener: (message: FromCompanion) => void): unknown;
  on(event: 'error', listener: (error: Error) => void): unknown;
  once(event: 'exit', listener: (code: number | null, signal: string | null) => void): unknown;
  kill?(signal?: NodeJS.Signals): boolean;
}

/** What the runner gets from a start: the addresses, and the launch it was started with. */
export interface Running {
  url: string;
  relay: string;
  launch: Launch;
}

export interface RunnerDeps {
  spawn(launch: Launch): ChildLike;
  /** The launch for a start: ports chosen, token read, origins and roots as they are now. */
  launch(): Promise<Launch>;
  log(line: string): void;
  now?(): number;
  /** Whether a page is in a room on this relay now: the companion is restarted at once after an exit then. */
  roomsOn?(relay: string): boolean;
  /** Told when the companion is listening, after any start or restart. */
  started?(running: Running): void;
  /** Told when it exits or its relay's state changes. */
  stopped?(): void;
}

/** Exits not asked for, of a companion that had listened: five in a minute leave it stopped. */
export const EXITS_PER_MINUTE = 5;
/** How long a start waits for `listening`. */
export const START_MS = 30_000;
/**
 * How long a companion asked to close has before it is killed: one that never listened (a start that timed out), and one
 * that does not exit for a restart. Its relay's own close is bounded at five seconds (`server.mts`), so a companion still
 * running after this is hung and writing nothing; the design's "no kill" is about cutting room writes (section 9.16).
 */
export const KILL_AFTER_MS = 10_000;

/**
 * The companion's process (P37): started on demand, once at a time; restarted on its ports when it exits unasked, at once while
 * a page is in a room on its relay and else at the next need; left stopped after five exits in a minute, as the relay's own
 * watch is. `close` posts `close` and kills nothing.
 */
export class CompanionRunner {
  private child: ChildLike | undefined;
  private starting: Promise<Running> | undefined;
  private current: Running | undefined;
  private readonly exits: number[] = [];
  private closing = false;
  private nextInvite = 1;
  private readonly invites = new Map<number, { resolve: () => void; reject: (e: Error) => void }>();
  /** Five exits in a minute: no more starts until the window reloads. */
  gaveUp = false;

  constructor(private readonly deps: RunnerDeps) {}

  /** The companion, when it is listening. */
  get running(): Running | undefined {
    return this.current;
  }

  /** Started if need be; its addresses once it listens. */
  start(): Promise<Running> {
    if (this.current) return Promise.resolve(this.current);
    if (this.gaveUp) return Promise.reject(new Error('stopped'));
    this.starting ??= this.spawn().finally(() => { this.starting = undefined; });
    return this.starting;
  }

  private async spawn(): Promise<Running> {
    this.closing = false;
    const launch = await this.deps.launch();
    const child = this.deps.spawn(launch);
    this.child = child;
    return new Promise<Running>((resolve, reject) => {
      let listening = false;
      /** Given up on: it did not listen in time. Asked to close, killed if it does not; nothing it says counts (Fable's review, section 9.19). */
      let abandoned = false;
      let settled = false;
      let gone = false;
      const fail = (e: Error): void => { if (!settled) { settled = true; reject(e); } };
      const timer = setTimeout(() => {
        if (listening) return;
        abandoned = true;
        if (this.child === child) this.child = undefined;
        this.deps.log(`did not start in ${START_MS / 1000} s: stopping it`);
        this.ask(child, { type: 'close' });
        setTimeout(() => { if (!gone) child.kill?.('SIGKILL'); }, KILL_AFTER_MS);
        fail(new Error('the companion did not start in time'));
      }, START_MS);
      child.on('message', (m) => {
        if (!m || typeof m !== 'object' || abandoned) return;
        if (m.type === 'listening' && !listening) {
          listening = true;
          clearTimeout(timer);
          this.current = { url: m.url, relay: m.relay, launch };
          this.deps.log(`listening at ${m.url}, its relay at ${m.relay}`);
          this.deps.started?.(this.current);
          settled = true;
          resolve(this.current);
        } else if (m.type === 'relay') {
          this.deps.log(`its relay is ${m.state}`);
        } else if (m.type === 'invited') {
          const waiting = this.invites.get(m.id);
          if (!waiting) return;
          this.invites.delete(m.id);
          if ('error' in m) waiting.reject(new Error(m.error)); else waiting.resolve();
        }
      });
      // Every error, not the first alone: a second one with no listener would throw in the extension host.
      child.on('error', (e) => {
        this.deps.log(listening ? `its process reported: ${e.message}` : `could not start: ${e.message}`);
        if (!listening) { clearTimeout(timer); fail(e); }
      });
      child.once('exit', (code, signal) => {
        gone = true;
        clearTimeout(timer);
        const how = signal ?? `code ${code}`;
        if (abandoned) { this.deps.log(`the companion that did not start has stopped (${how})`); return; }
        if (this.child === child) this.child = undefined;
        const was = this.current;
        if (this.current?.launch === launch) this.current = undefined;
        for (const w of this.invites.values()) w.reject(new Error('the companion stopped'));
        this.invites.clear();
        this.deps.stopped?.();
        // A start that failed is its caller's to say, and is not counted: each start is a person's or VS Code's act.
        if (!listening) { this.deps.log(`exited before it listened (${how})`); fail(new Error(`the companion exited before it listened (${how})`)); return; }
        if (this.closing) { this.deps.log('stopped'); return; }
        this.deps.log(`exited (${how})`);
        this.exited(was);
      });
    });
  }

  /** An exit not asked for: counted, and restarted at once while a page is in a room on its relay. */
  private exited(was: Running | undefined): void {
    const now = this.deps.now?.() ?? Date.now();
    this.exits.push(now);
    while (this.exits.length > 0 && now - this.exits[0]! > 60_000) this.exits.shift();
    if (this.exits.length >= EXITS_PER_MINUTE) {
      this.gaveUp = true;
      this.deps.log(`stopped ${EXITS_PER_MINUTE} times in a minute and is left stopped until the window reloads`);
      return;
    }
    if (was && this.deps.roomsOn?.(was.relay)) {
      this.deps.log('restarting it at once: a page is in a room on its relay');
      this.start().catch((e: unknown) => this.deps.log(`the restart failed: ${(e as Error).message}`));
    }
  }

  /** Posted to the companion when it runs; dropped otherwise (the next start passes what it needs). */
  post(message: ToCompanion): void {
    if (this.child) this.ask(this.child, message);
  }

  /** A message to one child; a send that fails (the channel closing) is logged, never an `'error'` event. */
  private ask(child: ChildLike, message: ToCompanion): void {
    if (!child.connected) return;
    child.send(message, (e) => { if (e) this.deps.log(`could not tell it ${message.type}: ${e.message}`); });
  }

  /** The agent invited into a room on the companion's relay (`invite` over the channel). */
  async invite(room: string, token: string): Promise<void> {
    await this.start();
    const id = this.nextInvite++;
    await new Promise<void>((resolve, reject) => {
      this.invites.set(id, { resolve, reject });
      this.post({ type: 'invite', id, room, token });
    });
  }

  /** `close` posted: the companion writes its rooms and exits by itself. Nothing is killed (section 9.16 finding 4). */
  close(): void {
    this.closing = true;
    this.post({ type: 'close' });
  }

  /**
   * Closed, and started again on the launch as it is now (a changed token, P39); resolves once it listens. A companion that
   * does not exit within `KILL_AFTER_MS` is hung: it is killed, or VS Code's start of the server would wait for ever.
   */
  async restart(): Promise<Running> {
    const child = this.child;
    if (child) {
      const gone = new Promise<void>((resolve) => child.once('exit', () => resolve()));
      this.close();
      const exited = await Promise.race([gone.then(() => true), new Promise<boolean>((r) => setTimeout(() => r(false), KILL_AFTER_MS))]);
      if (!exited) {
        this.deps.log(`did not stop in ${KILL_AFTER_MS / 1000} s for a restart: killed`);
        child.kill?.('SIGKILL');
        await gone;
      }
    }
    this.current = undefined;
    return this.start();
  }
}
