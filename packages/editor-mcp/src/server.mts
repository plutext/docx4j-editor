// SPDX-License-Identifier: GPL-3.0-or-later
//
// The companion (ED-005 section 8.1): one process on `127.0.0.1` serving the built editor at `/`,
// the companion contract's probe at `/companion/` (`app/companion.ts`), and, from step F3, MCP; and
// the co-editing relay, in a worker thread on a loopback port of its own (`relay-worker.mts`, ED-005
// section 9.3 item 1 (i)), so that it forwards while a tool call holds this process's event loop. A
// document the agent opens is shared into a room on that relay by a headless `EditorDocument`,
// which joins it over loopback as a page does (`RoomConnection`), and the room's link opens in any
// browser on the machine. Requests are checked against DNS
// rebinding: the `Host` must be this server's own loopback name. Nothing is written to stdout,
// which is the MCP protocol's (vocabulary section 6). MCP (step F3, `docs/developer/agent.md`) is served
// over streamable HTTP under `/mcp`, and over stdio by the bin (`cli.mts`).
import { createServer, type IncomingMessage, type Server, type ServerResponse } from 'node:http';
import { createHash, randomBytes, timingSafeEqual } from 'node:crypto';
import { readFile, stat } from 'node:fs/promises';
import { basename, dirname, extname, join, normalize, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Worker } from 'node:worker_threads';
import WebSocket from 'ws';
import { RoomConnection, RoomRefusedError } from '@docx4j/editor-relay/client';
import { EditorDocument, newEditStore, isAgentStopped, onAgents, resumeAgent } from '@docx4j/editor-model';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { existingUnderRoots } from './paths.mjs';
import { STOPPED_HINT, type ToolHost, type ToolRoom } from './kit.mjs';
import { ToolError } from './errors.mjs';
import { createMcpServer } from './tools.mjs';

export interface CompanionOptions {
  /** The folders an agent may open files under (vocabulary section 6: no default). */
  roots: readonly string[];
  /** The agent's name, on its tracked changes, its comments and in the People list (default "Claude"). */
  name?: string;
  /** Default 4174 (docx4j-mcp section 18's); 0 picks a free one. */
  port?: number;
  /** Default `127.0.0.1`: local only. */
  host?: string;
  /** The relay's room files (default `./rooms` under the first root's parent is not assumed: `./.docx4j-editor-rooms`). */
  data?: string;
  /**
   * The relay's own loopback port (ED-005 section 9.3 item 1 (i)). Default the companion's port and
   * one (4175), or a free one when the companion's port is 0; a default that is taken falls back to
   * a free one, since the room links carry the relay's address.
   */
  relayPort?: number;
  /** The built application's folder; default the package's `dist/app`, else the repository's `packages/editor/dist`. */
  app?: string;
  /** Whether the agent may accept and reject tracked changes (ED-005 proposal 12; default no). */
  allowResolve?: boolean;
  /** The inline cap on a result's characters (vocabulary section 6; default 200,000). */
  maxInlineChars?: number;
  /** Where the companion reports what it does: never stdout (default stderr). */
  log?: (line: string) => void;
  /**
   * Page origins allowed beside the local-page rule (ED-014 section 4.5, P38): at the probe's CORS, the invitation and the
   * relay. VS Code's webview (`vscode-webview://…`) reaches only the relay, since its page asks the extension for the rest;
   * the probe and the invitation serve the companion's other uses, a page on a deployed site.
   */
  allowOrigins?: readonly string[];
  /**
   * MCP's credential (ED-014 P39): with one, `/mcp` serves only a request bearing `Authorization: Bearer <token>`, whatever
   * its `Origin`, and refuses the rest with 403. Without one, as run by hand, `/mcp` asks for none and refuses a page of
   * another origin.
   */
  mcpToken?: string;
  /** Told when the relay's state changes after it first listens (`relayState`): the channel's `relay` (ED-014 section 4.5). */
  onRelayState?: (state: 'up' | 'restarting' | 'down') => void;
}

/** A document the agent has open: its room, the headless session that holds it, and its link. */
export interface AgentRoom extends ToolRoom {
  /** The room id: the `document` handle every tool takes (ED-005 section 8.3). */
  readonly id: string;
  /** The file it was opened from; none for a room a person invited it into from their browser. */
  readonly path: string | undefined;
  readonly session: EditorDocument;
  /** The link a person opens to join the room in the editor. */
  readonly link: string;
  /** Leave the room: the direct connection closed, the session no longer written. */
  close(): Promise<void>;
}

export interface Companion {
  /** The relay's WebSocket address (`ws://127.0.0.1:4175`), once listening: its own port, in a worker thread. */
  readonly relay: string;
  /**
   * Whether the relay is running: `up`; `restarting` after its worker stopped (restarted on the same
   * port, so the room links and the connections hold); `down` when the restarts failed, until the
   * companion is restarted (the tools then answer `relay.down`).
   */
  readonly relayState: 'up' | 'restarting' | 'down';
  readonly rooms: ReadonlyMap<string, AgentRoom>;
  /** Start the HTTP server; resolves with its base URL (`http://127.0.0.1:4174`). */
  listen(): Promise<string>;
  /** Open a file under a root into a new room, the agent joined; a file already open answers its room. */
  open(path: string): Promise<AgentRoom>;
  /** Join a room on this relay a person shared, with its token (the invitation, `POST <companion>/rooms`). */
  join(room: string, token: string): Promise<AgentRoom>;
  /** The folders the agent may open files under, from now on: a workspace's folders changing (ED-014 section 4.5). */
  setRoots(roots: readonly string[]): void;
  /** One more page origin allowed, at the probe, the invitation and the relay, from now on (ED-014 section 4.5). */
  allowOrigin(origin: string): void;
  /** The companion as the tools see it. */
  readonly host: ToolHost;
  /** A new MCP server over the companion's rooms (one per transport). */
  mcp(): McpServer;
  /** Leave every room, stop the relay and the server. */
  close(): Promise<void>;
  /** End the relay's worker as a crash would, without `close` (tests: the watch and the restart). */
  crashRelay(): Promise<void>;
}

/**
 * The relay worker's module: beside this one when built; from the package's `dist` when this module
 * runs from source (the suites), since a worker takes a JavaScript file.
 */
const RELAY_WORKER = import.meta.url.endsWith('.mts') ? new URL('../dist/relay-worker.mjs', import.meta.url) : new URL('./relay-worker.mjs', import.meta.url);

/** The agent's colour in the People list and on its caret (`#rrggbb`, as the presence plugin draws it). */
export const AGENT_COLOR = '#6d28d9';

/**
 * The application's page as the companion serves it: marked with where the companion answers, so
 * that the page probes it. A page without the mark (a static deployment) probes nothing
 * (`packages/editor/src/app/companion.ts`, `COMPANION_META`).
 */
export function markedPage(html: string): string {
  const meta = '<meta name="docx4j-companion" content="/companion">';
  const head = /<head(\s[^>]*)?>/i.exec(html);
  if (head) return html.slice(0, head.index + head[0].length) + meta + html.slice(head.index + head[0].length);
  const doctype = /^\s*<!doctype[^>]*>/i.exec(html);
  return doctype ? doctype[0] + meta + html.slice(doctype[0].length) : meta + html;
}

const TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.svg': 'image/svg+xml', '.png': 'image/png',
  '.woff2': 'font/woff2', '.woff': 'font/woff', '.ttf': 'font/ttf', '.wasm': 'application/wasm', '.map': 'application/json',
  '.ico': 'image/x-icon', '.txt': 'text/plain; charset=utf-8',
};

/** The built application: the package's copy, else the repository's build (a development checkout). */
async function findApp(given: string | undefined): Promise<string | undefined> {
  const here = dirname(fileURLToPath(import.meta.url));
  for (const dir of given ? [given] : [join(here, 'app'), join(here, '..', '..', 'editor', 'dist')]) {
    try { if ((await stat(join(dir, 'index.html'))).isFile()) return dir; } catch { /* next */ }
  }
  return undefined;
}

/** The loopback names this server answers to, for the `Host` check. */
function hostsFor(port: number): Set<string> {
  return new Set([`127.0.0.1:${port}`, `localhost:${port}`, `[::1]:${port}`]);
}

/** Whether an `Origin` is a page on this machine (as the relay allows by default), for the probe's CORS. */
function localOrigin(origin: string | undefined): boolean {
  if (!origin) return false;
  try {
    const url = new URL(origin);
    return url.hostname === 'localhost' || url.hostname === '127.0.0.1' || url.hostname === '[::1]';
  } catch {
    return false;
  }
}

/** A refused invitation, with its HTTP status. */
class InviteError extends Error {
  constructor(readonly status: number, message: string) {
    super(message);
  }
}

/** Why an invitation (`join`) was refused, as `POST <companion>/rooms` answers it: its status and its message. */
export function invitationRefusal(e: unknown): { status: number; error: string } {
  return { status: e instanceof InviteError ? e.status : 500, error: (e as Error)?.message ?? String(e) };
}

/** Whether an `Authorization` header bears this token (`Bearer <token>`), compared in constant time. */
export function bearsToken(header: string | undefined, token: string): boolean {
  const match = /^Bearer\s+(\S+)\s*$/i.exec(header ?? '');
  if (!match) return false;
  const digest = (text: string): Buffer => createHash('sha256').update(text).digest();
  return timingSafeEqual(digest(match[1]!), digest(token));
}

export function createCompanion(options: CompanionOptions): Companion {
  const name = options.name ?? 'Claude';
  const host = options.host ?? '127.0.0.1';
  const log = options.log ?? ((line: string) => { process.stderr.write(`${line}\n`); });
  const rooms = new Map<string, AgentRoom>();
  let roots: readonly string[] = [...options.roots];
  const allowed = new Set(options.allowOrigins ?? []);
  /** A page on this machine (the local rule), or one of the origins allowed. */
  const pageAllowed = (origin: string | undefined): boolean => localOrigin(origin) || (origin !== undefined && allowed.has(origin));
  let relayUrl = '';
  let relayWorker: Worker | undefined;
  let relayState: 'up' | 'restarting' | 'down' = 'up';
  /** Settles when a restart in progress has succeeded or failed. */
  let relayRestarted: Promise<void> = Promise.resolve();
  /** When the worker stopped, lately: past five stops in a minute the relay is left down. */
  const relayStops: number[] = [];
  let server: Server | undefined;
  let base = '';
  let hosts = new Set<string>();
  let appDir: string | undefined;

  const probe = (): unknown => ({ tools: [], relay: relayUrl, agent: name, rooms: rooms.size });
  const toolHost: ToolHost = {
    name,
    allowResolve: options.allowResolve === true,
    // Read at each call, so a workspace's folders changing reaches the next `open` and `save`.
    get roots() { return roots; },
    maxInlineChars: options.maxInlineChars ?? 200_000,
    rooms,
    open: (path) => open(path),
    relayReady: () => relayReady(),
  };

  /**
   * MCP over streamable HTTP, stateless: a server and a transport per request, the rooms the
   * companion's. A request from a page (an `Origin`) other than the companion's own is refused:
   * another local site must not drive the agent's tools. With a token (ED-014 P39), only a request
   * bearing it is served, whatever its `Origin`: the token is the credential.
   *
   * Only POST reaches the transport (ED-014 section 9.16). A stateless transport answers a GET for
   * notifications with an event stream on a transport made for that request alone, so nothing ever
   * comes down it, and VS Code's GET would hold a server, a transport and a socket for the window's
   * life. 405 is the streamable HTTP answer for a server that offers no stream, and VS Code and the
   * SDK's client then stop asking.
   */
  async function serveMcp(req: IncomingMessage, res: ServerResponse): Promise<void> {
    if (req.method !== 'POST') { res.writeHead(405, { Allow: 'POST' }).end(); return; }
    const origin = req.headers.origin;
    if (options.mcpToken !== undefined) {
      if (!bearsToken(req.headers.authorization, options.mcpToken)) { res.writeHead(403).end('Forbidden'); return; }
    } else if (origin !== undefined && origin !== base) { res.writeHead(403).end('Forbidden'); return; }
    const server = createMcpServer(toolHost);
    const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined, enableJsonResponse: true });
    res.on('close', () => { void transport.close(); void server.close(); });
    await server.connect(transport);
    await transport.handleRequest(req, res);
  }

  async function serveStatic(req: IncomingMessage, res: ServerResponse, path: string): Promise<void> {
    if (!appDir) { res.writeHead(404).end('The editor is not built: run `npm run build -w @docx4j/editor`, or start with --app.'); return; }
    const rel = normalize(decodeURIComponent(path === '/' ? '/index.html' : path)).replace(/^([/\\])+/, '');
    const file = join(appDir, rel);
    if (!file.startsWith(appDir + sep) && file !== appDir) { res.writeHead(403).end(); return; }
    try {
      const body = rel === 'index.html' ? markedPage(await readFile(file, 'utf8')) : await readFile(file);
      res.writeHead(200, { 'Content-Type': TYPES[extname(file)] ?? 'application/octet-stream', 'Cache-Control': 'no-store' });
      res.end(req.method === 'HEAD' ? undefined : body);
    } catch {
      res.writeHead(404).end();
    }
  }

  async function handle(req: IncomingMessage, res: ServerResponse): Promise<void> {
    // DNS rebinding: a page on another name that resolves here must not reach the agent's rooms.
    if (!hosts.has(req.headers.host ?? '')) { res.writeHead(421).end('Misdirected Request'); return; }
    const url = new URL(req.url ?? '/', base);
    const origin = req.headers.origin;
    if (url.pathname === '/companion' || url.pathname === '/companion/' || url.pathname === '/companion/rooms') {
      if (pageAllowed(origin)) res.setHeader('Access-Control-Allow-Origin', origin!);
      if (req.method === 'OPTIONS') { res.writeHead(204, { 'Access-Control-Allow-Methods': 'GET, POST', 'Access-Control-Allow-Headers': 'Content-Type' }).end(); return; }
      if (url.pathname === '/companion/rooms') {
        // An invitation only from this companion's pages, another page on this machine, or an origin allowed, as the relay allows.
        if (req.method !== 'POST') { res.writeHead(405).end(); return; }
        if (origin !== undefined && !pageAllowed(origin)) { res.writeHead(403).end('Forbidden'); return; }
        await serveInvite(req, res);
        return;
      }
      res.writeHead(200, { 'Content-Type': 'application/json' }).end(JSON.stringify(probe()));
      return;
    }
    if (url.pathname === '/mcp') { await serveMcp(req, res); return; }
    if (req.method !== 'GET' && req.method !== 'HEAD') { res.writeHead(405).end(); return; }
    await serveStatic(req, res, url.pathname);
  }

  /**
   * The relay in its worker (`relay-worker.mts`): its WebSocket address once it listens. A default
   * port that is taken falls back to a free one; a port the caller named does not.
   */
  async function startRelay(companionPort: number): Promise<string> {
    const named = options.relayPort !== undefined;
    let started = await startWorker(options.relayPort ?? (companionPort === 0 ? 0 : companionPort + 1));
    if (!started.url && !named && started.code === 'EADDRINUSE') {
      log(`docx4j-editor-mcp: port ${companionPort + 1} is taken; the relay takes a free one`);
      started = await startWorker(0);
    }
    if (!started.url) throw new Error(`the relay did not start: ${started.message}`);
    watch(started.worker);
    return started.url;
  }

  /** A relay worker on a port: its address once it listens, or why it did not. */
  function startWorker(port: number): Promise<{ worker: Worker; url?: string; code?: string; message?: string }> {
    return new Promise((resolve) => {
      const worker = new Worker(RELAY_WORKER, { workerData: { data: options.data ?? './.docx4j-editor-rooms', port, host, allowOrigins: [...allowed] } });
      const failed = (e: Error): void => resolve({ worker, message: e.message });
      worker.once('error', failed);
      worker.once('message', (m: { type: string; url?: string; code?: string; message?: string }) => {
        worker.off('error', failed);
        if (m.type === 'listening') resolve({ worker, url: m.url! });
        else { void worker.terminate(); resolve({ worker, ...(m.code ? { code: m.code } : {}), message: m.message ?? 'the relay did not start' }); }
      });
    });
  }

  /**
   * The running relay watched (the third review point of ED-005, section 12.47): a worker that stops
   * other than by `stopRelay` is restarted on its own port, so that the room links and every session's
   * connection, which reconnects by itself, reach it again, and its rooms come back from its files.
   * Restarts that fail, or more than five stops in a minute, leave it down: logged, and the tools
   * answer `relay.down` until the companion is restarted.
   */
  function watch(worker: Worker): void {
    relayWorker = worker;
    setRelayState('up');
    worker.on('error', (e) => { if (relayWorker === worker) log(`docx4j-editor-mcp: the relay failed: ${e.message}`); });
    worker.once('exit', (code) => {
      if (relayWorker !== worker) return;   // stopped by stopRelay, or replaced
      relayWorker = undefined;
      const now = Date.now();
      relayStops.push(now);
      while (relayStops.length > 0 && now - relayStops[0]! > 60_000) relayStops.shift();
      if (relayStops.length > 5) {
        setRelayState('down');
        log('docx4j-editor-mcp: the relay stopped five times in a minute and is left down; restart the companion');
        return;
      }
      log(`docx4j-editor-mcp: the relay stopped (exit code ${code}); restarting it on its port`);
      setRelayState('restarting');
      relayRestarted = restartRelay();
    });
  }

  async function restartRelay(): Promise<void> {
    const port = Number(new URL(relayUrl).port);
    let reason = '';
    for (const delay of [100, 500, 2000]) {
      await new Promise((resolve) => setTimeout(resolve, delay));
      if (relayState !== 'restarting') return;   // the companion closed meanwhile
      const started = await startWorker(port);
      if (started.url) {
        watch(started.worker);
        log(`docx4j-editor-mcp: the relay is running again at ${relayUrl}`);
        return;
      }
      reason = started.message ?? '';
    }
    setRelayState('down');
    log(`docx4j-editor-mcp: the relay could not be restarted (${reason}); restart the companion`);
  }

  /** The relay's state, told to `onRelayState` when it changes; not at a close, which the caller asked for. */
  function setRelayState(state: 'up' | 'restarting' | 'down'): void {
    const was = relayState;
    relayState = state;
    if (state !== was) options.onRelayState?.(state);
  }

  /** Whether the relay is running, waiting for a restart in progress. */
  async function relayReady(): Promise<boolean> {
    if (relayState === 'restarting') await relayRestarted;
    return relayState === 'up';
  }

  function relayDown(): ToolError {
    return new ToolError('relay.down', 'the companion\'s relay stopped and could not be restarted, so nothing reaches the rooms', 'tell the person the companion needs restarting; nothing you do now would reach the document');
  }

  /** Stop the relay: its rooms written, its connections closed, the worker ended. */
  async function stopRelay(): Promise<void> {
    const worker = relayWorker;
    relayState = 'down';
    if (!worker) return;
    relayWorker = undefined;
    await new Promise<void>((resolve) => {
      const done = setTimeout(resolve, 5000);
      worker.once('message', (m: { type?: string }) => { if (m?.type === 'closed') { clearTimeout(done); resolve(); } });
      worker.postMessage({ type: 'close' });
    });
    await worker.terminate();
  }

  /** The agent's connection to a room on the relay, as a page's: `create` makes the room (Share's path). */
  function connect(store: { readonly ydoc: import('yjs').Doc }, id: string, token: string, create: boolean): RoomConnection {
    return new RoomConnection(store, { room: id, token, relay: relayUrl }, { WebSocketPolyfill: WebSocket, create, user: { name, color: AGENT_COLOR, agent: true } });
  }

  /**
   * The agent in a room on the relay, through its connection (the session's store bound to the
   * provider, as a page's is); the agent in the room's awareness, as the connection's user; and a
   * person's Stop (the room's `agents` map, ED-005 section 12.28) heard, after which the companion
   * leaves and refuses the agent's calls there.
   */
  /**
   * The Stops the companion has heard, by room (docs/developer/agent.md section 2.11): each covers its room's
   * file and every file the agent saved the room to, so that `open` cannot take the agent round a
   * Stop into a new room of the same document. A person's invitation back into the room clears it;
   * leaving the room (`close`) does not.
   */
  const stops = new Map<string, { path: string | undefined; saved: Set<string> }>();

  /**
   * A session on a room joined afresh, as a page's join makes one: a new store synced from the relay
   * through a new connection, and the session opened on it. An invitation joins this way, and so
   * does a renewal after a refused call. The relay's refusals are the invitation's 404 and 403.
   */
  async function joinSession(id: string, token: string): Promise<{ session: EditorDocument; connection: RoomConnection }> {
    const store = newEditStore();
    const connection = connect(store, id, token, false);
    try {
      await connection.synced;
    } catch (e) {
      if (e instanceof RoomRefusedError) throw new InviteError(e.status === 'unknown' ? 404 : 403, e.message);
      throw e;
    }
    const baseFile = store.room.get('base') as Uint8Array | undefined;
    if (!baseFile) { connection.close(); throw new InviteError(409, 'the room holds no document yet'); }
    const session = await EditorDocument.open(baseFile, { store, mode: 'join', repair: 'word' });
    session.trackingAuthor = name;
    return { session, connection };
  }

  function enter(id: string, token: string, first: { session: EditorDocument; connection: RoomConnection }, path: string | undefined, saved: Iterable<string> = []): AgentRoom {
    // The agent in the room's awareness, as a page's user is (`app/presence.ts`): its name and
    // colour, marked as an agent, set by the connection (`RoomConnectionOptions.user`).
    const link = `${base}/#${new URLSearchParams({ room: id, token, relay: relayUrl }).toString()}`;
    let current = first;
    let chain: Promise<unknown> = Promise.resolve();
    let left = false;
    let offStop: () => void = () => undefined;
    /** The session and its connection let go: the Stop no longer heard, the agent gone from the awareness, the connection closed. */
    const detach = (): void => {
      offStop();
      current.connection.awareness.setLocalState(null);
      current.connection.close();
    };
    /** Out of the room. */
    const leave = async (): Promise<void> => {
      if (left) return;
      left = true;
      detach();
    };
    const room: AgentRoom = {
      id, path, link,
      get session(): EditorDocument { return current.session; },
      stopped: false,
      saved: new Set(saved),
      queue: <T,>(work: () => Promise<T>): Promise<T> => {
        const run = chain.then(work, work);
        chain = run.catch(() => undefined);
        return run;
      },
      people: () => [...current.connection.awareness.getStates().values()]
        .map((state) => (state as { user?: { name?: string; agent?: boolean } }).user)
        .filter((user): user is { name?: string; agent?: boolean } => user !== undefined)
        .map((user) => ({ name: user.name ?? '', agent: user.agent === true })),
      // A refused call (ED-005 proposal 42) leaves the session's tree holding it, so the agent lets
      // that session go and joins the room afresh: the room's state is the truth, and the refused
      // call wrote nothing to it.
      renew: async () => {
        detach();
        current = await joinSession(id, token);
        hear();
        log(`docx4j-editor-mcp: a call refused in room ${id}; joined it afresh`);
      },
      close: async () => {
        await leave();
        rooms.delete(id);
      },
    };
    /** A person's Stop: out of the room, the room kept so that every call on it answers agent.stopped (docs/developer/agent.md section 2.11). */
    const hear = (): void => {
      const store = current.session.store;
      offStop = onAgents(store, () => {
        if (room.stopped || !isAgentStopped(store, name)) return;
        room.stopped = true;
        stops.set(id, { path: room.path, saved: room.saved });
        log(`docx4j-editor-mcp: stopped in room ${id} by a person in it`);
        void leave();
      });
    };
    hear();
    rooms.set(id, room);
    return room;
  }

  async function open(path: string): Promise<AgentRoom> {
    if (!server) throw new Error('the companion is not listening');
    if (!(await relayReady())) throw relayDown();
    const file = await existingUnderRoots(roots, path);
    // A room of the file the person has not stopped the agent in answers; else a Stop that covers the file refuses.
    for (const room of rooms.values()) if (room.path === file && !room.stopped) return room;
    for (const [id, stop] of stops) {
      if (stop.path === file || stop.saved.has(file)) {
        const what = stop.path === file ? 'its file' : 'a copy you saved from it';
        throw new ToolError('agent.stopped', `the person pressed Stop for you in this document; ${basename(file)} is ${what}`, STOPPED_HINT, { document: id, path: file });
      }
    }
    const bytes = new Uint8Array(await readFile(file));
    // A file Word opens only through a repair is read as Word reads it, as the editor's Open reads it (Jason, 2026-11-10).
    const session = await EditorDocument.open(bytes, { repair: 'word' });
    session.trackingAuthor = name;
    const id = randomBytes(16).toString('base64url');
    const token = randomBytes(32).toString('base64url');
    await session.share(bytes, basename(file));
    const connection = connect(session.store, id, token, true);
    await connection.synced;
    const room = enter(id, token, { session, connection }, file);
    log(`docx4j-editor-mcp: ${basename(file)} is open in room ${id}: ${room.link}`);
    return room;
  }

  /**
   * A person's invitation (ED-005 section 8.2, "the browser first"): the agent joins a room on this
   * relay that a page shared, with the room's token, its session made from the room as a page's
   * join makes one. An agent stopped there before comes back: the invitation clears its Stop, and the
   * room keeps the file it was opened from and the files it was saved to.
   */
  async function joinRoom(id: string, token: string): Promise<AgentRoom> {
    if (!server) throw new Error('the companion is not listening');
    if (!(await relayReady())) throw new InviteError(503, 'the companion\'s relay stopped and could not be restarted');
    const existing = rooms.get(id);
    if (existing && !existing.stopped) return existing;
    if (existing) rooms.delete(id);
    const stop = stops.get(id);
    const joined = await joinSession(id, token);
    const store = joined.session.store;
    resumeAgent(store, name);
    stops.delete(id);
    const room = enter(id, token, joined, existing?.path ?? stop?.path, existing?.saved ?? stop?.saved);
    log(`docx4j-editor-mcp: invited into room ${id} (${String(store.room.get('name') ?? 'a shared document')})`);
    return room;
  }

  /** `POST <companion>/rooms` with `{ room, token }`: the invitation, from a page on this machine. */
  async function serveInvite(req: IncomingMessage, res: ServerResponse): Promise<void> {
    let body = '';
    for await (const chunk of req) body += String(chunk);
    let args: { room?: unknown; token?: unknown };
    try { args = JSON.parse(body) as typeof args; } catch { res.writeHead(400).end('not JSON'); return; }
    if (typeof args.room !== 'string' || typeof args.token !== 'string') { res.writeHead(400).end('room and token are needed'); return; }
    try {
      const room = await joinRoom(args.room, args.token);
      res.writeHead(200, { 'Content-Type': 'application/json' }).end(JSON.stringify({ document: room.id, agent: name }));
    } catch (e) {
      const { status, error } = invitationRefusal(e);
      res.writeHead(status, { 'Content-Type': 'application/json' }).end(JSON.stringify({ error }));
    }
  }

  return {
    get relay(): string { return relayUrl; },
    get relayState() { return relayState; },
    async crashRelay(): Promise<void> { await relayWorker?.terminate(); },
    rooms,
    host: toolHost,
    mcp: () => createMcpServer(toolHost),
    async listen(): Promise<string> {
      appDir = await findApp(options.app);
      if (!appDir) log('docx4j-editor-mcp: no built editor found; the room links will not open until it is built (npm run build -w @docx4j/editor)');
      server = createServer((req, res) => { handle(req, res).catch((e: unknown) => { log(`docx4j-editor-mcp: ${String(e)}`); if (!res.headersSent) res.writeHead(500); res.end(); }); });
      // No WebSocket is served here: the relay has a port of its own.
      server.on('upgrade', (_req, socket) => { socket.destroy(); });
      await new Promise<void>((resolve, reject) => { server!.once('error', reject); server!.listen(options.port ?? 4174, host, () => resolve()); });
      const address = server.address();
      const port = typeof address === 'object' && address ? address.port : options.port ?? 4174;
      hosts = hostsFor(port);
      base = `http://${host === '::1' ? '[::1]' : host}:${port}`;
      relayUrl = await startRelay(options.port === 0 ? 0 : port);
      log(`docx4j-editor-mcp: the editor and the probe at ${base}, the relay at ${relayUrl}`);
      return base;
    },
    open,
    join: joinRoom,
    setRoots(next: readonly string[]): void { roots = [...next]; },
    allowOrigin(origin: string): void {
      if (allowed.has(origin)) return;
      allowed.add(origin);
      relayWorker?.postMessage({ type: 'allowOrigin', origin });
    },
    async close(): Promise<void> {
      for (const room of [...rooms.values()]) await room.close();
      await stopRelay();
      await new Promise<void>((resolve) => { if (server) server.close(() => resolve()); else resolve(); });
    },
  };
}
