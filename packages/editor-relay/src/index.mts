// SPDX-License-Identifier: GPL-3.0-or-later
//
// The co-editing relay (ED-004 section 4.2): a Hocuspocus server holding each room's Yjs
// document, with this repository's configuration. A room is made only by a connection that asks
// to make it, `create:<token>` (the editor's Share), whose token becomes the room's (only its
// SHA-256 is kept); a plain token joins a room that exists and is refused with `Unknown Room`
// otherwise, so a mistyped link or a room the idle rule deleted is said as such and never made by
// a join. Every connection must bring the room's token, and may ask to be read-only by sending
// `readonly:<token>`. Each room's state
// is a plain file (`<data>/<room>.bin`) beside its record (`<data>/<room>.json`: the token's hash,
// when it was made, when it was last written); no native module. A room's size is capped, and a
// room nobody has written to for `idleDays` is deleted. Local by default: `127.0.0.1`, and only
// local pages (or clients that send no `Origin`, such as Node) unless origins are named.
//
// Embedding (E4's companion): `createRelay(options)` gives the server; `listen()` starts its own
// HTTP server, `handleUpgrade` mounts it on someone else's (Vite's, under `/relay`), and
// `openRoom(id, token)` is the in-process peer's way in, a direct connection checked against the
// same token.
import { createHash, timingSafeEqual } from 'node:crypto';
import { createServer, type IncomingMessage, type Server as HttpServer } from 'node:http';
import type { Duplex } from 'node:stream';
import { mkdir, readFile, readdir, rename, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { Hocuspocus, type DirectConnection } from '@hocuspocus/server';
import { Database } from '@hocuspocus/extension-database';
import { WebSocketServer } from 'ws';
import * as Y from 'yjs';

export interface RelayOptions {
  /** The port `listen` binds (default 4175; 0 picks a free one). */
  port?: number;
  /** The address `listen` binds (default `127.0.0.1`: local only). */
  host?: string;
  /** The folder of room files (default `./rooms`). */
  data?: string;
  /** A room's largest state in bytes; an update past it is refused (default 200 MB). */
  maxRoom?: number;
  /** Days without a write after which a room is deleted (default 30). */
  idleDays?: number;
  /**
   * The `Origin` headers allowed to connect. Undefined: pages on `localhost`, `127.0.0.1` or
   * `[::1]`, at any port. `['*']`: any. A connection without an `Origin` (not a browser) is allowed.
   */
  origins?: readonly string[];
  /** Milliseconds between a change and its write to disk (Hocuspocus's debounce; default 2000). */
  debounce?: number;
  /** The clock, for the idle rule's tests. */
  now?: () => number;
}

/** What a room's record holds, beside its state. */
interface RoomRecord {
  tokenHash: string;
  created: string;
  touched: string;
}

/** Room ids are URL-safe and long enough to be unguessable; anything else is refused before it names a file. */
const ROOM_ID = /^[A-Za-z0-9_-]{16,128}$/;
const READ_ONLY = 'readonly:';
const DAY = 24 * 60 * 60 * 1000;

/**
 * A refusal. Hocuspocus closes the connection with the error's `code` and `reason`, which the
 * application reads to say why and to stop reconnecting (the provider otherwise retries).
 */
export class RelayError extends Error {
  override name = 'RelayError';
  constructor(message: string, readonly code = 4403, readonly reason = 'Forbidden') {
    super(message);
  }
}

/** The close code of a room past its size limit. */
export const ROOM_FULL = { code: 4413, reason: 'Room Full' } as const;

/** The refusal of a join to a room the relay does not hold (never made, or deleted by the idle rule). */
export const UNKNOWN_ROOM = { code: 4404, reason: 'Unknown Room' } as const;

const CREATE = 'create:';

export interface Relay {
  /** The Hocuspocus instance (its documents, connections and hooks). */
  readonly hocuspocus: Hocuspocus;
  /** Start an HTTP server of the relay's own; resolves with its WebSocket URL. */
  listen(): Promise<string>;
  /** Hand an HTTP upgrade to the relay (for a server that is not the relay's own). */
  handleUpgrade(request: IncomingMessage, socket: Duplex, head: Buffer): void;
  /** An in-process peer's connection to a room, under the room's token; `create` makes the room when the relay does not hold it. */
  openRoom(room: string, token: string, options?: { create?: boolean }): Promise<DirectConnection>;
  /** Delete the rooms idle for longer than `idleDays` that nobody has open; their ids. */
  sweep(): Promise<string[]>;
  /** Close every connection, write every open room, stop the server. */
  destroy(): Promise<void>;
}

export function createRelay(options: RelayOptions = {}): Relay {
  const data = options.data ?? './rooms';
  const maxRoom = options.maxRoom ?? 200 * 1024 * 1024;
  const idleDays = options.idleDays ?? 30;
  const now = options.now ?? Date.now;
  const origins = options.origins;
  /** Each open room's size in bytes, an upper bound kept as updates arrive. */
  const sizes = new Map<string, number>();
  /** When each open room last changed: what a write records as `touched` (a write at unload with no change is not a touch). */
  const changed = new Map<string, number>();
  /** Claims in progress, so two first connections cannot both claim a room. */
  const claims = new Map<string, Promise<RoomRecord>>();
  /** Once destroyed, a store a debounce still had pending writes nothing; the writes under way are awaited by `destroy`. */
  let destroyed = false;
  const writes = new Set<Promise<void>>();

  const file = (room: string, ext: string): string => join(data, `${room}.${ext}`);
  const hash = (token: string): string => createHash('sha256').update(token).digest('hex');

  async function readRecord(room: string): Promise<RoomRecord | undefined> {
    try {
      return JSON.parse(await readFile(file(room, 'json'), 'utf8')) as RoomRecord;
    } catch {
      return undefined;
    }
  }

  async function writeAtomically(path: string, content: string | Uint8Array): Promise<void> {
    await mkdir(data, { recursive: true });
    const temporary = `${path}.${process.pid}.tmp`;
    await writeFile(temporary, content);
    await rename(temporary, path);
  }

  /** The room's record, made with this token's hash when the room is new and `create` asks for it; a different token is refused. */
  async function authorise(room: string, token: string, create: boolean): Promise<RoomRecord> {
    if (!ROOM_ID.test(room)) throw new RelayError('not a room id');
    if (token === '') throw new RelayError('no token');
    const pending = claims.get(room);
    if (pending) await pending.catch(() => undefined);
    let record = await readRecord(room);
    if (!record) {
      if (!create) throw new RelayError('no such room', UNKNOWN_ROOM.code, UNKNOWN_ROOM.reason);
      const stamp = new Date(now()).toISOString();
      const made: RoomRecord = { tokenHash: hash(token), created: stamp, touched: stamp };
      const claim = writeAtomically(file(room, 'json'), JSON.stringify(made)).then(() => made);
      claims.set(room, claim);
      try { record = await claim; } finally { claims.delete(room); }
      return record;
    }
    const given = Buffer.from(hash(token), 'hex');
    const kept = Buffer.from(record.tokenHash, 'hex');
    if (given.length !== kept.length || !timingSafeEqual(given, kept)) throw new RelayError('wrong token for this room');
    return record;
  }

  function originAllowed(origin: string | undefined): boolean {
    if (origin === undefined) return true;
    if (origins?.includes('*')) return true;
    if (origins) return origins.includes(origin);
    try {
      const { hostname } = new URL(origin);
      return hostname === 'localhost' || hostname === '127.0.0.1' || hostname === '[::1]';
    } catch {
      return false;
    }
  }

  const hocuspocus = new Hocuspocus({
    quiet: true,
    debounce: options.debounce ?? 2000,
    extensions: [
      new Database({
        fetch: async ({ documentName }) => {
          try {
            const state = new Uint8Array(await readFile(file(documentName, 'bin')));
            sizes.set(documentName, state.length);
            return state;
          } catch {
            sizes.set(documentName, 0);
            return null;
          }
        },
        store: async ({ documentName, state }) => {
          if (destroyed) return;
          const write = (async () => {
            await writeAtomically(file(documentName, 'bin'), state);
            sizes.set(documentName, state.length);
            const at = changed.get(documentName);
            const record = await readRecord(documentName);
            if (record && at !== undefined) await writeAtomically(file(documentName, 'json'), JSON.stringify({ ...record, touched: new Date(at).toISOString() }));
          })();
          writes.add(write);
          try { await write; } finally { writes.delete(write); }
        },
      }),
    ],
    async onConnect({ requestHeaders }) {
      if (!originAllowed(requestHeaders.origin)) throw new RelayError('origin not allowed');
    },
    async onAuthenticate({ documentName, token, connectionConfig }) {
      const readOnly = token.startsWith(READ_ONLY);
      const create = token.startsWith(CREATE);
      const secret = readOnly ? token.slice(READ_ONLY.length) : create ? token.slice(CREATE.length) : token;
      await authorise(documentName, secret, create);
      if (readOnly) connectionConfig.readOnly = true;
    },
    async beforeHandleMessage({ documentName, update }) {
      // `update` is the whole incoming message; its length bounds what it can add.
      if ((sizes.get(documentName) ?? 0) + update.length > maxRoom) throw new RelayError(`room limit of ${maxRoom} bytes reached`, ROOM_FULL.code, ROOM_FULL.reason);
    },
    async onChange({ documentName, update }) {
      sizes.set(documentName, (sizes.get(documentName) ?? 0) + update.length);
      changed.set(documentName, now());
    },
    async afterUnloadDocument({ documentName }) {
      sizes.delete(documentName);
      changed.delete(documentName);
    },
  });

  // The frame cap sits above the room limit, so an update past the limit reaches
  // `beforeHandleMessage` and is refused with the reason `Room Full` rather than by `ws` with 1009.
  const sockets = new WebSocketServer({ noServer: true, maxPayload: maxRoom + 1024 * 1024 });
  sockets.on('connection', (socket, request) => {
    // A frame past `maxPayload` makes `ws` close the socket with an error event; the refusal is the close.
    socket.on('error', () => undefined);
    // A cast (ED-004 section 10.6 item 3): Hocuspocus types the socket as the DOM's WebSocket; `ws`'s is what it uses.
    hocuspocus.handleConnection(socket as never, request);
  });
  let http: HttpServer | undefined;

  const handleUpgrade = (request: IncomingMessage, socket: Duplex, head: Buffer): void => {
    sockets.handleUpgrade(request, socket, head, (ws) => sockets.emit('connection', ws, request));
  };

  async function sweep(): Promise<string[]> {
    let names: string[];
    try { names = await readdir(data); } catch { return []; }
    const gone: string[] = [];
    for (const name of names) {
      if (!name.endsWith('.json')) continue;
      const room = name.slice(0, -'.json'.length);
      if (hocuspocus.documents.has(room)) continue;
      const record = await readRecord(room);
      if (!record || now() - Date.parse(record.touched) <= idleDays * DAY) continue;
      await rm(file(room, 'bin'), { force: true });
      await rm(file(room, 'json'), { force: true });
      gone.push(room);
    }
    return gone;
  }
  const sweeper = setInterval(() => { void sweep(); }, 60 * 60 * 1000);
  sweeper.unref();

  return {
    hocuspocus,
    handleUpgrade,
    listen: () => new Promise((resolve, reject) => {
      http = createServer((_request, response) => { response.writeHead(426, { 'content-type': 'text/plain' }).end('docx4j-editor-relay: connect with a WebSocket\n'); });
      http.on('upgrade', handleUpgrade);
      http.once('error', reject);
      http.listen(options.port ?? 4175, options.host ?? '127.0.0.1', () => {
        const address = http!.address();
        const port = typeof address === 'object' && address ? address.port : options.port;
        void sweep();
        resolve(`ws://${options.host ?? '127.0.0.1'}:${port}`);
      });
    }),
    async openRoom(room: string, token: string, options: { create?: boolean } = {}): Promise<DirectConnection> {
      await authorise(room, token, options.create === true);
      return hocuspocus.openDirectConnection(room);
    },
    sweep,
    async destroy(): Promise<void> {
      clearInterval(sweeper);
      for (const document of [...hocuspocus.documents.values()]) {
        // A hand-built payload cast to the hook's type (ED-004 section 10.6 item 3): the first place a Hocuspocus minor may break.
        await hocuspocus.storeDocumentHooks(document, { clientsCount: 0, context: {}, document, documentName: document.name, instance: hocuspocus, requestHeaders: {}, requestParameters: new URLSearchParams(), socketId: 'relay' } as never, true);
      }
      // Every room is stored above. A debounced store that fires later writes nothing (its data folder may be
      // gone, as a test's is), and one under way is let finish.
      destroyed = true;
      await Promise.allSettled([...writes]);
      hocuspocus.closeConnections();
      for (const client of sockets.clients) client.terminate();
      sockets.close();
      if (http) await new Promise<void>((resolve) => http!.close(() => resolve()));
    },
  };
}

/** The size of a Yjs update or state, for callers that report it. */
export function stateSize(doc: Y.Doc): number {
  return Y.encodeStateAsUpdate(doc).length;
}
