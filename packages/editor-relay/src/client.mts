// SPDX-License-Identifier: GPL-3.0-or-later
//
// A peer's connection to a room on the relay (ED-004 sections 4.3 and 4.4): the provider that keeps
// a Yjs document in step with the relay's copy of the room, and what the relay's refusals mean.
// The application's pages use it, and so does the MCP companion, whose sessions join their rooms
// over loopback as a page does (ED-005 section 9.3 item 1 (i)). It is the package's `./client`
// entry, apart from the server, so that a browser bundle takes the provider and not the server.
// The provider binds the given `Y.Doc` itself, so the caller's store and the provider share one
// document and nothing is copied; its transactions carry the provider as their origin.
import { HocuspocusProvider, HocuspocusProviderWebsocket } from '@hocuspocus/provider';
import type * as Y from 'yjs';

/** A room on a relay: its id, its token, and the relay's WebSocket address. */
export interface RoomAddress {
  room: string;
  token: string;
  relay: string;
}

/** Who a peer is, for the others' carets and people lists. */
export interface RoomUser {
  name: string;
  /** `#rrggbb`: y-prosemirror draws only that form. */
  color: string;
  /** An agent, badged in the people list (ED-005 section 12.28). */
  agent?: boolean;
}

/**
 * Where the connection is: `connecting` (the first attempt, or again after a drop), `connected`
 * (synced with the relay), `offline` (the socket closed; edits go on locally and are sent when it
 * reconnects), `refused` (a wrong token, or an origin the relay does not allow; no retry),
 * `unknown` (the relay holds no such room: a mistyped link, or a room its idle rule deleted; no
 * retry), `full` (the room is at the relay's size limit; no retry), `closed` (left).
 *
 * `offline` comes from the socket closing or from the browser's own `offline` event. A relay that
 * dies while the network stays up is noticed only when the provider's message timeout (30 s)
 * passes without a message, so for up to that long the status still says `connected`; edits made
 * meanwhile are kept and sent on reconnect. That is the provider's design, not a fault.
 */
export type RoomStatus = 'connecting' | 'connected' | 'offline' | 'refused' | 'unknown' | 'full' | 'closed';

/** The relay's close reasons (`RelayError`), which stop the retries. */
const REFUSED = 'Forbidden';
const FULL = 'Room Full';
const UNKNOWN = 'Unknown Room';

export class RoomRefusedError extends Error {
  override name = 'RoomRefusedError';
  constructor(readonly status: 'refused' | 'unknown' | 'full', reason: string) {
    super(reason);
  }
}

export interface RoomConnectionOptions {
  /** Told of every change of status. */
  onStatus?: (status: RoomStatus) => void;
  /** A WebSocket class, for Node (the companion, the tests); the browser's own otherwise. */
  WebSocketPolyfill?: unknown;
  /**
   * Make the room if the relay does not hold it (`create:<token>`): Share, and a sharer's recovery
   * (which remakes a room the idle rule deleted from the local copy). A join never makes a room.
   */
  create?: boolean;
  /** Who this peer is, for the others' carets and people list (ED-004 section 5). */
  user?: RoomUser;
}

export class RoomConnection {
  readonly provider: HocuspocusProvider;
  private readonly socket: HocuspocusProviderWebsocket;
  private current: RoomStatus = 'connecting';
  /** Resolves once the store has synced with the relay the first time; rejects if the relay refuses the room. */
  readonly synced: Promise<void>;

  constructor(readonly store: { readonly ydoc: Y.Doc }, readonly link: RoomAddress, private readonly options: RoomConnectionOptions = {}) {
    let resolve!: () => void;
    let reject!: (e: Error) => void;
    this.synced = new Promise((res, rej) => { resolve = res; reject = rej; });
    // A caller that only awaits later must not see an unhandled rejection meanwhile.
    this.synced.catch(() => undefined);
    this.socket = new HocuspocusProviderWebsocket({
      url: link.relay,
      ...(options.WebSocketPolyfill ? { WebSocketPolyfill: options.WebSocketPolyfill } : {}),
    });
    const stop = (status: 'refused' | 'unknown' | 'full', reason: string): void => {
      if (this.current === 'closed') return;
      this.close();
      this.set(status);
      reject(new RoomRefusedError(status, reason));
    };
    this.provider = new HocuspocusProvider({
      websocketProvider: this.socket,
      name: link.room,
      document: store.ydoc,
      token: options.create ? `create:${link.token}` : link.token,
      onSynced: ({ state }) => { if (state) { this.set('connected'); resolve(); } },
      onStatus: ({ status }) => {
        if (this.current === 'refused' || this.current === 'unknown' || this.current === 'full' || this.current === 'closed') return;
        if (status === 'connecting') this.set('connecting');
        else if (status === 'disconnected') this.set('offline');
        else if (status === 'connected' && this.provider?.isSynced) this.set('connected');
      },
      onAuthenticationFailed: ({ reason }) => stop(reason === UNKNOWN ? 'unknown' : 'refused', reason),
      onClose: ({ event }) => {
        // The provider reconnects after any close; the relay's refusals stop it (ED-004 section 10.4 item 4).
        // 1009 is a frame past the socket's cap (a relay that caps frames at its room limit).
        if (event.reason === FULL || event.code === 1009) stop('full', event.reason || FULL);
        else if (event.reason === REFUSED) stop('refused', event.reason);
      },
    });
    this.provider.attach();
    if (options.user) this.setUser(options.user);
    // The browser's own word on the network: offline at once (a dead socket otherwise shows only
    // after the provider's message timeout, 30 s), and a new connection the moment it is back.
    if (typeof window !== 'undefined') {
      window.addEventListener('offline', this.wentOffline);
      window.addEventListener('online', this.cameOnline);
    }
  }

  private readonly wentOffline = (): void => {
    if (this.current === 'closed' || this.current === 'refused' || this.current === 'unknown' || this.current === 'full') return;
    this.socket.disconnect();
    this.set('offline');
  };

  private readonly cameOnline = (): void => {
    if (this.current !== 'offline') return;
    this.set('connecting');
    void this.socket.connect();
  };

  /** The awareness the peers' carets and the people list read. */
  get awareness(): NonNullable<HocuspocusProvider['awareness']> {
    return this.provider.awareness!;
  }

  /** Tell the others who this peer is (a name changed in Options). */
  setUser(user: RoomUser): void {
    this.provider.awareness?.setLocalStateField('user', user);
  }

  get status(): RoomStatus {
    return this.current;
  }

  private set(status: RoomStatus): void {
    if (this.current === status) return;
    this.current = status;
    this.options.onStatus?.(status);
  }

  /** Leave the room: the provider and its socket closed for good. The store and its local copy stay. */
  close(): void {
    if (this.current === 'closed') return;
    this.current = 'closed';
    if (typeof window !== 'undefined') {
      window.removeEventListener('offline', this.wentOffline);
      window.removeEventListener('online', this.cameOnline);
    }
    this.provider.destroy();
    this.socket.destroy();
    // UPSTREAM-6: `destroy` leaves the reconnect `onClose` scheduled, and `connect` revives the socket.
    this.socket.connect = async () => undefined;
    this.options.onStatus?.('closed');
  }
}
