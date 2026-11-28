// SPDX-License-Identifier: GPL-3.0-or-later
//
// The channel between the companion and the process that started it (ED-014 section 4.5, P37): Node's IPC channel, which
// `--ipc` asks for and the spawn opens (`stdio: [..., 'ipc']`). The VS Code extension starts the companion with one. No
// imports, so the extension shares the file as it is (`packages/editor-vscode/src/shared.ts`).
//
// - **The companion posts** `listening` once it listens, and `relay` when its relay stops or is running again; `invited`
//   answers an `invite`.
// - **The starter posts** `allowOrigin` for a page origin learned after the start, `roots` when the workspace's folders
//   change, `invite` for a room a page shared, and `close`.
// - **When the channel closes,** the starter has gone (a quit, a reload, a crash): the companion writes its rooms and exits,
//   as it does at `close`. So no companion outlives the process that started it.

/** From the companion. */
export type FromCompanion =
  /** It listens: its own address (`http://127.0.0.1:<port>`) and its relay's (`ws://127.0.0.1:<port>`). */
  | { type: 'listening'; url: string; relay: string }
  /** The relay stopped and is restarting, is running again, or is down until the companion restarts. */
  | { type: 'relay'; state: 'up' | 'restarting' | 'down' }
  /** The agent joined the room an `invite` named. */
  | { type: 'invited'; id: number }
  /** It did not: the status `POST /companion/rooms` would answer (404 no such room, 403 a wrong token, 409 no document yet). */
  | { type: 'invited'; id: number; status: number; error: string };

/** To the companion. */
export type ToCompanion =
  /** One more page origin allowed, at the relay (and the probe and the invitation). */
  | { type: 'allowOrigin'; origin: string }
  /** The folders the agent may open files under, replacing the ones it had. */
  | { type: 'roots'; roots: string[] }
  /** The agent invited into a room on the companion's relay, as `POST /companion/rooms` invites it. */
  | { type: 'invite'; id: number; room: string; token: string }
  /** Write the rooms and exit. */
  | { type: 'close' };
