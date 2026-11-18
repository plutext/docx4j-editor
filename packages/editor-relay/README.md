# @docx4j/editor-relay

The co-editing relay for [`@docx4j/editor`](../editor): a [Hocuspocus](https://tiptap.dev/docs/hocuspocus)
server holding each shared room's Yjs document, with this repository's configuration. Browsers
connect to it through the editor's Share and Join; an in-process peer (the MCP companion's
agent) connects directly.

```
npx docx4j-editor-relay --port 4175 --data ./rooms
```

| Option | Default | |
|---|---|---|
| `--port` | 4175 | the port to listen on |
| `--host` | `127.0.0.1` | the address to bind: this machine only unless changed |
| `--data` | `./rooms` | the folder of room files, `<room>.bin` (the state) and `<room>.json` (the token's hash, when made, when last changed) |
| `--max-room` | 209715200 | a room's largest size in bytes; past it the connection is closed with the reason `Room Full` |
| `--idle-days` | 30 | a room nobody has changed for this long, and nobody has open, is deleted |
| `--origin` | local pages | a page origin allowed to connect, repeatable; `*` for any. Without it, pages on `localhost`, `127.0.0.1` or `[::1]`; a client that sends no `Origin` (Node) is always allowed |

**Rooms and tokens.** A room is made only by a connection that asks to make it, sending
`create:<token>` (the editor's Share): its token becomes the room's, and only its SHA-256 is kept.
A plain token joins a room that exists; for a room the relay does not hold (a mistyped link, a room
the idle rule deleted) it is refused with the reason `Unknown Room`, and nothing is made.
Every connection must bring the room's token; `readonly:<token>` connects read-only. Room ids are
16 to 128 URL-safe characters (the editor makes 128-bit ones).

**TLS.** The token travels in the socket's first message, so a relay reachable from anywhere but
this machine must sit behind a TLS proxy (nginx, Caddy, a cloud load balancer) and be given to the
editor as a `wss:` address. The plain `ws:` case is the local one only: a page served over
`https:` cannot open a `ws:` relay anyway.

**What the relay can read.** Everything shared through it: the relay holds each room's document
and the base file in it. Run it where the people sharing trust whoever runs it.

**Embedding.**

```ts
import { createRelay } from '@docx4j/editor-relay';

const relay = createRelay({ data: './rooms' });
await relay.listen();                           // or relay.handleUpgrade(request, socket, head) on your own server
const agent = await relay.openRoom(room, token); // a direct connection, checked against the room's token ({ create: true } to make it)
await agent.transact((doc) => { /* write as a peer */ });
```

Design: [`docs/user/co-editing.md`](https://github.com/plutext/docx4j-editor/blob/main/docs/user/co-editing.md) for people sharing a document, and
[`docs/developer/architecture.md`](https://github.com/plutext/docx4j-editor/blob/main/docs/developer/architecture.md) for the layers a room replicates.

Licence: `GPL-3.0-or-later` (see `LICENSE`); contributions under Apache-2.0 (`CONTRIBUTING.md`).
