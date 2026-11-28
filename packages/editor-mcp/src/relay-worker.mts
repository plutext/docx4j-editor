// SPDX-License-Identifier: GPL-3.0-or-later
//
// The companion's relay, in a worker thread of its own (ED-005 section 9.3 item 1 (i)). A tool
// call's tree run holds the companion's event loop for its whole length (about 0.7 s on a 255-page
// document, section 12.31 item 3); with the relay on that loop, every room's messages waited for it,
// people's typing to each other included. In a worker the relay has an event loop of its own, so it
// forwards while a call runs, and it ends with the companion's process. The companion's sessions
// join its rooms over loopback, as a page does (`RoomConnection`).
//
// The protocol with the companion: the worker answers `{ type: 'listening', url }` or
// `{ type: 'error', message, code }` once, and `{ type: 'closed' }` after a `{ type: 'close' }`.
// `{ type: 'allowOrigin', origin }` allows one more page origin from then on (ED-014 section 4.5).
import { parentPort, workerData } from 'node:worker_threads';
import { createRelay } from '@docx4j/editor-relay';

export interface RelayWorkerData {
  data: string;
  port: number;
  host: string;
  /** Page origins allowed beside the local rule: a VS Code webview's (ED-014 section 4.5). */
  allowOrigins?: string[];
}

const port = parentPort!;
const { data, port: listenPort, host, allowOrigins } = workerData as RelayWorkerData;
const relay = createRelay({ data, port: listenPort, host, allowOrigins: allowOrigins ?? [] });

port.on('message', (message: { type?: string; origin?: unknown }) => {
  if (message?.type === 'allowOrigin' && typeof message.origin === 'string') { relay.allowOrigin(message.origin); return; }
  if (message?.type !== 'close') return;
  void relay.destroy().finally(() => port.postMessage({ type: 'closed' }));
});

try {
  const url = await relay.listen();
  port.postMessage({ type: 'listening', url });
} catch (e) {
  port.postMessage({ type: 'error', message: (e as Error).message, code: (e as NodeJS.ErrnoException).code });
}
