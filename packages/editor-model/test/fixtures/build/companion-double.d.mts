// SPDX-License-Identifier: GPL-3.0-or-later
import type { IncomingMessage, ServerResponse } from 'node:http';

/** The companion double's `paginate`: every page-break marker moved one paragraph down (ED-003 section 6.5). */
export function paginateDouble(bytes: Uint8Array): Promise<Uint8Array>;

/** The companion's contract on a Node request handler under `base` (`/companion`). */
export function companionRoutes(base?: string): (req: IncomingMessage, res: ServerResponse, next: () => void) => void;
