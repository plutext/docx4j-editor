// SPDX-License-Identifier: GPL-3.0-or-later
import type { IncomingMessage, ServerResponse } from 'node:http';

/** The ask double's answer text for a prompt (ED-003 section 7.5): a ts block, prose, a text block or a wml block by the prompt's words; null for a refusal. */
export function answerFor(prompt: string): string | null;

/** The Messages API's response for a request body. */
export function messageFor(body: unknown): unknown;

/** The route on a Node request handler under `base` (`/ask`). */
export function askRoutes(base?: string): (req: IncomingMessage, res: ServerResponse, next: () => void) => void;
