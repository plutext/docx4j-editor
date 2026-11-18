// SPDX-License-Identifier: GPL-3.0-or-later
//
// The tools' errors (docs/developer/agent.md section 2.8): an MCP tool error whose structured content is
// `{ code, message, hint, ... }`. The codes are dotted, as docx4j-python CR-003's are, so that an
// agent recovers the same way whichever docx4j server answered; where the meaning is CR-003's, so is
// the code.

/** Every code a tool answers with (docs/developer/agent.md section 2.8). */
export type ErrorCode =
  | 'address.not_found'
  | 'range.invalid'
  | 'target.invalid'
  | 'location.invalid'
  | 'style.not_found'
  | 'document.not_open'
  | 'target.stale'
  | 'markdown.unsupported'
  | 'ooxml.invalid'
  | 'arguments.invalid'
  | 'path.not_allowed'
  | 'path.not_found'
  | 'path.exists'
  | 'tracking.not_allowed'
  | 'resolve.not_allowed'
  | 'agent.stopped'
  | 'relay.down'
  | 'binding.no_node'
  | 'binding.document_element'
  | 'binding.no_repeat_node'
  | 'internal.error';

/** A failed call: what the agent reads to recover. */
export class ToolError extends Error {
  constructor(
    readonly code: ErrorCode,
    message: string,
    readonly hint: string,
    /** The code's own fields (`address`, `nearest`, `current`, ...). */
    readonly details: Record<string, unknown> = {},
  ) {
    super(message);
    this.name = 'ToolError';
  }
}

/** The result an MCP client gets for a failed call. */
export interface ErrorResult {
  isError: true;
  content: { type: 'text'; text: string }[];
  structuredContent: Record<string, unknown>;
  [key: string]: unknown;
}

/** A thrown value as the call's error result: a `ToolError` as it is, anything else as `internal.error`. */
export function errorResult(e: unknown): ErrorResult {
  const error = e instanceof ToolError
    ? e
    : new ToolError('internal.error', e instanceof Error ? e.message : String(e), 'this is the companion\'s fault, not the call\'s: report it');
  const structured = { code: error.code, message: error.message, hint: error.hint, ...error.details };
  return { isError: true, content: [{ type: 'text', text: `${error.code}: ${error.message}\nhint: ${error.hint}` }], structuredContent: structured };
}
