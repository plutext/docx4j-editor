// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Ask box's one call to the Anthropic API (ED-014 section 4.6, P45): made by the page in a browser and by the VS Code
// extension, which takes this file through its `shared.ts`, so that both hosts send the same body. It imports the SDK and
// the protocol's shapes alone.
//
// Nothing comes from the environment (section 9.20, finding 2): the SDK reads `ANTHROPIC_BASE_URL` for a base URL not
// given and `ANTHROPIC_AUTH_TOKEN` for an `authToken` not given, and in VS Code's extension host those are the
// environment VS Code was started from. So the address, the key and `authToken: null` are always passed.
import Anthropic from '@anthropic-ai/sdk';
import type { AskAnswerData, AskFailureData, AskRequestData } from './host-protocol.js';

/** Anthropic's own address: an empty base URL means it, passed as such. */
export const ANTHROPIC_ADDRESS = 'https://api.anthropic.com';

export interface SendOptions {
  apiKey: string;
  /** The API's address: Anthropic's when empty, else a proxy or a test double (`<origin>/ask`). */
  baseURL?: string;
  signal?: AbortSignal;
  /** The page's own call in a browser: the SDK refuses a browser otherwise, and says so in a header (section 9.20 item 9). */
  browser?: boolean;
}

/** The request as one message, for a chat or a model with no system role: the contract and the context, a rule, the question (ED-003 section 11.25). */
export function chatPrompt(request: AskRequestData): string {
  return `${request.system}\n\n---\n\n${request.prompt}`;
}

/** One Messages API call: the system prompt, one user message, the answer's limit and adaptive thinking. */
export async function sendAnthropic(request: AskRequestData, options: SendOptions): Promise<AskAnswerData> {
  const client = new Anthropic({
    apiKey: options.apiKey,
    authToken: null,
    baseURL: options.baseURL && options.baseURL.trim() !== '' ? options.baseURL.trim() : ANTHROPIC_ADDRESS,
    // The key is the user's, in the user's own page (ED-001 section 11 item 5).
    ...(options.browser ? { dangerouslyAllowBrowser: true } : {}),
    maxRetries: 1,
  });
  const response = await client.messages.create({
    model: request.model,
    max_tokens: request.maxTokens,
    system: request.system,
    messages: [{ role: 'user', content: request.prompt }],
    thinking: { type: 'adaptive' },
  }, { ...(options.signal ? { signal: options.signal } : {}) });
  const text = response.content.filter((b): b is Anthropic.TextBlock => b.type === 'text').map((b) => b.text).join('\n');
  const answer: AskAnswerData = { text, model: response.model, stopReason: response.stop_reason, usage: { input: response.usage.input_tokens, output: response.usage.output_tokens } };
  if (response.stop_reason === 'refusal') {
    const details = response.stop_details;
    answer.refusal = { category: details?.category ?? null, explanation: details?.explanation ?? null };
  }
  return answer;
}

/** Whether an error is the SDK's for a request its signal cancelled. */
export function isCancelled(e: unknown): boolean {
  return e instanceof Anthropic.APIUserAbortError || (e as { name?: unknown } | null)?.name === 'AbortError';
}

/** A failed call as data, to word in the page or to post from the extension: its kind, its status and its message. */
export function failureOf(e: unknown): AskFailureData {
  const message = (e as Error | null)?.message ?? String(e);
  if (e instanceof Anthropic.AuthenticationError) return { kind: 'key', status: 401, message };
  if (e instanceof Anthropic.RateLimitError) return { kind: 'rateLimit', status: 429, message };
  if (e instanceof Anthropic.APIConnectionError) return { kind: 'connection', message };
  if (e instanceof Anthropic.APIError) return { kind: 'api', ...(typeof e.status === 'number' ? { status: e.status } : {}), message };
  return { kind: 'other', message };
}
