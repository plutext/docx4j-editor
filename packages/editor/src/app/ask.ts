// SPDX-License-Identifier: GPL-3.0-or-later
//
// The ask mode's provider adapter (ED-003 section 7.5; section 9 items 5 and 30): one interface,
// one implementation in E2.e, the Anthropic API through its SDK in the browser with the user's
// own key (bring your own, kept in local storage, entered in Options and written nowhere else)
// and a base URL an organisation proxy (E4) or the suites' test double can take. One request
// per question, no streaming, no server-side fallback to another model: the user chose the
// model and is shown the request; a refusal is shown as such with its explanation. The
// organisation-proxy and companion routes are E4's and reuse this interface.
import Anthropic from '@anthropic-ai/sdk';
import type { AskRequest } from './ask-context.js';
import { t } from './i18n.js';

export const OPTION_ASK_KEY = 'docx4j-editor.ask.key';
export const OPTION_ASK_BASE_URL = 'docx4j-editor.ask.baseUrl';
export const OPTION_ASK_MODEL = 'docx4j-editor.ask.model';
export const OPTION_ASK_SEND_AT_ONCE = 'docx4j-editor.ask.sendAtOnce';

export interface AskModel {
  id: string;
  label: string;
}

/** Section 9 item 5: the newest Claude models, Sonnet 5 by default for the box's latency. */
export const ASK_MODELS: readonly AskModel[] = [
  { id: 'claude-sonnet-5', label: 'Claude Sonnet 5' },
  { id: 'claude-opus-5', label: 'Claude Opus 5' },
  { id: 'claude-fable-5-1', label: 'Claude Fable 5.1' },
];

export interface AskAnswer {
  text: string;
  model: string;
  stopReason: string | null;
  /** The refusal's category and explanation when the model declined (`stop_reason: 'refusal'`). */
  refusal?: { category: string | null; explanation: string | null };
  usage?: { input: number; output: number };
}

export interface ProviderOptions {
  apiKey: string;
  /** The API's base URL: the provider's own when empty, else a proxy or a test double (`<origin>/ask`). */
  baseURL?: string;
  signal?: AbortSignal;
}

export interface AskProvider {
  id: string;
  name: string;
  models: readonly AskModel[];
  send(request: AskRequest, options: ProviderOptions): Promise<AskAnswer>;
}

export const anthropicProvider: AskProvider = {
  id: 'anthropic',
  name: 'Anthropic API',
  models: ASK_MODELS,
  async send(request, options) {
    const client = new Anthropic({
      apiKey: options.apiKey,
      ...(options.baseURL ? { baseURL: options.baseURL } : {}),
      dangerouslyAllowBrowser: true,   // the key is the user's, in the user's own page (ED-001 section 11 item 5)
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
    const answer: AskAnswer = { text, model: response.model, stopReason: response.stop_reason, usage: { input: response.usage.input_tokens, output: response.usage.output_tokens } };
    if (response.stop_reason === 'refusal') {
      const details = response.stop_details;
      answer.refusal = { category: details?.category ?? null, explanation: details?.explanation ?? null };
    }
    return answer;
  },
};

export function askOption(key: string): string {
  try { return (localStorage.getItem(key) ?? '').trim(); } catch { return ''; }
}

export function setAskOption(key: string, value: string): void {
  try { if (value === '') localStorage.removeItem(key); else localStorage.setItem(key, value); } catch { /* ignore */ }
}

/** A message for the status line from a failed request. */
export function describeError(e: unknown): string {
  if (e instanceof Anthropic.AuthenticationError) return t('ask.error.key');
  if (e instanceof Anthropic.RateLimitError) return t('ask.error.rateLimit');
  if (e instanceof Anthropic.APIConnectionError) return t('ask.error.connection', { error: e.message });
  if (e instanceof Anthropic.APIError) return t('ask.error.api', { status: e.status ?? '', error: e.message });
  return (e as Error).message ?? String(e);
}
