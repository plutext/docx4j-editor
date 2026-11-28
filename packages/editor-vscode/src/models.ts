// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Ask box's requests in the extension (ED-014 section 4.6, P44 to P46), with no `vscode` import: `ask.ts` gives it
// VS Code's calls. Every request the page sends goes through here, by one of two routes:
// - **VS Code's models** (`vscode`): the request as one user message (`chatPrompt`, since the stable API has no system
//   role), counted first against the model's input, sent with the `LanguageModelChat` object the last listing gave, never
//   selected again by its id (section 9.20 finding 1);
// - **the Anthropic API with the person's key** (`key`): the key from SecretStorage, the address from a setting the page
//   cannot write, the call the page makes in a browser (`sendAnthropic`).
// A request runs under its page: the page's going cancels it.
import { chatPrompt, failureOf, modelKey, type AskAnswerData, type AskFailureData, type AskRequestData, type HostModelData, type ToPage } from './shared.js';

/** A model as `vscode.lm` gives it (`LanguageModelChat`): its fields, and its two calls. */
export interface ChatModel {
  readonly id: string;
  readonly name: string;
  readonly vendor: string;
  readonly family: string;
  readonly version: string;
  readonly maxInputTokens: number;
  sendRequest(messages: unknown[], options?: { justification?: string; modelOptions?: Record<string, unknown> }, token?: unknown): PromiseLike<{ text: AsyncIterable<string> }>;
  countTokens(text: string, token?: unknown): PromiseLike<number>;
}

/** A model's data, as the page lists it. */
export function modelData(m: ChatModel): HostModelData {
  return { vendor: m.vendor, id: m.id, version: m.version, name: m.name, family: m.family, maxInputTokens: m.maxInputTokens };
}

/** A failure the page words, from a `LanguageModelError` by its `code` (both releases' `vscode.d.ts`), or any other error. */
export function lmFailure(e: unknown): AskFailureData {
  const message = (e as Error | null)?.message ?? String(e);
  switch ((e as { code?: unknown } | null)?.code) {
    case 'NoPermissions': return { kind: 'noPermissions', message };
    case 'Blocked': return { kind: 'blocked', message };
    case 'NotFound': return { kind: 'notFound', message };
    default: return { kind: 'other', message };
  }
}

/** A request with no answer, and why. */
export class AskFailed extends Error {
  constructor(readonly failure: AskFailureData) {
    super(failure.message);
  }
}

/** The answer's limit a page may ask for: the browser asks for 4,096 (ED-003 section 7.5). */
export const MAX_ANSWER_TOKENS = 32_000;

/** VS Code's side, as `ask.ts` gives it. */
export interface AskDeps<M extends ChatModel> {
  /** `vscode.lm.selectChatModels()`; none where VS Code lacks the API (P30). */
  select(): PromiseLike<readonly M[]>;
  /** `LanguageModelChatMessage.User(text)`. */
  userMessage(text: string): unknown;
  /** A VS Code cancellation token that follows the signal, and its disposal. */
  tokenFor(signal: AbortSignal): { token: unknown; dispose(): void };
  /** The consent dialog's justification, in VS Code's language. */
  justification(): string;
  /** The person's Anthropic key, from SecretStorage. */
  key(): PromiseLike<string | undefined>;
  /** The key's address, the setting `docx4j.ask.baseUrl` at application scope; empty for Anthropic's own. */
  address(): string;
  /** The browser's call, `sendAnthropic`. */
  sendKey(request: AskRequestData, options: { apiKey: string; baseURL: string; signal: AbortSignal }): Promise<AskAnswerData>;
}

export type Asked = Extract<ToPage, { type: 'asked' }>;

/** A request as the page posted it, checked: the page is the extension's own, but a script in it can post anything. */
export function checkedRequest(value: unknown): AskRequestData | undefined {
  const r = value as Partial<AskRequestData> | null;
  if (!r || typeof r !== 'object') return undefined;
  if (typeof r.model !== 'string' || r.model === '' || typeof r.system !== 'string' || typeof r.prompt !== 'string') return undefined;
  if (typeof r.maxTokens !== 'number' || !Number.isInteger(r.maxTokens) || r.maxTokens < 1) return undefined;
  return { model: r.model, system: r.system, prompt: r.prompt, maxTokens: Math.min(r.maxTokens, MAX_ANSWER_TOKENS) };
}

/** A promise that rejects when the signal aborts, and never settles otherwise. */
function abortOf(signal: AbortSignal): Promise<never> {
  const out = new Promise<never>((_, reject) => {
    const fail = (): void => reject(new AskFailed({ kind: 'other', message: 'cancelled' }));
    if (signal.aborted) fail(); else signal.addEventListener('abort', fail, { once: true });
  });
  out.catch(() => undefined);
  return out;
}

/** Every occurrence of the key taken out of a message: an error's words must never carry it to the page. */
function withoutKey(failure: AskFailureData, key: string | undefined): AskFailureData {
  if (!key || !failure.message.includes(key)) return failure;
  return { ...failure, message: failure.message.split(key).join('…') };
}

/**
 * The requests of every page, by the page's channel `P`: VS Code's models as last listed, kept by `<vendor>/<id>`, and each
 * running request's cancellation, which the page's Cancel, its going, or its loading again ends.
 */
export class AskService<M extends ChatModel, P = unknown> {
  private kept = new Map<string, M>();
  private listedOnce = false;
  private readonly running = new Map<P, Map<number, AbortController>>();

  constructor(private readonly deps: AskDeps<M>) {}

  /** Whether a page has asked for the models: only then does a change in VS Code's models list them again (P46). */
  get listed(): boolean {
    return this.listedOnce;
  }

  /** VS Code's models, listed and kept by their key. A listing that fails lists none. */
  async models(): Promise<HostModelData[]> {
    let models: readonly M[];
    try { models = await this.deps.select(); } catch { models = []; }
    this.listedOnce = true;
    const kept = new Map<string, M>();
    for (const m of models) if (!kept.has(modelKey(m))) kept.set(modelKey(m), m);
    this.kept = kept;
    return [...kept.values()].map(modelData);
  }

  /** How many requests run for this page. */
  runningFor(page: P): number {
    return this.running.get(page)?.size ?? 0;
  }

  /** A page's request, answered as `asked`: its answer, its failure, or its cancellation. */
  async ask(page: P, id: number, route: unknown, value: unknown): Promise<Asked> {
    const controller = new AbortController();
    let mine = this.running.get(page);
    if (!mine) { mine = new Map(); this.running.set(page, mine); }
    mine.get(id)?.abort();
    mine.set(id, controller);
    let key: string | undefined;
    try {
      const request = checkedRequest(value);
      if (!request) throw new AskFailed({ kind: 'other', message: 'the request is not one the Ask box makes' });
      let work: Promise<AskAnswerData>;
      if (route === 'vscode') work = this.viaModel(request, controller.signal);
      else if (route === 'key') {
        key = await this.deps.key();
        work = this.viaKey(request, key, controller.signal);
      } else throw new AskFailed({ kind: 'other', message: `no route ${String(route)}` });
      // A cancel answers at once. VS Code 1.105.0 passes it to neither the provider nor the caller's wait (section 9.21), so
      // the request may run on there to its end, and what it gives then is dropped.
      work.catch(() => undefined);
      const answer = await Promise.race([work, abortOf(controller.signal)]);
      return controller.signal.aborted ? { type: 'asked', id, cancelled: true } : { type: 'asked', id, answer };
    } catch (e) {
      if (controller.signal.aborted) return { type: 'asked', id, cancelled: true };
      const failure = e instanceof AskFailed ? e.failure : { kind: 'other' as const, message: (e as Error)?.message ?? String(e) };
      return { type: 'asked', id, failure: withoutKey(failure, key) };
    } finally {
      if (mine.get(id) === controller) mine.delete(id);
      if (mine.size === 0 && this.running.get(page) === mine) this.running.delete(page);
    }
  }

  /** The person's Cancel. An id this page has not running is ignored. */
  cancel(page: P, id: number): void {
    this.running.get(page)?.get(id)?.abort();
  }

  /** The page went (its panel disposed) or loaded again (a second `ready`): its requests are cancelled (section 9.20 item 5). */
  pageGone(page: P): void {
    const mine = this.running.get(page);
    if (!mine) return;
    this.running.delete(page);
    for (const controller of mine.values()) controller.abort();
  }

  private async viaModel(request: AskRequestData, signal: AbortSignal): Promise<AskAnswerData> {
    const model = this.kept.get(request.model);
    if (!model) throw new AskFailed({ kind: 'notFound', message: request.model });
    const text = chatPrompt(request);
    const { token, dispose } = this.deps.tokenFor(signal);
    try {
      // VS Code passes a message over the model's input to the provider as it is (measured, section 9.21), so it is counted
      // first. A count that fails refuses nothing; a model that is gone throws at once, here and from `sendRequest`.
      let count: number | undefined;
      try { count = await model.countTokens(text, token); } catch { count = undefined; }
      if (signal.aborted) throw new AskFailed({ kind: 'other', message: 'cancelled' });
      if (count !== undefined && model.maxInputTokens > 0 && count > model.maxInputTokens) {
        throw new AskFailed({ kind: 'tooLong', message: `${count} > ${model.maxInputTokens}`, tokens: count, limit: model.maxInputTokens });
      }
      let out = '';
      try {
        const response = await model.sendRequest([this.deps.userMessage(text)], { justification: this.deps.justification(), modelOptions: { max_tokens: request.maxTokens } }, token);
        for await (const part of response.text) out += part;
      } catch (e) {
        throw new AskFailed(lmFailure(e));
      }
      return { text: out, model: model.name, stopReason: null };
    } finally {
      dispose();
    }
  }

  private async viaKey(request: AskRequestData, key: string | undefined, signal: AbortSignal): Promise<AskAnswerData> {
    // Refused before a client is made: the SDK then never looks for a key of its own (section 9.20 finding 2).
    if (!key) throw new AskFailed({ kind: 'noKey', message: 'no Anthropic API key is set' });
    try {
      return await this.deps.sendKey(request, { apiKey: key, baseURL: this.deps.address().trim(), signal });
    } catch (e) {
      throw new AskFailed(failureOf(e));
    }
  }
}
