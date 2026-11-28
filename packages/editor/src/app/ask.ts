// SPDX-License-Identifier: GPL-3.0-or-later
//
// The ask mode's provider adapter (ED-003 section 7.5; section 9 items 5 and 30): one interface,
// one implementation in E2.e, the Anthropic API through its SDK in the browser with the user's
// own key (bring your own, kept in local storage, entered in Options and written nowhere else)
// and a base URL an organisation proxy (E4) or the suites' test double can take. One request
// per question, no streaming, no server-side fallback to another model: the user chose the
// model and is shown the request; a refusal is shown as such with its explanation. The call
// itself is `ask-send.ts`'s, which VS Code's extension makes too (ED-014 section 4.6, P45): in
// VS Code the page holds no key and sends through the host (`host().ask()`).
import { modelKey, PAGE_COMMANDS, type AskAnswerData, type AskFailureData, type HostModelData } from './host-protocol.js';
import type { AskRequest } from './ask-context.js';
import { sendAnthropic, failureOf, isCancelled } from './ask-send.js';
import { AskFailure, host } from './host.js';
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

export type AskAnswer = AskAnswerData;

/** `ask.model` keeps a VS Code model as `vscode:<vendor>/<id>`, and an Anthropic model by its bare id (P46). */
export const VSCODE_MODEL = 'vscode:';

/** The VS Code model a kept choice names, when it is offered. */
export function vscodeModelOf(choice: string, models: readonly HostModelData[]): HostModelData | undefined {
  if (!choice.startsWith(VSCODE_MODEL)) return undefined;
  const key = choice.slice(VSCODE_MODEL.length);
  return models.find((m) => modelKey(m) === key);
}

/**
 * The model the Ask box offers first in VS Code (P46): the one kept in `ask.model` while it is offered; else the first VS
 * Code model whose family or id names Claude, since the prompt and the fenced forms were written for Claude (ED-003 section
 * 7.5); else the first listed. Undefined when nothing is offered: the chat route alone.
 */
export function chooseModel(kept: string, models: readonly HostModelData[], keyStored: boolean): string | undefined {
  const offered = [...models.map((m) => `${VSCODE_MODEL}${modelKey(m)}`), ...(keyStored ? ASK_MODELS.map((m) => m.id) : [])];
  if (kept !== '' && offered.includes(kept)) return kept;
  const claude = models.find((m) => /claude/i.test(m.family) || /claude/i.test(m.id));
  if (claude) return `${VSCODE_MODEL}${modelKey(claude)}`;
  return offered[0];
}

/** Where a request goes: VS Code's model or the key, through the host; the page's own call in a browser; a chat by copy and paste. */
export type AskRoute = 'vscode' | 'key' | 'browser' | 'chat';

/** The model the box has chosen, the route a request to it takes, and its name for the request and the sending line. */
export interface AskChoice {
  /** The select's value: `vscode:<vendor>/<id>` or an Anthropic model's id; undefined when nothing is offered. */
  model: string | undefined;
  /** The VS Code model chosen, when the choice is one. */
  vsModel: HostModelData | undefined;
  route: AskRoute;
  label: string;
}

/**
 * The choice. In VS Code, over VS Code's models as listed: `chooseModel`'s model, sent to VS Code's model or with the key the
 * extension holds, else the chat. In a browser (`models` unused, `keyStored` the key in Options): kept or the first model,
 * sent with the key or by the chat.
 */
export function askChoice(inVsCode: boolean, kept: string, models: readonly HostModelData[], keyStored: boolean): AskChoice {
  const labelOf = (id: string): string => ASK_MODELS.find((m) => m.id === id)?.label ?? id;
  if (!inVsCode) {
    const model = kept || ASK_MODELS[0]!.id;
    return { model, vsModel: undefined, route: keyStored ? 'browser' : 'chat', label: labelOf(model) };
  }
  const model = chooseModel(kept, models, keyStored);
  if (model === undefined) return { model, vsModel: undefined, route: 'chat', label: '' };
  const vsModel = vscodeModelOf(model, models);
  if (vsModel) return { model, vsModel, route: 'vscode', label: vsModel.name };
  return { model, vsModel, route: keyStored && ASK_MODELS.some((m) => m.id === model) ? 'key' : 'chat', label: labelOf(model) };
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
  send: (request, options) => sendAnthropic(request, { ...options, browser: true }),
};

export function askOption(key: string): string {
  try { return (localStorage.getItem(key) ?? '').trim(); } catch { return ''; }
}

export function setAskOption(key: string, value: string): void {
  try { if (value === '') localStorage.removeItem(key); else localStorage.setItem(key, value); } catch { /* ignore */ }
}

/** The command that sets the key in VS Code (ED-014 section 4.6). */
const SET_KEY = 'docx4j.ask.setKey';

/**
 * The name the palette shows for the command that sets the key (ED-014 section 4.7, P49): VS Code's, in VS Code's language,
 * which is the one a person can find there whatever the page's; English where the host names none.
 */
export function setKeyCommand(): string {
  return host().commandName(SET_KEY) ?? PAGE_COMMANDS[SET_KEY]!;
}

/**
 * A failure's words, the same in either host for each kind (ED-014 section 4.6), but where they name the place the key is
 * set: Options in a browser, the command in VS Code, whose Options cannot show the key (section 9.23 finding 1).
 */
export function describeFailure(f: AskFailureData): string {
  switch (f.kind) {
    case 'key': return host().has('askOptions') ? t('ask.error.key') : t('ask.error.key.vscode', { command: setKeyCommand() });
    case 'rateLimit': return t('ask.error.rateLimit');
    case 'connection': return t('ask.error.connection', { error: f.message });
    case 'api': return t('ask.error.api', { status: f.status ?? '', error: f.message });
    case 'noPermissions': return t('ask.error.noPermissions');
    case 'blocked': return t('ask.error.blocked', { error: f.message });
    case 'notFound': return t('ask.error.notFound');
    case 'tooLong': return t('ask.error.tooLong', { tokens: f.tokens ?? 0, limit: f.limit ?? 0 });
    case 'noKey': return t('ask.error.noKey', { command: setKeyCommand() });
    case 'other': return f.message;
  }
}

/** A message for the status line from a failed request: an SDK error in a browser, the host's failure in VS Code. */
export function describeError(e: unknown): string {
  if (e instanceof AskFailure) return describeFailure(e.failure);
  return describeFailure(failureOf(e));
}

/** Whether a failed request was the person's Cancel: nothing to say. */
export function wasCancelled(e: unknown): boolean {
  return isCancelled(e);
}
