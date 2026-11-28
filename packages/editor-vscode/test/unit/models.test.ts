// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Ask box's requests in the extension (ED-014 section 4.6), with no `vscode`: the listing kept by vendor and id, both
// routes, the count, the failures by kind, Cancel and a page's going, and the key never in what is posted.
import { describe, expect, test } from 'vitest';
import { AskService, checkedRequest, lmFailure, modelData, MAX_ANSWER_TOKENS, type AskDeps, type ChatModel } from '../../src/models.js';
import { chatPrompt, type AskAnswerData, type AskRequestData } from '../../src/shared.js';

/** A `LanguageModelError` as VS Code throws one: an `Error` with a `code`. */
function lmError(code: string, message = code): Error {
  return Object.assign(new Error(message), { name: 'LanguageModelError', code });
}

interface Sent { messages: unknown[]; options: { justification?: string; modelOptions?: Record<string, unknown> } | undefined; token: unknown }

interface FakeModel extends ChatModel {
  sent: Sent[];
  counted: string[];
}

function fakeModel(fields: Partial<ChatModel> & { vendor: string; id: string }, how: {
  answer?: string[];
  count?: number | 'reject' | 'throw';
  send?: 'throw' | 'reject' | 'hang' | 'midway';
  error?: Error;
} = {}): FakeModel {
  const sent: Sent[] = [];
  const counted: string[] = [];
  const m: FakeModel = {
    name: fields.name ?? `${fields.vendor} ${fields.id}`,
    family: fields.family ?? 'fam',
    version: fields.version ?? '1',
    maxInputTokens: fields.maxInputTokens ?? 100_000,
    vendor: fields.vendor,
    id: fields.id,
    sent,
    counted,
    countTokens(text: string) {
      counted.push(text);
      if (how.count === 'throw') throw how.error ?? lmError('NotFound');
      if (how.count === 'reject') return Promise.reject(new Error('the count fails'));
      return Promise.resolve(how.count ?? Math.ceil(text.length / 4));
    },
    sendRequest(messages, options, token) {
      sent.push({ messages, options, token });
      if (how.send === 'throw') throw how.error ?? lmError('NotFound');
      if (how.send === 'reject') return Promise.reject(how.error ?? lmError('Blocked'));
      if (how.send === 'hang') return new Promise(() => undefined);
      const parts = how.answer ?? ['an ', 'answer'];
      const error = how.error;
      // A stream that fails after its first part, as a provider's dropped connection does.
      if (how.send === 'midway') return Promise.resolve({ text: (async function* () { yield parts[0]!; throw error ?? lmError('Blocked'); })() });
      return Promise.resolve({ text: (async function* () { for (const p of parts) yield p; })() });
    },
  };
  return m;
}

interface Harness {
  service: AskService<FakeModel, string>;
  keySends: { request: AskRequestData; options: { apiKey: string; baseURL: string; signal: AbortSignal } }[];
  tokens: { signal: AbortSignal; disposed: boolean }[];
}

function harness(models: FakeModel[] | (() => Promise<FakeModel[]>), opts: { key?: string; address?: string; sendKey?: AskDeps<FakeModel>['sendKey'] } = {}): Harness {
  const keySends: Harness['keySends'] = [];
  const tokens: Harness['tokens'] = [];
  const service = new AskService<FakeModel, string>({
    select: () => (typeof models === 'function' ? models() : Promise.resolve(models)),
    userMessage: (text) => ({ role: 'user', text }),
    tokenFor: (signal) => { const t = { signal, disposed: false }; tokens.push(t); return { token: t, dispose: () => { t.disposed = true; } }; },
    justification: () => 'JUSTIFICATION',
    key: () => Promise.resolve(opts.key),
    address: () => opts.address ?? '',
    sendKey: opts.sendKey ?? (async (request, options) => {
      keySends.push({ request, options });
      return { text: 'from the key', model: request.model, stopReason: 'end_turn' } satisfies AskAnswerData;
    }),
  });
  return { service, keySends, tokens };
}

const request = (model: string, extra: Partial<AskRequestData> = {}): AskRequestData => ({ model, system: 'SYSTEM', prompt: 'QUESTION', maxTokens: 4096, ...extra });

describe('the listing (P46)', () => {
  test('a model\'s data, from its fields', () => {
    expect(modelData(fakeModel({ vendor: 'copilot', id: 'gpt', name: 'GPT', family: 'gpt-x', version: '2', maxInputTokens: 64_000 })))
      .toEqual({ vendor: 'copilot', id: 'gpt', version: '2', name: 'GPT', family: 'gpt-x', maxInputTokens: 64_000 });
  });

  test('two vendors\' models of one id are kept apart, each sent to by its own object', async () => {
    const a = fakeModel({ vendor: 'copilot', id: 'claude-x' }, { answer: ['from copilot'] });
    const b = fakeModel({ vendor: 'anthropic', id: 'claude-x' }, { answer: ['from anthropic'] });
    const { service } = harness([a, b]);
    expect((await service.models()).map((m) => `${m.vendor}/${m.id}`)).toEqual(['copilot/claude-x', 'anthropic/claude-x']);
    const out = await service.ask('page', 1, 'vscode', request('anthropic/claude-x'));
    expect(out).toEqual({ type: 'asked', id: 1, answer: { text: 'from anthropic', model: 'anthropic claude-x', stopReason: null } });
    expect(a.sent).toHaveLength(0);
    expect(b.sent).toHaveLength(1);
  });

  test('a listing that fails lists none; whether a page has listed is kept', async () => {
    const { service } = harness(() => Promise.reject(new Error('no lm')));
    expect(service.listed).toBe(false);
    expect(await service.models()).toEqual([]);
    expect(service.listed).toBe(true);
  });
});

describe('VS Code\'s models', () => {
  test('one user message, chatPrompt of the request; the answer\'s limit in modelOptions; the justification; the answer\'s text joined', async () => {
    const m = fakeModel({ vendor: 'v', id: 'm', name: 'Model M' }, { answer: ['first, ', 'second'] });
    const { service, tokens } = harness([m]);
    await service.models();
    const out = await service.ask('page', 7, 'vscode', request('v/m'));
    expect(out).toEqual({ type: 'asked', id: 7, answer: { text: 'first, second', model: 'Model M', stopReason: null } });
    expect(m.sent).toHaveLength(1);
    expect(m.sent[0]!.messages).toEqual([{ role: 'user', text: chatPrompt(request('v/m')) }]);
    expect(m.sent[0]!.messages).toEqual([{ role: 'user', text: 'SYSTEM\n\n---\n\nQUESTION' }]);
    expect(m.sent[0]!.options).toEqual({ justification: 'JUSTIFICATION', modelOptions: { max_tokens: 4096 } });
    expect(m.counted).toEqual(['SYSTEM\n\n---\n\nQUESTION']);
    // The token passed to both calls, disposed at the end.
    expect(m.sent[0]!.token).toBe(tokens[0]);
    expect(tokens[0]!.disposed).toBe(true);
  });

  test('over the model\'s input: refused with both figures, nothing sent', async () => {
    const m = fakeModel({ vendor: 'v', id: 'm', maxInputTokens: 100 }, { count: 150 });
    const { service } = harness([m]);
    await service.models();
    expect(await service.ask('page', 1, 'vscode', request('v/m'))).toEqual({ type: 'asked', id: 1, failure: { kind: 'tooLong', message: '150 > 100', tokens: 150, limit: 100 } });
    expect(m.sent).toHaveLength(0);
  });

  test('a count that fails refuses nothing', async () => {
    const m = fakeModel({ vendor: 'v', id: 'm', maxInputTokens: 1 }, { count: 'reject' });
    const { service } = harness([m]);
    await service.models();
    expect((await service.ask('page', 1, 'vscode', request('v/m'))).type).toBe('asked');
    expect(m.sent).toHaveLength(1);
  });

  test('a model that is gone throws NotFound at once from both calls: a failure, not a throw (section 9.20 finding 3)', async () => {
    const m = fakeModel({ vendor: 'v', id: 'm' }, { count: 'throw', send: 'throw', error: lmError('NotFound', 'v/m') });
    const { service } = harness([m]);
    await service.models();
    expect(await service.ask('page', 1, 'vscode', request('v/m'))).toEqual({ type: 'asked', id: 1, failure: { kind: 'notFound', message: 'v/m' } });
  });

  test('a LanguageModelError by its code; any other error by its message', async () => {
    expect(lmFailure(lmError('NoPermissions', 'not allowed'))).toEqual({ kind: 'noPermissions', message: 'not allowed' });
    expect(lmFailure(lmError('Blocked', 'quota'))).toEqual({ kind: 'blocked', message: 'quota' });
    expect(lmFailure(lmError('NotFound', 'gone'))).toEqual({ kind: 'notFound', message: 'gone' });
    expect(lmFailure(new Error('something else'))).toEqual({ kind: 'other', message: 'something else' });
    const m = fakeModel({ vendor: 'v', id: 'm' }, { send: 'reject', error: lmError('NoPermissions', 'refused') });
    const { service } = harness([m]);
    await service.models();
    expect(await service.ask('page', 2, 'vscode', request('v/m'))).toEqual({ type: 'asked', id: 2, failure: { kind: 'noPermissions', message: 'refused' } });
  });

  test('a stream that fails part way is a failure by its error, the part already read dropped (section 9.23)', async () => {
    const m = fakeModel({ vendor: 'v', id: 'm' }, { send: 'midway', answer: ['half an answer'], error: lmError('Blocked', 'the quota ran out') });
    const { service, tokens } = harness([m]);
    await service.models();
    expect(await service.ask('page', 3, 'vscode', request('v/m'))).toEqual({ type: 'asked', id: 3, failure: { kind: 'blocked', message: 'the quota ran out' } });
    expect(tokens[0]!.disposed).toBe(true);
  });

  test('a model no longer offered (or never listed) is refused as gone', async () => {
    const { service } = harness([fakeModel({ vendor: 'v', id: 'm' })]);
    expect(await service.ask('page', 1, 'vscode', request('v/m'))).toEqual({ type: 'asked', id: 1, failure: { kind: 'notFound', message: 'v/m' } });
    await service.models();
    expect(await service.ask('page', 2, 'vscode', request('v/other'))).toEqual({ type: 'asked', id: 2, failure: { kind: 'notFound', message: 'v/other' } });
  });
});

describe('the key\'s route (P45)', () => {
  test('refused before any call with no key stored', async () => {
    const { service, keySends } = harness([]);
    expect(await service.ask('page', 1, 'key', request('claude-sonnet-5'))).toEqual({ type: 'asked', id: 1, failure: { kind: 'noKey', message: 'no Anthropic API key is set' } });
    expect(keySends).toHaveLength(0);
  });

  test('sent with the stored key and the setting\'s address, the request as the page shows it', async () => {
    const { service, keySends } = harness([], { key: 'sk-the-key', address: ' https://proxy.example ' });
    const out = await service.ask('page', 3, 'key', request('claude-sonnet-5'));
    expect(out).toEqual({ type: 'asked', id: 3, answer: { text: 'from the key', model: 'claude-sonnet-5', stopReason: 'end_turn' } });
    expect(keySends).toHaveLength(1);
    expect(keySends[0]!.request).toEqual(request('claude-sonnet-5'));
    expect(keySends[0]!.options.apiKey).toBe('sk-the-key');
    expect(keySends[0]!.options.baseURL).toBe('https://proxy.example');
  });

  test('an SDK failure as data; the key never in what is posted, even where an error\'s words carry it', async () => {
    const { service } = harness([], { key: 'sk-secret-1', sendKey: () => Promise.reject(new Error('bad key sk-secret-1 refused')) });
    const out = await service.ask('page', 4, 'key', request('claude-sonnet-5'));
    expect(out).toEqual({ type: 'asked', id: 4, failure: { kind: 'other', message: 'bad key … refused' } });
    expect(JSON.stringify(out)).not.toContain('sk-secret-1');
  });

  test('nothing the extension posts holds the key: answers, failures, cancellations, the listing', async () => {
    const key = 'sk-ant-never-posted';
    const posts: unknown[] = [];
    const m = fakeModel({ vendor: 'v', id: 'm' }, { answer: ['ok'] });
    const { service } = harness([m], { key, sendKey: async (req) => ({ text: 'fine', model: req.model, stopReason: 'end_turn' }) });
    posts.push(await service.models());
    posts.push(await service.ask('page', 1, 'key', request('claude-sonnet-5')));
    posts.push(await service.ask('page', 2, 'vscode', request('v/m')));
    const pending = service.ask('page', 3, 'key', request('claude-sonnet-5'));
    service.cancel('page', 3);
    posts.push(await pending);
    expect(JSON.stringify(posts)).not.toContain(key);
  });
});

describe('Cancel and a page\'s going (section 9.20 item 5)', () => {
  test('a Cancel answers at once, though the model ignores its token (VS Code 1.105.0, section 9.21)', async () => {
    const m = fakeModel({ vendor: 'v', id: 'm' }, { send: 'hang' });
    const { service, tokens } = harness([m]);
    await service.models();
    const pending = service.ask('page', 5, 'vscode', request('v/m'));
    await new Promise((r) => setTimeout(r, 0));
    expect(service.runningFor('page')).toBe(1);
    service.cancel('page', 5);
    expect(await pending).toEqual({ type: 'asked', id: 5, cancelled: true });
    expect(tokens[0]!.signal.aborted).toBe(true);
    expect(service.runningFor('page')).toBe(0);
  });

  test('the key\'s request gets the signal the Cancel aborts', async () => {
    let seen: AbortSignal | undefined;
    const { service } = harness([], { key: 'k', sendKey: (_r, o) => { seen = o.signal; return new Promise(() => undefined); } });
    const pending = service.ask('page', 1, 'key', request('claude-sonnet-5'));
    await new Promise((r) => setTimeout(r, 0));
    service.cancel('page', 1);
    expect(await pending).toEqual({ type: 'asked', id: 1, cancelled: true });
    expect(seen?.aborted).toBe(true);
  });

  test('a page that goes, or loads again, has every request cancelled; another page\'s run on', async () => {
    const m = fakeModel({ vendor: 'v', id: 'm' }, { send: 'hang' });
    const { service, tokens } = harness([m]);
    await service.models();
    const a1 = service.ask('A', 1, 'vscode', request('v/m'));
    const a2 = service.ask('A', 2, 'vscode', request('v/m'));
    void service.ask('B', 1, 'vscode', request('v/m'));
    await new Promise((r) => setTimeout(r, 0));
    service.pageGone('A');
    expect(await a1).toEqual({ type: 'asked', id: 1, cancelled: true });
    expect(await a2).toEqual({ type: 'asked', id: 2, cancelled: true });
    expect(tokens.map((t) => t.signal.aborted)).toEqual([true, true, false]);
    expect(service.runningFor('B')).toBe(1);
  });

  test('a Cancel for an id the page has not running is ignored', async () => {
    const { service } = harness([fakeModel({ vendor: 'v', id: 'm' })]);
    await service.models();
    expect(() => service.cancel('page', 99)).not.toThrow();
    expect(() => service.cancel('nobody', 1)).not.toThrow();
    expect(() => service.pageGone('nobody')).not.toThrow();
    expect((await service.ask('page', 1, 'vscode', request('v/m'))).type).toBe('asked');
  });
});

describe('a request as the page posted it', () => {
  test('checked: its fields, and the answer\'s limit capped', () => {
    expect(checkedRequest(request('m'))).toEqual(request('m'));
    expect(checkedRequest(request('m', { maxTokens: 1_000_000 }))?.maxTokens).toBe(MAX_ANSWER_TOKENS);
    for (const bad of [null, 'text', {}, { ...request('m'), model: '' }, { ...request('m'), system: 3 }, { ...request('m'), maxTokens: 0 }, { ...request('m'), maxTokens: 1.5 }]) {
      expect(checkedRequest(bad)).toBeUndefined();
    }
  });

  test('a malformed request or an unknown route is a failure', async () => {
    const { service } = harness([]);
    expect(await service.ask('page', 1, 'vscode', { model: 3 })).toEqual({ type: 'asked', id: 1, failure: { kind: 'other', message: 'the request is not one the Ask box makes' } });
    expect(await service.ask('page', 2, 'elsewhere', request('m'))).toEqual({ type: 'asked', id: 2, failure: { kind: 'other', message: 'no route elsewhere' } });
  });
});
