// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Ask box's call to the Anthropic API (ED-014 section 4.6, P45), in Node as VS Code's extension host makes it: the body
// the browser sends, nothing read from the environment, no browser header outside a browser, and each SDK failure as data.
import { afterEach, beforeEach, describe, expect, test } from 'vitest';
import { createServer, type IncomingHttpHeaders, type Server } from 'node:http';
import type { AddressInfo } from 'node:net';
import { messageFor } from '../../editor-model/test/fixtures/build/ask-double.mjs';
import { ANTHROPIC_ADDRESS, chatPrompt, failureOf, isCancelled, sendAnthropic } from '../src/app/ask-send.js';
import type { AskRequestData } from '../src/app/host-protocol.js';

interface Seen { url: string; headers: IncomingHttpHeaders; body: unknown }

/** The browser suite's double, answering as `askRoutes` does, recording each request; `status` answers an error instead. */
async function double(status?: number): Promise<{ address: string; seen: Seen[]; close(): Promise<void> }> {
  const seen: Seen[] = [];
  const server: Server = createServer((req, res) => {
    const chunks: Buffer[] = [];
    req.on('data', (c: Buffer) => chunks.push(c));
    req.on('end', () => {
      const body = JSON.parse(Buffer.concat(chunks).toString('utf8') || 'null') as unknown;
      seen.push({ url: req.url ?? '', headers: req.headers, body });
      res.setHeader('content-type', 'application/json');
      if (status !== undefined) {
        res.statusCode = status;
        res.end(JSON.stringify({ type: 'error', error: { type: status === 401 ? 'authentication_error' : status === 429 ? 'rate_limit_error' : 'api_error', message: `the double answers ${status}` } }));
        return;
      }
      res.end(JSON.stringify(messageFor(body)));
    });
  });
  await new Promise<void>((r) => server.listen(0, '127.0.0.1', () => r()));
  const { port } = server.address() as AddressInfo;
  return { address: `http://127.0.0.1:${port}`, seen, close: () => new Promise((r) => server.close(() => r())) };
}

const request: AskRequestData = { model: 'claude-sonnet-5', system: 'SYSTEM', prompt: 'a heading for the report', maxTokens: 4096 };

const ENV = ['ANTHROPIC_BASE_URL', 'ANTHROPIC_AUTH_TOKEN', 'ANTHROPIC_API_KEY'] as const;
let saved: Record<string, string | undefined> = {};
beforeEach(() => { saved = Object.fromEntries(ENV.map((k) => [k, process.env[k]])); });
afterEach(() => { for (const k of ENV) { if (saved[k] === undefined) delete process.env[k]; else process.env[k] = saved[k]; } });

describe('sendAnthropic', () => {
  test('the body the browser sends: the system prompt, one user message, the limit and adaptive thinking; the answer', async () => {
    const d = await double();
    try {
      const answer = await sendAnthropic(request, { apiKey: 'sk-test', baseURL: d.address });
      expect(d.seen).toHaveLength(1);
      expect(d.seen[0]!.url).toBe('/v1/messages');
      expect(d.seen[0]!.body).toEqual({
        model: 'claude-sonnet-5',
        max_tokens: 4096,
        system: 'SYSTEM',
        messages: [{ role: 'user', content: 'a heading for the report' }],
        thinking: { type: 'adaptive' },
      });
      expect(d.seen[0]!.headers['x-api-key']).toBe('sk-test');
      expect(answer.model).toBe('claude-sonnet-5');
      expect(answer.stopReason).toBe('end_turn');
      expect(answer.text).toContain('```wml');
      expect(answer.usage).toEqual({ input: 1, output: 1 });
    } finally { await d.close(); }
  });

  test('nothing from the environment: ANTHROPIC_BASE_URL, ANTHROPIC_AUTH_TOKEN and ANTHROPIC_API_KEY unread (section 9.20 finding 2)', async () => {
    const d = await double();
    const trap = await double();
    process.env['ANTHROPIC_BASE_URL'] = trap.address;
    process.env['ANTHROPIC_AUTH_TOKEN'] = 'token-from-the-environment';
    process.env['ANTHROPIC_API_KEY'] = 'key-from-the-environment';
    try {
      await sendAnthropic(request, { apiKey: 'sk-given', baseURL: d.address });
      expect(trap.seen).toHaveLength(0);
      expect(d.seen).toHaveLength(1);
      expect(d.seen[0]!.headers['authorization']).toBeUndefined();
      expect(d.seen[0]!.headers['x-api-key']).toBe('sk-given');
    } finally { await d.close(); await trap.close(); }
  });

  test('an empty address is Anthropic\'s own, never the environment\'s', async () => {
    process.env['ANTHROPIC_BASE_URL'] = 'http://127.0.0.1:9/never';
    const urls: string[] = [];
    const real = globalThis.fetch;
    // The SDK takes the global `fetch` when the client is made, as it takes VS Code's in the extension host (section 9.21).
    globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
      urls.push(typeof input === 'string' ? input : input instanceof URL ? input.href : input.url);
      return new Response(JSON.stringify(messageFor(JSON.parse(String(init?.body)))), { status: 200, headers: { 'content-type': 'application/json' } });
    }) as typeof fetch;
    try {
      await sendAnthropic(request, { apiKey: 'sk', baseURL: '' });
      await sendAnthropic(request, { apiKey: 'sk', baseURL: '   ' });
      await sendAnthropic(request, { apiKey: 'sk' });
    } finally { globalThis.fetch = real; }
    expect(urls).toEqual([`${ANTHROPIC_ADDRESS}/v1/messages`, `${ANTHROPIC_ADDRESS}/v1/messages`, `${ANTHROPIC_ADDRESS}/v1/messages`]);
  });

  test('the browser\'s header only for the page\'s own call in a browser (section 9.20 item 9)', async () => {
    const d = await double();
    try {
      await sendAnthropic(request, { apiKey: 'sk', baseURL: d.address });
      await sendAnthropic(request, { apiKey: 'sk', baseURL: d.address, browser: true });
      expect(d.seen.map((s) => s.headers['anthropic-dangerous-direct-browser-access'] ?? null)).toEqual([null, 'true']);
      expect(d.seen[0]!.headers['x-stainless-runtime']).toBe('node');
    } finally { await d.close(); }
  });

  test('a refusal: its category and explanation', async () => {
    const d = await double();
    try {
      const answer = await sendAnthropic({ ...request, prompt: 'refuse this' }, { apiKey: 'sk', baseURL: d.address });
      expect(answer.stopReason).toBe('refusal');
      expect(answer.refusal).toEqual({ category: 'test', explanation: 'The double declines prompts that say refuse' });
    } finally { await d.close(); }
  });
});

describe('failureOf', () => {
  const failing = async (status?: number, address?: string): Promise<unknown> => {
    const d = status === undefined ? undefined : await double(status);
    try {
      await sendAnthropic(request, { apiKey: 'sk', baseURL: address ?? d!.address });
    } catch (e) { return e; } finally { await d?.close(); }
    throw new Error('no failure');
  };

  test('each SDK error class as a kind, with its status', async () => {
    expect(failureOf(await failing(401))).toMatchObject({ kind: 'key', status: 401 });
    expect(failureOf(await failing(429))).toMatchObject({ kind: 'rateLimit', status: 429 });
    expect(failureOf(await failing(400))).toMatchObject({ kind: 'api', status: 400 });
    // A closed port: the SDK's connection error, after its one retry.
    const closed = await double();
    await closed.close();
    expect(failureOf(await failing(undefined, closed.address))).toMatchObject({ kind: 'connection', message: 'Connection error.' });
    expect(failureOf(new Error('anything else'))).toEqual({ kind: 'other', message: 'anything else' });
  }, 30_000);

  test('a request its signal cancelled', async () => {
    const d = await double();
    const controller = new AbortController();
    controller.abort();
    try {
      const e = await sendAnthropic(request, { apiKey: 'sk', baseURL: d.address, signal: controller.signal }).then(() => undefined, (x: unknown) => x);
      expect(isCancelled(e)).toBe(true);
      expect(isCancelled(new Error('no'))).toBe(false);
    } finally { await d.close(); }
  });
});

describe('chatPrompt', () => {
  test('unchanged: the system prompt, a rule, the question', () => {
    expect(chatPrompt({ model: 'm', system: 'CONTRACT', prompt: 'QUESTION', maxTokens: 4096 })).toBe('CONTRACT\n\n---\n\nQUESTION');
  });
});
