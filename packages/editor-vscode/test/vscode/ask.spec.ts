// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Ask box in a real VS Code (ED-014 section 4.6, S6.4). The suite's driver registers a language model, vendor
// `docx4j-test`, which answers as the browser suite's double does; the key's route goes to that double, served here, at the
// address a user setting names.
// - VS Code's model: listed, sent the request shown as one message with its answer's limit, the answer inserted; a model whose
//   input the request exceeds refuses before sending; the page reaches nothing outside and its policy refuses nothing.
// - The key: set through the command's password box, sent with by the extension, never in the page; removed, its group gone.
// - Neither a model nor a key: the chat route alone.
// - An editor closed mid-request: the model's token cancelled (1.147.0; 1.105.0 passes no cancel on, section 9.21).
// - Restricted Mode: VS Code's model listed and answering.
import { test, expect, type Frame, type Locator } from '@playwright/test';
import { createServer, type IncomingHttpHeaders, type Server } from 'node:http';
import type { AddressInfo } from 'node:net';
import { VsCode, bodyText, caretAtEndOf, sleep } from './harness.js';
import { askRoutes } from '../../../editor-model/test/fixtures/build/ask-double.mjs';

test.describe.configure({ mode: 'serial' });

/** The chat prompt's rule between the system prompt and the message (`chatPrompt`). */
const RULE = '\n\n---\n\n';

interface Received { model: string; messages: { role: number; text: string }[]; modelOptions: Record<string, unknown> | null; cancelled: boolean }

let vs: VsCode;
let server: Server;
let address = '';
const seen: { url: string; headers: IncomingHttpHeaders }[] = [];

test.beforeAll(async () => {
  // The key's double: the browser suite's `askRoutes` under `/ask`, recording each request's headers.
  const routes = askRoutes('/ask');
  server = createServer((req, res) => {
    seen.push({ url: req.url ?? '', headers: req.headers });
    routes(req, res, () => { res.statusCode = 404; res.end(); });
  });
  await new Promise<void>((r) => server.listen(0, '127.0.0.1', () => r()));
  address = `http://127.0.0.1:${(server.address() as AddressInfo).port}/ask`;
  vs = await VsCode.launch({ profile: 'ask', lm: true, files: { 'a.docx': 'HelloWordOnline.docx', 'b.docx': 'HelloWordOnline.docx' } });
});

test.afterAll(async () => {
  await vs?.close();
  await new Promise<void>((r) => server?.close(() => r()));
});

test.afterEach(async () => {
  expect(vs.problems, 'content security policy refusals').toEqual([]);
});

const received = (): Promise<Received[]> => vs.ctl<Received[]>({ op: 'lmReceived' });

/**
 * A button pressed from the keyboard. On VS Code 1.105.0 the suite's mouse clicks over the DevTools Protocol did not reach the
 * perspective tabs, where the keyboard did (section 9.22; the cause was not found): the keys work on every release.
 */
async function press(target: Locator): Promise<void> {
  await target.press('Enter');
}

/** A checkbox set from the keyboard. */
async function setCheck(target: Locator, on: boolean): Promise<void> {
  if ((await target.isChecked()) !== on) await target.press('Space');
  await expect(target).toBeChecked({ checked: on });
}

/** The Developer perspective's Ask box in the page. */
async function askBox(frame: Frame) {
  const tab = frame.getByTestId('perspective-developer');
  await press(tab);
  await expect(tab).toHaveAttribute('aria-selected', 'true');
  const box = frame.getByTestId('ask-box');
  await expect(box).toBeVisible();
  return box;
}

/** The request as the page shows it: its system prompt and its message, as the <pre>s hold them. */
async function shown(frame: Frame): Promise<{ system: string; prompt: string }> {
  const pres = frame.getByTestId('ask-request').locator('pre');
  return { system: (await pres.nth(0).textContent()) ?? '', prompt: (await pres.nth(1).textContent()) ?? '' };
}

/** Every resource the page fetched that is not the webview's own. */
async function outside(frame: Frame): Promise<string[]> {
  const urls = await frame.evaluate(() => performance.getEntriesByType('resource').map((e) => e.name));
  // The buffer holds 250 entries by default and records nothing after them: a full one would hold nothing to the test (section 9.23).
  expect(urls.length).toBeLessThan(250);
  return urls.filter((u) => !/^(vscode-webview:|blob:|data:)/.test(u) && !/^https:\/\/[^/]*\.vscode-cdn\.net\//.test(u));
}

test('VS Code\'s model: listed, sent the request shown as one message, the answer inserted; a request too long refused before sending', async () => {
  const frame = await vs.open('a.docx');
  const box = await askBox(frame);
  const select = box.getByTestId('ask-model');
  await expect(select.locator('option')).toHaveText(['Test Claude (docx4j-test)', 'Tiny Claude (docx4j-test)']);
  await expect(select).toHaveValue('vscode:docx4j-test/claude-test');
  await caretAtEndOf(frame, 0);
  await box.getByTestId('ask-prompt').fill('Add a heading about the double');
  await press(box.getByTestId('ask-ask'));
  await expect(box.getByTestId('ask-one-message')).toContainText('Test Claude');
  const request = await shown(frame);
  expect(request.prompt).toContain('Add a heading about the double');
  const before = (await received()).length;
  await press(box.getByTestId('ask-send'));
  await expect(box.getByTestId('ask-preview')).toContainText('The double answers: Add a heading about the double');
  await expect(frame.getByTestId('status')).toContainText('Answered by Test Claude');
  const got = (await received()).slice(before);
  expect(got).toHaveLength(1);
  expect(got[0]!.model).toBe('claude-test');
  // One message, the user's (role 1), the request shown: the system prompt, the rule, the message.
  expect(got[0]!.messages).toEqual([{ role: 1, text: `${request.system}${RULE}${request.prompt}` }]);
  expect(got[0]!.modelOptions).toEqual({ max_tokens: 4096 });
  await press(box.getByTestId('ask-insert'));
  await expect.poll(() => bodyText(frame)).toContain('The double answers: Add a heading about the double');
  // A model whose input the request exceeds: refused with both figures, nothing sent to it.
  await select.selectOption('vscode:docx4j-test/claude-tiny');
  await setCheck(box.getByTestId('ask-send-at-once'), true);
  await box.getByTestId('ask-prompt').fill('Another heading');
  await press(box.getByTestId('ask-ask'));
  await expect(box.getByTestId('ask-error')).toContainText('too long for this model');
  await expect(box.getByTestId('ask-error')).toContainText('where it takes 1,000');
  expect((await received()).length).toBe(before + 1);
  expect(await outside(frame)).toEqual([]);
  await setCheck(box.getByTestId('ask-send-at-once'), false);
  await vs.closeAll();
});

test('the key: set through the command\'s password box, sent by the extension to the setting\'s address, never in the page; removed, its group gone', async () => {
  await vs.ctl({ op: 'config', section: 'docx4j.ask', key: 'baseUrl', value: address });
  const frame = await vs.open('a.docx');
  const box = await askBox(frame);
  await expect(box.getByTestId('ask-group-key')).toHaveCount(0);
  void vs.exec('docx4j.ask.setKey').catch(() => undefined);
  const input = vs.page.locator('.quick-input-widget input:not([type="checkbox"])');
  await input.waitFor({ timeout: 15_000 });
  await expect(input).toHaveAttribute('type', 'password');
  await input.fill('sk-suite-key');
  await sleep(200);
  await vs.page.keyboard.press('Enter');
  // The key's state reaches the page as a push: the Anthropic group appears.
  await expect(box.getByTestId('ask-group-key')).toHaveCount(1);
  await vs.exec('notifications.clearAll').catch(() => undefined);
  await box.getByTestId('ask-model').selectOption('claude-sonnet-5');
  await caretAtEndOf(frame, 0);
  await box.getByTestId('ask-prompt').fill('Add a paragraph from the key');
  await press(box.getByTestId('ask-ask'));
  await expect(box.getByTestId('ask-request')).toHaveAttribute('data-route', 'key');
  const before = seen.length;
  await press(box.getByTestId('ask-send'));
  await expect(box.getByTestId('ask-preview')).toContainText('The double answers: Add a paragraph from the key');
  const sent = seen.slice(before);
  expect(sent.map((s) => s.url)).toEqual(['/ask/v1/messages']);
  expect(sent[0]!.headers['x-api-key']).toBe('sk-suite-key');
  expect(sent[0]!.headers['authorization']).toBeUndefined();
  expect(sent[0]!.headers['anthropic-dangerous-direct-browser-access']).toBeUndefined();
  await press(box.getByTestId('ask-insert'));
  await expect.poll(() => bodyText(frame)).toContain('The double answers: Add a paragraph from the key');
  // The key is nowhere in the page: not in its storage, not in what was written into it.
  expect(await frame.evaluate(() => [Object.values(localStorage).join('\n'), document.documentElement.outerHTML].some((s) => s.includes('sk-suite-key')))).toBe(false);
  expect(await outside(frame)).toEqual([]);
  // A page opened while a key is stored has the Anthropic group from its boot data (the boot had dropped it, section 9.26).
  await vs.closeAll();
  const again = await vs.open('a.docx');
  const box2 = await askBox(again);
  await expect(box2.getByTestId('ask-group-key')).toHaveCount(1);
  await vs.exec('docx4j.ask.removeKey');
  await expect(box2.getByTestId('ask-group-key')).toHaveCount(0);
  await vs.exec('notifications.clearAll').catch(() => undefined);
  await vs.closeAll();
});

test('neither a model nor a key: the chat route alone; a model registered later is offered as a question starts', async () => {
  await vs.ctl({ op: 'lmUnregister' });
  const frame = await vs.open('b.docx');
  const box = await askBox(frame);
  await box.getByTestId('ask-prompt').fill('A question for a chat');
  await expect(box.getByTestId('ask-model')).toHaveCount(0);
  await press(box.getByTestId('ask-ask'));
  await expect(box.getByTestId('ask-no-key')).toContainText('VS Code offers none');
  await expect(box.getByTestId('ask-copy-prompt')).toBeVisible();
  await expect(box.getByTestId('ask-send')).toHaveCount(0);
  // A provider registering is an event neither release sends on (section 9.21): the page asks again when a question starts.
  await vs.ctl({ op: 'lmRegister' });
  await frame.getByTestId('ask-prompt').blur();
  await frame.getByTestId('ask-prompt').focus();
  await expect(box.getByTestId('ask-model')).toHaveValue('vscode:docx4j-test/claude-test');
  await vs.closeAll();
});

test('an editor closed mid-request: the request cancelled at the model (1.147.0), or left to run on there (1.105.0)', async () => {
  const frame = await vs.open('a.docx');
  const box = await askBox(frame);
  await expect(box.getByTestId('ask-model')).toHaveValue('vscode:docx4j-test/claude-test');
  await setCheck(box.getByTestId('ask-send-at-once'), true);
  const before = (await received()).length;
  await box.getByTestId('ask-prompt').fill('SLOW, then closed');
  await press(box.getByTestId('ask-ask'));
  await expect(box.getByTestId('ask-sending')).toBeVisible();
  await expect.poll(async () => (await received()).length).toBe(before + 1);
  await setCheck(box.getByTestId('ask-send-at-once'), false);
  await vs.closeAll();
  if (vs.olderThan(1, 147)) {
    // VS Code 1.105.0 passes no cancel to a provider: the request runs on, its answer dropped with the page.
    await sleep(2000);
    expect((await received())[before]!.cancelled).toBe(false);
  } else {
    await expect.poll(async () => (await received())[before]!.cancelled, { timeout: 15_000 }).toBe(true);
  }
});

test('Restricted Mode: VS Code\'s model listed and answering; nothing runs', async () => {
  await vs.close();
  vs = await VsCode.launch({ profile: 'ask-restricted', lm: true, trust: true, files: { 'a.docx': 'HelloWordOnline.docx' } });
  expect((await vs.ctl<{ trusted: boolean }>({ op: 'ping' })).trusted).toBe(false);
  const frame = await vs.open('a.docx');
  const box = await askBox(frame);
  await expect(box.getByTestId('ask-model')).toHaveValue('vscode:docx4j-test/claude-test');
  await caretAtEndOf(frame, 0);
  await setCheck(box.getByTestId('ask-send-at-once'), true);
  await box.getByTestId('ask-prompt').fill('Write a script that italicises the selection');
  await press(box.getByTestId('ask-ask'));
  // A script answer lands in the console unrun, in any workspace.
  await expect(box.getByTestId('ask-answer')).toHaveAttribute('data-form', 'script');
  await expect(box.getByTestId('ask-script-note')).toContainText('has not run');
  await setCheck(box.getByTestId('ask-send-at-once'), false);
  await vs.closeAll();
});
