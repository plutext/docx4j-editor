// SPDX-License-Identifier: GPL-3.0-or-later
//
// The agent's companion in a real VS Code (ED-014 section 4.5, S5.5). The test plays the agent: it calls the companion's
// `/mcp` over HTTP at the address "Copy the Agent's MCP Address" copies, with the token the Claude Code command carries.
// - The MCP server: not offered before the person admits the agent, offered after; VS Code starts it and calls a tool.
// - Invite from the page: the companion started, the room connected, the agent's edit a tracked change under its name.
// - Stop in People: the next call answers agent.stopped.
// - Share: the link opened in a browser beside VS Code, its typing reaching VS Code's page.
// - Join with the link the agent's own `open` gives: an untitled editor in the room, dirty.
// - A kill and a restart: the rooms rejoined, the address unchanged, the agent out until invited again.
// - A Save As of the joined editor: still in the room, saved.
// - The window closed: no companion left. And Restricted Mode: no Invite, no Share, no server.
import { test, expect, chromium, type Frame } from '@playwright/test';
import { execFileSync } from 'node:child_process';
import { readdirSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { VsCode, caretAtEndOf, sleep } from './harness.js';

test.describe.configure({ mode: 'serial' });

let vs: VsCode;
/** Jason's clipboard, put back after the suite reads the address through it. */
let kept: string | undefined;
let agent: { address: string; token: string } | undefined;

const SERVER = 'plutext.docx4j-editor/docx4j editor';
/** VS Code writes a crashed window's editor state every 60 s (section 9.9 item 12). */
const STATE_WRITTEN_MS = 62_000;

test.beforeAll(async () => {
  vs = await VsCode.launch({
    profile: 'agent',
    files: { 'a.docx': 'HelloWordOnline.docx', 'b.docx': 'HelloWordOnline.docx', 'c.docx': 'HelloWordOnline.docx' },
    // VS Code's MCP service runs only with its AI features on.
    settings: { 'chat.disableAIFeatures': false, 'docx4j.agent.name': 'Agent' },
  });
  kept = await vs.ctl<string>({ op: 'clipboardRead' }).catch(() => undefined);
});

test.afterAll(async () => {
  if (vs && kept !== undefined) await vs.ctl({ op: 'clipboardWrite', text: kept }).catch(() => undefined);
  await vs?.close();
  for (const pid of companions()) { try { process.kill(pid, 'SIGKILL'); } catch { /* gone */ } }
});

test.afterEach(async () => {
  expect(vs.problems, 'content security policy refusals').toEqual([]);
});

/** The companions the extension started for this suite's profiles: their bundle, their rooms under the suite's results. */
function companions(): number[] {
  const out = execFileSync('ps', ['-eo', 'pid=,args='], { encoding: 'utf8' });
  return out.split('\n').filter((l) => l.includes('/companion/cli.mjs') && l.includes('test-results/vscode/agent')).map((l) => Number(l.trim().split(/\s+/)[0]));
}

const alive = (pid: number): boolean => { try { process.kill(pid, 0); return true; } catch { return false; } };

/** "docx4j editor: Copy the Agent's MCP Address": the address from the clipboard, and the token from the Claude Code command. */
async function copyAddress(): Promise<{ address: string; token: string }> {
  await vs.ctl({ op: 'clipboardWrite', text: '' });
  void vs.exec('docx4j.copyAgentAddress').catch(() => undefined);
  let address = '';
  await expect.poll(async () => { address = await vs.ctl<string>({ op: 'clipboardRead' }); return address; }, { timeout: 30_000 }).toMatch(/^http:\/\/127\.0\.0\.1:\d+\/mcp$/);
  const button = vs.page.locator('.notification-toast .monaco-button', { hasText: 'Copy the Claude Code Command' });
  await button.waitFor();
  await sleep(400);
  await button.click();
  let line = '';
  await expect.poll(async () => { line = await vs.ctl<string>({ op: 'clipboardRead' }); return line; }).toContain('claude mcp add');
  await vs.exec('notifications.clearAll').catch(() => undefined);
  const token = /Bearer ([A-Za-z0-9_-]+)"/.exec(line)![1]!;
  expect(line).toBe(`claude mcp add --transport http docx4j-editor ${address} --header "Authorization: Bearer ${token}"`);
  return { address, token };
}

let rpc = 1;
/** A tool called as an MCP client calls it: one stateless POST with the token. */
async function call(name: string, args: Record<string, unknown> = {}): Promise<{ isError?: boolean; structuredContent?: Record<string, unknown> }> {
  const r = await fetch(agent!.address, {
    method: 'POST',
    headers: { 'content-type': 'application/json', accept: 'application/json, text/event-stream', authorization: `Bearer ${agent!.token}` },
    body: JSON.stringify({ jsonrpc: '2.0', id: rpc++, method: 'tools/call', params: { name, arguments: args } }),
  });
  expect(r.status).toBe(200);
  return ((await r.json()) as { result: { isError?: boolean; structuredContent?: Record<string, unknown> } }).result;
}

/** MCP: List Servers, as its rows' text. */
async function mcpServers(): Promise<string[]> {
  void vs.exec('workbench.mcp.listServer').catch(() => undefined);
  await vs.page.locator('.quick-input-widget').waitFor({ state: 'visible' });
  await sleep(1000);
  const rows = await vs.page.evaluate(() => [...document.querySelectorAll('.quick-input-list .monaco-list-row')].map((e) => (e as HTMLElement).innerText.replace(/\s+/g, ' ')));
  await vs.page.keyboard.press('Escape');
  return rows;
}

const status = (frame: Frame) => frame.getByTestId('status');
const roomStatus = (frame: Frame) => frame.getByTestId('room-status');

/** The backups the extension wrote, in its workspace storage. */
async function backups(): Promise<string[]> {
  const { workspace } = await vs.ctl<{ workspace: string | null }>({ op: 'storagePath' });
  const folder = join(workspace!, '..', 'plutext.docx4j-editor');
  try { return readdirSync(folder).filter((n) => statSync(join(folder, n)).isFile()); } catch { return []; }
}

test('the MCP server: not offered before the person admits the agent; after the address command VS Code lists it, starts it and calls a tool (P40)', async () => {
  expect((await mcpServers()).some((r) => r.includes('docx4j editor'))).toBe(false);
  expect(companions()).toEqual([]);
  agent = await copyAddress();
  expect(companions()).toHaveLength(1);
  await expect.poll(async () => (await mcpServers()).some((r) => r.includes('docx4j editor')), { timeout: 15_000 }).toBe(true);
  await vs.exec('workbench.mcp.startServer', SERVER, { waitForLiveTools: true });
  // VS Code's own call, as its agent makes one: through the language model tools, VS Code asking the person first.
  // Named from the server's own name, `mcp_docx4j-editor_list_documents` in 1.147 (section 9.17 item 4); 1.105 names it otherwise.
  let tool = '';
  await expect.poll(async () => {
    tool = (await vs.ctl<{ name: string; tags: string[] }[]>({ op: 'lmTools' })).find((t) => t.name.endsWith('list_documents') && t.tags.includes('mcp'))?.name ?? '';
    return tool;
  }, { timeout: 15_000 }).not.toBe('');
  console.log(`VS Code ${vs.version} names the tool ${tool}`);
  const answer = vs.ctl<string>({ op: 'invokeTool', name: tool });
  await expect.poll(() => vs.dialog(), { timeout: 15_000 }).toContain('List the open documents');
  await sleep(400);
  await vs.clickDialogButton('Yes');
  expect(JSON.parse(await answer)).toEqual({ documents: [] });
  // And the same server as Claude Code reaches it, with the token.
  expect((await call('list_documents')).structuredContent).toEqual({ documents: [] });
  // Still one companion: VS Code's start went to the running one.
  expect(companions()).toHaveLength(1);
});

test('Invite from the page: the document shared into the companion\'s relay, the agent in the room, its edit a tracked change under its name', async () => {
  const frame = await vs.open('a.docx');
  await expect(frame.getByTestId('invite-agent')).toHaveText('Invite Agent');
  await frame.getByTestId('invite-agent').click();
  await expect(roomStatus(frame)).toHaveAttribute('data-status', 'connected', { timeout: 30_000 });
  await expect(status(frame)).toContainText('Agent joined', { timeout: 30_000 });
  await expect(frame.locator('[data-testid="person"][data-agent="true"]')).toHaveCount(1);
  const { documents } = (await call('list_documents')).structuredContent as { documents: { document: string; path: string | null }[] };
  expect(documents).toHaveLength(1);
  expect(documents[0]!.path).toBeNull();
  const inserted = await call('insert', { document: documents[0]!.document, text: 'Written by the agent', location: 'End' });
  expect(inserted.isError).not.toBe(true);
  await expect(frame.locator('.docx-editor .ProseMirror ins.insertion[data-author="Agent"]').first()).toContainText('Written by the agent', { timeout: 15_000 });
});

test('Stop in People: the agent leaves, and its next call answers agent.stopped', async () => {
  const frame = await vs.frameFor('a.docx');
  const { documents } = (await call('list_documents')).structuredContent as { documents: { document: string }[] };
  await frame.getByTestId('agent-stop').click();
  await expect(frame.locator('[data-testid="person"][data-agent="true"]')).toHaveCount(0, { timeout: 15_000 });
  const refused = await call('outline', { document: documents[0]!.document });
  expect(refused.structuredContent?.['code']).toBe('agent.stopped');
  // Invited again, it comes back: the restart case below needs it to have worked in this room.
  await frame.getByTestId('invite-agent').click();
  await expect(status(frame)).toContainText('Agent joined', { timeout: 30_000 });
  expect((await call('outline', { document: documents[0]!.document })).isError).not.toBe(true);
});

test('Share: the link opens in a browser beside VS Code, on the companion\'s address, and its typing reaches VS Code\'s page', async () => {
  const frame = await vs.open('b.docx');
  await frame.getByTestId('share').click();
  const link = await frame.getByTestId('share-link').inputValue({ timeout: 30_000 });
  expect(link.startsWith(agent!.address.replace(/\/mcp$/, '/#room='))).toBe(true);
  await expect(frame.getByTestId('share-local')).toBeVisible();
  await vs.page.keyboard.press('Escape');
  const browser = await chromium.launch();
  try {
    const page = await (await browser.newContext()).newPage();
    await page.goto(link);
    await page.getByTestId('join-accept').click({ timeout: 30_000 });
    await expect(page.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 30_000 });
    await page.evaluate(() => {
      const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { nodeSize: number } } }; caretAt(p: number): void; focus(): void } }).docx4jEditor;
      v.caretAt(v.state.doc.child(0).nodeSize - 1);
      v.focus();
    });
    await page.keyboard.type(' typed beside VS Code', { delay: 10 });
    await expect(frame.locator('.docx-editor .ProseMirror').first()).toContainText('typed beside VS Code', { timeout: 15_000 });
  } finally {
    await browser.close();
  }
});

test('Join with the link the agent\'s own open gives: an untitled editor in the room, dirty', async () => {
  const opened = await call('open', { path: vs.path('c.docx') });
  const link = (opened.structuredContent as { link: string }).link;
  void vs.exec('docx4j.join').catch(() => undefined);
  const input = vs.page.locator('.quick-input-widget input:not([type="checkbox"])');
  await input.waitFor();
  await input.fill(link);
  await sleep(300);
  await vs.page.keyboard.press('Enter');
  // Room.docx, named before the room gave its own: VS Code cannot rename an untitled editor, and the page takes its name.
  const frame = await vs.frameFor('Room.docx', 60_000);
  await expect(roomStatus(frame)).toHaveAttribute('data-status', 'connected', { timeout: 30_000 });
  await expect(frame.locator('[data-testid="person"][data-agent="true"]')).toHaveCount(1);
  await expect.poll(async () => (await vs.tab('Room.docx'))?.isDirty).toBe(true);
  // A second Join of the same link, left unedited: the restart below restores it in its room too (section 9.19 finding 1).
  void vs.exec('docx4j.join').catch(() => undefined);
  await input.waitFor();
  await input.fill(link);
  await sleep(300);
  await vs.page.keyboard.press('Enter');
  const second = await vs.frameFor('Room 2.docx', 60_000);
  await expect(roomStatus(second)).toHaveAttribute('data-status', 'connected', { timeout: 30_000 });
  await expect.poll(async () => (await vs.tab('Room 2.docx'))?.isDirty).toBe(true);
});

test('a kill and a restart: the rooms rejoined, the address unchanged, the agent out until invited again; the joined editor restored in its room', async () => {
  // VS Code's state is written a minute after the edits, before the kill.
  test.setTimeout(480_000);
  // A local edit in each, so that both are backed up.
  for (const name of ['a.docx', 'Room.docx']) {
    const frame = await vs.frameFor(name);
    await caretAtEndOf(frame, 0);
    await vs.page.keyboard.type(' before the kill', { delay: 20 });
  }
  // Room 2.docx is left as it joined: backed up with nothing of its own.
  await expect.poll(async () => (await backups()).length, { timeout: 20_000 }).toBeGreaterThanOrEqual(3);
  const before = agent!.address;
  await sleep(STATE_WRITTEN_MS);
  vs = await vs.restart('SIGKILL');
  for (const label of ['a.docx', 'Room.docx']) {
    await expect.poll(async () => (await vs.tab(label))?.isDirty, { timeout: 60_000, message: label }).toBe(true);
    const tab = await vs.tab(label);
    await vs.exec('vscode.openWith', { $parse: tab!.uri }, 'docx4j.editor');
    const frame = await vs.frameFor(label);
    await expect(roomStatus(frame), label).toHaveAttribute('data-status', 'connected', { timeout: 30_000 });
    await expect(frame.locator('.docx-editor .ProseMirror').first()).toContainText('before the kill');
  }
  // The unedited Join: joined again, the room's text, not a new empty document.
  await expect.poll(async () => (await vs.tab('Room 2.docx'))?.isDirty, { timeout: 60_000 }).toBe(true);
  await vs.exec('vscode.openWith', { $parse: (await vs.tab('Room 2.docx'))!.uri }, 'docx4j.editor');
  const unedited = await vs.frameFor('Room 2.docx');
  await expect(roomStatus(unedited)).toHaveAttribute('data-status', 'connected', { timeout: 30_000 });
  await expect(unedited.locator('.docx-editor .ProseMirror').first()).toContainText('before the kill');
  const a = await vs.frameFor('a.docx');
  await expect(a.locator('[data-testid="person"][data-agent="true"]')).toHaveCount(0);
  await expect(status(a)).toContainText('Agent is not in the room until you invite it again.', { timeout: 15_000 });
  agent = await copyAddress();
  expect(agent.address).toBe(before);
  // Its tab in front again: Room.docx was opened after it.
  await vs.open('a.docx');
  await a.getByTestId('invite-agent').click();
  await expect(a.locator('[data-testid="person"][data-agent="true"]')).toHaveCount(1, { timeout: 30_000 });
});

test('a Save As of the joined editor: the new editor is still in the room, and saved', async () => {
  const label = 'Room.docx';
  // The active editor is the one VS Code's Save As saves.
  const tab = await vs.tab(label);
  await vs.exec('vscode.openWith', { $parse: tab!.uri }, 'docx4j.editor');
  await vs.frameFor(label);
  await expect.poll(async () => (await vs.tab(label))?.isActive).toBe(true);
  await vs.saveAs(vs.path('room-saved.docx'));
  await vs.afterSaveAs(label, async () => (await vs.tab('room-saved.docx')) !== undefined && (await vs.tab(label)) === undefined);
  const frame = await vs.frameFor('room-saved.docx');
  await expect(roomStatus(frame)).toHaveAttribute('data-status', 'connected', { timeout: 30_000 });
  await expect(frame.getByTestId('name')).toHaveText('room-saved.docx');
  await expect.poll(async () => (await vs.tab('room-saved.docx'))?.isDirty).toBe(false);
});

test('the window closed: no companion is left', async () => {
  const running = companions();
  expect(running).toHaveLength(1);
  void vs.exec('workbench.action.quit').catch(() => undefined);
  await expect.poll(() => running.some(alive), { timeout: 30_000 }).toBe(false);
});

test('Restricted Mode: no Invite, no Share, no server, and the address command says why', async () => {
  await vs.close();
  vs = await VsCode.launch({ profile: 'agent-restricted', trust: true, files: { 'a.docx': 'HelloWordOnline.docx' }, settings: { 'chat.disableAIFeatures': false } });
  const frame = await vs.open('a.docx');
  await expect(frame.getByTestId('toolbar')).toBeVisible();
  await expect(frame.getByTestId('invite-agent')).toHaveCount(0);
  await expect(frame.getByTestId('share')).toHaveCount(0);
  void vs.exec('docx4j.copyAgentAddress').catch(() => undefined);
  await expect.poll(async () => (await vs.notifications()).join(' ')).toContain('Restricted Mode');
  expect((await mcpServers()).some((r) => r.includes('docx4j editor'))).toBe(false);
  expect(companions()).toEqual([]);
});
