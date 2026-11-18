// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 8.9, step F1: the companion serves the application, the relay and the probe on
// loopback, refuses another `Host` (DNS rebinding), opens only files under its roots, and shares a
// file into a room that a page joins through the printed link; the agent is in the room's
// awareness as a page's user is, and what it writes reaches the page as its tracked change.
import { afterEach, describe, expect, test } from 'vitest';
import { mkdtemp, mkdir, rm, symlink, writeFile, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { request } from 'node:http';
import { createServer as createNetServer } from 'node:net';
import * as Y from 'yjs';
import WebSocket from 'ws';
import { HocuspocusProvider, HocuspocusProviderWebsocket } from '@hocuspocus/provider';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import { PlainText, Texts } from '@docx4j/core-ts';
import { EditorDocument, newEditStore, agentLog, stopAgent, isAgentStopped, runCommand, trackedChangesOf } from '@docx4j/editor-model';
import { createCompanion, existingUnderRoots, PathError, type Companion } from '../src/index.mjs';
import { markedPage } from '../src/server.mjs';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');
const cleanups: (() => Promise<void> | void)[] = [];
afterEach(async () => { for (const c of cleanups.splice(0).reverse()) await c(); });

/** A companion on a free port over a fresh root holding `tracked-changes.docx`, with a one-page application. */
async function start(): Promise<{ companion: Companion; base: string; root: string }> {
  const root = await mkdtemp(join(tmpdir(), 'companion-root-'));
  const app = await mkdtemp(join(tmpdir(), 'companion-app-'));
  const data = await mkdtemp(join(tmpdir(), 'companion-rooms-'));
  await writeFile(join(app, 'index.html'), '<!doctype html><title>docx4j editor</title>');
  await writeFile(join(root, 'lease.docx'), await readFile(join(fixtures, 'tracked-changes.docx')));
  await writeFile(join(root, 'notes.docx'), await readFile(join(fixtures, 'footnote.docx')));
  const lines: string[] = [];
  const companion = createCompanion({ roots: [root], port: 0, app, data, log: (l) => lines.push(l) });
  const base = await companion.listen();
  cleanups.push(() => rm(root, { recursive: true, force: true }), () => rm(app, { recursive: true, force: true }), () => rm(data, { recursive: true, force: true }));
  cleanups.push(() => companion.close());
  return { companion, base, root };
}

/** A GET with a given `Host`, as a page on another name that resolves here would send. */
function get(base: string, path: string, host?: string, extra: { method?: string; headers?: Record<string, string>; body?: string } = {}): Promise<{ status: number; body: string }> {
  const url = new URL(path, base);
  return new Promise((resolve, reject) => {
    const req = request({ hostname: url.hostname, port: url.port, path: url.pathname, method: extra.method ?? 'GET', headers: { ...(host ? { Host: host } : {}), ...(extra.headers ?? {}) } }, (res) => {
      let body = '';
      res.on('data', (c: Buffer) => { body += c.toString(); });
      res.on('end', () => resolve({ status: res.statusCode ?? 0, body }));
    });
    req.on('error', reject);
    req.end(extra.body);
  });
}

/** A page's connection to a room, as `app/room.ts` makes one, from the parts of a room link. */
function page(link: string): { doc: Y.Doc; provider: HocuspocusProvider; synced: Promise<void> } {
  const params = new URLSearchParams(link.slice(link.indexOf('#') + 1));
  const doc = new Y.Doc();
  const socket = new HocuspocusProviderWebsocket({ url: params.get('relay')!, WebSocketPolyfill: WebSocket, maxAttempts: 1 });
  let synced!: () => void;
  const out = { doc, provider: undefined as unknown as HocuspocusProvider, synced: new Promise<void>((r) => { synced = r; }) };
  out.provider = new HocuspocusProvider({ websocketProvider: socket, name: params.get('room')!, document: doc, token: params.get('token')!, onSynced: () => synced() });
  out.provider.attach();
  cleanups.push(() => { out.provider.destroy(); socket.destroy(); socket.connect = async () => undefined; });
  return out;
}

async function until(check: () => boolean, ms = 5000, what = 'a condition'): Promise<void> {
  const end = Date.now() + ms;
  while (!check()) {
    if (Date.now() > end) throw new Error(`timed out waiting for ${what}`);
    await new Promise((r) => setTimeout(r, 20));
  }
}

// The third review point of ED-005 (section 12.47): a relay worker that stops after starting is noticed.
describe('the relay watched after it starts', () => {
  async function client(companion: Companion): Promise<Client> {
    const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
    await companion.mcp().connect(serverSide);
    const c = new Client({ name: 'relay-test', version: '1' });
    await c.connect(clientSide);
    cleanups.push(() => c.close());
    return c;
  }

  test('a worker that stops is restarted on its port: the room link holds, the agent reconnects, and its edits reach a page again', async () => {
    const { companion } = await start();
    const c = await client(companion);
    const opened = await c.callTool({ name: 'open', arguments: { path: 'lease.docx' } }) as unknown as { structuredContent: { document: string; link: string } };
    const { document, link } = opened.structuredContent;
    const relay = companion.relay;
    await companion.crashRelay();
    await until(() => companion.relayState === 'up' && companion.relay === relay, 10_000, 'the relay restarted on its port');
    // A page joining by the same link, retrying until the room is back (the agent's own connection remakes it).
    const params = new URLSearchParams(link.slice(link.indexOf('#') + 1));
    expect(params.get('relay')).toBe(relay);
    const doc = new Y.Doc();
    const socket = new HocuspocusProviderWebsocket({ url: relay, WebSocketPolyfill: WebSocket });
    let synced = false;
    const provider = new HocuspocusProvider({ websocketProvider: socket, name: params.get('room')!, document: doc, token: params.get('token')!, onSynced: () => { synced = true; } });
    provider.attach();
    cleanups.push(() => { provider.destroy(); socket.destroy(); socket.connect = async () => undefined; });
    // The relay writes a room to its file up to two seconds after a change, so a room made just before the stop
    // comes back from its peers: the agent's connection brings the whole document when it reconnects.
    const store = newEditStore(doc);
    await until(() => synced && store.room.get('base') !== undefined, 15_000, 'the document back in the restarted relay');
    const joined = await EditorDocument.open(store.room.get('base') as Uint8Array, { store, mode: 'join' });
    const inserted = await c.callTool({ name: 'insert', arguments: { document, text: 'After the restart', location: 'End' } }) as { isError?: boolean };
    expect(inserted.isError).not.toBe(true);
    await until(() => joined.doc.textContent.includes('After the restart'), 10_000, "the agent's edit on the page");
  }, 60_000);

  test('restarts that fail leave it down: open and an edit refused with relay.down, a read still answered from the agent\'s copy', async () => {
    const { companion } = await start();
    const c = await client(companion);
    const opened = await c.callTool({ name: 'open', arguments: { path: 'lease.docx' } }) as unknown as { structuredContent: { document: string } };
    const { document } = opened.structuredContent;
    const port = Number(new URL(companion.relay).port);
    await companion.crashRelay();
    // The port taken before the first restart (100 ms later), so that every restart fails.
    // It drops what it accepts: the agent's connection keeps retrying the port, and a held socket would keep it open.
    const blocker = createNetServer((socket) => socket.destroy());
    await new Promise<void>((resolve, reject) => { blocker.once('error', reject); blocker.listen(port, '127.0.0.1', () => resolve()); });
    cleanups.push(() => new Promise<void>((resolve) => blocker.close(() => resolve())));
    await until(() => companion.relayState === 'down', 10_000, 'the relay left down');
    const inserted = await c.callTool({ name: 'insert', arguments: { document, text: 'Nowhere', location: 'End' } }) as { isError?: boolean; structuredContent: { code?: string } };
    expect(inserted.isError).toBe(true);
    expect(inserted.structuredContent.code).toBe('relay.down');
    const again = await c.callTool({ name: 'open', arguments: { path: 'lease.docx' } }) as { isError?: boolean; structuredContent: { code?: string } };
    expect(again.structuredContent.code).toBe('relay.down');
    const read = await c.callTool({ name: 'outline', arguments: { document } }) as { isError?: boolean };
    expect(read.isError).not.toBe(true);
  }, 60_000);
});

describe('the companion, step F1 (ED-005 section 8.9)', () => {
  test('marks the page it serves inside its head, so that the page probes it (release plan D8)', () => {
    const meta = '<meta name="docx4j-companion" content="/companion">';
    expect(markedPage('<!doctype html>\n<html lang="en">\n  <head>\n    <meta charset="UTF-8" />\n  </head>')).toBe(`<!doctype html>\n<html lang="en">\n  <head>${meta}\n    <meta charset="UTF-8" />\n  </head>`);
    expect(markedPage('<!DOCTYPE html><HEAD class="x"><title>t</title>')).toBe(`<!DOCTYPE html><HEAD class="x">${meta}<title>t</title>`);
    expect(markedPage('<!doctype html><title>docx4j editor</title>')).toBe(`<!doctype html>${meta}<title>docx4j editor</title>`);
    // `<header>` is not the head.
    expect(markedPage('<body><header>h</header>')).toBe(`${meta}<body><header>h</header>`);
  });

  test('serves the application and the probe on loopback, and refuses another Host', async () => {
    const { base } = await start();
    expect((await get(base, '/')).body).toContain('docx4j editor');
    // The page it serves is marked, so that the page probes it; a static site's page is not, and probes nothing.
    expect((await get(base, '/')).body).toContain('<meta name="docx4j-companion" content="/companion">');
    expect((await get(base, '/index.html')).body).toContain('<meta name="docx4j-companion" content="/companion">');
    const probe = JSON.parse((await get(base, '/companion/')).body) as { tools: string[]; relay: string; agent: string };
    // The relay on a loopback port of its own, in a worker (ED-005 section 9.3 item 1 (i)).
    expect(probe).toMatchObject({ tools: [], agent: 'Claude' });
    expect(probe.relay).toMatch(/^ws:\/\/127\.0\.0\.1:\d+$/);
    expect(probe.relay).not.toBe(base.replace(/^http/, 'ws'));
    expect((await get(base, '/', 'attacker.example:80')).status).toBe(421);
    expect((await get(base, '/companion/', 'attacker.example')).status).toBe(421);
  });

  test('opens a file only under a root: a relative path from the first, an absolute one inside, never one outside or through a symlink out', async () => {
    const { root } = await start();
    const outside = await mkdtemp(join(tmpdir(), 'companion-outside-'));
    cleanups.push(() => rm(outside, { recursive: true, force: true }));
    await writeFile(join(outside, 'secret.docx'), 'x');
    await mkdir(join(root, 'sub'));
    await symlink(join(outside, 'secret.docx'), join(root, 'sub', 'link.docx'));
    await expect(existingUnderRoots([root], 'lease.docx')).resolves.toMatch(/lease\.docx$/);
    await expect(existingUnderRoots([root], join(root, 'lease.docx'))).resolves.toMatch(/lease\.docx$/);
    await expect(existingUnderRoots([root], join(outside, 'secret.docx'))).rejects.toMatchObject({ code: 'outside_root' });
    await expect(existingUnderRoots([root], 'sub/link.docx')).rejects.toMatchObject({ code: 'outside_root' });
    await expect(existingUnderRoots([root], '../x.docx')).rejects.toBeInstanceOf(PathError);
    await expect(existingUnderRoots([], 'lease.docx')).rejects.toMatchObject({ code: 'outside_root' });
  });

  test('shares a file into a room a page joins through the printed link: the same document, the agent in the awareness, and its tracked edit reaching the page', async () => {
    const { companion, base } = await start();
    const room = await companion.open('lease.docx');
    expect(room.link.startsWith(`${base}/#room=${room.id}&token=`)).toBe(true);
    const p = page(room.link);
    await p.synced;
    // The page's document is the agent's.
    const store = newEditStore(p.doc);
    const joined = await EditorDocument.open(store.room.get('base') as Uint8Array, { store, mode: 'join' });
    expect(joined.doc.textContent).toBe(room.session.doc.textContent);
    // The agent is a user in the room's awareness, as a page's is.
    await until(() => [...p.provider.awareness!.getStates().values()].some((s) => (s['user'] as { name?: string } | undefined)?.name === 'Claude'), 5000, `the agent in the page's awareness (states: ${JSON.stringify([...p.provider.awareness!.getStates().entries()])})`);
    const agent = [...p.provider.awareness!.getStates().values()].find((s) => (s['user'] as { name?: string } | undefined)?.name === 'Claude')!;
    expect(agent['user']).toMatchObject({ name: 'Claude', agent: true });
    // What the agent writes reaches the page as its tracked change.
    const run = await room.session.runOnTree(({ body }) => { body.insertParagraph('The agent was here.', 'End'); }, 'agent', { author: 'Claude', mode: 'TrackAll' });
    expect(run.error).toBeUndefined();
    await until(() => joined.doc.textContent.includes('The agent was here.'), 5000, "the agent's edit on the page");
    let tracked = false;
    joined.doc.descendants((n) => { if (n.isText && n.text!.includes('The agent was here') && n.marks.some((m) => m.type.name === 'insertion' && m.attrs['author'] === 'Claude')) tracked = true; return !tracked; });
    expect(tracked).toBe(true);
    // And the page's typing reaches the agent.
    const at = joined.doc.child(1).nodeSize;
    joined.replace(joined.doc.type.create(joined.doc.attrs, joined.doc.content.replaceChild(0, joined.doc.child(0).type.create(joined.doc.child(0).attrs, joined.doc.type.schema.text('Retitled by the page')))));
    expect(at).toBeGreaterThan(0);
    await until(() => room.session.doc.child(0).textContent === 'Retitled by the page', 5000, "the page's edit at the agent");
  });

  // ED-005 proposal 42, section 9.3 item 1 (ii): the person's typing reaches the paragraph an agent's call is changing,
  // while the call runs. The call's run is wrapped once here so that the typing is sent as it starts, and has reached the
  // companion's socket (the relay forwards from its own thread) by the time the run lets input in before its merge.
  test('a call meets the person\'s typing in its paragraph while it runs: refused with target.stale, nothing written, the words the person\'s; the agent joins afresh, and the retry lands', async () => {
    const { companion } = await start();
    const room = await companion.open('lease.docx');
    const p = page(room.link);
    await p.synced;
    const store = newEditStore(p.doc);
    const person = await EditorDocument.open(store.room.get('base') as Uint8Array, { store, mode: 'join' });
    person.trackingAuthor = 'Ada';
    const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
    await companion.mcp().connect(serverSide);
    const client = new Client({ name: 'guard-test', version: '1' });
    await client.connect(clientSide);
    cleanups.push(() => client.close());
    const call = async (name: string, args: Record<string, unknown>) => client.callTool({ name, arguments: args }) as Promise<{ isError?: boolean; structuredContent?: Record<string, unknown> }>;
    const block = (s: EditorDocument): number => { let at = -1; s.doc.forEach((n, pos) => { if (at < 0 && n.textContent.includes('Some text')) at = pos + 1; }); return at; };
    const first = room.session;
    const real = first.runOnTree.bind(first);
    first.runOnTree = (async (...args: Parameters<typeof real>) => {
      expect(runCommand(person, 'insertText', { text: 'new ' }, { from: block(person) + 5 }).ok).toBe(true);
      const until = Date.now() + 30;
      while (Date.now() < until) { /* the companion's thread held, as a call's own work holds it; the relay forwards meanwhile */ }
      return real(...args);
    }) as typeof first.runOnTree;
    const refused = await call('edit', { document: room.id, address: { contains: 'Some text' }, text: 'Other words.' });
    expect(refused.isError).toBe(true);
    expect(refused.structuredContent).toMatchObject({ code: 'target.stale', current: { address: expect.stringMatching(/^body\/\d+$/), text: 'Some new text.' } });
    // Nothing written: the person's words stand, with no change of Claude's; the agent's session is a new one, holding them.
    await new Promise((r) => setTimeout(r, 200));
    expect(person.doc.textContent).toContain('Some new text.');
    expect(trackedChangesOf(person.doc).filter((c) => c.author === 'Claude')).toEqual([]);
    expect(room.session).not.toBe(first);
    await until(() => room.session.doc.textContent.includes('Some new text.'), 5000, 'the person\'s words in the new session');
    await expect(first.runOnTree(() => undefined)).rejects.toThrow(/spent/);
    // The agent reads again and redoes it: the retry lands as its tracked change, on both sides.
    const retry = await call('edit', { document: room.id, address: { contains: 'Some new text' }, text: 'Other words.', if_text: 'Some new text.' });
    expect(retry.isError, JSON.stringify(retry.structuredContent)).toBeFalsy();
    await until(() => trackedChangesOf(person.doc).some((c) => c.author === 'Claude' && c.type === 'Added' && c.text.includes('Other words.')), 5000, 'the retry at the person');
  });

  test('MCP over streamable HTTP under /mcp: the tools listed and called; a page of another origin refused', async () => {
    const { companion, base } = await start();
    const client = new Client({ name: 'http-test', version: '1' });
    await client.connect(new StreamableHTTPClientTransport(new URL(`${base}/mcp`)));
    cleanups.push(() => client.close());
    const { tools } = await client.listTools();
    expect(tools.map((t) => t.name)).toContain('replace_text');
    const opened = await client.callTool({ name: 'open', arguments: { path: 'lease.docx' } }) as unknown as { structuredContent: { document: string } };
    expect(companion.rooms.has(opened.structuredContent.document)).toBe(true);
    const refused = await get(base, '/mcp', undefined, { method: 'POST', headers: { Origin: 'http://localhost:9999', 'Content-Type': 'application/json', Accept: 'application/json, text/event-stream' }, body: JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'tools/list' }) });
    expect(refused.status).toBe(403);
  });

  test('in the room: the agent\'s tool call reaches a page as its tracked change, and the page\'s typing reaches the agent\'s next read', async () => {
    const { companion } = await start();
    const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
    await companion.mcp().connect(serverSide);
    const client = new Client({ name: 'room-test', version: '1' });
    await client.connect(clientSide);
    cleanups.push(() => client.close());
    const opened = await client.callTool({ name: 'open', arguments: { path: 'lease.docx' } }) as unknown as { structuredContent: { document: string; link: string } };
    const document = opened.structuredContent.document;
    const p = page(opened.structuredContent.link);
    await p.synced;
    const store = newEditStore(p.doc);
    const joined = await EditorDocument.open(store.room.get('base') as Uint8Array, { store, mode: 'join' });
    const inserted = await client.callTool({ name: 'insert', arguments: { document, text: 'From the agent', location: 'End' } }) as { isError?: boolean };
    expect(inserted.isError).not.toBe(true);
    await until(() => joined.doc.textContent.includes('From the agent'), 5000, "the agent's paragraph on the page");
    let tracked = false;
    joined.doc.descendants((n) => { if (n.isText && n.text === 'From the agent' && n.marks.some((m) => m.type.name === 'insertion' && m.attrs['author'] === 'Claude')) tracked = true; return !tracked; });
    expect(tracked).toBe(true);
    // The page retitles the document; the agent reads it as it is now.
    joined.replace(joined.doc.type.create(joined.doc.attrs, joined.doc.content.replaceChild(0, joined.doc.child(0).type.create(joined.doc.child(0).attrs, joined.doc.type.schema.text('Retitled on the page')))));
    const room = companion.rooms.get(document)!;
    await until(() => room.session.doc.child(0).textContent === 'Retitled on the page', 5000, "the page's edit at the agent");
    const read = await client.callTool({ name: 'read', arguments: { document, address: 'body/0' } }) as unknown as { structuredContent: { markdown: string } };
    expect(read.structuredContent.markdown).toBe('<!-- body/0 -->\nRetitled on the page');
  });

  test('the browser first: a page shares a room on the companion\'s relay and invites the agent; each call is logged in the room; a person\'s Stop ends its calls there; inviting again brings it back', async () => {
    const { companion, base } = await start();
    const bytes = await readFile(join(fixtures, 'tracked-changes.docx'));
    // The page shares, as the application's Share does: a room made on the relay, the document written to it.
    const doc = new Y.Doc();
    const session = await EditorDocument.open(new Uint8Array(bytes), { store: newEditStore(doc) });
    const roomId = 'invited-room-0123456789abcdef';
    const token = 'a-token-for-the-invited-room-0123456789';
    const socket = new HocuspocusProviderWebsocket({ url: companion.relay, WebSocketPolyfill: WebSocket, maxAttempts: 1 });
    let synced!: () => void;
    const ready = new Promise<void>((r) => { synced = r; });
    const provider = new HocuspocusProvider({ websocketProvider: socket, name: roomId, document: doc, token: `create:${token}`, onSynced: () => synced() });
    provider.attach();
    cleanups.push(() => { provider.destroy(); socket.destroy(); socket.connect = async () => undefined; });
    await ready;
    await session.share(new Uint8Array(bytes), 'lease.docx');
    const post = (body: unknown, origin?: string) => get(base, '/companion/rooms', undefined, { method: 'POST', headers: { 'Content-Type': 'application/json', ...(origin ? { Origin: origin } : {}) }, body: JSON.stringify(body) });
    expect((await post({ room: roomId, token: 'wrong' })).status).toBe(403);
    expect((await post({ room: 'no-such-room-0123456789', token })).status).toBe(404);
    expect((await post({ room: roomId, token }, 'http://attacker.example')).status).toBe(403);
    const invited = await post({ room: roomId, token }, base);
    expect(invited.status, invited.body).toBe(200);
    expect(JSON.parse(invited.body)).toEqual({ document: roomId, agent: 'Claude' });
    const agentIn = (): boolean => [...provider.awareness!.getStates().values()].some((st) => (st['user'] as { agent?: boolean } | undefined)?.agent === true);
    await until(agentIn, 5000, 'the agent in the page\'s awareness');
    // A call through MCP: the page sees Claude's insertion, and the room's log has the call.
    const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
    await companion.mcp().connect(serverSide);
    const client = new Client({ name: 'invite-test', version: '1' });
    await client.connect(clientSide);
    cleanups.push(() => client.close());
    const inserted = await client.callTool({ name: 'insert', arguments: { document: roomId, text: 'Invited and writing', location: 'End' } }) as { isError?: boolean };
    expect(inserted.isError).not.toBe(true);
    await until(() => session.doc.textContent.includes('Invited and writing'), 5000, 'the agent\'s paragraph on the page');
    await until(() => agentLog(session.store).length > 0, 5000, 'the log on the page');
    expect(agentLog(session.store).at(-1)).toMatchObject({ agent: 'Claude', tool: 'insert', summary: 'Inserted a paragraph at the end of the body', ok: true });
    // A room the agent was invited into has no file of its own: save needs output_path.
    const saved = await client.callTool({ name: 'save', arguments: { document: roomId } }) as { isError?: boolean; structuredContent?: { code?: string } };
    expect(saved.structuredContent?.code).toBe('arguments.invalid');
    // The person's Stop, through the room: the agent leaves and its calls there are refused.
    stopAgent(session.store, 'Claude');
    await until(() => companion.rooms.get(roomId)?.stopped === true, 5000, 'the companion hearing the Stop');
    await until(() => !agentIn(), 5000, 'the agent gone from the page\'s awareness');
    const refused = await client.callTool({ name: 'outline', arguments: { document: roomId } }) as { isError?: boolean; structuredContent?: { code?: string } };
    expect(refused.structuredContent?.code).toBe('agent.stopped');
    // Invited again, it comes back, the Stop cleared.
    expect((await post({ room: roomId, token }, base)).status).toBe(200);
    await until(() => !isAgentStopped(session.store, 'Claude'), 5000, 'the Stop cleared on the page');
    await until(agentIn, 5000, 'the agent back in the page\'s awareness');
    const again = await client.callTool({ name: 'outline', arguments: { document: roomId } }) as { isError?: boolean };
    expect(again.isError).not.toBe(true);
  });

  test('a Stop covers the document: its file and a copy the agent saved refused by open, leaving the room too; another file opens; the person\'s invitation brings it back with the room\'s file (ED-005 section 12.29)', async () => {
    const { companion, base, root } = await start();
    await writeFile(join(root, 'other.docx'), await readFile(join(fixtures, 'tracked-changes.docx')));
    const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
    await companion.mcp().connect(serverSide);
    const client = new Client({ name: 'stop-test', version: '1' });
    await client.connect(clientSide);
    cleanups.push(() => client.close());
    type Answer = { isError?: boolean; structuredContent?: Record<string, unknown> };
    const call = async (name: string, args: Record<string, unknown>): Promise<Answer> => client.callTool({ name, arguments: args }) as Promise<Answer>;
    // The agent opens the file and saves a copy; a person joins through the link.
    const opened = (await call('open', { path: 'lease.docx' })).structuredContent!;
    const document = opened['document'] as string;
    const link = opened['link'] as string;
    expect((await call('save', { document, output_path: 'lease-reviewed.docx' })).isError).not.toBe(true);
    const p = page(link);
    await p.synced;
    const store = newEditStore(p.doc);
    await EditorDocument.open(store.room.get('base') as Uint8Array, { store, mode: 'join' });
    // The person's Stop.
    stopAgent(store, 'Claude');
    await until(() => companion.rooms.get(document)?.stopped === true, 5000, 'the companion hearing the Stop');
    // What Claude Code did in the by-hand run: open the file again, then the copy. Both are refused, and the hint offers no way round.
    for (const path of ['lease.docx', 'lease-reviewed.docx']) {
      const refused = await call('open', { path });
      expect(refused.isError).toBe(true);
      expect(refused.structuredContent).toMatchObject({ code: 'agent.stopped', document });
      expect(String(refused.structuredContent!['hint'])).not.toMatch(/open the file again/);
      expect(String(refused.structuredContent!['hint'])).toContain('do not open its file or a copy of it again');
    }
    expect(companion.rooms.size).toBe(1);
    // Leaving the stopped room does not lift the Stop.
    expect((await call('close', { document })).isError).not.toBe(true);
    expect((await call('open', { path: 'lease.docx' })).structuredContent).toMatchObject({ code: 'agent.stopped' });
    // Another file is not the stopped document.
    const other = await call('open', { path: 'other.docx' });
    expect(other.isError).not.toBe(true);
    // The person invites the agent back from the page: the same room, its file kept, the copy open again.
    const params = new URLSearchParams(link.slice(link.indexOf('#') + 1));
    const invited = await get(base, '/companion/rooms', undefined, { method: 'POST', headers: { 'Content-Type': 'application/json', Origin: base }, body: JSON.stringify({ room: document, token: params.get('token') }) });
    expect(invited.status, invited.body).toBe(200);
    await until(() => !isAgentStopped(store, 'Claude'), 5000, 'the Stop cleared on the page');
    const back = companion.rooms.get(document)!;
    expect(back.path).toMatch(/lease\.docx$/);
    expect((await call('open', { path: 'lease.docx' })).structuredContent).toMatchObject({ document });
    const saved = await call('save', { document });
    expect(saved.isError, JSON.stringify(saved.structuredContent)).not.toBe(true);
    expect((await call('open', { path: 'lease-reviewed.docx' })).isError).not.toBe(true);
  });
});

// ED-008 step 4: the reading tools in a room, a person's page in it. What the agent reads is the room's document as
// it is when the call runs (the person's typing and their notes included), in the view asked for; what if_text is
// compared with is the accepted text, whatever view the agent last read (docs/developer/agent.md sections 2.2 and 2.13).
describe('the reading tools in a room (ED-008 step 4)', () => {
  async function inRoomWithPerson(file: string) {
    const { companion } = await start();
    const room = await companion.open(file);
    const p = page(room.link);
    await p.synced;
    const store = newEditStore(p.doc);
    const person = await EditorDocument.open(store.room.get('base') as Uint8Array, { store, mode: 'join' });
    person.trackingAuthor = 'Ada';
    const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
    await companion.mcp().connect(serverSide);
    const client = new Client({ name: 'rooms-reading-test', version: '1' });
    await client.connect(clientSide);
    cleanups.push(() => client.close());
    const call = async (name: string, args: Record<string, unknown>) => {
      const r = await client.callTool({ name, arguments: { document: room.id, ...args } }) as { isError?: boolean; structuredContent?: Record<string, unknown> };
      return { isError: r.isError === true, data: r.structuredContent ?? {} };
    };
    const text = async (view: string): Promise<string> => {
      const r = await call('extract_text', { view });
      expect(r.isError, JSON.stringify(r.data)).toBe(false);
      return r.data['text'] as string;
    };
    return { room, person, call, text };
  }

  test('the agent\'s own pending insertion is in the accepted text, marked in markup, out of the original; the person\'s session reads the same', async () => {
    const { room, person, call, text } = await inRoomWithPerson('lease.docx');
    const inserted = await call('insert', { text: 'The agent was here.', location: 'End' });
    expect(inserted.isError, JSON.stringify(inserted.data)).toBe(false);
    expect(await text('accepted')).toContain('\nThe agent was here.\n');
    expect(await text('markup')).toContain('{++The agent was here.++}');
    expect(await text('original')).not.toContain('The agent was here.');
    // The person's page, once the agent's edit reached it, reads the same text through the same model.
    await until(() => person.doc.textContent.includes('The agent was here.'), 5000, "the agent's insertion on the page");
    for (const view of ['accepted', 'markup', 'original'] as const) {
      const theirs = await person.readTree(async ({ textPackage }) => PlainText.of(await Texts.of(textPackage, { view })), { stories: true });
      expect(theirs, view).toBe(await text(view));
    }
    expect(room.session.doc.textContent).toBe(person.doc.textContent);
  });

  test('the person\'s typing between two reads: the second read has it, in every view, and find gives it a span', async () => {
    const { room, person, call, text } = await inRoomWithPerson('lease.docx');
    const before = await text('accepted');
    expect(before).not.toContain('Ada typed this');
    let at = -1;
    person.doc.forEach((n, pos) => { if (at < 0 && n.textContent === 'Some text.') at = pos + 1; });
    expect(runCommand(person, 'insertText', { text: 'Ada typed this. ' }, { from: at }).ok).toBe(true);
    await until(() => room.session.doc.textContent.includes('Ada typed this.'), 5000, "the person's typing at the agent");
    for (const view of ['accepted', 'original', 'markup']) expect(await text(view), view).toContain('Ada typed this. Some text.');
    const found = await call('find', { text: 'Ada typed' });
    expect(found.data['hits']).toEqual([expect.objectContaining({ offset: 0, length: 9, snippet: 'Ada typed this. Some text.' })]);
    const outline = (await call('outline', {})).data['blocks'] as { address: string; text: string }[];
    expect(outline.find((b) => b.text === 'Ada typed this. Some text.')?.address).toBe((found.data['hits'] as { address: string }[])[0]!.address);
  });

  test('a note the person edits is in the agent\'s next read: the session\'s stories written into their parts first', async () => {
    const { room, person, text } = await inRoomWithPerson('notes.docx');
    const note = (await person.stories.footnotes()).find((n) => !n.orphan && n.doc.textContent.includes('Footnote 1'))!;
    let at = 1;
    note.doc.firstChild!.forEach((child, offset) => { if (at === 1 && child.isText) at = 1 + offset; });
    expect(runCommand(person, 'insertText', { text: 'Ada noted ' }, { from: at }, note).ok).toBe(true);
    await until(() => room.session.store.footnotes.get(String(note.id))?.toString().includes('Ada noted') ?? false, 5000, "the person's note at the agent");
    expect(await text('accepted')).toContain('Ada noted');
    // A read writes nothing the person would see: the agent's session saves as the room holds the document.
    expect((await room.session.save()).selfCheck?.equal).toBe(true);
  });

  test('if_text after a markup read: the marked text is refused as stale, with the accepted text, which then lands', async () => {
    const { call } = await inRoomWithPerson('lease.docx');
    const markup = (await call('outline', { view: 'markup' })).data['blocks'] as { address: string; text: string }[];
    const marked = markup.find((b) => b.text.includes('{++An insertion++}'))!;
    const refused = await call('set_format', { address: marked.address, anchor: 'change tracking', bold: true, if_text: marked.text });
    expect(refused.isError).toBe(true);
    expect(refused.data).toMatchObject({ code: 'target.stale', current: { address: marked.address, text: 'Here is some change tracking. An insertion Followed by.' } });
    const accepted = (await call('outline', {})).data['blocks'] as { address: string; text: string }[];
    const landed = await call('set_format', { address: marked.address, anchor: 'change tracking', bold: true, if_text: accepted.find((b) => b.address === marked.address)!.text });
    expect(landed.isError, JSON.stringify(landed.data)).toBe(false);
  });
});
