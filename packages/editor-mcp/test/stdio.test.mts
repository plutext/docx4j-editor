// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 8.9, step F5: the scripted client end to end. The bin (`dist/cli.mjs`, built) is
// started over stdio as an MCP client starts it, stdout the protocol's alone; the agent opens a file
// and a page joins its room through the link; the agent reads, edits by Markdown with if_text, is
// refused on a paragraph the person changed and decides again, comments; the person accepts one of
// its changes; the agent lists what is left and saves; the file holds what the room holds.
import { afterEach, describe, expect, test } from 'vitest';
import { mkdtemp, readFile, rm, stat, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import * as Y from 'yjs';
import WebSocket from 'ws';
import { HocuspocusProvider, HocuspocusProviderWebsocket } from '@hocuspocus/provider';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, newEditStore, runCommand, trackedChangesOf } from '@docx4j/editor-model';

const here = dirname(fileURLToPath(import.meta.url));
const fixtures = join(here, '..', '..', 'editor-model', 'test', 'fixtures');
const cli = join(here, '..', 'dist', 'cli.mjs');
const cleanups: (() => Promise<void> | void)[] = [];
afterEach(async () => { for (const c of cleanups.splice(0).reverse()) await c(); });

async function until(check: () => boolean, ms = 10_000, what = 'a condition'): Promise<void> {
  const end = Date.now() + ms;
  while (!check()) {
    if (Date.now() > end) throw new Error(`timed out waiting for ${what}`);
    await new Promise((r) => setTimeout(r, 20));
  }
}

type Structured = Record<string, unknown>;

describe('the scripted client over stdio (ED-005 section 8.9, step F5)', () => {
  test('started through a link, as npm\'s bin and npx start it, the companion answers', async () => {
    // 0.1.0 compared process.argv[1] with its own file, so through node_modules/.bin it exited without a word.
    try { await stat(cli); } catch { throw new Error(`build the companion first (npm run build -w @docx4j/editor-mcp): no ${cli}`); }
    const bin = await mkdtemp(join(tmpdir(), 'stdio-bin-'));
    const root = await mkdtemp(join(tmpdir(), 'stdio-root-'));
    const data = await mkdtemp(join(tmpdir(), 'stdio-rooms-'));
    cleanups.push(() => rm(bin, { recursive: true, force: true }), () => rm(root, { recursive: true, force: true }), () => rm(data, { recursive: true, force: true }));
    const link = join(bin, 'docx4j-editor-mcp');
    await symlink(cli, link);
    const transport = new StdioClientTransport({ command: process.execPath, args: [link, '--root', root, '--port', '0', '--data', data], stderr: 'pipe' });
    const client = new Client({ name: 'linked', version: '1' });
    await client.connect(transport);
    cleanups.push(() => client.close());
    const { tools } = await client.listTools();
    expect(tools.map((t) => t.name)).toContain('open');
  });

  test('an agent\'s session with a person in the room, from open to save', async () => {
    try { await stat(cli); } catch { throw new Error(`build the companion first (npm run build -w @docx4j/editor-mcp): no ${cli}`); }
    const root = await mkdtemp(join(tmpdir(), 'stdio-root-'));
    const data = await mkdtemp(join(tmpdir(), 'stdio-rooms-'));
    const app = await mkdtemp(join(tmpdir(), 'stdio-app-'));
    await writeFile(join(app, 'index.html'), '<!doctype html><title>docx4j editor</title>');
    await writeFile(join(root, 'lease.docx'), await readFile(join(fixtures, 'tracked-changes.docx')));
    cleanups.push(() => rm(root, { recursive: true, force: true }), () => rm(data, { recursive: true, force: true }), () => rm(app, { recursive: true, force: true }));

    // The companion as Claude Code starts it: the bin over stdio, stderr its log.
    const transport = new StdioClientTransport({ command: process.execPath, args: [cli, '--root', root, '--port', '0', '--data', data, '--app', app], stderr: 'pipe' });
    let log = '';
    transport.stderr?.on('data', (chunk: Buffer) => { log += chunk.toString(); });
    const protocolErrors: Error[] = [];
    const client = new Client({ name: 'scripted-agent', version: '1' });
    client.onerror = (e) => { protocolErrors.push(e); };
    await client.connect(transport);
    cleanups.push(() => client.close());
    const call = async (name: string, args: Structured): Promise<{ isError: boolean; data: Structured }> => {
      const r = await client.callTool({ name, arguments: args }) as { isError?: boolean; structuredContent?: Structured };
      return { isError: r.isError === true, data: r.structuredContent ?? {} };
    };
    const ok = async (name: string, args: Structured): Promise<Structured> => {
      const r = await call(name, args);
      expect(r.isError, `${name}: ${JSON.stringify(r.data)}`).toBe(false);
      return r.data;
    };

    // The agent opens the file; the person joins through the link it gives.
    const opened = await ok('open', { path: 'lease.docx' });
    const document = opened['document'] as string;
    const params = new URLSearchParams((opened['link'] as string).split('#')[1]);
    const pageDoc = new Y.Doc();
    const socket = new HocuspocusProviderWebsocket({ url: params.get('relay')!, WebSocketPolyfill: WebSocket, maxAttempts: 1 });
    let synced!: () => void;
    const ready = new Promise<void>((r) => { synced = r; });
    const provider = new HocuspocusProvider({ websocketProvider: socket, name: params.get('room')!, document: pageDoc, token: params.get('token')!, onSynced: () => synced() });
    provider.attach();
    cleanups.push(() => { provider.destroy(); socket.destroy(); socket.connect = async () => undefined; });
    await ready;
    const store = newEditStore(pageDoc);
    const person = await EditorDocument.open(store.room.get('base') as Uint8Array, { store, mode: 'join' });
    person.trackingAuthor = 'Ada';

    // It reads the outline and the paragraph it will change.
    const blocks = (await ok('outline', { document }))['blocks'] as { address: string; text: string }[];
    const target = blocks.find((b) => b.text === 'Some text.')!;
    expect((await ok('read', { document, address: target.address }))['markdown']).toBe(`<!-- ${target.address} -->\nSome text.`);

    // It rewrites the paragraph, with if_text; the person sees the agent's tracked change.
    const edited = await ok('edit', { document, address: target.address, markdown: 'Some **revised** text, as the agent proposes.', if_text: 'Some text.' });
    const proposal = (edited['created'] as { address: string; text: string }[])[0]!;
    expect(proposal.text).toBe('Some revised text, as the agent proposes.');
    await until(() => person.doc.textContent.includes('as the agent proposes'), 10_000, 'the agent\'s paragraph at the person');

    // Its own edit added a paragraph, so the ordinals after it moved: it reads the outline again (this file has no paragraph ids).
    const indent = ((await ok('outline', { document }))['blocks'] as { address: string; text: string }[]).find((b) => b.text === 'Left indent')!;
    // The person retitles "Left indent"; the agent's edit on what it read is refused, and it decides again.
    const at = (text: string): number => { let found = -1; person.doc.descendants((n, pos) => { if (found < 0 && n.type.name === 'paragraph' && n.textContent === text) found = pos + 1; return found < 0; }); return found; };
    runCommand(person, 'insertText', { text: 'Wide ' }, { from: at('Left indent') });
    const outlineTexts = async (): Promise<string[]> => ((await ok('outline', { document }))['blocks'] as { text: string }[]).map((b) => b.text);
    let seen = await outlineTexts();
    for (let i = 0; i < 200 && !seen.includes('Wide Left indent'); i++) { await new Promise((r) => setTimeout(r, 50)); seen = await outlineTexts(); }
    expect(seen).toContain('Wide Left indent');
    const stale = await call('replace_text', { document, find: 'Left', replace: 'Start', scope: indent.address, if_text: 'Left indent' });
    expect(stale.isError).toBe(true);
    expect(stale.data['code']).toBe('target.stale');
    const current = (stale.data['current'] as { text: string }).text;
    expect(current).toBe('Wide Left indent');
    const replaced = await ok('replace_text', { document, find: 'Left', replace: 'Start', scope: indent.address, if_text: current });
    expect((replaced['touched'] as { text: string }[])[0]!.text).toBe('Wide Start indent');

    // It comments on its own words; the person's session has the thread.
    const commented = await ok('comment', { document, address: proposal.address, anchor: 'revised', text: 'Tightened, as you asked.' });
    const commentId = (commented['comment'] as { id: string }).id;
    // Its words are its own pending insertion: the comment stays on them, in Word's form (core-ts 0.3.1), with no warning.
    expect(commented['warnings']).toEqual([]);
    // A peer's comment is adopted when the comments are read (`comments()`), as the margin reads them on each change.
    let threads = await person.stories.comments();
    for (let i = 0; i < 200 && !threads.some((t) => t.id === commentId); i++) { await new Promise((r) => setTimeout(r, 50)); threads = await person.stories.comments(); }
    expect(threads.find((t) => t.id === commentId)).toMatchObject({ author: 'Claude', done: false });

    // The person accepts the agent's rewrite of "Left indent" and leaves the rest pending.
    await until(() => trackedChangesOf(person.doc).some((c) => c.author === 'Claude' && c.text === 'Start'), 10_000, 'the replacement at the person');
    const replacement = trackedChangesOf(person.doc).find((c) => c.author === 'Claude' && c.text === 'Start')!;
    expect(runCommand(person, 'trackedChange.accept', { key: person.changeKey(replacement) }).ok).toBe(true);
    const leftOver = async (): Promise<{ text: string }[]> => ((await ok('get_tracked_changes', { document, author: 'Claude' }))['changes'] as { text: string }[]);
    let pending = await leftOver();
    for (let i = 0; i < 100 && pending.some((c) => c.text === 'Start'); i++) { await new Promise((r) => setTimeout(r, 50)); pending = await leftOver(); }
    expect(pending.some((c) => c.text === 'Start')).toBe(false);
    expect(pending.some((c) => c.text.includes('as the agent proposes'))).toBe(true);

    // It saves a copy; the file holds what the room holds: its pending proposal, the accepted word plain, the person's typing.
    const saved = await ok('save', { document, output_path: 'lease-reviewed.docx' });
    expect(saved['path']).toBe(join(await realRoot(root), 'lease-reviewed.docx'));
    const bytes = new Uint8Array(await readFile(saved['path'] as string));
    const main = (await WordprocessingMLPackage.load(bytes)).getMainDocumentPart();
    await main.getContents();
    const xml = await main.getXml();
    expect(xml).toMatch(/<w:ins w:id="\d+"[^>]* w:author="Claude"[^>]*>(?:(?!<\/w:ins>).)*as the agent proposes/);
    expect(xml).toMatch(/<w:commentRangeStart w:id="\d+"\/>/);
    const reopened = await EditorDocument.open(bytes);
    const texts = await reopened.readTree(({ body }) => body.paragraphs.map((p) => p.text));
    expect(texts).toContain('Some revised text, as the agent proposes.');
    expect(texts).toContain('Wide Start indent');
    // The accepted word is plain; the agent's deletion of the old one is still pending, as the person left it.
    expect(trackedChangesOf(reopened.doc).filter((c) => c.author === 'Claude').map((c) => [c.type, c.text])).toContainEqual(['Deleted', 'Left']);
    expect(trackedChangesOf(reopened.doc).some((c) => c.author === 'Claude' && c.text === 'Start')).toBe(false);

    // stdout carried the protocol alone; stderr the companion's log.
    expect(protocolErrors).toEqual([]);
    expect(log).toMatch(/docx4j-editor-mcp: the editor and the probe at http:\/\/127\.0\.0\.1:\d+, the relay at ws:\/\/127\.0\.0\.1:\d+/);
    expect(log).toContain('lease.docx is open in room');
  }, 120_000);
});

async function realRoot(root: string): Promise<string> {
  const { realpath } = await import('node:fs/promises');
  return realpath(root);
}
