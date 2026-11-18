// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 8.9, step F1, in a browser: the companion (`@docx4j/editor-mcp`, started in the
// test on a free port) serves the built editor and its relay; a page opens the room link it printed
// for a file, joins, and has the document; the agent is in the People list; and what the agent
// writes arrives as its tracked change.
import { test, expect } from '@playwright/test';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createCompanion } from '@docx4j/editor-mcp';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { fixturePath, download, comparePackages } from './helpers.js';

const here = dirname(fileURLToPath(import.meta.url));
const body = '.docx-editor .ProseMirror[aria-label="Document"]';

test("a page opens the companion's room link: the document, the agent in People, and the agent's edit as its tracked change", async ({ page }) => {
  const root = await mkdtemp(join(tmpdir(), 'companion-root-'));
  const data = await mkdtemp(join(tmpdir(), 'companion-rooms-'));
  await writeFile(join(root, 'lease.docx'), await readFile(join(here, '..', '..', 'editor-model', 'test', 'fixtures', 'tracked-changes.docx')));
  const companion = createCompanion({ roots: [root], port: 0, app: join(here, '..', 'dist'), data, log: () => undefined });
  try {
    await companion.listen();
    const room = await companion.open('lease.docx');
    await page.goto(room.link);
    await page.getByTestId('join-accept').click();
    await page.locator(body).waitFor();
    await expect(page.locator(body)).toContainText('Docx sample document');
    // The agent drawn as its initials in its colour, its name on the button.
    await expect(page.getByTestId('person')).toHaveAttribute('aria-label', /Claude/, { timeout: 10_000 });
    await expect(page.getByTestId('person')).toHaveCSS('background-color', 'rgb(109, 40, 217)');
    const run = await room.session.runOnTree(({ body: b }) => { b.insertParagraph('The agent was here.', 'End'); }, 'agent', { author: 'Claude', mode: 'TrackAll' });
    expect(run.error).toBeUndefined();
    await expect(page.locator(`${body} ins.insertion[data-author="Claude"]`)).toContainText('The agent was here.', { timeout: 10_000 });
  } finally {
    await companion.close();
    await rm(root, { recursive: true, force: true });
    await rm(data, { recursive: true, force: true });
  }
});

// ED-005 section 8.9, step F4: the browser first. A page the companion serves opens a file, invites the agent (its
// companion found at the page's own /companion), sees it badged in People with a Stop, sees its call as its tracked
// change and in the review panel's log, and stops it: it leaves, and its calls there are refused.
test('a page invites the agent, sees it in People, sees its call and its log, and stops it', async ({ page }) => {
  const root = await mkdtemp(join(tmpdir(), 'companion-root-'));
  const data = await mkdtemp(join(tmpdir(), 'companion-rooms-'));
  const companion = createCompanion({ roots: [root], port: 0, app: join(here, '..', 'dist'), data, log: () => undefined });
  const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: 'invite-e2e', version: '1' });
  try {
    const base = await companion.listen();
    await companion.mcp().connect(serverSide);
    await client.connect(clientSide);
    await page.goto(base);
    await page.getByTestId('open').setInputFiles(fixturePath('tracked-changes.docx'));
    await page.locator(body).waitFor();
    await page.getByTestId('invite-agent').click();
    await expect(page.getByTestId('status')).toContainText('Claude joined');
    await expect(page.locator('[data-testid="person"][data-agent="true"]')).toHaveAttribute('aria-label', /Claude/);
    await expect(page.getByTestId('agent-badge')).toHaveText('Agent');
    const [document] = [...companion.rooms.keys()];
    expect(document).toBeDefined();
    // The agent's call arrives as its tracked change, and in the log.
    const result = await client.callTool({ name: 'insert', arguments: { document, text: 'Written by the invited agent', location: 'End' } }) as { isError?: boolean };
    expect(result.isError).not.toBe(true);
    await expect(page.locator(`${body} ins.insertion[data-author="Claude"]`)).toContainText('Written by the invited agent');
    await page.getByTestId('perspective-review').click();
    await expect(page.getByTestId('agent-log-entry').first()).toContainText('Inserted a paragraph at the end of the body');
    // A click on the entry goes to the paragraph the call made.
    await page.getByTestId('agent-log-entry').first().getByRole('button').click();
    await expect.poll(() => page.evaluate(() => {
      const v = (window as unknown as { docx4jEditor: { state: { selection: { $from: { parent: { textContent: string } } } } } }).docx4jEditor;
      return v.state.selection.$from.parent.textContent;
    })).toBe('Written by the invited agent');
    // Stop: the agent leaves People, the companion hears it, and its calls there are refused.
    await page.getByTestId('agent-stop').click();
    await expect(page.getByTestId('status')).toContainText('Claude was stopped');
    await expect(page.locator('[data-testid="person"][data-agent="true"]')).toHaveCount(0, { timeout: 10_000 });
    await expect.poll(() => companion.rooms.get(document!)?.stopped).toBe(true);
    const refused = await client.callTool({ name: 'outline', arguments: { document } }) as { structuredContent?: { code?: string } };
    expect(refused.structuredContent?.code).toBe('agent.stopped');
  } finally {
    await client.close();
    await companion.close();
    await rm(root, { recursive: true, force: true });
    await rm(data, { recursive: true, force: true });
  }
});

// Jason's by-hand run (ED-005 section 12.29 item 6): Stop and Invite worked, but nothing on the page said so.
// At an ordinary width the status line, sharing the app bar with its buttons, was squeezed to nothing; and the
// companion's log of every call, reads included, and an invitation clearing a Stop marked the document unsaved.
test('the agent stopped and invited back, as a person sees it: the status bar says so, the returning agent is marked, and neither its reads nor the invitation make the document unsaved', async ({ page }) => {
  await page.setViewportSize({ width: 1440, height: 900 });
  const root = await mkdtemp(join(tmpdir(), 'companion-root-'));
  const data = await mkdtemp(join(tmpdir(), 'companion-rooms-'));
  await writeFile(join(root, 'lease.docx'), await readFile(join(here, '..', '..', 'editor-model', 'test', 'fixtures', 'tracked-changes.docx')));
  const companion = createCompanion({ roots: [root], port: 0, app: join(here, '..', 'dist'), data, log: () => undefined });
  const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: 'status-e2e', version: '1' });
  try {
    await companion.listen();
    await companion.mcp().connect(serverSide);
    await client.connect(clientSide);
    const opened = (await client.callTool({ name: 'open', arguments: { path: 'lease.docx' } }) as unknown as { structuredContent: { document: string; link: string } }).structuredContent;
    await page.goto(opened.link);
    await page.getByTestId('join-accept').click();
    await expect(page.locator(body)).toContainText('Docx sample document');
    await expect(page.getByTestId('agent-badge')).toBeVisible();
    // The agent reads: its calls are logged in the room (the Review perspective shows the log), which is not an edit.
    await page.getByTestId('perspective-review').click();
    await client.callTool({ name: 'outline', arguments: { document: opened.document } });
    await expect(page.getByTestId('agent-log-entry').first()).toBeAttached();
    await expect(page.getByTestId('name')).toHaveText('lease.docx');
    // Stop: the status bar says so, where it can be seen.
    const status = page.getByTestId('status');
    await page.getByTestId('agent-stop').click();
    await expect(status).toHaveText('Claude was stopped and has left the document');
    await expect(status).toBeVisible();
    expect((await status.boundingBox())!.width).toBeGreaterThan(200);
    // Invite: the status bar says the agent joined, its chip rings as it arrives, and the document is still saved.
    await page.getByTestId('invite-agent').click();
    await expect(status).toHaveText('Claude joined: its edits come as its tracked changes');
    await expect(status).toBeVisible();
    const chip = page.locator('[data-testid="person"][data-agent="true"]');
    await expect(chip).toBeVisible();
    await expect(chip).toHaveCSS('animation-name', 'agent-arrives');
    await expect(page.getByTestId('name')).toHaveText('lease.docx');
  } finally {
    await client.close();
    await companion.close();
    await rm(root, { recursive: true, force: true });
    await rm(data, { recursive: true, force: true });
  }
});

// ED-005 section 9.3 item 1 (i), step G2: the relay in a worker of its own. A tool call's tree run holds the companion's
// event loop for its whole length (section 12.31 item 3); with the relay on that loop, two people's typing to each other
// waited for it. Here the companion's thread (this test's) is held for 1.5 s, as a long call holds it, while page A types
// on a timer of its own; page B must have the words before the hold ends. With the relay on the companion's loop, B would
// receive them only after it.
test('two pages type to each other while the companion\'s thread is busy: the relay forwards on its own event loop', async ({ browser }) => {
  const root = await mkdtemp(join(tmpdir(), 'companion-root-'));
  const data = await mkdtemp(join(tmpdir(), 'companion-rooms-'));
  await writeFile(join(root, 'lease.docx'), await readFile(join(here, '..', '..', 'editor-model', 'test', 'fixtures', 'tracked-changes.docx')));
  const companion = createCompanion({ roots: [root], port: 0, app: join(here, '..', 'dist'), data, log: () => undefined });
  const contexts = [await browser.newContext(), await browser.newContext()];
  try {
    await companion.listen();
    const room = await companion.open('lease.docx');
    const [a, b] = await Promise.all(contexts.map((c) => c.newPage()));
    for (const page of [a!, b!]) {
      await page.goto(room.link);
      await page.getByTestId('join-accept').click();
      await expect(page.locator(body)).toContainText('Docx sample document');
    }
    // B notes when the words arrive, by its own clock.
    await b!.evaluate(() => {
      const w = window as unknown as { docx4jEditor: { state: { doc: { textContent: string } } }; seenAt?: number };
      const timer = setInterval(() => { if (w.docx4jEditor.state.doc.textContent.includes('typed while the companion is busy')) { w.seenAt = Date.now(); clearInterval(timer); } }, 5);
    });
    // A types in 150 ms, on its own timer: this test's thread will be held by then.
    await a!.evaluate(() => {
      const v = (window as unknown as { docx4jEditor: { view: { state: { tr: { insertText(t: string, pos: number): unknown } }; dispatch(tr: unknown): void } } }).docx4jEditor.view;
      setTimeout(() => { v.dispatch(v.state.tr.insertText('typed while the companion is busy ', 1)); }, 150);
    });
    const holdEnd = Date.now() + 1500;
    while (Date.now() < holdEnd) { /* the companion's event loop held, as a long tool call holds it */ }
    const seenAt = await b!.evaluate(() => (window as unknown as { seenAt?: number }).seenAt);
    expect(seenAt, 'B had the words').toBeDefined();
    expect(seenAt!, 'B had them before the hold ended').toBeLessThan(holdEnd - 500);
    // And the agent's session has them once its thread is free.
    await expect.poll(() => room.session.doc.textContent.includes('typed while the companion is busy'), { timeout: 10_000 }).toBe(true);
  } finally {
    for (const c of contexts) await c.close();
    await companion.close();
    await rm(root, { recursive: true, force: true });
    await rm(data, { recursive: true, force: true });
  }
});


// ED-005 section 9.3, step G2: the agent writes a paragraph while page A types in it and page B watches. Where the typing
// lands against the call is timing: before it (the call then replaces the typed text), inside it (the call refused with
// target.stale and redone, proposal 42), or in flight (the stated limitation: the typing takes the agent's deletion). In
// every case the three agree: the same text, the same changes, and the two pages' downloads the same bytes.
test('the agent edits a paragraph while page A types in it and page B watches: the three agree, and the two downloads are the same bytes', async ({ browser }) => {
  const root = await mkdtemp(join(tmpdir(), 'companion-root-'));
  const data = await mkdtemp(join(tmpdir(), 'companion-rooms-'));
  await writeFile(join(root, 'lease.docx'), await readFile(join(here, '..', '..', 'editor-model', 'test', 'fixtures', 'tracked-changes.docx')));
  const companion = createCompanion({ roots: [root], port: 0, app: join(here, '..', 'dist'), data, log: () => undefined });
  const contexts = [await browser.newContext(), await browser.newContext()];
  try {
    await companion.listen();
    const room = await companion.open('lease.docx');
    const [a, b] = await Promise.all(contexts.map((c) => c.newPage()));
    for (const page of [a!, b!]) {
      await page.goto(room.link);
      await page.getByTestId('join-accept').click();
      await expect(page.locator(body)).toContainText('Docx sample document');
    }
    const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
    await companion.mcp().connect(serverSide);
    const client = new Client({ name: 'three-agree', version: '1' });
    await client.connect(clientSide);
    // A types into "Some text." on its own timer, a few keystrokes apart, while the agent's call runs.
    await a!.evaluate(() => {
      const v = (window as unknown as { docx4jEditor: { view: { state: { doc: { forEach(f: (n: { textContent: string }, pos: number) => void): void }; tr: { insertText(t: string, pos: number): unknown } }; dispatch(tr: unknown): void } } }).docx4jEditor.view;
      let i = 0;
      const tick = setInterval(() => {
        let at = -1;
        v.state.doc.forEach((n, pos) => { if (at < 0 && n.textContent.includes('text.')) at = pos + 1 + n.textContent.indexOf('text.'); });
        if (at > 0) v.dispatch(v.state.tr.insertText('ab'[i % 2]!, at));
        if (++i >= 6) clearInterval(tick);
      }, 15);
    });
    let result: { isError?: boolean; structuredContent?: { code?: string } } = {};
    for (let attempt = 0; attempt < 5; attempt++) {
      result = await client.callTool({ name: 'edit', arguments: { document: room.id, address: { contains: 'text.' }, text: 'The agent\'s paragraph.' } }) as typeof result;
      if (!result.isError) break;
      expect(result.structuredContent?.code).toBe('target.stale');
    }
    expect(result.isError).toBeFalsy();
    // The three agree: A, B and the agent's session hold the same text and the same changes.
    const shape = (page: typeof a) => page!.evaluate(() => {
      const v = (window as unknown as { docx4jEditor: { state: { doc: { textContent: string } } } }).docx4jEditor;
      return v.state.doc.textContent;
    });
    // All three read each time: the agent's session receives the last keystrokes over the relay like the pages.
    await expect.poll(async () => { const [x, y] = [await shape(a), await shape(b)]; return x === y && y === room.session.doc.textContent; }, { timeout: 10_000 }).toBe(true);
    expect(room.session.doc.textContent).toContain('The agent\'s paragraph.');
    const [da, db] = [await download(a!), await download(b!)];
    expect(comparePackages(da, db, ['docProps/app.xml'])).toEqual([]);
    await client.close();
  } finally {
    for (const c of contexts) await c.close();
    await companion.close();
    await rm(root, { recursive: true, force: true });
    await rm(data, { recursive: true, force: true });
  }
});
