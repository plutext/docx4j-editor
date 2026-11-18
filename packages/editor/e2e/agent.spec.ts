// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 8.9, step F5: a page and the companion together, the agent first. An MCP client
// (the SDK's in-memory transport, as a script drives the agent) opens a file; a person opens the link
// it gives. The agent rewrites a paragraph in Markdown and comments on it; the person sees both as
// the agent's, accepts its insertion and rejects its deletion in the Review perspective; the person
// types, the agent's edit on what it read is refused and it decides again; the log shows its calls;
// the person's download holds the room as it is; a Stop holds against reopening the file, and Invite
// brings the agent back.
import { test, expect } from '@playwright/test';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { ZipPartStore } from '@docx4j/core-ts';
import { EditorDocument, trackedChangesOf } from '@docx4j/editor-model';
import { createCompanion } from '@docx4j/editor-mcp';
import { download } from './helpers.js';

const here = dirname(fileURLToPath(import.meta.url));
const body = '.docx-editor .ProseMirror[aria-label="Document"]';
type Structured = Record<string, unknown>;

test('the agent first: its Markdown edit and comment reach the person, who accepts and rejects them; a stale edit is refused; the log shows its calls; the download holds the room; a Stop holds against reopening the file, and Invite brings it back', async ({ page }) => {
  const root = await mkdtemp(join(tmpdir(), 'agent-root-'));
  const data = await mkdtemp(join(tmpdir(), 'agent-rooms-'));
  await writeFile(join(root, 'lease.docx'), await readFile(join(here, '..', '..', 'editor-model', 'test', 'fixtures', 'tracked-changes.docx')));
  await writeFile(join(root, 'notes.docx'), await readFile(join(here, '..', '..', 'editor-model', 'test', 'fixtures', 'footnote.docx')));
  const companion = createCompanion({ roots: [root], port: 0, app: join(here, '..', 'dist'), data, log: () => undefined });
  const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: 'agent-e2e', version: '1' });
  const call = async (name: string, args: Structured): Promise<{ isError: boolean; data: Structured }> => {
    const r = await client.callTool({ name, arguments: args }) as { isError?: boolean; structuredContent?: Structured };
    return { isError: r.isError === true, data: r.structuredContent ?? {} };
  };
  const ok = async (name: string, args: Structured): Promise<Structured> => {
    const r = await call(name, args);
    expect(r.isError, `${name}: ${JSON.stringify(r.data)}`).toBe(false);
    return r.data;
  };
  try {
    await companion.listen();
    await companion.mcp().connect(serverSide);
    await client.connect(clientSide);

    // The agent opens the file; the person opens its link.
    const opened = await ok('open', { path: 'lease.docx' });
    const document = opened['document'] as string;
    await page.goto(opened['link'] as string);
    await page.getByTestId('join-accept').click();
    await expect(page.locator(body)).toContainText('Docx sample document');

    // It rewrites a paragraph in Markdown, a heading and a list among it, and comments on a word.
    const blocks = (await ok('outline', { document }))['blocks'] as { address: string; text: string }[];
    const target = blocks.find((b) => b.text === 'Some text.')!;
    const edited = await ok('edit', { document, address: target.address, if_text: 'Some text.', markdown: '## Terms from the agent\n\nThe tenant pays **monthly**, as [agreed](https://example.com/terms).\n\n- first point\n- second point' });
    const paragraph = (edited['created'] as { address: string; text: string }[])[1]!;
    await ok('comment', { document, address: paragraph.address, anchor: 'monthly', text: 'Was quarterly: please confirm.' });
    await expect(page.locator(`${body} ins.insertion[data-author="Claude"]`).first()).toContainText('Terms from the agent');
    await expect(page.locator(`${body} del.deletion[data-author="Claude"]`)).toContainText('Some text.');
    await expect(page.locator('[data-testid="comment-bubble"]', { hasText: 'Claude' })).toContainText('Was quarterly: please confirm.');
    const headingStyle = await page.evaluate(() => {
      const view = (window as unknown as { docx4jEditor: { state: { doc: { descendants(f: (n: { type: { name: string }; attrs: Record<string, unknown>; textContent: string }) => boolean | void): void } } } }).docx4jEditor;
      let style: unknown = null;
      view.state.doc.descendants((n) => { if (n.type.name === 'paragraph' && n.textContent === 'Terms from the agent') style = n.attrs['pStyle']; return style === null; });
      return style;
    });
    expect(headingStyle).toBe('Heading2');

    // In the Review perspective: the person accepts the agent's insertion and rejects its deletion.
    await page.getByTestId('perspective-review').click();
    const insertion = page.locator('[data-testid="review-entry"][data-author="Claude"][data-type="Added"]');
    const deletion = page.locator('[data-testid="review-entry"][data-author="Claude"][data-type="Deleted"]');
    // The paragraph the edit replaced and the first of what it wrote are one entry, a replacement
    // (proposal 41, ED-005 section 9.7). Two more insertions: the comment's markers go outside the
    // w:ins its words sit in (the engine's rule, CR-002 section 13), so its reference run parts the
    // insertion before the link's (ED-005 section 12.29); and the link's own text is a change of its
    // own, as Office JS lists an insertion that runs into a link and out again (core-ts check 30).
    const replaced = page.locator('[data-testid="review-entry"][data-author="Claude"][data-type="Replaced"]');
    await expect(replaced).toHaveCount(1);
    await expect(insertion).toHaveCount(2);
    // To accept what it wrote and reject what it deleted, the person splits the replacement.
    await replaced.getByTestId('review-split').click();
    await expect(insertion).toHaveCount(3);
    await expect(deletion).toHaveAttribute('data-split', 'true');
    while (await insertion.count() > 0) await insertion.first().getByTestId('review-accept').click();
    await deletion.getByTestId('review-reject').click();
    await expect(page.locator('[data-testid="review-entry"][data-author="Claude"]')).toHaveCount(0);
    await expect(page.locator('[data-testid="review-entry"][data-author="Claude"]')).toHaveCount(0);
    await expect.poll(async () => ((await ok('get_tracked_changes', { document, author: 'Claude' }))['changes'] as unknown[]).length).toBe(0);

    // The person types into "Left indent"; the agent's edit on what it read is refused, and it decides again.
    const indent = ((await ok('outline', { document }))['blocks'] as { address: string; text: string }[]).find((b) => b.text === 'Left indent')!;
    await page.evaluate(() => {
      const view = (window as unknown as { docx4jEditor: { state: { doc: { descendants(f: (n: { type: { name: string }; textContent: string }, pos: number) => boolean | void): void }; tr: { insertText(t: string, pos: number): unknown } }; view: { dispatch(tr: unknown): void } } }).docx4jEditor;
      let at = -1;
      view.state.doc.descendants((n, pos) => { if (at < 0 && n.type.name === 'paragraph' && n.textContent === 'Left indent') at = pos + 1; return at < 0; });
      view.view.dispatch(view.state.tr.insertText('Wide ', at));
    });
    // Once the typing has reached the agent's copy of the room, its edit on what it read before is refused.
    await expect.poll(async () => ((await ok('outline', { document }))['blocks'] as { text: string }[]).some((b) => b.text === 'Wide Left indent')).toBe(true);
    const stale = await call('replace_text', { document, find: 'Left', replace: 'Start', scope: indent.address, if_text: 'Left indent' });
    expect(stale.data['code']).toBe('target.stale');
    const current = (stale.data['current'] as { text: string }).text;
    expect(current).toBe('Wide Left indent');
    await ok('replace_text', { document, find: 'Left', replace: 'Start', scope: indent.address, if_text: current });
    await expect(page.locator(`${body} ins.insertion[data-author="Claude"]`, { hasText: 'Start' })).toBeVisible();

    // The log shows its calls, the refused one marked.
    const log = page.getByTestId('agent-log-entry');
    await expect(log.first()).toContainText('Replaced one match of "Left" with "Start" in one paragraph');
    await expect(page.locator('[data-testid="agent-log-entry"][data-ok="false"]')).toContainText('target.stale');
    await expect(page.locator('[data-testid="agent-log-entry"][data-tool="open"]')).toContainText('Joined the room');

    // The person's download holds the room as it is: the heading accepted and plain, "Some text." back, the replacement pending, the comment.
    const bytes = await download(page);
    const reopened = await EditorDocument.open(bytes);
    const texts = await reopened.readTree(({ body: b }) => b.paragraphs.map((p) => p.text));
    expect(texts).toEqual(expect.arrayContaining(['Terms from the agent', 'The tenant pays monthly, as agreed.', 'Some text.', 'Wide Start indent']));
    const pending = trackedChangesOf(reopened.doc).filter((c) => c.author === 'Claude').map((c) => [c.type, c.text]);
    expect(pending).toEqual(expect.arrayContaining([['Deleted', 'Left'], ['Added', 'Start']]));
    const comments = new TextDecoder().decode(new ZipPartStore(bytes).loadSync('word/comments.xml'));
    expect(comments).toContain('Was quarterly: please confirm.');

    // The agent's replace_text is one entry (proposal 41): its Accept resolves both halves, and one Ctrl+Z brings both back.
    const replacement = page.locator('[data-testid="review-entry"][data-author="Claude"][data-type="Replaced"]');
    await expect(replacement).toHaveCount(1);
    await expect(replacement.locator('.snippet')).toHaveText('“Left” with “Start”');
    await replacement.getByTestId('review-accept').click();
    await expect(page.locator('[data-testid="review-entry"][data-author="Claude"]')).toHaveCount(0);
    await expect.poll(async () => ((await ok('get_tracked_changes', { document, author: 'Claude' }))['changes'] as unknown[]).length).toBe(0);
    await page.locator(body).focus();
    await page.keyboard.press('Control+z');
    await expect(replacement).toHaveCount(1);
    // the engine's tracked Replace writes the insertion first (core-ts 0.6.0; ED-008 section 8.8 item 9)
    await expect.poll(async () => ((await ok('get_tracked_changes', { document, author: 'Claude' }))['changes'] as { kind: string }[]).map((c) => c.kind)).toEqual(['insertion', 'deletion']);

    // The person presses Stop: the agent's calls on the room are refused, and so is opening its file again
    // (Claude Code did that in the by-hand run, ED-005 section 12.29). Invite brings it back into the same room.
    // Named for what it does and for whom (ED-005 section 12.41); the log a list of buttons.
    await expect(page.getByRole('button', { name: /^Stop Claude/ })).toBeVisible();
    await expect(page.getByTestId('agent-log').getByRole('list').getByRole('listitem').first().getByRole('button')).toBeVisible();
    await page.getByTestId('agent-stop').click();
    await expect(page.locator('[data-testid="person"][data-agent="true"]')).toHaveCount(0);
    expect((await call('outline', { document })).data['code']).toBe('agent.stopped');
    const reopened2 = await call('open', { path: 'lease.docx' });
    expect(reopened2.data).toMatchObject({ code: 'agent.stopped', document });
    await page.getByTestId('invite-agent').click();
    await expect(page.locator('[data-testid="person"][data-agent="true"]')).toHaveCount(1);
    expect((await ok('open', { path: 'lease.docx' }))['document']).toBe(document);

    // Another document's link, pasted into this tab: only the hash changes and the page does not reload, but
    // it offers the join all the same (Jason's by-hand run, ED-005 section 12.29).
    const second = await ok('open', { path: 'notes.docx' });
    await page.evaluate((hash) => { location.hash = hash; }, new URL(second['link'] as string).hash);
    await page.getByTestId('join-accept').click();
    await expect(page.locator(body)).toContainText('Here is a document with 2 footnotes');
    await ok('insert', { document: second['document'] as string, text: 'The agent, in the second document', location: 'End' });
    await expect(page.locator(`${body} ins.insertion[data-author="Claude"]`, { hasText: 'The agent, in the second document' })).toBeVisible();
  } finally {
    await client.close();
    await companion.close();
    await rm(root, { recursive: true, force: true });
    await rm(data, { recursive: true, force: true });
  }
});
