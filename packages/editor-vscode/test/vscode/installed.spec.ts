// SPDX-License-Identifier: GPL-3.0-or-later
//
// The extension as a person installs it (ED-014 section 4.7, P51): the VSIX `scripts/package.mjs` makes from the build, and
// the suite's driver beside it, installed through VS Code's command line into a window with no development path. What only
// an installed extension does:
// - a document opens, takes an edit and saves, the saved parts compared with the source's;
// - Invite starts the companion from the VSIX's own `companion/`, serving its own `app/`, and the agent's `list_documents`
//   over `/mcp` names the document;
// - the words, through a stand-in French language pack (section 9.25 item 3): the palette's French title, a refused Save As
//   in French, and the page's boot naming the key's command as VS Code's palette does (P49).
// Hot exit is not here: an installed window restores a custom editor after a crash as a development host does, from VS
// Code's editor state alone (section 9.25 item 1), which `hotexit.spec.ts` holds.
import { test, expect } from '@playwright/test';
import { execFileSync } from 'node:child_process';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { EXTENSION, FIXTURES, VsCode, caretAtEndOf, sleep } from './harness.js';
import { vsixes } from './vsix.js';
import { mainPart, partsThatDiffer } from './zip.js';

test.describe.configure({ mode: 'serial' });

const version = (JSON.parse(readFileSync(join(EXTENSION, 'package.json'), 'utf8')) as { version: string }).version;
const INSTALLED = `plutext.docx4j-editor-${version}`;

/** The companions this spec's profiles started: their command lines. */
function companions(): string[] {
  const out = execFileSync('ps', ['-eo', 'pid=,args='], { encoding: 'utf8' });
  return out.split('\n').filter((l) => l.includes('/companion/cli.mjs') && l.includes('test-results/vscode/installed/'));
}

test.describe('installed', () => {
  let vs: VsCode;
  /** Jason's clipboard, put back after the spec reads the address through it. */
  let kept: string | undefined;

  test.beforeAll(async () => {
    const { extension, driver } = vsixes();
    vs = await VsCode.launch({
      profile: 'installed',
      installed: [extension, driver],
      files: { 'edit.docx': 'HelloWordOnline.docx', 'agent.docx': 'HelloWordOnline.docx' },
      // VS Code's MCP service runs only with its AI features on.
      settings: { 'chat.disableAIFeatures': false },
    });
    kept = await vs.ctl<string>({ op: 'clipboardRead' }).catch(() => undefined);
  });

  test.afterAll(async () => {
    if (vs && kept !== undefined) await vs.ctl({ op: 'clipboardWrite', text: kept }).catch(() => undefined);
    await vs?.close();
    for (const line of companions()) { try { process.kill(Number(line.trim().split(/\s+/)[0]), 'SIGKILL'); } catch { /* gone */ } }
  });

  test.afterEach(() => {
    expect(vs.problems, 'content security policy refusals').toEqual([]);
  });

  test('the window has no development path: the extension and the driver are installed ones', async () => {
    expect(readdirSync(join(vs.folder, '..', 'exts'))).toEqual(expect.arrayContaining([INSTALLED, 'plutext-test.docx4j-editor-test-driver-0.0.0']));
    const manifest = await vs.ctl<{ version: string } | null>({ op: 'manifest', id: 'plutext.docx4j-editor' });
    expect(manifest?.version).toBe(version);
  });

  test('a document opens, takes an edit and saves: the main part and docProps/app.xml written, every other part the source\'s', async () => {
    const frame = await vs.open('edit.docx');
    await caretAtEndOf(frame, 0);
    await vs.page.keyboard.type(' Typed in the installed editor', { delay: 20 });
    await expect.poll(async () => (await vs.tab('edit.docx'))?.isDirty).toBe(true);
    await vs.page.keyboard.press('Control+s');
    await expect.poll(async () => (await vs.tab('edit.docx'))?.isDirty).toBe(false);
    expect(mainPart(vs.path('edit.docx'))).toContain('Typed in the installed editor');
    expect(partsThatDiffer(join(FIXTURES, 'HelloWordOnline.docx'), vs.path('edit.docx'))).toEqual(['docProps/app.xml', 'word/document22.xml']);
    await vs.closeAll();
  });

  test('Invite: the companion from the VSIX\'s companion/, serving its app/; list_documents over /mcp names the document', async () => {
    const frame = await vs.open('agent.docx');
    await frame.getByTestId('invite-agent').click();
    await expect(frame.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 30_000 });
    await expect(frame.getByTestId('status')).toContainText('Agent joined', { timeout: 30_000 });
    const lines = companions();
    expect(lines).toHaveLength(1);
    const exts = join(vs.folder, '..', 'exts', INSTALLED);
    expect(lines[0]).toContain(`${join(exts, 'companion', 'cli.mjs')} `);
    expect(lines[0]).toContain(`--app ${join(exts, 'app')} `);
    // The address and its token, as "Copy the Agent's MCP Address" and its Claude Code command give them.
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
    const r = await fetch(address, {
      method: 'POST',
      headers: { 'content-type': 'application/json', accept: 'application/json, text/event-stream', authorization: `Bearer ${token}` },
      body: JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name: 'list_documents', arguments: {} } }),
    });
    expect(r.status).toBe(200);
    const { documents } = ((await r.json()) as { result: { structuredContent: { documents: { document: string; people: { agent: boolean }[] }[] } } }).result.structuredContent;
    expect(documents).toHaveLength(1);
    expect(documents[0]!.people.some((p) => p.agent)).toBe(true);
    // The page the companion serves is the VSIX's own application.
    const served = await (await fetch(address.replace(/\/mcp$/, '/'))).text();
    expect(served).toContain('<meta name="docx4j-companion"');
    expect(served.replace(/<meta name="docx4j-companion"[^>]*>/, '')).toBe(readFileSync(join(exts, 'app', 'index.html'), 'utf8'));
  });
});

test.describe('the words, through a language pack', () => {
  const fr = (file: string): Record<string, string> => JSON.parse(readFileSync(join(EXTENSION, file), 'utf8')) as Record<string, string>;

  test('VS Code in French: the palette\'s title, a refused Save As, and the boot\'s name of the key\'s command, in French (P48, P49)', async () => {
    const { extension, driver, pack } = vsixes();
    const launch = { profile: 'installed-fr', installed: [extension, driver, pack], locale: 'fr', files: { 'source.docx': 'HelloWordOnline.docx' } };
    // VS Code takes a pack installed through its command line from its second start (section 9.25 item 2).
    let vs = await VsCode.launch(launch);
    await vs.close();
    vs = await VsCode.launch({ ...launch, keep: true });
    try {
      expect((await vs.ctl<{ language: string }>({ op: 'ping' })).language).toBe('fr');
      const manifest = fr('package.nls.fr.json');
      const messages = fr('l10n/bundle.l10n.fr.json');
      // The palette.
      void vs.exec('workbench.action.showCommands').catch(() => undefined);
      const input = vs.page.locator('.quick-input-widget input:not([type="checkbox"])');
      await input.waitFor({ timeout: 15_000 });
      await input.fill('>docx4j');
      await expect.poll(() => vs.page.evaluate(() => [...document.querySelectorAll('.quick-input-list .monaco-list-row')].map((e) => (e as HTMLElement).innerText.replace(/\s+/g, ' ')).join('\n')))
        .toContain(`docx4j editor: ${manifest['command.ask.removeKey']}`);
      await vs.page.keyboard.press('Escape');
      // The page names the key's command as the palette does.
      const frame = await vs.open('source.docx');
      const boot = JSON.parse(await frame.evaluate(() => document.querySelector('meta[name="docx4j-host"]')!.getAttribute('content')!)) as { language: string; commands: Record<string, string> };
      expect(boot.language).toBe('fr');
      expect(boot.commands).toEqual({ 'docx4j.ask.setKey': `docx4j editor: ${manifest['command.ask.setKey']}` });
      // A message the extension shows.
      await caretAtEndOf(frame, 0);
      await vs.saveAs(vs.path('notes.md'));
      const refused = messages["{0} is not a Word document's name: save it as .docx, .dotx, .docm or .dotm, or as .xml for Flat OPC."]!.replace('{0}', 'notes.md');
      await expect.poll(() => vs.notifications()).toContainEqual(expect.stringContaining(refused));
      await vs.closeAll();
    } finally {
      await vs.close();
    }
  });
});
