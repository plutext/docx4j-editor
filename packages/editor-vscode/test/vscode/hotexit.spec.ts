// SPDX-License-Identifier: GPL-3.0-or-later
//
// Hot exit (ED-014 section 4.3, P8, P20) in a real VS Code: an edit across a kill of VS Code's main process; a Markdown
// conversion, not edited, across a kill; a rename in the explorer while dirty; a page loaded again after a save; and
// steady typing deferring the backup.
import { test, expect } from '@playwright/test';
import { readdirSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { VsCode, bodyText, caretAtEndOf, sleep } from './harness.js';
import { mainPart } from './zip.js';

test.describe.configure({ mode: 'serial' });

let vs: VsCode;

test.beforeAll(async () => {
  vs = await VsCode.launch({
    profile: 'hotexit',
    files: { 'kill.docx': 'HelloWordOnline.docx', 'saved.docx': 'HelloWordOnline.docx', 'rename.docx': 'HelloWordOnline.docx', 'reload.docx': 'HelloWordOnline.docx', 'steady.docx': 'HelloWordOnline.docx' },
  });
});

test.afterAll(async () => { await vs?.close(); });

test.afterEach(async () => {
  expect(vs.problems, 'content security policy refusals').toEqual([]);
});

async function tabIs(label: string, dirty: boolean): Promise<void> {
  await expect.poll(async () => (await vs.tab(label))?.isDirty, { message: `${label} ${dirty ? 'dirty' : 'clean'}` }).toBe(dirty);
}

/** The backup files the extension wrote, in its workspace storage. */
async function backups(): Promise<{ name: string; mtimeMs: number }[]> {
  const { workspace } = await vs.ctl<{ workspace: string | null }>({ op: 'storagePath' });
  // The driver's storage path is its own; the extension's is beside it, under its id.
  const folder = join(workspace!, '..', 'plutext.docx4j-editor');
  try { return readdirSync(folder).map((name) => ({ name, mtimeMs: statSync(join(folder, name)).mtimeMs })); } catch { return []; }
}

/**
 * VS Code gives an Extension Development Host window, as the suite runs the extension, no backup folder of its own
 * (`doOpenInBrowserWindow`: no `backupPath` with `extensionDevelopmentPath`), so after a crash it restores a custom editor
 * from its saved editor state alone, which names the backup and which it writes every 60 s (`DEFAULT_FLUSH_INTERVAL`).
 * A kill is made after that (ED-014 section 9.9). An installed extension's window has the folder: S7 checks that path
 * with the packaged VSIX.
 */
const STATE_WRITTEN_MS = 62_000;

/** Opens the restored editors, which VS Code shows as tabs before it resolves them. */
async function openTab(label: string): Promise<void> {
  const tab = await vs.tab(label);
  expect(tab, label).toBeTruthy();
  await vs.exec('vscode.openWith', { $parse: tab!.uri }, 'docx4j.editor');
}

test('an edit, a pause of more than a second, VS Code killed and started again: the editor dirty, with the edit', async () => {
  const frame = await vs.open('kill.docx');
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.type(' kept across a kill', { delay: 20 });
  await tabIs('kill.docx', true);
  await expect.poll(async () => (await backups()).length, { timeout: 15_000 }).toBeGreaterThan(0);
  await sleep(STATE_WRITTEN_MS);
  vs = await vs.restart('SIGKILL');
  await expect.poll(async () => (await vs.tab('kill.docx'))?.isDirty, { timeout: 60_000 }).toBe(true);
  await openTab('kill.docx');
  const again = await vs.frameFor('kill.docx');
  await expect.poll(() => bodyText(again)).toContain('kept across a kill');
  await expect(again.getByTestId('name')).toHaveText('kill.docx •');
  await vs.closeAll();
});

test('a save, then an edit, then a kill: the restored document shows no banner, and Save writes (Fable\'s review, section 9.11)', async () => {
  const frame = await vs.open('saved.docx');
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.type(' saved', { delay: 20 });
  await tabIs('saved.docx', true);
  await vs.page.keyboard.press('Control+s');
  await tabIs('saved.docx', false);
  await vs.page.keyboard.type(' then edited', { delay: 20 });
  await tabIs('saved.docx', true);
  await sleep(STATE_WRITTEN_MS);
  vs = await vs.restart('SIGKILL');
  await expect.poll(async () => (await vs.tab('saved.docx'))?.isDirty, { timeout: 60_000 }).toBe(true);
  await openTab('saved.docx');
  const again = await vs.frameFor('saved.docx');
  await expect.poll(() => bodyText(again)).toContain('saved then edited');
  // The session's own save is no change on disk: the backup recorded the bytes it wrote.
  await sleep(3000);
  await expect(again.getByTestId('disk-offer')).toHaveCount(0);
  await vs.page.keyboard.press('Control+s');
  await tabIs('saved.docx', false);
  expect(mainPart(vs.path('saved.docx'))).toContain('then edited');
  await vs.closeAll();
});

test('a .md opened and not edited, VS Code killed and started again: the conversion back (section 9.8 finding 3)', async () => {
  writeFileSync(vs.path('notes.md'), '# Notes\n\nA paragraph with **bold** text.\n');
  await vs.exec('vscode.openWith', { $uri: vs.path('notes.md') }, 'docx4j.markdown');
  const frame = await vs.frameFor('notes.docx');
  await expect.poll(() => bodyText(frame)).toContain('A paragraph with bold text.');
  await tabIs('notes.docx', true);
  await expect.poll(async () => (await backups()).length, { timeout: 15_000 }).toBeGreaterThan(0);
  await sleep(STATE_WRITTEN_MS);
  vs = await vs.restart('SIGKILL');
  await expect.poll(async () => (await vs.tab('notes.docx'))?.isDirty, { timeout: 60_000 }).toBe(true);
  await openTab('notes.docx');
  const again = await vs.frameFor('notes.docx');
  await expect.poll(() => bodyText(again)).toContain('A paragraph with bold text.');
  // An untitled document's close asks; Don't Save leaves nothing.
  await vs.closeAll();
});

test('a rename while dirty: the edit kept at the new name, which VS Code alone would lose (section 9.9)', async () => {
  const frame = await vs.open('rename.docx');
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.type(' kept across a rename', { delay: 20 });
  await tabIs('rename.docx', true);
  await sleep(2000);
  expect(await vs.ctl({ op: 'rename', from: vs.path('rename.docx'), to: vs.path('renamed.docx') })).toBe(true);
  const again = await vs.frameFor('renamed.docx');
  await expect.poll(() => bodyText(again)).toContain('kept across a rename');
  await tabIs('renamed.docx', true);
  // Saved under its own name, and watched there.
  await vs.page.keyboard.press('Control+s');
  await tabIs('renamed.docx', false);
  await vs.closeAll();
});

test('a page loaded again after a save opens the saved document, not the last backup (section 9.8 finding 4)', async () => {
  const frame = await vs.open('reload.docx');
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.type(' one', { delay: 20 });
  await sleep(2000);
  await vs.page.keyboard.press('Control+s');
  await tabIs('reload.docx', false);
  const t0 = await frame.evaluate(() => performance.timeOrigin);
  await vs.exec('workbench.action.webview.reloadWebviewAction');
  // A second ready: the page loads again (S3.1 item 5).
  let again = frame;
  await expect.poll(async () => {
    for (const f of await vs.frames()) {
      try { if (await f.evaluate(() => performance.timeOrigin) !== t0) { again = f; return true; } } catch { /* going */ }
    }
    return false;
  }).toBe(true);
  again = await vs.frameFor('reload.docx');
  await expect.poll(() => bodyText(again)).toContain(' one');
  await expect(again.getByTestId('name')).toHaveText('reload.docx');
  await tabIs('reload.docx', false);
  await vs.closeAll();
});

test('steady typing for five seconds, then a pause: no backup during the typing (P20)', async () => {
  const frame = await vs.open('steady.docx');
  await caretAtEndOf(frame, 0);
  const before = Date.now();
  const typing = Date.now() + 5000;
  while (Date.now() < typing) await vs.page.keyboard.type('x', { delay: 150 });
  const stopped = Date.now();
  const during = (await backups()).filter((b) => b.mtimeMs > before + 500 && b.mtimeMs < stopped);
  expect(during, 'backups written while typing').toEqual([]);
  await expect.poll(async () => (await backups()).filter((b) => b.mtimeMs >= stopped).length, { timeout: 15_000 }).toBeGreaterThan(0);
  await vs.closeAll();
});
