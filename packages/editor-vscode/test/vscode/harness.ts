// SPDX-License-Identifier: GPL-3.0-or-later
//
// The suite in VS Code (ED-014 P31): a real VS Code, started on a scratch profile with the extension and the suite's driver
// as development extensions, reached over the Chrome DevTools Protocol by Playwright, as S0 did. The page is driven with
// real keys and clicks; VS Code's side (its commands, its tabs, its clipboard) through the driver on a loopback port.
//
// - **VS Code:** `DOCX4J_VSCODE` names an executable; otherwise `@vscode/test-electron` downloads the version
//   `DOCX4J_VSCODE_VERSION` names (1.147.0 by default) into `.vscode-test/`, git-ignored.
// - **The profile:** a folder of its own under `test-results/`, with `window.dialogStyle: custom` and VS Code's own file
//   dialog (`files.simpleDialog.enable`), which keys can drive. On Linux VS Code runs with `--ozone-platform=x11`, without
//   which it crashes on Jason's machine (NVIDIA under Wayland); under xvfb it is X11 anyway.
// - **Connected first:** Playwright connects before any webview opens. A frame made before the connection is not seen.
// - **Stopped whole:** VS Code runs in a process group of its own, which a close kills, a crash's SIGKILL included. Killing
//   the main process alone left its renderers and extension host running, which loaded the machine and held a profile.
// - **No sandbox on Linux CI**, as `@vscode/test-electron` launches VS Code (microsoft/vscode#84238). A downloaded VS Code's
//   `chrome-sandbox` is not setuid root, so Chromium's sandbox needs an unprivileged user namespace, and without one VS Code
//   aborts at its start (reproduced under `bwrap --disable-userns`). On GitHub's Ubuntu runner every case had waited two
//   minutes for a page that never came. A run here keeps the sandbox, as a person's VS Code does.
// - **Secrets in plain text on Linux CI** (`--password-store=basic`): with no OS keyring, VS Code meets the first secret
//   stored (the Ask key) with a modal dialog. Given the switch, its main process turns plain-text encryption on at start
//   and stores the secret without asking (read in 1.147.0's `EncryptionMainService`).
// - **VS Code's output** goes to `vscode.log` beside its profile (a restart adds to it). A launch that fails ends at once
//   if VS Code has exited, and says the log's last lines either way.
// - `DOCX4J_SUITE_DEBUG=1` logs the DevTools connection's attempts.
import { chromium, expect, type Browser, type Frame, type Page } from '@playwright/test';
import { spawn, spawnSync, type ChildProcess } from 'node:child_process';
import { closeSync, cpSync, existsSync, mkdirSync, openSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { createServer } from 'node:net';
import { join } from 'node:path';

// The package is CommonJS, as VS Code loads the extension with `require`: Playwright compiles the specs so too.
const here = __dirname;
export const EXTENSION = join(here, '..', '..');
export const DRIVER = join(here, 'driver');
export const FIXTURES = join(EXTENSION, '..', 'editor-model', 'test', 'fixtures');
const RESULTS = join(EXTENSION, 'test-results', 'vscode');

export const sleep = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

async function freePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const server = createServer();
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address() as { port: number };
      server.close(() => resolve(port));
    });
  });
}

let executable: Promise<string> | undefined;

/** The VS Code to run: `DOCX4J_VSCODE`, else the pinned version downloaded once. */
export function vscodeExecutable(): Promise<string> {
  executable ??= (async () => {
    if (process.env['DOCX4J_VSCODE']) return process.env['DOCX4J_VSCODE'];
    const { downloadAndUnzipVSCode } = await import('@vscode/test-electron');
    return downloadAndUnzipVSCode({ version: process.env['DOCX4J_VSCODE_VERSION'] ?? '1.147.0', cachePath: join(EXTENSION, '.vscode-test') });
  })();
  return executable;
}

export interface LaunchOptions {
  /** The profile's name under `test-results/vscode/`; a second launch with the same name starts on the same profile. */
  profile: string;
  /** Fixtures copied into the workspace folder, by their names there: `{ 'a.docx': 'HelloWordOnline.docx' }`. */
  files?: Record<string, string>;
  /** VS Code's settings, over the suite's. */
  settings?: Record<string, unknown>;
  /** Trust off (the default) makes the workspace trusted; on, with no prompt, it opens in Restricted Mode. */
  trust?: boolean;
  /** Keep the profile and the workspace as an earlier launch left them. */
  keep?: boolean;
  /** A window with no folder: the workspace folder is made, for the files, but not opened. */
  noFolder?: boolean;
  /** More development extensions, beside the extension and the driver. */
  extensions?: string[];
  /** The driver registers the Ask box's test model, vendor `docx4j-test`, at its activation (ED-014 section 4.6). */
  lm?: boolean;
  /**
   * VSIXs installed through VS Code's command line into the profile before the window starts, the extension's and the
   * driver's among them, in place of the development paths (ED-014 P51). The window is then no Extension Development Host,
   * and it has the backup folder an installed extension's window has. Installed once, at a launch that does not `keep`.
   */
  installed?: string[];
  /** VS Code's display language (`--locale`); English when absent. */
  locale?: string;
}

export interface Tab {
  group: number;
  label: string;
  isDirty: boolean;
  isActive: boolean;
  viewType: string | null;
  kind: string | null;
  uri: string | null;
}

/** One VS Code: its process, its workbench page, the driver, and the page's policy violations heard. */
export class VsCode {
  /** Content security policy refusals the pages logged, and errors they threw. */
  readonly problems: string[] = [];
  /** VS Code's version, as the driver reports it at launch. */
  version = '';

  private constructor(
    readonly options: LaunchOptions,
    readonly folder: string,
    private readonly child: ChildProcess,
    readonly browser: Browser,
    readonly page: Page,
    private readonly driverPort: number,
  ) {
    page.on('console', (m) => {
      const url = m.location().url;
      const fromPage = url.startsWith('vscode-webview://') || url.includes('vscode-resource') || url === '';
      if (fromPage && m.type() === 'error' && /Content Security Policy/.test(m.text())) this.problems.push(m.text());
    });
  }

  static async launch(options: LaunchOptions): Promise<VsCode> {
    const root = join(RESULTS, options.profile);
    const folder = join(root, 'work');
    if (!options.keep) {
      rmSync(root, { recursive: true, force: true });
      mkdirSync(join(root, 'user', 'User'), { recursive: true });
      mkdirSync(folder, { recursive: true });
      for (const [name, from] of Object.entries(options.files ?? {})) cpSync(join(FIXTURES, from), join(folder, name));
    }
    writeFileSync(join(root, 'user', 'User', 'settings.json'), JSON.stringify({
      'window.dialogStyle': 'custom',
      'files.simpleDialog.enable': true,
      'workbench.startupEditor': 'none',
      'update.mode': 'none',
      'extensions.autoUpdate': false,
      'telemetry.telemetryLevel': 'off',
      'workbench.enableExperiments': false,
      'chat.disableAIFeatures': true,
      'security.workspace.trust.startupPrompt': 'never',
      'window.restoreWindows': 'all',
      'files.hotExit': 'onExitAndWindowClose',
      // The suite's folders are inside this checkout: no Git, whose offer to open the repository otherwise comes up.
      'git.enabled': false,
      'git.openRepositoryInParentFolders': 'never',
      ...options.settings,
    }, null, 2));
    const cdp = await freePort();
    const driverPort = await freePort();
    if (options.installed && options.extensions) throw new Error('installed and extensions together: an installed window takes no development path');
    const profile = [`--user-data-dir=${join(root, 'user')}`, `--extensions-dir=${join(root, 'exts')}`];
    if (options.installed && !options.keep) await install(options.installed, profile);
    const development = options.installed ? [] : [EXTENSION, DRIVER, ...(options.extensions ?? [])];
    const args = [
      ...(process.platform === 'linux' ? ['--ozone-platform=x11'] : []),
      ...(process.platform === 'linux' && process.env['CI'] ? ['--no-sandbox', '--disable-gpu-sandbox', '--password-store=basic'] : []),
      ...profile,
      ...development.map((e) => `--extensionDevelopmentPath=${e}`),
      ...(options.locale ? [`--locale=${options.locale}`] : []),
      `--remote-debugging-port=${cdp}`,
      '--skip-welcome', '--skip-release-notes', '--disable-telemetry', '--disable-updates',
      ...(options.trust ? [] : ['--disable-workspace-trust']),
      ...(options.noFolder ? [] : [folder]),
    ];
    const env: NodeJS.ProcessEnv = {
      ...process.env,
      DOCX4J_DRIVER_PORT: String(driverPort),
      DOCX4J_DRIVER_LM: options.lm ? 'on' : 'off',
      // The Ask box's double, which an installed driver cannot find by its own path (ED-014 section 9.24 item 4).
      DOCX4J_DRIVER_DOUBLE: join(FIXTURES, 'build', 'ask-double.mjs'),
    };
    delete env['ELECTRON_RUN_AS_NODE'];
    // A process group of its own, so that a close stops VS Code's helpers with it: a renderer left running an export
    // would load the machine for the next launch.
    if (process.env['DOCX4J_SUITE_DEBUG']) console.log('launch:', args.join(' '));
    const log = join(root, 'vscode.log');
    const out = openSync(log, 'a');
    const child = spawn(await vscodeExecutable(), args, { env, stdio: ['ignore', out, out], detached: true });
    closeSync(out);
    const { browser, page } = await connect(cdp, child, log);
    const vs = new VsCode(options, folder, child, browser, page, driverPort);
    await vs.waitForDriver();
    vs.version = (await vs.ctl<{ version: string }>({ op: 'ping' })).version;
    return vs;
  }

  /** VS Code's main process. */
  get pid(): number | undefined {
    return this.child.pid;
  }

  /** A path in the workspace folder. */
  path(name: string): string {
    return join(this.folder, name);
  }

  async ctl<T = unknown>(op: Record<string, unknown>): Promise<T> {
    const r = await fetch(`http://127.0.0.1:${this.driverPort}/`, { method: 'POST', body: JSON.stringify(op) });
    const j = await r.json() as { ok: boolean; out?: T; error?: string };
    if (!j.ok) throw new Error(j.error);
    return j.out as T;
  }

  private async waitForDriver(): Promise<void> {
    for (let i = 0; i < 240; i++) {
      try { await this.ctl({ op: 'ping' }); return; } catch { await sleep(500); }
    }
    throw new Error('the driver never answered');
  }

  exec<T = unknown>(command: string, ...args: unknown[]): Promise<T> {
    return this.ctl<T>({ op: 'exec', command, args });
  }

  tabs(): Promise<Tab[]> {
    return this.ctl<Tab[]>({ op: 'tabs' });
  }

  async tab(label: string): Promise<Tab | undefined> {
    return (await this.tabs()).find((t) => t.label === label);
  }

  /**
   * Opens a file of the workspace folder in its default editor, and waits for its page's document. VS Code 1.105.0 opens a
   * file asked for within a few seconds of its start in the text editor, before it knows the custom editors (measured, ED-014
   * section 9.22): a Word file opened so is closed and opened again, for up to 15 s, after which the text editor is a failure.
   */
  async open(name: string): Promise<Frame> {
    const until = Date.now() + 15_000;
    for (;;) {
      await this.exec('vscode.open', { $uri: this.path(name) });
      const tab = await this.tab(name);
      if (!tab || tab.kind !== 'text' || !/\.(docx|dotx|docm|dotm)$/i.test(name) || Date.now() > until) break;
      await this.ctl({ op: 'closeTab', label: tab.label });
      await sleep(1000);
    }
    return this.frameFor(name);
  }

  /** The page showing this document, once its view is up: by its name, a " •" after it while unsaved. */
  async frameFor(name: string, timeout = 120_000): Promise<Frame> {
    const until = Date.now() + timeout;
    while (Date.now() < until) {
      for (const frame of this.page.frames()) {
        try {
          const shown = await frame.evaluate(() => document.querySelector('[data-testid="name"]')?.textContent ?? null);
          if (shown !== null && shown.replace(/ •$/, '') === name) {
            await frame.locator('.ProseMirror[aria-label="Document"]').waitFor({ timeout: Math.max(1000, until - Date.now()) });
            return frame;
          }
        } catch { /* a frame that went, or is not the application's */ }
      }
      await sleep(250);
    }
    if (process.env['DOCX4J_SUITE_DEBUG']) {
      const seen = [];
      for (const frame of this.page.frames()) seen.push(`${frame.url().slice(0, 60)} ${await frame.evaluate(() => `${!!document.querySelector('meta[name="docx4j-host"]')} ${document.querySelector('[data-testid="name"]')?.textContent ?? '-'} ${document.querySelector('[data-testid="opening"]') ? 'opening' : ''}`).catch((e) => String(e).slice(0, 80))}`);
      console.log('frames:', seen, 'tabs:', (await this.tabs()).map((t) => t.label));
    }
    throw new Error(`no page shows ${name}`);
  }

  /** The application's pages now. */
  async frames(): Promise<Frame[]> {
    const out: Frame[] = [];
    for (const frame of this.page.frames()) {
      try { if (await frame.evaluate(() => !!document.querySelector('meta[name="docx4j-host"]'))) out.push(frame); } catch { /* not one */ }
    }
    return out;
  }

  /** VS Code's notifications, as their text. */
  async notifications(): Promise<string[]> {
    return this.page.evaluate(() => [...document.querySelectorAll('.notification-toast')].map((e) => (e as HTMLElement).innerText.replace(/\s+/g, ' ').trim()));
  }

  /** VS Code's own dialog (`window.dialogStyle: custom`), as its text, or null. */
  async dialog(): Promise<string | null> {
    return this.page.evaluate(() => (document.querySelector('.monaco-dialog-box') as HTMLElement | null)?.innerText ?? null);
  }

  async clickDialogButton(label: string): Promise<void> {
    await this.page.locator('.monaco-dialog-box .monaco-button', { hasText: label }).first().click();
  }

  /**
   * VS Code's Save As through its own file dialog: the command, the path typed, Enter. Onto a file that exists, the dialog
   * asks in the quick input itself ("… already exists. Are you sure you want to overwrite it?"); `replace` answers Yes.
   */
  async saveAs(path: string, options: { replace?: boolean } = {}): Promise<void> {
    void this.exec('workbench.action.files.saveAs').catch(() => undefined);
    const input = this.page.locator('.quick-input-widget input:not([type="checkbox"])');
    await input.waitFor({ timeout: 15_000 });
    await input.fill(path);
    await sleep(200);
    await this.page.keyboard.press('Enter');
    if (options.replace) {
      await expect(this.page.locator('.quick-input-widget')).toContainText('already exists', { timeout: 15_000 });
      // Yes is the prompt's first item, and Enter takes it.
      await sleep(300);
      await this.page.keyboard.press('Enter');
    }
  }

  /** Whether this VS Code is older than `major.minor`. */
  olderThan(major: number, minor: number): boolean {
    const [a = 0, b = 0] = this.version.split('.').map(Number);
    return a < major || (a === major && b < minor);
  }

  /**
   * After a Save As onto another file, until `done`. VS Code before 1.147 (1.105 among them) leaves the old editor dirty
   * and asks "Do you want to save the changes you made to <old>?" as it closes it. Jason accepted that prompt (ED-014
   * section 9.10), and the README tells a person to answer Don't Save, as this does. From 1.147.0 on the prompt is a
   * failure, since that release clears the dirty state. Says whether the prompt came.
   */
  async afterSaveAs(oldLabel: string, done: () => Promise<boolean>, timeout = 60_000): Promise<boolean> {
    let prompted = false;
    const until = Date.now() + timeout;
    while (Date.now() < until) {
      const dialog = await this.dialog();
      if (dialog?.includes(`the changes you made to ${oldLabel}`)) {
        if (!this.olderThan(1, 147)) throw new Error(`VS Code ${this.version} asked about ${oldLabel} after a Save As: ${dialog}`);
        prompted = true;
        await this.clickDialogButton("Don't Save");
      }
      if (await done()) return prompted;
      await sleep(250);
    }
    throw new Error(`the Save As from ${oldLabel} did not settle in ${timeout} ms`);
  }

  /** Every editor closed, unsaved edits discarded: VS Code's save prompt answered Don't Save for each. */
  async closeAll(): Promise<void> {
    await this.exec('notifications.clearAll').catch(() => undefined);
    const closing = this.exec('workbench.action.closeAllEditors').catch(() => undefined);
    const until = Date.now() + 30_000;
    while (Date.now() < until) {
      const done = await Promise.race([closing.then(() => true), sleep(250).then(() => false)]);
      if (done && (await this.tabs()).length === 0) return;
      const dialog = await this.dialog();
      if (dialog?.includes('Do you want to save')) await this.clickDialogButton("Don't Save");
      else if (done) await this.exec('workbench.action.closeAllEditors').catch(() => undefined);
    }
    throw new Error(`editors left open: ${(await this.tabs()).map((t) => t.label).join(', ')}`);
  }

  /** Stops VS Code: SIGTERM, or SIGKILL for a crash, and waits for it to go. */
  async close(signal: NodeJS.Signals = 'SIGTERM'): Promise<void> {
    await this.browser.close().catch(() => undefined);
    const group = (sig: NodeJS.Signals): void => { try { process.kill(-this.child.pid!, sig); } catch { /* gone */ } };
    if (this.child.exitCode === null && this.child.signalCode === null) {
      const gone = new Promise<void>((r) => this.child.once('exit', () => r()));
      // A crash is every process at once; a close is the main process's, which closes the rest.
      if (signal === 'SIGKILL') group('SIGKILL'); else this.child.kill(signal);
      await Promise.race([gone, sleep(15_000)]);
      if (this.child.exitCode === null && this.child.signalCode === null) this.child.kill('SIGKILL');
    }
    // Whatever of VS Code's is still running goes now.
    group('SIGKILL');
    // VS Code's helper processes can outlive the main one for a moment: give its profile's locks time to go.
    await sleep(1500);
  }

  /** The same profile, started again. */
  async restart(signal: NodeJS.Signals = 'SIGTERM'): Promise<VsCode> {
    await this.close(signal);
    return VsCode.launch({ ...this.options, keep: true });
  }
}

/** Installs VSIXs into a profile through VS Code's command line, as a person's `code --install-extension` does. */
async function install(vsixes: string[], profile: string[]): Promise<void> {
  const { resolveCliPathFromVSCodeExecutablePath } = await import('@vscode/test-electron');
  const cli = resolveCliPathFromVSCodeExecutablePath(await vscodeExecutable());
  for (const vsix of vsixes) {
    const env: NodeJS.ProcessEnv = { ...process.env };
    delete env['ELECTRON_RUN_AS_NODE'];
    const run = spawnSync(cli, [...profile, '--install-extension', vsix, '--force'], { env, encoding: 'utf8', timeout: 120_000 });
    if (run.status !== 0) throw new Error(`VS Code did not install ${vsix} (${run.status ?? run.signal}): ${run.stderr || run.stdout}`);
  }
}

async function connect(port: number, child: ChildProcess, log: string): Promise<{ browser: Browser; page: Page }> {
  const until = Date.now() + 120_000;
  while (Date.now() < until) {
    if (child.exitCode !== null || child.signalCode !== null) {
      throw new Error(`VS Code exited (${child.exitCode ?? child.signalCode}) before its workbench page came. ${logTail(log)}`);
    }
    try {
      const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`);
      for (let i = 0; i < 120; i++) {
        for (const context of browser.contexts()) for (const page of context.pages()) if (page.url().includes('workbench')) return { browser, page };
        await sleep(250);
      }
      if (process.env['DOCX4J_SUITE_DEBUG']) console.log('pages:', browser.contexts().flatMap((c) => c.pages().map((p) => p.url())));
      await browser.close();
    } catch (e) {
      if (process.env['DOCX4J_SUITE_DEBUG']) console.log('connect:', String(e).slice(0, 200));
      await sleep(500);
    }
  }
  throw new Error(`no workbench page over the DevTools Protocol. ${logTail(log)}`);
}

/** The last lines VS Code wrote, for a launch's failure. */
function logTail(log: string): string {
  let text = '';
  try { text = readFileSync(log, 'utf8').trimEnd(); } catch { /* none */ }
  return text ? `Its output ended (${log}):\n${text.split('\n').slice(-20).join('\n')}` : `It wrote nothing (${log}).`;
}

/** The body's caret put at the end of block `i`, through the view (a click races the browser's selection report). */
export async function caretAtEndOf(frame: Frame, i: number): Promise<void> {
  await frame.evaluate((index) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { nodeSize: number } } }; caretAt(p: number): void; focus(): void } }).docx4jEditor;
    let pos = 0;
    for (let k = 0; k < index; k++) pos += v.state.doc.child(k).nodeSize;
    v.caretAt(pos + v.state.doc.child(index).nodeSize - 1);
    v.focus();
  }, i);
}

/** Text inserted at the body's selection through the view: an edit the page makes, which a busy page runs when it is free. */
export function insertInPage(frame: Frame, text: string): Promise<void> {
  return frame.evaluate((t) => {
    const view = (window as unknown as { docx4jEditor: { view: { dispatch(tr: unknown): void; state: { tr: { insertText(t: string): unknown } } } } }).docx4jEditor.view;
    view.dispatch(view.state.tr.insertText(t));
  }, text);
}

/**
 * Before another program writes a file the suite has just opened: VS Code 1.90's file watcher missed a change made within
 * about 1.5 s of the document opening, and reported one made a few seconds later (ED-014 section 9.10). No person writes so
 * soon; 1.147.0 reported both.
 */
export const watcherSettles = (): Promise<void> => sleep(3000);

/** The body's text, as the page holds it. */
export function bodyText(frame: Frame): Promise<string> {
  return frame.evaluate(() => (window as unknown as { docx4jEditor: { state: { doc: { textContent: string } } } }).docx4jEditor.state.doc.textContent);
}

export function exists(path: string): boolean {
  return existsSync(path);
}
