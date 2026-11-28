// SPDX-License-Identifier: GPL-3.0-or-later
//
// The build as a VS Code webview loads it (ED-014 section 4.2, S2.4), in Chromium. The page is the one the extension's
// writer makes (`webviewPage` over the built `index.html`): section 3.5's policy, a `<base>`, the nonce and the boot
// data. It is served from a server of this spec's own on `127.0.0.1`, inside an `<iframe>` with the webview's sandbox
// flags (section 9.1 item 7), and the application's files come from the preview server, a second origin, as a webview's
// come from VS Code's resource origin. Not served through `page.route`: a routed page is not local to Chromium.
//
// What it cannot reproduce: VS Code's `vscode-webview:` scheme, the service worker that serves the resource origin,
// and Electron's Chromium. S3's extension-host test covers those.
import { test, expect, type Frame, type Page } from '@playwright/test';
import { createServer, type Server } from 'node:http';
import { readdirSync, readFileSync } from 'node:fs';
import { randomBytes } from 'node:crypto';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { AddressInfo } from 'node:net';
import { webviewPage } from '../src/app/webview-page.js';
import { VSCODE_API_INIT, b64, ofType, send } from './vscode-page.js';
import { fixture } from './helpers.js';

/** The application's files, as the webview reaches them on VS Code's resource origin. */
const FILES = 'http://localhost:4173/';
const FILES_ORIGIN = new URL(FILES).origin;
/** The webview's sandbox (S0, section 9.1 item 7): no `allow-popups`, no `allow-modals`. */
const SANDBOX = 'allow-scripts allow-same-origin allow-forms allow-pointer-lock allow-downloads';
const NONCE = randomBytes(16).toString('base64');
const here = dirname(fileURLToPath(import.meta.url));
const workerFile = (prefix: string): string => readdirSync(join(here, '..', 'dist', 'assets')).find((f) => f.startsWith(prefix) && f.endsWith('.js'))!;

let server: Server;
let pageOrigin: string;

test.beforeAll(async () => {
  server = createServer((req, res) => {
    const url = new URL(req.url ?? '/', 'http://127.0.0.1');
    if (url.pathname === '/') {
      // VS Code's host page: the frame its content runs in, given the window.
      res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
      res.end(`<!doctype html><html><head><meta charset="utf-8"><title>host</title><style>html,body{margin:0;height:100%}iframe{border:0;width:100%;height:100%;display:block}</style></head><body><iframe sandbox="${SANDBOX}" src="/page${url.search}"></iframe></body></html>`);
      return;
    }
    if (url.pathname === '/page') {
      const boot = { options: JSON.parse(url.searchParams.get('boot') ?? '{}') as Record<string, string> };
      fetch(`${FILES}index.html`).then((r) => r.text()).then((built) => {
        const page = webviewPage(built, { base: FILES, source: `'self' ${FILES_ORIGIN}`, nonce: NONCE, boot });
        res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
        res.end(page);
      }, (e: unknown) => { res.writeHead(500); res.end(String(e)); });
      return;
    }
    res.writeHead(404);
    res.end();
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  pageOrigin = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
});

test.afterAll(async () => {
  await new Promise<void>((resolve) => server.close(() => resolve()));
});

/** In the frame before any script: what the policy refused, what was thrown, and each worker started with its URL's scheme. */
const RECORD_INIT = (): void => {
  if (window === window.top) return;
  const record = { violations: [] as string[], errors: [] as string[], workers: [] as { scheme: string; name?: string; type: string }[], NativeWorker: window.Worker };
  (window as unknown as { __webview: typeof record }).__webview = record;
  document.addEventListener('securitypolicyviolation', (e) => { record.violations.push(`${e.violatedDirective} ${e.blockedURI}`); }, true);
  window.addEventListener('error', (e) => { record.errors.push(String(e.message)); });
  window.addEventListener('unhandledrejection', (e) => { record.errors.push(String((e.reason as Error | undefined)?.message ?? e.reason)); });
  const Native = window.Worker;
  window.Worker = class extends Native {
    constructor(url: string | URL, options?: WorkerOptions) {
      record.workers.push({ scheme: new URL(String(url), document.baseURI).protocol, name: options?.name, type: options?.type ?? 'classic' });
      super(url, options);
    }
  };
};

type Recorded = { violations: string[]; errors: string[]; workers: { scheme: string; name?: string; type: string }[] };
const recordOf = (frame: Frame): Promise<Recorded> => frame.evaluate(() => {
  const r = (window as unknown as { __webview: Recorded }).__webview;
  return { violations: [...r.violations], errors: [...r.errors], workers: r.workers.map((w) => ({ ...w })) };
});

/** The page VS Code would load, in its frame, waiting for the extension's `open`. */
async function webview(page: Page, boot: Record<string, string> = {}): Promise<Frame> {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(e.message));
  (page as unknown as { __errors: string[] }).__errors = errors;
  await page.addInitScript(VSCODE_API_INIT);
  await page.addInitScript(RECORD_INIT);
  await page.goto(`${pageOrigin}/?boot=${encodeURIComponent(JSON.stringify(boot))}`);
  const frame = await (await page.waitForSelector('iframe')).contentFrame();
  await frame!.getByTestId('opening').waitFor();
  return frame!;
}

/** The extension's `open`, and the waits `openFixture` has. */
async function open(frame: Frame, name: string): Promise<void> {
  await send(frame, { type: 'open', name, bytes: b64(await fixture(name)) });
  await frame.locator('.docx-editor .ProseMirror[aria-label="Document"]').waitFor();
  await frame.getByTestId('toolbar').waitFor();
}

/** A probe's violations heard, and forgotten before the case goes on. */
async function forget(frame: Frame): Promise<void> {
  await frame.evaluate(() => { (window as unknown as { __webview: Recorded }).__webview.violations.length = 0; });
}

/** After each case: nothing the policy refused, nothing thrown, in the frame or the tab. */
async function clean(page: Page, frame: Frame): Promise<void> {
  const record = await recordOf(frame);
  expect(record.violations, 'securitypolicyviolation').toEqual([]);
  expect(record.errors, 'errors in the frame').toEqual([]);
  expect((page as unknown as { __errors: string[] }).__errors, 'page errors').toEqual([]);
}

const caretInBlock = (frame: Frame, i: number): Promise<void> => frame.evaluate((index) => {
  const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { nodeSize: number } } }; caretAt(pos: number): void } }).docx4jEditor;
  let pos = 0;
  for (let k = 0; k < index; k++) pos += v.state.doc.child(k).nodeSize;
  v.caretAt(pos + 1);
}, i);

test('the condition: the page on its own origin under the policy, its files on another, where new Worker throws a SecurityError; a fixture opened through open', async ({ page }) => {
  const frame = await webview(page);
  expect(await frame.evaluate(() => location.origin)).toBe(pageOrigin);
  expect(await frame.evaluate(() => document.baseURI)).toBe(FILES);
  expect(await frame.evaluate(() => document.querySelector('meta[http-equiv="Content-Security-Policy"]')?.getAttribute('content'))).toContain(`'nonce-${NONCE}'`);
  // As in S0 (section 9.1 item 8): a worker's script on the files' origin cannot be started from the page.
  const tsWorker = `${FILES}assets/${workerFile('ts.worker-')}`;
  expect(await frame.evaluate((url) => {
    const Native = (window as unknown as { __webview: { NativeWorker: typeof Worker } }).__webview.NativeWorker;
    try { new Native(url); return 'started'; } catch (e) { return (e as Error).name; }
  }, tsWorker)).toBe('SecurityError');
  // The record hears the policy: a fetch outside connect-src is refused before it leaves the page.
  await frame.evaluate(() => fetch('http://localhost:4174/never').catch(() => undefined));
  await expect.poll(async () => (await recordOf(frame)).violations).toEqual(['connect-src http://localhost:4174/never']);
  await forget(frame);
  // Two rules `default-src` does not cover (ED-014 section 9.7 note 2). A `<base>` to anywhere but the files is refused:
  // while it stands first, the base falls back to the page's own address, never to the one it names.
  expect(await frame.evaluate(() => {
    const other = document.createElement('base');
    other.href = 'http://localhost:4174/';
    document.head.prepend(other);
    const during = document.baseURI;
    other.remove();
    return { fellBackToThePage: during === location.href, after: document.baseURI };
  })).toEqual({ fellBackToThePage: true, after: FILES });
  await expect.poll(async () => (await recordOf(frame)).violations).toEqual(['base-uri http://localhost:4174/']);
  await forget(frame);
  // And no form is submitted: the application's forms are all handled in the page.
  await frame.evaluate(() => {
    const form = document.createElement('form');
    form.action = 'http://localhost:4174/never';
    form.id = 'never-sent';
    document.body.append(form);
    form.submit();
  });
  await expect.poll(async () => (await recordOf(frame)).violations).toEqual(['form-action http://localhost:4174/never?']);
  await frame.evaluate(() => document.getElementById('never-sent')?.remove());
  await forget(frame);
  // The sandbox as VS Code's: no popups, no modals.
  expect(await frame.evaluate(() => window.open('about:blank'))).toBeNull();
  expect(await ofType(frame, 'ready')).toEqual([{ type: 'ready', origin: pageOrigin }]);
  await open(frame, 'tracked-changes.docx');
  await expect(frame.locator('.docx-editor .ProseMirror[aria-label="Document"] > p').nth(4)).toHaveText('Some text.');
  await clean(page, frame);
});

test("the console: TypeScript's diagnostic from a worker started from a Blob, and a Run inserting a paragraph (P23)", async ({ page }) => {
  const frame = await webview(page);
  await open(frame, 'tracked-changes.docx');
  await frame.getByTestId('perspective-developer').click();
  const consolePanel = frame.getByTestId('panel-console');
  await expect(consolePanel.locator('.console-editor .monaco-editor')).toBeVisible();
  await expect(frame.getByTestId('console-types')).toContainText('types loaded');
  await frame.waitForFunction(() => (window as unknown as { docx4jConsole?: unknown }).docx4jConsole !== undefined);
  const setScript = (s: string): Promise<void> => frame.evaluate((v) => { (window as unknown as { docx4jConsole: { setValue(v: string): void } }).docx4jConsole.setValue(v); }, s);
  // Monaco's diagnostics, from the TypeScript worker.
  await setScript(`await Word.run(async (context) => {\n  const body = context.document.body;\n  body.insertFileFromBase64('AAAA', 'End');\n});`);
  await frame.getByTestId('console-check').click();
  const output = frame.getByTestId('console-output');
  await expect(output).toContainText('line 3: not in the subset');
  await expect(output).toContainText('insertFileFromBase64');
  // The Run imports the compiled script from a Blob URL (`console-run.ts`), allowed by the nonce, which an import() takes
  // from the nonced script that calls it (section 9.6): the policy has no `blob:` in script-src (P23 as amended, 9.9).
  await setScript(`const p = body.paragraphs[4].insertParagraph('Run from a Blob', 'After');\nreturn p.text;`);
  await frame.getByTestId('console-run').click();
  await expect(output).toContainText('Run from a Blob');
  await expect(frame.locator('.docx-editor .ProseMirror[aria-label="Document"] > p').nth(5)).toHaveText('Run from a Blob');
  const record = await recordOf(frame);
  expect(record.workers.filter((w) => w.name === 'typescript' || w.name === 'javascript')).not.toEqual([]);
  for (const w of record.workers) expect(w, 'every worker started from a Blob, as a classic script').toMatchObject({ scheme: 'blob:', type: 'classic' });
  await clean(page, frame);
});

test("reveal codes: Monaco's find widget and folds, the editor worker started from a Blob, Monaco's stylesheet loaded with the nonce", async ({ page }) => {
  const frame = await webview(page);
  await open(frame, 'tracked-changes.docx');
  await frame.getByTestId('perspective-developer').click();
  await frame.getByTestId('panel-xml').waitFor();
  await caretInBlock(frame, 4);
  await frame.waitForFunction(() => (window as unknown as { docx4jXmlEditors?: { [id: string]: { getValue(): string } } }).docx4jXmlEditors?.['xml-editor']?.getValue().includes('Some text.'));
  const xml = frame.getByTestId('xml-editor');
  await xml.locator('.monaco-editor textarea').focus();
  await page.keyboard.press('Control+f');
  const find = xml.locator('.find-widget.visible');
  await expect(find).toHaveCount(1);
  // Monaco's own stylesheet places the widget: a lazy chunk's CSS, through the base.
  await expect(find).toHaveCSS('position', 'absolute');
  await page.keyboard.press('Escape');
  // The folds, from Monaco's folding model.
  await expect.poll(() => frame.evaluate(async () => {
    const ed = (window as unknown as { docx4jXmlEditors: { [id: string]: { getContribution(id: string): { getFoldingModel(): Promise<{ regions: { length: number } } | null> } } } }).docx4jXmlEditors['xml-editor']!;
    return (await ed.getContribution('editor.contrib.folding').getFoldingModel())?.regions.length ?? 0;
  })).toBeGreaterThan(0);
  // The editor worker (no TypeScript here), started from a Blob.
  await expect.poll(async () => (await recordOf(frame)).workers.length).toBeGreaterThan(0);
  for (const w of (await recordOf(frame)).workers) expect(w).toMatchObject({ scheme: 'blob:', type: 'classic' });
  expect((await recordOf(frame)).workers.some((w) => w.name !== 'typescript' && w.name !== 'javascript')).toBe(true);
  // The preload helper put the nonce from the csp-nonce meta on the links it made.
  const made = await frame.evaluate(() => [...document.querySelectorAll('link')].filter((l) => l.nonce !== '').map((l) => ({ rel: l.rel, nonce: l.nonce })));
  expect(made.some((l) => l.rel === 'stylesheet')).toBe(true);
  for (const l of made) expect(l.nonce).toBe(NONCE);
  await clean(page, frame);
});

test('the chrome in French from the boot options: a locale chunk through the base', async ({ page }) => {
  const fr = JSON.parse(readFileSync(join(here, '..', 'src', 'locales', 'fr.json'), 'utf8')) as Record<string, string>;
  const frame = await webview(page, { 'docx4j-editor.language': 'fr' });
  await expect(frame.locator('html')).toHaveAttribute('lang', 'fr');
  await expect(frame.getByTestId('opening')).toHaveText(fr['app.opening']!);
  await open(frame, 'tracked-changes.docx');
  await expect(frame.locator('button[data-command="font.bold"]')).toHaveAttribute('title', fr['toolbar.bold']!);
  await clean(page, frame);
});

test('the Review panel, a lazy panel, drawn with its CSS', async ({ page }) => {
  const frame = await webview(page);
  await open(frame, 'tracked-changes.docx');
  await frame.getByTestId('perspective-review').click();
  const panel = frame.getByTestId('panel-review');
  await expect(panel).toBeVisible();
  const entry = panel.getByTestId('review-entry').first();
  await expect(entry).toBeVisible();
  await expect(entry).toHaveCSS('display', 'flex');
  await clean(page, frame);
});
