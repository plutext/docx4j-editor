// SPDX-License-Identifier: GPL-3.0-or-later
//
// Monaco's workers started by their script's origin (ED-014 section 4.2, P22), over fakes of the page: directly on
// the page's origin, as a module worker where the URL names one (the dev server's); from a Blob of the fetched script
// on another origin (a VS Code webview's resource origin), fetched once. The browser spec `e2e/webview.spec.ts` runs
// the real thing on two origins.
import { describe, expect, test, vi } from 'vitest';
import { workerStarter, workerTypeOf, type WorkerPlace } from '../src/app/monaco-env.js';

interface Started { url: string; options: WorkerOptions }

function fakePlace(page: string, answer: (url: string) => Response | Promise<Response> = () => new Response('self.onmessage = () => {};', { status: 200 })) {
  const started: Started[] = [];
  const fetched: string[] = [];
  const blobs: Blob[] = [];
  const place: WorkerPlace = {
    base: () => page,
    origin: () => new URL(page).origin,
    fetch: async (url) => { fetched.push(url); return answer(url); },
    createObjectURL: (blob) => { blobs.push(blob); return `blob:${new URL(page).origin}/${blobs.length}`; },
    startWorker: (url, options) => { started.push({ url, options }); return { url } as unknown as Worker; },
  };
  return { place, started, fetched, blobs };
}

/** A Blob's text through the FileReader, which jsdom's Blob has where it has no `text()`. */
function textOf(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = () => reject(reader.error);
    reader.readAsText(blob);
  });
}

describe("Monaco's workers by their script's origin", () => {
  test('the module sets the environment over the URLs Vite gives the two workers, started on the page itself', () => {
    const made: Started[] = [];
    vi.stubGlobal('Worker', class { constructor(url: string, options: WorkerOptions) { made.push({ url, options }); } });
    try {
      const env = (self as unknown as { MonacoEnvironment: { getWorker(id: string, label: string): unknown } }).MonacoEnvironment;
      env.getWorker('workerMain.js', 'typescript');
      env.getWorker('workerMain.js', 'xml');
    } finally {
      vi.unstubAllGlobals();
    }
    // Vitest serves the imports as the dev server does: each worker's own file, a module, on the page's origin.
    expect(made.map((m) => new URL(m.url).origin)).toEqual([location.origin, location.origin]);
    expect(made[0]!.url).toMatch(/\/language\/typescript\/ts\.worker\.js\?worker_file&type=module$/);
    expect(made[1]!.url).toMatch(/\/editor\/editor\.worker\.js\?worker_file&type=module$/);
    expect(made.map((m) => m.options)).toEqual([{ name: 'typescript', type: 'module' }, { name: 'xml', type: 'module' }]);
  });

  test('a script on the page\'s origin is started directly, as a classic worker', () => {
    const { place, started, fetched } = fakePlace('http://localhost:4173/index.html');
    const start = workerStarter(place);
    start('http://localhost:4173/assets/ts.worker-abc.js', 'typescript');
    expect(started).toEqual([{ url: 'http://localhost:4173/assets/ts.worker-abc.js', options: { name: 'typescript', type: 'classic' } }]);
    expect(fetched).toEqual([]);
  });

  test("a URL that names type=module, as the dev server's does, is started as a module worker, resolved against the page", () => {
    const { place, started } = fakePlace('http://localhost:5173/');
    const start = workerStarter(place);
    start('/node_modules/monaco-editor/esm/vs/editor/editor.worker.js?worker_file&type=module', 'editorWorkerService');
    expect(started).toEqual([{
      url: 'http://localhost:5173/node_modules/monaco-editor/esm/vs/editor/editor.worker.js?worker_file&type=module',
      options: { name: 'editorWorkerService', type: 'module' },
    }]);
    expect(workerTypeOf('/w.js?worker_file&type=classic')).toBe('classic');
    expect(workerTypeOf('/w.js?type=modules')).toBe('classic');
    expect(workerTypeOf('/w.js')).toBe('classic');
  });

  test('a script on another origin is fetched and started from a Blob typed text/javascript', async () => {
    const { place, started, fetched, blobs } = fakePlace('vscode-webview://abc/index.html');
    const start = workerStarter(place);
    const worker = start('https://file+.vscode-resource.vscode-cdn.net/ext/app/assets/ts.worker-abc.js', 'typescript');
    expect(worker).toBeInstanceOf(Promise);
    expect(started).toEqual([]);
    await worker;
    expect(fetched).toEqual(['https://file+.vscode-resource.vscode-cdn.net/ext/app/assets/ts.worker-abc.js']);
    expect(blobs).toHaveLength(1);
    expect(blobs[0]!.type).toBe('text/javascript');
    expect(await textOf(blobs[0]!)).toBe('self.onmessage = () => {};');
    expect(started).toEqual([{ url: 'blob:null/1', options: { name: 'typescript', type: 'classic' } }]);
  });

  test('a relative URL under a <base> on another origin goes the Blob route', async () => {
    const { place, fetched, started } = fakePlace('http://127.0.0.1:9000/page');
    place.base = () => 'http://localhost:4173/';
    await workerStarter(place)('assets/editor.worker-abc.js', 'editorWorkerService');
    expect(fetched).toEqual(['http://localhost:4173/assets/editor.worker-abc.js']);
    expect(started[0]!.url).toMatch(/^blob:/);
  });

  test('two workers of one script fetch it once', async () => {
    const { place, started, fetched, blobs } = fakePlace('http://127.0.0.1:9000/');
    const start = workerStarter(place);
    const url = 'http://localhost:4173/assets/ts.worker-abc.js';
    await Promise.all([start(url, 'typescript'), start(url, 'javascript')]);
    await start(url, 'typescript');
    expect(fetched).toEqual([url]);
    expect(blobs).toHaveLength(1);
    expect(started.map((s) => s.url)).toEqual(['blob:http://127.0.0.1:9000/1', 'blob:http://127.0.0.1:9000/1', 'blob:http://127.0.0.1:9000/1']);
    expect(started.map((s) => s.options.name)).toEqual(['typescript', 'javascript', 'typescript']);
  });

  test('a failed fetch rejects the promise Monaco awaits, and the next start fetches again', async () => {
    let answers = 0;
    const { place, started, fetched } = fakePlace('http://127.0.0.1:9000/', () => {
      answers += 1;
      if (answers === 1) return Promise.reject(new TypeError('Failed to fetch'));
      if (answers === 2) return new Response('Not Found', { status: 404, statusText: 'Not Found' });
      return new Response('self.onmessage = () => {};', { status: 200 });
    });
    const start = workerStarter(place);
    const url = 'http://localhost:4173/assets/ts.worker-abc.js';
    await expect(start(url, 'typescript')).rejects.toThrow('Failed to fetch');
    await expect(start(url, 'typescript')).rejects.toThrow(`Monaco's worker script ${url} could not be fetched: 404 Not Found`);
    await start(url, 'typescript');
    expect(fetched).toEqual([url, url, url]);
    expect(started).toHaveLength(1);
  });
});
