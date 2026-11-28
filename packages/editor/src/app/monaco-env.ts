// SPDX-License-Identifier: GPL-3.0-or-later
//
// Monaco's workers (ED-003 section 7.3; ED-005 section 3.8; ED-014 section 4.2, P22): the TypeScript worker for the
// console, the editor worker for everything else (the XML editors have no language worker). Set by `monaco.ts`, the
// one importer of Monaco, which only the console's and the XML editors' chunks load, so the Edit perspective's bundle
// never carries Monaco.
//
// A worker is started by where its script is, never by the host:
// - on the page's origin (a site, the companion, the preview and dev servers), directly, as Vite's own `?worker`
//   constructor would: a module worker where the URL's query names one (the dev server serves the worker as an ES
//   module, `…?worker_file&type=module`), a classic one otherwise (the build's workers are classic bundles);
// - on another origin (a VS Code webview, whose files are on its resource origin; a site whose files are on a CDN),
//   from a Blob of the script, fetched once and kept for the page's life, since `new Worker` refuses a script on
//   another origin with a SecurityError (ED-014 section 9.1 item 8). The script never becomes a string: the response's
//   Blob is wrapped in one typed `text/javascript`. Only a self-contained script runs from a Blob, as the build's are.
import editorWorkerUrl from 'monaco-editor/esm/vs/editor/editor.worker?worker&url';
import tsWorkerUrl from 'monaco-editor/esm/vs/language/typescript/ts.worker?worker&url';

/** What starting a worker needs from the page: the page's own by default, fakes in `test/monaco-env.test.ts`. */
export interface WorkerPlace {
  /** The address a relative worker URL resolves against: the document's base, as `new Worker` resolves it. */
  base(): string;
  /** The page's own origin. */
  origin(): string;
  fetch(url: string): Promise<Response>;
  createObjectURL(blob: Blob): string;
  startWorker(url: string, options: WorkerOptions): Worker;
}

export const pagePlace: WorkerPlace = {
  base: () => globalThis.document?.baseURI ?? self.location.href,
  origin: () => self.location.origin,
  fetch: (url) => fetch(url),
  createObjectURL: (blob) => URL.createObjectURL(blob),
  startWorker: (url, options) => new Worker(url, options),
};

/** The type a worker's URL names, as Vite writes it on the dev server (`type=module`); classic without one. */
export function workerTypeOf(url: string): WorkerType {
  return /[?&]type=module(?:[&#]|$)/.test(url) ? 'module' : 'classic';
}

/** A starter of workers by their script's origin, the Blob of each script on another origin made once. */
export function workerStarter(place: WorkerPlace = pagePlace): (url: string, name: string) => Worker | Promise<Worker> {
  const blobs = new Map<string, Promise<string>>();
  const blobOf = (href: string): Promise<string> => {
    let made = blobs.get(href);
    if (!made) {
      const fetched = place.fetch(href).then(async (response) => {
        if (!response.ok) throw new Error(`Monaco's worker script ${href} could not be fetched: ${response.status} ${response.statusText}`.trimEnd());
        // The script's bytes, not `response.blob()`: a Blob made in another realm is not a part every `Blob` takes (jsdom's,
        // on Node 22, wrote it as the text "[object Blob]").
        return place.createObjectURL(new Blob([await response.arrayBuffer()], { type: 'text/javascript' }));
      });
      // A failure is not kept, so that the next editor fetches again.
      fetched.catch(() => { if (blobs.get(href) === fetched) blobs.delete(href); });
      blobs.set(href, fetched);
      made = fetched;
    }
    return made;
  };
  return (url, name) => {
    const href = new URL(url, place.base()).href;
    const options: WorkerOptions = { name, type: workerTypeOf(url) };
    if (new URL(href).origin === place.origin()) return place.startWorker(href, options);
    return blobOf(href).then((blobUrl) => place.startWorker(blobUrl, options));
  };
}

const start = workerStarter();

(self as unknown as { MonacoEnvironment?: unknown }).MonacoEnvironment = {
  getWorker(_: string, label: string) {
    if (label === 'typescript' || label === 'javascript') return start(tsWorkerUrl, label);
    return start(editorWorkerUrl, label);
  },
};
