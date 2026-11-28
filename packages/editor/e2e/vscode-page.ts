// SPDX-License-Identifier: GPL-3.0-or-later
//
// VS Code's side of the page, as the browser specs stand in for it (ED-014 sections 4.1 and 4.2): the fake
// `acquireVsCodeApi`, what the page posts, and the extension's messages sent to it. `host-vscode.spec.ts` serves the
// page on the preview server's origin; `webview.spec.ts` serves the page the extension's writer makes, on an origin of
// its own in a sandboxed frame. A message crosses `evaluate` as JSON, a byte array as `{ b64 }`.
import type { Frame, Page } from '@playwright/test';

/** In the page before any script: the fake API, what is posted, every keydown as the window's bubble phase sees it. */
export const VSCODE_API_INIT = (): void => {
  const w = window as unknown as Record<string, unknown>;
  const posted: unknown[] = [];
  const keys: KeyboardEvent[] = [];
  let acquired = 0;
  w['__posted'] = posted;
  w['__keys'] = keys;
  w['acquireVsCodeApi'] = () => {
    acquired++;
    if (acquired > 1) throw new Error('An instance of the VS Code API has already been acquired');
    return { postMessage: (m: unknown) => { posted.push(m); } };
  };
  window.addEventListener('keydown', (e) => { keys.push(e); });
  const toB64 = (bytes: Uint8Array): string => {
    let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    return btoa(s);
  };
  const fromB64 = (b64: string): Uint8Array => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
  /** A message as JSON can carry it, and back: a byte array as `{ b64 }`. */
  const out = (v: unknown): unknown => v instanceof Uint8Array ? { b64: toB64(v) } : Array.isArray(v) ? v.map(out) : v && typeof v === 'object' ? Object.fromEntries(Object.entries(v).map(([k, x]) => [k, out(x)])) : v;
  const back = (v: unknown): unknown => v && typeof v === 'object' && 'b64' in (v as object) ? fromB64((v as { b64: string }).b64) : Array.isArray(v) ? v.map(back) : v && typeof v === 'object' ? Object.fromEntries(Object.entries(v).map(([k, x]) => [k, back(x)])) : v;
  w['__take'] = () => posted.map(out);
  w['__send'] = (m: unknown) => { window.dispatchEvent(new MessageEvent('message', { data: back(m) })); };
};

export type Posted = { type: string; [key: string]: unknown };
/** Where the page runs: the tab itself, or the frame VS Code's host page holds it in. */
export type PageLike = Page | Frame;

export const b64 = (bytes: Uint8Array): { b64: string } => ({ b64: Buffer.from(bytes).toString('base64') });
export const bytesOf = (v: unknown): Uint8Array => new Uint8Array(Buffer.from((v as { b64: string }).b64, 'base64'));

export const posted = (page: PageLike): Promise<Posted[]> => (page as Page).evaluate(() => (window as unknown as { __take: () => Posted[] }).__take());
export const ofType = async (page: PageLike, type: string): Promise<Posted[]> => (await posted(page)).filter((m) => m.type === type);
export const send = (page: PageLike, message: Record<string, unknown>): Promise<void> => (page as Page).evaluate((m) => (window as unknown as { __send: (m: unknown) => void }).__send(m), message);
