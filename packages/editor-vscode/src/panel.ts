// SPDX-License-Identifier: GPL-3.0-or-later
//
// A page's HTML (ED-014 section 4.3): `webviewPage` over the application's `index.html`, read once at activation, with
// the webview's own source, a nonce of 16 random bytes and the boot data (P19, P27).
import * as vscode from 'vscode';
import { randomBytes } from 'node:crypto';
import { webviewPage, type HostBoot } from './shared.js';

export interface App {
  /** The application's folder in the extension, `app/`. */
  root: vscode.Uri;
  indexHtml: string;
  /** The palette's names of the commands the page names, in VS Code's language (P49): read once, at activation. */
  commands: Record<string, string>;
}

/** The folder as the webview reaches it, ending in `/`: `asWebviewUri` of a folder returns none (section 9.8 finding 15). */
export function baseOf(webview: vscode.Webview, root: vscode.Uri): string {
  const base = webview.asWebviewUri(root).toString();
  return base.endsWith('/') ? base : `${base}/`;
}

/**
 * The page for a webview, with its options set: scripts, and the application's folder as its only resource root. `connect`:
 * the companion's relay and a room's on another relay, for `connect-src` (ED-014 section 4.5).
 */
export function writePage(webview: vscode.Webview, app: App, boot: HostBoot, connect: readonly string[] = []): void {
  webview.options = { enableScripts: true, localResourceRoots: [app.root] };
  webview.html = webviewPage(app.indexHtml, {
    base: baseOf(webview, app.root),
    source: webview.cspSource,
    nonce: randomBytes(16).toString('base64'),
    boot,
    ...(connect.length > 0 ? { connect } : {}),
  });
}

function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** A page of one line with no script: the Markdown launcher's, shown while it opens the Word document (P26). */
export function writeNotice(webview: vscode.Webview, text: string): void {
  webview.options = { enableScripts: false, localResourceRoots: [] };
  webview.html = `<!doctype html><html lang="${escapeHtml(vscode.env.language)}"><head><meta charset="utf-8">`
    + `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'">`
    + `<style>body{font-family:var(--vscode-font-family);color:var(--vscode-foreground);padding:2em}</style></head>`
    + `<body><p>${escapeHtml(text)}</p></body></html>`;
}
