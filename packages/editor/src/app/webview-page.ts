// SPDX-License-Identifier: GPL-3.0-or-later
//
// The page a VS Code webview loads (ED-014 section 4.2, P24): the built `index.html` with section 3.5's content
// security policy, a `<base>` at the application's folder, the nonce and the boot data. The extension (`src/webview.ts`,
// S3) calls `webviewPage` at each `resolveCustomEditor`; `e2e/webview.spec.ts` serves the page it writes, so the
// browser suite holds the page VS Code will load. It imports only `host-protocol.ts`, which imports nothing, so the
// extension shares it as it is.
//
// It refuses what would load nothing, or load under a policy other than the one asked for, and says why: the policy is
// written from its arguments, so one that could end a directive or an attribute is refused, never escaped into place.
import { HOST_META, type HostBoot } from './host-protocol.js';

export interface WebviewPageOptions {
  /** The application's folder as the webview reaches it, ending in `/`: `asWebviewUri` of the extension's copy. */
  base: string;
  /** The webview's own sources: `webview.cspSource` (`'self' https://*.vscode-cdn.net` in S0). */
  source: string;
  /** Fresh for each page. */
  nonce: string;
  /** More addresses for `connect-src`: the companion's relay, and a room's on another relay (ED-014 section 4.5). */
  connect?: readonly string[];
  /** The boot data (P19). */
  boot: HostBoot;
}

/** A page `webviewPage` will not write, with the reason. */
export class WebviewPageError extends Error {
  override name = 'WebviewPageError';
}

/** A nonce as the policy and an attribute can hold it raw: base64 and base64url. */
const NONCE = /^[A-Za-z0-9+/=_-]+$/;
/** What would end a directive (`;`), start a second policy (`,`), end an attribute or a tag, or break a line. */
const UNSAFE = /[;,"`<>\u0000-\u001f\u007f]/;

function check(options: Pick<WebviewPageOptions, 'source' | 'nonce' | 'connect'>): void {
  if (!NONCE.test(options.nonce)) throw new WebviewPageError(`the nonce may hold only A-Z, a-z, 0-9, +, /, =, _ and -: ${JSON.stringify(options.nonce)}`);
  if (options.source.trim() === '') throw new WebviewPageError('the source is empty: the page could load none of its files');
  if (UNSAFE.test(options.source)) throw new WebviewPageError(`the source holds a character that would rewrite the policy: ${JSON.stringify(options.source)}`);
  for (const entry of options.connect ?? []) {
    if (entry === '' || /\s/.test(entry) || UNSAFE.test(entry)) throw new WebviewPageError(`a connect entry must be one source expression: ${JSON.stringify(entry)}`);
  }
}

/** An attribute value, double-quoted: `&`, `"` and `<` escaped. */
function attribute(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
}

/**
 * Section 3.5's policy. `script-src` has no `blob:` (P23 as amended, ED-014 section 9.9): the console's Run and a macro
 * import a Blob, and the import takes the nonce of the script that calls it, measured in Chromium (section 9.6) and in VS
 * Code's Electron (section 9.9). `base-uri` and `form-action` are stated because `default-src` covers neither (section
 * 9.7 note 2): a `<base>` only to the files, and no form submitted, the application's forms all being handled in the page.
 */
export function webviewPolicy(options: Pick<WebviewPageOptions, 'source' | 'nonce' | 'connect'>): string {
  check(options);
  const source = options.source.trim().split(/\s+/).join(' ');
  return [
    `default-src 'none'`,
    `base-uri ${source}`,
    `form-action 'none'`,
    `script-src 'nonce-${options.nonce}' ${source}`,
    `style-src ${source} 'unsafe-inline'`,
    `img-src ${source} blob: data:`,
    `font-src ${source} blob: data:`,
    `worker-src blob:`,
    `connect-src ${[source, ...(options.connect ?? [])].join(' ')}`,
  ].join('; ');
}

/** The boot meta (`HOST_META`), its content the boot data as JSON, escaped as an attribute; the host reads it back through `getAttribute`. */
export function hostBootMeta(boot: HostBoot): string {
  const data: HostBoot = {
    options: boot.options,
    ...(typeof boot.language === 'string' ? { language: boot.language } : {}),
    ...(typeof boot.trusted === 'boolean' ? { trusted: boot.trusted } : {}),
    ...(typeof boot.agent === 'string' ? { agent: boot.agent } : {}),
    // S6's and S7's fields: whether a key is stored, and the palette's names of the commands the page names. The writer
    // copies each field by name, so a new one is added here too (ED-014 section 9.26: both had been dropped).
    ...(typeof boot.askKey === 'boolean' ? { askKey: boot.askKey } : {}),
    ...(boot.commands && Object.keys(boot.commands).length > 0 ? { commands: boot.commands } : {}),
  };
  return `<meta name="${HOST_META}" content="${attribute(JSON.stringify(data))}">`;
}

const HEAD = /<head(?:\s[^>]*)?>/i;
const CHARSET = /<meta\s[^>]*\bcharset\s*=[^>]*>/i;

/** The built `index.html` as the webview's page: the inserts after its charset (or its `<head>`), and the nonce on its scripts. */
export function webviewPage(indexHtml: string, options: WebviewPageOptions): string {
  const policy = webviewPolicy(options);
  let base: URL;
  try {
    base = new URL(options.base);
  } catch {
    throw new WebviewPageError(`the base must be an absolute address: ${JSON.stringify(options.base)}`);
  }
  if (!options.base.endsWith('/')) throw new WebviewPageError(`the base must end in /, or the last segment is dropped: ${JSON.stringify(options.base)}`);
  const head = HEAD.exec(indexHtml);
  if (!head) throw new WebviewPageError('the page has no <head>');
  if (/<base[\s>/]/i.test(indexHtml)) throw new WebviewPageError('the page already holds a <base>: the first one wins');
  if (/<meta\s[^>]*http-equiv\s*=\s*["']?content-security-policy/i.test(indexHtml)) throw new WebviewPageError('the page already holds a content security policy');
  if (new RegExp(`<meta\\s[^>]*name\\s*=\\s*["']?${HOST_META}["'\\s>]`, 'i').test(indexHtml)) throw new WebviewPageError('the page already holds the boot meta');
  if (/<script\b[^>]*\snonce\s*=/i.test(indexHtml)) throw new WebviewPageError('a script in the page already has a nonce: the first attribute wins');
  const charset = CHARSET.exec(indexHtml.slice(head.index));
  const at = charset ? head.index + charset.index + charset[0].length : head.index + head[0].length;
  // Whatever comes before the policy and the base loads under neither.
  if (/<(?:script|link|style)\b/i.test(indexHtml.slice(0, at))) throw new WebviewPageError('a script, link or style comes before the charset: it would load outside the policy and the base');
  const inserts = [
    `<meta http-equiv="Content-Security-Policy" content="${attribute(policy)}">`,
    `<base href="${attribute(base.href)}">`,
    `<meta property="csp-nonce" nonce="${attribute(options.nonce)}">`,
    hostBootMeta(options.boot),
  ].map((tag) => `\n    ${tag}`).join('');
  const page = indexHtml.slice(0, at) + inserts + indexHtml.slice(at);
  return page.replace(/<script\b/gi, `<script nonce="${attribute(options.nonce)}"`);
}
