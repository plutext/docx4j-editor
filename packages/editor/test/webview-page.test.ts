// SPDX-License-Identifier: GPL-3.0-or-later
//
// The page a VS Code webview loads (ED-014 section 4.2, P24): section 3.5's policy with no `blob:` in `script-src` (P23 as
// amended, section 9.9) and with `base-uri` and `form-action` (section 9.7 note 2), the inserts after the charset and before every script and link, the nonce, the boot data read back by the host's
// own reader, and each refusal. `e2e/webview.spec.ts` loads the page it writes in Chromium.
import { describe, expect, test } from 'vitest';
import { webviewPage, webviewPolicy, hostBootMeta, WebviewPageError, type WebviewPageOptions } from '../src/app/webview-page.js';
import { readHostBoot } from '../src/app/host.js';
import { HOST_META } from '../src/app/host-protocol.js';

/** The built `index.html` under the relative base (P21), as `vite build` writes it. */
const BUILT = `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1, interactive-widget=resizes-content" />
    <title>docx4j editor</title>
    <script type="module" crossorigin src="./assets/index-HpstMlDr.js"></script>
    <link rel="stylesheet" crossorigin href="./assets/index-zfcf0NWJ.css">
  </head>
  <body>
    <div id="root"></div>
  </body>
</html>
`;

const SOURCE = `'self' https://*.vscode-cdn.net`;
const options = (more: Partial<WebviewPageOptions> = {}): WebviewPageOptions => ({
  base: 'https://file+.vscode-resource.vscode-cdn.net/home/me/.vscode/extensions/docx4j/app/',
  source: SOURCE,
  nonce: 'r4nd0m+N0nce/=_-',
  boot: { options: {} },
  ...more,
});

const parse = (html: string): Document => new DOMParser().parseFromString(html, 'text/html');

describe('the policy', () => {
  test("section 3.5's, with no blob: in script-src, and base-uri and form-action, which default-src does not cover", () => {
    expect(webviewPolicy(options())).toBe([
      `default-src 'none'`,
      `base-uri 'self' https://*.vscode-cdn.net`,
      `form-action 'none'`,
      `script-src 'nonce-r4nd0m+N0nce/=_-' 'self' https://*.vscode-cdn.net`,
      `style-src 'self' https://*.vscode-cdn.net 'unsafe-inline'`,
      `img-src 'self' https://*.vscode-cdn.net blob: data:`,
      `font-src 'self' https://*.vscode-cdn.net blob: data:`,
      `worker-src blob:`,
      `connect-src 'self' https://*.vscode-cdn.net`,
    ].join('; '));
  });

  test('the companion and its relay in connect-src', () => {
    const policy = webviewPolicy(options({ connect: ['http://127.0.0.1:4174', 'ws://127.0.0.1:4175'] }));
    expect(policy.split('; ').at(-1)).toBe(`connect-src 'self' https://*.vscode-cdn.net http://127.0.0.1:4174 ws://127.0.0.1:4175`);
    expect(policy.split('; ').slice(0, -1)).toEqual(webviewPolicy(options()).split('; ').slice(0, -1));
  });
});

describe('the page', () => {
  test('the inserts come after the charset, before the viewport, every script and every link', () => {
    const page = webviewPage(BUILT, options());
    const head = parse(page).head;
    const tags = [...head.children].map((el) => el.outerHTML.replace(/\s(content|href|src|nonce)="[^"]*"/g, ' $1'));
    expect(tags).toEqual([
      '<meta charset="utf-8">',
      '<meta http-equiv="Content-Security-Policy" content>',
      '<base href>',
      '<meta property="csp-nonce" nonce>',
      `<meta name="${HOST_META}" content>`,
      '<meta name="viewport" content>',
      '<title>docx4j editor</title>',
      '<script nonce type="module" crossorigin="" src></script>',
      '<link rel="stylesheet" crossorigin="" href>',
    ]);
    const doc = parse(page);
    expect(doc.querySelector('meta[http-equiv="Content-Security-Policy"]')!.getAttribute('content')).toBe(webviewPolicy(options()));
    expect(doc.querySelector('base')!.getAttribute('href')).toBe(options().base);
    // The charset stays within the first bytes, however long the boot data.
    expect(page.indexOf('<meta charset')).toBe(BUILT.indexOf('<meta charset'));
    // Nothing else moves.
    expect(page.replace(/\n\s*<meta http-equiv[^\n]*|\n\s*<base[^\n]*|\n\s*<meta property="csp-nonce"[^\n]*|\n\s*<meta name="docx4j-host"[^\n]*/g, '').replace(/ nonce="[^"]*"/g, '')).toBe(BUILT);
  });

  test('the script resolves against the base, and so does what the page fetches', () => {
    const doc = parse(webviewPage(BUILT, options()));
    const script = doc.querySelector('script')!;
    expect(new URL(script.getAttribute('src')!, doc.querySelector('base')!.href).href).toBe(`${options().base}assets/index-HpstMlDr.js`);
  });

  test('the nonce on every script, and in the csp-nonce meta for the preload helper', () => {
    const two = BUILT.replace('</head>', '<script>window.x = 1;</script>\n  </head>');
    const doc = parse(webviewPage(two, options()));
    const scripts = [...doc.querySelectorAll('script')];
    expect(scripts).toHaveLength(2);
    for (const script of scripts) expect(script.getAttribute('nonce')).toBe(options().nonce);
    expect(doc.querySelector('meta[property="csp-nonce"]')!.getAttribute('nonce')).toBe(options().nonce);
  });

  test('a page with no charset takes the inserts straight after <head>', () => {
    const page = webviewPage(BUILT.replace('    <meta charset="utf-8" />\n', ''), options());
    expect(parse(page).head.firstElementChild!.getAttribute('http-equiv')).toBe('Content-Security-Policy');
    expect(page).toMatch(/<head>\n {4}<meta http-equiv="Content-Security-Policy"/);
  });

  test('the boot data holding ", &, < and text outside ASCII is read back by the host as it was written', () => {
    const boot = { options: { 'docx4j-editor.authorName': 'Zoë "the editor" <zoe> & Ōno 名前 \u{1F600}', 'docx4j-editor.language': 'fr' } };
    const doc = parse(webviewPage(BUILT, options({ boot })));
    expect(readHostBoot(doc)).toEqual(boot);
    const meta = hostBootMeta(boot);
    expect(meta).not.toMatch(/content="[^"]*"[^>]+"/);
    expect(meta).toContain('&quot;');
    expect(meta).toContain('&amp;');
    expect(meta).toContain('&lt;zoe>');
    expect(readHostBoot(parse(`<head>${meta}</head>`))).toEqual(boot);
  });

  test('VS Code\'s language and the workspace\'s trust go with the options, and are read back (P27)', () => {
    const boot = { options: {}, language: 'zh-cn', trusted: false };
    expect(readHostBoot(parse(webviewPage(BUILT, options({ boot }))))).toEqual(boot);
    // Neither is written when not given, nor anything else.
    const meta = hostBootMeta({ options: {}, ...({ other: 'x' } as object) });
    expect(meta).not.toContain('language');
    expect(meta).not.toContain('other');
  });

  test('whether a key is stored and the palette\'s names go with the options, and are read back (ED-014 sections 4.6 and 4.7)', () => {
    const boot = { options: {}, askKey: true, commands: { 'docx4j.ask.setKey': 'docx4j editor: Définir la clé API Anthropic…' } };
    expect(readHostBoot(parse(webviewPage(BUILT, options({ boot }))))).toEqual(boot);
    expect(readHostBoot(parse(webviewPage(BUILT, options({ boot: { options: {}, askKey: false, commands: {} } }))))).toEqual({ options: {}, askKey: false });
  });

  test('a base with a quote is written as an address, never as markup', () => {
    const doc = parse(webviewPage(BUILT, options({ base: 'https://example.test/a"b<c/' })));
    expect(doc.querySelectorAll('base')).toHaveLength(1);
    expect(doc.querySelector('base')!.getAttribute('href')).toBe('https://example.test/a%22b%3Cc/');
  });
});

describe('the refusals', () => {
  const refused = (html: string, more: Partial<WebviewPageOptions>, why: RegExp): void => {
    expect(() => webviewPage(html, options(more))).toThrow(WebviewPageError);
    expect(() => webviewPage(html, options(more))).toThrow(why);
  };

  test('a base that is not absolute, or does not end in /', () => {
    refused(BUILT, { base: './app/' }, /absolute/);
    refused(BUILT, { base: '/app/' }, /absolute/);
    refused(BUILT, { base: 'https://file+.vscode-resource.vscode-cdn.net/ext/app' }, /end in \//);
  });

  test('a nonce outside base64 and base64url', () => {
    for (const nonce of ['', 'a b', "a'b", 'a;b', 'a"b', 'a<b', 'nonce\n']) refused(BUILT, { nonce }, /nonce/);
    expect(() => webviewPolicy(options({ nonce: "x' 'unsafe-eval" }))).toThrow(WebviewPageError);
  });

  test('a source or a connect entry that would rewrite the policy', () => {
    for (const source of ['', '  ', "'self'; script-src *", "'self', script-src *", `"self"`, '`self`', "'self' <x>", "'self'\nscript-src *", "'self'\u0000"]) refused(BUILT, { source }, /source/);
    for (const entry of ['', 'http://127.0.0.1:4174 *', 'http://a;b', 'http://a,b', 'ws://"a"', 'http://a\tb']) refused(BUILT, { connect: [entry] }, /connect entry/);
  });

  test('a page with no <head>', () => {
    refused('<!doctype html><html><body><script src="./a.js"></script></body></html>', {}, /no <head>/);
  });

  test('a page already holding a <base>, a policy or the boot meta', () => {
    refused(BUILT.replace('<title>', '<base href="https://elsewhere.test/">\n    <title>'), {}, /<base>/);
    refused(BUILT.replace('<title>', `<meta http-equiv="content-security-policy" content="default-src *">\n    <title>`), {}, /content security policy/);
    refused(BUILT.replace('<title>', `<meta name="${HOST_META}" content="{}">\n    <title>`), {}, /boot meta/);
    refused(BUILT.replace('<script ', '<script nonce="old" '), {}, /already has a nonce/);
  });

  test('a script, link or style before the charset, which would load outside the policy and the base', () => {
    refused(BUILT.replace('<meta charset="utf-8" />', '<script src="./early.js"></script>\n    <meta charset="utf-8" />'), {}, /before the charset/);
    refused(BUILT.replace('<meta charset="utf-8" />', '<link rel="stylesheet" href="./early.css">\n    <meta charset="utf-8" />'), {}, /before the charset/);
  });
});
