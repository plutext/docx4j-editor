// SPDX-License-Identifier: GPL-3.0-or-later
import { defineConfig, type Plugin } from 'vite';
import react from '@vitejs/plugin-react';
import { companionRoutes } from '../editor-model/test/fixtures/build/companion-double.mjs';
import { askRoutes } from '../editor-model/test/fixtures/build/ask-double.mjs';
import { mkdtempSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { Server as HttpServer } from 'node:http';
import { createRelay, type Relay } from '@docx4j/editor-relay';
import { NOTICES_FILE, allowed, describePackage, noticesText, packageDirOf, withoutLicenceText } from '../../scripts/notices.mjs';

/**
 * The companion's contract served by the test double under `/companion` on the dev and preview
 * servers (ED-003 section 6.5): the browser tests point the companion option at it, and a
 * developer can too. Not part of the built site: a static deployment has no companion unless
 * one runs beside it.
 */
function companionDouble(): Plugin {
  return {
    name: 'docx4j-companion-double',
    configureServer(server) { server.middlewares.use(companionRoutes('/companion')); },
    configurePreviewServer(server) { server.middlewares.use(companionRoutes('/companion')); },
  };
}

/** The ask mode's provider double under `/ask` (ED-003 section 7.5): the Messages API's shape, answered by the prompt's words; the tests point the base-URL option at it. */
function askDouble(): Plugin {
  return {
    name: 'docx4j-ask-double',
    configureServer(server) { server.middlewares.use(askRoutes('/ask')); },
    configurePreviewServer(server) { server.middlewares.use(askRoutes('/ask')); },
  };
}

/**
 * The co-editing relay under `/relay` on the dev and preview servers (ED-004 section 4.2), its
 * rooms in a temporary folder per server: the browser tests point the relay option at it, and a
 * developer can too. Not part of the built site: a deployment runs `docx4j-editor-relay` (or E4's
 * companion) and sets `VITE_RELAY_URL`. Vite's own HMR socket is left to Vite.
 */
function relayMount(): Plugin {
  let relay: Relay | undefined;
  const mount = (http: HttpServer | null): void => {
    if (!http) return;
    relay ??= createRelay({ data: mkdtempSync(join(tmpdir(), 'docx4j-relay-')), debounce: 200 });
    http.on('upgrade', (request, socket, head) => {
      if (request.url === '/relay' || request.url?.startsWith('/relay/') || request.url?.startsWith('/relay?')) relay!.handleUpgrade(request, socket, head);
    });
    http.once('close', () => { void relay?.destroy(); relay = undefined; });
  };
  return {
    name: 'docx4j-relay',
    configureServer(server) { mount(server.httpServer as HttpServer | null); },
    configurePreviewServer(server) { mount(server.httpServer as HttpServer | null); },
  };
}

/**
 * The site's LICENSE.txt and THIRD-PARTY-NOTICES.txt, at its root, so that a deployment (the demo,
 * the companion's copy in editor-mcp's dist/app) carries both. LICENSE.txt is the repository's LICENSE
 * (the site is GPL-3.0-or-later like the rest; `.txt` so that a static host shows it rather than
 * offering a download), and the notices point to it. Hosting the site also distributes the
 * bundled packages, whose licences ask that their notices go with them. The notices are written from
 * the modules the build actually took in (a devDependency bundled counts too), and the build fails on
 * one whose licence the licence check would refuse (scripts/notices.mjs, shared with
 * scripts/check-licences.mjs).
 */
function thirdPartyNotices(): Plugin {
  return {
    name: 'docx4j-third-party-notices',
    apply: 'build',
    generateBundle() {
      const dirs = new Set<string>();
      for (const id of this.getModuleIds()) {
        const dir = packageDirOf(id);
        if (dir) dirs.add(dir);
      }
      const parties = [...dirs].map((dir) => describePackage(dir)).filter((p) => p !== undefined);
      const refused = parties.filter((p) => !allowed(p.licence));
      if (refused.length > 0) this.error(`bundled packages outside the licence allowlist: ${refused.map((p) => `${p.name}@${p.version} (${p.licence ?? 'no licence'})`).join(', ')}`);
      for (const p of withoutLicenceText(parties)) this.warn(`${p.name}@${p.version} ships no licence file: supply its project's in scripts/licences/`);
      this.emitFile({ type: 'asset', fileName: NOTICES_FILE, source: noticesText('the site (@docx4j/editor)', parties, 'LICENSE.txt') });
      this.emitFile({ type: 'asset', fileName: 'LICENSE.txt', source: readFileSync(fileURLToPath(new URL('../../LICENSE', import.meta.url))) });
    },
  };
}

// A static site: `vite build` gives dist/, served from anywhere, no server (ED-001 section 11
// item 4). The engine is ESM with no Node imports outside its /node subpath, so it bundles as is;
// the object model loads its 103 mapping modules through a static registry (objects CR-005,
// 0.2.0), so no bundler hint is needed. ED-002 section 10.3 item 1 records the hint that was
// here until 2026-09-25, and why it went.
//
// A relative base (ED-014 P21): `index.html` names `./assets/…` and the chunks find one another relative to themselves,
// so one build serves a site's root (the preview server, the companion's `/`), a VS Code webview under the `<base>` its
// page writes (`app/webview-page.ts`), and any folder it is copied to. The dev server and `vite preview` take it as `/`;
// the demo passes its own `--base`, which wins.
export default defineConfig({
  base: './',
  plugins: [react(), companionDouble(), askDouble(), relayMount(), thirdPartyNotices()],
  build: {
    target: 'es2022',
    sourcemap: true,
  },
  // Class names survive minification: the engine's errors name the class (`Word.Body.insertFileFromBase64 is not supported`, ED-003 section 7.3) and a minified `Word.Lc` tells nobody anything.
  esbuild: { keepNames: true },
  server: {
    fs: { allow: ['../..'] },
  },
});
