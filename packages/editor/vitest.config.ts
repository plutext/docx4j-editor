// SPDX-License-Identifier: GPL-3.0-or-later
import { defineConfig } from 'vitest/config';
import react from '@vitejs/plugin-react';
import { fileURLToPath } from 'node:url';

/**
 * Two projects. The jsdom one runs the application's tests against the DOM the application runs
 * against: jsonix is its browser entry there (jsonix 3.3.0's `node` entry would parse the custom
 * XML parts with xmldom while the XPath engines, the native `document.evaluate` and FontoXPath,
 * evaluate over jsdom's DOM: two DOMs and quietly wrong answers, ED-003 section 11.8), and the
 * `@docx4j` packages are transformed by Vite so that the alias reaches their imports. The node
 * one runs the files that say `@vitest-environment node` (a package saved under jsdom cannot be
 * reloaded there, ED-003 section 11.2 item 18), with jsonix's own Node entry.
 */
const jsonixBrowser = fileURLToPath(new URL('../../node_modules/@docx4j/jsonix/jsonix.mjs', import.meta.url));
const nodeFiles = ['test/ask-send.test.ts', 'test/newdoc.test.ts', 'test/room-client.test.ts', 'test/normalisation.test.ts', 'test/changed-since.test.ts'];

export default defineConfig({
  plugins: [react()],
  test: {
    testTimeout: 60_000,
    // Node 25 and later have a `localStorage` global of their own, undefined without `--localstorage-file`, and under
    // jsdom it hides jsdom's: every test that stores something failed on Node 26. Off, jsdom's is the one the tests see.
    poolOptions: { forks: { execArgv: ['--no-experimental-webstorage'] } },
    projects: [
      {
        extends: true,
        resolve: { alias: { '@docx4j/jsonix': jsonixBrowser } },
        test: { name: 'jsdom', environment: 'jsdom', include: ['test/**/*.test.{ts,tsx}'], exclude: nodeFiles, server: { deps: { inline: [/@docx4j\//] } } },
      },
      { extends: true, test: { name: 'node', environment: 'node', include: nodeFiles } },
    ],
  },
});
