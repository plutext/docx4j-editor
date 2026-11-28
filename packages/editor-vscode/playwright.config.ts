// SPDX-License-Identifier: GPL-3.0-or-later
//
// The suite in VS Code (ED-014 P31): `npm run test:vscode -w docx4j-editor`, after the root build. One VS Code at a time:
// each spec starts its own on a profile of its own, and drives it over the DevTools Protocol. No browser of Playwright's
// own is used. On Linux without a display, run it under `xvfb-run -a`.
import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: 'test/vscode',
  testMatch: '*.spec.ts',
  // The README's screenshot is made on demand (`npm run screenshot`), never in the suite.
  testIgnore: 'screenshot.spec.ts',
  workers: 1,
  fullyParallel: false,
  // A save of the 255-page fixture takes about 33 s in the page (ED-014 section 9.9).
  timeout: 240_000,
  expect: { timeout: 30_000 },
  retries: process.env['CI'] ? 1 : 0,
  reporter: process.env['CI'] ? [['list'], ['html', { open: 'never' }]] : 'list',
  outputDir: 'test-results/playwright',
});
