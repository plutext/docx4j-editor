// SPDX-License-Identifier: GPL-3.0-or-later
// ED-005 section 6.4's spike (e2e/spikes/review-spike.spec.ts) in Chromium and Firefox. WebKit needs
// system libraries this host lacks (libicu74, libflite1), recorded in ED-005 section 12.
import { defineConfig } from '@playwright/test';
import base from './playwright.config.js';

export default defineConfig({
  ...base,
  testDir: 'e2e/spikes',
  testIgnore: [],
  projects: [
    { name: 'chromium', use: { browserName: 'chromium' } },
    { name: 'firefox', use: { browserName: 'firefox' } },
  ],
});
