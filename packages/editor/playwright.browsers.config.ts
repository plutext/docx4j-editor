// SPDX-License-Identifier: GPL-3.0-or-later
// ED-005 section 9.5 (E4.g step G4): the review spike and the review, tracking and comments specs in
// Firefox and WebKit (and since ED-006 P6 the print spec, its PDF parts Chromium's), apart from `npm run e2e`, which stays Chromium's. Firefox runs on this host
// (`npm run e2e:browsers -- --project firefox`); WebKit needs system libraries this host lacks
// (libicu74, libflite1), so it runs on CI's Ubuntu through the `browsers` workflow, started by hand
// and never gating a push (proposal 40). What fails is a finding: fixed in the editor, or stated in
// `docs/user/review.md`.
import { defineConfig } from '@playwright/test';
import base from './playwright.config.js';

export default defineConfig({
  ...base,
  testDir: 'e2e',
  testIgnore: [],
  testMatch: ['spikes/review-spike.spec.ts', 'review.spec.ts', 'tracking.spec.ts', 'comments.spec.ts', 'comment-split.spec.ts', 'excluded-characters.spec.ts', 'bookmarks.spec.ts', 'print.spec.ts'],
  projects: [
    { name: 'firefox', use: { browserName: 'firefox' } },
    { name: 'webkit', use: { browserName: 'webkit' } },
  ],
});
