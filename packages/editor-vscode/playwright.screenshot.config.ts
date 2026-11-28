// SPDX-License-Identifier: GPL-3.0-or-later
//
// The README's screenshot (ED-014 section 4.7): `npm run screenshot -w docx4j-editor`, after the root build. The suite's
// configuration, its one spec `screenshot.spec.ts`, which the suite itself leaves out.
import { defineConfig } from '@playwright/test';
import suite from './playwright.config';

export default defineConfig({ ...suite, testMatch: 'screenshot.spec.ts', testIgnore: [], retries: 0 });
