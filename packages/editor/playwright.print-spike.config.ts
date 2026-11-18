// SPDX-License-Identifier: GPL-3.0-or-later
// ED-006 step P0: the print spike alone, in Chromium and Firefox (WebKit lacks system libraries on
// this host). Firefox is launched with its silent print-to-file preferences, so that `window.print()`
// writes a PDF the spike can read; Playwright's `page.pdf` is Chromium's only.
import { defineConfig } from '@playwright/test';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import base from './playwright.config.js';

export const FIREFOX_PDF = join(tmpdir(), `print-spike-firefox-${process.pid}.pdf`);
process.env['PRINT_SPIKE_FIREFOX_PDF'] = FIREFOX_PDF;

export default defineConfig({
  ...base,
  testDir: 'e2e',
  testIgnore: [],
  testMatch: ['spikes/print-spike.spec.ts'],
  webServer: undefined,
  projects: [
    { name: 'chromium', use: { browserName: 'chromium' } },
    {
      name: 'firefox',
      use: {
        browserName: 'firefox',
        launchOptions: {
          firefoxUserPrefs: {
            'print.always_print_silent': true,
            'print.show_print_progress': false,
            'print_printer': 'Mozilla Save to PDF',
            'print.printer_Mozilla_Save_to_PDF.print_to_file': true,
            'print.printer_Mozilla_Save_to_PDF.print_to_filename': FIREFOX_PDF,
            'print.printer_Mozilla_Save_to_PDF.print_bgcolor': true,
          },
        },
      },
    },
  ],
});
