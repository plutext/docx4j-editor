// SPDX-License-Identifier: GPL-3.0-or-later
import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    // The modules with no `vscode` import (ED-014 section 4.3); the suite in VS Code is Playwright's (`test:vscode`).
    include: ['test/unit/**/*.test.ts'],
  },
});
