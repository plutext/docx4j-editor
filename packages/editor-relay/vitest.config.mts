// SPDX-License-Identifier: GPL-3.0-or-later
import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['test/**/*.test.mts'],
    // The fixtures are real documents; a 400-page one takes a while to unmarshal.
    testTimeout: 60_000,
    hookTimeout: 60_000,
  },
});
