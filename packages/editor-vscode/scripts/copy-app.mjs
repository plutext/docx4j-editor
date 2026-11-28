// SPDX-License-Identifier: GPL-3.0-or-later
//
// The built application into the extension's own `app/`, which its webviews load (ED-014 section 4.3): the editor's
// `dist` without its source maps, as the companion's `copy-app.mjs` copies it (335 files, 20.3 MB, section 9.6). Run after
// `npm run build -w @docx4j/editor`. `app/` is git-ignored.
import { cp, rm, stat } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const from = join(here, '..', '..', 'editor', 'dist');
const to = join(here, '..', 'app');
try {
  await stat(join(from, 'index.html'));
} catch {
  console.error(`copy-app: no built application at ${from}; build @docx4j/editor first`);
  process.exit(1);
}
await rm(to, { recursive: true, force: true });
await cp(from, to, { recursive: true, filter: (src) => !src.endsWith('.map') });
console.error(`copy-app: ${from} -> ${to}`);
