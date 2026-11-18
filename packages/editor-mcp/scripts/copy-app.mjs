// SPDX-License-Identifier: GPL-3.0-or-later
//
// The built application into the companion's own `dist/app`, so that the published package serves
// it without the (private) `@docx4j/editor` package beside it (ED-005 section 8.1). Run after
// `npm run build -w @docx4j/editor`; without that build the companion falls back to the
// repository's `packages/editor/dist`, or to `--app`.
import { cp, rm, stat } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const from = join(here, '..', '..', 'editor', 'dist');
const to = join(here, '..', 'dist', 'app');
try {
  await stat(join(from, 'index.html'));
} catch {
  console.error(`copy-app: no built application at ${from}; the companion will look for one at run time`);
  process.exit(0);
}
await rm(to, { recursive: true, force: true });
await cp(from, to, { recursive: true });
console.error(`copy-app: ${from} -> ${to}`);
