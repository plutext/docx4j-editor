// SPDX-License-Identifier: GPL-3.0-or-later
//
// The companion as the extension ships it (ED-014 section 4.5, P36): `@docx4j/editor-mcp`'s `dist/cli.mjs` and its relay's
// worker, `dist/relay-worker.mjs`, bundled by esbuild into `companion/` (git-ignored) with every package they import, since
// the VSIX carries no `node_modules`. ESM for Node 22, as VS Code 1.105's Electron runs Node 22.19 and 1.147's Node 24; the
// banner gives the CommonJS packages bundled into ESM the `require` they call. The companion's served application is the
// extension's own `app/`, passed as `--app` at the spawn, so it ships once.
//
// Beside the two files: THIRD-PARTY-NOTICES.txt, from the packages esbuild took in (scripts/notices.mjs, as the site's
// build writes its own), the build failing on one the licence check would refuse. Run after the companion's build.
import { build } from 'esbuild';
import { rm, stat, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { NOTICES_FILE, allowed, describePackage, noticesText, packageDirOf, withoutLicenceText } from '../../../scripts/notices.mjs';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const from = join(root, '..', 'editor-mcp', 'dist');
const to = join(root, 'companion');
try {
  await stat(join(from, 'cli.mjs'));
} catch {
  console.error(`bundle-companion: no built companion at ${from}; build @docx4j/editor-mcp first`);
  process.exit(1);
}
await rm(to, { recursive: true, force: true });
const result = await build({
  entryPoints: { cli: join(from, 'cli.mjs'), 'relay-worker': join(from, 'relay-worker.mjs') },
  outdir: to,
  outExtension: { '.js': '.mjs' },
  bundle: true,
  format: 'esm',
  platform: 'node',
  target: 'node22',
  // ws's optional native helpers: absent, ws uses its own JavaScript.
  external: ['bufferutil', 'utf-8-validate'],
  banner: { js: "import { createRequire as __cr } from 'node:module'; const require = __cr(import.meta.url);" },
  legalComments: 'none',
  metafile: true,
  logLevel: 'warning',
});

const dirs = new Set();
for (const input of Object.keys(result.metafile.inputs)) {
  const dir = packageDirOf(resolve(input));
  if (dir) dirs.add(dir);
}
const parties = [...dirs].map((dir) => describePackage(dir)).filter((p) => p !== undefined);
const refused = parties.filter((p) => !allowed(p.licence));
if (refused.length > 0) {
  console.error(`bundle-companion: bundled packages outside the licence allowlist: ${refused.map((p) => `${p.name}@${p.version} (${p.licence ?? 'no licence'})`).join(', ')}`);
  process.exit(1);
}
for (const p of withoutLicenceText(parties)) console.warn(`bundle-companion: ${p.name}@${p.version} ships no licence file: supply its project's in scripts/licences/`);
await writeFile(join(to, NOTICES_FILE), noticesText('the companion as the VS Code extension ships it (@docx4j/editor-mcp, bundled)', parties, '../LICENSE'));
const size = async (name) => ((await stat(join(to, name))).size / 1024 / 1024).toFixed(1);
console.error(`bundle-companion: ${from} -> ${to} (cli.mjs ${await size('cli.mjs')} MB, relay-worker.mjs ${await size('relay-worker.mjs')} MB, ${parties.length} packages)`);
