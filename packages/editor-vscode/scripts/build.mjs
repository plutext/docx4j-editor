// SPDX-License-Identifier: GPL-3.0-or-later
//
// The extension's code into `dist/extension.js` (ED-014 section 4.3): one CommonJS file for VS Code's extension host,
// with the application's `host-protocol.ts`, `webview-page.ts` and `ask-send.ts` bundled in, and the packages they import
// (the Anthropic SDK among them, section 4.6). `vscode` is the host's own module. No runtime dependency, so nothing else
// goes into the VSIX.
//
// Beside the manifest: THIRD-PARTY-NOTICES.txt, the notices of the packages esbuild took in (ED-014 section 4.7, P47), as
// `bundle-companion.mjs` writes the companion's, the build failing on one the licence check would refuse. It names the
// other two bundles' notices, `app/`'s and `companion/`'s, so the VSIX's root says where each is.
import { build } from 'esbuild';
import { writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { NOTICES_FILE, allowed, describePackage, noticesText, packageDirOf, withoutLicenceText } from '../../../scripts/notices.mjs';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const watch = process.argv.includes('--watch');
const options = {
  entryPoints: [join(root, 'src', 'extension.ts')],
  outfile: join(root, 'dist', 'extension.js'),
  bundle: true,
  format: 'cjs',
  platform: 'node',
  target: 'node20',
  external: ['vscode'],
  sourcemap: false,
  legalComments: 'none',
  logLevel: 'warning',
};
if (watch) {
  const { context } = await import('esbuild');
  const ctx = await context(options);
  await ctx.watch();
} else {
  const result = await build({ ...options, metafile: true });
  await writeNotices(result.metafile);
}

/** THIRD-PARTY-NOTICES.txt for `dist/extension.js`, from the modules esbuild bundled. */
async function writeNotices(metafile) {
  const dirs = new Set();
  for (const input of Object.keys(metafile.inputs)) {
    const dir = packageDirOf(resolve(input));
    if (dir) dirs.add(dir);
  }
  const parties = [...dirs].map((dir) => describePackage(dir)).filter((p) => p !== undefined);
  const refused = parties.filter((p) => !allowed(p.licence));
  if (refused.length > 0) {
    console.error(`build: bundled packages outside the licence allowlist: ${refused.map((p) => `${p.name}@${p.version} (${p.licence ?? 'no licence'})`).join(', ')}`);
    process.exit(1);
  }
  for (const p of withoutLicenceText(parties)) console.warn(`build: ${p.name}@${p.version} ships no licence file: supply its project's in scripts/licences/`);
  const others = 'The extension carries two more bundles, each with its own notices: the application, in app/\n'
    + '(app/THIRD-PARTY-NOTICES.txt), and the agent\'s companion, in companion/ (companion/THIRD-PARTY-NOTICES.txt).';
  await writeFile(join(root, NOTICES_FILE), noticesText('the VS Code extension\'s own code (docx4j-editor, dist/extension.js)', parties, 'LICENSE', [others]));
  console.error(`build: ${NOTICES_FILE} (${parties.length} packages: ${parties.map((p) => p.name).sort().join(', ')})`);
}
