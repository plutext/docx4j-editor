// SPDX-License-Identifier: GPL-3.0-or-later
//
// The VSIX (ED-014 section 4.7, P47): the built extension packaged by `vsce` (pinned, a devDependency), then read back and
// checked. Run after the root build; `npm run vsix` at the root does both.
//
//   node scripts/package.mjs [--out <folder>]   packages into <folder>/docx4j-editor-<version>.vsix (by default this
//                                                package's own folder, where the file is git-ignored), then checks it
//   node scripts/package.mjs --check <vsix>     checks a VSIX made before
//   --from <folder>                              packages another extension's folder in place of this one (the tests')
//
// What goes in is the manifest's `files`, an allowlist: `vsce` refuses an entry that matches no file, so a missing notice
// or folder stops it. What came out is checked in the VSIX's own names (`vsce` puts every file under `extension/`, renames
// the README `extension/readme.md` and `LICENSE` `extension/LICENSE.txt`, and adds two files of its own): under the
// budget, no source map, nothing outside the allowlist, the three bundles and their notices present, and the editor's
// version. `vsce`'s own failures end the process they run in (`process.exit(1)`, never a throw), so a caller runs this
// script as a child process and reads its exit code.
import { createVSIX } from '@vscode/vsce';
import { existsSync, mkdirSync, readFileSync, realpathSync, statSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';

const here = join(dirname(fileURLToPath(import.meta.url)), '..');

/** The VSIX's limit: half as much again as the 6.3 MB of section 4.7's trial. A source map, or all of Monaco's languages, would cross it. */
export const BUDGET = 10 * 1024 * 1024;

/** Every name the VSIX may hold. */
const ALLOWED = [
  /^extension\.vsixmanifest$/,
  /^\[Content_Types\]\.xml$/,
  /^extension\/package\.json$/,
  /^extension\/readme\.md$/,
  /^extension\/LICENSE\.txt$/,
  /^extension\/(dist|app|companion|l10n)\/./,
  /^extension\/package\.nls(\.[a-z]{2,3}(-[a-z]+)?)?\.json$/,
  /^extension\/THIRD-PARTY-NOTICES\.txt$/,
  /^extension\/media\/icon\.png$/,
];

/** The names the VSIX must hold: each bundle, and each bundle's notices. */
const REQUIRED = [
  'extension/dist/extension.js',
  'extension/THIRD-PARTY-NOTICES.txt',
  'extension/app/index.html',
  'extension/app/THIRD-PARTY-NOTICES.txt',
  'extension/companion/cli.mjs',
  'extension/companion/relay-worker.mjs',
  'extension/companion/THIRD-PARTY-NOTICES.txt',
];

/** The editor's version, which the extension's follows (the packages move in lockstep). */
export function editorVersion() {
  return JSON.parse(readFileSync(join(here, '..', 'editor', 'package.json'), 'utf8')).version;
}

/** The VSIX's entry names, and its manifest's text, read without unpacking the rest. */
export function contentsOf(path) {
  const names = [];
  const files = unzipSync(new Uint8Array(readFileSync(path)), {
    filter: (file) => { names.push(file.name); return file.name === 'extension/package.json'; },
  });
  const manifest = files['extension/package.json'];
  return { names, manifest: manifest ? strFromU8(manifest) : undefined };
}

/** What is wrong with the VSIX at `path`, as sentences; empty when nothing is. */
export function problemsOf(path, version = editorVersion()) {
  const out = [];
  const size = statSync(path).size;
  if (size > BUDGET) out.push(`it is ${mb(size)}, over the budget of ${mb(BUDGET)}`);
  const { names, manifest } = contentsOf(path);
  const maps = names.filter((n) => n.endsWith('.map'));
  if (maps.length > 0) out.push(`it holds source maps: ${list(maps)}`);
  const outside = names.filter((n) => !ALLOWED.some((a) => a.test(n)) && !n.endsWith('.map'));
  if (outside.length > 0) out.push(`it holds what the allowlist leaves out: ${list(outside)}`);
  const missing = REQUIRED.filter((n) => !names.includes(n));
  if (missing.length > 0) out.push(`it lacks ${list(missing)}`);
  const packaged = manifest ? JSON.parse(manifest).version : undefined;
  if (packaged !== version) out.push(`its version is ${packaged ?? 'missing'}, where the editor's is ${version}`);
  return out;
}

const mb = (bytes) => `${(bytes / 1024 / 1024).toFixed(2)} MB`;
const list = (names) => (names.length > 5 ? `${names.slice(0, 5).join(', ')} and ${names.length - 5} more` : names.join(', '));

function check(path) {
  const problems = problemsOf(path);
  if (problems.length > 0) {
    console.error(`package: ${path} is refused:\n${problems.map((p) => `- ${p}`).join('\n')}`);
    process.exit(1);
  }
  console.error(`package: ${path}: ${mb(statSync(path).size)}, ${contentsOf(path).names.length} files, checked`);
}

function argument(name) {
  const at = process.argv.indexOf(name);
  return at >= 0 ? process.argv[at + 1] : undefined;
}

/** Run as a script, not imported: Node takes the entry module's real path, and `argv[1]` as typed (a symlinked checkout). */
function isMain() {
  try { return realpathSync(process.argv[1]) === realpathSync(fileURLToPath(import.meta.url)); } catch { return false; }
}

if (isMain()) {
  const only = argument('--check');
  if (only) {
    check(resolve(only));
  } else {
    const from = resolve(argument('--from') ?? here);
    const out = resolve(argument('--out') ?? from);
    const absent = ['dist/extension.js', 'app/index.html', 'companion/cli.mjs'].filter((f) => !existsSync(join(from, f)));
    if (absent.length > 0) {
      console.error(`package: ${from} lacks ${absent.join(', ')}: run \`npm run build\` at the repository's root first`);
      process.exit(1);
    }
    const manifest = JSON.parse(readFileSync(join(from, 'package.json'), 'utf8'));
    mkdirSync(out, { recursive: true });
    const vsix = join(out, `${manifest.name}-${manifest.version}.vsix`);
    // No `node_modules`, and no `npm list`: the extension has no runtime dependency (section 4.3).
    await createVSIX({ cwd: from, packagePath: vsix, dependencies: false });
    check(vsix);
  }
}
