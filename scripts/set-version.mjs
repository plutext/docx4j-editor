// SPDX-License-Identifier: GPL-3.0-or-later
//
// Sets the editor's version in every workspace package, which share one, and the packages' references to one another
// (RELEASING.md). `npm version` sets the versions alone and leaves the references, so the application's exact
// reference to @docx4j/editor-model would then match no workspace and npm would fetch the published package instead.
// Afterwards, `npm install` brings the lockfile along.
//
//   node scripts/set-version.mjs 0.1.1
import { existsSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const version = process.argv[2] ?? '';
if (!/^\d+\.\d+\.\d+$/.test(version)) {
  console.error('usage: node scripts/set-version.mjs <major.minor.patch>');
  process.exit(1);
}
const packages = join(dirname(fileURLToPath(import.meta.url)), '..', 'packages');
const manifests = readdirSync(packages, { withFileTypes: true })
  .filter((d) => d.isDirectory() && existsSync(join(packages, d.name, 'package.json')))
  .map((d) => {
    const file = join(packages, d.name, 'package.json');
    return { file, json: JSON.parse(readFileSync(file, 'utf8')) };
  });
const names = new Set(manifests.map((m) => m.json.name));
for (const { file, json } of manifests) {
  json.version = version;
  for (const field of ['dependencies', 'devDependencies', 'peerDependencies', 'optionalDependencies']) {
    const deps = json[field];
    if (!deps) continue;
    // A caret range stays a caret range, an exact version stays exact.
    for (const name of Object.keys(deps)) if (names.has(name)) deps[name] = deps[name].startsWith('^') ? `^${version}` : version;
  }
  writeFileSync(file, `${JSON.stringify(json, null, 2)}\n`);
  console.log(`${json.name} ${version}`);
}
