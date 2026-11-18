// SPDX-License-Identifier: GPL-3.0-or-later
//
// The third-party notices that go with every copy the repository distributes: the site (written by
// the editor's Vite build from the packages actually bundled), and each npm package (written by
// check-licences.mjs from the package's production dependency tree). MIT and BSD ask that their
// copyright and permission notices accompany copies, and Apache-2.0 its NOTICE files, so the file
// holds each package's name, version and licence, and the text of every licence or notice file the
// package ships.
//
// Also the licence allowlist, which the licence check and the site's build share: permissive only,
// even though this repository is GPL (see check-licences.mjs).
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { dirname, join, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

export const NOTICES_FILE = 'THIRD-PARTY-NOTICES.txt';

export const ALLOWED = new Set([
  'MIT', 'ISC', 'BSD-2-Clause', 'BSD-3-Clause', 'Apache-2.0', '0BSD', 'CC0-1.0', 'Unlicense',
  'BlueOak-1.0.0', 'Python-2.0', 'MIT-0', 'CC-BY-4.0',
]);

/** SPDX expression → true when any alternative of an OR is allowed, or every part of an AND; keep it simple. */
export function allowed(expression) {
  if (typeof expression !== 'string') return false;
  const e = expression.replace(/^\(|\)$/g, '').trim();
  if (e.includes(' OR ')) return e.split(' OR ').some((part) => allowed(part.trim()));
  if (e.includes(' AND ')) return e.split(' AND ').every((part) => allowed(part.trim()));
  return ALLOWED.has(e);
}

/** A package.json's licence as an SPDX expression, whichever of the old forms it uses. */
export function licenceOf(pkg) {
  if (typeof pkg.license === 'string') return pkg.license;
  if (pkg.license?.type) return pkg.license.type;
  if (Array.isArray(pkg.licenses)) return pkg.licenses.map((l) => l.type).join(' OR ');
  return undefined;
}

/**
 * The folder of the package a bundled module came from, or undefined for one of the repository's
 * own (its workspaces resolve to `packages/`, never through node_modules) or a virtual module.
 */
export function packageDirOf(moduleId) {
  const path = moduleId.replace(/[?#].*$/, '').split(sep).join('/');
  if (path.startsWith('\0')) return undefined;
  const at = path.lastIndexOf('/node_modules/');
  if (at < 0) return undefined;
  const rest = path.slice(at + '/node_modules/'.length).split('/');
  const name = rest[0].startsWith('@') ? `${rest[0]}/${rest[1]}` : rest[0];
  return `${path.slice(0, at)}/node_modules/${name}`;
}

/**
 * Licence texts for packages that ship none, taken from the project's repository (scripts/licences/):
 * the licence's copyright notice must accompany copies all the same.
 */
const SUPPLIED = [
  { match: (name) => name.startsWith('@hocuspocus/'), file: 'hocuspocus.md', from: 'https://github.com/ueberdosis/hocuspocus/blob/v3.4.4/LICENSE.md' },
  // The repository's root LICENSE is Apache-2.0, its libraries' MIT, as the package declares.
  { match: (name) => name === 'standardwebhooks', file: 'standardwebhooks.md', from: 'https://github.com/standard-webhooks/standard-webhooks/blob/main/libraries/LICENSE' },
];
const LICENCES = join(dirname(fileURLToPath(import.meta.url)), 'licences');

/** The packages among `packages` that ship no licence file and have none supplied. */
export function withoutLicenceText(packages) {
  return packages.filter((p) => p && noticeFiles(p.dir).length === 0 && !SUPPLIED.some((s) => s.match(p.name)));
}

const NOTICE_FILE = /^(licen[cs]e|copying|notice|copyright)([-._].*)?$/i;

/** The licence and notice files a package ships, licence first. */
function noticeFiles(dir) {
  let names;
  try { names = readdirSync(dir, { withFileTypes: true }); } catch { return []; }
  return names
    .filter((e) => e.isFile() && NOTICE_FILE.test(e.name))
    .map((e) => e.name)
    .sort((a, b) => (/^notice/i.test(a) === /^notice/i.test(b) ? (a < b ? -1 : 1) : /^notice/i.test(a) ? 1 : -1));
}

/** One entry per package folder: name, version, licence, and where it is. */
export function describePackage(dir) {
  const file = join(dir, 'package.json');
  if (!existsSync(file)) return undefined;
  const pkg = JSON.parse(readFileSync(file, 'utf8'));
  const repository = typeof pkg.repository === 'string' ? pkg.repository : pkg.repository?.url;
  return { name: pkg.name, version: pkg.version, licence: licenceOf(pkg), homepage: pkg.homepage ?? repository, author: typeof pkg.author === 'string' ? pkg.author : pkg.author?.name, dir };
}

const RULE = '='.repeat(80);
const THIN = '-'.repeat(80);

/**
 * THIRD-PARTY-NOTICES.txt for `what` (the site, or a package), over the packages given (from
 * describePackage), each once by name and version; `licenceFile` is the GPL's file beside it.
 */
export function noticesText(what, packages, licenceFile = 'LICENSE') {
  const unique = new Map();
  for (const p of packages) if (p) unique.set(`${p.name}@${p.version}`, p);
  const sorted = [...unique.values()].sort((a, b) => (a.name === b.name ? (a.version < b.version ? -1 : 1) : a.name < b.name ? -1 : 1));
  const out = [
    `Third-party notices for ${what}`,
    '',
    'This is part of the docx4j editor, free software under the GNU General Public License, version 3',
    `or later (see ${licenceFile}). It includes or depends on the ${sorted.length} third-party packages below, each`,
    'under its own licence. Their licence and notice files follow, as each package ships them.',
    '',
  ];
  for (const p of sorted) {
    out.push(RULE, `${p.name} ${p.version}`, `Licence: ${p.licence ?? 'not stated'}`);
    if (p.homepage) out.push(`From: ${p.homepage}`);
    out.push(THIN);
    const files = noticeFiles(p.dir);
    const supplied = files.length === 0 ? SUPPLIED.find((s) => s.match(p.name)) : undefined;
    if (supplied) {
      out.push(`The package ships no licence file; this is its project's, from ${supplied.from}:`, '');
      out.push(readFileSync(join(LICENCES, supplied.file), 'utf8').replace(/\r\n?/g, '\n').trimEnd(), '');
      continue;
    }
    if (files.length === 0) {
      out.push(`The package ships no licence file; its package.json declares ${p.licence ?? 'no licence'}${p.author ? `, its author ${p.author}` : ''}.`, '');
      continue;
    }
    for (const [i, name] of files.entries()) {
      if (i > 0) out.push('', `(${name})`, '');
      out.push(readFileSync(join(p.dir, name), 'utf8').replace(/\r\n?/g, '\n').trimEnd());
    }
    out.push('');
  }
  return `${out.join('\n')}\n`;
}
