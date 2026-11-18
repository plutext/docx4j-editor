#!/usr/bin/env node
// SPDX-License-Identifier: GPL-3.0-or-later
//
// The licence check ED-001 section 11 item 12 asks for in CI: every package in the production
// dependency tree of every workspace must carry a permissive licence. Permissive only, even
// though this repository is GPL: decisions 15 and 18 (the commercial licence and commercially
// licensed plugins) depend on Plutext holding the rights to everything it ships, and one
// copyleft dependency owned by someone else would take that away.
//
// Reads node_modules directly (no dependency of its own, so nothing to check the checker with),
// walks the `dependencies` of each workspace package transitively, and fails on any licence
// outside the allowlist, on a missing licence field, and on an expression it cannot read.
import { readFileSync, readdirSync, existsSync, writeFileSync } from 'node:fs';
import { join, dirname, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { allowed, licenceOf, describePackage, noticesText, withoutLicenceText, NOTICES_FILE } from './notices.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

function readJson(path) {
  return JSON.parse(readFileSync(path, 'utf8'));
}

/** Resolve a dependency the way Node does: nearest node_modules up the tree. */
function resolvePackage(name, from) {
  let dir = from;
  for (;;) {
    const candidate = join(dir, 'node_modules', name, 'package.json');
    if (existsSync(candidate)) return candidate;
    const parent = dirname(dir);
    if (parent === dir) return undefined;
    dir = parent;
  }
}

const root = readJson(join(ROOT, 'package.json'));
const workspaces = (root.workspaces ?? []).flatMap((pattern) => {
  if (!pattern.endsWith('/*')) return [join(ROOT, pattern)];
  const base = join(ROOT, pattern.slice(0, -2));
  return readdirSync(base).map((d) => join(base, d)).filter((d) => existsSync(join(d, 'package.json')));
});

const seen = new Map(); // name@version -> { licence, ok, path }
const problems = [];
// Our own packages depend on each other; they are the GPL side of the boundary, not third parties.
const ours = new Set(workspaces.map((ws) => readJson(join(ws, 'package.json')).name));

function visit(pkgJsonPath, chain) {
  const pkg = readJson(pkgJsonPath);
  const key = `${pkg.name}@${pkg.version}`;
  if (seen.has(key)) return;
  if (ours.has(pkg.name)) {
    if (pkg.license !== 'GPL-3.0-or-later') problems.push(`${key}: our package must be GPL-3.0-or-later, is ${pkg.license}`);
    seen.set(key, { licence: pkg.license, ok: true });
    for (const dep of Object.keys(pkg.dependencies ?? {})) {
      const resolved = resolvePackage(dep, dirname(pkgJsonPath));
      if (resolved) visit(resolved, [...chain, pkg.name]);
    }
    return;
  }
  const licence = licenceOf(pkg);
  const ok = allowed(licence);
  seen.set(key, { licence, ok });
  if (!ok) problems.push(`${key}: ${licence ?? 'no licence field'}  (via ${chain.join(' > ')})`);
  const deps = { ...(pkg.dependencies ?? {}), ...(pkg.peerDependencies ?? {}) };
  const optionalPeers = new Set(Object.entries(pkg.peerDependenciesMeta ?? {}).filter(([, m]) => m?.optional).map(([n]) => n));
  for (const dep of Object.keys(deps)) {
    const resolved = resolvePackage(dep, dirname(pkgJsonPath));
    if (!resolved) {
      if (optionalPeers.has(dep) || pkg.optionalDependencies?.[dep]) continue;
      problems.push(`${key}: dependency ${dep} is not installed`);
      continue;
    }
    visit(resolved, [...chain, pkg.name]);
  }
}

for (const ws of workspaces) {
  const pkg = readJson(join(ws, 'package.json'));
  if (pkg.license !== 'GPL-3.0-or-later') problems.push(`${pkg.name}: package.json license must be GPL-3.0-or-later, is ${pkg.license}`);
  for (const dep of Object.keys(pkg.dependencies ?? {})) {
    const resolved = resolvePackage(dep, ws);
    if (!resolved) { problems.push(`${pkg.name}: dependency ${dep} is not installed`); continue; }
    visit(resolved, [pkg.name]);
  }
}

const counted = [...seen.values()];
console.log(`checked ${counted.length} packages in the production dependency trees of ${workspaces.length} workspace(s)`);
if (problems.length > 0) {
  console.error('licence check FAILED:');
  for (const p of problems) console.error('  ' + p);
  process.exit(1);
}
console.log('licence check OK: every dependency is permissively licensed');

// Each npm package's THIRD-PARTY-NOTICES.txt (in its `files`, git-ignored, written here and by its
// `prepack`): the third-party packages of its production tree, through our own packages' trees. The
// site's (the private @docx4j/editor) is written by its Vite build from what the bundle holds.
function thirdParties(ws) {
  const found = new Map(); // package folder -> description
  const walked = new Set();
  const walk = (pkgJsonPath, isWorkspace) => {
    if (walked.has(pkgJsonPath)) return;
    walked.add(pkgJsonPath);
    const pkg = readJson(pkgJsonPath);
    const dir = dirname(pkgJsonPath);
    if (!isWorkspace && !ours.has(pkg.name)) found.set(dir, describePackage(dir));
    const deps = { ...(pkg.dependencies ?? {}), ...(isWorkspace || ours.has(pkg.name) ? {} : pkg.peerDependencies ?? {}) };
    for (const dep of Object.keys(deps)) {
      const resolved = resolvePackage(dep, dir);
      if (resolved) walk(resolved, false);
    }
  };
  walk(join(ws, 'package.json'), true);
  return [...found.values()];
}

for (const ws of workspaces) {
  const pkg = readJson(join(ws, 'package.json'));
  if (pkg.private) continue;
  const parties = thirdParties(ws);
  writeFileSync(join(ws, NOTICES_FILE), noticesText(pkg.name, parties));
  console.log(`wrote ${relative(ROOT, join(ws, NOTICES_FILE))}`);
  // Not a failure: the licence is known and permissive. But an MIT or BSD notice names its copyright
  // holder, so find the project's text and add it to scripts/licences/ (notices.mjs, SUPPLIED).
  for (const p of withoutLicenceText(parties)) console.warn(`  ${p.name}@${p.version} ships no licence file: supply its project's in scripts/licences/`);
}
