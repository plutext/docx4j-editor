#!/usr/bin/env node
// SPDX-License-Identifier: GPL-3.0-or-later
//
// Copies the text extraction conformance suite into packages/editor-model/test/fixtures/text-extraction/
// (ED-008 step 2), and writes its manifest:
//
//   node scripts/copy-text-extraction-suite.mjs <docx4j checkout> <docx4j commit> <core-ts checkout> <core-ts tag>
//
// From docx4j at the commit (git archive, never the working tree), unedited: the probe's directory
// (docx4j-core-tests/src/test/resources/text-extraction/phase0/: the fixtures, the tables, the goldens) and the
// suite's three contracts beside it (README.md, GOLDENS.md, INSPECTION.md), and the licence of the fonts
// extraction-probe-inspect.docx embeds (support/inspect-fonts/OFL.txt and its README.md: the OFL asks that
// the notice go with every copy). Use the commit the installed
// core-ts release's own copy is at (its test/text-extraction/manifest.json), so that the editor holds the
// engine to the suite the engine was released against.
//
// From core-ts at the tag, unedited: the runner (test/text-extraction/conformance.mjs, its mirror of docx4j's
// ExtractionConformance), which the model's test runs over the installed engine through its own adapter.
//
// manifest.json records both commits and, for each file, its source path and SHA-1; the model's test checks
// every file against it. A recopy replaces the directory whole: review the difference before committing it.
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const repo = dirname(dirname(fileURLToPath(import.meta.url)));
const [docx4j, commit, coreTs, tag] = process.argv.slice(2);
if (!docx4j || !commit || !coreTs || !tag) {
  console.error('usage: node scripts/copy-text-extraction-suite.mjs <docx4j checkout> <docx4j commit> <core-ts checkout> <core-ts tag>');
  process.exit(2);
}

const SUITE = 'docx4j-core-tests/src/test/resources/text-extraction';
const RUNNER = 'test/text-extraction/conformance.mjs';
const target = join(repo, 'packages', 'editor-model', 'test', 'fixtures', 'text-extraction');

const git = (dir, ...args) => execFileSync('git', ['-C', dir, ...args], { maxBuffer: 1 << 30 });
const describe = (dir, ref) => {
  const full = git(dir, 'rev-parse', '--verify', `${ref}^{commit}`).toString().trim();
  return { commit: full, date: git(dir, 'show', '-s', '--format=%cs', full).toString().trim(), subject: git(dir, 'show', '-s', '--format=%s', full).toString().trim() };
};
const source = describe(docx4j, commit);
const engine = { tag, ...describe(coreTs, tag) };

// The folder holds the copy alone: the runner's types, the editor's own, are in test/text-extraction-suite.mts.
rmSync(target, { recursive: true, force: true });
mkdirSync(join(target, 'suite'), { recursive: true });
// The archive's paths are the repository's: strip the suite's prefix, keeping phase0/, the three contracts and the fonts' licence.
const paths = [`${SUITE}/phase0`, `${SUITE}/README.md`, `${SUITE}/GOLDENS.md`, `${SUITE}/INSPECTION.md`, `${SUITE}/support/inspect-fonts/OFL.txt`, `${SUITE}/support/inspect-fonts/README.md`];
const archive = git(docx4j, 'archive', '--format=tar', source.commit, ...paths);
execFileSync('tar', ['-x', '-C', join(target, 'suite'), `--strip-components=${SUITE.split('/').length}`], { input: archive });
writeFileSync(join(target, 'conformance.mjs'), git(coreTs, 'show', `${engine.commit}:${RUNNER}`));

const sha1 = (file) => createHash('sha1').update(readFileSync(file)).digest('hex');
const files = [];
const walk = (dir) => {
  for (const e of readdirSync(dir, { withFileTypes: true }).sort((a, b) => (a.name < b.name ? -1 : 1))) {
    const path = join(dir, e.name);
    if (e.isDirectory()) { walk(path); continue; }
    const here = relative(target, path).split(sep).join('/');
    const from = here === 'conformance.mjs' ? { 'core-ts': RUNNER } : { docx4j: `${SUITE}/${here.slice('suite/'.length)}` };
    files.push({ ...from, sha1: sha1(path), path: here });
  }
};
walk(target);
writeFileSync(join(target, 'manifest.json'), `${JSON.stringify({ docx4j: source, 'core-ts': engine, files }, null, 1)}\n`);
console.log(`${files.length} files from docx4j ${source.commit.slice(0, 9)} and core-ts ${tag} into ${relative(repo, target)}`);
