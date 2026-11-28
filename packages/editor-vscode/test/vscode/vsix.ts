// SPDX-License-Identifier: GPL-3.0-or-later
//
// The VSIXs an installed window runs (ED-014 section 4.7, P51), made once per run into `test-results/vscode/vsix/`:
// - the extension's, by its own `scripts/package.mjs` over the build the suite runs after, its checks included;
// - the suite's driver, and the stand-in French language pack (`language-pack/`, an empty translation of VS Code, which
//   VS Code takes, section 9.25 item 3), by `vsce`'s command line. Neither has a repository field or a licence file, which
//   `vsce` would otherwise ask about.
// Each runs as a child process: `vsce`'s failures end the process they run in, which would end the suite's worker with them.
import { spawnSync } from 'node:child_process';
import { mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { DRIVER, EXTENSION } from './harness.js';

const OUT = join(EXTENSION, 'test-results', 'vscode', 'vsix');
export const LANGUAGE_PACK = join(__dirname, 'language-pack');

function run(args: string[], cwd: string): void {
  const r = spawnSync(process.execPath, args, { cwd, encoding: 'utf8', timeout: 300_000 });
  if (r.status !== 0) throw new Error(`${args.join(' ')} failed (${r.status ?? r.signal}):\n${r.stderr}${r.stdout}`);
}

const VSCE = join(dirname(require.resolve('@vscode/vsce/package.json')), 'vsce');

let made: { extension: string; driver: string; pack: string } | undefined;

/** The three VSIXs, packaged at the first call. */
export function vsixes(): { extension: string; driver: string; pack: string } {
  if (made) return made;
  mkdirSync(OUT, { recursive: true });
  run([join(EXTENSION, 'scripts', 'package.mjs'), '--out', OUT], EXTENSION);
  const version = (require(join(EXTENSION, 'package.json')) as { name: string; version: string }).version;
  const extension = join(OUT, `docx4j-editor-${version}.vsix`);
  const driver = join(OUT, 'driver.vsix');
  const pack = join(OUT, 'language-pack-fr.vsix');
  for (const [folder, file] of [[DRIVER, driver], [LANGUAGE_PACK, pack]] as const) {
    run([VSCE, 'package', '--no-dependencies', '--allow-missing-repository', '--skip-license', '--out', file], folder);
  }
  made = { extension, driver, pack };
  return made;
}
