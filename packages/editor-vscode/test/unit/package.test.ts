// SPDX-License-Identifier: GPL-3.0-or-later
//
// The VSIX's script (ED-014 section 4.7, P47), over small extension folders made here: the manifest and words the
// extension's own, its bundles stand-ins. Each is packaged by `scripts/package.mjs` run as a child process, since `vsce`'s
// failures end the process they run in, and its exit code read: a good folder packaged and checked; a source map, a path
// outside the allowlist, a missing notice and another version refused; a folder not built refused before `vsce` runs; and a
// VSIX over the budget refused by `--check`.
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { spawnSync } from 'node:child_process';
import { copyFileSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { randomBytes } from 'node:crypto';
import { zipSync, unzipSync } from 'fflate';
// @ts-expect-error: a plain script, no declarations
import { BUDGET, contentsOf, editorVersion } from '../../scripts/package.mjs';

const root = join(import.meta.dirname, '..', '..');
const script = join(root, 'scripts', 'package.mjs');
let scratch: string;

beforeAll(() => { scratch = mkdtempSync(join(tmpdir(), 'docx4j-vsix-')); });
afterAll(() => { rmSync(scratch, { recursive: true, force: true }); });

/** An extension folder as the build leaves one, with `change` applied to its manifest and files. */
function folder(name: string, change: { manifest?: (m: Record<string, unknown>) => void; files?: Record<string, string | null> } = {}): string {
  const dir = join(scratch, name);
  const manifest = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')) as Record<string, unknown>;
  change.manifest?.(manifest);
  const files: Record<string, string | null> = {
    'package.json': JSON.stringify(manifest, null, 2),
    'README.md': '# A stand-in\n\nFor the script\'s tests.\n',
    'LICENSE': 'GPL-3.0-or-later, the stand-in\'s.\n',
    'THIRD-PARTY-NOTICES.txt': 'The extension\'s notices.\n',
    'dist/extension.js': 'exports.activate = () => {};\n',
    'app/index.html': '<!doctype html><title>stand-in</title>\n',
    'app/THIRD-PARTY-NOTICES.txt': 'The application\'s notices.\n',
    'companion/cli.mjs': 'export {};\n',
    'companion/relay-worker.mjs': 'export {};\n',
    'companion/THIRD-PARTY-NOTICES.txt': 'The companion\'s notices.\n',
    ...change.files,
  };
  for (const [path, text] of Object.entries(files)) {
    if (text === null) continue;
    mkdirSync(dirname(join(dir, path)), { recursive: true });
    writeFileSync(join(dir, path), text);
  }
  for (const copied of ['package.nls.json', 'l10n/bundle.l10n.json', 'media/icon.png']) {
    mkdirSync(dirname(join(dir, copied)), { recursive: true });
    copyFileSync(join(root, copied), join(dir, copied));
  }
  return dir;
}

function run(...args: string[]): { status: number | null; stderr: string } {
  const r = spawnSync(process.execPath, [script, ...args], { encoding: 'utf8', timeout: 120_000 });
  return { status: r.status, stderr: `${r.stderr}${r.stdout}` };
}

const vsixOf = (dir: string): string => join(dir, `docx4j-editor-${editorVersion()}.vsix`);

describe('the VSIX (P47)', () => {
  test('a built folder: packaged and checked, its files in the VSIX\'s own names', () => {
    const dir = folder('good');
    const r = run('--from', dir);
    expect(r.status, r.stderr).toBe(0);
    expect(r.stderr).toMatch(/files, checked/);
    const { names, manifest } = contentsOf(vsixOf(dir)) as { names: string[]; manifest: string };
    expect(names).toContain('extension/readme.md');
    expect(names).toContain('extension/LICENSE.txt');
    expect(names).toContain('extension/media/icon.png');
    expect(names).toContain('extension.vsixmanifest');
    expect(names.some((n) => n.startsWith('extension/node_modules/'))).toBe(false);
    expect(JSON.parse(manifest).version).toBe(editorVersion());
  });

  test('--out: the VSIX written in another folder', () => {
    const dir = folder('out');
    const out = join(scratch, 'elsewhere');
    expect(run('--from', dir, '--out', out).status).toBe(0);
    expect(contentsOf(vsixOf(out)).names).toContain('extension/dist/extension.js');
  });

  test('a source map in dist/: refused', () => {
    const r = run('--from', folder('map', { files: { 'dist/extension.js.map': '{"version":3}' } }));
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('it holds source maps: extension/dist/extension.js.map');
  });

  test('a path the allowlist leaves out, put in by the manifest\'s files: refused', () => {
    const r = run('--from', folder('outside', {
      manifest: (m) => { (m['files'] as string[]).push('src/'); },
      files: { 'src/extension.ts': 'export {};\n' },
    }));
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('what the allowlist leaves out: extension/src/extension.ts');
  });

  test('the companion\'s notices missing: refused by the check; the extension\'s own missing: refused by vsce', () => {
    const companion = run('--from', folder('no-companion-notices', { files: { 'companion/THIRD-PARTY-NOTICES.txt': null } }));
    expect(companion.status).toBe(1);
    expect(companion.stderr).toContain('it lacks extension/companion/THIRD-PARTY-NOTICES.txt');
    // `files` names it, and an entry that matches no file ends vsce's process.
    const own = run('--from', folder('no-notices', { files: { 'THIRD-PARTY-NOTICES.txt': null } }));
    expect(own.status).toBe(1);
    expect(own.stderr).toContain('THIRD-PARTY-NOTICES.txt');
  });

  test('a version other than the editor\'s: refused', () => {
    const r = run('--from', folder('version', { manifest: (m) => { m['version'] = '9.9.9'; } }));
    expect(r.status).toBe(1);
    expect(r.stderr).toContain(`its version is 9.9.9, where the editor's is ${editorVersion()}`);
  });

  test('a folder not built: refused before vsce runs, naming the build', () => {
    const r = run('--from', folder('unbuilt', { files: { 'app/index.html': null } }));
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('lacks app/index.html: run `npm run build`');
  });

  test('--check: a VSIX over the budget refused', () => {
    const dir = folder('budget');
    expect(run('--from', dir).status).toBe(0);
    const entries = unzipSync(new Uint8Array(readFileSync(vsixOf(dir))));
    entries['extension/app/assets/large.bin'] = new Uint8Array(randomBytes(BUDGET + 1024));
    const large = join(scratch, 'large.vsix');
    writeFileSync(large, zipSync(entries, { level: 0 }));
    const r = run('--check', large);
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/over the budget of 10\.00 MB/);
    expect(run('--check', vsixOf(dir)).status).toBe(0);
  });
});
