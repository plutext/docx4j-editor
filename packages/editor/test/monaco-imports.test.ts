// SPDX-License-Identifier: GPL-3.0-or-later
//
// Monaco trimmed (ED-014 section 4.2, P14): `app/monaco.ts` is the application's one importer of Monaco, with three
// languages of the package's 81, and `app/monaco-env.ts` the one of its workers. A single `import … from
// 'monaco-editor'` anywhere else would bring every language back into the build unseen, and an importer that skipped
// `app/monaco.ts` could start an editor before its workers are set. A source scan, since the build would not fail.
// A type-only `import type … from 'monaco-editor'` is refused too, though it costs nothing at run time: the types come
// through `app/monaco.ts` as well, so there is one way in (ED-014 section 9.7 note 5).
import { describe, expect, test } from 'vitest';
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const SRC = join(dirname(fileURLToPath(import.meta.url)), '..', 'src');

function sources(dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) return sources(path);
    return /\.(?:ts|tsx|mts|js|mjs)$/.test(entry.name) ? [path] : [];
  });
}

/** Every module specifier a source names: static imports and exports, side-effect imports, `import()`. A comment on lines of its own is
 * not read; a comment's opener inside a string (a glob in `import.meta.glob`) is left alone. */
function specifiersOf(source: string): string[] {
  const text = source.replace(/^\s*\/\*[\s\S]*?\*\//gm, '').replace(/^\s*\/\/.*$/gm, '');
  const found: string[] = [];
  const patterns = [
    /\b(?:import|export)\s[^'";]*?\bfrom\s*(['"])([^'"]+)\1/g,
    /\bimport\s*(['"])([^'"]+)\1/g,
    /\bimport\s*\(\s*(['"`])([^'"`]+)\1\s*\)/g,
  ];
  for (const pattern of patterns) for (const match of text.matchAll(pattern)) found.push(match[2]!);
  return found;
}

const files = sources(SRC).map((path) => ({ name: relative(SRC, path).split(sep).join('/'), text: readFileSync(path, 'utf8') }));

describe('Monaco through app/monaco.ts alone', () => {
  test('the scan reads every form of import', () => {
    expect(specifiersOf([
      "import * as monaco from 'monaco-editor';",
      "import type { editor } from \"monaco-editor/esm/vs/editor/editor.api.js\";",
      "import 'monaco-editor/esm/vs/basic-languages/xml/xml.contribution.js';",
      "export * from 'monaco-editor/esm/vs/editor/editor.api.js';",
      "const m = await import('monaco-editor');",
      "import {\n  a,\n  b,\n} from 'monaco-editor';",
      "// A single `import … from 'monaco-editor'` in a comment",
      "/** or `import 'monaco-editor'` in a doc comment */",
    ].join('\n')).filter((s) => s.startsWith('monaco-editor'))).toHaveLength(6);
  });

  test('no source but app/monaco.ts and app/monaco-env.ts imports the package', () => {
    const importers = files.filter((f) => specifiersOf(f.text).some((s) => s === 'monaco-editor' || s.startsWith('monaco-editor/'))).map((f) => f.name);
    expect(importers.sort()).toEqual(['app/monaco-env.ts', 'app/monaco.ts']);
  });

  test('app/monaco.ts takes the editor and three languages, never the whole package', () => {
    const monaco = files.find((f) => f.name === 'app/monaco.ts')!;
    expect(specifiersOf(monaco.text).filter((s) => s.startsWith('monaco-editor')).sort()).toEqual([
      'monaco-editor/esm/vs/basic-languages/javascript/javascript.contribution.js',
      'monaco-editor/esm/vs/basic-languages/typescript/typescript.contribution.js',
      'monaco-editor/esm/vs/basic-languages/xml/xml.contribution.js',
      'monaco-editor/esm/vs/editor/edcore.main.js',
      'monaco-editor/esm/vs/editor/editor.api.js',
      'monaco-editor/esm/vs/language/typescript/monaco.contribution.js',
    ]);
  });

  test('the workers are set by app/monaco.ts, which nothing goes round', () => {
    const importers = files.filter((f) => specifiersOf(f.text).some((s) => /(?:^|\/)monaco-env(?:\.js)?$/.test(s))).map((f) => f.name);
    expect(importers).toEqual(['app/monaco.ts']);
  });
});
