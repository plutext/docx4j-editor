// SPDX-License-Identifier: GPL-3.0-or-later
//
// The engine's declarations for the console's completion (ED-002 section 6.3): the engine's
// `dist/**/*.d.mts`, and the object model's index, builders, helpers and the `wml` module (the
// other 93 modules are the size of the editor and stay out), loaded as raw strings on first use
// into Monaco's extra libraries, with a synthetic package.json per package so bare imports
// resolve, and a globals file declaring the console's bindings.
import * as monaco from 'monaco-editor';

const engine = import.meta.glob('../../../../node_modules/@docx4j/core-ts/dist/**/*.d.mts', { query: '?raw', import: 'default' }) as Record<string, () => Promise<string>>;
const objects = import.meta.glob([
  '../../../../node_modules/@docx4j/generated-objects-ts/dist/index.d.mts',
  '../../../../node_modules/@docx4j/generated-objects-ts/dist/builders/wml.d.mts',
  '../../../../node_modules/@docx4j/generated-objects-ts/dist/helpers/wml.d.mts',
  '../../../../node_modules/@docx4j/generated-objects-ts/modules/org_docx4j_wml.d.ts',
  '../../../../node_modules/@docx4j/generated-objects-ts/modules/org_docx4j_wml.el.d.mts',
], { query: '?raw', import: 'default' }) as Record<string, () => Promise<string>>;

const GLOBALS = `
import type { WordprocessingMLPackage, Body, MainDocumentPart, Paragraph, Range, Font } from '@docx4j/core-ts';
import type * as officeJs from '@docx4j/core-ts/office-js';
import type * as wmlBuilders from '@docx4j/generated-objects-ts/builders/wml';
import type * as elFactories from '@docx4j/generated-objects-ts/modules/org_docx4j_wml.el';
declare global {
  /** The package (the document layer). */
  const pkg: WordprocessingMLPackage;
  /** The content API over the tree, in Office JS shapes. */
  const body: Body;
  const main: MainDocumentPart;
  /** Element factories of the wml module: el.p({ ... }), el.r({ ... }). */
  const el: typeof elFactories;
  /** The builders: wml.p('text'), wml.r, wml.t, wml.tbl, and the wml\`...\` template. */
  const wml: typeof wmlBuilders;
  /** The selection: the selected paragraphs, their text and addresses; the range when the selection lies within one paragraph; the font of the range or of every selected run. */
  const selection: { paragraphs: Paragraph[]; text: string; addresses: string[]; range: Range | undefined; font: Font };
  /** The editor's commands and the selection's effective formatting. */
  const editor: { run(name: string, args?: unknown): boolean; formatting(): unknown };
  const perspective: { id: string; panel: { open: boolean; tab: string } };
  /**
   * Office JS's Word namespace over this document (the subset @docx4j/core-ts implements): an
   * add-in's await Word.run(async (context) => { ... }) runs unchanged, one undo step. A member
   * the subset lacks is a type error here and a NotSupportedError at the run.
   */
  const Word: Omit<typeof officeJs.Word, 'run'> & {
    run<T>(batch: (context: officeJs.RequestContext) => Promise<T> | T, options?: officeJs.RunOptions): Promise<T>;
  };
}
export {};
`;

/** The other module declarations of the object model (`a:`, `wp:`, `pic:`, ...), loaded on demand (ED-003 section 7.6). */
const modules = import.meta.glob('../../../../node_modules/@docx4j/generated-objects-ts/modules/*.d.ts', { query: '?raw', import: 'default' }) as Record<string, () => Promise<string>>;

/** How many module declarations "load all types" adds (the `wml` module is loaded from the start). */
export const MODULE_COUNT = Object.keys(modules).filter((p) => !p.endsWith('/org_docx4j_wml.d.ts')).length;

let loaded: Promise<void> | undefined;
let loadedAll: Promise<void> | undefined;

/** Load every module declaration once (after the base set); later calls await the same promise. */
export function loadAllModuleDeclarations(): Promise<void> {
  loadedAll ??= (async () => {
    await loadDeclarations();
    const ts = monaco.languages.typescript.typescriptDefaults;
    for (const [path, load] of Object.entries(modules)) {
      if (path.endsWith('/org_docx4j_wml.d.ts')) continue;
      ts.addExtraLib(await load(), virtualPath(path));
    }
  })();
  return loadedAll;
}

/**
 * The path Monaco's resolver sees for a glob key: `file:///node_modules/...`. A glob that leaves
 * the package root comes back from Vite as an absolute path, not the relative pattern written
 * (ED-003 section 11.15 item 5), so the key is cut at `node_modules/` rather than stripped of
 * leading `../`.
 */
function virtualPath(globPath: string): string {
  const at = globPath.indexOf('node_modules/');
  return 'file:///' + (at >= 0 ? globPath.slice(at) : globPath.replace(/^(\.\.\/)+/, ''));
}

/** Load once; later calls await the same promise. */
export function loadDeclarations(): Promise<void> {
  loaded ??= (async () => {
    const ts = monaco.languages.typescript.typescriptDefaults;
    ts.setCompilerOptions({
      target: monaco.languages.typescript.ScriptTarget.ES2020,
      module: monaco.languages.typescript.ModuleKind.ESNext,
      moduleResolution: monaco.languages.typescript.ModuleResolutionKind.NodeJs,
      allowNonTsExtensions: true,
      // Off: the worker takes any extension it does not know (`.d.mts`, a model with none) for
      // JavaScript while `allowJs` is on, and then no declaration parses and no type error shows
      // (ED-003 section 11.15 item 5).
      allowJs: false,
      strict: false,
      noEmit: true,
      lib: ['es2022', 'dom'],
      typeRoots: ['file:///node_modules/@types'],
    });
    ts.setDiagnosticsOptions({ noSemanticValidation: false, noSyntaxValidation: false, diagnosticCodesToIgnore: [1375, 1378] });
    // Monaco reads `filePath`, not `path`: a lib without it is registered under a random name and resolves nothing (ED-003 section 11.15 item 5).
    const libs: { content: string; filePath: string }[] = [];
    for (const [path, load] of Object.entries(engine)) libs.push({ content: await load(), filePath: virtualPath(path) });
    for (const [path, load] of Object.entries(objects)) {
      const content = await load();
      const p = virtualPath(path);
      libs.push({ content, filePath: p });
      // Bare subpath imports (`.../builders/wml`) resolve without the exports map only to `.d.ts`.
      if (p.endsWith('.d.mts')) libs.push({ content, filePath: p.replace(/\.d\.mts$/, '.d.ts') });
    }
    libs.push({ content: JSON.stringify({ name: '@docx4j/core-ts', types: './dist/index.d.mts' }), filePath: 'file:///node_modules/@docx4j/core-ts/package.json' });
    // The subpaths Monaco's resolver cannot reach through an exports map (ED-003 section 7.3's `Word` global).
    libs.push({ content: "export * from './dist/office-js/index.mjs';\n", filePath: 'file:///node_modules/@docx4j/core-ts/office-js.d.ts' });
    libs.push({ content: JSON.stringify({ name: '@docx4j/generated-objects-ts', types: './dist/index.d.ts' }), filePath: 'file:///node_modules/@docx4j/generated-objects-ts/package.json' });
    libs.push({ content: GLOBALS, filePath: 'file:///globals.d.ts' });
    ts.setExtraLibs(libs);
  })();
  return loaded;
}
