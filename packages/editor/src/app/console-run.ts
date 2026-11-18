// SPDX-License-Identifier: GPL-3.0-or-later
//
// Running a console script (ED-001 section 3.5; ED-002 section 6.3): Sucrase strips the
// TypeScript, the code is wrapped as an async function taking the bindings as parameters (so
// top-level await works and nothing is global), imported from a blob URL, and called inside
// `runOnTree`, so the whole run is one Yjs transaction and one undo step. Errors carry the line
// of the user's code: Sucrase keeps line numbers, and the wrapper adds one line. A macro run
// (ED-003 section 7.4) is the same call without the console.
import { transform } from 'sucrase';
import * as builders from '@docx4j/generated-objects-ts/builders/wml';
import * as el from '@docx4j/generated-objects-ts/modules/org_docx4j_wml.el';
import { Word as officeJs, run as officeRun, ItemNotFoundError, type RunOptions, type RequestContext } from '@docx4j/core-ts/office-js';
import { selectionSpec, selectionBinding, type EditorDocument, type SelectionFormatting, type SelectionSpec, type SelectionBinding } from '@docx4j/editor-model';
import type { EditorSurface } from '../view/surface.js';

export const BINDINGS = ['pkg', 'body', 'main', 'el', 'wml', 'selection', 'editor', 'perspective', 'console', 'Word'] as const;

export interface ConsoleOutput {
  log(...args: unknown[]): void;
  warn(...args: unknown[]): void;
  error(...args: unknown[]): void;
}

export interface EditorBinding {
  run(name: string, args?: unknown): boolean;
  formatting(): SelectionFormatting;
  /** The selection as indexes and offsets, for `selection` to be built over the tree. */
  selection(): SelectionSpec;
}

/** The `editor` binding over a surface: its commands, its formatting, its selection. */
export function editorBinding(surface: EditorSurface): EditorBinding {
  return {
    run: (name, args) => surface.run(name, args),
    formatting: () => surface.formatting(),
    selection: () => selectionSpec(surface.state),
  };
}

export interface RunOutcome {
  value?: unknown;
  error?: unknown;
  /** The line in the user's code an error came from, when it can be told. */
  line?: number;
  changed: boolean;
}

/** The user's code as a module whose default export takes the bindings. */
export function wrap(code: string): string {
  return `export default async ({ ${BINDINGS.join(', ')} }) => {\n${code}\n};\n`;
}

/** TypeScript to JavaScript, line for line. */
export function compile(code: string): string {
  return transform(wrap(code), { transforms: ['typescript'], disableESTransforms: true }).code;
}

/** The line of the user's code an error refers to, from a stack trace mentioning the blob URL. */
export function lineOf(error: unknown, blobUrl: string): number | undefined {
  const stack = (error as { stack?: string } | undefined)?.stack ?? '';
  const escaped = blobUrl.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const m = new RegExp(`${escaped}:(\\d+):(\\d+)`).exec(stack);
  if (!m) return undefined;
  const line = Number(m[1]) - 1;   // the wrapper's first line
  return line >= 1 ? line : undefined;
}

/**
 * The console's `Word` (ED-003 section 7.3): the shim's object with `run` bound to the session's
 * package and the selection, so an add-in's `await Word.run(async (context) => { ... })` pastes
 * unchanged and runs inside the console's own tree run. `getSelection()` is the selection's
 * range: within one paragraph as before; across paragraphs one range from the first selected
 * paragraph's offset to the last's, the last mark held when the selection takes it (core-ts 0.6.0,
 * CR-014; ED-008 section 8.8 item 8), so a script reads, deletes and replaces what the person
 * selected; a whole paragraph where the binding gives no range; or nothing selected.
 */
export function wordBinding(session: EditorDocument, selection: SelectionBinding): typeof officeJs {
  const range = selection.selectionRange ?? selection.range ?? selection.paragraphs[0]?.getRange('Whole');
  return Object.freeze({
    ...officeJs,
    run: <T>(batch: (context: RequestContext) => Promise<T> | T, options: RunOptions = {}) => {
      if (!range && options.selection === undefined) {
        // `getSelection()` with nothing selected: the shim reports it at the call, as Office JS's null object would.
        return officeRun(session.pkg, (context) => {
          Object.defineProperty(context.document, 'getSelection', { value: () => { throw new ItemNotFoundError('nothing is selected'); } });
          return batch(context);
        }, options);
      }
      return officeRun(session.pkg, batch, { ...(range ? { selection: range } : {}), ...options });
    },
  }) as unknown as typeof officeJs;
}

/**
 * Run a script against the session: compile, import from a blob URL, call inside `runOnTree`
 * with the bindings. `console` output goes to `output`; the return value comes back.
 */
export async function runScript(session: EditorDocument, editor: EditorBinding, perspective: { id: string; panel: { open: boolean; tab: string } }, code: string, output: ConsoleOutput): Promise<RunOutcome> {
  let js: string;
  try {
    js = compile(code);
  } catch (e) {
    return { error: e, line: (e as { loc?: { line?: number } }).loc?.line !== undefined ? (e as { loc: { line: number } }).loc.line - 1 : undefined, changed: false };
  }
  const url = URL.createObjectURL(new Blob([js], { type: 'text/javascript' }));
  let mod: { default: (bindings: Record<string, unknown>) => Promise<unknown> };
  try {
    mod = (await import(/* @vite-ignore */ url)) as typeof mod;
  } catch (e) {
    URL.revokeObjectURL(url);
    return { error: e, line: lineOf(e, url), changed: false };
  }
  const spec = editor.selection();
  const result = await session.runOnTree(async ({ pkg, body, main }) => {
    const selection = selectionBinding(body, spec);
    return mod.default({ pkg, body, main, el, wml: builders, selection, editor: { run: editor.run.bind(editor), formatting: editor.formatting.bind(editor) }, perspective, console: output, Word: wordBinding(session, selection) });
  });
  URL.revokeObjectURL(url);
  const out: RunOutcome = { changed: result.changed };
  if ('value' in result) out.value = result.value;
  if (result.error !== undefined) { out.error = result.error; const line = lineOf(result.error, url); if (line !== undefined) out.line = line; }
  return out;
}
