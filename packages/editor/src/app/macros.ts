// SPDX-License-Identifier: GPL-3.0-or-later
//
// Macros in the application (ED-001 section 3.5; ED-003 section 7.4): the browser's list in
// local storage, the document's through the model's part, the trust record for document macros
// (ED-001 section 11 item 5, written into the UI: a document macro is code from wherever the
// document came from, shown before its first run in this browser, the script's hash remembered
// so the dialog returns when the script changes), the deployment flag that disables document
// macros, and a run, which is a console run without the console: one undo step.
import { readDocumentMacros, writeDocumentMacros, type EditorDocument, type Macro } from '@docx4j/editor-model';
import type { EditorSurface } from '../view/surface.js';
import { runScript, editorBinding, type ConsoleOutput, type RunOutcome } from './console-run.js';
import { t } from './i18n.js';
import { host } from './host.js';

const BROWSER_KEY = 'docx4j-editor.macros';
const TRUSTED_KEY = 'docx4j-editor.macros.trusted';

export type MacroHome = 'browser' | 'document';

export interface StoredMacro extends Macro {
  home: MacroHome;
}

/** Whether this deployment offers document macros at all (`VITE_DOCUMENT_MACROS=off` at build time turns them off; the part is kept). */
export function documentMacrosBuilt(): boolean {
  const env = (import.meta as unknown as { env?: Record<string, string | undefined> }).env;
  return env?.['VITE_DOCUMENT_MACROS'] !== 'off';
}

/**
 * Whether document macros are offered: the build's flag, and the host's feature (ED-014 P27): VS Code offers them only in a
 * trusted workspace, as its boot data said when the page was written.
 */
export function documentMacrosEnabled(): boolean {
  return documentMacrosBuilt() && host().has('documentMacros');
}

// ---------------------------------------------------------------- the browser's list

export function browserMacros(): Macro[] {
  try {
    const raw = localStorage.getItem(BROWSER_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((m): m is Macro => typeof m === 'object' && m !== null && typeof (m as Macro).name === 'string' && typeof (m as Macro).script === 'string');
  } catch {
    return [];
  }
}

export function saveBrowserMacros(macros: readonly Macro[]): void {
  try { localStorage.setItem(BROWSER_KEY, JSON.stringify(macros)); } catch { /* ignore */ }
}

/** Add or replace by name. */
export function putBrowserMacro(macro: Macro): Macro[] {
  const list = browserMacros().filter((m) => m.name !== macro.name);
  list.push(macro);
  saveBrowserMacros(list);
  return list;
}

export function deleteBrowserMacro(name: string): Macro[] {
  const list = browserMacros().filter((m) => m.name !== name);
  saveBrowserMacros(list);
  return list;
}

// ---------------------------------------------------------------- the document's list

export async function documentMacros(session: EditorDocument): Promise<Macro[]> {
  return documentMacrosEnabled() ? readDocumentMacros(session) : [];
}

export async function putDocumentMacro(session: EditorDocument, macro: Macro): Promise<Macro[]> {
  const list = (await readDocumentMacros(session)).filter((m) => m.name !== macro.name);
  list.push(macro);
  await writeDocumentMacros(session, list);
  return list;
}

export async function deleteDocumentMacro(session: EditorDocument, name: string): Promise<Macro[]> {
  const list = (await readDocumentMacros(session)).filter((m) => m.name !== name);
  await writeDocumentMacros(session, list);
  return list;
}

// ---------------------------------------------------------------- trust

/** SHA-256 of the script, hex. */
export async function scriptHash(script: string): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(script));
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

function trustedHashes(): Set<string> {
  try {
    const raw = localStorage.getItem(TRUSTED_KEY);
    const parsed = raw ? (JSON.parse(raw) as unknown) : [];
    return new Set(Array.isArray(parsed) ? parsed.filter((h): h is string => typeof h === 'string') : []);
  } catch {
    return new Set();
  }
}

/** Whether a document macro's script has been shown and run in this browser before. */
export async function isTrusted(script: string): Promise<boolean> {
  return trustedHashes().has(await scriptHash(script));
}

export async function trust(script: string): Promise<void> {
  const set = trustedHashes();
  set.add(await scriptHash(script));
  try { localStorage.setItem(TRUSTED_KEY, JSON.stringify([...set])); } catch { /* ignore */ }
}

// ---------------------------------------------------------------- running

export interface MacroRun {
  outcome: RunOutcome;
  /** What the script logged, in order. */
  lines: string[];
}

/** Run a macro over a surface: a console run, one undo step; the output collected. */
export async function runMacro(macro: Macro, session: EditorDocument, surface: EditorSurface, perspective: { id: string; panel: { open: boolean; tab: string } }): Promise<MacroRun> {
  const lines: string[] = [];
  const show = (v: unknown): string => (typeof v === 'string' ? v : (() => { try { return JSON.stringify(v, (k, x) => (k === 'PARENT' ? undefined : x)) ?? String(v); } catch { return String(v); } })());
  const output: ConsoleOutput = {
    log: (...a) => lines.push(a.map(show).join(' ')),
    warn: (...a) => lines.push(a.map(show).join(' ')),
    error: (...a) => lines.push(a.map(show).join(' ')),
  };
  const outcome = await runScript(session, editorBinding(surface), perspective, macro.script, output);
  return { outcome, lines };
}

/** One line for the status bar after a run. */
export function describeRun(macro: Macro, run: MacroRun): string {
  const { outcome } = run;
  if (outcome.error !== undefined) {
    const e = outcome.error as Error;
    const error = e.message ?? String(e);
    return outcome.line !== undefined ? t('macro.failedAt', { name: macro.name, error, line: outcome.line }) : t('macro.failed', { name: macro.name, error });
  }
  if (outcome.value === undefined) return t(outcome.changed ? 'macro.ran' : 'macro.ranUnchanged', { name: macro.name });
  const value = typeof outcome.value === 'string' ? outcome.value : JSON.stringify(outcome.value);
  return t(outcome.changed ? 'macro.ranValue' : 'macro.ranValueUnchanged', { name: macro.name, value });
}
