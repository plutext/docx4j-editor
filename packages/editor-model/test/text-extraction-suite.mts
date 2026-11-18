// SPDX-License-Identifier: GPL-3.0-or-later
//
// The text extraction conformance suite (`fixtures/text-extraction/`): docx4j's probe and core-ts's runner,
// copied unedited by `scripts/copy-text-extraction-suite.mjs` (ED-008 step 2). The public repository leaves the
// folder out (the release plan's "what crosses the fence", Jason, 2026-11-18), so the tests that read it load the
// runner from here and skip where `suitePresent` is false. The runner's types are declared here, the one part of
// it that is the editor's own; only what the tests use is declared.
import { existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

/** The copy's folder, and the probe's directory within it. */
export const SUITE = join(dirname(fileURLToPath(import.meta.url)), 'fixtures', 'text-extraction');
export const PHASE0 = join(SUITE, 'suite', 'phase0');

/** Whether the copy is here: false in the public repository. */
export const suitePresent = existsSync(join(SUITE, 'conformance.mjs')) && existsSync(PHASE0);

/** A row of a suite table: its columns by the header's names. */
export type Row = Record<string, string>;

/** What the runner asks of an engine (docx4j's `TextsAdapter` contract, GOLDENS.md). */
export interface Engine {
  implementedStatements(): readonly string[] | Promise<readonly string[]>;
  extract(fixture: string, view: string, scenario: string, format: string): Promise<{ text: string; findings: Record<string, string> }>;
}

export interface RunResult {
  passed: number;
  pending: number;
  /** One line per row checked: fixture, case, item, view, scenario, status, reason, tab-separated, after a header line. */
  report: string[];
}

/** core-ts's runner (`conformance.mjs`), as the tests use it. */
export interface Runner {
  table(file: string): Row[];
  /** WD REQ-1201's normalisation, under which texts are compared. */
  normalize(value: string): string;
  /** Markdown with each image's destination written `…`. */
  imageDestinations(markdown: string): string;
  /** Runs the suite in a directory against an engine (null: every row pending). Throws at the first row that fails. */
  run(dir: string, engine: Engine | null): Promise<RunResult>;
  HTML_READER_ABSENT: string;
}

/** The runner, loaded from the copy; only where `suitePresent`. */
export async function loadRunner(): Promise<Runner> {
  if (!suitePresent) throw new Error(`the text extraction suite is not here (${SUITE})`);
  return (await import(pathToFileURL(join(SUITE, 'conformance.mjs')).href)) as Runner;
}
