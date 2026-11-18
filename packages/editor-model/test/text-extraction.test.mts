// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-008 step 2: the installed engine's text model held to the text extraction specification's conformance
// suite, so that a failure in the companion's reading tools (ED-008 step 3) is the engine's or the copy's
// before it is a tool's. The suite is docx4j's probe (`fixtures/text-extraction/suite/`, copied unedited at the
// docx4j commit the installed core-ts release was held to), and the runner is core-ts's mirror of docx4j's
// `ExtractionConformance`, copied unedited from the release's tag (`manifest.json`). The adapter below is the
// editor's: core-ts's `TextsAdapter` mirror over the package as installed, claiming the statements the release
// claims, so that every row the engine passes passes here and every row it leaves pending pends here. The public
// repository leaves the copy out, and there the suite skips (`text-extraction-suite.mts`).
import { createHash } from 'node:crypto';
import { readFileSync, readdirSync } from 'node:fs';
import { join, relative, sep } from 'node:path';
import { describe, expect, test } from 'vitest';
import { WordprocessingMLPackage, Texts, PlainText, TextOptions, MarkdownExporter } from '@docx4j/core-ts';
import { SUITE as base, PHASE0 as phase0, suitePresent, loadRunner, type Engine, type Runner } from './text-extraction-suite.mjs';

const { run, table, imageDestinations, HTML_READER_ABSENT } = suitePresent ? await loadRunner() : ({} as Runner);
const manifest = (suitePresent ? JSON.parse(readFileSync(join(base, 'manifest.json'), 'utf8')) : {}) as {
  docx4j: { commit: string }; 'core-ts': { tag: string; commit: string };
  files: { sha1: string; path: string }[];
};

/**
 * The statements core-ts 0.8.0's adapter claims (its `test/text-extraction/adapter.mjs` at the tag, CR-012
 * sections 15, 16, 23 and 24): a release that claims more moves rows from pending to passed, and the counts
 * below are recounted at its install, with the engine's own.
 */
const IMPLEMENTED = [
  'REQ-201', 'REQ-202', 'REQ-204',
  'REQ-301', 'REQ-302', 'REQ-303', 'REQ-304', 'REQ-305', 'REQ-306', 'REQ-307',
  'REQ-401', 'REQ-402', 'REQ-403', 'REQ-404', 'REQ-405', 'REQ-406',
  'REQ-501', 'REQ-502', 'REQ-503', 'REQ-504', 'REQ-505', 'REQ-507', 'REQ-508', 'REQ-509', 'REQ-510',
  'REQ-514', 'REQ-515', 'REQ-516', 'REQ-517', 'REQ-518', 'REQ-520',
  'REQ-601', 'REQ-602', 'REQ-603', 'REQ-605', 'REQ-606',
  'REQ-701', 'REQ-702', 'REQ-703', 'REQ-704', 'REQ-706', 'REQ-707', 'REQ-708', 'REQ-709',
  'REQ-801', 'REQ-802', 'REQ-1001', 'REQ-1202',
  'REQ-1002', 'REQ-607', 'REQ-521', 'REQ-710',
];

/** A scenario's options: `default` (and `markdown`), or the WD's options as `name: value`, several `; ` apart. */
function optionsOf(view: string, scenario: string | undefined): TextOptions {
  const options = TextOptions.of({ view: view as 'accepted' });
  if (scenario === undefined || scenario === 'default' || scenario === 'markdown' || scenario.startsWith('html reader')) return options;
  for (const option of scenario.split(';')) {
    const colon = option.indexOf(':');
    if (colon < 0) throw new Error(`Unknown scenario: ${scenario}`);
    options.set(option.slice(0, colon).trim(), option.slice(colon + 1).trim());
  }
  return options;
}

/** The engine as installed: a package per fixture and repair setting, findings from the warnings and the metadata. */
function textsAdapter(): Engine {
  const packages = new Map<string, Promise<WordprocessingMLPackage>>();
  return {
    implementedStatements: () => [...IMPLEMENTED, HTML_READER_ABSENT],
    async extract(fixture, view, scenario, format) {
      if (format !== 'txt' && format !== 'md') throw new Error(`format ${format}`);
      const options = optionsOf(view, scenario);
      const key = `${fixture}\u0000${options.repair}`;
      let pkg = packages.get(key);
      if (!pkg) packages.set(key, (pkg = WordprocessingMLPackage.load(readFileSync(fixture), { repair: options.repair })));
      const loaded = await pkg;
      const result = await Texts.of(loaded, options);
      const findings: Record<string, string> = {};
      for (const w of result.warnings) {
        const k = `warning\t${w.kind}\t${w.selector ?? ''}`;
        findings[k] = String(Number(findings[k] ?? 0) + 1);
      }
      for (const [k, v] of Object.entries(result.metadata)) {
        if (v !== null && typeof v === 'object') for (const [pk, pv] of Object.entries(v)) findings[`metadata\t${pk}\t`] = String(pv);
        else findings[`metadata\t${k}\t`] = String(v);
      }
      return { text: format === 'md' ? await MarkdownExporter.export(result, loaded) : PlainText.of(result), findings };
    },
  };
}

const ENGINE = textsAdapter();

describe.skipIf(!suitePresent)('the text extraction conformance suite over the installed engine (ED-008 step 2)', () => {
  test('the copy is docx4j\'s and core-ts\'s, unedited: every file as manifest.json records it, and nothing more', () => {
    expect(manifest.docx4j.commit).toMatch(/^[0-9a-f]{40}$/);
    expect(manifest['core-ts'].tag).toBe('0.8.0');
    const sha1 = (file: string): string => createHash('sha1').update(readFileSync(file)).digest('hex');
    for (const f of manifest.files) expect(sha1(join(base, f.path)), f.path).toBe(f.sha1);
    const listed = new Set(manifest.files.map((f) => f.path));
    const walk = (dir: string): string[] => readdirSync(dir, { withFileTypes: true }).flatMap((e) => e.isDirectory() ? walk(join(dir, e.name)) : [relative(base, join(dir, e.name)).split(sep).join('/')]);
    for (const file of walk(base)) if (file !== 'manifest.json') expect(listed.has(file), `${file} is not the suite's: recopy, never edit`).toBe(true);
  });

  test('every row the engine claims passes, and every other row pends for a statement it does not claim: the release\'s own counts', async () => {
    const result = await run(phase0, ENGINE);
    // core-ts 0.8.0's own test over this suite (test/text-extraction.test.mjs at the tag): 795 passed, 977 pending.
    expect(result.passed + result.pending).toBe(1772);
    expect(result.passed, 'a row moved: the engine changed, or the copy did').toBe(795);
    const claimed = new Set(IMPLEMENTED);
    for (const line of result.report.slice(1).filter((l) => l.includes('\tPENDING\t') && !/\tPENDING\t\d+ rows? of /.test(l))) {
      const missing = line.split('\t')[6]!.split('; ')[1]!.split(', ');
      expect(missing.length > 0 && missing.every((r) => !claimed.has(r)), line).toBe(true);
    }
  }, 120_000);

  test('the whole-document goldens (REQ-1203, not yet claimed): exact, but for the lines of the engine\'s phase 5', async () => {
    // docx4j's goldens are written with their trailing line ends trimmed to one
    const asWritten = (text: string): string => `${text.replace(/\n+$/, '')}\n`;
    // the lines phase 5 adds: LaTeX math, the chart and SmartArt, the expanded chunks
    const PENDING = /MTH32|CHT34_|SMA35_|ALT4[3-7]_|^\[chart: /;
    const without = (text: string): string => text.split('\n').filter((line) => !PENDING.test(line)).join('\n');
    const PENDING_MD = /MTH32|CHT34\\_|SMA35\\_|ALT4[3-7]\\_/;
    const withoutMd = (text: string): string => text.split('\n\n').filter((block) => !PENDING_MD.test(block)).join('\n\n');
    let exact = 0;
    for (const row of table(join(phase0, 'goldens.tsv'))) {
      const ours = asWritten(imageDestinations((await ENGINE.extract(join(phase0, row['fixture']!), row['view']!, row['scenario']!, row['format']!)).text));
      const golden = readFileSync(join(phase0, row['golden']!), 'utf8');
      if (ours === golden) { exact++; continue; }
      expect(row['fixture'], `${row['golden']} differs`).toBe('extraction-probe.docx');
      if (row['format'] === 'md') expect(withoutMd(ours), row['golden']).toBe(withoutMd(golden));
      else expect(without(ours), row['golden']).toBe(without(golden));
    }
    expect(exact, 'the external and structure probes\' goldens, in every view and in Markdown').toBe(8);
  }, 60_000);
});
