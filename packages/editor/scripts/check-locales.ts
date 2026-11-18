// SPDX-License-Identifier: GPL-3.0-or-later
//
// `npm run locales:check` (ED-003 section 12.2): the check of `locales.ts`, printed, failing on an
// error. `UPDATE_COVERAGE=1` writes the coverage table into docs/developer/localisation.md between its
// markers; without it a stale table is reported as a warning.
import { readFileSync, writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { checkLocales, coverageLines, coverageTable } from './locales.js';

const GUIDE = join(dirname(fileURLToPath(import.meta.url)), '..', '..', '..', 'docs', 'developer', 'localisation.md');
const START = '<!-- locales:coverage -->';
const END = '<!-- /locales:coverage -->';

const report = checkLocales();
console.log(`en: ${report.keys} keys`);
for (const line of coverageLines(report)) console.log(line);

const guide = readFileSync(GUIDE, 'utf8');
const from = guide.indexOf(START);
const to = guide.indexOf(END);
if (from >= 0 && to > from) {
  const next = `${guide.slice(0, from + START.length)}\n${coverageTable(report)}\n${guide.slice(to)}`;
  if (next !== guide) {
    if (process.env['UPDATE_COVERAGE'] === '1') { writeFileSync(GUIDE, next); console.log('docs/developer/localisation.md: coverage table written'); }
    else console.warn('warning: docs/developer/localisation.md\'s coverage table is stale (UPDATE_COVERAGE=1 npm run locales:check writes it)');
  }
}

if (report.errors.length > 0) {
  for (const e of report.errors) console.error(e);
  console.error(`${report.errors.length} error${report.errors.length === 1 ? '' : 's'}`);
  process.exit(1);
}
