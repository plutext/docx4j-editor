// SPDX-License-Identifier: GPL-3.0-or-later
//
// The locales check as a test (ED-003 section 12.5 item 2): `en.json` against the sources and
// the registry, every shipped file against `en.json`. `npm run locales:check` prints the same
// report with the coverage per language.
import { describe, expect, test } from 'vitest';
import { checkLocales, keysInSource } from '../scripts/locales.js';

describe('the locales check', () => {
  test('the message files agree with the sources, the registry and each other', () => {
    const report = checkLocales();
    expect(report.errors).toEqual([]);
    expect(report.keys).toBeGreaterThan(900);
  });

  test('the keys a source looks up: literals in the first argument, a template literal as a prefix', () => {
    const found = keysInSource([
      "t('toolbar.bold'); tr('a.b', { x: 1 });",
      "t(kind === 'header' ? 'view.headerOfPage' : 'view.footerOfPage', { page });",
      't(`tab.${id}`); tOr(`command.${spec.name}`, spec.label);',
      "format('not.a.lookup'); editor.run('history.undo');",
    ].join('\n'));
    expect([...found.keys].sort()).toEqual(['a.b', 'toolbar.bold', 'view.footerOfPage', 'view.headerOfPage']);
    expect([...found.prefixes].sort()).toEqual(['command.', 'tab.']);
  });
});
