// SPDX-License-Identifier: GPL-3.0-or-later
//
// The extension's words (P29, and ED-014 section 4.7, P48): the check `scripts/l10n.mjs --check` runs, over the files as
// they are, and over drafts made wrong one way at a time. The bundle holds every message the code passes to
// `vscode.l10n.t` and nothing else; the manifest's `%key%`s are all in `package.nls.json`; the drafts are the nineteen,
// named by VS Code's ids. `node scripts/l10n.mjs` writes the English bundle again.
import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'vitest';
// @ts-expect-error: a plain script, no declarations
import { BUNDLE, NLS, PRODUCT, checkWords, draftProblems, languages, messages, messagesIn, vscodeIdOf } from '../../scripts/l10n.mjs';

type Draft = Record<string, string>;
const problems = draftProblems as (name: string, english: Draft, draft: Draft, kind: 'manifest' | 'bundle') => string[];

describe('the extension\'s words', () => {
  test('the bundle is the messages the code uses', () => {
    const bundle = JSON.parse(readFileSync(BUNDLE, 'utf8')) as Draft;
    expect(Object.keys(bundle).sort()).toEqual((messages() as string[]).slice().sort());
    for (const [key, value] of Object.entries(bundle)) expect(value).toBe(key);
  });

  test('a message read from a single-quoted literal, its escapes undone; any other form named by its line (Fable\'s review, section 9.27)', () => {
    const text = [
      "vscode.l10n.t('The editor for {0} is not ready.', name);",
      "vscode.l10n.t(",
      "  'It\\'s {0}.', x);",
      'vscode.l10n.t("Double");',
      'vscode.l10n.t(`Template ${x}`);',
      'vscode.l10n.t(words);',
    ].join('\n');
    expect(messagesIn(text)).toEqual({ found: ['The editor for {0} is not ready.', "It's {0}."], unreadable: [4, 5, 6] });
  });

  test('the check finds nothing wrong, and every draft is whole', () => {
    const { errors, coverage, total } = checkWords() as { errors: string[]; coverage: { id: string; manifest: number; bundle: number }[]; total: { manifest: number; bundle: number } };
    expect(errors).toEqual([]);
    expect(coverage.map((c) => c.id)).toEqual(languages());
    for (const c of coverage) expect(c, c.id).toEqual({ id: c.id, manifest: total.manifest, bundle: total.bundle, draft: true });
  });

  test('the nineteen named by VS Code\'s ids: the application\'s zh-Hans, zh-Hant and pt as zh-cn, zh-tw and pt-br (P48)', () => {
    expect(vscodeIdOf('zh-Hans')).toBe('zh-cn');
    expect(vscodeIdOf('zh-Hant')).toBe('zh-tw');
    expect(vscodeIdOf('pt')).toBe('pt-br');
    expect(vscodeIdOf('fr')).toBe('fr');
    const ids = languages() as string[];
    expect(ids).toHaveLength(19);
    expect(ids).toEqual(expect.arrayContaining(['zh-cn', 'zh-tw', 'pt-br', 'ar', 'uk']));
    expect(ids.some((id) => /[A-Z]/.test(id) || id === 'pt')).toBe(false);
  });
});

describe('a draft made wrong', () => {
  const nls = JSON.parse(readFileSync(NLS, 'utf8')) as Draft;
  const english: Draft = {
    'The editor for {0} is not ready.': 'The editor for {0} is not ready.',
    'Join a Room': 'Join a Room',
  };
  const good: Draft = { _status: 'draft', 'The editor for {0} is not ready.': 'L\'éditeur de {0} n\'est pas prêt.', 'Join a Room': 'Rejoindre une salle' };

  test('a good one, and one lacking a key (VS Code shows English\'s): nothing', () => {
    expect(problems('fr', english, good, 'bundle')).toEqual([]);
    expect(problems('fr', english, { _status: 'draft', 'Join a Room': 'Rejoindre une salle' }, 'bundle')).toEqual([]);
  });

  test('no _status; a key English lacks; an empty message', () => {
    const { _status: _, ...unmarked } = good;
    expect(problems('fr', english, unmarked, 'bundle')).toEqual(['fr: no _status marking it a draft']);
    expect(problems('fr', english, { ...good, 'Join a room': 'x' }, 'bundle')).toEqual(['fr: "Join a room" is not in English']);
    expect(problems('fr', english, { ...good, 'Join a Room': ' ' }, 'bundle')).toEqual(['fr: "Join a Room" is not a message']);
  });

  test('a placeholder lost, renamed or added', () => {
    const key = 'The editor for {0} is not ready.';
    for (const wrong of ['L\'éditeur n\'est pas prêt.', 'L\'éditeur de {nom} n\'est pas prêt.', 'L\'éditeur de {0} ({1}) n\'est pas prêt.']) {
      expect(problems('fr', english, { ...good, [key]: wrong }, 'bundle'), wrong).toHaveLength(1);
    }
  });

  test('the manifest\'s backquoted code changed, and the product\'s name translated', () => {
    const base = nls['config.ask.baseUrl']!;
    expect(base).toContain('`https://api.anthropic.com`');
    const translated = base.replace('`https://api.anthropic.com`', '`https://api.anthropic.fr`');
    expect(problems('fr', nls, { _status: 'draft', 'config.ask.baseUrl': translated }, 'manifest')).toEqual([
      'fr: "config.ask.baseUrl" changes the code English quotes (`https://api.anthropic.com`)',
    ]);
    expect(nls['category']).toBe(PRODUCT);
    expect(problems('fr', nls, { _status: 'draft', category: 'éditeur docx4j' }, 'manifest')).toEqual([
      'fr: "category" lacks "docx4j editor", which is never translated',
    ]);
  });
});
