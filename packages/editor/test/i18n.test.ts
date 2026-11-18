// SPDX-License-Identifier: GPL-3.0-or-later
//
// The formatter and the language match (ED-003 section 12.5 item 1): arguments, plurals in
// English, French, Ukrainian and Arabic through `Intl.PluralRules`, select, apostrophe quoting,
// a missing key's fallback reported once, the match of the browser's languages to the files.
import { describe, expect, test } from 'vitest';
import { format, t, onMissingKey, matchLanguage, directionOf, messageError, argumentsOf, pluralCategoriesOf, availableLanguages, SOURCE_LANGUAGE } from '../src/app/i18n.js';

describe('the formatter', () => {
  test('arguments by name, numbers in the language, an absent argument left as it is', () => {
    expect(format('Saved {name} ({how})', { name: 'a.docx', how: 'download' }, 'en')).toBe('Saved a.docx (download)');
    expect(format('{n} bytes', { n: 12345.5 }, 'en')).toBe('12,345.5 bytes');
    expect(format('{n} octets', { n: 12345.5 }, 'fr')).toBe('12 345,5 octets');
    expect(format('Hello {who}', {}, 'en')).toBe('Hello {who}');
  });

  test('plurals in English and French: French counts 0 and 1 as one', () => {
    const m = '{count, plural, one {# paragraph} other {# paragraphs}}';
    expect(format(m, { count: 1 }, 'en')).toBe('1 paragraph');
    expect(format(m, { count: 0 }, 'en')).toBe('0 paragraphs');
    expect(format(m, { count: 1500 }, 'en')).toBe('1,500 paragraphs');
    const fr = '{count, plural, one {# paragraphe} many {# de paragraphes} other {# paragraphes}}';
    expect(format(fr, { count: 0 }, 'fr')).toBe('0 paragraphe');
    expect(format(fr, { count: 1 }, 'fr')).toBe('1 paragraphe');
    expect(format(fr, { count: 2 }, 'fr')).toBe('2 paragraphes');
  });

  test('Ukrainian: one, few, many and other', () => {
    const m = '{count, plural, one {# абзац} few {# абзаци} many {# абзаців} other {# абзацу}}';
    expect(format(m, { count: 1 }, 'uk')).toBe('1 абзац');
    expect(format(m, { count: 21 }, 'uk')).toBe('21 абзац');
    expect(format(m, { count: 3 }, 'uk')).toBe('3 абзаци');
    expect(format(m, { count: 5 }, 'uk')).toBe('5 абзаців');
    expect(format(m, { count: 11 }, 'uk')).toBe('11 абзаців');
    expect(format(m, { count: 1.5 }, 'uk')).toBe('1,5 абзацу');
  });

  test('Arabic: all six categories, chosen by Intl.PluralRules without a special case', () => {
    const m = '{count, plural, zero {z} one {o} two {t} few {f} many {m} other {x}}';
    const got = [0, 1, 2, 3, 11, 100].map((n) => format(m, { count: n }, 'ar'));
    expect(got).toEqual(['z', 'o', 't', 'f', 'm', 'x']);
    expect(pluralCategoriesOf(m)).toEqual(new Set(['zero', 'one', 'two', 'few', 'many', 'other']));
  });

  test('an exact match beats the category; # is the number inside a plural only', () => {
    const m = '{count, plural, =0 {none} one {# item} other {# items}} #';
    expect(format(m, { count: 0 }, 'en')).toBe('none #');
    expect(format(m, { count: 1 }, 'en')).toBe('1 item #');
  });

  test('select, with a plural inside it', () => {
    const m = '{position, select, first {the first} other {the last}} child; {kind, select, header {{n, plural, one {# header} other {# headers}}} other {{n} others}}';
    expect(format(m, { position: 'first', kind: 'header', n: 2 }, 'en')).toBe('the first child; 2 headers');
    expect(format(m, { position: 'x', kind: 'y', n: 2 }, 'en')).toBe('the last child; 2 others');
  });

  test("apostrophes: a lone one is itself, '' is one, '{' quotes a brace", () => {
    expect(format("l'éditeur de {name}", { name: 'Jean' }, 'fr')).toBe("l'éditeur de Jean");
    expect(format("it''s", {}, 'en')).toBe("it's");
    expect(format("'{'literal'}' {x}", { x: 1 }, 'en')).toBe('{literal} 1');
    expect(format("{count, plural, other {'#' is #}}", { count: 3 }, 'en')).toBe('# is 3');
  });

  test('a message outside the subset is reported by the check and shown as it is', () => {
    expect(messageError('{count, plural, one {x}}')).toContain('"other"');
    expect(messageError('{a, bogus}')).toContain('unknown type');
    expect(messageError('unmatched }')).toContain('unmatched');
    expect(format('{broken', {}, 'en')).toBe('{broken');
  });

  test('the arguments a message uses, for the check against English', () => {
    expect(argumentsOf('{a} and {n, plural, one {# {b}} other {{c}}}')).toEqual(new Set(['a', 'n', 'b', 'c']));
  });
});

describe('the lookup', () => {
  test('a key is looked up in English, the source; a missing key is itself, reported once', () => {
    expect(t('toolbar.bold')).toBe('Bold');
    expect(t('status.saved.changed', { how: t('status.saved.how.download') })).toBe('Saved with changes (download)');
    const missing: string[] = [];
    const stop = onMissingKey((k) => missing.push(k));
    expect(t('no.such.key')).toBe('no.such.key');
    expect(t('no.such.key')).toBe('no.such.key');
    stop();
    expect(missing).toEqual(['no.such.key']);
  });

  test('English is shipped first', () => {
    expect(availableLanguages()[0]).toBe(SOURCE_LANGUAGE);
  });
});

describe('the language match', () => {
  const shipped = ['en', 'fr', 'pt', 'uk', 'zh-Hans', 'zh-Hant', 'ar'];

  test('the exact tag, then the language with its likely script, then the language alone', () => {
    expect(matchLanguage(['fr-CA'], shipped)).toBe('fr');
    expect(matchLanguage(['pt-PT'], shipped)).toBe('pt');
    expect(matchLanguage(['pt-PT'], [...shipped, 'pt-PT'])).toBe('pt-PT');
    expect(matchLanguage(['zh-TW'], shipped)).toBe('zh-Hant');
    expect(matchLanguage(['zh-HK'], shipped)).toBe('zh-Hant');
    expect(matchLanguage(['zh-CN'], shipped)).toBe('zh-Hans');
    expect(matchLanguage(['zh'], shipped)).toBe('zh-Hans');
    expect(matchLanguage(['uk-UA'], shipped)).toBe('uk');
  });

  test('the preferences in order; none matching is none, and the application takes English', () => {
    expect(matchLanguage(['sw'], shipped)).toBeUndefined();
    expect(matchLanguage(['sw', 'fr-FR', 'en'], shipped)).toBe('fr');
    expect(matchLanguage(['not a tag', 'ar-EG'], shipped)).toBe('ar');
    expect(matchLanguage([], shipped)).toBeUndefined();
  });

  test('the chrome is right-to-left for Arabic, left-to-right otherwise', () => {
    expect(directionOf('ar')).toBe('rtl');
    expect(directionOf('ar-EG')).toBe('rtl');
    expect(directionOf('fr')).toBe('ltr');
    expect(directionOf('zh-Hant')).toBe('ltr');
  });
});

describe('arguments in a right-to-left language', () => {
  test('a string argument is isolated in Arabic, a number is formatted for Arabic (Latin digits in current locale data) and not isolated, and nothing changes in French', () => {
    expect(format('تم الربط بـ {xpath}', { xpath: '/invoice[1]/name[1]' }, 'ar')).toBe('تم الربط بـ ⁨/invoice[1]/name[1]⁩');
    expect(format('{count, plural, zero {لا شيء} one {واحد} two {اثنان} few {# قليلة} many {# كثيرة} other {# عنصر}}', { count: 3 }, 'ar')).toBe(`${new Intl.NumberFormat('ar').format(3)} قليلة`);
    expect(format('Lié à {xpath}', { xpath: '/a' }, 'fr')).toBe('Lié à /a');
  });
});
