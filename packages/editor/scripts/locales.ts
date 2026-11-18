// SPDX-License-Identifier: GPL-3.0-or-later
//
// The locales check (ED-003 section 12.2; `npm run locales:check`, in CI): the message files
// against the sources and against each other. Run under vite-node, so it reads the files through
// the same parser the application formats them with (`app/i18n.ts`). What it checks:
//
// - every key the sources look up is in `en.json`, and every key in `en.json` is looked up (a
//   dynamic key, `t(`tab.${id}`)`, counts every key under its prefix as used);
// - every registered command has its label, `command.<name>`, and each argument its description,
//   `command.<name>.arg.<arg>` (the palette's data-driven keys);
// - every message in every file parses as the subset; a translation uses no argument English
//   lacks; each plural names only the categories its language has (`Intl.PluralRules`), and all
//   of them (a missing one falls to `other`, which reads wrong: Ukrainian's `few`);
// - a file other than English has no key English lacks, and its missing keys are counted: a
//   file may be incomplete (English fills in at run time) but never wrong.
//
// `check-locales.ts` runs it: the coverage per language (the README's line), exit 1 on an error.
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join, dirname, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { listCommands } from '@docx4j/editor-model';
import { messageError, argumentsOf, pluralCategoriesOf, type Messages } from '../src/app/i18n.js';

const PACKAGE = join(dirname(fileURLToPath(import.meta.url)), '..');
const SOURCES = join(PACKAGE, 'src');
const LOCALES = join(SOURCES, 'locales');

/** The lookups whose first argument is a key. */
const LOOKUPS = ['t', 'tr', 'tNodes', 'tOr', 'plainOr', 'hasMessage'];

/** A key: dotted words (`toolbar.bold`, `command.font.bold.arg.value`). A literal without a dot in a lookup's first argument is a comparison, not a key. */
const KEY = /^[A-Za-z][\w-]*(\.[\w-]+)+$/;

/** Messages taken as plain text, not ICU (`plainOr`): a command argument's description quotes code in braces. */
export const isPlain = (key: string): boolean => /^command\..+\.arg\.[^.]+$/.test(key);

export interface LocalesReport {
  errors: string[];
  /** Per language other than English: keys translated, keys missing (English fills in). */
  coverage: { tag: string; translated: number; missing: number; total: number; draft: boolean }[];
  keys: number;
}

function sourceFiles(dir: string): string[] {
  const out: string[] = [];
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) { if (name !== 'locales') out.push(...sourceFiles(path)); }
    else if (/\.tsx?$/.test(name) && !name.endsWith('.d.ts')) out.push(path);
  }
  return out;
}

/**
 * The keys a source looks up: the string literals in the first argument of each lookup (a
 * conditional names both), and the fixed head of a template literal as a prefix.
 */
export function keysInSource(text: string): { keys: Set<string>; prefixes: Set<string> } {
  const keys = new Set<string>();
  const prefixes = new Set<string>();
  const call = new RegExp(`(?<![\\w.])(?:${LOOKUPS.join('|')})\\(`, 'g');
  for (const m of text.matchAll(call)) {
    let i = m.index! + m[0].length;
    let depth = 0;
    // The first argument: up to a comma or the closing parenthesis at depth 0, strings skipped whole.
    while (i < text.length) {
      const c = text[i]!;
      if (c === "'" || c === '"' || c === '`') {
        const quote = c;
        const from = i + 1;
        i++;
        let dynamic = false;
        while (i < text.length && text[i] !== quote) {
          if (text[i] === '\\') i++;
          else if (quote === '`' && text[i] === '$' && text[i + 1] === '{') {
            if (!dynamic) prefixes.add(text.slice(from, i));
            dynamic = true;
            let braces = 0;
            for (i += 1; i < text.length; i++) { if (text[i] === '{') braces++; else if (text[i] === '}' && --braces === 0) break; }
          }
          i++;
        }
        if (!dynamic && KEY.test(text.slice(from, i))) keys.add(text.slice(from, i));
        i++;
        continue;
      }
      if (c === '(' || c === '[' || c === '{') depth++;
      else if (c === ')' || c === ']' || c === '}') { if (depth === 0) break; depth--; }
      else if (c === ',' && depth === 0) break;
      i++;
    }
  }
  return { keys, prefixes };
}

function readLocale(tag: string): Messages {
  return JSON.parse(readFileSync(join(LOCALES, `${tag}.json`), 'utf8')) as Messages;
}

/** The plural categories a language has, `other` among them. */
function categoriesOf(tag: string): Set<string> {
  return new Set(new Intl.PluralRules(tag).resolvedOptions().pluralCategories);
}

function checkMessages(tag: string, messages: Messages, en: Messages, errors: string[]): void {
  const categories = categoriesOf(tag);
  for (const [key, message] of Object.entries(messages)) {
    if (key.startsWith('_')) continue;
    if (typeof message !== 'string') { errors.push(`${tag}.json: ${key} is not a string`); continue; }
    if (tag !== 'en' && en[key] === undefined) { errors.push(`${tag}.json: ${key} is not in en.json`); continue; }
    if (message === '' || isPlain(key)) continue;
    const problem = messageError(message);
    if (problem) { errors.push(`${tag}.json: ${key}: ${problem}`); continue; }
    if (tag !== 'en') {
      const known = argumentsOf(en[key]!);
      for (const a of argumentsOf(message)) if (!known.has(a)) errors.push(`${tag}.json: ${key} uses {${a}}, which the English message does not`);
    }
    const used = pluralCategoriesOf(message);
    if (used.size > 0) {
      for (const c of used) if (!c.startsWith('=') && !categories.has(c)) errors.push(`${tag}.json: ${key}: "${c}" is not a plural category of ${tag} (${[...categories].join(', ')})`);
      for (const c of categories) if (!used.has(c)) errors.push(`${tag}.json: ${key}: the plural lacks "${c}", a category of ${tag}`);
    }
  }
}

export function checkLocales(): LocalesReport {
  const errors: string[] = [];
  const en = readLocale('en');
  const enKeys = Object.keys(en).filter((k) => !k.startsWith('_'));

  // The sources against English.
  const used = new Set<string>();
  const prefixes = new Set<string>();
  for (const file of sourceFiles(SOURCES)) {
    const found = keysInSource(readFileSync(file, 'utf8'));
    for (const k of found.keys) {
      used.add(k);
      if (en[k] === undefined) errors.push(`${relative(PACKAGE, file)}: "${k}" is not in en.json`);
    }
    for (const p of found.prefixes) {
      prefixes.add(p);
      if (!enKeys.some((k) => k.startsWith(p))) errors.push(`${relative(PACKAGE, file)}: no key in en.json starts with "${p}"`);
    }
  }
  for (const k of enKeys) if (!used.has(k) && ![...prefixes].some((p) => p !== '' && k.startsWith(p))) errors.push(`en.json: ${k} is not used in the sources`);

  // The data-driven keys: the palette's labels and argument descriptions.
  for (const spec of listCommands()) {
    if (en[`command.${spec.name}`] === undefined) errors.push(`en.json: no label for the command ${spec.name} (command.${spec.name})`);
    for (const arg of Object.keys(spec.args ?? {})) if (en[`command.${spec.name}.arg.${arg}`] === undefined) errors.push(`en.json: no description for ${spec.name}'s argument ${arg}`);
  }

  // Every file's messages.
  checkMessages('en', en, en, errors);
  const coverage: LocalesReport['coverage'] = [];
  for (const name of readdirSync(LOCALES).sort()) {
    const tag = name.replace(/\.json$/, '');
    if (tag === 'en' || !name.endsWith('.json')) continue;
    let messages: Messages;
    try { messages = readLocale(tag); } catch (e) { errors.push(`${name}: ${(e as Error).message}`); continue; }
    checkMessages(tag, messages, en, errors);
    if (typeof messages['_status'] !== 'string') errors.push(`${name}: no _status entry ("draft, not yet reviewed by a native speaker", or "reviewed by …")`);
    const translated = enKeys.filter((k) => typeof messages[k] === 'string' && messages[k] !== '').length;
    coverage.push({ tag, translated, missing: enKeys.length - translated, total: enKeys.length, draft: !/^reviewed/i.test(messages['_status'] ?? '') });
  }
  return { errors, coverage, keys: enKeys.length };
}

/** The coverage as a Markdown table, for docs/developer/localisation.md: the language in its own name, the keys, the status. */
export function coverageTable(report: LocalesReport): string {
  const name = (tag: string): string => { try { return new Intl.DisplayNames(['en'], { type: 'language' }).of(tag) ?? tag; } catch { return tag; } };
  const rows = report.coverage.map((c) => `| ${name(c.tag)} | \`${c.tag}\` | ${c.translated}/${c.total} | ${c.draft ? 'draft' : 'reviewed'} |`);
  return ['| Language | File | Keys | Status |', '|---|---|---|---|', `| English | \`en\` | ${report.keys}/${report.keys} | source |`, ...rows].join('\n');
}

/** The coverage as one line per language, for the console. */
export function coverageLines(report: LocalesReport): string[] {
  return report.coverage.map((c) => `${c.tag}: ${c.translated}/${c.total}${c.missing > 0 ? ` (${c.missing} missing)` : ''}${c.draft ? ', draft' : ''}`);
}
