// SPDX-License-Identifier: GPL-3.0-or-later
//
// Localisation of the application's own words (ED-003 section 12; ED-001 decision 22; section 9
// item 37): flat message files per language under `src/locales/`, English the source and the
// fallback, looked up by stable keys through `t(key, args)`. The messages are ICU MessageFormat's
// subset translators know: `{name}` arguments, `{count, plural, one {# paragraph} other {#
// paragraphs}}` with `Intl.PluralRules` choosing the category (and `=0` exact matches), `{kind,
// select, ...}`, `{n, number}`, `{d, date}`, apostrophe quoting as ICU's ('' for a quote, '{' for
// a literal brace; any other apostrophe is itself). A number argument is formatted as a number
// in the language (grouped: `1,500`), so an identifier goes in as a string. The document is never
// translated: its text, its style names and its labels are the document's.
import { Fragment, createElement, useSyncExternalStore, type ReactNode } from 'react';

export type Messages = Record<string, string>;

// The message files are the site's own (section 12.7: nothing from a server). English is in the
// bundle; every other language is a chunk of its own (a file is some 60 kB, 15 kB compressed),
// loaded when it is chosen, before the first render when the browser asks for it at start.
const tagOf = (path: string): string => /([^/]+)\.json$/.exec(path)![1]!;
const english = import.meta.glob<Messages>('../locales/en.json', { eager: true, import: 'default' });
const loaders = new Map(Object.entries(import.meta.glob<Messages>(['../locales/*.json', '!../locales/en.json'], { import: 'default' })).map(([path, load]) => [tagOf(path), load]));

const catalogue = new Map<string, Messages>();
for (const [path, messages] of Object.entries(english)) catalogue.set(tagOf(path), messages);

export const SOURCE_LANGUAGE = 'en';

/** The languages shipped, English first, then by tag. */
export function availableLanguages(): string[] {
  return [SOURCE_LANGUAGE, ...[...loaders.keys()].sort()];
}

function shipped(tag: string): boolean {
  return tag === SOURCE_LANGUAGE || loaders.has(tag);
}

/** A language's messages, loaded once; English's are there from the start. */
export async function loadMessages(tag: string): Promise<Messages | undefined> {
  const have = catalogue.get(tag);
  if (have) return have;
  const load = loaders.get(tag);
  if (!load) return undefined;
  const messages = await load();
  catalogue.set(tag, messages);
  return messages;
}

/** The messages of a language once loaded, or undefined. */
export function messagesOf(tag: string): Messages | undefined {
  return catalogue.get(tag);
}

/**
 * A loaded file's `_status` entry: "draft, not yet reviewed by a native speaker" until a review,
 * then "reviewed by …". Every file but English's has one (the check says so). Read from the loaded
 * file only: importing it eagerly as well would pull every language into the main chunk.
 */
export function statusOf(tag: string): string | undefined {
  return catalogue.get(tag)?.['_status'];
}

/** Whether a loaded language's file is still a draft: its status does not say it was reviewed. */
export function isDraft(tag: string): boolean {
  const s = statusOf(tag);
  return s !== undefined && !/^reviewed/i.test(s);
}

// ---------------------------------------------------------------- the language match

function canonical(tag: string): string | undefined {
  try { return Intl.getCanonicalLocales(tag)[0]; } catch { return undefined; }
}

/**
 * The best shipped language for the user's preferences, in order: the exact tag (`pt-PT` once a
 * `pt-PT.json` ships), the language with its likely script (`zh-TW` and `zh-HK` to `zh-Hant`,
 * `zh-CN` to `zh-Hans`), the language alone (`fr-CA` to `fr`). Undefined when nothing matches.
 */
export function matchLanguage(preferences: readonly string[], available: readonly string[]): string | undefined {
  const byLower = new Map(available.map((a) => [a.toLowerCase(), a]));
  for (const pref of preferences) {
    const tag = canonical(pref);
    if (!tag) continue;
    const exact = byLower.get(tag.toLowerCase());
    if (exact) return exact;
    let locale: Intl.Locale;
    try { locale = new Intl.Locale(tag).maximize(); } catch { continue; }
    if (locale.script) {
      const scripted = byLower.get(`${locale.language}-${locale.script}`.toLowerCase());
      if (scripted) return scripted;
    }
    const bare = byLower.get(locale.language.toLowerCase());
    if (bare) return bare;
  }
  return undefined;
}

const RTL = new Set(['ar', 'fa', 'he', 'ur', 'ps', 'sd', 'ug', 'yi', 'dv', 'ckb']);

/** The chrome's direction for a language: right-to-left for Arabic and its neighbours. */
export function directionOf(tag: string): 'ltr' | 'rtl' {
  let lang = tag.split('-')[0]!.toLowerCase();
  try { lang = new Intl.Locale(tag).language; } catch { /* the plain split will do */ }
  return RTL.has(lang) ? 'rtl' : 'ltr';
}

/** Whether an element lays out right-to-left (the chrome's direction where it sits): arrow keys and pointer arithmetic follow it. */
export function isRightToLeft(el: Element | null): boolean {
  return el !== null && getComputedStyle(el).direction === 'rtl';
}

/** The language's name in itself, for the Options list: "Français", "العربية". */
export function nativeName(tag: string): string {
  try { return new Intl.DisplayNames([tag], { type: 'language' }).of(tag) ?? tag; } catch { return tag; }
}

// ---------------------------------------------------------------- the current language

export const LANGUAGE_KEY = 'docx4j-editor.language';

/** The override in Options, or null to follow the browser. */
export function languageOption(): string | null {
  try {
    const v = localStorage.getItem(LANGUAGE_KEY);
    return v && shipped(v) ? v : null;
  } catch { return null; }
}

function browserLanguages(): readonly string[] {
  if (typeof navigator === 'undefined') return [];
  return navigator.languages?.length ? navigator.languages : navigator.language ? [navigator.language] : [];
}

/** The host's display language, given at start (VS Code's `env.language`, ED-014 P27): preferred to the browser's own. */
let hostLanguage: string | undefined;

function preferredLanguages(): readonly string[] {
  return hostLanguage ? [hostLanguage, ...browserLanguages()] : browserLanguages();
}

/** The language at start: the override, else the best match of the host's and then the browser's, else English. */
export function initialLanguage(): string {
  return languageOption() ?? matchLanguage(preferredLanguages(), availableLanguages()) ?? SOURCE_LANGUAGE;
}

/** English until `languageReady` has loaded the language the start asks for. */
let current = SOURCE_LANGUAGE;
const listeners = new Set<() => void>();

export function language(): string {
  return current;
}

/** Switch to a language once its messages are in; every `useT` re-renders. */
async function switchTo(next: string): Promise<void> {
  await loadMessages(next);
  if (next === current) return;
  current = next;
  for (const l of [...listeners]) l();
}

/**
 * Choose a language (null: follow the browser again), persisted; resolves once the application
 * shows it. A file that does not load (a tab opened before a redeploy asks for a chunk whose
 * hashed name is gone) rejects and leaves the language as it was (the choice is kept, so a reload shows it); the caller says so.
 */
export async function setLanguage(tag: string | null): Promise<void> {
  try {
    if (tag === null) localStorage.removeItem(LANGUAGE_KEY); else localStorage.setItem(LANGUAGE_KEY, tag);
  } catch { /* storage may be unavailable */ }
  const next = tag !== null && shipped(tag) ? tag : matchLanguage(preferredLanguages(), availableLanguages()) ?? SOURCE_LANGUAGE;
  await switchTo(next);
}

/**
 * The language at start, loaded (main.tsx awaits it before the first render, so a French browser
 * never sees the English interface first). A file that fails to load leaves English. `host` is the
 * host's display language, if it has one (VS Code's, ED-014 P27).
 */
export async function languageReady(host?: string): Promise<void> {
  hostLanguage = host;
  const want = initialLanguage();
  current = SOURCE_LANGUAGE;
  try { await switchTo(want); } catch { /* English stays */ }
}

export function onLanguageChange(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

// ---------------------------------------------------------------- the formatter

type Part =
  | string
  | { arg: string; type?: undefined }
  | { arg: string; type: 'number' | 'date' | 'time'; style?: string }
  | { arg: string; type: 'plural' | 'selectordinal' | 'select'; options: Map<string, Part[]> };

class MessageSyntaxError extends Error {}

/** Parse a message into text and arguments; `inPlural` makes `#` special. */
function parse(message: string): Part[] {
  let i = 0;
  const fail = (what: string): never => { throw new MessageSyntaxError(`${what} at ${i} in "${message}"`); };
  const ws = (): void => { while (i < message.length && /\s/.test(message[i]!)) i++; };
  const word = (): string => {
    const start = i;
    while (i < message.length && !/[\s,{}]/.test(message[i]!)) i++;
    if (i === start) fail('a name expected');
    return message.slice(start, i);
  };
  const parts = (depth: number, inPlural: boolean): Part[] => {
    const out: Part[] = [];
    let text = '';
    const flush = (): void => { if (text) { out.push(text); text = ''; } };
    while (i < message.length) {
      const c = message[i]!;
      if (c === "'") {
        const next = message[i + 1];
        if (next === "'") { text += "'"; i += 2; continue; }
        if (next === '{' || next === '}' || (inPlural && next === '#')) {
          // Quoted literal text until the next lone apostrophe.
          i++;
          while (i < message.length) {
            if (message[i] === "'") {
              if (message[i + 1] === "'") { text += "'"; i += 2; continue; }
              i++;
              break;
            }
            text += message[i];
            i++;
          }
          continue;
        }
        text += c; i++; continue;
      }
      if (c === '}') {
        if (depth === 0) fail('an unmatched }');
        break;
      }
      if (c === '#' && inPlural) { flush(); out.push({ arg: '#' }); i++; continue; }
      if (c === '{') {
        flush();
        i++;
        ws();
        const arg = word();
        ws();
        if (message[i] === '}') { i++; out.push({ arg }); continue; }
        if (message[i] !== ',') fail('"," or "}" expected');
        i++;
        ws();
        const type = word();
        ws();
        if (type === 'number' || type === 'date' || type === 'time') {
          let style: string | undefined;
          if (message[i] === ',') { i++; ws(); style = word(); ws(); }
          if (message[i] !== '}') fail('"}" expected');
          i++;
          out.push(style === undefined ? { arg, type } : { arg, type, style });
          continue;
        }
        if (type !== 'plural' && type !== 'select' && type !== 'selectordinal') fail(`an unknown type ${type}`);
        if (message[i] !== ',') fail('"," expected');
        i++;
        const options = new Map<string, Part[]>();
        for (;;) {
          ws();
          if (message[i] === '}') { i++; break; }
          if (i >= message.length) fail('an unclosed argument');
          const selector = word();
          ws();
          if (message[i] !== '{') fail('"{" expected');
          i++;
          options.set(selector, parts(depth + 1, type !== 'select' || inPlural));
          if (message[i] !== '}') fail('"}" expected');
          i++;
        }
        if (!options.has('other')) fail('an "other" option expected');
        out.push({ arg, type: type as 'plural' | 'select' | 'selectordinal', options });
        continue;
      }
      text += c;
      i++;
    }
    flush();
    return out;
  };
  const result = parts(0, false);
  if (i < message.length) fail('trailing text');
  return result;
}

const parsed = new Map<string, Part[]>();

function parsedOf(message: string): Part[] {
  let p = parsed.get(message);
  if (!p) { p = parse(message); parsed.set(message, p); }
  return p;
}

/** Whether a message parses as the subset; the error's message otherwise (the locales check uses it). */
export function messageError(message: string): string | undefined {
  try { parse(message); return undefined; } catch (e) { return (e as Error).message; }
}

/** The argument names a message uses, `#` excluded (the locales check compares them per key). */
export function argumentsOf(message: string): Set<string> {
  const names = new Set<string>();
  const walk = (ps: Part[]): void => {
    for (const p of ps) {
      if (typeof p === 'string' || p.arg === '#') continue;
      names.add(p.arg);
      if ('options' in p) for (const o of p.options.values()) walk(o);
    }
  };
  walk(parse(message));
  return names;
}

/** The plural selectors a message's plural arguments use (`one`, `few`, `=0`), for the check. */
export function pluralCategoriesOf(message: string): Set<string> {
  const out = new Set<string>();
  const walk = (ps: Part[]): void => {
    for (const p of ps) {
      if (typeof p === 'string' || !('options' in p)) continue;
      if (p.type === 'plural') for (const k of p.options.keys()) out.add(k);
      for (const o of p.options.values()) walk(o);
    }
  };
  walk(parse(message));
  return out;
}

export type Args = Record<string, unknown>;

const numberFormats = new Map<string, Intl.NumberFormat>();
const pluralRules = new Map<string, Intl.PluralRules>();

function numberFormat(locale: string, style?: string): Intl.NumberFormat {
  const key = `${locale}|${style ?? ''}`;
  let f = numberFormats.get(key);
  if (!f) {
    f = new Intl.NumberFormat(locale, style === 'percent' ? { style: 'percent' } : style === 'integer' ? { maximumFractionDigits: 0 } : {});
    numberFormats.set(key, f);
  }
  return f;
}

function rules(locale: string, type: 'cardinal' | 'ordinal'): Intl.PluralRules {
  const key = `${locale}|${type}`;
  let r = pluralRules.get(key);
  if (!r) { r = new Intl.PluralRules(locale, { type }); pluralRules.set(key, r); }
  return r;
}

/**
 * A string argument in a right-to-left language, isolated (U+2068 FIRST STRONG ISOLATE … U+2069
 * POP DIRECTIONAL ISOLATE): an XPath, a file name or the engine's English message keeps its own
 * order inside an Arabic sentence instead of being reordered with it. Invisible, and absent in a
 * left-to-right language.
 */
function isolated(text: string, rtl: boolean): string {
  return rtl && text !== '' ? `\u2068${text}\u2069` : text;
}

function render(parts: Part[], args: Args, locale: string, hash: number | undefined, rtl = directionOf(locale) === 'rtl'): string {
  let out = '';
  for (const p of parts) {
    if (typeof p === 'string') { out += p; continue; }
    if (p.arg === '#') { out += hash === undefined ? '#' : numberFormat(locale).format(hash); continue; }
    const value = args[p.arg];
    switch (p.type) {
      case undefined:
        out += value === undefined ? `{${p.arg}}` : typeof value === 'number' ? numberFormat(locale).format(value) : isolated(String(value), rtl);
        break;
      case 'number':
        out += typeof value === 'number' ? numberFormat(locale, p.style).format(value) : String(value);
        break;
      case 'date':
      case 'time': {
        const d = value instanceof Date ? value : new Date(value as string | number);
        const dateStyle = (p.style ?? 'medium') as 'short' | 'medium' | 'long' | 'full';
        out += Number.isNaN(d.getTime()) ? String(value) : new Intl.DateTimeFormat(locale, p.type === 'date' ? { dateStyle } : { timeStyle: dateStyle }).format(d);
        break;
      }
      case 'select': {
        const option = p.options.get(String(value)) ?? p.options.get('other')!;
        out += render(option, args, locale, hash, rtl);
        break;
      }
      case 'plural':
      case 'selectordinal': {
        const n = Number(value);
        const exact = p.options.get(`=${n}`);
        const option = exact ?? p.options.get(rules(locale, p.type === 'plural' ? 'cardinal' : 'ordinal').select(n)) ?? p.options.get('other')!;
        out += render(option, args, locale, n, rtl);
        break;
      }
    }
  }
  return out;
}

/** A message formatted in a locale; a message that does not parse is returned as it is. */
export function format(message: string, args: Args = {}, locale: string = current): string {
  let parts: Part[];
  try { parts = parsedOf(message); } catch { return message; }
  return render(parts, args, locale, undefined);
}

// ---------------------------------------------------------------- the lookup

const reported = new Set<string>();
const missingListeners = new Set<(key: string) => void>();

/** Called once per key that no file has, English included (the tests listen). */
export function onMissingKey(listener: (key: string) => void): () => void {
  missingListeners.add(listener);
  return () => missingListeners.delete(listener);
}

function report(key: string): void {
  if (reported.has(key)) return;
  reported.add(key);
  for (const l of missingListeners) l(key);
  if (missingListeners.size === 0) console.warn(`docx4j-editor: no message for "${key}"`);
}

function lookup(key: string, locale: string): { message: string; locale: string } | undefined {
  const own = catalogue.get(locale)?.[key];
  if (own !== undefined && own !== '') return { message: own, locale };
  const en = catalogue.get(SOURCE_LANGUAGE)?.[key];
  if (en !== undefined) return { message: en, locale: SOURCE_LANGUAGE };
  return undefined;
}

/** The message for `key` in the current language, English where the language lacks it, the key itself (reported once) where English does too. */
export function t(key: string, args?: Args): string {
  const found = lookup(key, current);
  if (!found) { report(key); return key; }
  // An English message falling back is formatted as English: its plural options are English's.
  return format(found.message, args, found.locale);
}

/**
 * A data-driven message with an English form of its own (a command's label, a plugin
 * perspective's tab): the file's message when a file has the key, else `fallback`. Not
 * reported when missing: a perspective or command registered outside the shipped set has
 * no key.
 */
export function tOr(key: string, fallback: string, args?: Args): string {
  const found = lookup(key, current);
  return found ? format(found.message, args, found.locale) : format(fallback, args, SOURCE_LANGUAGE);
}

/**
 * A message taken as plain text, not formatted (a command argument's description, which quotes
 * code with braces: `{ xpath: '…' }`): the file's text, else `fallback`.
 */
export function plainOr(key: string, fallback: string): string {
  const own = catalogue.get(current)?.[key];
  if (own !== undefined && own !== '') return own;
  return catalogue.get(SOURCE_LANGUAGE)?.[key] ?? fallback;
}

/** Whether any file has the key (English's file is the list of keys). */
export function hasMessage(key: string): boolean {
  return catalogue.get(SOURCE_LANGUAGE)?.[key] !== undefined;
}

/**
 * A message with elements among its arguments (a name in bold, a button inside a sentence): the
 * elements stand in as markers while the message is formatted, so a translation may move them.
 */
export function tNodes(key: string, args: Record<string, ReactNode>): ReactNode {
  const plain: Args = {};
  const nodes: ReactNode[] = [];
  for (const [name, value] of Object.entries(args)) {
    if (value === null || typeof value !== 'object') plain[name] = value;
    else { plain[name] = `\u0000${nodes.length}\u0000`; nodes.push(value); }
  }
  const text = t(key, plain);
  if (nodes.length === 0) return text;
  return text.split('\u0000').map((piece, i) => (i % 2 === 1 ? createElement(Fragment, { key: i }, nodes[Number(piece)]) : piece));
}

export type Translate = typeof t;

/** A number in the current language (`Intl.NumberFormat`). */
export function formatNumber(n: number): string {
  return numberFormat(current).format(n);
}

/** A date in the current language (`Intl.DateTimeFormat`). */
const DEFAULT_DATE: Intl.DateTimeFormatOptions = { dateStyle: 'medium', timeStyle: 'short' };
const dateFormatters = new Map<string, Intl.DateTimeFormat>();

/** The formatter for the current language and `options`, made once: a tracked change's mark view formats a date per run (ED-005 section 12.91). */
function dateFormatter(options: Intl.DateTimeFormatOptions): Intl.DateTimeFormat {
  const key = `${current}|${options === DEFAULT_DATE ? 'default' : JSON.stringify(options)}`;
  let f = dateFormatters.get(key);
  if (!f) { f = new Intl.DateTimeFormat(current, options); dateFormatters.set(key, f); }
  return f;
}

export function formatDate(d: Date | string | number, options: Intl.DateTimeFormatOptions = DEFAULT_DATE): string {
  const date = d instanceof Date ? d : new Date(d);
  return Number.isNaN(date.getTime()) ? String(d) : dateFormatter(options).format(date);
}

/** A list in the current language: "a, b and c" (`Intl.ListFormat`). */
export function formatList(items: readonly string[], type: 'conjunction' | 'disjunction' | 'unit' = 'conjunction'): string {
  try { return new Intl.ListFormat(current, { type }).format(items); } catch { return items.join(', '); }
}

/**
 * The lookup for a React component: it re-renders when the language changes. The function
 * returned changes identity with the language, so a memo that depends on it recomputes.
 */
export function useT(): Translate {
  const lang = useSyncExternalStore(onLanguageChange, language, language);
  return translators.get(lang) ?? makeTranslator(lang);
}

const translators = new Map<string, Translate>();

function makeTranslator(lang: string): Translate {
  const fn: Translate = (key, args) => t(key, args);
  translators.set(lang, fn);
  return fn;
}

/** The current language as a React value (the root's `lang` and `dir`). */
export function useLanguage(): string {
  return useSyncExternalStore(onLanguageChange, language, language);
}
