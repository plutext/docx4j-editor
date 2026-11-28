// SPDX-License-Identifier: GPL-3.0-or-later
//
// The extension's words (P29, and ED-014 section 4.7, P48): the manifest's in `package.nls.json`, and every
// `vscode.l10n.t('…')` in `src/` in `l10n/bundle.l10n.json`, English, the message its own key. Beside them, a draft of
// each for every one of the application's languages other than English, named by VS Code's own language id
// (`package.nls.<id>.json`, `l10n/bundle.l10n.<id>.json`): VS Code reads a bundle at its own id alone, so the application's
// `zh-Hans`, `zh-Hant` and `pt` are `zh-cn`, `zh-tw` and `pt-br` here.
//
//   node scripts/l10n.mjs           writes the English bundle from the sources
//   node scripts/l10n.mjs --check   checks every file (the root's `locales:check`, and `test/unit/l10n.test.ts`), and
//                                   prints the coverage per language; exit 1 on an error
//
// The check: the bundle is the sources' messages; every `%key%` of the manifest is in `package.nls.json` (`vsce` throws on
// one that is not) and every entry there is used; the drafts are the nineteen, named by VS Code's ids; and in each draft,
// `_status` marks it, no key English lacks, a message's placeholders English's (VS Code fills `{0}` and `{name}` alike and
// quotes nothing), a manifest value's backquoted code English's, and "docx4j editor" where English has it. A draft may lack
// a key, which VS Code shows in English; the coverage counts them.
import { existsSync, readdirSync, readFileSync, realpathSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const APP_LOCALES = join(root, '..', 'editor', 'src', 'locales');

/** The extension's sources: every `.ts` under `src/`, its folders included. */
function sourceFiles(dir = join(root, 'src')) {
  const out = [];
  for (const name of readdirSync(dir).sort()) {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) out.push(...sourceFiles(path));
    else if (name.endsWith('.ts')) out.push(path);
  }
  return out;
}

/**
 * In one source's text, each message passed to `l10n.t` as a single-quoted literal, and the line of each call whose message
 * is anything else (a double-quoted or template literal, a variable): the bundle cannot hold those, so the check refuses them.
 */
export function messagesIn(text) {
  const found = [];
  const unreadable = [];
  for (const m of text.matchAll(/\bl10n\.t\(\s*/g)) {
    const literal = /^'((?:[^'\\]|\\.)*)'/.exec(text.slice(m.index + m[0].length));
    if (literal) found.push(literal[1].replace(/\\(.)/g, '$1'));
    else unreadable.push(text.slice(0, m.index).split('\n').length);
  }
  return { found, unreadable };
}

/** The messages, in the order the sources first use them. */
export function messages() {
  const out = [];
  for (const file of sourceFiles()) {
    for (const message of messagesIn(readFileSync(file, 'utf8')).found) if (!out.includes(message)) out.push(message);
  }
  return out;
}

export const BUNDLE = join(root, 'l10n', 'bundle.l10n.json');
export const NLS = join(root, 'package.nls.json');

/** The product's name, never translated (section 4.7). */
export const PRODUCT = 'docx4j editor';

/** The application's tags VS Code names otherwise: its language packs' ids (P48). */
const VSCODE_IDS = { 'zh-Hans': 'zh-cn', 'zh-Hant': 'zh-tw', pt: 'pt-br' };

/** VS Code's id for one of the application's language tags. */
export function vscodeIdOf(tag) {
  return VSCODE_IDS[tag] ?? tag.toLowerCase();
}

/** VS Code's ids for the application's languages other than English, from its locale files. */
export function languages() {
  return readdirSync(APP_LOCALES).filter((f) => f.endsWith('.json') && f !== 'en.json').map((f) => vscodeIdOf(f.slice(0, -'.json'.length))).sort();
}

const placeholders = (text) => [...text.matchAll(/\{[^{}]*\}/g)].map((m) => m[0]).sort();
const code = (text) => [...text.matchAll(/`[^`]*`/g)].map((m) => m[0]).sort();
const count = (text, word) => text.split(word).length - 1;
const same = (a, b) => a.length === b.length && a.every((x, i) => x === b[i]);

/**
 * What is wrong with one draft against English, as sentences. `kind` is `manifest` for a `package.nls` file, whose values
 * may hold backquoted code, or `bundle`.
 */
export function draftProblems(name, english, draft, kind) {
  const out = [];
  if (typeof draft._status !== 'string') out.push(`${name}: no _status marking it a draft`);
  for (const [key, value] of Object.entries(draft)) {
    if (key === '_status') continue;
    if (!(key in english)) { out.push(`${name}: "${key}" is not in English`); continue; }
    if (typeof value !== 'string' || value.trim() === '') { out.push(`${name}: "${key}" is not a message`); continue; }
    const en = english[key];
    if (!same(placeholders(value), placeholders(en))) out.push(`${name}: "${key}" has the placeholders ${placeholders(value).join(' ') || 'none'}, where English has ${placeholders(en).join(' ') || 'none'}`);
    if (kind === 'manifest' && !same(code(value), code(en))) out.push(`${name}: "${key}" changes the code English quotes (${code(en).join(' ') || 'none'})`);
    if (count(value, PRODUCT) < count(en, PRODUCT)) out.push(`${name}: "${key}" lacks "${PRODUCT}", which is never translated`);
  }
  return out;
}

const read = (path) => JSON.parse(readFileSync(path, 'utf8'));

/** Every file checked: the errors, and per language the keys each draft has of English's. */
export function checkWords() {
  const errors = [];
  const bundle = read(BUNDLE);
  const nls = read(NLS);
  const sources = messages();
  const keys = Object.keys(bundle);
  for (const m of sources) if (!(m in bundle)) errors.push(`bundle.l10n.json lacks a message the sources use: "${m}" (run node scripts/l10n.mjs)`);
  for (const k of keys) if (!sources.includes(k)) errors.push(`bundle.l10n.json holds a message no source uses: "${k}" (run node scripts/l10n.mjs)`);
  for (const [k, v] of Object.entries(bundle)) if (v !== k) errors.push(`bundle.l10n.json: "${k}" is not its own message`);
  for (const file of sourceFiles()) {
    for (const line of messagesIn(readFileSync(file, 'utf8')).unreadable) errors.push(`${file.slice(root.length + 1)}:${line}: l10n.t is given a message that is not a single-quoted literal, which the bundle cannot hold`);
  }
  const manifest = readFileSync(join(root, 'package.json'), 'utf8');
  const used = new Set([...manifest.matchAll(/"%([^%"]+)%"/g)].map((m) => m[1]));
  for (const k of used) if (!(k in nls)) errors.push(`package.nls.json lacks %${k}%, which the manifest uses (vsce throws on it)`);
  for (const k of Object.keys(nls)) if (!used.has(k)) errors.push(`package.nls.json holds "${k}", which the manifest does not use`);

  const ids = languages();
  const named = (dir, pattern) => readdirSync(dir).map((f) => pattern.exec(f)?.[1]).filter((id) => id !== undefined);
  for (const id of named(root, /^package\.nls\.(.+)\.json$/)) if (!ids.includes(id)) errors.push(`package.nls.${id}.json is not named by VS Code's id for one of the application's languages`);
  for (const id of named(join(root, 'l10n'), /^bundle\.l10n\.(.+)\.json$/)) if (!ids.includes(id)) errors.push(`l10n/bundle.l10n.${id}.json is not named by VS Code's id for one of the application's languages`);

  const coverage = [];
  for (const id of ids) {
    const row = { id, manifest: 0, bundle: 0, draft: true };
    for (const [kind, path, english] of [['manifest', join(root, `package.nls.${id}.json`), nls], ['bundle', join(root, 'l10n', `bundle.l10n.${id}.json`), bundle]]) {
      if (!existsSync(path)) { errors.push(`${path.slice(root.length + 1)} is missing`); continue; }
      const draft = read(path);
      errors.push(...draftProblems(path.slice(root.length + 1), english, draft, kind));
      row[kind] = Object.keys(draft).filter((k) => k !== '_status' && k in english).length;
    }
    coverage.push(row);
  }
  return { errors, coverage, total: { manifest: Object.keys(nls).length, bundle: keys.length } };
}

/** Run as a script, not imported: Node takes the entry module's real path, and `argv[1]` as typed (a symlinked checkout). */
function isMain() {
  try { return realpathSync(process.argv[1]) === realpathSync(fileURLToPath(import.meta.url)); } catch { return false; }
}

if (isMain()) {
  if (process.argv.includes('--check')) {
    const { errors, coverage, total } = checkWords();
    console.log(`The extension's words: ${total.manifest} in the manifest, ${total.bundle} messages`);
    for (const row of coverage) console.log(`  ${row.id.padEnd(6)} manifest ${row.manifest}/${total.manifest}, messages ${row.bundle}/${total.bundle} (draft)`);
    for (const e of errors) console.error(`l10n: ${e}`);
    process.exit(errors.length > 0 ? 1 : 0);
  } else {
    const bundle = Object.fromEntries(messages().map((m) => [m, m]));
    writeFileSync(BUNDLE, `${JSON.stringify(bundle, null, 2)}\n`);
    console.error(`l10n: ${Object.keys(bundle).length} messages -> ${BUNDLE}`);
  }
}
