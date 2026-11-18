// SPDX-License-Identifier: GPL-3.0-or-later
//
// The portability lint for pasted Office JS (ED-001 section 3.5's caveat; ED-003 section 7.3):
// code that works in the console is not guaranteed to work in Word, because Office JS requires
// `load()` and `context.sync()` before a property is read and this package does not. A pure
// function over the script's text finds, inside a `Word.run` callback, the reads of loadable
// properties on a variable with no earlier `load('...')` naming that property (or `items/property`
// on the collection a loop variable came from), followed by a sync; and it offers the `load` and
// `sync` lines to insert. Heuristic over tokens, not a type-directed analysis: it catches the shape
// add-in code has (`paragraphs.load('items/text'); await context.sync(); for (const p of
// paragraphs.items) ... p.text`) and nothing more.

/** The properties Office JS makes the caller load before reading; a method call needs no load, and `items` is implied by any `items/x` load. */
export const LOADABLE = new Set([
  'text', 'style', 'styleBuiltIn', 'alignment', 'isNullObject', 'leftIndent', 'rightIndent', 'firstLineIndent',
  'spaceBefore', 'spaceAfter', 'lineSpacing', 'outlineLevel', 'tag', 'title', 'value', 'font', 'bold', 'italic', 'underline',
  'size', 'name', 'color', 'highlightColor', 'type', 'id', 'hyperlink', 'values', 'rowCount', 'isListItem', 'changeTrackingMode',
]);

export interface LintWarning {
  /** 1-based line of the read. */
  line: number;
  object: string;
  property: string;
  message: string;
}

export interface LintInsertion {
  /** 1-based line before which the text goes. */
  line: number;
  text: string;
}

export interface LintResult {
  warnings: LintWarning[];
  /** The `load` and `sync` lines that would satisfy the earliest unloaded reads, when any. */
  insertion?: LintInsertion;
}

interface Loaded {
  /** Properties loaded on the object, with the line of the load. */
  props: Map<string, number>;
}

const CONTEXT_RE = /Word\.run\(\s*(?:async\s*)?\(?\s*(\w+)\s*\)?\s*=>/;
const LOAD_RE = /([\w.]+)\.load\(\s*(['"`])([^'"`]*)\2\s*\)/g;
const ITEMS_LOOP_RE = /for\s*\(\s*(?:const|let|var)\s+(\w+)\s+of\s+([\w.]+?)(?:\.items)?\s*\)/g;
const ITEMS_INDEX_RE = /(?:const|let|var)\s+(\w+)\s*=\s*([\w.]+?)(?:\.items\s*\[[^\]]*\]|\.getFirst(?:OrNullObject)?\(\)|\.getLast(?:OrNullObject)?\(\))/g;
const READ_RE = /(?<![\w.])(\w+)\.(\w+)\b(?!\s*\()/g;

export function lintOfficeJs(code: string): LintResult {
  const m = CONTEXT_RE.exec(code);
  if (!m) return { warnings: [] };
  const context = m[1]!;
  const lines = code.split('\n');
  const start = code.slice(0, m.index).split('\n').length - 1;   // the line of Word.run
  const loaded = new Map<string, Loaded>();
  /** A loop or index variable and the collection it came from. */
  const derived = new Map<string, string>();
  const syncLines: number[] = [];
  const warnings: LintWarning[] = [];
  /** Unloaded reads by object, for the insertion. */
  const unloaded = new Map<string, { line: number; props: Set<string> }>();

  const propsLoadedOn = (obj: string, prop: string): number | undefined => {
    const direct = loaded.get(obj)?.props.get(prop);
    if (direct !== undefined) return direct;
    const from = derived.get(obj);
    if (from) return loaded.get(from)?.props.get(`items/${prop}`);
    return undefined;
  };

  for (let i = start; i < lines.length; i++) {
    const line = lines[i]!;
    const n = i + 1;
    if (new RegExp(`\\b${context}\\.sync\\(`).test(line)) syncLines.push(n);
    for (const lm of line.matchAll(LOAD_RE)) {
      const entry = loaded.get(lm[1]!) ?? { props: new Map() };
      for (const p of lm[3]!.split(',').map((x) => x.trim()).filter(Boolean)) {
        entry.props.set(p, n);
        if (p.startsWith('items/')) entry.props.set('items', n);
      }
      loaded.set(lm[1]!, entry);
    }
    for (const dm of line.matchAll(ITEMS_LOOP_RE)) derived.set(dm[1]!, dm[2]!);
    for (const dm of line.matchAll(ITEMS_INDEX_RE)) derived.set(dm[1]!, dm[2]!);
    for (const rm of line.matchAll(READ_RE)) {
      const [whole, obj, prop] = rm as unknown as [string, string, string];
      if (!LOADABLE.has(prop) || obj === context || obj === 'Word' || obj === 'console' || obj === 'Math') continue;
      // A write needs no load; a `.load(` or `.sync(` call is not a read.
      const after = line.slice(rm.index! + whole.length);
      if (/^\s*=[^=]/.test(after)) continue;
      if (/^\s*\./.test(after) && /^\s*\.(load|sync)\(/.test(after)) continue;
      const loadedAt = propsLoadedOn(obj, prop);
      if (loadedAt !== undefined && syncLines.some((s) => s >= loadedAt && s <= n)) continue;
      const message = loadedAt === undefined
        ? `line ${n}: \`${obj}.${prop}\` is read without \`load('${derived.has(obj) ? 'items/' : ''}${prop}')\`: fine here, fails in Word`
        : `line ${n}: \`${obj}.${prop}\` is read without \`await ${context}.sync()\` after its load: fine here, fails in Word`;
      warnings.push({ line: n, object: obj, property: prop, message });
      if (loadedAt === undefined) {
        const target = derived.get(obj) ?? obj;
        const key = derived.has(obj) ? `items/${prop}` : prop;
        const u = unloaded.get(target) ?? { line: n, props: new Set() };
        u.props.add(key);
        unloaded.set(target, u);
      }
    }
  }
  const result: LintResult = { warnings };
  if (unloaded.size > 0) {
    const [target, u] = [...unloaded.entries()].sort((a, b) => a[1].line - b[1].line)[0]!;
    // Before the loop that binds a derived variable, else before the read itself.
    let at = u.line;
    for (let i = start; i < u.line - 1; i++) {
      const l = lines[i]!;
      if ([...l.matchAll(ITEMS_LOOP_RE), ...l.matchAll(ITEMS_INDEX_RE)].some((dm) => dm[2] === target)) { at = i + 1; break; }
    }
    const indent = /^\s*/.exec(lines[at - 1] ?? '')?.[0] ?? '';
    result.insertion = { line: at, text: `${indent}${target}.load('${[...u.props].join(',')}');\n${indent}await ${context}.sync();\n` };
  }
  return result;
}
