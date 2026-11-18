// SPDX-License-Identifier: GPL-3.0-or-later
//
// Optional breaks (ED-007 section 8.1 item 13; Jason, 2026-10-15): a zero-width space, U+200B, Word's "No-Width
// Optional Break", is a character a line may break at and nothing is drawn for. The fit offers to put one after an
// inner period and before an inner capital of the words it had to break mid-word, and the caret can insert one.
import { register } from './spec.mjs';
import { tableAround } from './tables.mjs';

/** Word's No-Width Optional Break. */
export const OPTIONAL_BREAK = '​';

/**
 * A word for breaking: a run of characters with no blank, hyphen, dash, soft hyphen or optional break, since a line may
 * break after any of those. Two characters at least; a global, sticky-free pattern to `matchAll` over.
 */
export const WORD_RUN = /[^\s\-­​‐‒-―]{2,}/gu;

/**
 * The offsets in a word where an optional break goes: after an inner period before a letter ("Gamma.|Den") and before a
 * capital after a lower-case letter ("Eta|Eastern"). Empty when the word has no such place.
 */
export function breakOffsetsIn(word: string): number[] {
  const out: number[] = [];
  const chars = Array.from(word);
  let offset = 0;
  for (let i = 0; i < chars.length; i++) {
    const prev = chars[i - 1];
    const cur = chars[i]!;
    if (i > 0 && i < chars.length - 1 + 1 && ((prev === '.' && /\p{L}/u.test(cur)) || (prev !== undefined && /\p{Ll}/u.test(prev) && /\p{Lu}/u.test(cur)))) out.push(offset);
    offset += cur.length;
  }
  return out;
}

register<void>({
  name: 'range.insertOptionalBreak',
  label: 'Insert optional break',
  typing: true,
  api: () => `selection.range.insertText(${JSON.stringify(OPTIONAL_BREAK)}, 'Replace');`,
  run: () => (state, dispatch) => {
    dispatch?.(state.tr.insertText(OPTIONAL_BREAK).scrollIntoView());
    return true;
  },
});

register<{ words: string[] }>({
  name: 'table.breakWords',
  label: 'Break words at periods and capitals',
  args: { words: 'the words to break, as they read in the table at the selection; an optional break goes after each inner period and before each inner capital' },
  isEnabled: (state) => tableAround(state) !== undefined,
  run: ({ words }) => (state, dispatch) => {
    const found = tableAround(state);
    if (!found) return false;
    const want = new Set(words.filter((w) => breakOffsetsIn(w).length > 0));
    if (want.size === 0) return false;
    const inserts: number[] = [];
    state.doc.nodesBetween(found.pos, found.pos + found.node.nodeSize, (node, pos) => {
      if (!node.isText) return true;
      for (const m of (node.text ?? '').matchAll(WORD_RUN)) {
        if (!want.has(m[0])) continue;
        for (const offset of breakOffsetsIn(m[0])) inserts.push(pos + m.index + offset);
      }
      return false;
    });
    if (inserts.length === 0) return false;
    const tr = state.tr;
    for (const at of inserts.sort((a, b) => b - a)) tr.insertText(OPTIONAL_BREAK, at);
    dispatch?.(tr);
    return true;
  },
});
