// SPDX-License-Identifier: GPL-3.0-or-later
//
// The characters XML 1.0 excludes (ED-005 section 12.114). No escape can write one, so a part holding one is not
// well-formed: the engine refuses to save it and a browser refuses to read it. They arrive in text from elsewhere,
// since Office JS's text carries breaks and hyphens as such characters and a text file can hold a form feed, and none
// may stay in the edit layer, or the document can no longer be saved. Each is written as what it stands for and the
// rest are dropped, as the engine's writers do and as Word does (core-ts CR-002 section 60, check 55).
//
// `recordTransaction` applies the rule to what a transaction inserts, so a paste, a drop, a replacement and a script
// are all covered on every surface. Text that never passes through a transaction, Markdown and a comment's text, is
// made into a run's items by `runItemsFor`.
//
// The rule is the engine's `officeText.mts` (core-ts CR-002 section 60), whose `dropExcluded` and `officeItems` make a run's
// items since core-ts 0.7.0 (UPSTREAM-25, closed). The rule over a transaction's steps is the editor's own, the engine
// having no such thing; it keeps the two hyphens as characters, as `runItemsFor` does before the engine's writer.
import { dropExcluded, officeItems, type Element } from '@docx4j/core-ts';
import type { Mark } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';
import { ReplaceAroundStep, ReplaceStep, type Step } from 'prosemirror-transform';

/** Any character outside XML 1.0's `Char` production; with the `u` flag an unpaired surrogate is one. */
const EXCLUDED = /[^\t\n\r\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/u;
const EVERY_EXCLUDED = new RegExp(EXCLUDED.source, 'gu');

/** The characters that are breaks in Office JS's text. */
const BREAKS: ReadonlyMap<string, 'line' | 'page' | 'column'> = new Map([['\u000B', 'line'], ['\u000C', 'page'], ['\u000E', 'column']]);

/**
 * The two hyphens of Office JS's text, as the characters Word shows them as. The engine writes `w:noBreakHyphen` and
 * `w:softHyphen`; the edit layer holds those only as opaque atoms of a source element, which a new one does not have.
 */
const HYPHENS: ReadonlyMap<string, string> = new Map([['\u001E', '\u2011'], ['\u001F', '\u00AD']]);

/** A text as the edit layer comes to hold it, a break a line feed: what to compare with a text that may hold such characters. */
export function heldText(text: string): string {
  return text.replace(EVERY_EXCLUDED, (character) => (BREAKS.has(character) ? '\n' : HYPHENS.get(character) ?? ''));
}

/**
 * A run's items for a text, as the engine's writers make them: a break for each break character, the rest of the
 * characters XML 1.0 excludes dropped. The two hyphens are written as the characters Word shows them as first, as a
 * transaction writes them, where the engine writes `w:noBreakHyphen` and `w:softHyphen`: the edit layer would hold those
 * as opaque atoms, reported among the content it cannot edit.
 */
export function runItemsFor(text: string): Element[] {
  return officeItems(dropExcluded(text.replace(/[\u001E\u001F]/g, (c) => HYPHENS.get(c)!)));
}

/** Whether a step inserts such a character. It reads the step's own content, so typing costs a character's test. */
function inserts(step: Step): boolean {
  if (!(step instanceof ReplaceStep) && !(step instanceof ReplaceAroundStep)) return false;
  let found = false;
  step.slice.content.descendants((node) => {
    found ||= node.isText && EXCLUDED.test(node.text!);
    return !found;
  });
  return found;
}

/**
 * What a transaction inserted, with these characters written as above, in the same transaction: one undo step, and
 * tracked with what it came in. Only the characters are replaced, since replacing a text node would mark the text
 * around them as changed. The whole of each text node is read, as half of a surrogate pair at a step's edge is not one.
 */
export function writeExcludedCharacters(tr: Transaction): void {
  const found = new Map<number, { character: string; marks: readonly Mark[] }>();
  tr.steps.forEach((step, i) => {
    if (!inserts(step)) return;
    const later = tr.mapping.slice(i + 1);
    step.getMap().forEach((_oldStart, _oldEnd, from, to) => {
      tr.doc.nodesBetween(later.map(from, -1), later.map(to, 1), (node, pos) => {
        if (!node.isText) return;
        for (const match of node.text!.matchAll(EVERY_EXCLUDED)) found.set(pos + match.index, { character: match[0], marks: node.marks });
      });
    });
  });
  const { schema } = tr.doc.type;
  // From the last to the first, so that each position still stands.
  for (const [at, { character, marks }] of [...found].sort(([a], [b]) => b - a)) {
    const kind = BREAKS.get(character);
    const hyphen = HYPHENS.get(character);
    if (kind !== undefined) tr.replaceWith(at, at + 1, schema.nodes['hardBreak']!.create({ kind }));
    else if (hyphen !== undefined) tr.replaceWith(at, at + 1, schema.text(hyphen, marks));
    else tr.delete(at, at + 1);
  }
}
