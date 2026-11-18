// SPDX-License-Identifier: GPL-3.0-or-later
//
// Replace and Replace all (ED-005 section 3.3): the matcher of `find.mts` over the command's
// state, each one transaction and so one undo step, co-editing like any edit; a hit is found in
// the state the command is dispatched against, so in a room a peer's concurrent edit either lands
// before (and is searched) or after (and is kept). Their `api` lines are Office JS where Office
// JS has the shape (`search`, then `insertText(…, 'Replace')` from the last hit to the first,
// since a replacement of another length moves every later hit in its paragraph); a replacement
// with group references (`$1`, Word's `\1`) is the engine's `replaceText` extension, and a
// regular expression is the engine's `matchRegExp` extension of `SearchOptions` (CR-002 section 26).
// A hit holding a paragraph mark (Word's wildcards across one, `^p`) is never replaced: the engine's
// `replaceText` refuses before changing anything, and so do both commands (`HOLDS_MARK`).
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import { Finder, hasGroupReference, hitSelection, searchOptionsOf, type FindOptions, type Hit } from '../find.mjs';
import { register, sessionOf } from './spec.mjs';
import { stateFor } from './run.mjs';
import type { EditorDocument } from '../session.mjs';
import type { Story } from '../stories.mjs';

export interface ReplaceArgs extends FindOptions {
  /** What to find. */
  text: string;
  /** What to put in its place; `$1` (a regular expression) or `\1` (wildcards) for a group. */
  replacement: string;
}

/** The transaction meta a replace carries: how many hits it replaced (the bar's status line). */
export const REPLACED = 'docx4j-find-replaced';

export function replacedCountOf(tr: Transaction): number | undefined {
  return tr.getMeta(REPLACED) as number | undefined;
}

/** The query as a finder over the state's session, or undefined for an empty or invalid one (the bar says why). */
export function finderFor(state: EditorState, args: ReplaceArgs): Finder | undefined {
  if (!args.text) return undefined;
  try {
    const session = sessionOf(state);
    return new Finder(args.text, args, session?.ctx.ids, { resolver: session?.pkg.propertyResolverOrUndefined });
  } catch {
    return undefined;
  }
}

/** The options as a line's argument: nothing when none is set. */
function optionsArg(args: FindOptions): string {
  const o = searchOptionsOf(args);
  return Object.keys(o).length === 0 ? '' : `, ${JSON.stringify(o)}`;
}

const ARGS: Record<string, string> = {
  text: 'what to find',
  replacement: 'what to put in its place ($1 or \\1 for a group)',
  matchCase: 'boolean: match case',
  matchWholeWord: 'boolean: whole words only',
  matchWildcards: "boolean: Word's wildcards",
  matchRegExp: 'boolean: a regular expression (not with wildcards)',
};
const BOOLEAN = ['false', 'true'] as const;
const CHOICES = { matchCase: BOOLEAN, matchWholeWord: BOOLEAN, matchWildcards: BOOLEAN, matchRegExp: BOOLEAN };

/**
 * Word's Replace: the selection, when it is a hit, replaced and the next hit selected; otherwise
 * the next hit from the selection selected (the first press finds). Wraps at the end.
 */
register<ReplaceArgs>({
  name: 'find.replace',
  label: 'Replace',
  args: ARGS,
  choices: CHOICES,
  api: (args) => {
    if (!args.text || hasGroupReference(args.replacement, args)) return undefined;
    return `const r = selection.range; const hit = r && r.paragraph.search(${JSON.stringify(args.text)}${optionsArg(args)}).find((h) => h.start === r.start && h.end === r.end); if (hit) hit.insertText(${JSON.stringify(args.replacement)}, 'Replace');`;
  },
  run: (args) => (state, dispatch) => {
    const finder = finderFor(state, args);
    if (!finder) return false;
    const hits = finder.all(state.doc);
    if (hits.length === 0) return false;
    const { from, to } = state.selection;
    const current = selectedHit(state, hits);
    if (current?.mark) return false;
    const tr = state.tr;
    let after = from;
    if (current) {
      after = current.from + finder.replacementFor(current, args.replacement).length;
      tr.setMeta(REPLACED, finder.replace(tr, [current], args.replacement));
    } else if (hits.some((h) => h.from === from && h.to === to)) return false;
    // The next hit after the replacement (or from the selection), searched in the document as it now is; the first when none follows.
    const fresh = current ? finder.all(tr.doc) : hits;
    const next = fresh.find((h) => h.from >= after) ?? fresh[0];
    if (next) tr.setSelection(hitSelection(tr.doc, next));
    else tr.setSelection(TextSelection.create(tr.doc, after));
    dispatch?.(tr.scrollIntoView());
    return true;
  },
});

/** The hit the selection is: one whose selection (`hitSelection`, across a mark) is the state's. */
export function selectedHit(state: EditorState, hits: readonly Hit[]): Hit | undefined {
  const { from, to } = state.selection;
  return hits.find((h) => {
    if (h.from !== from) return false;
    if (!h.mark) return h.to === to;
    const sel = hitSelection(state.doc, h);
    return sel.from === from && sel.to === to;
  });
}

/** Every hit replaced, last to first, in one transaction. */
register<ReplaceArgs>({
  name: 'find.replaceAll',
  label: 'Replace all',
  args: ARGS,
  choices: CHOICES,
  api: (args) => {
    if (!args.text) return undefined;
    const q = JSON.stringify(args.text);
    const r = JSON.stringify(args.replacement);
    if (hasGroupReference(args.replacement, args)) return `body.replaceText(${q}, ${r}${optionsArg(args)});`;
    return `const hits = body.search(${q}${optionsArg(args)}); for (let i = hits.length - 1; i >= 0; i--) hits[i].insertText(${r}, 'Replace');`;
  },
  run: (args) => (state, dispatch) => {
    const finder = finderFor(state, args);
    if (!finder) return false;
    const hits = finder.all(state.doc);
    if (hits.length === 0 || hits.some((h) => h.mark)) return false;
    const tr = state.tr;
    tr.setMeta(REPLACED, finder.replace(tr, hits, args.replacement));
    dispatch?.(tr);
    return true;
  },
});

/**
 * `find.replaceAll` over a note or a header, headless, without closing the undo step as
 * `runCommand` does, so the bar's Replace all over the body and then the stories is one undo step
 * (ED-005 section 3.7): the manager merges transactions within its capture window. The number
 * replaced.
 */
export function replaceAllInStory(session: EditorDocument, story: Story, args: ReplaceArgs): number {
  const state = stateFor(session, undefined, undefined, story);
  const finder = finderFor(state, args);
  if (!finder) return 0;
  const hits = finder.all(state.doc);
  if (hits.length === 0 || hits.some((h) => h.mark)) return 0;
  const tr = state.tr;
  finder.replace(tr, hits, args.replacement);
  session.replaceStory(story, tr.doc);
  return hits.length;
}
