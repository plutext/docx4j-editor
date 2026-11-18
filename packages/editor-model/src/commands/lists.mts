// SPDX-License-Identifier: GPL-3.0-or-later
//
// Lists, the attribute half (ED-003 section 5.3, step C1): the commands that change `numPr` on
// the selected paragraphs' nodes and nothing else, so each is one transaction, one Yjs update
// and one undo step at any document size. Where a paragraph's list comes from its style, the
// emulator's `numRefFor` says so over the same `{ pStyle, numPr }` the labels use, and a level
// change writes the style's `w:numId` beside the level, as the engine's `listItem.level` does.
// The definition commands (start, restart) are step C2, over the engine's paragraph-free verbs.
import type { Command, EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type { Emulator, NumberingDefinitions } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { IndentAttrs, NumPrAttrs, ParagraphAttrs } from '../schema.mjs';
import type { EditorDocument } from '../session.mjs';
import { paragraphsIn, updateParagraphs } from './paragraph.mjs';
import { register, sessionOf } from './spec.mjs';

/** Word's nine levels. */
export const LEVELS = 9;

export interface Numbering {
  emulator: Emulator;
  definitions: NumberingDefinitions;
}

/** The numbering part's emulator and definitions; undefined where the document has no numbering part, or it has not been read. */
export function numberingParts(session: EditorDocument | undefined): Numbering | undefined {
  return session ? numberingOfPart(session.main) : undefined;
}

/** The same, from a main document part: the preview's, whose package is a clone with lists of its own (ED-005 section 12.117). */
export function numberingOfPart(main: { numberingDefinitionsPart?: { getEmulator(): Emulator; definitions: NumberingDefinitions } | undefined }): Numbering | undefined {
  const part = main.numberingDefinitionsPart;
  if (!part) return undefined;
  try {
    return { emulator: part.getEmulator(), definitions: part.definitions };
  } catch {
    return undefined;   // the part has not been read (a session opened without the resolver)
  }
}

/** The same, from an editor state carrying the session. */
export function numberingOf(state: EditorState): Numbering | undefined {
  return numberingParts(sessionOf(state));
}

/**
 * The indent a list paragraph lays out with (ED-003 section 5.4): the level's `w:ind` (the
 * override level's, else the abstract level's, else the linked style's: the definitions'
 * `indOf`) with the paragraph's own `ind` over it field by field, as the resolver folds them.
 * Null where the paragraph is not a list item or its level states no indent, in which case
 * the paragraph's own attribute is all there is (the node view applies that as before).
 */
export function levelIndentOf(attrs: ParagraphAttrs, numbering: Numbering | undefined): IndentAttrs | null {
  if (!numbering) return null;
  const ref = listRefOf(attrs, numbering.emulator);
  if (!ref) return null;
  const ind = numbering.definitions.indOf(ref.numId!, String(ref.ilvl));
  if (!ind) return null;
  const own = attrs.ind;
  const level: IndentAttrs = { start: ind.start ?? ind.left ?? null, end: ind.end ?? ind.right ?? null, firstLine: ind.firstLine ?? null, hanging: ind.hanging ?? null };
  if (!own) return level;
  const out: IndentAttrs = { start: own.start ?? level.start, end: own.end ?? level.end, firstLine: level.firstLine, hanging: level.hanging };
  // The first-line and hanging indents are one setting with two signs: the paragraph's own replaces the level's.
  if (own.firstLine !== null || own.hanging !== null) { out.firstLine = own.firstLine; out.hanging = own.hanging; }
  return out;
}

/** The only `w:pPr` the emulator's `numRefFor` reads, from a paragraph node's attributes (ED-003 section 5.4). */
export function numberingPPrOf(attrs: Pick<ParagraphAttrs, 'pStyle' | 'numPr'>): wml.PPr {
  const pPr: wml.PPr = { TYPE_NAME: 'org_docx4j_wml.PPr' };
  if (attrs.pStyle !== null) pPr.pStyle = { val: attrs.pStyle };
  if (attrs.numPr) {
    pPr.numPr = { ilvl: { val: attrs.numPr.ilvl } };
    if (attrs.numPr.numId !== null) pPr.numPr.numId = { val: Number(attrs.numPr.numId) };
  }
  return pPr;
}

/** Where a paragraph's numbering resolves to (its own `w:numPr`, else its style's), or undefined when it is not a list item. */
export function listRefOf(attrs: ParagraphAttrs, emulator: Emulator): NumPrAttrs | undefined {
  const ref = emulator.numRefFor(numberingPPrOf(attrs));
  if (ref.notNumbered || ref.numId === undefined) return undefined;
  return { numId: ref.numId, ilvl: Number(ref.ilvl ?? '0') };
}

/** Whether the paragraph's style numbers it: the reference over the style alone. */
function styleNumbers(attrs: ParagraphAttrs, emulator: Emulator): boolean {
  return listRefOf({ ...attrs, numPr: null }, emulator) !== undefined;
}

/**
 * The level the list defines nearest to `level`, at or below it (`levelExistences`); 0 to 8
 * when the definition cannot be read; undefined when the list defines no level at or below.
 */
export function clampLevel(numId: string, level: number, definitions: NumberingDefinitions | undefined): number | undefined {
  let n = Math.max(0, Math.min(LEVELS - 1, Math.trunc(level)));
  const list = definitions?.list(numId);
  if (!list) return n;
  while (n >= 0 && !list.levelExists(n)) n--;
  return n < 0 ? undefined : n;
}

/** The selected paragraphs that are list items, with where each resolves to. */
function listItemsIn(state: EditorState): { node: PMNode; pos: number; ref: NumPrAttrs }[] {
  const numbering = numberingOf(state);
  if (!numbering) return [];
  const out: { node: PMNode; pos: number; ref: NumPrAttrs }[] = [];
  for (const p of paragraphsIn(state)) {
    const ref = listRefOf(p.node.attrs as ParagraphAttrs, numbering.emulator);
    if (ref) out.push({ ...p, ref });
  }
  return out;
}

const someListItem = (state: EditorState): boolean => listItemsIn(state).length > 0;

// ---------------------------------------------------------------------------------------------
// attach, detach
// ---------------------------------------------------------------------------------------------

register<{ listId: string | number; level?: number }>({
  name: 'paragraph.attachToList',
  label: 'Attach to list',
  args: { listId: "the list's w:numId", level: 'the level, 0 to 8; 0 by default' },
  run: ({ listId, level }) => (state, dispatch) => {
    const numId = String(listId);
    const numbering = numberingOf(state);
    // Refused when the list does not exist: a w:numId naming no w:num numbers nothing.
    if (!numbering?.definitions.list(numId)) return false;
    const ilvl = clampLevel(numId, level ?? 0, numbering.definitions);
    if (ilvl === undefined) return false;
    return updateParagraphs(() => ({ numPr: { numId, ilvl } }))(state, dispatch);
  },
});

/**
 * Office JS `paragraph.detachFromList()`: the paragraph's own `w:numPr` goes; where its style
 * still numbers it, `w:numId` 0 takes its place, which is how Word switches an inherited list
 * off (the engine's verb, ECMA-376 17.9.18).
 */
register<void>({
  name: 'paragraph.detachFromList',
  label: 'Detach from list',
  run: () => (state, dispatch) => {
    const emulator = numberingOf(state)?.emulator;
    return updateParagraphs((a) => {
      const numPr: NumPrAttrs | null = emulator && styleNumbers(a, emulator) ? { numId: '0', ilvl: 0 } : null;
      return a.numPr === null && numPr === null ? null : { numPr };
    })(state, dispatch);
  },
  isEnabled: (state) => someListItem(state) || paragraphsIn(state).some((p) => (p.node.attrs as ParagraphAttrs).numPr !== null),
});

// ---------------------------------------------------------------------------------------------
// the level: set, indent, outdent
// ---------------------------------------------------------------------------------------------

/** The `numPr` that puts a list item at `level`, or undefined when the list has no such level; null detaches (outdent at 0). */
function levelPatch(ref: NumPrAttrs, level: number, definitions: NumberingDefinitions): NumPrAttrs | null | undefined {
  if (level < 0) return null;
  const ilvl = clampLevel(ref.numId!, level, definitions);
  return ilvl === undefined ? undefined : { numId: ref.numId, ilvl };
}

/** Change the level of every selected list item by `read`; true when the selection held one. */
function changeLevels(state: EditorState, dispatch: Parameters<Command>[1], read: (ref: NumPrAttrs) => number): boolean {
  const numbering = numberingOf(state);
  if (!numbering) return false;
  const items = listItemsIn(state);
  if (items.length === 0) return false;
  const tr = state.tr;
  let any = false;
  for (const { node, pos, ref } of items) {
    const attrs = node.attrs as ParagraphAttrs;
    const numPr = levelPatch(ref, read(ref), numbering.definitions);
    if (numPr === undefined) continue;
    // Where the paragraph resolves already (a clamped level equal to its own): nothing to write.
    if (numPr !== null && numPr.numId === ref.numId && numPr.ilvl === ref.ilvl) continue;
    // Outdent at level 0 detaches, as Word does; a style-numbered paragraph is switched off with numId 0.
    const value = numPr === null && styleNumbers(attrs, numbering.emulator) ? { numId: '0', ilvl: 0 } : numPr;
    if (value === null && attrs.numPr === null) continue;
    tr.setNodeMarkup(pos, undefined, { ...attrs, numPr: value });
    any = true;
  }
  if (any) dispatch?.(tr);
  return true;
}

register<{ value: number }>({
  name: 'listItem.level',
  label: 'List level',
  args: { value: 'the level, 0 to 8, clamped to the levels the list defines' },
  run: ({ value }) => (state, dispatch) => changeLevels(state, dispatch, () => Math.max(0, value)),
  isEnabled: someListItem,
  isActive: (state, args) => args !== undefined && (() => { const items = listItemsIn(state); return items.length > 0 && items.every((i) => i.ref.ilvl === args.value); })(),
});

/**
 * Tab at the start of a list item (`atStart`), and the toolbar's Increase Indent on one: the
 * level plus one, clamped to the levels the list defines. With `atStart` the command applies
 * only to an empty selection at the start of a list item's text, so that the keymap can bind
 * Tab here before the table bindings and `insertTab`.
 */
const atStartOfListItem = (state: EditorState): boolean => {
  const { $from, empty } = state.selection;
  return empty && $from.parentOffset === 0 && $from.parent.type.name === 'paragraph' && someListItem(state);
};

register<{ atStart?: boolean } | void>({
  name: 'list.indent',
  label: 'Increase list level',
  args: { atStart: 'true: only at the start of a list item (the Tab key)' },
  run: (args) => (state, dispatch) => {
    if (args?.atStart && !atStartOfListItem(state)) return false;
    return changeLevels(state, dispatch, (ref) => ref.ilvl + 1);
  },
  isEnabled: someListItem,
});

register<{ atStart?: boolean } | void>({
  name: 'list.outdent',
  label: 'Decrease list level',
  args: { atStart: 'true: only at the start of a list item (Shift+Tab)' },
  run: (args) => (state, dispatch) => {
    if (args?.atStart && !atStartOfListItem(state)) return false;
    return changeLevels(state, dispatch, (ref) => ref.ilvl - 1);
  },
  isEnabled: someListItem,
});

// ---------------------------------------------------------------------------------------------
// continue
// ---------------------------------------------------------------------------------------------

/** Every paragraph of the document in story order (tables and controls descended into), with its position. */
function allParagraphs(doc: PMNode): { node: PMNode; pos: number }[] {
  const out: { node: PMNode; pos: number }[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name === 'paragraph') { out.push({ node, pos }); return false; }
    return true;
  });
  return out;
}

/**
 * Word's "Continue Numbering": the caret's list item, and the later paragraphs of its `w:num`,
 * move to the nearest preceding list of the same abstract definition, so that the counting
 * carries on from it. A restart is a `w:num` of its own over the shared definition (ED-003
 * section 5.3), which is what this undoes; nothing to continue from refuses.
 */
register<void>({
  name: 'list.continue',
  label: 'Continue numbering',
  run: () => (state, dispatch) => {
    const numbering = numberingOf(state);
    if (!numbering) return false;
    const first = listItemsIn(state)[0];
    if (!first) return false;
    const abstractOf = (numId: string): string | undefined => numbering.definitions.list(numId)?.abstractDefinition?.id;
    const own = abstractOf(first.ref.numId!);
    if (own === undefined) return false;
    const paragraphs = allParagraphs(state.doc);
    const at = paragraphs.findIndex((p) => p.pos === first.pos);
    let target: string | undefined;
    for (let i = at - 1; i >= 0 && target === undefined; i--) {
      const ref = listRefOf(paragraphs[i]!.node.attrs as ParagraphAttrs, numbering.emulator);
      if (ref && ref.numId !== first.ref.numId && abstractOf(ref.numId!) === own) target = ref.numId!;
    }
    if (target === undefined) return false;
    const tr = state.tr;
    for (let i = at; i < paragraphs.length; i++) {
      const { node, pos } = paragraphs[i]!;
      const attrs = node.attrs as ParagraphAttrs;
      const ref = listRefOf(attrs, numbering.emulator);
      if (!ref || ref.numId !== first.ref.numId) continue;
      tr.setNodeMarkup(pos, undefined, { ...attrs, numPr: { numId: target, ilvl: ref.ilvl } });
    }
    dispatch?.(tr);
    return true;
  },
  isEnabled: someListItem,
});

// ---------------------------------------------------------------------------------------------
// the definition commands (step C2): start, restart, toggle; and what Enter and Backspace need
// ---------------------------------------------------------------------------------------------

export type ListKind = 'bullet' | 'number';

/** What kind of list a paragraph is in (its level's `w:numFmt`), or undefined when it is not a list item. */
export function listKindOf(attrs: ParagraphAttrs, numbering: Numbering | undefined): ListKind | undefined {
  if (!numbering) return undefined;
  const ref = listRefOf(attrs, numbering.emulator);
  if (!ref) return undefined;
  return numbering.definitions.list(ref.numId!)?.level(String(ref.ilvl))?.isBullet ? 'bullet' : 'number';
}

/** The `numPr` that takes a paragraph out of its list: none, or `numId` 0 where its style numbers it (the engine's detach rule). */
export function detachPatch(attrs: ParagraphAttrs, numbering: Numbering | undefined): NumPrAttrs | null {
  return numbering && styleNumbers(attrs, numbering.emulator) ? { numId: '0', ilvl: 0 } : null;
}

/** The paragraph just before the first selected one, in story order. */
function paragraphBefore(state: EditorState, pos: number): PMNode | undefined {
  let previous: PMNode | undefined;
  let found: PMNode | undefined;
  state.doc.descendants((node, at) => {
    if (found) return false;
    if (node.type.name === 'opaqueBlock') return false;
    if (node.type.name !== 'paragraph') return true;
    if (at === pos) { found = previous; return false; }
    previous = node;
    return false;
  });
  return found;
}

/** The session's numbering after a part write: the emulator over the definitions the write refreshed. */
function numberingAfterWrite(state: EditorState): Numbering | undefined {
  return numberingOf(state);
}

/**
 * Office JS `paragraph.startNewList({ bullet, level })` over every selected paragraph: the
 * `w:num` is made by `prepare` (the session's `newList`, the part write) and `run` attaches the
 * paragraphs to it at the level, one transaction.
 */
register<{ bullet?: boolean; level?: number; numId?: string }>({
  name: 'list.start',
  label: 'Start a new list',
  // `numId` is filled in by `prepare`, not asked of anyone: the palette lists the two arguments below.
  args: { bullet: 'true for a bulleted list; a numbered one by default', level: 'the level, 0 by default' },
  choices: { bullet: ['false', 'true'] },
  prepare: async (state, args) => {
    const session = sessionOf(state);
    if (!session || paragraphsIn(state).length === 0) return args;
    return { ...args, numId: await session.newList(args.bullet === true, state.doc) };
  },
  run: (args) => (state, dispatch) => {
    if (args.numId === undefined) return false;
    const numbering = numberingAfterWrite(state);
    const ilvl = clampLevel(args.numId, args.level ?? 0, numbering?.definitions) ?? 0;
    return updateParagraphs(() => ({ numPr: { numId: args.numId!, ilvl } }))(state, dispatch);
  },
  isEnabled: (state) => paragraphsIn(state).length > 0,
});

/**
 * Office JS `paragraph.restartList()`: a `w:num` over the same definition with a
 * `w:startOverride` (`prepare`, the session's `restartList`), and this paragraph and the later
 * paragraphs of its list moved to it (`run`), which is the mirror of `list.continue`.
 */
register<{ from?: string; numId?: string } | void>({
  name: 'list.restart',
  label: 'Restart numbering',
  // `from` and `numId` are filled in by `prepare`: no argument is asked of anyone.
  prepare: async (state, args) => {
    const session = sessionOf(state);
    const first = listItemsIn(state)[0];
    if (!session || !first) return args ?? {};
    const from = first.ref.numId!;
    return { ...(args ?? {}), from, numId: await session.restartList(from) };
  },
  run: (args) => (state, dispatch) => {
    if (!args?.from || !args.numId) return false;
    const numbering = numberingAfterWrite(state);
    const first = listItemsIn(state)[0];
    if (!numbering || !first) return false;
    const tr = state.tr;
    let any = false;
    for (const { node, pos } of allParagraphs(state.doc)) {
      if (pos < first.pos) continue;
      const attrs = node.attrs as ParagraphAttrs;
      const ref = listRefOf(attrs, numbering.emulator);
      if (!ref || ref.numId !== args.from) continue;
      tr.setNodeMarkup(pos, undefined, { ...attrs, numPr: { numId: args.numId, ilvl: ref.ilvl } });
      any = true;
    }
    if (!any) return false;
    dispatch?.(tr);
    return true;
  },
  isEnabled: someListItem,
});

export type ToggleCase = { kind: 'detach' } | { kind: 'attach'; numId: string; ilvl: number } | { kind: 'start' };

/**
 * What the Bullets or Numbering button does over the selection (ED-003 section 5.3): every
 * selected paragraph already in a list of that kind, detach; else a list of that kind ending
 * in the paragraph just before, attach and continue it; else start a new one.
 */
export function toggleCase(state: EditorState, bullet: boolean): ToggleCase {
  const numbering = numberingOf(state);
  const wanted: ListKind = bullet ? 'bullet' : 'number';
  const paragraphs = paragraphsIn(state);
  if (numbering && paragraphs.length > 0 && paragraphs.every((p) => listKindOf(p.node.attrs as ParagraphAttrs, numbering) === wanted)) return { kind: 'detach' };
  const before = paragraphs[0] ? paragraphBefore(state, paragraphs[0].pos) : undefined;
  if (numbering && before) {
    const attrs = before.attrs as ParagraphAttrs;
    if (listKindOf(attrs, numbering) === wanted) {
      const ref = listRefOf(attrs, numbering.emulator)!;
      return { kind: 'attach', numId: ref.numId!, ilvl: ref.ilvl };
    }
  }
  return { kind: 'start' };
}

register<{ bullet: boolean; numId?: string }>({
  name: 'list.toggle',
  label: 'Bullets or numbering on or off',
  args: { bullet: 'true for bullets, false for numbering' },
  choices: { bullet: ['true', 'false'] },
  prepare: async (state, args) => {
    const session = sessionOf(state);
    if (!session || toggleCase(state, args.bullet).kind !== 'start' || paragraphsIn(state).length === 0) return args;
    return { ...args, numId: await session.newList(args.bullet, state.doc) };
  },
  run: (args) => (state, dispatch) => {
    const c = toggleCase(state, args.bullet);
    const numbering = numberingAfterWrite(state);
    if (c.kind === 'detach') return updateParagraphs((a) => ({ numPr: detachPatch(a, numbering) }))(state, dispatch);
    if (c.kind === 'attach') return updateParagraphs(() => ({ numPr: { numId: c.numId, ilvl: c.ilvl } }))(state, dispatch);
    if (args.numId === undefined) return false;
    return updateParagraphs(() => ({ numPr: { numId: args.numId!, ilvl: 0 } }))(state, dispatch);
  },
  isEnabled: (state) => paragraphsIn(state).length > 0,
  isActive: (state, args) => {
    if (!args) return false;
    const numbering = numberingOf(state);
    const first = paragraphsIn(state)[0];
    return first !== undefined && listKindOf(first.node.attrs as ParagraphAttrs, numbering) === (args.bullet ? 'bullet' : 'number');
  },
});
