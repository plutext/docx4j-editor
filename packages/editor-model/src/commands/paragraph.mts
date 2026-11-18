// SPDX-License-Identifier: GPL-3.0-or-later
//
// Paragraph formatting (ED-002 section 4.4): Office JS `Paragraph` property names in points,
// applied to every paragraph the selection touches by changing one attribute; export writes the
// difference, so the rest of `w:pPr` never moves.
import type { Command, EditorState, Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type { Alignment } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { IndentAttrs, ParagraphAttrs, SpacingAttrs, TabAttrs } from '../schema.mjs';
import { register, sessionOf, TOGGLE_CHOICES, selectionSpans } from './spec.mjs';

export const TWIPS_PER_POINT = 20;
export const twips = (pt: number): number => Math.round(pt * TWIPS_PER_POINT);
export const points = (tw: number | null): number | null => (tw === null ? null : tw / TWIPS_PER_POINT);

/** Word's indent button step: half an inch. */
export const INDENT_STEP = 720;

/** Every paragraph the selection touches, outermost order, with its position; a cell selection's, cell by cell. */
export function paragraphsIn(state: EditorState): { node: PMNode; pos: number }[] {
  const out: { node: PMNode; pos: number }[] = [];
  for (const span of selectionSpans(state)) {
    state.doc.nodesBetween(span.from, spanEnd(state.doc, span.from, span.to), (node, pos) => {
      if (node.type.name === 'paragraph') { out.push({ node, pos }); return false; }
      return true;
    });
  }
  return out;
}

/**
 * Where the blocks a selection covers end: a selection that runs to the start of the next paragraph (the selection
 * bar's, which takes a paragraph's mark, ED-005 sections 12.74 and 12.81; Shift+Down past a mark) touches that
 * paragraph at its start and no further, and Word counts nothing of it as selected, so a paragraph command and a
 * block control stop before it (section 12.81 item 5).
 */
export function blocksEnd(state: EditorState): number {
  const { from, to } = state.selection;
  return spanEnd(state.doc, from, to);
}

/** `blocksEnd` for one span of the selection. */
function spanEnd(doc: PMNode, from: number, to: number): number {
  const $to = doc.resolve(to);
  return to > from && $to.parentOffset === 0 && $to.parent.isTextblock && $to.depth > 0 ? $to.before() : to;
}

/** Apply an attribute update to every paragraph in the selection; false when there is none. */
export function updateParagraphs(update: (attrs: ParagraphAttrs, node: PMNode) => Partial<ParagraphAttrs> | null): Command {
  return (state, dispatch) => {
    const targets = paragraphsIn(state);
    if (targets.length === 0) return false;
    const tr: Transaction = state.tr;
    let any = false;
    for (const { node, pos } of targets) {
      const patch = update(node.attrs as ParagraphAttrs, node);
      if (patch) { tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...patch }); any = true; }
    }
    if (!any) return false;
    dispatch?.(tr);
    return true;
  };
}

/** A value shared by every paragraph in the selection, else 'Mixed'. */
export function paragraphValue<T>(state: EditorState, read: (attrs: ParagraphAttrs) => T): T | 'Mixed' {
  const values = paragraphsIn(state).map((p) => read(p.node.attrs as ParagraphAttrs));
  if (values.length === 0) return 'Mixed';
  return values.every((v) => JSON.stringify(v) === JSON.stringify(values[0])) ? values[0]! : 'Mixed';
}

/** The `api` line of a paragraph property over every selected paragraph (ED-003 section 7.2). */
const paragraphApi = (property: string, value: unknown): string => `for (const p of selection.paragraphs) p.${property} = ${JSON.stringify(value)};`;

const ind = (attrs: ParagraphAttrs): IndentAttrs => attrs.ind ?? { start: null, end: null, firstLine: null, hanging: null };
const spacing = (attrs: ParagraphAttrs): SpacingAttrs => attrs.spacing ?? { before: null, after: null, line: null, lineRule: null };

// ---------------------------------------------------------------------------------------------
// style, alignment
// ---------------------------------------------------------------------------------------------

register<{ value: string | null }>({
  name: 'paragraph.style',
  label: 'Paragraph style',
  args: { value: "a paragraph style id; '' or null for the default style" },
  api: ({ value }) => (value === '' || value === null ? undefined : paragraphApi('styleId', value)),
  run: ({ value }) => updateParagraphs(() => ({ pStyle: value === '' ? null : value })),
  isActive: (state, args) => args !== undefined && paragraphValue(state, (a) => a.pStyle) === (args.value === '' ? null : args.value),
});

export const ALIGNMENT_TO_JC: Record<Alignment, string> = { Left: 'left', Centered: 'center', Right: 'right', Justified: 'both' };
export const JC_TO_ALIGNMENT: Record<string, Alignment> = { left: 'Left', start: 'Left', center: 'Centered', right: 'Right', end: 'Right', both: 'Justified', distribute: 'Justified' };

register<{ value: Alignment | 'Unknown' }>({
  name: 'paragraph.alignment',
  label: 'Alignment',
  args: { value: "'Left' | 'Centered' | 'Right' | 'Justified'; 'Unknown' clears" },
  choices: { value: ['Left', 'Centered', 'Right', 'Justified', 'Unknown'] },
  api: ({ value }) => paragraphApi('alignment', value),
  // A value the paragraph inherits anyway is not written: Word's Ctrl+L on a centred paragraph over a
  // left-aligned style removes the w:jc, tracking on (core-ts check 26) and off (Jason's Word 2010 and
  // Word 15 measurements, ED-005 section 12.33), and the engine's setter does the same from CR-002
  // section 38 (core-ts 426e7b5, unreleased; until 0.3.1 the api case is a known difference).
  run: ({ value }) => (state, dispatch) => updateParagraphs((attrs) => {
    if (value === 'Unknown') return { jc: null };
    return { jc: inheritedAlignment(state, attrs) === value ? null : ALIGNMENT_TO_JC[value] ?? null };
  })(state, dispatch),
  isActive: (state, args) => args !== undefined && paragraphValue(state, (a) => (a.jc ? JC_TO_ALIGNMENT[a.jc] ?? 'Unknown' : 'Unknown')) === args.value,
});

/**
 * The alignment a paragraph resolves to with no w:jc of its own: its style's, else the defaults',
 * else left, as Word lays out a paragraph with none (the engine's `inheritedAlignment`). Undefined
 * without a resolver, where the value is written as asked.
 */
function inheritedAlignment(state: EditorState, attrs: ParagraphAttrs): Alignment | undefined {
  const resolver = sessionOf(state)?.pkg.propertyResolverOrUndefined;
  if (!resolver) return undefined;
  const pPr = (attrs.pStyle ? { pStyle: { val: attrs.pStyle } } : {}) as wml.PPr;
  const jc = resolver.getEffectivePPr(pPr).jc?.val;
  return jc ? JC_TO_ALIGNMENT[jc] ?? 'Left' : 'Left';
}

// ---------------------------------------------------------------------------------------------
// indents and spacing, in points
// ---------------------------------------------------------------------------------------------

register<{ value: number }>({
  name: 'paragraph.leftIndent',
  label: 'Left indent',
  args: { value: 'points' },
  api: ({ value }) => paragraphApi('leftIndent', value),
  run: ({ value }) => updateParagraphs((a) => ({ ind: { ...ind(a), start: twips(value) } })),
});

register<{ value: number }>({
  name: 'paragraph.rightIndent',
  label: 'Right indent',
  args: { value: 'points' },
  api: ({ value }) => paragraphApi('rightIndent', value),
  run: ({ value }) => updateParagraphs((a) => ({ ind: { ...ind(a), end: twips(value) } })),
});

register<{ value: number }>({
  name: 'paragraph.firstLineIndent',
  label: 'First line indent',
  args: { value: 'points; negative is a hanging indent' },
  api: ({ value }) => paragraphApi('firstLineIndent', value),
  run: ({ value }) => updateParagraphs((a) => ({
    ind: value < 0 ? { ...ind(a), firstLine: null, hanging: twips(-value) } : { ...ind(a), firstLine: twips(value), hanging: null },
  })),
});

register<{ value: number }>({
  name: 'paragraph.spaceBefore',
  label: 'Space before',
  args: { value: 'points' },
  api: ({ value }) => paragraphApi('spaceBefore', value),
  run: ({ value }) => updateParagraphs((a) => ({ spacing: { ...spacing(a), before: twips(value) } })),
});

register<{ value: number }>({
  name: 'paragraph.spaceAfter',
  label: 'Space after',
  args: { value: 'points' },
  api: ({ value }) => paragraphApi('spaceAfter', value),
  run: ({ value }) => updateParagraphs((a) => ({ spacing: { ...spacing(a), after: twips(value) } })),
});

register<{ value: number }>({
  name: 'paragraph.lineSpacing',
  label: 'Line spacing',
  args: { value: 'points' },
  api: ({ value }) => paragraphApi('lineSpacing', value),
  // Word's form, as its Office JS setter writes it (core-ts check 71b): an auto rule, w:line 20 x the points, so 12 is
  // single spacing. A paragraph the command does not reach keeps the rule its file states.
  run: ({ value }) => updateParagraphs((a) => ({ spacing: { ...spacing(a), line: twips(value), lineRule: 'auto' } })),
});

// ---------------------------------------------------------------------------------------------
// tab stops (ED-003 section 8.5, step F4)
// ---------------------------------------------------------------------------------------------

/** One stop as the command takes it: `position` in points; `alignment` and `leader` as `w:tab` states them. */
export interface TabStopArg {
  position: number;
  alignment?: string | undefined;
  leader?: string | undefined;
}

export const TAB_ALIGNMENTS: readonly string[] = ['left', 'center', 'right', 'decimal', 'bar', 'num', 'clear'];
export const TAB_LEADERS: readonly string[] = ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'];

/** The attribute a list of stops becomes: twips, ascending, one per position (the last given wins), `none` leaders dropped; undefined when a value is not one `w:tab` takes. */
export function tabAttrsOf(stops: readonly TabStopArg[]): TabAttrs[] | undefined {
  const byPos = new Map<number, TabAttrs>();
  for (const s of stops) {
    if (typeof s.position !== 'number' || !Number.isFinite(s.position)) return undefined;
    const val = s.alignment ?? 'left';
    const leader = s.leader ?? 'none';
    if (!TAB_ALIGNMENTS.includes(val) || !TAB_LEADERS.includes(leader)) return undefined;
    byPos.set(twips(s.position), { pos: twips(s.position), val, leader: leader === 'none' ? null : leader });
  }
  return [...byPos.values()].sort((a, b) => a.pos - b.pos);
}

register<{ stops: TabStopArg[] }>({
  name: 'paragraph.tabStops',
  label: 'Tab stops',
  args: { stops: "[{ position: points, alignment?: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear', leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot' }]: the paragraph's own stops, replacing them; 'clear' cancels a style's stop at that position; [] removes them all" },
  // Editor-only: Office JS has no tab-stop model to name a line after (section 8.5), so there is no api form.
  run: ({ stops }) => (state, dispatch) => {
    const tabs = tabAttrsOf(Array.isArray(stops) ? stops : []);
    if (!tabs) return false;
    return updateParagraphs(() => ({ tabs: tabs.length === 0 ? null : tabs }))(state, dispatch);
  },
});

register<void>({
  name: 'paragraph.increaseIndent',
  label: 'Increase indent',
  // The direct indent, as the editor steps it (the engine's `leftIndent` reads the effective one; ED-003 section 11.15).
  api: () => 'for (const p of selection.paragraphs) p.leftIndent = p.formatting({ direct: true }).leftIndent + 36;',
  run: () => updateParagraphs((a) => ({ ind: { ...ind(a), start: (ind(a).start ?? 0) + INDENT_STEP } })),
});

register<void>({
  name: 'paragraph.decreaseIndent',
  label: 'Decrease indent',
  api: () => 'for (const p of selection.paragraphs) p.leftIndent = Math.max(0, p.formatting({ direct: true }).leftIndent - 36);',
  run: () => updateParagraphs((a) => {
    const start = ind(a).start ?? 0;
    if (start === 0) return null;
    return { ind: { ...ind(a), start: Math.max(0, start - INDENT_STEP) } };
  }),
  isEnabled: (state) => paragraphsIn(state).some((p) => ((p.node.attrs as ParagraphAttrs).ind?.start ?? 0) > 0),
});

register<void>({
  name: 'paragraph.clearFormatting',
  label: 'Clear paragraph formatting',
  run: () => updateParagraphs(() => ({ pStyle: null, jc: null, ind: null, spacing: null, bidi: false, numPr: null, tabs: null })),
});

register<{ value: boolean | 'toggle' }>({
  name: 'paragraph.bidi',
  label: 'Right-to-left paragraph',
  args: { value: "boolean | 'toggle'" },
  choices: { value: TOGGLE_CHOICES },
  run: ({ value }) => (state, dispatch) => {
    const on = value === 'toggle' ? paragraphValue(state, (a) => a.bidi) !== true : value;
    return updateParagraphs(() => ({ bidi: on }))(state, dispatch);
  },
  isActive: (state) => paragraphValue(state, (a) => a.bidi) === true,
});
