// SPDX-License-Identifier: GPL-3.0-or-later
//
// Run formatting (ED-002 section 4.3): Office JS `Font` property names, each a change to the
// marks of section 3.3 over the selection, with grapheme-safe edges and the atoms' `fmt`.
import { type Command, type EditorState, type Transaction } from 'prosemirror-state';
import { Mark, MarkType, type Node as PMNode, type ResolvedPos, type Schema } from 'prosemirror-model';
import { readRunOptions, type UnderlineType } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { highlightHexValue, highlightNameForColor } from '@docx4j/generated-objects-ts/helpers/wml';
import type { Element } from '@docx4j/core-ts';
import { FORMATTING_MARKS, type FontAttrs } from '../schema.mjs';
import { formattingKey, marksFromKey } from '../export.mjs';
import { formattingMarks } from '../import.mjs';
import { graphemeBoundary } from '../text.mjs';
import { register, sessionOf, selectionSpans, TOGGLE_CHOICES } from './spec.mjs';
import { effectiveFormatting } from './reads.mjs';

/** The inline node types that carry `runSource` and `fmt` (UPSTREAM-2). */
export const ATOM_TYPES = new Set(['tab', 'hardBreak', 'image', 'footnoteRef', 'opaqueInline']);

/**
 * An atom that formatting applies to: one inside a run (`runSource`), or one created in the
 * editor (`fmt`). A paragraph-level opaque item (a bookmark, a comment marker, an empty run)
 * has no run of its own, so formatting neither reads nor writes it.
 */
export function isFormattableAtom(node: PMNode): boolean {
  if (!ATOM_TYPES.has(node.type.name)) return false;
  if (node.type.name !== 'opaqueInline') return true;
  return node.attrs['runSource'] !== null || node.attrs['fmt'] !== null;
}

/** A change to a set of formatting marks. */
export type MarkChange = (marks: readonly Mark[], schema: Schema) => readonly Mark[];

// ---------------------------------------------------------------------------------------------
// Ranges and atoms
// ---------------------------------------------------------------------------------------------

/** The text node a position falls inside, with its start position, when it is inside one. */
function textNodeAround($pos: ResolvedPos): { node: PMNode; start: number } | undefined {
  if ($pos.textOffset === 0) return undefined;
  const node = $pos.parent.child($pos.index());
  return node.isText ? { node, start: $pos.pos - $pos.textOffset } : undefined;
}

/** `from` moved back and `to` forward to grapheme boundaries (ED-001 section 3.7). */
export function graphemeSafeRange(doc: PMNode, from: number, to: number): { from: number; to: number } {
  const a = textNodeAround(doc.resolve(from));
  if (a) from = a.start + graphemeBoundary(a.node.text ?? '', from - a.start, 'backward');
  const b = textNodeAround(doc.resolve(to));
  if (b) to = b.start + graphemeBoundary(b.node.text ?? '', to - b.start, 'forward');
  return { from, to };
}

/** The formatting marks an atom stands for: its `fmt`, else its source run's, else none. */
export function atomMarks(node: PMNode, state: EditorState): readonly Mark[] {
  const fmt = node.attrs['fmt'] as string | null;
  if (typeof fmt === 'string') return marksFromKey(fmt, state.schema);
  const runSource = node.attrs['runSource'] as string | null;
  const session = sessionOf(state);
  if (runSource && session) {
    const run = session.ctx.ids.get(runSource) as Element<wml.R> | undefined;
    if (run) return formattingMarks(readRunOptions(run.value.rPr), run.value.rPr, state.schema);
  }
  return [];
}

function isFormatting(mark: Mark): boolean {
  return (FORMATTING_MARKS as readonly string[]).includes(mark.type.name);
}

/** Apply the difference between two mark sets to a text range. */
function applyMarkDiff(tr: Transaction, from: number, to: number, before: readonly Mark[], after: readonly Mark[]): void {
  for (const m of after) {
    if (!isFormatting(m)) continue;
    const old = before.find((b) => b.type === m.type);
    if (!old || !old.eq(m)) tr.addMark(from, to, m);
  }
  for (const m of before) {
    if (!isFormatting(m)) continue;
    if (!after.some((a) => a.type === m.type)) tr.removeMark(from, to, m.type);
  }
}

/**
 * Apply a change over the selection: stored marks for a caret; otherwise every text node in each
 * grapheme-safe span (one, or a cell selection's cells) and, for the atoms in it, their `fmt`.
 */
export function applyMarkChange(change: MarkChange): Command {
  return (state, dispatch) => {
    const tr = state.tr;
    if (state.selection.empty) {
      const current = state.storedMarks ?? state.selection.$from.marks();
      tr.setStoredMarks(change(current, state.schema));
      dispatch?.(tr);
      return true;
    }
    for (const span of selectionSpans(state)) {
      const r = graphemeSafeRange(state.doc, span.from, span.to);
      applyMarkChangeIn(tr, state, r.from, r.to, change);
    }
    dispatch?.(tr);
    return true;
  };
}

/**
 * Apply a change to every text node and formattable atom between `from` and `to` of `tr`'s
 * document, as part of a larger transaction (a link's look, links.mts); the range is taken as
 * given, so its edges must already be grapheme boundaries.
 */
export function applyMarkChangeIn(tr: Transaction, state: EditorState, from: number, to: number, change: MarkChange): void {
  const atoms: { pos: number; node: PMNode; key: string }[] = [];
  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (node.isText) {
      const s = Math.max(pos, from);
      const e = Math.min(pos + node.nodeSize, to);
      applyMarkDiff(tr, s, e, node.marks, change(node.marks, state.schema));
      return false;
    }
    if (ATOM_TYPES.has(node.type.name)) {
      if (!isFormattableAtom(node)) return false;
      const current = atomMarks(node, state);
      const key = formattingKey(change(current, state.schema));
      if (key !== formattingKey(current)) atoms.push({ pos, node, key });
      return false;
    }
    return true;
  });
  for (const a of atoms) tr.setNodeMarkup(a.pos, undefined, { ...a.node.attrs, fmt: a.key });
}

// ---------------------------------------------------------------------------------------------
// Reading the selection's marks (isActive)
// ---------------------------------------------------------------------------------------------

/** The mark sets of the selection: the caret's marks, or one set per text node and atom in the range. */
export function markSetsIn(state: EditorState): (readonly Mark[])[] {
  if (state.selection.empty) return [state.storedMarks ?? state.selection.$from.marks()];
  const out: (readonly Mark[])[] = [];
  for (const { from, to } of selectionSpans(state)) {
    state.doc.nodesBetween(from, to, (node) => {
      if (node.isText) { out.push(node.marks); return false; }
      if (ATOM_TYPES.has(node.type.name)) { if (isFormattableAtom(node)) out.push(atomMarks(node, state)); return false; }
      return true;
    });
  }
  return out;
}

/** True when every mark set in the selection has a mark of the type (and attrs, when given). */
export function allHave(state: EditorState, type: MarkType, attrs?: Record<string, unknown>): boolean {
  const sets = markSetsIn(state);
  if (sets.length === 0) return false;
  return sets.every((marks) => marks.some((m) => m.type === type && (attrs === undefined || Object.entries(attrs).every(([k, v]) => m.attrs[k] === v))));
}

export const withMark = (marks: readonly Mark[], mark: Mark): readonly Mark[] => mark.addToSet(marks);
export const withoutMark = (marks: readonly Mark[], type: MarkType): readonly Mark[] => type.removeFromSet(marks);

type Toggle = boolean | 'toggle';

function resolveToggle(state: EditorState, value: Toggle, active: () => boolean): boolean {
  return value === 'toggle' ? !active() : value;
}

// ---------------------------------------------------------------------------------------------
// font.bold, italic, and the other booleans
// ---------------------------------------------------------------------------------------------

/** The `api` line of a boolean font property (ED-003 section 7.2): a toggle reads the property first. */
function booleanApi(property: string): (args: { value: Toggle }) => string {
  return ({ value }) => (value === 'toggle' ? `selection.font.${property} = !selection.font.${property};` : `selection.font.${property} = ${value};`);
}

function booleanMark(name: string, markName: string, label: string, property?: string): void {
  register<{ value: Toggle }>({
    name,
    label,
    args: { value: "boolean | 'toggle'" },
    choices: { value: TOGGLE_CHOICES },
    ...(property ? { api: booleanApi(property) } : {}),
    run: ({ value }) => (state, dispatch) => {
      const type = state.schema.marks[markName]!;
      const on = resolveToggle(state, value, () => allHave(state, type));
      return applyMarkChange((marks) => (on ? withMark(marks, type.create()) : withoutMark(marks, type)))(state, dispatch);
    },
    isActive: (state) => allHave(state, state.schema.marks[markName]!),
  });
}

booleanMark('font.bold', 'bold', 'Bold', 'bold');
booleanMark('font.italic', 'italic', 'Italic', 'italic');
booleanMark('font.rtl', 'rtl', 'Right-to-left run');

register<{ value: Toggle }>({
  name: 'font.strikeThrough',
  label: 'Strikethrough',
  args: { value: "boolean | 'toggle'" },
  choices: { value: TOGGLE_CHOICES },
  api: booleanApi('strikeThrough'),
  run: ({ value }) => (state, dispatch) => {
    const type = state.schema.marks['strike']!;
    const on = resolveToggle(state, value, () => allHave(state, type, { double: false }));
    return applyMarkChange((marks) => {
      const cur = marks.find((m) => m.type === type);
      if (on) return withMark(marks, type.create({ double: false }));
      return cur && cur.attrs['double'] !== true ? withoutMark(marks, type) : [...marks];
    })(state, dispatch);
  },
  isActive: (state) => allHave(state, state.schema.marks['strike']!, { double: false }),
});

register<{ value: Toggle }>({
  name: 'font.doubleStrikeThrough',
  label: 'Double strikethrough',
  args: { value: "boolean | 'toggle'" },
  choices: { value: TOGGLE_CHOICES },
  api: booleanApi('doubleStrikeThrough'),
  run: ({ value }) => (state, dispatch) => {
    const type = state.schema.marks['strike']!;
    const on = resolveToggle(state, value, () => allHave(state, type, { double: true }));
    return applyMarkChange((marks) => {
      const cur = marks.find((m) => m.type === type);
      if (on) return withMark(marks, type.create({ double: true }));
      return cur && cur.attrs['double'] === true ? withoutMark(marks, type) : [...marks];
    })(state, dispatch);
  },
  isActive: (state) => allHave(state, state.schema.marks['strike']!, { double: true }),
});

function vertAlignCommand(name: string, val: 'subscript' | 'superscript', label: string): void {
  register<{ value: Toggle }>({
    name,
    label,
    args: { value: "boolean | 'toggle'" },
    choices: { value: TOGGLE_CHOICES },
    api: booleanApi(val),
    run: ({ value }) => (state, dispatch) => {
      const type = state.schema.marks['vertAlign']!;
      const on = resolveToggle(state, value, () => allHave(state, type, { val }));
      return applyMarkChange((marks) => {
        const cur = marks.find((m) => m.type === type);
        if (on) return withMark(marks, type.create({ val }));
        return cur && cur.attrs['val'] === val ? withoutMark(marks, type) : [...marks];
      })(state, dispatch);
    },
    isActive: (state) => allHave(state, state.schema.marks['vertAlign']!, { val }),
  });
}

vertAlignCommand('font.subscript', 'subscript', 'Subscript');
vertAlignCommand('font.superscript', 'superscript', 'Superscript');

register<{ value: UnderlineType | 'None' | Toggle }>({
  name: 'font.underline',
  label: 'Underline',
  args: { value: "an underline name ('Single', 'Double', 'Wave', ...), 'None', boolean or 'toggle'" },
  choices: { value: ['toggle', 'None', 'Single', 'Double', 'Thick', 'Dotted', 'DottedHeavy', 'DashLine', 'DashLineHeavy', 'DashLineLong', 'DashLineLongHeavy', 'DotDashLine', 'DotDashLineHeavy', 'TwoDotDashLine', 'TwoDotDashLineHeavy', 'Wave', 'WaveHeavy', 'WaveDouble', 'Word'] },
  api: ({ value }) => (value === 'toggle' ? "selection.font.underline = selection.font.underline === 'None' ? 'Single' : 'None';"
    : `selection.font.underline = ${JSON.stringify(value === true ? 'Single' : value === false ? 'None' : value)};`),
  run: ({ value }) => (state, dispatch) => {
    const type = state.schema.marks['underline']!;
    let val: string | null;
    if (value === 'toggle') val = allHave(state, type) ? null : 'Single';
    else if (value === true) val = 'Single';
    else if (value === false || value === 'None') val = null;
    else val = value;
    return applyMarkChange((marks) => (val === null ? withoutMark(marks, type) : withMark(marks, type.create({ val }))))(state, dispatch);
  },
  isActive: (state, args) => allHave(state, state.schema.marks['underline']!, args && typeof args.value === 'string' && args.value !== 'toggle' && args.value !== 'None' ? { val: args.value } : undefined),
});

// ---------------------------------------------------------------------------------------------
// The compound font mark: name, size, color, highlight
// ---------------------------------------------------------------------------------------------

/** Set one field of the `font` mark, removing the mark when every field is null. */
export function setFontField(field: keyof FontAttrs, value: string | number | null): MarkChange {
  return (marks, schema) => {
    const type = schema.marks['font']!;
    const cur = marks.find((m) => m.type === type);
    const attrs: FontAttrs = { name: null, size: null, color: null, highlight: null, ...(cur?.attrs as Partial<FontAttrs> | undefined) };
    (attrs as unknown as Record<string, unknown>)[field] = value;
    const empty = attrs.name === null && attrs.size === null && attrs.color === null && attrs.highlight === null;
    return empty ? withoutMark(marks, type) : withMark(marks, type.create(attrs));
  };
}

function fontFieldValue(state: EditorState, field: keyof FontAttrs): unknown | 'Mixed' {
  const sets = markSetsIn(state);
  const values = sets.map((marks) => (marks.find((m) => m.type.name === 'font')?.attrs[field] as unknown) ?? null);
  return values.every((v) => v === values[0]) ? (values[0] ?? null) : 'Mixed';
}

register<{ value: string }>({
  name: 'font.name',
  label: 'Font',
  args: { value: "a font name; '' clears" },
  api: ({ value }) => (value === '' ? undefined : `selection.font.name = ${JSON.stringify(value)};`),
  run: ({ value }) => applyMarkChange(setFontField('name', value === '' ? null : value)),
  isActive: (state, args) => args !== undefined && fontFieldValue(state, 'name') === args.value,
});

register<{ value: number }>({
  name: 'font.size',
  label: 'Font size',
  args: { value: 'points; 0 clears' },
  api: ({ value }) => (value === 0 ? undefined : `selection.font.size = ${value};`),
  run: ({ value }) => applyMarkChange(setFontField('size', value === 0 ? null : value)),
  isActive: (state, args) => args !== undefined && fontFieldValue(state, 'size') === args.value,
});

register<{ value: string }>({
  name: 'font.color',
  label: 'Font colour',
  args: { value: "'#RRGGBB', 'auto'; '' clears" },
  api: ({ value }) => (value === '' ? undefined : `selection.font.color = ${JSON.stringify(value)};`),
  run: ({ value }) => applyMarkChange(setFontField('color', value === '' ? null : value)),
  isActive: (state, args) => args !== undefined && fontFieldValue(state, 'color') === args.value,
});

/**
 * The `font` mark's highlight for a command argument: null to clear, else the '#RRGGBB' of a highlight name or of a
 * hex in the palette, the form the projection reads from a file (`readRunOptions`), so the save's self-check sees
 * the same mark after a reload (a name stored as given had refused every save with a highlight applied by name,
 * 2026-10-31); undefined for a value outside the palette, which the export could not write.
 */
function highlightArg(value: string | null | undefined): string | null | undefined {
  if (value === '' || value === null || value === undefined) return null;
  // A highlight name is Word's before anything else: the object model reads a colour as Office JS does (objects 0.8.0,
  // CR-014), where `green` is CSS's #008000, Word's darkGreen (ED-008 section 8.8 install item 18 (g)).
  const own = highlightHexValue(value);
  if (own !== undefined) return own;
  const name = highlightNameForColor(value);
  return name === undefined ? undefined : highlightHexValue(name);
}

register<{ value: string | null }>({
  name: 'font.highlightColor',
  label: 'Highlight',
  args: { value: "a highlight name or its '#RRGGBB'; null or '' clears" },
  // The hex the command applies, which Office JS reads as the highlight meant; a name it would read as CSS's (`green`).
  api: ({ value }) => (value === '' || value === null ? 'selection.font.highlightColor = null;' : `selection.font.highlightColor = ${JSON.stringify(highlightArg(value) ?? value)};`),
  run: ({ value }) => {
    const hex = highlightArg(value);
    return hex === undefined ? () => false : applyMarkChange(setFontField('highlight', hex));
  },
  isActive: (state, args) => args !== undefined && fontFieldValue(state, 'highlight') === highlightArg(args.value),
});

register<{ value: string }>({
  name: 'font.style',
  label: 'Character style',
  args: { value: "a character style id; '' clears" },
  run: ({ value }) => applyMarkChange((marks, schema) => {
    const type = schema.marks['runStyle']!;
    return value === '' ? withoutMark(marks, type) : withMark(marks, type.create({ id: value }));
  }),
  isActive: (state, args) => args !== undefined && allHave(state, state.schema.marks['runStyle']!, { id: args.value }),
});

// ---------------------------------------------------------------------------------------------
// Grow, shrink, clear
// ---------------------------------------------------------------------------------------------

/** Word's font-size list, the steps Ctrl+Shift+period and comma take. */
export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

/** The size the selection shows: its direct size when uniform, else the effective size, else 11. */
function currentSize(state: EditorState): number {
  const direct = fontFieldValue(state, 'size');
  if (typeof direct === 'number') return direct;
  const effective = effectiveFormatting(state).font.size;
  return typeof effective === 'number' && effective > 0 ? effective : 11;
}

function stepSize(size: number, direction: 1 | -1): number {
  if (direction > 0) return FONT_SIZES.find((s) => s > size) ?? Math.min(size + 1, 1638);
  return [...FONT_SIZES].reverse().find((s) => s < size) ?? Math.max(size - 1, 1);
}

register<void>({
  name: 'font.grow',
  label: 'Grow font',
  run: () => (state, dispatch) => applyMarkChange(setFontField('size', stepSize(currentSize(state), 1)))(state, dispatch),
});

register<void>({
  name: 'font.shrink',
  label: 'Shrink font',
  run: () => (state, dispatch) => applyMarkChange(setFontField('size', stepSize(currentSize(state), -1)))(state, dispatch),
});

register<void>({
  name: 'font.clear',
  label: 'Clear character formatting',
  run: () => applyMarkChange((marks) => marks.filter((m) => !isFormatting(m))),
});
