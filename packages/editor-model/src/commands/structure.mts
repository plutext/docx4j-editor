// SPDX-License-Identifier: GPL-3.0-or-later
//
// Text and structure (ED-002 section 4.5): typing, the paragraph split and join with Word's
// paragraph-mark rules, breaks and tabs.
import { TextSelection, type Command, type EditorState } from 'prosemirror-state';
import type { Node as PMNode, ResolvedPos } from 'prosemirror-model';
import { formattingKey } from '../export.mjs';
import type { ParagraphAttrs } from '../schema.mjs';
import { register, sessionOf, storyOf, NEW_MARK_META } from './spec.mjs';
import { numberingOf, listRefOf, detachPatch } from './lists.mjs';
import { readControl } from '../sdtpr.mjs';

register<{ text: string }>({
  name: 'insertText',
  label: 'Insert text',
  args: { text: 'the text; replaces the selection' },
  typing: true,
  api: ({ text }) => `selection.range.insertText(${JSON.stringify(text)}, 'Replace');`,
  run: ({ text }) => (state, dispatch) => {
    dispatch?.(state.tr.insertText(text).scrollIntoView());
    return true;
  },
});

/** The style's "style for following paragraph" (`w:next`), through the resolver; undefined when unknown. */
export function nextStyleOf(state: EditorState, styleId: string | null): string | null | undefined {
  const resolver = sessionOf(state)?.pkg.propertyResolverOrUndefined;
  if (!resolver) return undefined;
  const style = resolver.getStyle(styleId ?? resolver.defaultParagraphStyleId);
  const next = style?.next?.val;
  if (next === undefined) return undefined;
  return next === resolver.defaultParagraphStyleId ? null : next;
}

/**
 * Enter: split the paragraph at the caret. Word inserts a new paragraph mark at the caret, so
 * the text after it keeps the old mark: the second half keeps `sourceId`, `sectPr` and `markRPr`
 * and the first half is a new paragraph with a copy of the other attributes, and `copiedFrom`
 * naming the source element whose whole `w:pPr` the export takes as its base (ED-005 section 12.102). At the end of the
 * text, the new empty paragraph (the second half) takes the style's `w:next`, and so is the old
 * mark's paragraph with its style changed (core-ts check 29: tracked, Word records a `w:pPrChange`
 * on it). The transaction names the first half as the paragraph whose mark it made
 * (`NEW_MARK_META`), for the tracker.
 *
 * A run-level control is never split (ED-005 section 12.123): the copy in the new paragraph would carry its `w:id`, its tag
 * and its binding. A caret at a control's inner end splits after the control and one at its inner start before it, as the
 * caret just outside it does (`outOfRunControls`). In the middle of a plain text control that allows line breaks
 * (`w:multiLine`) Enter is a line break, Word's soft break; in the middle of any other it does nothing.
 */
register<void>({
  name: 'insertParagraph',
  label: 'New paragraph',
  run: () => (state, dispatch) => {
    const { $from, $to } = state.selection;
    const at = paragraphAround($from);
    if (at === undefined || at !== paragraphAround($to)) return false;
    const tr = state.tr;
    if (!state.selection.empty) tr.deleteSelection();
    const out = outOfRunControls(tr.selection.$from);
    if ('control' in out) {
      const control = readControl(out.control.attrs['sdtPr']);
      if (control.kind !== 'PlainText' || !control.multiLine) return false;
      const $caret = tr.selection.$from;
      const marks = state.storedMarks ?? $caret.marks();
      tr.replaceSelectionWith(state.schema.nodes['hardBreak']!.create({ runSource: runSourceAt($caret), fmt: formattingKey(marks), kind: 'line' }), false);
      tr.setStoredMarks(marks);
      dispatch?.(tr.scrollIntoView());
      return true;
    }
    const $pos = tr.doc.resolve(out.pos);
    const para = $pos.parent;
    const attrs = para.attrs as ParagraphAttrs;
    const atEnd = $pos.parentOffset === para.content.size;
    const before = $pos.before();
    // Enter on an empty list item ends the list rather than adding an empty item (Word; ED-003 section 5.3).
    const numbering = numberingOf(state);
    if (para.content.size === 0 && numbering && listRefOf(attrs, numbering.emulator)) {
      tr.setNodeMarkup(before, undefined, { ...attrs, numPr: detachPatch(attrs, numbering) });
      dispatch?.(tr.scrollIntoView());
      return true;
    }
    tr.split($pos.pos);
    const first = tr.doc.nodeAt(before)!;
    const secondPos = before + first.nodeSize;
    tr.setNodeMarkup(before, undefined, { ...attrs, sourceId: null, sectPr: null, copiedFrom: attrs.sourceId ?? attrs.copiedFrom });
    let second: Partial<ParagraphAttrs> = {};
    if (atEnd) {
      const next = nextStyleOf(state, attrs.pStyle);
      if (next !== undefined && next !== attrs.pStyle) second = { pStyle: next };
    }
    tr.setNodeMarkup(secondPos, undefined, { ...attrs, ...second });
    tr.setMeta(NEW_MARK_META, [before]);
    tr.setSelection(TextSelection.create(tr.doc, secondPos + 1)).scrollIntoView();
    dispatch?.(tr);
    return true;
  },
  isEnabled: (state) => paragraphAround(state.selection.$from) !== undefined && paragraphAround(state.selection.$from) === paragraphAround(state.selection.$to),
});

/** The start of the paragraph a position is in, directly or inside run-level controls; undefined elsewhere. */
function paragraphAround($pos: ResolvedPos): number | undefined {
  let depth = $pos.depth;
  while (depth > 0 && $pos.node(depth).type.name === 'sdtInline') depth--;
  return depth > 0 && $pos.node(depth).type.name === 'paragraph' ? $pos.before(depth) : undefined;
}

/** Whether a position is inside a run-level control, at any depth. */
export function inRunControl($pos: ResolvedPos): boolean {
  for (let depth = $pos.depth; depth > 0; depth--) if ($pos.node(depth).type.name === 'sdtInline') return true;
  return false;
}

/**
 * Where Enter splits for a caret inside run-level controls: at a control's inner end after it, at its inner start before
 * it (an empty control's caret is at its end), outwards through nested controls to a position in the paragraph; or the
 * innermost control whose middle the caret is in.
 */
function outOfRunControls($caret: ResolvedPos): { pos: number } | { control: PMNode } {
  let $at = $caret;
  while ($at.parent.type.name === 'sdtInline') {
    if ($at.parentOffset === $at.parent.content.size) $at = $at.doc.resolve($at.after());
    else if ($at.parentOffset === 0) $at = $at.doc.resolve($at.before());
    else return { control: $at.parent };
  }
  return { pos: $at.pos };
}

/** The paragraph before or after the caret's, when it is a sibling paragraph. */
function siblingParagraph($pos: ResolvedPos, direction: -1 | 1): PMNode | undefined {
  const depth = $pos.depth - 1;
  const parent = $pos.node(depth);
  const index = $pos.index(depth) + direction;
  if (index < 0 || index >= parent.childCount) return undefined;
  const node = parent.child(index);
  return node.type.name === 'paragraph' ? node : undefined;
}

/**
 * Join two paragraphs at a boundary. The merged paragraph keeps the first's attributes and
 * `sourceId`, except `sectPr`, which comes from the second: deleting a section-break mark joins
 * the text into the following section, as Word does. An empty first paragraph is removed instead.
 */
function joinAt(state: EditorState, dispatch: ((tr: import('prosemirror-state').Transaction) => void) | undefined, boundary: number, first: PMNode, second: PMNode, caret: number): boolean {
  const tr = state.tr;
  if (first.content.size === 0) {
    tr.delete(boundary - first.nodeSize, boundary);
  } else {
    tr.join(boundary);
    tr.setNodeMarkup(boundary - first.nodeSize, undefined, { ...first.attrs, sectPr: (second.attrs as ParagraphAttrs).sectPr });
    tr.setSelection(TextSelection.create(tr.doc, caret));
  }
  dispatch?.(tr.scrollIntoView());
  return true;
}

register<void>({
  name: 'joinBackward',
  label: 'Delete the paragraph mark before',
  run: () => (state, dispatch) => {
    const { $from, empty } = state.selection;
    if (!empty || $from.parentOffset !== 0 || $from.parent.type.name !== 'paragraph') return false;
    // Backspace at the start of a list item takes it out of the list before it joins anything (Word).
    const numbering = numberingOf(state);
    const attrs = $from.parent.attrs as ParagraphAttrs;
    if (numbering && listRefOf(attrs, numbering.emulator)) {
      dispatch?.(state.tr.setNodeMarkup($from.before(), undefined, { ...attrs, numPr: detachPatch(attrs, numbering) }).scrollIntoView());
      return true;
    }
    const prev = siblingParagraph($from, -1);
    if (!prev) return false;
    const boundary = $from.before();
    return joinAt(state, dispatch, boundary, prev, $from.parent, boundary - 1);
  },
});

register<void>({
  name: 'joinForward',
  label: 'Delete the paragraph mark after',
  run: () => (state, dispatch) => {
    const { $from, empty } = state.selection;
    if (!empty || $from.parentOffset !== $from.parent.content.size || $from.parent.type.name !== 'paragraph') return false;
    const next = siblingParagraph($from, 1);
    if (!next) return false;
    const boundary = $from.after();
    return joinAt(state, dispatch, boundary, $from.parent, next, $from.pos);
  },
});

/** The run an atom inserted at the caret joins: the text (or atom) before, else after. */
function runSourceAt($pos: ResolvedPos): string | null {
  for (const n of [$pos.nodeBefore, $pos.nodeAfter]) {
    if (!n) continue;
    if (n.isText) {
      const id = n.marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'];
      if (typeof id === 'string') return id;
    } else if ('runSource' in n.attrs && typeof n.attrs['runSource'] === 'string') {
      return n.attrs['runSource'] as string;
    }
  }
  return null;
}

export function insertAtom(makeNode: (state: EditorState, attrs: Record<string, unknown>) => PMNode): Command {
  return (state, dispatch) => {
    if (state.selection.$from.parent.type.name !== 'paragraph' && state.selection.$from.parent.type.name !== 'sdtInline') return false;
    const tr = state.tr;
    if (!state.selection.empty) tr.deleteSelection();
    const $pos = tr.selection.$from;
    const marks = state.storedMarks ?? $pos.marks();
    const node = makeNode(state, { runSource: runSourceAt($pos), fmt: formattingKey(marks) });
    tr.replaceSelectionWith(node, false);
    tr.setStoredMarks(marks);
    dispatch?.(tr.scrollIntoView());
    return true;
  };
}

register<{ type: 'Page' | 'Line' | 'Column' }>({
  name: 'insertBreak',
  label: 'Insert break',
  args: { type: "'Page' | 'Line' | 'Column'" },
  choices: { type: ['Page', 'Line', 'Column'] },
  // A page or column break has no meaning in a note or a header (ED-003 section 6.4): only a line break there.
  run: ({ type }) => (state, dispatch, view) => (type !== 'Line' && storyOf(state) !== undefined ? false : insertAtom((s, attrs) => s.schema.nodes['hardBreak']!.create({ ...attrs, kind: type === 'Page' ? 'page' : type === 'Column' ? 'column' : 'line' }))(state, dispatch, view)),
  isEnabled: (state, args) => args?.type === 'Line' || storyOf(state) === undefined,
});

register<void>({
  name: 'insertTab',
  label: 'Insert tab',
  run: () => insertAtom((state, attrs) => state.schema.nodes['tab']!.create(attrs)),
});
