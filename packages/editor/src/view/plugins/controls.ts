// SPDX-License-Identifier: GPL-3.0-or-later
//
// Content controls in the view (ED-003 section 4.9): the control at the caret is marked so the
// gutters and the Template perspective's panels agree on which one is current, and a data node
// dragged from the Data tab onto a control binds it (section 4.5's "bind by drag", which is what
// Word's XML Mapping pane does with a click). The decorations are a handful per transaction (the
// current control's node, or the rows or cells of its group), never one per block (ED-002
// section 10.5 item 10).
import { Plugin, PluginKey, TextSelection, type EditorState } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { controlAt, controlAround, commandNamed, rowPosition, isNodeTarget, sdtPrOf, controlWritesOf, type ControlTarget, type ControlWrite, type EditorDocument } from '@docx4j/editor-model';

export const controlsKey = new PluginKey<DecorationSet>('docx4j-controls');

/** The drag type a data node travels under (DataPanel.tsx sets it; the drop handler reads it). */
export const BIND_MIME = 'application/x-docx4j-bind';

/** What a dragged data node carries: enough for `contentControl.bind`. */
export interface BindDrag {
  xpath: string;
  prefixMappings: string;
  storeItemID: string;
}

/** The meta a transaction carries when the controls map changed outside the view (an undo, a peer), so the marks and labels refresh. */
export const CONTROLS_CHANGED = 'docx4j-controls-changed';

export interface ControlsPluginOptions {
  /** After a drop: whether the bind applied, what was dropped, and where. */
  onDrop?: (outcome: { bound: boolean; drag: BindDrag; target: ControlTarget | undefined }) => void;
}

/**
 * The current control's node, or its group's rows or cells, marked `sdt-current`; the spec
 * carries the properties so a change re-renders the node views. A transaction's own control
 * writes are the freshest properties: the state applies before the view puts them in the map.
 */
function currentControl(session: EditorDocument, state: EditorState, writes: readonly ControlWrite[] = []): DecorationSet {
  const target = controlAt(state);
  if (!target) return DecorationSet.empty;
  const written = isNodeTarget(target) ? undefined : writes.find((w) => w.groupId === target.groupId);
  const spec = { sdtPr: written ? written.entry?.sdtPr ?? null : sdtPrOf(target, session.controlEntry) };
  const decos: Decoration[] = [];
  if (isNodeTarget(target)) {
    decos.push(Decoration.node(target.pos, target.pos + target.node.nodeSize, { class: 'sdt-current' }, spec));
  } else if (target.form === 'row') {
    const [first, last] = target.rows;
    for (let r = first; r <= last; r++) {
      const pos = rowPosition(target.table, target.tablePos, r);
      decos.push(Decoration.node(pos, pos + target.table.child(r).nodeSize, { class: 'sdt-current' }, spec));
    }
  } else {
    const row = target.table.child(target.row);
    const rowPos = rowPosition(target.table, target.tablePos, target.row);
    const [first, last] = target.cells;
    let cellPos = rowPos + 1;
    row.forEach((cell, _o, i) => {
      if (i >= first && i <= last) decos.push(Decoration.node(cellPos, cellPos + cell.nodeSize, { class: 'sdt-current' }, spec));
      cellPos += cell.nodeSize;
    });
  }
  return DecorationSet.create(state.doc, decos);
}

/** Bind the innermost control at the drop point to the dragged node, then select the control's content. */
function dropBinding(view: EditorView, event: DragEvent, options: ControlsPluginOptions): boolean {
  const data = event.dataTransfer?.getData(BIND_MIME);
  if (!data) return false;
  event.preventDefault();
  let drag: BindDrag;
  try { drag = JSON.parse(data) as BindDrag; } catch { return true; }
  const at = view.posAtCoords({ left: event.clientX, top: event.clientY });
  const target = at ? controlAround(view.state.doc.resolve(at.pos)) : undefined;
  if (!at || !target) { options.onDrop?.({ bound: false, drag, target: undefined }); return true; }
  view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(at.pos))));
  const bound = commandNamed('contentControl.bind')!.run({ xpath: drag.xpath, prefixMappings: drag.prefixMappings, storeItemID: drag.storeItemID })(view.state, view.dispatch);
  commandNamed('contentControl.select')!.run(isNodeTarget(target) ? { pos: target.pos } : { groupId: target.groupId })(view.state, view.dispatch);
  view.focus();
  options.onDrop?.({ bound, drag, target });
  return true;
}

/** Whether the text at a position is laid out right to left (a paragraph's own `w:bidi` or its style's), from what is drawn. */
function isRightToLeftAt(view: EditorView, pos: number): boolean {
  try {
    const { node } = view.domAtPos(pos, -1);
    const el = node.nodeType === 3 ? node.parentElement : (node as Element);
    return el !== null && getComputedStyle(el).direction === 'rtl';
  } catch {
    return false;
  }
}

const graphemes = typeof Intl !== 'undefined' && 'Segmenter' in Intl ? new Intl.Segmenter(undefined, { granularity: 'grapheme' }) : undefined;

/** The length, in code units, of the first (or last) grapheme cluster of a text; 1 where the browser cannot segment. */
function graphemeLength(text: string, last: boolean): number {
  if (text === '' || !graphemes) return 1;
  let length = 1;
  for (const seg of graphemes.segment(text)) {
    length = seg.segment.length;
    if (!last) break;
  }
  return length;
}

/**
 * A run control's two borders are caret stops, as Word's are: with the caret inside the control at
 * its end, Right steps out after it (at a paragraph's end, where the browser has nowhere to go and
 * typing went on inside the control: Jason, 2026-10-08); at its start, Left steps out before it;
 * from just outside, one step goes in at the edge, before the first or after the last character;
 * and a step from a neighbouring character onto a border or an inner edge is taken here too, since
 * the border and the inner edge are one spot on screen and the browser lands on either.
 */
export function stepAcrossControlBorder(view: EditorView, event: KeyboardEvent): boolean {
  if (event.ctrlKey || event.metaKey || event.altKey || event.shiftKey) return false;
  if (event.key !== 'ArrowRight' && event.key !== 'ArrowLeft') return false;
  const sel = view.state.selection;
  if (!sel.empty) return false;
  const { doc } = view.state;
  const $caret = sel.$from;
  const inControl = $caret.parent.type.name === 'sdtInline';
  const forward = (event.key === 'ArrowRight') !== isRightToLeftAt(view, $caret.pos);
  const isControl = (n: PMNode | null | undefined): boolean => n?.type.name === 'sdtInline';
  let to: number | undefined;
  if (forward) {
    const next = $caret.nodeAfter;
    if (inControl && $caret.parentOffset === $caret.parent.content.size) to = $caret.after();
    else if (isControl(next)) to = $caret.pos + 1;
    else if (next?.isText && next.text !== undefined) {
      const $to = doc.resolve($caret.pos + graphemeLength(next.text, false));
      if (isControl($to.nodeAfter) || (isControl($to.parent) && $to.parentOffset === $to.parent.content.size)) to = $to.pos;
    }
  } else {
    const prev = $caret.nodeBefore;
    if (inControl && $caret.parentOffset === 0) to = $caret.before();
    else if (isControl(prev)) to = $caret.pos - 1;
    else if (prev?.isText && prev.text !== undefined) {
      const $to = doc.resolve($caret.pos - graphemeLength(prev.text, true));
      if (isControl($to.nodeBefore) || (isControl($to.parent) && $to.parentOffset === 0)) to = $to.pos;
    }
  }
  if (to === undefined) return false;
  view.dispatch(view.state.tr.setSelection(TextSelection.create(doc, to)).scrollIntoView());
  return true;
}

export function controlsPlugin(session: EditorDocument, options: ControlsPluginOptions = {}): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key: controlsKey,
    state: {
      init: (_config, state) => currentControl(session, state),
      // Recomputed on a change of the document, the selection, or the properties (a command's
      // control writes change the map alone, and the map change from elsewhere arrives as a meta).
      apply: (tr, set, _old, state) => {
        const writes = controlWritesOf(tr);
        return tr.docChanged || tr.selectionSet || tr.getMeta(CONTROLS_CHANGED) || writes.length > 0 ? currentControl(session, state, writes) : set;
      },
    },
    props: {
      decorations: (state) => controlsKey.getState(state) ?? DecorationSet.empty,
      handleDrop: (view, event) => dropBinding(view, event, options),
      handleTextInput: (view, from, to, text) => typeAtControlEdge(view, from, to, text),
      handleKeyDown: (view, event) => stepAcrossControlBorder(view, event),
    },
  });
}

/**
 * Text typed with the caret just outside a run control, at either edge, which the browser types inside
 * the control's element: inserted where the caret is, outside the control, as Word types it. Before,
 * typing after a control went into it; with its contents locked, it was refused (ED-005 section 12.60).
 */
export function typeAtControlEdge(view: EditorView, from: number, to: number, text: string): boolean {
  const sel = view.state.selection;
  if (!sel.empty || from !== to || from === sel.from) return false;
  const $caret = sel.$from;
  const inEnd = $caret.nodeBefore?.type.name === 'sdtInline' && from === sel.from - 1;
  const inStart = $caret.nodeAfter?.type.name === 'sdtInline' && from === sel.from + 1;
  if (!inEnd && !inStart) return false;
  view.dispatch(view.state.tr.insertText(text, sel.from).scrollIntoView());
  return true;
}
