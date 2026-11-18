// SPDX-License-Identifier: GPL-3.0-or-later
//
// Word's selection bar (ED-005 section 12.74): a click in the page's left margin, beside a line of
// a paragraph, selects that line; when the line is the paragraph's last (a paragraph on one line
// included) the selection takes the paragraph mark too, running to the start of the next paragraph,
// so that typing or Delete removes the mark as Word's does. The margin is the page container's
// padding, outside the editor's element, which ProseMirror never hears about.
import { TextSelection, type Selection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { ResolvedPos } from 'prosemirror-model';
import { CellSelection } from 'prosemirror-tables';

/** The depth of the row around a position, or -1 (ED-005 section 12.81: a click beside a table row selects the row). */
function rowDepth($pos: ResolvedPos): number {
  for (let d = $pos.depth; d > 0; d--) if ($pos.node(d).type.name === 'table_row') return d;
  return -1;
}

/** The depth of the paragraph (the nearest textblock) around a position, or -1: a position inside an inline control is in its paragraph's line. */
function textblockDepth($pos: ResolvedPos): number {
  for (let d = $pos.depth; d > 0; d--) if ($pos.node(d).isTextblock) return d;
  return -1;
}

/** The row's cells as one selection: its first cell to its last. */
function rowSelection($pos: ResolvedPos, depth: number): CellSelection {
  const row = $pos.node(depth);
  const rowStart = $pos.start(depth);
  let last = rowStart;
  row.forEach((_cell, offset) => { last = rowStart + offset; });
  return CellSelection.create($pos.doc, rowStart, last);
}

/** The selection a margin click at `clientY` makes, or undefined where no line is beside it. */
export function lineSelectionAt(view: EditorView, clientY: number): Selection | undefined {
  const box = view.dom.getBoundingClientRect();
  if (clientY < box.top || clientY > box.bottom) return undefined;
  const a = view.posAtCoords({ left: box.left + 1, top: clientY });
  const b = view.posAtCoords({ left: box.right - 1, top: clientY });
  if (!a || !b) return undefined;
  const { doc } = view.state;
  let from = Math.min(a.pos, b.pos);
  let to = Math.max(a.pos, b.pos);
  // Beside a table row (either edge of the line inside it, the table's own border included): the row.
  for (const p of [a, b]) {
    const $p = doc.resolve(p.pos);
    const depth = rowDepth($p);
    if (depth >= 0) return rowSelection($p, depth);
    const $inside = p.inside >= 0 ? doc.resolve(p.inside) : undefined;
    if ($inside && $inside.nodeAfter?.type.name === 'table') {
      // The border of a table: the row at this height is the one whose box holds the point.
      const rows = rowAt(view, $inside.pos, clientY);
      if (rows !== undefined) return rows;
    }
  }
  const $from = doc.resolve(from);
  const depth = textblockDepth($from);
  if (depth < 0) return undefined;
  const start = $from.start(depth);
  const end = $from.end(depth);
  // An end of the line inside an inline node (a control): the whole node, as Word selects it.
  if ($from.depth > depth) from = $from.before(depth + 1);
  const $to = doc.resolve(to);
  if ($to.depth > depth && $to.sharedDepth(from) >= depth) to = $to.after(depth + 1);
  from = Math.max(from, start);
  to = Math.min(to, end);
  if (to === end) {
    // The paragraph's last line: the mark too, up to the next paragraph's start when one follows.
    const $after = doc.resolve($from.after(depth));
    if ($after.nodeAfter?.isTextblock) to = $after.pos + 1;
  }
  return TextSelection.create(doc, from, to);
}

/** The row of the table at `tablePos` whose box holds `clientY`, as a selection. */
function rowAt(view: EditorView, tablePos: number, clientY: number): CellSelection | undefined {
  const table = view.state.doc.nodeAt(tablePos);
  if (!table) return undefined;
  let pos = tablePos + 1;
  for (let r = 0; r < table.childCount; r++) {
    const row = table.child(r);
    const dom = view.nodeDOM(pos) as HTMLElement | null;
    const box = dom?.getBoundingClientRect();
    if (box && clientY >= box.top && clientY <= box.bottom) return rowSelection(view.state.doc.resolve(pos + 1), view.state.doc.resolve(pos + 1).depth);
    pos += row.nodeSize;
  }
  return undefined;
}

/** Installs the selection bar on the page container around `view.dom`; returns the remover. */
export function installSelectionBar(view: EditorView, container: HTMLElement): () => void {
  const onMouseDown = (event: MouseEvent): void => {
    if (event.button !== 0 || event.target !== container || view.isDestroyed || !view.editable) return;
    const box = view.dom.getBoundingClientRect();
    if (event.clientX >= box.left) return;   // the left margin alone
    const selection = lineSelectionAt(view, event.clientY);
    if (!selection) return;
    event.preventDefault();
    view.dispatch(view.state.tr.setSelection(selection).scrollIntoView());
    view.focus();
  };
  container.addEventListener('mousedown', onMouseDown);
  return () => container.removeEventListener('mousedown', onMouseDown);
}
