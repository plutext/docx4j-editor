// SPDX-License-Identifier: GPL-3.0-or-later
//
// Word's column selection (ED-005 section 12.107): the pointer over a table's top border, above a
// column, is a down arrow, and a click there selects the column; dragging along the border takes
// in the columns crossed. The columns are the table's grid columns, read from the `<colgroup>`
// the table view keeps in step with the cells, so a merged cell is selected with every column it
// covers, as Word selects it.
import { Plugin, PluginKey } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { CellSelection, TableMap } from 'prosemirror-tables';

export const columnSelectKey = new PluginKey('docx4j-column-select');

/** The class on a table while the pointer is on its top border. */
export const COLUMN_ARROW_CLASS = 'column-arrow';

/** How far above the top border the arrow still shows, and how far below it, in pixels. */
export const BAND_ABOVE = 5;
export const BAND_BELOW = 2;

/** The table whose top border holds the point, or null: the table under the point, else the one just below it. */
function tableOnTopBorder(view: EditorView, target: EventTarget | null, clientX: number, clientY: number): HTMLTableElement | null {
  const under = target instanceof Element ? target.closest('table') : null;
  const table = under && view.dom.contains(under) ? under : below(view, clientX, clientY);
  if (!table) return null;
  const top = table.getBoundingClientRect().top;
  return clientY >= top - BAND_ABOVE && clientY <= top + BAND_BELOW ? table : null;
}

/** The table just below the point, when the point is above its top border. */
function below(view: EditorView, clientX: number, clientY: number): HTMLTableElement | null {
  const el = view.dom.ownerDocument.elementFromPoint(clientX, clientY + BAND_ABOVE);
  const table = el?.closest('table') ?? null;
  return table && view.dom.contains(table) ? table : null;
}

/** The grid column whose box holds `clientX`, from the table's `<colgroup>`; the nearest edge column outside the table. */
export function columnAt(table: HTMLTableElement, clientX: number): number | undefined {
  const cols = Array.from(table.querySelectorAll(':scope > colgroup > col'));
  if (cols.length === 0) return undefined;
  const boxes = cols.map((col) => col.getBoundingClientRect());
  const i = boxes.findIndex((b) => clientX >= b.left && clientX <= b.right);
  if (i >= 0) return i;
  // Outside every column (a drag past the table's edge): the nearest column.
  let nearest = 0;
  let distance = Infinity;
  boxes.forEach((b, c) => {
    const d = clientX < b.left ? b.left - clientX : clientX - b.right;
    if (d < distance) { distance = d; nearest = c; }
  });
  return nearest;
}

/** The table node and the position of its first row for the table element. */
function tableAt(view: EditorView, table: HTMLTableElement): { node: PMNode; start: number } | undefined {
  const tbody = table.tBodies[0];
  if (!tbody) return undefined;
  let pos: number;
  try { pos = view.posAtDOM(tbody, 0); } catch { return undefined; }
  // The position before the first row lies inside the table, so the table is the innermost node around it.
  const $pos = view.state.doc.resolve(pos);
  for (let d = $pos.depth; d > 0; d--) {
    const node = $pos.node(d);
    if (node.type.spec.tableRole === 'table') return { node, start: $pos.start(d) };
  }
  return undefined;
}

/** The selection of the grid columns from `anchorCol` to `headCol`: the cells at the top and the bottom rows. */
export function columnSelection(doc: PMNode, table: PMNode, start: number, anchorCol: number, headCol: number): CellSelection {
  const map = TableMap.get(table);
  const clamp = (c: number): number => Math.max(0, Math.min(map.width - 1, c));
  const $anchor = doc.resolve(start + map.positionAt(0, clamp(anchorCol), table));
  const $head = doc.resolve(start + map.positionAt(map.height - 1, clamp(headCol), table));
  return CellSelection.colSelection($anchor, $head);
}

/**
 * The plugin, installed before `prosemirror-tables`' own handlers so that a click on the border
 * never starts their cell drag.
 */
export function columnSelectPlugin(): Plugin {
  let arrowOn: HTMLTableElement | null = null;
  const setArrow = (table: HTMLTableElement | null): void => {
    if (arrowOn === table) return;
    arrowOn?.classList.remove(COLUMN_ARROW_CLASS);
    table?.classList.add(COLUMN_ARROW_CLASS);
    arrowOn = table;
  };
  return new Plugin({
    key: columnSelectKey,
    props: {
      handleDOMEvents: {
        mousemove(view, event) {
          if (!view.editable || event.buttons !== 0) return false;
          setArrow(tableOnTopBorder(view, event.target, event.clientX, event.clientY));
          return false;
        },
        mouseleave() {
          setArrow(null);
          return false;
        },
        mousedown(view, event) {
          if (event.button !== 0 || !view.editable) return false;
          const table = tableOnTopBorder(view, event.target, event.clientX, event.clientY);
          if (!table) return false;
          const found = tableAt(view, table);
          const anchorCol = columnAt(table, event.clientX);
          if (!found || anchorCol === undefined) return false;
          event.preventDefault();
          const select = (headCol: number): void => {
            if (view.isDestroyed) return;
            const selection = columnSelection(view.state.doc, found.node, found.start, anchorCol, headCol);
            if (!view.state.selection.eq(selection)) view.dispatch(view.state.tr.setSelection(selection));
          };
          select(anchorCol);
          view.focus();
          // The drag along the border: the columns crossed join, the anchor column staying.
          const doc = view.dom.ownerDocument;
          const onMove = (e: MouseEvent): void => {
            const col = columnAt(table, e.clientX);
            if (col !== undefined) select(col);
          };
          const onUp = (): void => {
            doc.removeEventListener('mousemove', onMove);
            doc.removeEventListener('mouseup', onUp);
          };
          doc.addEventListener('mousemove', onMove);
          doc.addEventListener('mouseup', onUp);
          return true;
        },
      },
    },
    view() {
      return { destroy() { setArrow(null); } };
    },
  });
}
