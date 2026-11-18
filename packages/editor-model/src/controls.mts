// SPDX-License-Identifier: GPL-3.0-or-later
//
// Where the content controls are (ED-003 section 4.5's tree, 4.3's commands): block and run
// controls are nodes (`sdtBlock`, `sdtInline`); row and cell controls are groups on the rows'
// and cells' `sdt` stacks with their properties in the `controls` map (section 4.2). One walk
// lists them in document order with their nesting; one lookup finds the innermost control
// around a selection.
import type { Node as PMNode, ResolvedPos } from 'prosemirror-model';
import type { EditorState } from 'prosemirror-state';
import { CellSelection } from 'prosemirror-tables';
import type { SdtGroup } from './import.mjs';
import type { ControlLookup } from './export.mjs';
import { readControl, type ControlReading } from './sdtpr.mjs';
import { readTag, type TagReading } from './opendope.mjs';

export type ControlForm = 'block' | 'run' | 'row' | 'cell';

/** One control of the document, as the tree shows it. */
export interface ControlSummary {
  form: ControlForm;
  /** A node's position (before it); undefined for a group. */
  pos?: number;
  /** A group's id, and the table it is in. */
  groupId?: string;
  tablePos?: number;
  /** A row group's first and last row index (inclusive); a cell group's row. */
  rows?: [number, number];
  /** A cell group's first and last cell index in its row (inclusive). */
  cells?: [number, number];
  /** Nesting depth: controls and groups above this one. */
  depth: number;
  sourceId: string | null;
  sdtPr: unknown | null;
  reading: ControlReading;
  tag: TagReading;
}

/** Every control in document order, nested ones after their parent with a greater depth. */
export function listControls(doc: PMNode, controls: ControlLookup): ControlSummary[] {
  const out: ControlSummary[] = [];
  const summary = (base: Omit<ControlSummary, 'reading' | 'tag'>): ControlSummary => ({ ...base, reading: readControl(base.sdtPr), tag: readTag(readControl(base.sdtPr).tag) });
  const walk = (node: PMNode, pos: number, depth: number): void => {
    if (node.type.name === 'table') { walkTable(node, pos, depth); return; }
    node.forEach((child, offset) => {
      const childPos = pos + 1 + offset;
      if (child.type.name === 'sdtBlock' || child.type.name === 'sdtInline') {
        out.push(summary({ form: child.type.name === 'sdtBlock' ? 'block' : 'run', pos: childPos, depth, sourceId: child.attrs['sourceId'] as string | null, sdtPr: child.attrs['sdtPr'] as unknown }));
        walk(child, childPos, depth + 1);
      } else if (child.isBlock || child.type.name === 'table') {
        walk(child, childPos, depth);
      }
    });
  };
  const walkTable = (table: PMNode, tablePos: number, depth: number): void => {
    // A row group's extent: the rows carrying it.
    const extents = new Map<string, [number, number]>();
    table.forEach((row, _o, r) => {
      for (const g of (row.attrs['sdt'] as SdtGroup[] | null) ?? []) {
        const e = extents.get(g.groupId);
        if (!e) extents.set(g.groupId, [r, r]); else e[1] = r;
      }
    });
    const seen = new Set<string>();
    let rowPos = tablePos + 1;
    table.forEach((row, _o, r) => {
      const stack = (row.attrs['sdt'] as SdtGroup[] | null) ?? [];
      stack.forEach((g, d) => {
        if (seen.has(g.groupId)) return;
        seen.add(g.groupId);
        const entry = controls(g.groupId);
        out.push(summary({ form: 'row', groupId: g.groupId, tablePos, rows: extents.get(g.groupId)!, depth: depth + d, sourceId: g.groupId.startsWith('new:') ? null : g.groupId, sdtPr: entry?.sdtPr ?? null }));
      });
      const cellExtents = new Map<string, [number, number]>();
      row.forEach((cell, _co, c) => {
        for (const g of (cell.attrs['sdt'] as SdtGroup[] | null) ?? []) {
          const e = cellExtents.get(g.groupId);
          if (!e) cellExtents.set(g.groupId, [c, c]); else e[1] = c;
        }
      });
      const cellSeen = new Set<string>();
      let cellPos = rowPos + 1;
      row.forEach((cell, _co, c) => {
        const cellStack = (cell.attrs['sdt'] as SdtGroup[] | null) ?? [];
        cellStack.forEach((g, d) => {
          if (cellSeen.has(g.groupId)) return;
          cellSeen.add(g.groupId);
          const entry = controls(g.groupId);
          out.push(summary({ form: 'cell', groupId: g.groupId, tablePos, rows: [r, r], cells: cellExtents.get(g.groupId)!, depth: depth + stack.length + d, sourceId: g.groupId.startsWith('new:') ? null : g.groupId, sdtPr: entry?.sdtPr ?? null }));
        });
        void c;
        walk(cell, cellPos, depth + stack.length + cellStack.length);
        cellPos += cell.nodeSize;
      });
      rowPos += row.nodeSize;
    });
  };
  walk(doc, -1, 0);
  return out;
}

/** The innermost control around a selection, for the commands and the panel. */
export type NodeTarget = { form: 'block' | 'run'; pos: number; node: PMNode };
export type RowGroupTarget = { form: 'row'; groupId: string; tablePos: number; table: PMNode; rows: [number, number] };
export type CellGroupTarget = { form: 'cell'; groupId: string; tablePos: number; table: PMNode; row: number; cells: [number, number] };
export type ControlTarget = NodeTarget | RowGroupTarget | CellGroupTarget;

export function isNodeTarget(target: ControlTarget): target is NodeTarget {
  return target.form === 'block' || target.form === 'run';
}

/** The rows carrying a group in a table, as first and last index. */
export function rowExtent(table: PMNode, groupId: string): [number, number] | undefined {
  let first = -1;
  let last = -1;
  table.forEach((row, _o, r) => {
    if (((row.attrs['sdt'] as SdtGroup[] | null) ?? []).some((g) => g.groupId === groupId)) { if (first < 0) first = r; last = r; }
  });
  return first < 0 ? undefined : [first, last];
}

export function cellExtent(row: PMNode, groupId: string): [number, number] | undefined {
  let first = -1;
  let last = -1;
  row.forEach((cell, _o, c) => {
    if (((cell.attrs['sdt'] as SdtGroup[] | null) ?? []).some((g) => g.groupId === groupId)) { if (first < 0) first = c; last = c; }
  });
  return first < 0 ? undefined : [first, last];
}

/**
 * The innermost control around the selection: a node, else the caret's cell group, else its row
 * group. A cell selection (rows or cells selected as a block) takes the innermost group that
 * covers every selected cell: the row group when a whole repeat row is selected, even though
 * its first cell carries a control of its own.
 */
export function controlAt(state: EditorState): ControlTarget | undefined {
  if (state.selection instanceof CellSelection) return controlOfCells(state.selection);
  return controlAround(state.selection.$from);
}

function controlOfCells(sel: CellSelection): ControlTarget | undefined {
  const $a = sel.$anchorCell;
  const $h = sel.$headCell;
  const table = $a.node(-1);
  const tablePos = $a.before(-1);
  const rows: [number, number] = [Math.min($a.index(-1), $h.index(-1)), Math.max($a.index(-1), $h.index(-1))];
  if (rows[0] === rows[1]) {
    const row = $a.node();
    const cells: [number, number] = [Math.min($a.index(), $h.index()), Math.max($a.index(), $h.index())];
    const stack = ($a.nodeAfter?.attrs['sdt'] as SdtGroup[] | null) ?? [];
    for (let i = stack.length - 1; i >= 0; i--) {
      const extent = cellExtent(row, stack[i]!.groupId);
      if (extent && extent[0] <= cells[0] && extent[1] >= cells[1]) return { form: 'cell', groupId: stack[i]!.groupId, tablePos, table, row: rows[0], cells: extent };
    }
  }
  const stack = (table.child(rows[0]).attrs['sdt'] as SdtGroup[] | null) ?? [];
  for (let i = stack.length - 1; i >= 0; i--) {
    const extent = rowExtent(table, stack[i]!.groupId);
    if (extent && extent[0] <= rows[0] && extent[1] >= rows[1]) return { form: 'row', groupId: stack[i]!.groupId, tablePos, table, rows: extent };
  }
  return undefined;
}

/** The innermost control around a resolved position (the view's drop handler, the panels). */
export function controlAround($from: ResolvedPos): ControlTarget | undefined {
  for (let d = $from.depth; d >= 1; d--) {
    const node = $from.node(d);
    const pos = $from.before(d);
    if (node.type.name === 'sdtInline' || node.type.name === 'sdtBlock') return { form: node.type.name === 'sdtBlock' ? 'block' : 'run', pos, node };
    if (node.type.name === 'table_cell') {
      const stack = (node.attrs['sdt'] as SdtGroup[] | null) ?? [];
      const innermost = stack[stack.length - 1];
      if (innermost) {
        const row = $from.node(d - 1);
        const table = $from.node(d - 2);
        return { form: 'cell', groupId: innermost.groupId, tablePos: $from.before(d - 2), table, row: $from.index(d - 2), cells: cellExtent(row, innermost.groupId)! };
      }
    }
    if (node.type.name === 'table_row') {
      const stack = (node.attrs['sdt'] as SdtGroup[] | null) ?? [];
      const innermost = stack[stack.length - 1];
      if (innermost) {
        const table = $from.node(d - 1);
        return { form: 'row', groupId: innermost.groupId, tablePos: $from.before(d - 1), table, rows: rowExtent(table, innermost.groupId)! };
      }
    }
  }
  return undefined;
}

/** The properties of a target: the node's attribute, or the group's map entry. */
export function sdtPrOf(target: ControlTarget, controls: ControlLookup): unknown | null {
  if (isNodeTarget(target)) return target.node.attrs['sdtPr'] as unknown;
  return controls(target.groupId)?.sdtPr ?? null;
}

/** The position of row `r` of a table at `tablePos`. */
export function rowPosition(table: PMNode, tablePos: number, r: number): number {
  let pos = tablePos + 1;
  for (let i = 0; i < r; i++) pos += table.child(i).nodeSize;
  return pos;
}
