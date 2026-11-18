// SPDX-License-Identifier: GPL-3.0-or-later
//
// Tables in the view (ED-003 section 3.6): `prosemirror-tables`' cell selection and column
// resizing, and the plugin that keeps row-level control groups contiguous after any change
// (ED-001 section 4.1's invariant), so a row inserted inside a group by any means joins it.
//
// Both repairs run on local changes only (ED-004 section 3.5): a repair is made once, by the peer
// whose change broke the table, and travels like any edit. A transaction y-prosemirror makes from
// the Yjs side (a remote update, an undo or redo, the first binding of a fragment) is not
// repaired: an undo restores a state that was valid when it was made, and a repair at the first
// binding would mark an untouched document as edited.
import { touches } from './labels.js';
import { Plugin, PluginKey, type Transaction } from 'prosemirror-state';
import { ySyncPluginKey } from 'y-prosemirror';
import { columnResizing, tableEditing } from 'prosemirror-tables';
import { normalizeRowGroups, endCellsWithParagraph } from '@docx4j/editor-model';
import { CELL_MIN_WIDTH } from '../nodeviews/index.js';

export const rowGroupsKey = new PluginKey('docx4j-row-groups');

/** Whether every transaction that changed the document came from the Yjs side (y-prosemirror's `isChangeOrigin`). */
export function fromYjs(transactions: readonly Transaction[]): boolean {
  const changing = transactions.filter((tr) => tr.docChanged);
  return changing.length > 0 && changing.every((tr) => (tr.getMeta(ySyncPluginKey) as { isChangeOrigin?: boolean } | undefined)?.isChangeOrigin === true);
}

/**
 * A library plugin whose `appendTransaction` cannot be told about Yjs, rebuilt from its spec with
 * the guard in front. For `tableEditing` that skips its selection normalisation after a Yjs-side
 * change as well as its table fix-up; the selection is normalised again on the next local one.
 */
function localOnly(plugin: Plugin): Plugin {
  const append = plugin.spec.appendTransaction;
  if (!append) return plugin;
  return new Plugin({
    ...plugin.spec,
    appendTransaction(transactions, oldState, state) {
      return fromYjs(transactions) ? null : append.call(this, transactions, oldState, state);
    },
  });
}

export function rowGroupsPlugin(): Plugin {
  return new Plugin({
    key: rowGroupsKey,
    appendTransaction(transactions, _old, state) {
      if (fromYjs(transactions)) return null;
      // Only a change in or around a table can break a group's contiguity (ED-003 section 8.2: the whole-document pass ran on every keystroke).
      if (!transactions.some((tr) => tr.docChanged && touches(tr, (n) => n.type.name === 'table'))) return null;
      const tr = normalizeRowGroups(state.tr);
      return tr && tr.docChanged ? tr : null;
    },
  });
}

export const cellEndsKey = new PluginKey('docx4j-cell-ends');

/**
 * A cell's last block stays a paragraph (ED-005 section 12.109; Word refuses a `w:tc` ending in a
 * `w:tbl` and repairs it so): a paste or an Insert Table that leaves a table last in a cell gets an
 * empty paragraph after it, on local changes only, as the row groups are kept. The headless runner
 * does the same in the command's own transaction (`endCellsWithParagraph` in `run.mts`).
 */
export function cellEndsPlugin(): Plugin {
  return new Plugin({
    key: cellEndsKey,
    appendTransaction(transactions, _old, state) {
      if (fromYjs(transactions)) return null;
      if (!transactions.some((tr) => tr.docChanged && touches(tr, (n) => n.type.name === 'table'))) return null;
      const tr = endCellsWithParagraph(state.tr, true);
      return tr && tr.docChanged ? tr : null;
    },
  });
}

/** The table plugins, in the order `prosemirror-tables` wants: resizing before editing. */
export function tablesPlugins(): Plugin[] {
  // The table node view is the editor's own (nodeviews/index.ts), so the resizer registers none.
  return [columnResizing({ cellMinWidth: CELL_MIN_WIDTH, defaultCellMinWidth: CELL_MIN_WIDTH, View: null }), localOnly(tableEditing()), rowGroupsPlugin(), cellEndsPlugin()];
}
