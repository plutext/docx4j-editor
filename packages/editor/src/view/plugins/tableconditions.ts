// SPDX-License-Identifier: GPL-3.0-or-later
//
// A table style's conditional formatting on the page (ED-005 section 12.65): each cell its table style
// formats is given the class of its conditions (`tc-<style>-<mask>`, `tableConditionsOf`), and a paragraph
// whose own `w:cnfStyle` puts its text under others the class of those (`tp-<style>-<mask>`), which the
// document stylesheet draws (css.ts: the cell's shading, borders and margins, and its paragraphs' text,
// with the table level resolved in). A node decoration per such cell, nested in its table, never one per
// top-level block: the set is mapped on a change and rebuilt only for the top-level blocks the change
// replaced, as `decorations.ts` memoises its walk (ED-003 section 8.2). A styles part changed (an Apply,
// a peer) rebuilds it whole.
import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { tableConditionsOf, type EditorDocument } from '@docx4j/editor-model';

/** A cell's decoration within its top-level block: offsets from the block's start, and its class. */
interface CellMark { from: number; to: number; key: string }

interface Entry { cells: CellMark[]; keys: string[] }

export interface TableConditionsState {
  set: DecorationSet;
  memo: WeakMap<PMNode, Entry>;
  /** The classes of the blocks this transaction rebuilt: the view checks its stylesheet has rules for them. */
  changedKeys: ReadonlySet<string>;
}

export const tableConditionsKey = new PluginKey<TableConditionsState>('docx4j-table-conditions');

/** The meta that rebuilds every table's classes: the styles changed (`DocxEditorView.partsChanged`). */
export const TABLE_STYLES_CHANGED = 'docx4j-table-styles-changed';

const EMPTY: ReadonlySet<string> = new Set();

/** The cells a top-level block's tables (nested ones included) give a class, and the classes. */
function entryOf(session: EditorDocument, block: PMNode): Entry {
  const cells: CellMark[] = [];
  const keys = new Set<string>();
  const visit = (table: PMNode, start: number): void => {
    const conditions = tableConditionsOf(session, table);
    if (conditions.cells.size === 0 && conditions.paragraphs.size === 0) return;
    // The table's own cells, and the paragraphs whose own w:cnfStyle gives their text other conditions.
    table.descendants((n, pos) => {
      const c = conditions.cells.get(n) ?? conditions.paragraphs.get(n);
      if (c) {
        const from = start + 1 + pos;
        cells.push({ from, to: from + n.nodeSize, key: c.key });
        keys.add(c.key);
      }
      return n.type.name !== 'table' && !n.isTextblock;
    });
  };
  if (block.type.name === 'table') visit(block, 0);
  block.descendants((n, pos) => {
    if (n.type.name === 'table') visit(n, pos + 1);
    return n.isBlock;
  });
  return { cells, keys: [...keys] };
}

function rebuild(session: EditorDocument, doc: PMNode, prev?: TableConditionsState, tr?: Transaction): TableConditionsState {
  const memo = prev && tr ? prev.memo : new WeakMap<PMNode, Entry>();
  let set = prev && tr ? prev.set.map(tr.mapping, doc) : DecorationSet.empty;
  const changedKeys = new Set<string>();
  const add: Decoration[] = [];
  doc.forEach((block, pos) => {
    if (prev && tr && memo.has(block)) return;
    if (prev && tr) set = set.remove(set.find(pos, pos + block.nodeSize));
    const entry = entryOf(session, block);
    memo.set(block, entry);
    for (const c of entry.cells) add.push(Decoration.node(pos + c.from, pos + c.to, { class: c.key }));
    for (const k of entry.keys) changedKeys.add(k);
  });
  if (add.length > 0) set = set.add(doc, add);
  return { set, memo, changedKeys };
}

/** Every class the document's tables use, for the stylesheet's check (a block's memo, walked; no table read again). */
export function tableClassesOf(state: EditorState): Set<string> {
  const st = tableConditionsKey.getState(state);
  const out = new Set<string>();
  if (!st) return out;
  state.doc.forEach((block) => { for (const k of st.memo.get(block)?.keys ?? []) out.add(k); });
  return out;
}

export function tableConditionsPlugin(session: EditorDocument): Plugin<TableConditionsState> {
  return new Plugin<TableConditionsState>({
    key: tableConditionsKey,
    state: {
      init: (_config, state) => rebuild(session, state.doc),
      apply: (tr, prev, _old, state) => {
        if (tr.getMeta(TABLE_STYLES_CHANGED)) return rebuild(session, state.doc);
        if (!tr.docChanged) return prev.changedKeys.size === 0 ? prev : { ...prev, changedKeys: EMPTY };
        return rebuild(session, state.doc, prev, tr);
      },
    },
    props: {
      decorations: (state) => tableConditionsKey.getState(state)?.set ?? DecorationSet.empty,
    },
  });
}
