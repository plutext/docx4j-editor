// SPDX-License-Identifier: GPL-3.0-or-later
//
// The read-only elements grouped by kind (ED-003 section 8.6; ED-002 section 10.5 item 9's
// leftover): one row per kind with its count, opened to the addresses, a click on an address
// selecting the node; "in the selection" keeps the entries whose source element the selection
// touches. The banner's list and the Diagnostics tab are this component.
import { useMemo, useState } from 'react';
import type { EditorState } from 'prosemirror-state';
import type { OpaqueEntry } from '@docx4j/editor-model';
import type { EditorSurface } from '../../view/surface.js';
import { useEditorState, selectSource } from '../state.js';
import { useT } from '../i18n.js';

export interface KindGroup {
  kind: string;
  entries: OpaqueEntry[];
}

/** The entries by kind, the most frequent kind first, ties by name. */
export function groupByKind(entries: readonly OpaqueEntry[]): KindGroup[] {
  const map = new Map<string, OpaqueEntry[]>();
  for (const e of entries) {
    const list = map.get(e.kind);
    if (list) list.push(e); else map.set(e.kind, [e]);
  }
  return [...map].map(([kind, list]) => ({ kind, entries: list })).sort((a, b) => b.entries.length - a.entries.length || a.kind.localeCompare(b.kind));
}

/**
 * The source ids of every node the selection touches: the nodes in the range and the ancestors
 * around it (a cell's `vMerge` and a table's `tblGrid` are reported under the cell's and the
 * table's ids), and at a caret the atom on either side, since an invisible marker cannot be
 * selected by hand.
 */
export function sourceIdsInSelection(state: EditorState): Set<string> {
  const ids = new Set<string>();
  const add = (node: { attrs: Record<string, unknown> } | null | undefined): void => {
    const id = node?.attrs['sourceId'];
    if (typeof id === 'string') ids.add(id);
  };
  const { from, to, $from } = state.selection;
  state.doc.nodesBetween(from, to, (node) => { add(node); return true; });
  for (let d = 0; d <= $from.depth; d++) add($from.node(d));
  if (from === to) { add($from.nodeBefore); add($from.nodeAfter); }
  return ids;
}

export interface DiagnosticListProps {
  editor: EditorSurface;
  entries: readonly OpaqueEntry[];
  testId?: string;
}

export function DiagnosticList({ editor, entries, testId = 'diagnostic-list' }: DiagnosticListProps) {
  const t = useT();
  const [inSelection, setInSelection] = useState(false);
  // Subscribed to the view only while the filter is on: a list of hundreds is not re-rendered per keystroke otherwise.
  const state = useEditorState(inSelection ? editor : null);
  const ids = useMemo(() => (inSelection && state ? sourceIdsInSelection(state) : null), [inSelection, state]);
  const groups = useMemo(() => groupByKind(ids ? entries.filter((e) => ids.has(e.sourceId)) : entries), [entries, ids]);
  return (
    <div className="diagnostic-list" data-testid={testId}>
      <label className="meta filter"><input type="checkbox" data-testid={`${testId}-in-selection`} checked={inSelection} onChange={(e) => setInSelection(e.target.checked)} /> {t('diagnostics.inSelection')}</label>
      {groups.length === 0 && <p className="meta">{inSelection ? t('diagnostics.noneInSelection') : t('diagnostics.none')}</p>}
      {groups.map((g) => (
        <details key={g.kind} className="diagnostic-kind" data-kind={g.kind}>
          <summary><code>{g.kind}</code> <span className="count">{g.entries.length}</span></summary>
          <ul>
            {g.entries.map((e, i) => (
              <li key={`${e.sourceId}-${i}`}>
                <button type="button" className="link address" aria-label={t('diagnostics.address', { kind: e.kind, address: e.address })} onClick={() => selectSource(editor, e.sourceId)}>{e.address}</button>
                {e.inline ? '' : <span className="meta">{t('diagnostics.block')}</span>}
              </li>
            ))}
          </ul>
        </details>
      ))}
    </div>
  );
}
