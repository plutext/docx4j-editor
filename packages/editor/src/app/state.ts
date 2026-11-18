// SPDX-License-Identifier: GPL-3.0-or-later
//
// Rendering from the model (ED-002 section 5.4): React subscribes to the view and re-renders
// on every transaction; it holds no document state of its own.
import { useRef, useSyncExternalStore } from 'react';
import type { EditorState, Selection } from 'prosemirror-state';
import { NodeSelection, TextSelection } from 'prosemirror-state';
import { commandNamed, isInTable, rowShadingOf, controlAt, paragraphsIn, listKindOf, numberingOf, type EditorDocument, type SelectionFormatting, type ListKind, type ParagraphAttrs } from '@docx4j/editor-model';
import type { EditorSurface } from '../view/surface.js';

const noop = (): void => {};

export function useEditorState(editor: EditorSurface | null): EditorState | null {
  return useSyncExternalStore(
    (onChange) => (editor ? editor.subscribe(() => onChange()) : noop),
    () => editor?.state ?? null,
    () => null,
  );
}

/** The commands whose enabled state the toolbar shows. */
const TOOLBAR_ENABLED = ['history.undo', 'history.redo', 'paragraph.decreaseIndent', 'list.outdent', 'range.hyperlink', 'insertBreak', 'range.insertFootnote', 'range.insertEndnote', 'contentControl.insert', 'contentControl.wrapRows', 'contentControl.wrapCells', 'table.mergeCells', 'table.splitCell', 'table.insert'] as const;

/** Everything the toolbar renders from (ED-003 section 8.2): a small value a keystroke rarely changes. */
export interface ToolbarReading {
  fmt: SelectionFormatting;
  enabled: Record<string, boolean>;
  headerRow: boolean;
  /** The caret is in a hyperlink (the Hyperlink button pressed). */
  inLink: boolean;
  inControl: boolean;
  inTable: boolean;
  /** The selected rows' own shading (`#RRGGBB`, null for none, 'Mixed'), for the Row shading menu's bar; undefined outside a table. */
  rowShading: string | null | 'Mixed' | undefined;
  listKind: ListKind | undefined;
}

function readToolbar(editor: EditorSurface): ToolbarReading {
  const state = editor.state;
  const enabled: Record<string, boolean> = {};
  for (const name of TOOLBAR_ENABLED) enabled[name] = commandNamed(name)?.isEnabled?.(state) ?? true;
  const first = paragraphsIn(state)[0];
  return {
    fmt: editor.formatting(),
    enabled,
    headerRow: commandNamed('row.header')?.isActive?.(state) === true,
    inLink: commandNamed('range.hyperlink')?.isActive?.(state) === true,
    inControl: controlAt(state) !== undefined,
    inTable: isInTable(state),
    rowShading: rowShadingOf(state),
    listKind: first ? listKindOf(first.node.attrs as ParagraphAttrs, numberingOf(state)) : undefined,
  };
}

/**
 * The toolbar's reading, the same object while its content is the same (ED-003 section 8.2):
 * React compares snapshots by identity, so a keystroke that changes no formatting, no enabled
 * state and no context re-renders nothing. The reading is computed per transaction (a fraction
 * of a millisecond) and compared by its JSON.
 */
export function useToolbarReading(editor: EditorSurface): ToolbarReading {
  const last = useRef<{ editor: EditorSurface; key: string; reading: ToolbarReading } | null>(null);
  const snapshot = (): ToolbarReading => {
    const reading = readToolbar(editor);
    const key = JSON.stringify(reading);
    if (last.current && last.current.editor === editor && last.current.key === key) return last.current.reading;
    last.current = { editor, key, reading };
    return reading;
  };
  return useSyncExternalStore((onChange) => editor.subscribe(() => onChange()), snapshot, snapshot);
}

const revisions = new WeakMap<EditorDocument, number>();

/**
 * A counter that moves when the store changes outside the body: the `controls` map (a group's
 * properties, an undo of them) and the package map (a part edited in place bumps its revision
 * through `markPartChanged`). Panels that read the controls map or the custom XML parts
 * re-render on it, as they do on the editor state.
 */
export function useStoreRevision(session: EditorDocument | null): number {
  return useSyncExternalStore(
    (onChange) => {
      if (!session) return noop;
      const bump = (): void => { revisions.set(session, (revisions.get(session) ?? 0) + 1); onChange(); };
      session.store.controls.observeDeep(bump);
      session.store.package.observe(bump);
      return () => { session.store.controls.unobserveDeep(bump); session.store.package.unobserve(bump); };
    },
    () => (session ? revisions.get(session) ?? 0 : 0),
    () => 0,
  );
}

/** Select the node that came from a source element (the banner's list, the diagnostics panel). */
export function selectSource(editor: EditorSurface, sourceId: string): boolean {
  const { doc } = editor.state;
  let found: { pos: number; size: number; isText: boolean } | undefined;
  doc.descendants((node, pos) => {
    if (found) return false;
    if (node.attrs['sourceId'] === sourceId) { found = { pos, size: node.nodeSize, isText: node.isTextblock }; return false; }
    return true;
  });
  if (!found) return false;
  let sel: Selection;
  if (found.isText) sel = TextSelection.create(doc, found.pos + 1, found.pos + found.size - 1);
  else {
    try { sel = NodeSelection.create(doc, found.pos); } catch { sel = TextSelection.create(doc, found.pos); }
  }
  editor.view.dispatch(editor.state.tr.setSelection(sel).scrollIntoView());
  editor.focus();
  return true;
}
