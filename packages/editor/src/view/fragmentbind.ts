// SPDX-License-Identifier: GPL-3.0-or-later
//
// Word fragments shown (ED-005 proposal 46, part 3; section 12.119): the model's refresh (`refreshFragments`, the engine's
// bind pass in a tree run, out of the undo history) and then the view's stylesheet and list labels made again, since a fill
// copies styles and lists into the document's parts, which the view draws from. Run on the occasions the XHTML and picture
// refreshes run: the Control tab's switch and bind, the Data tab's bind and edits, the stale banner's Refresh.
import { refreshFragments, type EditorDocument, type FragmentBinding } from '@docx4j/editor-model';
import type { DocxEditorView } from './DocxEditorView.js';

export interface FragmentViewRefresh {
  changed: number;
  /** Nothing to say in the status line: the engine's notes are its own English, and a fill shows itself. */
  notes: string[];
}

export async function refreshFragmentControls(editor: DocxEditorView, session: EditorDocument, options: { ids?: ReadonlySet<number>; shape?: FragmentBinding['shape'] } = {}): Promise<FragmentViewRefresh> {
  const result = await refreshFragments(session, options);
  if (result.changed > 0) editor.partsChanged();
  return { changed: result.changed, notes: [] };
}
