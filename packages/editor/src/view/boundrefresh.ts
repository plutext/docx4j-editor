// SPDX-License-Identifier: GPL-3.0-or-later
//
// What a bind shows at once (ED-005 sections 12.88 and 12.124): the control just bound, at the caret, given its node as the
// view shows bound content, XHTML converted, a picture placed, a Word document filled. The Data tab's bind of a control in
// place and a node dropped on a control both run it. A picture or Word fragment control elsewhere is left as it is: the
// stale banner's Keep left it unfilled on purpose, and a bind here is no Refresh there.
import { controlAt, sdtPrOf, sdtId, type EditorDocument } from '@docx4j/editor-model';
import type { DocxEditorView } from './DocxEditorView.js';
import { refreshXhtmlControls } from './xhtmlbind.js';
import { refreshPictureControls } from './picturebind.js';
import { refreshFragmentControls } from './fragmentbind.js';

/** The refreshes for the control at the caret; what they could not do, each note once, for the status line. */
export async function refreshBoundAtCaret(editor: DocxEditorView, session: EditorDocument): Promise<string[]> {
  const target = controlAt(editor.state);
  const id = target ? sdtId(sdtPrOf(target, session.controlEntry)) : undefined;
  const ids = id !== undefined ? new Set([id]) : undefined;
  const results = await Promise.all([
    refreshXhtmlControls(editor.view, session),
    refreshPictureControls(editor.view, session, undefined, ids),
    refreshFragmentControls(editor, session, ids ? { ids } : {}),
  ]);
  return [...new Set(results.flatMap((r) => r.notes))];
}
