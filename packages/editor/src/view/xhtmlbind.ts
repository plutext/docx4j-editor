// SPDX-License-Identifier: GPL-3.0-or-later
//
// XHTML-bound controls shown (ED-005 proposal 44, section 12.75): a control whose tag binds it to a
// node holding XHTML markup (`isXhtmlControl`, the model's `xhtml.mts`) is given the markup converted,
// through the engine's converter (`@docx4j/core-ts/html`, re-exported by `paste.ts`), with the list
// definitions, relationships and styles the markup needs made first, as a paste's are. The Preview tab
// converts through the engine's bind step instead (ED-005 section 12.117).
// The write goes to the store out of the undo history, as the text bindings' refresh does
// (`applyBoundRefresh`), and nothing is written back to the node: an edit inside such a control is
// the document's alone. Not at open: a file round-trips as it is; the conversion runs when the
// control is bound to XHTML in the Control tab, when the Data tab changes a node, and when a text
// control bound to the same node writes it.
import { Fragment, type Node as PMNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { type EditorDocument, xhtmlControlsIn, boundNodeOf, isFlatOpc, sdtId, applyControlProperties, readControl, UNTRACKED_META, BOUND_WRITE_META } from '@docx4j/editor-model';
import { convertHtml, styleLookupFor, prepareConverted, toNodes, describeDropped, type PasteReport } from './paste.js';

export interface XhtmlRefresh {
  /** Controls whose content was written. */
  changed: number;
  /** What the markup had that the control could not hold, one line per control (the status line's). */
  notes: string[];
}

/** The value each control was last filled from, by its `w:id`: a refresh over an unchanged value converts nothing and makes no definitions. */
const filled = new WeakMap<EditorDocument, Map<number, string>>();

export async function refreshXhtmlControls(view: EditorView, session: EditorDocument, only?: ReadonlySet<Node>): Promise<XhtmlRefresh> {
  const result: XhtmlRefresh = { changed: 0, notes: [] };
  if (!session.template.loaded) return result;
  const controls = xhtmlControlsIn(view.state.doc, session.template);
  if (controls.length === 0) return result;
  let cache = filled.get(session);
  if (!cache) { cache = new Map(); filled.set(session, cache); }
  const lookup = await styleLookupFor(session);
  const built: { id: number | undefined; node: PMNode; pos: number; content: Fragment; report: PasteReport; value: string }[] = [];
  for (const c of controls) {
    const data = boundNodeOf(session, c.binding);
    if (!data || (only && !only.has(data.node))) continue;
    const value = data.nodeValue;
    if (isFlatOpc(value)) continue;
    const id = sdtId(c.node.attrs['sdtPr']);
    if (id !== undefined && cache.get(id) === value) continue;
    const converted = convertHtml(value, lookup);
    const prepared = await prepareConverted(view, session, converted);
    if ((view as { isDestroyed?: boolean }).isDestroyed) return result;
    let nodes = toNodes(prepared.kept, view.state.schema, prepared.lists, prepared.links);
    let content: Fragment;
    if (c.node.type.name === 'sdtInline') {
      // A run-level control holds inline content: the first paragraph's, the rest reported (docx4j's rule).
      const first = nodes.find((n) => n.type.name === 'paragraph');
      const rest = nodes.length - (first ? 1 : 0);
      if (rest > 0) prepared.report.dropped['blocks after the first, in a run-level control'] = rest;
      content = first ? first.content : Fragment.empty;
    } else {
      if (nodes.length === 0) nodes = [view.state.schema.nodes['paragraph']!.createAndFill()!];
      content = Fragment.from(nodes);
    }
    built.push({ id, node: c.node, pos: c.pos, content, report: prepared.report, value });
  }
  if (built.length === 0) return result;
  // The transaction over the state the view holds now: the document may have moved on during the awaits, so each control is found again.
  const state = view.state;
  const tr = state.tr;
  for (const b of built.reverse()) {
    let pos = b.pos;
    if (state.doc.nodeAt(pos) !== b.node) {
      pos = -1;
      state.doc.descendants((n, p) => { if (pos >= 0) return false; if (n === b.node) { pos = p; return false; } return true; });
      if (pos < 0) continue;
    }
    if (b.id !== undefined) cache.set(b.id, b.value);
    if (b.node.content.eq(b.content)) continue;
    tr.replaceWith(pos + 1, pos + b.node.nodeSize - 1, b.content);
    const sdtPr = b.node.attrs['sdtPr'];
    if (readControl(sdtPr).showingPlaceholder) tr.setNodeMarkup(pos, undefined, { ...b.node.attrs, sdtPr: applyControlProperties(sdtPr, { showingPlaceholder: false }) });
    result.changed++;
    if (Object.keys(b.report.dropped).length > 0) result.notes.push(describeDropped(b.report.dropped));
  }
  if (!tr.docChanged) return result;
  tr.setMeta(UNTRACKED_META, true).setMeta('addToHistory', false).setMeta(BOUND_WRITE_META, true);
  session.replaceOutOfHistory(tr.doc, state.doc);
  return result;
}
