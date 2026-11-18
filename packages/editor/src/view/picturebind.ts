// SPDX-License-Identifier: GPL-3.0-or-later
//
// Picture-bound controls shown (ED-005 proposal 45 part 3, section 12.76): a Word picture control
// bound to a node holding base64 image data, or a rich text control tagged `od:Handler=picture`
// with a `width`, is given the image: the part and the drawing made as a picture inserted by hand is
// (`session.addInlinePicture`), a block-level control holding it in a paragraph, a run-level one
// inline. The write goes to the store out of the undo history, as the XHTML refresh's does, and
// nothing is written back. A handler control without `width` keeps a floating drawing the editor
// holds opaque (REQ-061: the engine's bind step replaces its blip, released since core-ts 0.4.0 and
// not yet run here in a tree run); it is left and noted. Not at open: a file round-trips as it is.
import { Fragment, type Node as PMNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { type EditorDocument, pictureControlsIn, boundNodeOf, sdtId, applyControlProperties, readControl, UNTRACKED_META, BOUND_WRITE_META, looksLikeImageBase64, fittedInBox } from '@docx4j/editor-model';
export { looksLikeImageBase64 };
import { t } from '../app/i18n.js';

export interface PictureRefresh {
  changed: number;
  notes: string[];
}

const filled = new WeakMap<EditorDocument, Map<number, string>>();

/** The first picture a control holds, whose box Word fits the node's picture inside. */
function firstImage(control: PMNode): PMNode | undefined {
  let image: PMNode | undefined;
  control.descendants((n) => { if (image) return false; if (n.type.name === 'image') { image = n; return false; } return true; });
  return image;
}


/**
 * Place each picture-bound control's image from its node. `only` limits it to the controls bound to those data nodes (a Data
 * tab edit's), and `ids` to the controls with those `w:id`s (a bind or a width set in a tab acts on its own control alone, and
 * says nothing of the others; ED-005 section 12.117).
 */
export async function refreshPictureControls(view: EditorView, session: EditorDocument, only?: ReadonlySet<Node>, ids?: ReadonlySet<number>): Promise<PictureRefresh> {
  const result: PictureRefresh = { changed: 0, notes: [] };
  if (!session.template.loaded) return result;
  const controls = pictureControlsIn(view.state.doc, session.template);
  if (controls.length === 0) return result;
  let cache = filled.get(session);
  if (!cache) { cache = new Map(); filled.set(session, cache); }
  const built: { id: number | undefined; node: PMNode; pos: number; content: Fragment; value: string }[] = [];
  for (const c of controls) {
    const id = sdtId(c.node.attrs['sdtPr']);
    if (ids && (id === undefined || !ids.has(id))) continue;
    const data = boundNodeOf(session, c.picture.binding);
    if (!data || (only && !only.has(data.node))) continue;
    const value = data.nodeValue.trim();
    if (value === '') continue;
    // A control filled from this value at this width is left; a new width (the Control tab's field) places the picture again.
    const key = `${String(c.picture.width ?? '')}\u0000${value}`;
    if (id !== undefined && cache.get(id) === key) continue;
    if (c.picture.shape === 'handler' && c.picture.width === undefined) { result.notes.push(t('picture.note.floating')); continue; }
    if (!looksLikeImageBase64(value)) { result.notes.push(t('picture.note.notImage', { xpath: c.picture.binding.xpath })); continue; }
    // Word's picture control holding a picture: the node's fitted inside that picture's box, as Word shows it (section 12.126);
    // holding none (text, or nothing), at its natural size, as the engine fills it where Word leaves it (core-ts CR-005, E2).
    const held = c.picture.shape === 'picture' ? firstImage(c.node) : undefined;
    const box = held && typeof held.attrs['width'] === 'number' && typeof held.attrs['height'] === 'number' ? fittedInBox(value, { width: held.attrs['width'], height: held.attrs['height'] }) : undefined;
    let made: { sourceId: string; width: number | null; height: number | null };
    try {
      made = session.addInlinePicture(value, box ?? (c.picture.width !== undefined && c.picture.width !== 'auto' ? { width: c.picture.width } : {}));
    } catch (e) {
      result.notes.push((e as Error).message);
      continue;
    }
    const schema = view.state.schema;
    const image = schema.nodes['image']!.create({ sourceId: made.sourceId, width: made.width, height: made.height });
    const content = c.node.type.name === 'sdtInline' ? Fragment.from(image) : Fragment.from(schema.nodes['paragraph']!.create(null, image));
    built.push({ id, node: c.node, pos: c.pos, content, value: key });
  }
  if (built.length === 0) return result;
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
    tr.replaceWith(pos + 1, pos + b.node.nodeSize - 1, b.content);
    const sdtPr = b.node.attrs['sdtPr'];
    if (readControl(sdtPr).showingPlaceholder) tr.setNodeMarkup(pos, undefined, { ...b.node.attrs, sdtPr: applyControlProperties(sdtPr, { showingPlaceholder: false }) });
    result.changed++;
  }
  if (!tr.docChanged) return result;
  tr.setMeta(UNTRACKED_META, true).setMeta('addToHistory', false).setMeta(BOUND_WRITE_META, true);
  session.replaceOutOfHistory(tr.doc, state.doc);
  return result;
}
