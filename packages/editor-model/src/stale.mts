// SPDX-License-Identifier: GPL-3.0-or-later
//
// Bound controls whose content is older than their data (ED-005 section 12.89). Word shows a bound control's node
// value at open; the editor shows the file's content until a bind or a data edit, so that an unedited file saves
// byte for byte. At open the application counts the controls a refresh would change and offers it, once, as the
// user's action: a text control whose shown value is not the node's (the Data tab's rule, `refreshBoundControls`),
// an XHTML control whose text is not the markup's, a picture control that holds no image or another image's bytes, and
// Word's rich text mapped to a Word document whose text is not the document's (ED-005 proposal 46).
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorState } from 'prosemirror-state';
import type { Element } from '@docx4j/core-ts';
import { refreshBoundControls, boundNodeOf, isFlatOpc } from './bound.mjs';
import { xhtmlControlsIn } from './xhtml.mjs';
import { pictureControlsIn, looksLikeImageBase64, base64Bytes } from './pictures.mjs';
import { fragmentControlsIn, fragmentFactsOf } from './fragments.mjs';
import { blipEmbedOf } from './drawings.mjs';
import type { EditorDocument } from './session.mjs';

export interface StaleBoundControls {
  text: number;
  xhtml: number;
  picture: number;
  /** Word's rich text mapped to a Word document (ED-005 proposal 46), whose text is not the document's: Word fills it on open. */
  fragment: number;
  total: number;
}

/** The text of escaped markup, as a reader would see it: tags gone, the common entities decoded, whitespace collapsed. */
export function textOfMarkup(markup: string): string {
  return markup
    .replace(/<(br|p|div|li|tr|h[1-6])\b[^>]*>/gi, ' ')
    .replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'").replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ').trim();
}

const collapse = (text: string): string => text.replace(/\s+/g, ' ').trim();

/** The text controls a refresh would change: `refreshBoundControls`'s own count. */
export function staleTextControls(session: EditorDocument, state: EditorState): number {
  const tr = refreshBoundControls(session, state);
  if (!tr) return 0;
  // One `setNodeMarkup` per control written (bound.mts), after its content's replacement.
  return tr.steps.filter((step) => step.constructor.name === 'ReplaceAroundStep').length;
}

/** The XHTML controls whose text is not their node's markup's. */
export function staleXhtmlControls(session: EditorDocument, doc: PMNode): number {
  if (!session.template.loaded) return 0;
  let n = 0;
  for (const c of xhtmlControlsIn(doc, session.template)) {
    const node = boundNodeOf(session, c.binding);
    const value = node?.nodeValue.trim() ?? '';
    if (value === '') continue;
    if (textOfMarkup(value) !== collapse(c.node.textContent)) n++;
  }
  return n;
}

/** The image part an `image` node's drawing points at, through the main part's relationships. */
async function imageBytesOf(session: EditorDocument, sourceId: string): Promise<Uint8Array | undefined> {
  const embed = blipEmbedOf(session.ctx.ids.get(sourceId) as Element | undefined);
  if (!embed) return undefined;
  const part = session.main.relationshipsPart?.getPart(embed) as { getBytes?: () => Promise<Uint8Array> } | undefined;
  return part?.getBytes ? part.getBytes() : undefined;
}

/** The picture controls a refresh would fill: holding no image, or an image whose bytes are not the node's. */
export async function stalePictureControls(session: EditorDocument, doc: PMNode): Promise<number> {
  if (!session.template.loaded) return 0;
  let n = 0;
  for (const c of pictureControlsIn(doc, session.template)) {
    if (c.picture.shape === 'handler' && c.picture.width === undefined) continue;   // left for the processor (section 12.76)
    const node = boundNodeOf(session, c.picture.binding);
    const value = node?.nodeValue.trim() ?? '';
    if (value === '' || !looksLikeImageBase64(value)) continue;
    let image: PMNode | undefined;
    c.node.descendants((d) => { if (image) return false; if (d.type.name === 'image') { image = d; return false; } return true; });
    if (!image) { n++; continue; }
    const have = await imageBytesOf(session, image.attrs['sourceId'] as string);
    if (!have) { n++; continue; }
    const want = base64Bytes(value);
    if (have.length !== want.length || have.some((b, i) => b !== want[i])) n++;
  }
  return n;
}

/** Every bound control a refresh would change, by kind, over the document `state` holds. */
export async function staleBoundControls(session: EditorDocument, state: EditorState): Promise<StaleBoundControls> {
  const text = staleTextControls(session, state);
  const xhtml = staleXhtmlControls(session, state.doc);
  const picture = await stalePictureControls(session, state.doc);
  const fragment = await staleFragmentControls(session, state.doc);
  return { text, xhtml, picture, fragment, total: text + xhtml + picture + fragment };
}

/**
 * Word's rich text controls mapped by `w15:dataBinding` to a Word document as Flat OPC whose text is not the document's
 * (ED-005 proposal 46, decision (d)): Word 365 fills such a control from its node on open (the proposal's part 6(a)), so the
 * editor offers it. OpenDoPE's `od:progid` form is a processor's to fill, never Word's, and is not counted. Compared loosely,
 * whitespace aside, as the XHTML count is: converted formatting never reads as stale.
 */
export async function staleFragmentControls(session: EditorDocument, doc: PMNode): Promise<number> {
  if (!session.template.loaded) return 0;
  let n = 0;
  for (const c of fragmentControlsIn(doc, session.template)) {
    if (c.fragment.shape !== 'w15') continue;
    const value = boundNodeOf(session, c.fragment.binding)?.nodeValue ?? '';
    if (!isFlatOpc(value)) continue;
    const facts = await fragmentFactsOf(value);
    if (facts && facts.text.replace(/\s+/g, '') !== c.node.textContent.replace(/\s+/g, '')) n++;
  }
  return n;
}

