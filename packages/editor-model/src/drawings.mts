// SPDX-License-Identifier: GPL-3.0-or-later
//
// The drawing behind an `image` node (ED-005 section 12.106; ED-009). Its source is a `w:drawing`
// with one `wp:inline`: since ED-009 a drawing inside a kept `mc:AlternateContent`'s selected branch is
// projected as the run's own item, the node's `alt` naming the element, and the export writes the
// element back around it. A store kept before ED-009 (ED-005 section 12.106 item 3) has the element
// itself as the source, and the readers here still take that form: the selected branch
// (`selectedBranchOf`, the branch the engine's preprocessor once resolved the element to) holding such
// a drawing alone. Whatever reads the drawing (the picture's part, its size) comes through here rather
// than reading `el.value` as a drawing.
import { selectedBranchOf, type Element } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type * as mce from '@docx4j/generated-objects-ts/modules/org_docx4j_mce';

/** A `wp:inline`, as much of it as the editor reads. */
export interface InlineDrawing {
  extent?: { cx?: number; cy?: number };
  graphic?: { graphicData?: { any?: Element[] } };
}

/** The `w:drawing` with one `wp:inline` an element is, or the one a kept `mc:AlternateContent`'s selected branch holds alone. */
export function inlineDrawingOf(el: Element | undefined): Element<wml.Drawing> | undefined {
  if (!el) return undefined;
  const local = el.name.localPart;
  if (local === 'AlternateContent') {
    const branch = selectedBranchOf(el.value as mce.AlternateContent);
    return branch?.length === 1 ? inlineDrawingOf(branch[0]) : undefined;
  }
  if (local !== 'drawing') return undefined;
  const drawing = el.value as wml.Drawing;
  const first = drawing.anchorOrInline?.[0] as { TYPE_NAME?: string } | undefined;
  return drawing.anchorOrInline?.length === 1 && first?.TYPE_NAME?.endsWith('.Inline') ? (el as Element<wml.Drawing>) : undefined;
}

/** The `wp:inline` of an image node's source, if the source is one (see `inlineDrawingOf`). */
export function inlineOf(el: Element | undefined): InlineDrawing | undefined {
  return inlineDrawingOf(el)?.value.anchorOrInline?.[0] as InlineDrawing | undefined;
}

/** The relationship id of the picture an image node's source embeds (`a:blip r:embed`); undefined for a drawing that is not a picture. */
export function blipEmbedOf(el: Element | undefined): string | undefined {
  const pic = inlineOf(el)?.graphic?.graphicData?.any?.[0]?.value as { blipFill?: { blip?: { embed?: string } } } | undefined;
  return pic?.blipFill?.blip?.embed;
}
