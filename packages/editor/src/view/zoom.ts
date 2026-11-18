// SPDX-License-Identifier: GPL-3.0-or-later
//
// The page is zoomed by CSS `zoom` on its container (ED-005 section 12.110): the document larger or smaller while
// the chrome keeps its size. Client rectangles and the caret's coordinates come back zoomed, while an element's
// layout size (`offsetWidth`) and the pixel values a style states do not, so whatever compares the two divides
// the first by the scale here.

/** The factor between an element's client rectangle and its layout size: the zoom in force over it, 1 outside one. */
export function scaleOf(el: Element | null | undefined): number {
  if (!(el instanceof HTMLElement)) return 1;
  const layout = el.offsetWidth;
  if (!layout) return 1;
  const scale = el.getBoundingClientRect().width / layout;
  return Number.isFinite(scale) && scale > 0 ? scale : 1;
}
