// SPDX-License-Identifier: GPL-3.0-or-later
//
// The room a caret scrolled into view keeps from the window's edges (ED-013 section 3.10). In the compact layout the
// document scrolls with the page's root, under a sticky top bar and over a sticky status bar, and ProseMirror's
// scrolling knows neither: a caret it scrolled to the bottom edge sat under the status bar. The application sets the
// bars' heights here as they change; every surface reads them through getters each time it scrolls, so no live view is
// reconfigured. Elsewhere it is ProseMirror's default, 5 px.
const margin = { top: 5, bottom: 5 };

/** The space at the top and the bottom of the window that the application's sticky bars take, in CSS px. */
export function setScrollMargin(top: number, bottom: number): void {
  margin.top = top;
  margin.bottom = bottom;
}

/** Every editing surface's `scrollMargin` prop. */
export const scrollMargin = {
  get top(): number { return margin.top; },
  get bottom(): number { return margin.bottom; },
  left: 5,
  right: 5,
};
