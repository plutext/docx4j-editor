// SPDX-License-Identifier: GPL-3.0-or-later
//
// The application bar's menus (ED-005 section 12.112). A pane hangs from its button, and the bar wraps, so the button can
// land anywhere in the window and a pane anchored to it can reach past an edge, where the root's `overflow: clip`
// (section 12.111) leaves no scrolling to it. Measured here, since no stylesheet rule knows where the button landed.

/** The gap kept between a pane and the window's edge, half of the 1rem the stylesheet takes off a pane's width at most. */
const MARGIN = 7;

/** An open menu's pane moved back inside the window, and its height held above the bottom edge, its own scrolling reaching the rest. */
export function placeMenuPane(menu: HTMLDetailsElement): void {
  const pane = menu.querySelector<HTMLElement>(':scope > .menu-pane');
  if (!pane || !menu.open) return;
  pane.style.translate = '';
  pane.style.maxBlockSize = '';
  const { left, right, top } = pane.getBoundingClientRect();
  const width = document.documentElement.clientWidth;
  const shift = left < MARGIN ? MARGIN - left : right > width - MARGIN ? width - MARGIN - right : 0;
  if (shift !== 0) pane.style.translate = `${shift}px`;
  pane.style.maxBlockSize = `${Math.max(0, window.innerHeight - top - MARGIN)}px`;
}

/** Every open menu placed again, after anything that can move its button: the window's size, the perspective, the language. */
export function placeOpenMenus(): void {
  document.querySelectorAll<HTMLDetailsElement>('.app-bar .panels-menu[open]').forEach(placeMenuPane);
}
