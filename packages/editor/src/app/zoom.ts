// SPDX-License-Identifier: GPL-3.0-or-later
//
// The page's zoom (ED-005 section 12.110; Jason, 2026-11-01: the browser's Ctrl+plus enlarges the chrome with the
// document, wanted sometimes and not always): a factor over the page column and its comments margin alone, applied
// as CSS `zoom`, stepped by the status bar's control, the palette and Ctrl+wheel over the page, and kept per
// browser as Word keeps its last zoom. Word's range, 10% to 500%, in steps of 10%.

export const ZOOM_KEY = 'docx4j-editor.zoom';
export const ZOOM_MIN = 0.1;
export const ZOOM_MAX = 5;
export const ZOOM_STEP = 0.1;

/** The zoom held to the range, rounded to a whole percent. */
export function clampZoom(zoom: number): number {
  if (!Number.isFinite(zoom)) return 1;
  return Math.round(Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, zoom)) * 100) / 100;
}

/** The next step up or down: to the next multiple of the step, so 1.05 steps to 1.1 and 1.0 to 0.9. */
export function stepZoom(zoom: number, dir: 1 | -1): number {
  const steps = zoom / ZOOM_STEP;
  const next = dir > 0 ? Math.floor(steps + 1e-6) + 1 : Math.ceil(steps - 1e-6) - 1;
  return clampZoom(next * ZOOM_STEP);
}

export function loadZoom(): number {
  try {
    const v = localStorage.getItem(ZOOM_KEY);
    if (v !== null) return clampZoom(Number(v));
  } catch { /* no storage */ }
  return 1;
}

export function saveZoom(zoom: number): void {
  try {
    if (zoom === 1) localStorage.removeItem(ZOOM_KEY); else localStorage.setItem(ZOOM_KEY, String(zoom));
  } catch { /* no storage */ }
}
