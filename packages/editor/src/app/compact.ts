// SPDX-License-Identifier: GPL-3.0-or-later
//
// The compact layout for phones (ED-013): when it applies (section 3.1), the option that overrides it, and the
// on-screen keyboard's state (section 3.3). Pure rules first, so that the tests can hold them without a phone; the
// hooks below follow the browser.
//
// - The rule: a touch-only device (`(pointer: coarse) and (hover: none)`) whose shorter screen side is at most
//   600 CSS px. The screen's sides, not the window's: a phone stays compact when it turns to landscape.
// - The option: Automatic (the rule), Compact or Full, kept per browser.
// - The keyboard: open when the visual viewport is under 75% of its height with the keyboard closed, for the
//   viewport's current width (a keyboard changes the height, never the width; a rotation changes both, and the
//   screen's sides turn before the viewport's, so the width is what tells one orientation's heights from the other's).
//   With `interactive-widget=resizes-content` (P7) the layout viewport shrinks with it; without it, only the visual
//   viewport does, and both are read the same.
import { useEffect, useState, useSyncExternalStore } from 'react';

export type LayoutChoice = 'auto' | 'compact' | 'full';

/** The option's key in the browser's storage. */
export const OPTION_LAYOUT = 'docx4j-editor.layout';

/** The shorter screen side at or under which a touch-only device is a phone (ED-013 P1). */
export const PHONE_SIDE = 600;

/** The share of its closed height under which the viewport means an open keyboard. */
export const KEYBOARD_SHARE = 0.75;

export function isLayoutChoice(value: unknown): value is LayoutChoice {
  return value === 'auto' || value === 'compact' || value === 'full';
}

export function loadLayout(): LayoutChoice {
  try { const v = localStorage.getItem(OPTION_LAYOUT); return isLayoutChoice(v) ? v : 'auto'; } catch { return 'auto'; }
}

export function saveLayout(choice: LayoutChoice): void {
  try { if (choice === 'auto') localStorage.removeItem(OPTION_LAYOUT); else localStorage.setItem(OPTION_LAYOUT, choice); } catch { /* ignore */ }
}

/** What the rule reads from the device. */
export interface DeviceTraits {
  /** `(pointer: coarse) and (hover: none)`: no mouse or trackpad. */
  touchOnly: boolean;
  /** The screen's sides in CSS px. */
  screenWidth: number;
  screenHeight: number;
}

/** Whether the device is a phone by the rule (ED-013 P1). */
export function isPhone(traits: DeviceTraits): boolean {
  return traits.touchOnly && Math.min(traits.screenWidth, traits.screenHeight) <= PHONE_SIDE;
}

/** Whether the layout is compact, for the option and the device. */
export function isCompact(choice: LayoutChoice, traits: DeviceTraits): boolean {
  return choice === 'compact' || (choice === 'auto' && isPhone(traits));
}

const TOUCH_ONLY = '(pointer: coarse) and (hover: none)';

function traitsNow(): DeviceTraits {
  const touchOnly = typeof matchMedia === 'function' && matchMedia(TOUCH_ONLY).matches;
  const w = typeof screen !== 'undefined' ? screen.width : 0;
  const h = typeof screen !== 'undefined' ? screen.height : 0;
  return { touchOnly, screenWidth: w || (typeof innerWidth === 'number' ? innerWidth : 0), screenHeight: h || (typeof innerHeight === 'number' ? innerHeight : 0) };
}

/** Whether the layout is compact now, followed as the device changes (a mouse plugged in, the screen turned). */
export function useCompact(choice: LayoutChoice): boolean {
  const [traits, setTraits] = useState<DeviceTraits>(traitsNow);
  useEffect(() => {
    const update = (): void => setTraits((old) => { const next = traitsNow(); return sameTraits(old, next) ? old : next; });
    const query = typeof matchMedia === 'function' ? matchMedia(TOUCH_ONLY) : null;
    query?.addEventListener('change', update);
    addEventListener('resize', update);
    addEventListener('orientationchange', update);
    return () => { query?.removeEventListener('change', update); removeEventListener('resize', update); removeEventListener('orientationchange', update); };
  }, []);
  return isCompact(choice, traits);
}

function sameTraits(a: DeviceTraits, b: DeviceTraits): boolean {
  return a.touchOnly === b.touchOnly && a.screenWidth === b.screenWidth && a.screenHeight === b.screenHeight;
}

/** The keyboard's state, and the orientation it is judged in. */
export interface KeyboardState {
  open: boolean;
  landscape: boolean;
  /** The visible height in CSS px, for the bars fixed above the keyboard where `resizes-content` is not honoured. */
  height: number;
}

/**
 * The keyboard's state from the viewport's height now and the tallest seen at the same width, which is the height
 * with the keyboard closed. Pure, for the tests.
 */
export function keyboardOf(height: number, tallest: number): boolean {
  return tallest > 0 && height < tallest * KEYBOARD_SHARE;
}

/** The keyboard as the viewport shows it; followed while `enabled` (the compact layout), closed otherwise. */
export function useKeyboard(enabled: boolean): KeyboardState {
  const snapshot = useSyncExternalStore(enabled ? subscribeViewport : noopSubscribe, viewportSnapshot, () => CLOSED_KEY);
  if (!enabled) return CLOSED;
  const [open, landscape, height] = snapshot.split('|');
  return { open: open === '1', landscape: landscape === '1', height: Number(height) };
}

const CLOSED: KeyboardState = { open: false, landscape: false, height: 0 };
const CLOSED_KEY = '0|0|0';
/** The tallest viewport seen at each width: the keyboard's closed height. */
const tallest = new Map<number, number>();

function viewportSnapshot(): string {
  if (typeof window === 'undefined') return CLOSED_KEY;
  const height = Math.round(window.visualViewport?.height ?? innerHeight);
  const width = Math.round(innerWidth);
  const landscape = isLandscape();
  const closed = Math.max(tallest.get(width) ?? 0, height);
  tallest.set(width, closed);
  return `${keyboardOf(height, closed) ? 1 : 0}|${landscape ? 1 : 0}|${height}`;
}

/**
 * Whether the phone is turned to landscape: the screen's sides, or its orientation where the sides do not turn (iOS
 * reports an iPhone's screen as 390 × 844 in either orientation).
 */
function isLandscape(): boolean {
  if (typeof screen === 'undefined') return false;
  return screen.width > screen.height || (screen.orientation?.type?.startsWith('landscape') ?? false);
}

function subscribeViewport(onChange: () => void): () => void {
  const vv = window.visualViewport;
  vv?.addEventListener('resize', onChange);
  addEventListener('resize', onChange);
  addEventListener('orientationchange', onChange);
  return () => { vv?.removeEventListener('resize', onChange); removeEventListener('resize', onChange); removeEventListener('orientationchange', onChange); };
}

function noopSubscribe(): () => void {
  return () => {};
}

/**
 * Whether this is WebKit on an iPhone or iPad, where every browser is Safari underneath. An iPad asking for the desktop
 * site says Macintosh, and is told from a Mac by its touch points.
 */
export function isAppleTouch(userAgent: string, maxTouchPoints: number): boolean {
  return /\b(iPhone|iPad|iPod)\b/.test(userAgent) || (/\bMacintosh\b/.test(userAgent) && maxTouchPoints > 1);
}

/**
 * iOS Safari zooms the page in when an editable element with text under 16 px takes the focus: a document's 11 pt is
 * 14.67 px, so a tap into the text zoomed to 1.09, cut off the screen's end (the bar's Done half off it) and moved where
 * taps landed. `maximum-scale=1` stops that zoom, and iOS still lets a person pinch-zoom (it ignores the limit for that).
 * Chrome for Android would honour the limit and stop pinch-zoom, so it is set on iOS alone.
 */
export function stopFocusZoom(): void {
  if (typeof navigator === 'undefined' || !isAppleTouch(navigator.userAgent, navigator.maxTouchPoints ?? 0)) return;
  // Marked for the rules that answer Safari's ways (compact.css).
  document.documentElement.classList.add('apple-touch');
  const meta = document.querySelector<HTMLMetaElement>('meta[name="viewport"]');
  if (meta && !/maximum-scale/.test(meta.content)) meta.content = `${meta.content}, maximum-scale=1`;
}

/**
 * The layout viewport outside the visual one, above it and below it, in whole CSS px. Zero where the keyboard resizes the
 * page (`resizes-content`, Chrome for Android); on iOS Safari the keyboard covers the layout viewport's end, and Safari
 * scrolls the visual viewport inside it to keep the caret in sight, so the sticky bars, placed against the layout
 * viewport, end up behind the keyboard and above the screen. Pure, for the tests.
 */
export function viewportGaps(layoutHeight: number, offsetTop: number, visualHeight: number): { top: number; bottom: number } {
  const top = Math.max(0, Math.round(offsetTop));
  const bottom = Math.max(0, Math.round(layoutHeight - offsetTop - visualHeight));
  return { top, bottom };
}

/**
 * Follows the visual viewport while `enabled`, setting `--vv-top` and `--vv-bottom` on the root element: the compact
 * bars are moved by them into the visible part.
 */
export function useViewportGaps(enabled: boolean): void {
  useEffect(() => {
    const vv = window.visualViewport;
    const root = document.documentElement.style;
    const clear = (): void => { root.removeProperty('--vv-top'); root.removeProperty('--vv-bottom'); };
    if (!enabled || !vv) { clear(); return; }
    let last = '';
    // Safari reports the visual viewport's place only after a scroll has moved it, so a bar moved by the gaps lags and
    // then jumps. While the keyboard covers the page, a scroll marks the root `vv-moving` until it settles, and the bars
    // fade out meanwhile (compact.css).
    let settle: ReturnType<typeof setTimeout> | undefined;
    const moving = (): void => {
      if (!last || last === '0|0') return;
      document.documentElement.classList.add('vv-moving');
      clearTimeout(settle);
      settle = setTimeout(() => document.documentElement.classList.remove('vv-moving'), 250);
    };
    const update = (): void => {
      moving();
      // The layout viewport's height as the sticky bars are placed against it: the root's client height. Safari's
      // `innerHeight` changes as it scrolls with the keyboard up (699, 683, 531 in one scroll) where that does not.
      const { top, bottom } = viewportGaps(document.documentElement.clientHeight || innerHeight, vv.offsetTop, vv.height);
      const key = `${top}|${bottom}`;
      if (key === last) return;
      last = key;
      root.setProperty('--vv-top', `${top}px`);
      root.setProperty('--vv-bottom', `${bottom}px`);
    };
    update();
    vv.addEventListener('resize', update);
    vv.addEventListener('scroll', update);
    addEventListener('scroll', update, { passive: true });
    return () => {
      vv.removeEventListener('resize', update);
      vv.removeEventListener('scroll', update);
      removeEventListener('scroll', update);
      clearTimeout(settle);
      document.documentElement.classList.remove('vv-moving');
      clear();
    };
  }, [enabled]);
}

/** Whether the focus is in an editing surface (a ProseMirror view one can type in); followed while `enabled`. */
export function useSurfaceFocused(enabled: boolean): boolean {
  const focused = useSyncExternalStore(enabled ? subscribeFocus : noopSubscribe, surfaceFocused, () => false);
  return enabled && focused;
}

function surfaceFocused(): boolean {
  const el = typeof document !== 'undefined' ? document.activeElement : null;
  return el instanceof HTMLElement && el.closest('.ProseMirror[contenteditable="true"]') !== null;
}

function subscribeFocus(onChange: () => void): () => void {
  document.addEventListener('focusin', onChange);
  document.addEventListener('focusout', onChange);
  return () => { document.removeEventListener('focusin', onChange); document.removeEventListener('focusout', onChange); };
}

/**
 * Whether an editing surface has the focus from a tap into it: on iOS only a person's tap raises the keyboard, and focus
 * the application sets (at a document's open) does not. A tap on the text sets it; the focus leaving every surface (Done,
 * a sheet) clears it. Followed while `enabled`.
 */
export function useTappedIntoText(enabled: boolean): boolean {
  const tapped = useSyncExternalStore(enabled ? subscribeTap : noopSubscribe, () => tappedIntoText && surfaceFocused(), () => false);
  return enabled && tapped;
}

let tappedIntoText = false;

function subscribeTap(onChange: () => void): () => void {
  const down = (e: PointerEvent): void => {
    if (e.target instanceof Element && e.target.closest('.ProseMirror[contenteditable="true"]')) { tappedIntoText = true; onChange(); }
  };
  const out = (): void => {
    // The focus moves on within the task; read where it went after it.
    queueMicrotask(() => { if (!surfaceFocused()) { tappedIntoText = false; onChange(); } });
  };
  document.addEventListener('pointerdown', down, true);
  document.addEventListener('focusin', onChange);
  document.addEventListener('focusout', out);
  return () => {
    document.removeEventListener('pointerdown', down, true);
    document.removeEventListener('focusin', onChange);
    document.removeEventListener('focusout', out);
  };
}

/** A group of the bar above the keyboard (ED-013 section 3.4), besides Done, which is always there. */
export type BarGroup = 'review' | 'comment' | 'comments' | 'table';

/** Where the caret is, as the bar above the keyboard reads it. */
export interface BarContext {
  /** The Review perspective, the body's surface, and changes listed: Previous and Next. */
  review: boolean;
  /** The selection in a change Accept and Reject resolve. */
  inChange: boolean;
  /** A change before the selection, and one after it: Previous and Next, greyed without one (they do not go round). */
  hasPrevious: boolean;
  hasNext: boolean;
  /** A selection a comment can be made on. */
  selection: boolean;
  /** The comment threads at the selection: Comments (n), which opens their sheet (section 3.7). */
  comments: number;
  inTable: boolean;
}

/**
 * The bar's groups, in order: Review's (Accept and Reject in a change, Next whenever there is a change to go to), then
 * Comment on a selection, then Comments (n) on commented text, then the table's rows. In landscape with the keyboard
 * open the bar is a strip with room for one group, the first (P11).
 */
export function barGroups(ctx: BarContext, landscape: boolean): BarGroup[] {
  const groups: BarGroup[] = [];
  if (ctx.review) groups.push('review');
  if (ctx.selection) groups.push('comment');
  if (ctx.comments > 0) groups.push('comments');
  if (ctx.inTable) groups.push('table');
  return landscape ? groups.slice(0, 1) : groups;
}

/** Room kept between the caret and a bar, in CSS px. */
const CARET_ROOM = 4;

/**
 * How far to scroll the page so that the caret's line is in the band left between the bars, in CSS px (negative is up,
 * 0 is in view). The band is read from the bars as drawn, in the coordinates the caret's are in, so that it holds
 * whichever viewport Safari reports them against. Where the band is shorter than the line, the line's foot is kept,
 * where the next character goes. Pure, for the tests.
 */
export function caretScroll(caret: { top: number; bottom: number }, band: { top: number; bottom: number }): number {
  if (caret.bottom > band.bottom - CARET_ROOM) return caret.bottom - band.bottom + CARET_ROOM;
  if (caret.top < band.top + CARET_ROOM) return Math.max(caret.top - band.top - CARET_ROOM, caret.bottom - band.bottom + CARET_ROOM);
  return 0;
}

/** The caret's box in client coordinates, and what identifies the selection and document it is for. */
export interface CaretBox { top: number; bottom: number; stamp: readonly unknown[] }

/** A bar's box if it is drawn and not faded out, else null. */
function drawnBox(selector: string): DOMRect | null {
  const el = document.querySelector(selector);
  if (!el) return null;
  const style = getComputedStyle(el);
  if (style.display === 'none' || style.opacity === '0') return null;
  const box = el.getBoundingClientRect();
  return box.height > 0 ? box : null;
}

/**
 * While `enabled` (the compact layout with the keyboard and the bar above it up), keeps the caret's line between the top
 * bar and the bar above the keyboard after each selection change, as one types or taps. ProseMirror's own scrolling
 * measures against the layout viewport, which iOS Safari's keyboard covers, and in landscape the band is a few lines.
 * `caretBox` gives the caret's box in client coordinates, or null where no surface has the focus. A scroll the reader
 * makes is left alone: only a selection that moved, or text that changed, acts.
 */
export function useCaretInView(enabled: boolean, caretBox: () => CaretBox | null): void {
  useEffect(() => {
    if (!enabled) return;
    // Where the visible part is, in the coordinates boxes are reported in: iOS Safari reports them against the layout
    // viewport at one time and the visual one at another (both seen on an iPhone), so a box fixed at the layout viewport's
    // top says which, and the visual viewport's offset is added to it.
    const anchor = document.createElement('div');
    anchor.style.cssText = 'position: fixed; inset-block-start: 0; inset-inline-start: 0; inline-size: 0; block-size: 0; pointer-events: none;';
    anchor.setAttribute('aria-hidden', 'true');
    document.body.append(anchor);
    let frame = 0;
    // The selection and document last acted on: iOS reports a selection change as a scroll starts under a finger,
    // and a reader scrolling away from the caret was brought back to it. Only a caret that moved, or text that
    // changed, is brought into view.
    let seen: readonly unknown[] = [];
    const check = (): void => {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(() => {
        const caret = caretBox();
        if (!caret) return;
        const same = caret.stamp.length === seen.length && caret.stamp.every((v, i) => v === seen[i]);
        seen = caret.stamp;
        if (same) return;
        const vv = window.visualViewport;
        const top = anchor.getBoundingClientRect().top + (vv?.offsetTop ?? 0);
        const visible = { top, bottom: top + (vv?.height ?? innerHeight) };
        const above = drawnBox('.app .compact-bar');
        const below = drawnBox('.app > .keyboard-bar');
        const band = { top: Math.max(visible.top, above?.bottom ?? -Infinity), bottom: Math.min(visible.bottom, below?.top ?? Infinity) };
        const by = caretScroll(caret, band);
        if (by !== 0) window.scrollBy(0, by);
      });
    };
    document.addEventListener('selectionchange', check);
    check();
    return () => {
      cancelAnimationFrame(frame);
      anchor.remove();
      document.removeEventListener('selectionchange', check);
    };
  }, [enabled, caretBox]);
}
