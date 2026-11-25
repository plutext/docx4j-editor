// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-013 sections 3.1, 3.3, 3.4 and 3.7: when the compact layout applies, when the keyboard counts as open, and what the bar
// above it holds. The phone's
// figures are the S23 Ultra's (section 1.4): 384 × 824 CSS px, the viewport 686 high, 375 with the keyboard open.
import { describe, test, expect } from 'vitest';
import { isPhone, isCompact, keyboardOf, isLayoutChoice, loadLayout, saveLayout, barGroups, OPTION_LAYOUT, viewportGaps, isAppleTouch, caretScroll } from '../src/app/compact.js';

const phone = { touchOnly: true, screenWidth: 384, screenHeight: 824 };

describe('the rule (P1)', () => {
  test('a touch-only device whose shorter side is at most 600 px is a phone, in either orientation', () => {
    expect(isPhone(phone)).toBe(true);
    expect(isPhone({ ...phone, screenWidth: 824, screenHeight: 384 })).toBe(true);
    expect(isPhone({ touchOnly: true, screenWidth: 600, screenHeight: 960 })).toBe(true);
  });

  test('a tablet, or a device with a mouse, is not', () => {
    expect(isPhone({ touchOnly: true, screenWidth: 820, screenHeight: 1180 })).toBe(false);
    expect(isPhone({ ...phone, touchOnly: false })).toBe(false);
  });

  test('the option overrides it either way', () => {
    expect(isCompact('auto', phone)).toBe(true);
    expect(isCompact('full', phone)).toBe(false);
    expect(isCompact('compact', { touchOnly: false, screenWidth: 1920, screenHeight: 1080 })).toBe(true);
  });

  test('the option is kept per browser, Automatic as no entry', () => {
    saveLayout('full');
    expect(localStorage.getItem(OPTION_LAYOUT)).toBe('full');
    expect(loadLayout()).toBe('full');
    saveLayout('auto');
    expect(localStorage.getItem(OPTION_LAYOUT)).toBeNull();
    expect(loadLayout()).toBe('auto');
    localStorage.setItem(OPTION_LAYOUT, 'sideways');
    expect(loadLayout()).toBe('auto');
    localStorage.removeItem(OPTION_LAYOUT);
    expect(isLayoutChoice('compact')).toBe(true);
    expect(isLayoutChoice('')).toBe(false);
  });
});

describe('the keyboard (section 3.3)', () => {
  test("open under 75% of the viewport's closed height", () => {
    expect(keyboardOf(375, 686)).toBe(true);
    expect(keyboardOf(686, 686)).toBe(false);
    // Chrome's address bar hiding as the page scrolls changes the height by less: not a keyboard.
    expect(keyboardOf(686 - 91, 686)).toBe(false);
    // Landscape: 298 closed, 90 above Samsung Keyboard's split layout (section 1.5).
    expect(keyboardOf(90, 298)).toBe(true);
    expect(keyboardOf(300, 0)).toBe(false);
  });
});

describe('the bar above the keyboard (section 3.4, P3 and P11)', () => {
  const none = { review: false, inChange: false, hasPrevious: false, hasNext: false, selection: false, comments: 0, inTable: false };

  test('Done alone outside any context; a group per context, Review first, then Comment, then Comments (n), then the rows', () => {
    expect(barGroups(none, false)).toEqual([]);
    expect(barGroups({ ...none, inTable: true }, false)).toEqual(['table']);
    expect(barGroups({ ...none, selection: true, inTable: true }, false)).toEqual(['comment', 'table']);
    expect(barGroups({ ...none, comments: 2, inTable: true }, false)).toEqual(['comments', 'table']);
    expect(barGroups({ ...none, review: true, inChange: true, selection: true, comments: 1, inTable: true }, false)).toEqual(['review', 'comment', 'comments', 'table']);
  });

  test('in landscape with the keyboard open, the strip holds one group, the first', () => {
    expect(barGroups({ ...none, selection: true, inTable: true }, true)).toEqual(['comment']);
    expect(barGroups({ ...none, comments: 1, inTable: true }, true)).toEqual(['comments']);
    expect(barGroups({ ...none, review: true, inTable: true }, true)).toEqual(['review']);
    expect(barGroups(none, true)).toEqual([]);
  });
});

// iOS Safari (ED-013 section 9, the iPhone): its keyboard covers the layout viewport rather than shrinking it. The figures are
// an iPhone 13 Pro's in portrait (390 × 699 CSS px with the keyboard closed), read from the page on the phone.
describe('the iPhone', () => {
  test('the gaps are the layout viewport outside the visual one, above and below', () => {
    expect(viewportGaps(677, 22, 362)).toEqual({ top: 22, bottom: 293 });
    expect(viewportGaps(699, 0, 699)).toEqual({ top: 0, bottom: 0 });
  });

  test('a visual viewport reported past the layout one leaves no gap below, never a negative one', () => {
    expect(viewportGaps(295, 404, 295)).toEqual({ top: 404, bottom: 0 });
  });

  test('the caret is scrolled into the band between the bars, by as little as brings it in', () => {
    const band = { top: 44, bottom: 318 };
    expect(caretScroll({ top: 100, bottom: 115 }, band)).toBe(0);
    expect(caretScroll({ top: 320, bottom: 335 }, band)).toBe(335 - 318 + 4);
    expect(caretScroll({ top: 30, bottom: 45 }, band)).toBe(30 - 44 - 4);
  });

  test('in a band shorter than the line, the line\'s foot is kept, where the next character goes', () => {
    expect(caretScroll({ top: 0, bottom: 20 }, { top: 5, bottom: 16 })).toBe(20 - 16 + 4);
  });

  test('WebKit on an iPhone or an iPad is told from a Mac and from Android', () => {
    expect(isAppleTouch('Mozilla/5.0 (iPhone; CPU iPhone OS 26_6 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/26.6 Mobile/15E148 Safari/604.1', 5)).toBe(true);
    expect(isAppleTouch('Mozilla/5.0 (iPad; CPU OS 18_0 like Mac OS X) AppleWebKit/605.1.15', 5)).toBe(true);
    expect(isAppleTouch('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/26.0 Safari/605.1.15', 5)).toBe(true);
    expect(isAppleTouch('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/26.0 Safari/605.1.15', 0)).toBe(false);
    expect(isAppleTouch('Mozilla/5.0 (Linux; Android 16; SM-S918B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/157.0.0.0 Mobile Safari/537.36', 5)).toBe(false);
  });
});
