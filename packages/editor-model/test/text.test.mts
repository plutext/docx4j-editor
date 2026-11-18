// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 3.7 test 6: never split inside a surrogate pair or a grapheme cluster.
import { describe, test, expect } from 'vitest';
import { graphemeBoundary, isGraphemeBoundary } from '../src/index.mjs';

describe('graphemeBoundary', () => {
  const family = '👨‍👩‍👧‍👦';   // 4 code points joined by ZWJ: 11 UTF-16 code units
  const text = `ab${family}cd`;

  test('an offset inside an emoji sequence moves to its edge', () => {
    expect(family.length).toBe(11);
    for (let i = 3; i < 2 + family.length; i++) {
      expect(graphemeBoundary(text, i)).toBe(2);
      expect(graphemeBoundary(text, i, 'forward')).toBe(2 + family.length);
    }
  });

  test('offsets on boundaries stay put; 0 and length are boundaries', () => {
    expect(graphemeBoundary(text, 0)).toBe(0);
    expect(graphemeBoundary(text, 2)).toBe(2);
    expect(graphemeBoundary(text, 2 + family.length)).toBe(2 + family.length);
    expect(graphemeBoundary(text, text.length)).toBe(text.length);
    expect(isGraphemeBoundary(text, 1)).toBe(true);
    expect(isGraphemeBoundary(text, 4)).toBe(false);
  });

  test('a combining mark stays with its base', () => {
    const s = 'éx';   // e + combining acute
    expect(graphemeBoundary(s, 1)).toBe(0);
    expect(graphemeBoundary(s, 1, 'forward')).toBe(2);
    expect(graphemeBoundary(s, 2)).toBe(2);
  });

  test('a surrogate pair is never split', () => {
    const s = 'a𝄞b';   // U+1D11E, two code units
    expect(graphemeBoundary(s, 2)).toBe(1);
    expect(graphemeBoundary(s, 2, 'forward')).toBe(3);
  });
});
