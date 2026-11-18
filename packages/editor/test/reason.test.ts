// SPDX-License-Identifier: GPL-3.0-or-later
//
// A failed save says why (ED-005 section 12.114): the engine's outer error names the part, and the reason is its cause.
import { describe, test, expect } from 'vitest';
import { reasonOf } from '../src/app/files.js';

describe('reasonOf', () => {
  test('each cause after what it caused', () => {
    const e = new Error('Problem saving part /word/document.xml', { cause: new Error('Part /word/document.xml holds U+000B', { cause: new Error('deeper') }) });
    expect(reasonOf(e)).toBe('Problem saving part /word/document.xml: Part /word/document.xml holds U+000B: deeper');
  });
  test('an error with no cause is its message, and a cause that is no error ends the words', () => {
    expect(reasonOf(new Error('Refused'))).toBe('Refused');
    expect(reasonOf(new Error('Refused', { cause: 'a string' }))).toBe('Refused');
  });
  test('what is no error is its text', () => {
    expect(reasonOf('thrown')).toBe('thrown');
  });
});
