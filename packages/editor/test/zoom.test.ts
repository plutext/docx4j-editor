// SPDX-License-Identifier: GPL-3.0-or-later
// @vitest-environment node
//
// The page's zoom (ED-005 section 12.110): the steps, the range and the store.
import { describe, test, expect } from 'vitest';
import { clampZoom, stepZoom, ZOOM_MIN, ZOOM_MAX } from '../src/app/zoom.js';

describe('zoom', () => {
  test('steps by 10% to the next multiple of it', () => {
    expect(stepZoom(1, 1)).toBe(1.1);
    expect(stepZoom(1, -1)).toBe(0.9);
    expect(stepZoom(1.05, 1)).toBe(1.1);
    expect(stepZoom(1.05, -1)).toBe(1);
    expect(stepZoom(0.3, -1)).toBe(0.2);
    expect(stepZoom(1.1, 1)).toBe(1.2);   // not 1.2000000000000002
  });
  test('held to Word\'s range', () => {
    expect(stepZoom(ZOOM_MAX, 1)).toBe(ZOOM_MAX);
    expect(stepZoom(ZOOM_MIN, -1)).toBe(ZOOM_MIN);
    expect(clampZoom(9)).toBe(ZOOM_MAX);
    expect(clampZoom(0)).toBe(ZOOM_MIN);
    expect(clampZoom(Number.NaN)).toBe(1);
    expect(clampZoom(1.234)).toBe(1.23);
  });
});
