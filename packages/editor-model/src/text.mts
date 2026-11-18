// SPDX-License-Identifier: GPL-3.0-or-later
//
// The text model's one rule that is easy to get wrong (ED-001 section 3.7): offsets are UTF-16
// code units, and a run is never split inside a surrogate pair or a grapheme cluster.

let segmenter: Intl.Segmenter | undefined;

function graphemes(): Intl.Segmenter {
  return (segmenter ??= new Intl.Segmenter(undefined, { granularity: 'grapheme' }));
}

/**
 * The nearest grapheme boundary at or before `offset` (or at or after it, with
 * `direction: 'forward'`), in UTF-16 code units. An offset already on a boundary is returned
 * as is; 0 and `text.length` are always boundaries.
 */
export function graphemeBoundary(text: string, offset: number, direction: 'backward' | 'forward' = 'backward'): number {
  if (offset <= 0) return 0;
  if (offset >= text.length) return text.length;
  let last = 0;
  for (const { index } of graphemes().segment(text)) {
    if (index === offset) return offset;
    if (index > offset) return direction === 'backward' ? last : index;
    last = index;
  }
  return direction === 'backward' ? last : text.length;
}

/** True when `offset` falls on a grapheme boundary of `text`. */
export function isGraphemeBoundary(text: string, offset: number): boolean {
  return graphemeBoundary(text, offset) === offset;
}
