// SPDX-License-Identifier: GPL-3.0-or-later
//
// Tab stops (ED-003 section 5.6, step C5): a paragraph's effective stops, which the resolver
// merges through the whole chain (the catalogue's `mergeTabs`, the port of docx4j's
// `StyleUtil.apply(Tabs, Tabs)`: union, a stop at the same position replacing the inherited
// one, `w:clear` removing only, ascending order), and the pure computation of where a tab
// character ends: the next stop past its start, else the next multiple of the settings part's
// `w:defaultTabStop`. The start is a layout read the view makes; this module never lays out.
import type { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { IndentAttrs } from './schema.mjs';
import { paragraphPropertiesFor } from './commands/reads.mjs';
import { sessionOf } from './commands/spec.mjs';

/** Word's default interval, when the settings part states none (ECMA-376 17.15.1.25). */
export const DEFAULT_TAB_STOP = 720;

export interface TabStop {
  /** Twips from the leading edge of the text column. */
  pos: number;
  /** `left`, `center`, `right`, `decimal`, `num`, `start`, `end` (`bar` and `clear` are not stops a tab moves to). */
  val: string;
  leader?: string | undefined;
}

/** The stops of an effective `w:tabs`, in ascending order, without `clear` and `bar` entries. */
export function tabStopsOf(tabs: wml.Tabs | undefined): TabStop[] {
  return (tabs?.tab ?? [])
    .filter((t) => t.pos !== undefined && t.val !== 'clear' && t.val !== 'bar')
    .map((t) => ({ pos: t.pos, val: String(t.val), leader: t.leader === undefined ? undefined : String(t.leader) }))
    .sort((a, b) => a.pos - b.pos);
}

/**
 * A paragraph node's effective stops and the indent it lays out with, through the resolver:
 * the source element's `w:pPr` with the node's attributes applied (its own `w:tabs` from the
 * `tabs` attribute since F4, ED-003 section 8.5), folded through the chain. Empty when the
 * session has no resolver.
 */
export function paragraphTabs(node: PMNode, state: EditorState): { stops: TabStop[]; indent: IndentAttrs | null } {
  const resolver = sessionOf(state)?.pkg.propertyResolverOrUndefined;
  if (!resolver) return { stops: [], indent: null };
  const effective = resolver.getEffectivePPr(paragraphPropertiesFor(node, state));
  const ind = effective?.ind;
  const indent: IndentAttrs | null = ind
    ? { start: ind.start ?? ind.left ?? null, end: ind.end ?? ind.right ?? null, firstLine: ind.firstLine ?? null, hanging: ind.hanging ?? null }
    : null;
  return { stops: tabStopsOf(effective?.tabs), indent };
}

/**
 * Where a tab starting at `x` (twips from the column's leading edge) ends: the first stop past
 * `x`; else the first multiple of the default interval past both `x` and the last custom stop,
 * since Word's default stops begin after the custom ones (ECMA-376 17.3.1.38). A paragraph
 * with a hanging indent has Word's implicit stop at its left indent, where the lines sit.
 * The answer is the stop's position; for a centre, right or decimal stop the tab's width is
 * shorter by the text that follows it (`tabWidth`), which is the view's to measure.
 */
export function nextTabStop(x: number, stops: readonly TabStop[], defaultInterval = DEFAULT_TAB_STOP, indent: IndentAttrs | null = null): { pos: number; stop?: TabStop } {
  const interval = defaultInterval > 0 ? defaultInterval : DEFAULT_TAB_STOP;
  const candidates: TabStop[] = [...stops];
  if (indent && indent.hanging !== null && indent.hanging > 0 && indent.start !== null && indent.start > 0) {
    candidates.push({ pos: indent.start, val: 'left' });
    candidates.sort((a, b) => a.pos - b.pos);
  }
  for (const stop of candidates) if (stop.pos > x) return { pos: stop.pos, stop };
  const last = candidates.length > 0 ? candidates[candidates.length - 1]!.pos : 0;
  const from = Math.max(x, last);
  return { pos: (Math.floor(from / interval) + 1) * interval };
}

/** What the view measured after a tab: the text up to the next tab or the paragraph's end, in twips, and the part of it before the first decimal separator (null when it has none). */
export interface FollowingText {
  width: number;
  toDecimal: number | null;
}

/**
 * A tab's width in twips (ED-003 section 11.20, E2.f step F5): to the stop for a left stop
 * and for the defaults; for a right stop, less the text that follows the tab, so that text
 * ends at the stop; for a centre stop, less half of it; for a decimal stop, less the text
 * before the decimal separator (the whole text when it has none, as Word does). Text wider
 * than the room collapses the tab to nothing and the text simply follows it, as in Word.
 */
export function tabWidth(x: number, next: { pos: number; stop?: TabStop }, following: FollowingText): number {
  const room = next.pos - x;
  const val = next.stop?.val;
  let taken = 0;
  if (val === 'right' || val === 'end') taken = following.width;
  else if (val === 'center') taken = following.width / 2;
  else if (val === 'decimal') taken = following.toDecimal ?? following.width;
  return Math.max(0, room - taken);
}
