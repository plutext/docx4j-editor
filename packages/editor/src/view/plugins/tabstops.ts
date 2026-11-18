// SPDX-License-Identifier: GPL-3.0-or-later
//
// Tab widths (ED-003 section 5.6, step C5): each `tab` atom is as wide as the distance from
// where it starts to the paragraph's next effective stop (the model's `nextTabStop` over the
// resolver's merged `w:tabs` and the settings part's default interval). Where a tab starts is
// a layout read, so after a transaction that changed a paragraph holding a tab, the tabs of
// those paragraphs are measured in one pass after the paint (`coordsAtPos` against the text
// column's leading edge, each width written to the DOM before the next tab is read, since it
// moves it), and one transaction carries the widths as a meta, from which this plugin's node
// decorations set `inline-size` on the atoms. Confined to paragraphs that have a tab: a
// document with none pays nothing. For a centre, right or decimal stop the text after the tab
// (to the next tab or the paragraph's end) is measured in the same pass and the tab is shorter
// by it, so that text ends at, is centred on, or has its decimal separator at the stop (ED-003
// section 11.20); a leader is drawn by CSS from the stop's `w:leader`.
import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { scaleOf } from '../zoom.js';
import { paragraphTabs, nextTabStop, tabWidth, TWIPS_PER_PX, type EditorDocument, type ParagraphAttrs, type FollowingText } from '@docx4j/editor-model';

/** A measurement's result: the paragraphs measured (their ranges) and each tab's width. */
export interface TabWidths {
  ranges: [number, number][];
  tabs: { pos: number; width: number; leader?: string | undefined }[];
}

export const TAB_WIDTHS = new PluginKey<TabWidths>('docx4j-tab-widths');
/** Set by the view after a stylesheet change: every tab is measured again. */
export const TABS_REMEASURE = new PluginKey<boolean>('docx4j-tabs-remeasure');

export interface TabStopsState {
  decos: DecorationSet;
  /** Paragraph positions waiting for a measurement, or every paragraph with a tab. */
  pending: number[] | 'all' | null;
  /** How many measurements were dispatched (tests). */
  measured: number;
}

export const tabStopsKey = new PluginKey<TabStopsState>('docx4j-tab-stops');

const hasTab = (paragraph: PMNode): boolean => {
  let found = false;
  paragraph.forEach((child) => { if (child.type.name === 'tab') found = true; });
  return found;
};

/** The paragraphs with a tab in the ranges a transaction touched (its steps' maps). */
function paragraphsTouched(tr: Transaction, doc: PMNode): number[] {
  const out = new Set<number>();
  tr.steps.forEach((step, i) => {
    const rest = tr.mapping.slice(i + 1);
    step.getMap().forEach((_os, _oe, newStart, newEnd) => {
      const from = rest.map(newStart, -1);
      const to = rest.map(newEnd, 1);
      doc.nodesBetween(Math.max(0, from), Math.min(to, doc.content.size), (node, pos) => {
        if (node.type.name === 'paragraph') { if (hasTab(node)) out.add(pos); return false; }
        return true;
      });
    });
  });
  return [...out];
}

/** Every paragraph with a tab. */
function allParagraphsWithTabs(doc: PMNode): number[] {
  const out: number[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name === 'paragraph') { if (hasTab(node)) out.push(pos); return false; }
    return true;
  });
  return out;
}

function decorationsOf(doc: PMNode, previous: DecorationSet, widths: TabWidths): DecorationSet {
  let set = previous;
  for (const [from, to] of widths.ranges) set = set.remove(set.find(from, to));
  const added = widths.tabs.map((t) => Decoration.node(t.pos, t.pos + 1, {
    style: `inline-size: ${t.width}px`,
    ...(t.leader && t.leader !== 'none' ? { 'data-leader': t.leader } : {}),
  }));
  return set.add(doc, added);
}

/**
 * The measurement (after the paint): each pending paragraph's tabs from the column's leading
 * edge, the width from the model. Reads and writes are batched in passes so the document is
 * laid out again a few times rather than once per tab: a pass reads every tab's start, then
 * writes the widths that changed; the first tab of a paragraph is right after the first pass,
 * the second after the second (its start moved with the first's width), and a pass that
 * changes nothing ends it. In an environment that cannot lay out (jsdom) every start reads as
 * the edge, which gives each tab the first stop's width: harmless, and the flow is what a test
 * there checks.
 */
function measure(view: EditorView, session: EditorDocument, pending: number[] | 'all'): TabWidths {
  const state = view.state;
  const doc = state.doc;
  const positions = pending === 'all' ? allParagraphsWithTabs(doc) : pending.filter((p) => { const n = doc.nodeAt(p); return n?.type.name === 'paragraph' && hasTab(n); });
  const out: TabWidths = { ranges: [], tabs: [] };
  const column = view.dom.getBoundingClientRect();
  // The reads are client pixels, zoomed with the page; the widths written are layout pixels (`view/zoom.ts`).
  const scale = scaleOf(view.dom);
  const paragraphs = positions.map((pos) => {
    const paragraph = doc.nodeAt(pos)!;
    out.ranges.push([pos, pos + paragraph.nodeSize]);
    const tabs: number[] = [];
    paragraph.forEach((child, offset) => { if (child.type.name === 'tab') tabs.push(pos + 1 + offset); });
    return { pos, paragraph, tabs, end: pos + paragraph.nodeSize - 1, ...paragraphTabs(paragraph, state), rtl: (paragraph.attrs as ParagraphAttrs).bidi === true };
  });
  const x = (pos: number): { x: number; top: number; bottom: number } | undefined => {
    try { const c = view.coordsAtPos(pos); return { x: c.left, top: c.top, bottom: c.bottom }; } catch { return undefined; }
  };
  // On one line when the vertical extents overlap: a tab atom's box sits a pixel or two off the text's, so equal tops would not do.
  const sameLine = (a: { top: number; bottom: number }, b: { top: number; bottom: number }): boolean => a.top < b.bottom && b.top < a.bottom;
  /**
   * The text after a tab up to the next tab or the paragraph's end, in twips, and the part of
   * it before the first decimal separator: the difference of two layout reads on the same
   * line. Text that wraps to another line is wider than any room, except before a decimal
   * stop's separator, which is all such a stop needs.
   */
  const following = (p: typeof paragraphs[number], i: number): FollowingText => {
    const from = p.tabs[i]! + 1;
    const to = i + 1 < p.tabs.length ? p.tabs[i + 1]! : p.end;
    const none: FollowingText = { width: 0, toDecimal: null };
    if (to <= from) return none;
    const start = x(from);
    const end = x(to);
    if (!start || !end) return none;
    // Text that wraps is wider than any room, which collapses a right or centre stop's tab. A decimal
    // stop needs only the part before its separator, and what follows the separator may wrap, as in
    // Word. Reading the wrap as "no separator" collapsed a decimal tab, its text then fitted and the tab
    // widened again: the passes swung between the two and the fonts decided where the last one left it
    // (CI's did, ED-005 section 12.42).
    const width = sameLine(start, end) ? (Math.abs(end.x - start.x) / scale) * TWIPS_PER_PX : Number.MAX_SAFE_INTEGER;
    const text = doc.textBetween(from, to, '\u0000', '\u0000');
    const sep = text.indexOf(session.decimalSymbol);
    let toDecimal: number | null = null;
    if (sep >= 0) {
      const at = x(from + sep);
      if (at && sameLine(start, at)) toDecimal = (Math.abs(at.x - start.x) / scale) * TWIPS_PER_PX;
    }
    return { width, toDecimal };
  };
  const widths = new Map<number, { width: number; leader?: string | undefined }>();
  let remaining = paragraphs;
  for (let pass = 0; pass < 8 && remaining.length > 0; pass++) {
    // Reads.
    const wanted: { at: number; width: number; leader?: string | undefined; owner: typeof remaining[number] }[] = [];
    for (const p of remaining) {
      p.tabs.forEach((at, i) => {
        let startPx = 0;
        try {
          const c = view.coordsAtPos(at);
          startPx = (p.rtl ? column.right - c.left : c.left - column.left) / scale;
        } catch { startPx = 0; }
        const start = Math.max(0, startPx) * TWIPS_PER_PX;
        const next = nextTabStop(start, p.stops, session.defaultTabStop, p.indent);
        const aligned = next.stop !== undefined && next.stop.val !== 'left' && next.stop.val !== 'start' && next.stop.val !== 'num';
        const twipsWide = tabWidth(start, next, aligned ? following(p, i) : { width: 0, toDecimal: null });
        const width = Math.max(1, Math.round((twipsWide / TWIPS_PER_PX) * 100) / 100);
        wanted.push({ at, width, leader: next.stop?.leader, owner: p });
      });
    }
    // Writes, for the tabs whose width changed; their paragraphs are read again.
    const changed = new Set<typeof remaining[number]>();
    for (const w of wanted) {
      const before = widths.get(w.at);
      widths.set(w.at, { width: w.width, leader: w.leader });
      if (before && Math.abs(before.width - w.width) < 0.5) continue;
      const dom = view.nodeDOM(w.at) as HTMLElement | null;
      if (dom?.style) dom.style.inlineSize = `${w.width}px`;
      // A later tab of the paragraph starts where this one's width puts it; and a tab shortened by its following text may have moved that text off a wrap, so it is read again too.
      if (w.owner.tabs.length > 1 || w.owner.stops.some((s) => s.val !== 'left')) changed.add(w.owner);
    }
    remaining = [...changed];
  }
  for (const [pos, w] of widths) out.tabs.push({ pos, width: w.width, leader: w.leader });
  out.tabs.sort((a, b) => a.pos - b.pos);
  return out;
}

const schedule = (fn: () => void): void => { if (typeof requestAnimationFrame === 'function') requestAnimationFrame(fn); else setTimeout(fn, 0); };

export function tabStopsPlugin(session: EditorDocument): Plugin<TabStopsState> {
  return new Plugin<TabStopsState>({
    key: tabStopsKey,
    state: {
      init: (_config, state) => ({ decos: DecorationSet.empty, pending: allParagraphsWithTabs(state.doc).length > 0 ? 'all' : null, measured: 0 }),
      apply: (tr, value, _old, state) => {
        const widths = tr.getMeta(TAB_WIDTHS) as TabWidths | undefined;
        if (widths) return { decos: decorationsOf(state.doc, value.decos, widths), pending: null, measured: value.measured + 1 };
        if (tr.getMeta(TABS_REMEASURE)) return { ...value, pending: 'all' };
        if (!tr.docChanged) return value;
        const decos = value.decos.map(tr.mapping, tr.doc);
        const touched = paragraphsTouched(tr, state.doc);
        if (value.pending === 'all') return { decos, pending: 'all', measured: value.measured };
        const carried = (value.pending ?? []).map((p) => tr.mapping.map(p));
        const pending = [...new Set([...carried, ...touched])];
        return { decos, pending: pending.length > 0 ? pending : null, measured: value.measured };
      },
    },
    props: {
      decorations(state: EditorState) {
        return tabStopsKey.getState(state)?.decos ?? DecorationSet.empty;
      },
    },
    view(view) {
      let queued = false;
      const run = (): void => {
        queued = false;
        const pending = tabStopsKey.getState(view.state)?.pending;
        if (!pending || !view.dom.isConnected) return;
        const widths = measure(view, session, pending);
        view.dispatch(view.state.tr.setMeta(TAB_WIDTHS, widths).setMeta('addToHistory', false));
      };
      const check = (): void => {
        if (queued || !tabStopsKey.getState(view.state)?.pending) return;
        queued = true;
        schedule(run);
      };
      check();
      return { update: check, destroy: () => { queued = false; } };
    },
  });
}
