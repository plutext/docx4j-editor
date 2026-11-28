// SPDX-License-Identifier: GPL-3.0-or-later
//
// Rendering-only additions (ED-001 Appendix A; ED-002 sections 5.3 and 7.4): section classes and
// section-break rules, page boundaries from Word's markers and from page-starting section breaks
// (Appendix D.1, muted from the first edited block on) carrying the footer of the page that ends
// and the header of the page that starts, list labels from the Emulator, footnote reference
// numbers, font-by-script spans from the run-font selector, and in a story the field results the
// page decides (`PAGE`, `NUMPAGES`). The body's walk is whole on every change and memoised per
// top-level block (ED-003 section 8.2, section 9 item 32): a block a transaction did not touch
// is the same node object afterwards, so what the walk learned about it is read back rather
// than walked again; a story's walk is small and stays plain.
import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { TableMap } from 'prosemirror-tables';
import type { Node as PMNode } from 'prosemirror-model';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { Element } from '@docx4j/core-ts';
import { type EditorDocument, runPropertiesFor, paragraphPropertiesFor, sessionPlugin, formatNumber, noteLabels, type LiveLabel, type SourceIds } from '@docx4j/editor-model';
import { EditorState as PMEditorState } from 'prosemirror-state';
import { fontStack, sectionsNeedingClass, characterStyleRPr, styleOnText, cssId } from '../css.js';
import { fieldKeyword } from '../nodeviews/index.js';
import { t } from '../../app/i18n.js';
import { labelLookup, labelWalksOf, PARTS_CHANGED, type LabelLookup } from './labels.js';
import { proofingOf, documentLanguage, type RelProofing } from '../proofing.js';

export const decorationsKey = new PluginKey<DecorationsState>('docx4j-decorations');

/** The plugin's state: the decorations, and per boundary widget key what its DOM shows (applied in place, so a widget survives a staleness change; ED-003 section 6.4). */
export interface DecorationsState {
  set: DecorationSet;
  boundaries: Map<string, BoundaryStatus>;
  /** The number each note reference shows, by `kind:id` (ED-003 section 6.3: `w:footnotePr` honoured); the panel prints the same. */
  noteLabels: Map<string, string>;
}

export interface BoundaryStatus {
  stale: boolean;
  label: string;
}

/** The section types a break widget names (`view.sectionBreak.<type>`). */
const SECTION_TYPES = new Set(['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn']);

/** The meta a transaction carries when the application's language changed: every widget is made again in it (surface.ts). */
export const LANGUAGE_CHANGED = 'docx4j-language-changed';

/** One side of a page boundary: which page, in which section, and whether the page opens its section. */
export interface PageSide {
  /** The physical page, counted from 1: `NUMPAGES` and the order of the boundaries. */
  page: number;
  /** The number the page shows, from `w:pgNumType` (`start` restarts a section's count, `fmt` formats it; ED-003 section 6.4). */
  numeric: number;
  /** `numeric` formatted: `5`, `iv`, `C`. */
  number: string;
  section: number;
  firstOfSection: boolean;
}

export interface BoundaryInfo {
  kind: 'start' | 'marker' | 'section' | 'end';
  /** The page that ends above the boundary; none at the document's start. */
  above?: PageSide;
  /** The page that starts below it; none at the document's end. */
  below?: PageSide;
  /** Pages in the document, known once the walk is over (widgets draw later). */
  total: () => number;
  /** The sections of the document walked, complete once the walk is over, for `Stories.headerFooterFor`. */
  sections: (wml.SectPr | null)[];
  /** Below the first edited block: the boundary is where Word last saw it, not where it is. */
  stale: boolean;
}

/** Slots in a boundary widget for the footer above and the header below; returns the teardown. */
export type BoundaryMount = (slots: { above: HTMLElement; below: HTMLElement }, info: BoundaryInfo) => () => void;

export interface DecorationOptions {
  /**
   * A story (a footnote, a comment, a header; ED-002 section 7.1) rather than the body: its
   * nodes were projected into a context of their own, so the body's imported map does not know
   * them and would call them all changed; and its page-break markers, sections and reference
   * numbers mean nothing outside the body.
   */
  story?: boolean;
  /** In a story: complex fields whose keyword is here show this result rather than the cached one. */
  fields?: Record<string, string>;
  /** In the body: mount the footer and header at each page boundary (ED-002 section 7.4). */
  boundary?: BoundaryMount;
  /** The ids a story was projected with, when not the session's (the preview). */
  ids?: SourceIds;
  /** The plugin's registry of mounted boundary widgets by key, filled by the widgets and read by the plugin's view update. */
  boundaryRegistry?: Map<string, HTMLElement>;
}

function widget(className: string, text: string, tag = 'span'): HTMLElement {
  const el = document.createElement(tag);
  el.className = className;
  el.textContent = text;
  el.contentEditable = 'false';
  return el;
}

/** True when the block is not the node it was imported as (identity, then `eq`). */
function changed(session: EditorDocument, node: PMNode): boolean {
  const id = node.attrs['sourceId'] as string | null;
  if (!id) return true;
  const imported = session.ctx.imported.get(id);
  return !imported || (imported !== node && !imported.eq(node));
}

/** Whether a paragraph's text needs per-script fonts: anything outside Latin, digits and punctuation. */
const NON_LATIN = /[^ -ɏḀ-ỿ -⁯₠-⃏℀-⅏]/;

/**
 * A page boundary's spec says every event inside it is the story's (a header's, a footer's), never the body view's
 * (`eventBelongsToView` asks each widget up to the view), without stopping the event's bubble (ED-014 section 4.4).
 */
const storyEvent = (): boolean => true;

/** A page boundary widget: the footer slot, the label, the header slot; the mount fills the slots and is torn down with the widget. */
function boundaryWidget(className: string, key: string, label: string, info: BoundaryInfo, mount: BoundaryMount | undefined, registry: Map<string, HTMLElement> | undefined, row?: { columns: number }): { dom: HTMLElement; destroy: () => void } {
  const above = widget('boundary-footer', '', 'div');
  const text = widget('boundary-label', label, 'div');
  const below = widget('boundary-header', '', 'div');
  let dom: HTMLElement;
  if (row) {
    // A boundary inside a table (ED-003 section 11.14): a row of its own spanning the columns, between the table's rows, never inside a cell.
    dom = widget(`${className} in-table`, '', 'tr');
    const td = document.createElement('td');
    td.colSpan = row.columns;
    td.contentEditable = 'false';
    td.append(above, text, below);
    dom.append(td);
  } else {
    dom = widget(className, '', 'div');
    dom.append(above, text, below);
  }
  // A story view inside the body's editable surface: its events are its own. The widget's spec keeps the body view from
  // reading any of them (`stopEvent`, below), so the keys go on to the page's listeners and, in VS Code, to the key gate and
  // the workbench (ED-014 section 4.4): Ctrl+S saves from a header. The mouse events stop here, so that nothing around the
  // body (its selection bar, its table and column plugins' listeners) reads a selection from inside.
  for (const type of ['mousedown', 'mouseup', 'click']) dom.addEventListener(type, (e) => e.stopPropagation());
  registry?.set(key, dom);
  const unmount = mount ? mount({ above, below }, info) : () => {};
  const destroy = (): void => { unmount(); if (registry?.get(key) === dom) registry.delete(key); };
  return { dom, destroy };
}

/**
 * The label's widget (ED-003 section 5.4): with the level's default `w:suff` (a tab) it fills
 * the hanging indent, which is the space the paragraph's negative `text-indent` opens for it;
 * a space suffix is one space's worth; `nothing` is none. A level without a hanging indent
 * keeps the stylesheet's fixed width and gap.
 */
function labelWidget(label: LiveLabel): HTMLElement {
  const el = widget(`list-label${label.isBullet ? ' bullet' : ''}`, label.listString);
  const hanging = label.indent?.hanging ?? null;
  if (label.suff === 'tab' && hanging !== null && hanging > 0) el.style.cssText = `inline-size: ${hanging / 20}pt; min-inline-size: 0; margin-inline-end: 0`;
  else if (label.suff === 'space') el.style.cssText = 'min-inline-size: 0; margin-inline-end: 0.25em';
  else if (label.suff === 'nothing') el.style.cssText = 'min-inline-size: 0; margin-inline-end: 0';
  return el;
}

/** What resolves a run's theme colour: the package's `PropertyResolver`, once it is built. */
type Resolver = NonNullable<EditorDocument['pkg']['propertyResolverOrUndefined']>;

/** An inline font span of a paragraph, relative to the paragraph's position. */
interface RelSpan { from: number; to: number; font: string }

/** A stretch of text drawn in a run's theme colour, relative to the paragraph's position. */
interface RelColour { from: number; to: number; color: string }

/** A stretch of text whose run switches a property off itself, and the classes that draw it so, relative to the paragraph's position. */
interface RelOff { from: number; to: number; classes: string }

/** What the walk meets inside one top-level block, at offsets from the block's position, in order. */
type BlockEvent =
  | { kind: 'paragraph'; offset: number; node: PMNode; changed: boolean; spans: RelSpan[]; colours: RelColour[]; proofing: RelProofing[]; off: RelOff[]; styled: RelOff[] }
  | { kind: 'marker'; offset: number; row?: { offset: number; columns: number } }
  | { kind: 'ref'; offset: number; size: number; key: string };

interface BlockEntry {
  events: BlockEvent[];
  /** The block's part in the document's sequence of markers, sections and references: unchanged means the page and note numbers are. */
  sig: string;
  /** Marks the block's decorations (`spec.block`), so a patch can take them out by block. */
  token: object;
}

/**
 * The memo of the body's walk (ED-003 section 8.2): one entry per top-level block, keyed by the
 * node object, which ProseMirror keeps for a block a transaction did not touch. Reset when the
 * session's projection context changes (a tree run re-imports: what "changed" means moves).
 */
export class DecorationCache {
  entries = new WeakMap<PMNode, BlockEntry>();
  ctx: EditorDocument['ctx'] | undefined;
  /** Blocks walked afresh since the counter was last read (tests). */
  misses = 0;
  /** The last walk's sequence signature and label walks: equal again means the previous set can be mapped and patched rather than rebuilt. */
  lastSignature: string | undefined;
  lastWalks: number | undefined;
  /** Whether the last walk patched the previous set (tests). */
  patched = false;
}

/** A short hash of a string, empty for the empty string: a suffix for keys that must change with it. */
function hashOf(text: string): string {
  if (text === '') return '';
  let h = 5381;
  for (let i = 0; i < text.length; i++) h = ((h * 33) ^ text.charCodeAt(i)) >>> 0;
  return `-${h.toString(36)}`;
}

/** The font spans of a paragraph's text, relative to the paragraph (ED-001 section 3.7). */
function relSpans(state: EditorState, session: EditorDocument, paragraph: PMNode, selector: NonNullable<EditorDocument['main']['runFontSelectorOrUndefined']>): RelSpan[] {
  const out: RelSpan[] = [];
  const pPr = paragraphPropertiesFor(paragraph, state);
  paragraph.forEach((child, offset) => {
    if (!child.isText || !NON_LATIN.test(child.text ?? '')) return;
    const runSource = child.marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'] as string | undefined;
    const source = runSource ? (session.ctx.ids.get(runSource) as Element<wml.R> | undefined) : undefined;
    const rPr = runPropertiesFor(source?.value.rPr, child.marks);
    let spans;
    try { spans = selector.spans(pPr, rPr, child.text ?? ''); } catch { return; }
    let at = 1 + offset;
    for (const span of spans) {
      const end = at + span.text.length;
      if (span.documentFont) out.push({ from: at, to: end, font: span.documentFont });
      at = end;
    }
  });
  return out;
}

/**
 * The text of a paragraph a run's own theme colour draws differently from the value the `font` mark holds (ED-008
 * section 8.8): the projection reads `w:val`, a copy Word rewrites and which can be stale, where Word draws the
 * reference beside it. Only what differs is listed, so a file whose copies agree gets no decoration. Read through
 * what the marks mean, since a colour the person set has replaced the reference. The mark keeps `w:val`: holding the
 * resolved colour there would read as an edit, and the first change to the paragraph would write it over the reference.
 */
function themeColours(paragraph: PMNode, ids: SourceIds, resolver: Resolver): RelColour[] {
  const out: RelColour[] = [];
  paragraph.descendants((child, rel) => {
    if (!child.isText) return true;
    const runSource = child.marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'] as string | undefined;
    const source = runSource ? (ids.get(runSource) as Element<wml.R> | undefined) : undefined;
    if (!source?.value.rPr?.color?.themeColor) return false;
    const color = runPropertiesFor(source.value.rPr, child.marks)?.color;
    const drawn = color?.themeColor ? resolver.resolveColor(color) : undefined;
    if (!drawn) return false;
    const hex = `#${drawn.toHex()}`;
    const written = child.marks.find((m) => m.type.name === 'font')?.attrs['color'] as string | null | undefined;
    if (hex.toLowerCase() !== written?.toLowerCase()) out.push({ from: 1 + rel, to: 1 + rel + child.nodeSize, color: hex });
    return false;
  });
  return out;
}

/**
 * A theme colour's decoration: a class and a variable rather than `color` itself, since a decoration's element is the
 * innermost around the text and a tracked change's author colour, on the `ins` outside it, must still win (`editor.css`).
 */
function themeColourDeco(from: number, to: number, color: string, spec?: { block: object }): Decoration {
  return Decoration.inline(from, to, { class: 'theme-colour', style: `--theme-colour: ${color}` }, spec);
}

/** The properties a run can switch off itself that the stylesheet draws on, and the class that draws each off (`editor.css`). */
const SWITCHES: readonly (readonly ['b' | 'i' | 'caps' | 'smallCaps', string])[] = [['b', 'not-bold'], ['i', 'not-italic'], ['caps', 'not-caps'], ['smallCaps', 'not-small-caps']];

const isOff = (v: { val?: boolean } | undefined): boolean => v !== undefined && v.val === false;

/**
 * The text of a paragraph whose run switches bold, italic, caps or small caps off itself (`<w:b w:val="0"/>`), where a
 * paragraph or character style may switch it on (ED-005 section 12.132). The marks hold only what is on, so without
 * this the view drew the style's bold. Read through what the marks mean, since bold the person set replaces the "off".
 * Only a run whose own properties say "off" is read further, so a file without one gets no decoration.
 */
function switchedOff(paragraph: PMNode, ids: SourceIds): RelOff[] {
  const out: RelOff[] = [];
  paragraph.descendants((child, rel) => {
    if (!child.isText) return true;
    const runSource = child.marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'] as string | undefined;
    const own = runSource ? (ids.get(runSource) as Element<wml.R> | undefined)?.value.rPr : undefined;
    if (!own || !SWITCHES.some(([k]) => isOff(own[k]))) return false;
    const rPr = runPropertiesFor(own, child.marks);
    const classes = SWITCHES.filter(([k]) => isOff(rPr?.[k])).map(([, c]) => c);
    if (classes.length > 0) out.push({ from: 1 + rel, to: 1 + rel + child.nodeSize, classes: classes.join(' ') });
    return false;
  });
  return out;
}

/**
 * The text of a paragraph under a character style whose size or family its run does not state itself, and the classes
 * that draw the style's there (`rz-<id>`, `rf-<id>`; ED-005 section 12.133). The style's element no longer carries
 * them, so its box is the paragraph's and a line is as tall as the text drawn on it, as in Word (`styleOnText`). A run's
 * own size or family is its `font` mark's, inside the style's element, and needs nothing.
 */
function styleFonts(paragraph: PMNode, resolver: Resolver): RelOff[] {
  const out: RelOff[] = [];
  const styles = new Map<string, { size: boolean; family: boolean }>();
  paragraph.descendants((child, rel) => {
    if (!child.isText) return true;
    const id = child.marks.find((m) => m.type.name === 'runStyle')?.attrs['id'];
    if (typeof id !== 'string') return false;
    let onText = styles.get(id);
    if (!onText) { onText = styleOnText(characterStyleRPr(resolver, id)); styles.set(id, onText); }
    const font = child.marks.find((m) => m.type.name === 'font');
    const classes: string[] = [];
    if (onText.size && font?.attrs['size'] == null) classes.push(`rz-${cssId(id)}`);
    if (onText.family && font?.attrs['name'] == null) classes.push(`rf-${cssId(id)}`);
    if (classes.length > 0) out.push({ from: 1 + rel, to: 1 + rel + child.nodeSize, classes: classes.join(' ') });
    return false;
  });
  return out;
}

/** The row of the outermost table around a position inside a block, as an offset from the block, for a marker hoisted out of a cell (section 11.14). */
function rowWithin(block: PMNode, offset: number): { offset: number; columns: number } | undefined {
  if (offset === 0) return undefined;
  const $b = block.resolve(offset - 1);
  for (let d = 0; d <= $b.depth; d++) {
    if ($b.node(d).type.name !== 'table') continue;
    return { offset: 1 + $b.before(d + 1), columns: TableMap.get($b.node(d)).width };
  }
  return undefined;
}

/** Walk one top-level block once: its paragraphs (with whether each is changed and its font spans), its page-break markers and its note references. */
function collectBlock(block: PMNode, session: EditorDocument, withSession: EditorState, selector: EditorDocument['main']['runFontSelectorOrUndefined'], resolver: Resolver | undefined): BlockEntry {
  const events: BlockEvent[] = [];
  const visit = (node: PMNode, offset: number): boolean => {
    if (node.type.name === 'paragraph') {
      const spans = selector && resolver && NON_LATIN.test(node.textContent) ? relSpans(withSession, session, node, selector) : [];
      const colours = resolver ? themeColours(node, session.ctx.ids, resolver) : [];
      const proofing = resolver ? proofingOf(node, session.ctx.ids, resolver, documentLanguage(session)) : [];
      events.push({ kind: 'paragraph', offset, node, changed: changed(session, node), spans, colours, proofing, off: switchedOff(node, session.ctx.ids), styled: resolver ? styleFonts(node, resolver) : [] });
      return true;
    }
    if (node.type.name === 'opaqueInline' && node.attrs['kind'] === 'lastRenderedPageBreak') {
      const row = rowWithin(block, offset);
      events.push(row ? { kind: 'marker', offset, row } : { kind: 'marker', offset });
      return false;
    }
    if (node.type.name === 'footnoteRef') {
      events.push({ kind: 'ref', offset, size: node.nodeSize, key: `${String(node.attrs['kind'])}:${String(node.attrs['footnoteId'])}` });
      return false;
    }
    return !node.isLeaf;
  };
  if (visit(block, 0)) block.descendants((node, rel) => visit(node, 1 + rel));
  const sig = events.map((e) => (e.kind === 'marker' ? 'M' : e.kind === 'ref' ? `R${e.key}` : e.node.attrs['sectPr'] ? 'S' : '')).join('');
  return { events, sig, token: {} };
}

export function computeDecorations(session: EditorDocument, state: EditorState, labels: LabelLookup, options: DecorationOptions = {}, cache: DecorationCache = new DecorationCache(), previous?: DecorationsState, tr?: Transaction): DecorationsState {
  const decos: Decoration[] = [];
  const boundaryStatus = new Map<string, BoundaryStatus>();
  let section = 0;
  let page = 1;
  let firstOfSection = true;
  let stale = false;
  const story = options.story === true;
  const boundaries = options.boundary !== undefined;
  const total = (): number => page;
  // Filled as the walk meets section-break paragraphs, then the final section: the widgets draw later.
  const sections: (wml.SectPr | null)[] = [];
  const selector = session.main.runFontSelectorOrUndefined;
  const resolver = session.pkg.propertyResolverOrUndefined;
  const withSession = state.plugins.length > 0 ? state : PMEditorState.create({ doc: state.doc, plugins: [sessionPlugin(session)] });
  // The sections up front (the same list the walk builds), for the section classes and the page numbers.
  const secs = story ? [] : session.stories.sections(state.doc);
  // The sections' header and footer references, hashed into the signature and the boundary keys: a reference added
  // (a page number's footer, ED-005 section 12.95) remakes the boundary widgets, whose mounts looked the stories up.
  const refsSig = hashOf(secs.map((sp) => ((sp?.egHdrFtrReferences ?? []) as { type?: string; id?: string }[]).map((r) => `${r.type ?? ''}=${r.id ?? ''}`).join('+')).join(','));
  // Only paragraphs of a section narrower or wider than the first carry a section class: one
  // node decoration per paragraph of a long document made every keystroke quadratic (item 10).
  const classed = story ? new Set<number>() : sectionsNeedingClass(secs);
  // The page number a page shows (ED-003 section 6.4): `w:pgNumType/@start` restarts the count at
  // a section's first page, `@fmt` formats it; the physical count stays for NUMPAGES and the keys.
  const pgNum = (s: number): wml.CTPageNumber | undefined => secs[s]?.pgNumType;
  // The notes' numbers (section 6.3): a walk of their own over the references, only for a document
  // with a notes part; read back from the previous state when the walk below finds it can patch.
  let notes = new Map<string, string>();
  const notesFor = (patching: boolean): Map<string, string> => (!story && session.hasNotes ? (patching && previous ? previous.noteLabels : noteLabels(state.doc, secs, session.noteSettings)) : new Map<string, string>());
  let numeric = pgNum(0)?.start ?? 1;
  const side = (): PageSide => ({ page, numeric, number: formatNumber(numeric, pgNum(section)?.fmt), section, firstOfSection });
  const withDestroy = (w: { dom: HTMLElement; destroy: () => void }): HTMLElement => { (w.dom as HTMLElement & { docx4jDestroy?: () => void }).docx4jDestroy = w.destroy; return w.dom; };
  const destroyOf = (dom: Node): void => { (dom as HTMLElement & { docx4jDestroy?: () => void }).docx4jDestroy?.(); };
  const startInfo: BoundaryInfo = { kind: 'start', below: side(), total, sections, stale: false };
  if (boundaries) boundaryStatus.set('page-start', { stale: false, label: '' });
  let patch = false;
  let patchedSet: DecorationSet | undefined;
  const pushBoundary = (pos: number, kind: 'marker' | 'section', above: PageSide, below: PageSide, muted: boolean, sideOf: -1 | 1, row?: { columns: number }, token?: object): void => {
    const info: BoundaryInfo = { kind, above, below, total, sections, stale: muted };
    // Word's markers, or the companion's layout (ED-001 Appendix D.1: labelled computed), until an edit makes either approximate.
    const label = t(muted ? 'view.page.approximate' : session.paginationSource === 'companion' ? 'view.page.computed' : 'view.page', { number: below.number });
    // The key leaves staleness out, so a body edit that mutes the boundaries below it keeps their widgets (and the header views in them); the plugin's view applies the status in place.
    const key = `page-${kind}-${below.page}${refsSig}`;
    boundaryStatus.set(key, { stale: muted, label });
    if (!patch || token) decos.push(Decoration.widget(pos, () => withDestroy(boundaryWidget(`page-boundary${muted ? ' stale' : ''}`, key, label, info, options.boundary, options.boundaryRegistry, row)), { side: sideOf, key, destroy: destroyOf, stopEvent: storyEvent, ...(token ? { block: token } : {}) }));
  };
  if (story) {
    // A note or a header: its own list labels (ED-003 section 6.4), script fonts and field results; no pages, no sections. Small, and walked plain.
    state.doc.descendants((node, pos) => {
      if (node.type.name !== 'paragraph') return !node.isLeaf;
      const label = labels(node);
      if (label) decos.push(Decoration.widget(pos + 1, () => labelWidget(label), { side: -1, key: `label-${label.listString}-${label.suff}-${label.indent?.hanging ?? ''}` }));
      if (selector && resolver && NON_LATIN.test(node.textContent)) scriptSpans(withSession, node, pos, selector, decos);
      if (resolver) for (const c of themeColours(node, options.ids ?? session.ctx.ids, resolver)) decos.push(themeColourDeco(pos + c.from, pos + c.to, c.color));
      if (resolver) for (const p of proofingOf(node, options.ids ?? session.ctx.ids, resolver, documentLanguage(session))) decos.push(Decoration.inline(pos + p.from, pos + p.to, p.attrs));
      for (const o of switchedOff(node, options.ids ?? session.ctx.ids)) decos.push(Decoration.inline(pos + o.from, pos + o.to, { class: o.classes }));
      if (resolver) for (const o of styleFonts(node, resolver)) decos.push(Decoration.inline(pos + o.from, pos + o.to, { class: o.classes }));
      if (options.fields) fieldResults(options.ids ?? session.ctx.ids, node, pos, options.fields, decos);
      return true;
    });
  } else {
    let reset = false;
    // A part changed too: a block the memo keeps may be drawn from a theme that is no longer the document's.
    if (cache.ctx !== session.ctx || tr?.getMeta(PARTS_CHANGED)) { cache.entries = new WeakMap(); cache.ctx = session.ctx; reset = true; }
    // First pass: the memo, and which blocks are new to it.
    const entries: { pos: number; block: PMNode; entry: BlockEntry; miss: boolean }[] = [];
    state.doc.forEach((block, pos) => {
      let entry = cache.entries.get(block);
      let miss = false;
      if (!entry) { entry = collectBlock(block, session, withSession, selector, resolver); cache.entries.set(block, entry); cache.misses++; miss = true; }
      entries.push({ pos, block, entry, miss });
    });
    // With every section's start: a break's kind is read from the section after it (another block, or the final
    // section), so a change of it must not leave the break before it patched over.
    const signature = `${entries.map((e) => e.entry.sig).join('|')}#${secs.map((sp) => (sp as { type?: { val?: string } } | null)?.type?.val ?? '').join(',')}${refsSig}`;
    const walks = labelWalksOf(state);
    // The previous set can be mapped and patched when the document's sequence of markers, sections
    // and references is what it was (so every untouched page and note number stands), no untouched
    // paragraph's label moved, and what "changed" means did not: then only the blocks walked afresh
    // need their decorations made again, and ProseMirror's view keeps every other child's for free.
    patch = previous !== undefined && tr !== undefined && tr.docChanged && !reset && signature === cache.lastSignature && walks === cache.lastWalks;
    cache.lastSignature = signature;
    cache.lastWalks = walks;
    cache.patched = patch;
    notes = notesFor(patch);
    if (patch && previous && tr) {
      // The old blocks the fresh ones replace, by mapping their range back: their decorations go.
      const invert = tr.mapping.invert();
      const gone = new Set<object>();
      let low = Number.POSITIVE_INFINITY;
      let high = 0;
      for (const e of entries) {
        if (!e.miss) continue;
        const oldFrom = invert.map(e.pos, -1);
        const oldTo = invert.map(e.pos + e.block.nodeSize, 1);
        tr.before.forEach((n, p) => { if (p < oldTo && p + n.nodeSize > oldFrom) { const t = cache.entries.get(n)?.token; if (t) gone.add(t); } });
        low = Math.min(low, e.pos);
        high = Math.max(high, e.pos + e.block.nodeSize);
      }
      let set = previous.set.map(tr.mapping, tr.doc);
      if (gone.size > 0) set = set.remove(set.find(Math.max(0, low - 1), Math.min(state.doc.content.size, high + 1), (spec) => gone.has((spec as { block?: object }).block as object)));
      patchedSet = set;
    }
    for (const { pos, entry, miss } of entries) {
      const emit = !patch || miss;
      const token = entry.token;
      for (const ev of entry.events) {
        const at = pos + ev.offset;
        if (ev.kind === 'paragraph') {
          const node = ev.node;
          if (!stale && ev.changed) stale = true;
          if (emit && classed.has(section)) decos.push(Decoration.node(at, at + node.nodeSize, { class: `sect-${section}` }, { block: token }));
          const label = emit ? labels(node) : undefined;
          if (label) decos.push(Decoration.widget(at + 1, () => labelWidget(label), { side: -1, key: `label-${label.listString}-${label.suff}-${label.indent?.hanging ?? ''}`, block: token }));
          const sectPr = node.attrs['sectPr'] as wml.SectPr | null;
          if (sectPr) {
            sections.push(sectPr);
            // The break's kind is the next section's: `w:type` says how a section starts relative to the one
            // before it (ECMA-376 17.6.22), so the break after section k is section k + 1's (ED-006 section 9.9).
            const typeVal = (secs[section + 1] as { type?: { val?: string } } | null | undefined)?.type?.val ?? 'nextPage';
            const type = t(`view.sectionBreak.${SECTION_TYPES.has(typeVal) ? typeVal : 'nextPage'}`);
            // Its index is printing's (ED-006 section 3.3): the blocks after it take the next section's page.
            const index = section;
            if (emit) decos.push(Decoration.widget(at + node.nodeSize, () => { const el = widget('section-break', t('view.sectionBreak', { type }), 'div'); el.dataset['section'] = String(index); return el; }, { side: 1, key: `section-${section}-${type}`, block: token }));
            const endsPage = typeVal !== 'continuous' && typeVal !== 'nextColumn';
            const above = side();
            section++;
            if (endsPage) {
              // Word starts the section on a new page and writes no marker for it (checked over the fixtures).
              page++;
              numeric = pgNum(section)?.start ?? numeric + 1;
              firstOfSection = true;
              if (boundaries) pushBoundary(at + node.nodeSize, 'section', above, side(), stale, 1, undefined, emit ? token : undefined);
            }
          }
          if (emit) for (const sp of ev.spans) decos.push(Decoration.inline(at + sp.from, at + sp.to, { style: `font-family: ${fontStack(sp.font)}` }, { block: token }));
          if (emit) for (const c of ev.colours) decos.push(themeColourDeco(at + c.from, at + c.to, c.color, { block: token }));
          // A run's own language, or Word's "Do not check spelling" (view/proofing.ts).
          if (emit) for (const p of ev.proofing) decos.push(Decoration.inline(at + p.from, at + p.to, p.attrs, { block: token }));
          // A run's own bold, italic or caps switched off over a style's (section 12.132).
          if (emit) for (const o of ev.off) decos.push(Decoration.inline(at + o.from, at + o.to, { class: o.classes }, { block: token }));
          if (emit) for (const o of ev.styled) decos.push(Decoration.inline(at + o.from, at + o.to, { class: o.classes }, { block: token }));
        } else if (ev.kind === 'marker') {
          const above = side();
          page++;
          numeric++;
          firstOfSection = false;
          const below = side();
          const muted = stale;
          // With headers and footers off the same widget shows the label alone (its slots stay empty).
          if (ev.row) pushBoundary(pos + ev.row.offset, 'marker', above, below, muted, -1, { columns: ev.row.columns }, emit ? token : undefined);
          else pushBoundary(at, 'marker', above, below, muted, -1, undefined, emit ? token : undefined);
        } else if (emit) {
          decos.push(Decoration.node(at, at + ev.size, { 'data-n': notes.get(ev.key) ?? '?' }, { block: token }));
        }
      }
    }
  }
  sections.push((state.doc.attrs['finalSectPr'] as wml.SectPr | null) ?? null);
  if (boundaries) {
    const info: BoundaryInfo = { kind: 'end', above: side(), total, sections, stale };
    const muted = stale;
    const key = `page-end-${page}${refsSig}`;
    boundaryStatus.set(key, { stale: muted, label: '' });
    if (!patch) {
      decos.push(Decoration.widget(0, () => withDestroy(boundaryWidget('page-start', 'page-start', '', startInfo, options.boundary, options.boundaryRegistry)), { side: -1, key: 'page-start', destroy: destroyOf, stopEvent: storyEvent }));
      decos.push(Decoration.widget(state.doc.content.size, () => withDestroy(boundaryWidget(`page-end${muted ? ' stale' : ''}`, key, '', info, options.boundary, options.boundaryRegistry)), { side: 1, key, destroy: destroyOf, stopEvent: storyEvent }));
    }
  }
  const set = patch && patchedSet ? patchedSet.add(state.doc, decos) : DecorationSet.create(state.doc, decos);
  return { set, boundaries: boundaryStatus, noteLabels: notes };
}

/**
 * Complex fields in a story (ED-002 section 7.4): between a `separate` and an `end` field
 * character sits the cached result; when the view knows the keyword's value, the result is
 * hidden and the value shown in its place.
 */
function fieldResults(ids: SourceIds, paragraph: PMNode, pos: number, fields: Record<string, string>, decos: Decoration[]): void {
  let depth = 0;
  let instr = '';
  let resultFrom = -1;
  paragraph.forEach((child, offset) => {
    if (child.type.name !== 'opaqueInline') return;
    const kind = String(child.attrs['kind']);
    if (kind !== 'fldChar' && kind !== 'instrText') return;
    const id = child.attrs['sourceId'] as string | null;
    const el = id ? ids.get(id) : undefined;
    const at = pos + 1 + offset;
    if (kind === 'instrText') {
      if (depth === 1) instr += String((el?.value as wml.Text | undefined)?.value ?? '');
      return;
    }
    const type = (el?.value as wml.FldChar | undefined)?.fldCharType;
    if (type === 'begin') { depth++; if (depth === 1) { instr = ''; resultFrom = -1; } }
    else if (type === 'separate') { if (depth === 1) resultFrom = at + child.nodeSize; }
    else if (type === 'end') {
      if (depth === 1) {
        // A field with no cached result (no run between `separate` and `end`, or no `separate` at all, as a
        // generator writes a PAGE field) shows the value Word would compute, before its end (ED-005 section 12.98).
        const from = resultFrom >= 0 && resultFrom <= at ? resultFrom : at;
        const keyword = fieldKeyword(instr);
        const value = keyword ? fields[keyword] : undefined;
        if (value !== undefined) {
          if (from < at) decos.push(Decoration.inline(from, at, { class: 'field-hidden' }));
          decos.push(Decoration.widget(from, () => { const s = widget('field', value); s.title = t('view.field', { instruction: instr.trim() }); return s; }, { side: -1, key: `field-${from}-${value}` }));
        }
      }
      depth = Math.max(0, depth - 1);
    }
  });
}

/** Inline decorations with the font of each script span of a paragraph's text (ED-001 section 3.7). */
function scriptSpans(state: EditorState, paragraph: PMNode, pos: number, selector: NonNullable<EditorDocument['main']['runFontSelectorOrUndefined']>, decos: Decoration[]): void {
  const session = sessionOfState(state);
  if (!session) return;
  const pPr = paragraphPropertiesFor(paragraph, state);
  paragraph.forEach((child, offset) => {
    if (!child.isText || !NON_LATIN.test(child.text ?? '')) return;
    const runSource = child.marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'] as string | undefined;
    const source = runSource ? (session.ctx.ids.get(runSource) as Element<wml.R> | undefined) : undefined;
    const rPr = runPropertiesFor(source?.value.rPr, child.marks);
    let spans;
    try {
      spans = selector.spans(pPr, rPr, child.text ?? '');
    } catch {
      return;
    }
    let at = pos + 1 + offset;
    for (const span of spans) {
      const end = at + span.text.length;
      if (span.documentFont) decos.push(Decoration.inline(at, end, { style: `font-family: ${fontStack(span.documentFont)}` }));
      at = end;
    }
  });
}

function sessionOfState(state: EditorState): EditorDocument | undefined {
  for (const p of state.plugins) {
    const s = p.getState(state) as unknown;
    if (s && typeof s === 'object' && 'ctx' in (s as object) && 'pkg' in (s as object)) return s as EditorDocument;
  }
  return undefined;
}

/** The cache of a plugin instance, for tests that count the blocks walked afresh. */
export function decorationCacheOf(plugin: Plugin<DecorationsState>): DecorationCache | undefined {
  return (plugin as unknown as { docx4jCache?: DecorationCache }).docx4jCache;
}

export function decorationsPlugin(session: EditorDocument, options: DecorationOptions = {}): Plugin<DecorationsState> {
  // The labels come from the labels plugin, installed before this one in the body's view and in the editable story views.
  const registry = new Map<string, HTMLElement>();
  const opts: DecorationOptions = { ...options, boundaryRegistry: registry };
  const cache = new DecorationCache();
  const plugin = new Plugin<DecorationsState>({
    key: decorationsKey,
    state: {
      init: (_config, state) => computeDecorations(session, state, labelLookup(state), opts, cache),
      apply: (tr, value, _old, state) => (tr.docChanged || tr.getMeta(LANGUAGE_CHANGED) || tr.getMeta(PARTS_CHANGED) ? computeDecorations(session, state, labelLookup(state), opts, cache, value, tr) : { ...value, set: value.set.map(tr.mapping, tr.doc) }),
    },
    props: {
      decorations(state) {
        return decorationsKey.getState(state)?.set ?? DecorationSet.empty;
      },
    },
    // Staleness and the label applied to the mounted boundary widgets in place (their keys leave it out).
    view: () => ({
      update: (view) => {
        const boundaries = decorationsKey.getState(view.state)?.boundaries;
        if (!boundaries) return;
        for (const [key, dom] of registry) {
          const status = boundaries.get(key);
          if (!status) continue;
          dom.classList.toggle('stale', status.stale);
          const label = dom.querySelector('.boundary-label');
          if (label && label.textContent !== status.label) label.textContent = status.label;
        }
      },
    }),
  });
  (plugin as unknown as { docx4jCache?: DecorationCache }).docx4jCache = cache;
  return plugin;
}
