// SPDX-License-Identifier: GPL-3.0-or-later
//
// CSS generated from the document layer (ED-001 section 5; ED-002 section 5.3): the document
// defaults, a rule per paragraph and character style in use from the resolver's effective
// properties, font stacks with metric-compatible substitutes, and the page from the sections.
// Logical properties throughout (ED-001 section 3.7). Direct formatting is not here: it goes as
// inline styles from the node and mark views.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { Node as PMNode } from 'prosemirror-model';
import { SUBSTITUTE_ROWS, classOf, readRunOptions, type PropertyResolver, type RunFontSelector, type CellContext } from '@docx4j/core-ts';

/**
 * What resolves a colour as Word draws it (the engine's `PropertyResolver`, core-ts 0.6.0 CR-008 phase 2; ED-008
 * section 8.8 item 13): a theme reference wins over the value written beside it, which is a cache Word rewrites
 * and which can be stale. Without one, the written value is drawn.
 */
export type ColorResolver = Pick<PropertyResolver, 'resolveColor' | 'resolveBorderColor' | 'resolveUnderlineColor' | 'resolveFill' | 'resolvePatternColor'>;

/** A run colour as drawn: through the resolver where there is one, else the written six hex digits. */
export function runColorCss(color: wml.Color | undefined, colors?: ColorResolver): string | undefined {
  if (!color) return undefined;
  if (colors) { const c = colors.resolveColor(color); return c ? `#${c.toHex()}` : undefined; }
  return color.val && color.val !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(color.val) ? `#${color.val}` : undefined;
}
import { type EditorDocument, type CellConditions, sectionAt, stateFor, revive, tableConditionsOf, conditionalTcPr, highlightHexValue } from '@docx4j/editor-model';
import { cellDeclarations } from './tables.js';
import { paginationDeclarations } from './print.js';

const PT = 20;   // twips per point

/** `"Calibri", "Carlito", "Liberation Sans", sans-serif`: the name, its substitutes, its class's generic. */
export function fontStack(name: string | null | undefined): string {
  if (!name) return 'sans-serif';
  const row = SUBSTITUTE_ROWS.find((r) => r.documentFont.toLowerCase() === name.toLowerCase());
  const families = [name];
  for (const s of row?.substitutes ?? []) {
    const family = s.font.replace(/\s+(Regular|Light|Display Regular)$/i, '');
    if (!families.some((f) => f.toLowerCase() === family.toLowerCase())) families.push(family);
  }
  const cls = classOf(name);
  const generic = cls === 'SERIF' ? 'serif' : cls === 'MONO' ? 'monospace' : 'sans-serif';
  return families.map((f) => `"${f.replace(/"/g, '')}"`).join(', ') + ', ' + generic;
}

/** The document font a run shows Latin text in, theme references resolved through the selector. */
export function latinFontOf(selector: RunFontSelector | undefined, pPr: wml.PPr | undefined, rPr: wml.RPr | undefined): string {
  if (selector) {
    try {
      const span = selector.spans(pPr, rPr, 'a', { rPrIsEffective: true })[0];
      if (span?.documentFont) return span.documentFont;
    } catch { /* fall through to the direct name */ }
  }
  return readRunOptions(rPr).name;
}

function twipsPt(v: number | undefined): string | undefined {
  return v === undefined ? undefined : `${v / PT}pt`;
}

/** The CSS declarations of paragraph properties (effective or direct). */
export function paragraphDeclarations(pPr: wml.PPr | undefined): string[] {
  if (!pPr) return [];
  const out: string[] = [];
  const jc = pPr.jc?.val as string | undefined;
  if (jc) out.push(`text-align: ${jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'end' : jc === 'both' || jc === 'distribute' ? 'justify' : 'start'}`);
  const sp = pPr.spacing;
  if (sp) {
    if (sp.before !== undefined) out.push(`margin-block-start: ${sp.before / PT}pt`);
    if (sp.after !== undefined) out.push(`margin-block-end: ${sp.after / PT}pt`);
    if (sp.line !== undefined) {
      const rule = sp.lineRule ?? 'auto';
      out.push(rule === 'auto' ? `line-height: ${(sp.line / 240).toFixed(3)}` : `line-height: ${sp.line / PT}pt`);
    }
  }
  const ind = pPr.ind;
  if (ind) {
    const start = ind.start ?? ind.left;
    const end = ind.end ?? ind.right;
    const hanging = ind.hanging;
    // `w:left` is where the paragraph's lines sit; a hanging indent pulls the first line back from
    // it by that much (ECMA-376 17.3.1.12), which is where a list label sits (ED-003 section 5.4).
    if (start !== undefined) out.push(`padding-inline-start: ${start / PT}pt`);
    if (end !== undefined) out.push(`padding-inline-end: ${end / PT}pt`);
    if (hanging !== undefined) out.push(`text-indent: -${hanging / PT}pt`);
    else if (ind.firstLine !== undefined) out.push(`text-indent: ${ind.firstLine / PT}pt`);
  }
  if (pPr.bidi !== undefined && pPr.bidi.val !== false) out.push('direction: rtl');
  return out;
}

function on(v: { val?: boolean } | undefined): boolean {
  return v !== undefined && v.val !== false;
}

/** The CSS declarations of run properties; `font` names the family to use (resolved elsewhere). */
export function runDeclarations(rPr: wml.RPr | undefined, font?: string | null, colors?: ColorResolver): string[] {
  if (!rPr) return [];
  const out: string[] = [];
  if (font) out.push(`font-family: ${fontStack(font)}`);
  if (rPr.sz?.val !== undefined) out.push(`font-size: ${rPr.sz.val / 2}pt`);
  if (rPr.b !== undefined) out.push(`font-weight: ${on(rPr.b) ? 'bold' : 'normal'}`);
  if (rPr.i !== undefined) out.push(`font-style: ${on(rPr.i) ? 'italic' : 'normal'}`);
  const deco: string[] = [];
  if (rPr.u?.val !== undefined && rPr.u.val !== 'none') deco.push('underline');
  if (on(rPr.strike) || on(rPr.dstrike)) deco.push('line-through');
  if (deco.length > 0) out.push(`text-decoration: ${deco.join(' ')}`);
  const color = runColorCss(rPr.color, colors);
  if (color) out.push(`color: ${color}`);
  if (rPr.highlight?.val && rPr.highlight.val !== 'none') out.push(`background-color: ${highlightCss(rPr.highlight.val)}`);
  if (rPr.vertAlign?.val === 'superscript') out.push('vertical-align: super', 'font-size: smaller');
  else if (rPr.vertAlign?.val === 'subscript') out.push('vertical-align: sub', 'font-size: smaller');
  if (on(rPr.caps)) out.push('text-transform: uppercase');
  if (on(rPr.smallCaps)) out.push('font-variant: small-caps');
  if (rPr.spacing?.val !== undefined) out.push(`letter-spacing: ${rPr.spacing.val / PT}pt`);
  if (on(rPr.rtl)) out.push('direction: rtl', 'unicode-bidi: isolate');
  return out;
}

/**
 * The CSS colour of a highlight: a name through the object model's table (`HIGHLIGHT_COLORS`, the one the engine reads
 * and writes `w:highlight` by, so a file's highlight and one the toolbar applies by name draw alike; its values follow
 * the objects release installed: Word's own palette, green `#00FF00`, since objects CR-014 in 0.6.0), or a `#RRGGBB`
 * as it is.
 */
export function highlightCss(value: string): string {
  return highlightHexValue(value) ?? (value.startsWith('#') ? value : `#${value}`);
}

/** A character style's own properties, its `basedOn` chain folded (child over parent), without the document defaults. */
export function characterStyleRPr(resolver: PropertyResolver, styleId: string): wml.RPr | undefined {
  const chain: wml.RPr[] = [];
  const seen = new Set<string>();
  let id: string | undefined = styleId;
  while (id && !seen.has(id)) {
    seen.add(id);
    const style = resolver.getStyle(id);
    if (!style) break;
    if (style.rPr) chain.unshift(style.rPr);
    id = style.basedOn?.val;
  }
  if (chain.length === 0) return undefined;
  return Object.assign({}, ...chain) as wml.RPr;
}

/**
 * Whether a character style's size and its font family are drawn on its text, as the classes `rz-<id>` and `rf-<id>`
 * that the per-block walk puts where the run states no size or family of its own (`plugins/decorations.ts`), rather
 * than on the style's element (ED-005 section 12.133). That element is the outermost of a run's formatting (12.132),
 * so a size of its own made every line through it as tall as the style's font, even where the run's own size, inside
 * it, replaced the style's: a 14pt style under 11pt runs drew 16.1pt lines where Word draws 12.6pt. The element keeps
 * the paragraph's font. A raised or lowered style keeps its size on the element, as `font-size: smaller` there
 * depends on it.
 */
export function styleOnText(rPr: wml.RPr | undefined): { size: boolean; family: boolean } {
  const raised = rPr?.vertAlign?.val === 'superscript' || rPr?.vertAlign?.val === 'subscript';
  return { size: rPr?.sz?.val !== undefined && !raised, family: rPr?.rFonts !== undefined };
}

/** The style ids a document uses: paragraph styles and run styles (the view passes its state's document; `session.doc` rebuilds one from the store). */
export function stylesInUse(session: EditorDocument, doc: PMNode = session.doc, extra: readonly PMNode[] = []): { paragraph: Set<string>; character: Set<string> } {
  const paragraph = new Set<string>();
  const character = new Set<string>();
  // The stories read so far (notes, headers; ED-003 section 6.4) use styles the body may not.
  for (const d of [doc, ...extra]) {
    d.descendants((node) => {
      if (node.type.name === 'paragraph') {
        const id = node.attrs['pStyle'] as string | null;
        if (id) paragraph.add(id);
      }
      for (const m of node.marks) if (m.type.name === 'runStyle' && typeof m.attrs['id'] === 'string') character.add(m.attrs['id']);
      return true;
    });
  }
  return { paragraph, character };
}

/** Page geometry of a section, in points. */
export interface PageGeometry {
  width: number;
  height: number;
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export function pageGeometry(sectPr: wml.SectPr | null | undefined): PageGeometry {
  const w = sectPr?.pgSz?.w ?? 11906;
  const h = sectPr?.pgSz?.h ?? 16838;
  const m = sectPr?.pgMar;
  return { width: w / PT, height: h / PT, top: (m?.top ?? 1440) / PT, right: (m?.right ?? 1440) / PT, bottom: (m?.bottom ?? 1440) / PT, left: (m?.left ?? 1440) / PT };
}

/**
 * The sections whose text width differs from the first section's: the only ones whose
 * paragraphs need a `sect-i` class and a rule (the decorations emit one node decoration per
 * paragraph in such a section, and none otherwise: a class on every paragraph of a long
 * document costs ProseMirror quadratic time per keystroke, ED-002 section 10.5 item 10).
 */
export function sectionsNeedingClass(secs: readonly (wml.SectPr | null | undefined)[]): Set<number> {
  const out = new Set<number>();
  const first = pageGeometry(secs[0]);
  secs.forEach((s, i) => {
    if (i === 0) return;
    const g = pageGeometry(s);
    if (g.width === first.width && g.left === first.left && g.right === first.right) return;
    out.add(i);
  });
  return out;
}

/** The sections in document order: the paragraph-level ones then the final one. */
export function sections(session: EditorDocument, doc: PMNode = session.doc): wml.SectPr[] {
  const out: wml.SectPr[] = [];
  doc.descendants((node) => {
    if (node.type.name !== 'paragraph') return true;
    const s = node.attrs['sectPr'] as unknown;
    if (s) out.push(revive(s) as wml.SectPr);
    return false;
  });
  const final = session.doc.attrs['finalSectPr'] as unknown;
  out.push(final ? (revive(final) as wml.SectPr) : ({} as wml.SectPr));
  return out;
}

/**
 * Where a stylesheet's styles come from when they are not the session's, and the selector its rules go under: the
 * preview's (ED-005 section 12.117), drawn from its clone's resolver, whose styles can be ones the document does not
 * use or an XHTML conversion added, under the preview's container so that its rules win there.
 */
export interface CssSource {
  resolver: PropertyResolver | undefined;
  selector: RunFontSelector | undefined;
  scope: string;
}

/**
 * The whole stylesheet for a document: the page, the defaults, and a rule per style the document uses (the view
 * refreshes it when a style comes into use). With a `source`, the defaults and the styles alone, from its resolver,
 * under its scope: the page is the body's, and the table conditions are drawn by a plugin a read-only story has not.
 */
export function documentCss(session: EditorDocument, doc: PMNode = session.doc, extra: readonly PMNode[] = [], source?: CssSource): string {
  const css = documentRules(session, doc, extra, source);
  return source ? css.replaceAll('.docx-editor', `${source.scope} .docx-editor`) : css;
}

function documentRules(session: EditorDocument, doc: PMNode, extra: readonly PMNode[], source: CssSource | undefined): string {
  const resolver = source ? source.resolver : session.pkg.propertyResolverOrUndefined;
  const selector = source ? source.selector : session.main.runFontSelectorOrUndefined;
  const rules: string[] = [];
  if (!source) pageRules(session, doc, rules);
  if (!resolver) return rules.join('\n');
  const used = stylesInUse(session, doc, extra);
  styleRules(resolver, selector, used, rules);
  if (!source) tableConditionRules(session, resolver, selector, used, [doc, ...extra], rules);
  return rules.join('\n');
}

/** The page: the first section's width and margins on the editor itself, a later section of another width approximated. */
function pageRules(session: EditorDocument, doc: PMNode, rules: string[]): void {
  const secs = sections(session, doc);
  const first = pageGeometry(secs[0]);
  rules.push(`.docx-editor { inline-size: ${first.width}pt; padding-inline: ${first.left}pt ${first.right}pt; padding-block: ${first.top}pt ${first.bottom}pt; }`);
  // A page boundary spans the page, the document's start and end too (their header and footer had sat inside the
  // text column, inset by the margins a second time: ED-005 section 12.98); the header and footer inside sit within the margins.
  rules.push(`.docx-editor .page-boundary, .docx-editor .page-start, .docx-editor .page-end { margin-inline: -${first.left}pt -${first.right}pt; }`);
  rules.push(`.docx-editor .boundary-footer, .docx-editor .boundary-header { padding-inline: ${first.left}pt ${first.right}pt; }`);
  for (const i of sectionsNeedingClass(secs)) {
    // A later section of a different width: approximated per paragraph.
    const g = pageGeometry(secs[i]);
    const delta = (first.width - first.left - first.right) - (g.width - g.left - g.right);
    // On screen only: in print each section has a page of its own (ED-006 section 3.3).
    rules.push(`@media screen { .docx-editor .sect-${i} { margin-inline-end: ${delta}pt; } }`);
  }
}

/** The document defaults and a rule per style in use. */
function styleRules(resolver: PropertyResolver, selector: RunFontSelector | undefined, used: { paragraph: Set<string>; character: Set<string> }, rules: string[]): void {
  // Document defaults.
  const defaultId = resolver.defaultParagraphStyleId;
  const defaultPPr = resolver.getEffectivePPr(defaultId);
  const defaultRPr = defaultId !== undefined ? resolver.getEffectiveRPr(defaultId) : resolver.getEffectiveRPr(undefined, undefined);
  // The run defaults inherit from the editor; the paragraph defaults (the document defaults and
  // the default paragraph style: Word's 10pt after and 1.15 lines in a new document) go on every
  // paragraph, since margins do not inherit. A style's rule, later and more specific, wins.
  rules.push(`.docx-editor { ${runDeclarations(defaultRPr, latinFontOf(selector, defaultPPr, defaultRPr), resolver).join('; ')} }`);
  rules.push(`.docx-editor p { ${['margin-block: 0', ...paragraphDeclarations(defaultPPr)].join('; ')} }`);
  // Word's breaks and keeps, for printing only (ED-006 section 3.4): the defaults', then each style's.
  const print: string[] = [];
  const defaultBreaks = paginationDeclarations(defaultPPr);
  if (defaultBreaks.length > 0) print.push(`.docx-editor p { ${defaultBreaks.join('; ')} }`);
  for (const id of used.paragraph) {
    const pPr = resolver.getEffectivePPr(id);
    const rPr = resolver.getEffectiveRPr(id);
    const decl = [...paragraphDeclarations(pPr), ...runDeclarations(rPr, latinFontOf(selector, pPr, rPr), resolver)];
    if (decl.length > 0) rules.push(`.docx-editor .s-${cssId(id)} { ${decl.join('; ')} }`);
    const breaks = paginationDeclarations(pPr);
    if (breaks.length > 0) print.push(`.docx-editor .s-${cssId(id)} { ${breaks.join('; ')} }`);
  }
  if (print.length > 0) rules.push(`@media print {\n${print.join('\n')}\n}`);
  for (const id of used.character) {
    const rPr = characterStyleRPr(resolver, id);
    const onText = styleOnText(rPr);
    // The style's size and family are drawn on its text, where the run does not state its own (`styleOnText`).
    const own: wml.RPr | undefined = rPr && onText.size ? { ...rPr } : rPr;
    if (own && onText.size) delete own.sz;
    const decl = runDeclarations(own, undefined, resolver);
    if (decl.length > 0) rules.push(`.docx-editor .r-${cssId(id)} { ${decl.join('; ')} }`);
    const family = onText.family ? latinFontOf(selector, undefined, resolver.getEffectiveRPr(rPr)) : undefined;
    if (family) rules.push(`.docx-editor .rf-${cssId(id)} { font-family: ${fontStack(family)} }`);
    if (onText.size && rPr?.sz?.val !== undefined) rules.push(`.docx-editor .rz-${cssId(id)} { font-size: ${rPr.sz.val / 2}pt }`);
  }
}

/** A table style's conditional formatting, per class of conditions in use. */
function tableConditionRules(session: EditorDocument, resolver: PropertyResolver, selector: RunFontSelector | undefined, used: { paragraph: Set<string> }, docs: readonly PMNode[], rules: string[]): void {
  // A table style's conditional formatting (ED-005 section 12.65): per class of conditions a cell is under
  // (plugins/tableconditions.ts), the cell's properties, and each paragraph style's resolved with the table
  // level, Word's order (core-ts CR-007). A paragraph in a block control in the cell is reached too; a nested
  // table's cells have their own class.
  const { cells, paragraphs } = tableConditionsIn(session, docs);
  /** The text of each paragraph style (and none) under a table context, after the selector for its paragraphs. */
  const text = (context: CellContext, selectorOf: (p: string) => string): void => {
    for (const id of [undefined, ...used.paragraph]) {
      const pPr: wml.PPr = id !== undefined ? { pStyle: { val: id } } : {};
      const effectivePPr = resolver.getEffectivePPr(pPr, context);
      const effectiveRPr = resolver.getEffectiveRPr(undefined, pPr, context);
      const decl = [...paragraphDeclarations(effectivePPr), ...runDeclarations(effectiveRPr, latinFontOf(selector, effectivePPr, effectiveRPr), resolver)];
      if (decl.length > 0) rules.push(`${selectorOf(id !== undefined ? `.s-${cssId(id)}` : ':not([class^="s-"]):not([class*=" s-"])')} { ${decl.join('; ')} }`);
    }
  };
  for (const [key, c] of cells) {
    const cell = cellDeclarations(conditionalTcPr(c), resolver);
    if (cell.length > 0) rules.push(`.docx-editor td.${key} { ${cell.join('; ')} }`);
    if (c.cell.formatsText) text(c.cell, (p) => `.docx-editor td.${key} > p${p}, .docx-editor td.${key} > .sdt-block > p${p}`);
  }
  // After the cells' rules, so that a paragraph's own conditions win over its cell's.
  for (const [key, context] of paragraphs) text(context, (p) => `.docx-editor td > p.${key}${p}, .docx-editor td > .sdt-block > p.${key}${p}`);
}

/** The classes of conditions the documents' tables use, each with a cell or paragraph under it (`tableConditionsOf`), for the stylesheet. */
export function tableConditionsIn(session: EditorDocument, docs: readonly PMNode[]): { cells: Map<string, CellConditions>; paragraphs: Map<string, CellContext> } {
  const cells = new Map<string, CellConditions>();
  const paragraphs = new Map<string, CellContext>();
  for (const doc of docs) {
    doc.descendants((n) => {
      if (n.type.name === 'table') {
        const found = tableConditionsOf(session, n);
        for (const c of found.cells.values()) if (!cells.has(c.key)) cells.set(c.key, c);
        for (const p of found.paragraphs.values()) if (!paragraphs.has(p.key)) paragraphs.set(p.key, p.context);
      }
      return n.isBlock;
    });
  }
  return { cells, paragraphs };
}

/**
 * A paragraph style's look for its entry in the toolbar's style list, as Word's style gallery and pane
 * show a style: its run properties (font, size, bold, italic, colour, decoration), as React style
 * properties, the size kept between 9 and 16 points so a title does not swamp the list. Nothing
 * without the document's styles.
 */
export function stylePreview(session: EditorDocument, id: string): Record<string, string> {
  const resolver = session.pkg.propertyResolverOrUndefined;
  if (!resolver) return {};
  const pPr = resolver.getEffectivePPr(id);
  const rPr = resolver.getEffectiveRPr(id);
  const out: Record<string, string> = {};
  for (const declaration of runDeclarations(rPr, latinFontOf(session.main.runFontSelectorOrUndefined, pPr, rPr), session.pkg.propertyResolverOrUndefined)) {
    const colon = declaration.indexOf(':');
    if (colon < 0) continue;
    const property = declaration.slice(0, colon).trim();
    if (property === 'direction' || property === 'unicode-bidi' || property === 'vertical-align') continue;
    let value = declaration.slice(colon + 1).trim();
    if (property === 'font-size') {
      const points = Number.parseFloat(value);
      if (!Number.isFinite(points)) continue;
      value = `${Math.min(16, Math.max(9, points))}pt`;
    }
    out[property.replace(/-([a-z])/g, (_m, c: string) => c.toUpperCase())] = value;
  }
  return out;
}

/** A style id as a CSS class fragment. */
export function cssId(id: string): string {
  return id.replace(/[^A-Za-z0-9_-]/g, (c) => `_${c.charCodeAt(0).toString(16)}`);
}

/**
 * The inline style of a paragraph node's direct formatting; with `indent`, a list paragraph's
 * layout indent (the level's with the paragraph's own over it, `levelIndentOf`) in place of the
 * attribute.
 */
export function paragraphInlineStyle(attrs: Record<string, unknown>, indent?: { start: number | null; end: number | null; firstLine: number | null; hanging: number | null } | null, colors?: ColorResolver): string {
  const pPr: wml.PPr = {};
  if (attrs['jc']) pPr.jc = { val: attrs['jc'] as wml.Jc['val'] };
  const ind = indent ?? (attrs['ind'] as { start: number | null; end: number | null; firstLine: number | null; hanging: number | null } | null);
  if (ind) {
    pPr.ind = {};
    if (ind.start !== null) pPr.ind.start = ind.start;
    if (ind.end !== null) pPr.ind.end = ind.end;
    if (ind.firstLine !== null) pPr.ind.firstLine = ind.firstLine;
    if (ind.hanging !== null) pPr.ind.hanging = ind.hanging;
  }
  const sp = attrs['spacing'] as { before: number | null; after: number | null; line: number | null; lineRule: string | null } | null;
  if (sp) {
    pPr.spacing = {};
    if (sp.before !== null) pPr.spacing.before = sp.before;
    if (sp.after !== null) pPr.spacing.after = sp.after;
    if (sp.line !== null) pPr.spacing.line = sp.line;
    if (sp.lineRule !== null) pPr.spacing.lineRule = sp.lineRule as wml.PPrBase.Spacing['lineRule'];
  }
  if (attrs['bidi'] === true) pPr.bidi = {};
  const out = paragraphDeclarations(pPr);
  // The paragraph mark's own size and colour, for its pilcrow when the formatting marks are shown (ED-005 section
  // 12.80): Word draws the pilcrow in the mark's run formatting; what the mark inherits the pilcrow inherits with it.
  const mark = attrs['markRPr'] as { sz?: { val?: number }; color?: wml.Color } | null;
  if (mark?.sz?.val !== undefined) out.push(`--mark-size: ${mark.sz.val / 2}pt`);
  const markColor = runColorCss(mark?.color, colors);
  if (markColor) out.push(`--mark-color: ${markColor}`);
  return out.join('; ');
}

/** The section index (for `sect-<n>`) of a position, and whether the paragraph ends a section. */
export function sectionIndexAt(session: EditorDocument, pos: number): number {
  const state = stateFor(session, { from: pos });
  let n = 0;
  let found = -1;
  state.doc.descendants((node, nodePos) => {
    if (found >= 0) return false;
    if (node.type.name !== 'paragraph') return true;
    if (nodePos + node.nodeSize > pos) { found = n; return false; }
    if (node.attrs['sectPr']) n++;
    return false;
  });
  void sectionAt;
  return found < 0 ? n : found;
}
