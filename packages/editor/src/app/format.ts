// SPDX-License-Identifier: GPL-3.0-or-later
//
// What the caret sits in, resolved (the Format and Styles tabs, Jason 2026-09-24, after Word's
// Reveal Formatting): the run's, the paragraph's and the section's effective properties from
// the same resolver the stylesheet is generated from, so the tab and the page agree; and the
// style chains the resolver climbs for the run, the paragraph, the list and the table. The
// properties are shown as the XML they marshal to, element by element with its attributes,
// which is the one rendering that cannot drift from the object model.
import type { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import * as el from '@docx4j/generated-objects-ts/modules/org_docx4j_wml.el';
import { marshalString, unmarshalString, typeNameOf, segmentsOf, type Element } from '@docx4j/core-ts';
import { type EditorDocument, paragraphPropertiesFor, runPropertiesFor, sectionAt, atomMarks, isFormattableAtom, revive, ATOM_TYPES, tableAround } from '@docx4j/editor-model';
import type { EditorSurface } from '../view/surface.js';
import { xmlSyntaxError, XmlSyntaxError, pretty, blockAt, runUnderCaret } from './xml.js';

/** The direct properties at the caret, before resolution. */
export interface CaretContext {
  paragraph: PMNode | undefined;
  /** The paragraph's direct `w:pPr`, with its `w:pStyle`. */
  pPr: wml.PPr | undefined;
  /** The run's direct `w:rPr`, with its `w:rStyle`. */
  rPr: wml.RPr | undefined;
  sectPr: wml.SectPr | undefined;
  /** Whether the caret sits in a table. */
  inTable: boolean;
  /** The table's direct `w:tblPr`, with its `w:tblStyle`, when the caret is in one. */
  tblPr: wml.CTTblPrBase | undefined;
}

export function caretContext(session: EditorDocument, state: EditorState): CaretContext {
  const $from = state.selection.$from;
  let paragraph: PMNode | undefined;
  for (let d = $from.depth; d >= 0; d--) { const n = $from.node(d); if (n.type.name === 'paragraph') { paragraph = n; break; } }
  const pPr = paragraph ? paragraphPropertiesFor(paragraph, state) : undefined;
  let marks = state.storedMarks ?? $from.marks();
  const after = $from.nodeAfter;
  if (state.selection.empty && after && ATOM_TYPES.has(after.type.name) && isFormattableAtom(after)) {
    const runSource = after.attrs['runSource'] as string | null;
    const atom = atomMarks(after, state);
    marks = runSource ? [...atom, state.schema.marks['runSource']!.create({ sourceId: runSource })] : atom;
  }
  const runSource = marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'] as string | undefined;
  const source = runSource ? (session.ctx.ids.get(runSource) as Element<wml.R> | undefined) : undefined;
  const rPr = runPropertiesFor(source?.value.rPr, marks);
  const section = sectionAt(state);
  const sectPr = section.sectPr ? (revive(section.sectPr) as wml.SectPr) : undefined;
  const table = tableAround(state);
  const tblPr = table?.node.attrs['tblPr'] ? (revive(table.node.attrs['tblPr']) as wml.CTTblPrBase) : undefined;
  return { paragraph, pPr, rPr, sectPr, inTable: table !== undefined, tblPr };
}

export interface Effective {
  rPr: wml.RPr | undefined;
  pPr: wml.PPr | undefined;
  sectPr: wml.SectPr | undefined;
}

/** The effective properties at the caret, from the package's resolver. */
export function effectiveAt(session: EditorDocument, state: EditorState): Effective {
  const ctx = caretContext(session, state);
  const resolver = session.pkg.propertyResolverOrUndefined;
  if (!resolver) return { rPr: ctx.rPr, pPr: ctx.pPr, sectPr: ctx.sectPr };
  return {
    rPr: resolver.getEffectiveRPr(ctx.rPr, ctx.pPr),
    pPr: ctx.paragraph ? resolver.getEffectivePPr(ctx.pPr) : undefined,
    sectPr: ctx.sectPr,
  };
}

/** One element of a properties tree as shown: `w:sz` with `w:val="28"`, and its children. */
export interface PropertyRow {
  name: string;
  attributes: { name: string; value: string }[];
  children: PropertyRow[];
}

function rowsOf(node: globalThis.Element): PropertyRow {
  const attributes: { name: string; value: string }[] = [];
  for (const a of Array.from(node.attributes)) if (!a.name.startsWith('xmlns')) attributes.push({ name: a.name, value: a.value });
  const children: PropertyRow[] = [];
  for (const c of Array.from(node.children)) children.push(rowsOf(c));
  return { name: node.tagName, attributes, children };
}

/**
 * A properties object as rows: marshalled through the object model (wrapped in the element
 * that owns it, since `w:pPr` and `w:sectPr` have no factory of their own), then read back
 * with the browser's parser. Empty when the object has nothing to show.
 */
export async function propertyRows(kind: 'rPr' | 'pPr' | 'sectPr', value: wml.RPr | wml.PPr | wml.SectPr | undefined): Promise<PropertyRow[]> {
  if (!value || typeof DOMParser === 'undefined') return [];
  const owner = kind === 'rPr' ? el.rPr(value as wml.RPr) : kind === 'pPr' ? el.p({ pPr: value as wml.PPr }) : el.body({ sectPr: value as wml.SectPr });
  const xml = await marshalString(owner as Element);
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const root = doc.documentElement;
  const target = kind === 'rPr' ? root : Array.from(root.children).find((c) => c.localName === kind);
  if (!target) return [];
  return rowsOf(target).children;
}

export interface StyleChains {
  /**
   * The run's character style and what it is based on, the run's own first. A run with no
   * `w:rStyle` carries the default character style (Word's "Default Paragraph Font"), which is
   * listed then as the default: a character style is always in the stack.
   */
  character: wml.Style[];
  /** The numbering style the paragraph's list links to (`w:numStyleLink`), and its chain. */
  numbering: wml.Style[];
  /** The paragraph's style, or the default paragraph style, and what it is based on. */
  paragraph: wml.Style[];
  /** The table's style and what it is based on, when the caret is in a table (E2.a). */
  table: wml.Style[];
  /** Which ids the caret named directly, so the tab can say "applied" against "default". */
  applied: { character: string | undefined; paragraph: string | undefined; numId: string | undefined };
  /** The top of the stack: the run's own `w:rPr` and the paragraph's own `w:pPr`. */
  direct: { rPr: wml.RPr | undefined; pPr: wml.PPr | undefined };
  /** The bottom of the stack: `w:docDefaults` of the styles part. */
  defaults: { rPr: wml.RPr | undefined; pPr: wml.PPr | undefined; present: boolean };
}

/**
 * The styles part's tree: the live one once something unmarshalled the part, else a private
 * read cached per session (the resolver's own copy is not reachable, and a read parses the
 * whole part, which the caret's every move must not repeat).
 */
const privateStyles = new WeakMap<EditorDocument, Promise<wml.Styles | undefined>>();
async function stylesOf(session: EditorDocument): Promise<wml.Styles | undefined> {
  const part = session.main.styleDefinitionsPart;
  if (!part) return undefined;
  if (part.isUnmarshalled) return part.contents;
  let cached = privateStyles.get(session);
  if (!cached) {
    cached = part.readContents().catch(() => undefined);
    privateStyles.set(session, cached);
  }
  return cached;
}

/** The style chains at the caret, leaf first (the resolver's `ancestry` is root first). */
export async function styleChains(session: EditorDocument, state: EditorState): Promise<StyleChains> {
  const resolver = session.pkg.propertyResolverOrUndefined;
  const ctx = caretContext(session, state);
  const direct = { rPr: ctx.rPr, pPr: ctx.pPr };
  const styles = await stylesOf(session);
  const defaults = { rPr: styles?.docDefaults?.rPrDefault?.rPr, pPr: styles?.docDefaults?.pPrDefault?.pPr, present: styles?.docDefaults !== undefined };
  const empty: StyleChains = { character: [], numbering: [], paragraph: [], table: [], applied: { character: undefined, paragraph: undefined, numId: undefined }, direct, defaults };
  if (!resolver) return empty;
  const leafFirst = (id: string | undefined): wml.Style[] => (id ? resolver.ancestry(id).slice().reverse() : []);
  // The default character style, as the resolver finds it: the first `w:default="1"` character style.
  const defaultCharacter = styles?.style?.find((st) => st.type === 'character' && st._default === true && st.styleId)?.styleId;
  const character = ctx.rPr?.rStyle?.val ?? defaultCharacter;
  const paragraph = ctx.paragraph ? ctx.pPr?.pStyle?.val ?? resolver.getDefaultParagraphStyleId() : undefined;
  // The list: the paragraph's own numbering, else the one its style chain carries.
  const effectivePPr = ctx.paragraph ? resolver.getEffectivePPr(ctx.pPr) : undefined;
  const numId = effectivePPr?.numPr?.numId?.val !== undefined ? String(effectivePPr.numPr.numId.val) : undefined;
  let numbering: wml.Style[] = [];
  if (numId && numId !== '0') {
    try {
      const definitions = await session.main.numberingDefinitionsPart?.getDefinitions();
      const linked = definitions?.instanceListDefinitions.get(numId)?.abstractDefinition?.linkedStyleId;
      numbering = leafFirst(linked);
    } catch { numbering = []; }
  }
  return {
    character: leafFirst(character),
    numbering,
    paragraph: leafFirst(paragraph),
    table: ctx.inTable ? leafFirst(ctx.tblPr?.tblStyle?.val ?? styles?.style?.find((st) => st.type === 'table' && st._default === true && st.styleId)?.styleId) : [],
    applied: { character: ctx.rPr?.rStyle?.val, paragraph: ctx.pPr?.pStyle?.val, numId },
    direct,
    defaults,
  };
}

/** A layer of the stack other than a style: direct formatting, or the document defaults. */
export type LayerKind = 'rPr' | 'pPr' | 'rPrDefault' | 'pPrDefault';

const LAYER_ROOT: Record<LayerKind, string> = { rPr: 'w:rPr', pPr: 'w:pPr', rPrDefault: 'w:rPrDefault', pPrDefault: 'w:pPrDefault' };

/**
 * A layer's XML: marshalled inside the element that owns it (`w:r`, `w:p`, `w:docDefaults`),
 * the wanted child serialised back out. An absent layer is an empty element to start from.
 */
export async function layerXml(kind: LayerKind, value: wml.RPr | wml.PPr | undefined): Promise<string> {
  if (!value) return `<${LAYER_ROOT[kind]}/>`;
  const owner = kind === 'rPr' ? el.r({ rPr: value as wml.RPr })
    : kind === 'pPr' ? el.p({ pPr: value as wml.PPr })
    : kind === 'rPrDefault' ? el.docDefaults({ rPrDefault: { rPr: value as wml.RPr } })
    : el.docDefaults({ pPrDefault: { pPr: value as wml.PPr } });
  const xml = await marshalString(owner as Element);
  if (typeof DOMParser === 'undefined') return pretty(xml);
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const child = Array.from(doc.documentElement.children).find((c) => `${c.prefix ?? 'w'}:${c.localName}` === LAYER_ROOT[kind]);
  return pretty(child ? new XMLSerializer().serializeToString(child) : xml);
}

/** Parse an edited layer back, through the element that owns it. */
export async function parseLayer(kind: LayerKind, xml: string): Promise<wml.RPr | wml.PPr | undefined> {
  const syntax = xmlSyntaxError(xml, 'block');
  if (syntax) throw new XmlSyntaxError(syntax.message, syntax.line);
  const text = xml.trim().replace(/^<\?xml[^>]*\?>\s*/, '');
  if (kind === 'rPr') {
    const parsed = await unmarshalString(`<w:r ${STYLES_DECLARATIONS}>${text}</w:r>`);
    return (parsed.value as wml.R).rPr;
  }
  if (kind === 'pPr') {
    const parsed = await unmarshalString(`<w:p ${STYLES_DECLARATIONS}>${text}</w:p>`);
    return (parsed.value as wml.P).pPr;
  }
  const parsed = await unmarshalString(`<w:styles ${STYLES_DECLARATIONS}><w:docDefaults>${text}</w:docDefaults></w:styles>`);
  const defaults = (parsed.value as wml.Styles).docDefaults;
  return kind === 'rPrDefault' ? defaults?.rPrDefault?.rPr : defaults?.pPrDefault?.pPr;
}

/** Replace the document defaults' run or paragraph properties in the styles part, in place, as `replaceStyle` does. */
export async function replaceDocDefaults(session: EditorDocument, editor: EditorSurface, kind: 'rPrDefault' | 'pPrDefault', value: wml.RPr | wml.PPr | undefined): Promise<void> {
  const part = session.main.styleDefinitionsPart;
  if (!part) throw new Error('the document has no styles part');
  const styles = await part.readContents();
  const defaults = (styles.docDefaults ??= {});
  if (kind === 'rPrDefault') defaults.rPrDefault = { rPr: value as wml.RPr | undefined };
  else defaults.pPrDefault = { pPr: value as wml.PPr | undefined };
  part.setContents(styles);
  session.pkg.refreshPropertyResolver();
  session.markPartChanged(part.partName.toString());
  editor.refreshCss();
}

/**
 * Replace the run's own `w:rPr` or the paragraph's own `w:pPr` at the caret, through the tree
 * (decision 20), as reveal codes replaces a run: the edit is one undo step.
 */
export async function replaceDirect(session: EditorDocument, editor: EditorSurface, kind: 'rPr' | 'pPr', value: wml.RPr | wml.PPr | undefined): Promise<void> {
  if (editor.story) throw new Error('direct formatting in a note or a header is edited in place, not through the tree');
  const state = editor.state;
  const block = blockAt(state);
  if (!block) throw new Error('no block at the caret');
  const runInfo = kind === 'rPr' ? runUnderCaret(session, state, session.exportBlock(block.node)) : undefined;
  if (kind === 'rPr' && !runInfo) throw new Error('no run under the caret');
  const blockIndex = block.index;
  const result = await session.runOnTree(({ body }) => {
    const paragraph = body.content[blockIndex];
    if (!paragraph || typeNameOf(paragraph) !== 'org_docx4j_wml.P') throw new Error('editing inside a content control is not supported yet; edit the block in reveal codes');
    const p = paragraph.value as wml.P;
    if (kind === 'pPr') { if (value) p.pPr = value as wml.PPr; else delete p.pPr; return; }
    let runIndex = -1;
    let lastRun: object | undefined;
    for (const seg of segmentsOf(p as object)) {
      if (seg.run !== lastRun) { runIndex++; lastRun = seg.run; }
      if (runIndex === runInfo!.index) {
        const r = seg.runOwner[seg.runIndex]!.value as wml.R;
        if (value) r.rPr = value as wml.RPr; else delete r.rPr;
        return;
      }
    }
    throw new Error('the run is no longer where it was');
  });
  if (result.error) throw result.error instanceof Error ? result.error : new Error(String(result.error));
}

export async function styleXml(style: wml.Style): Promise<string> {
  // Shown, never parsed: a character XML 1.0 excludes is shown as held, not refused (objects 0.8.0, CR-020).
  return pretty(await marshalString(el.style(style) as Element, { checkCharacters: false }));
}

const STYLES_DECLARATIONS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';

/** Parse an edited `w:style`; the browser parser's line on a syntax error, the unmarshaller's message otherwise. */
export async function parseStyle(xml: string): Promise<wml.Style> {
  const syntax = xmlSyntaxError(xml, 'block');
  if (syntax) throw new XmlSyntaxError(syntax.message, syntax.line);
  const text = xml.trim().replace(/^<\?xml[^>]*\?>\s*/, '');
  const parsed = await unmarshalString(`<w:styles ${STYLES_DECLARATIONS}>${text}</w:styles>`);
  const styles = parsed.value as wml.Styles;
  const list = styles.style ?? [];
  if (list.length !== 1) throw new Error(`expected one w:style, found ${list.length}`);
  if (typeNameOf(parsed) === undefined) throw new Error('not a styles element');
  const style = list[0]!;
  if (!style.styleId) throw new Error('the style has no w:styleId');
  return style;
}

/**
 * Replace a style in the styles part with an edited one, in place: the part becomes live (it
 * is re-marshalled on save, so its round trip is spent), the resolver and the stylesheet are
 * refreshed, and the session records the part as touched. No undo in E1.
 */
export async function replaceStyle(session: EditorDocument, editor: EditorSurface, style: wml.Style): Promise<void> {
  const part = session.main.styleDefinitionsPart;
  if (!part) throw new Error('the document has no styles part');
  const styles = await part.readContents();
  const list = (styles.style ??= []);
  const at = list.findIndex((s) => s.styleId === style.styleId);
  if (at < 0) throw new Error(`no style with id ${style.styleId}`);
  list[at] = style;
  part.setContents(styles);
  session.pkg.refreshPropertyResolver();
  session.markPartChanged(part.partName.toString());
  editor.refreshCss();
}
