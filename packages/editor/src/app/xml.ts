// SPDX-License-Identifier: GPL-3.0-or-later
//
// XML helpers for reveal codes (ED-002 section 6.2): pretty-printing, the run under the caret,
// and parsing an edited fragment back into an element through the engine's unmarshaller (the
// schema check gives the error on a bad edit).
import { marshalString, unmarshalString, segmentsOf, typeNameOf, type Element } from '@docx4j/core-ts';
import type { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorDocument } from '@docx4j/editor-model';
import { TableMap } from 'prosemirror-tables';

/** A light indentation of marshalled XML, for reading; the text is not changed. */
export function pretty(xml: string): string {
  let depth = 0;
  return xml
    .replace(/^<\?xml[^>]*\?>/, '')
    .replace(/></g, '>\n<')
    .split('\n')
    .filter((l) => l !== '')
    .map((line) => {
      if (/^<\//.test(line)) depth = Math.max(0, depth - 1);
      const out = '  '.repeat(depth) + line;
      if (/^<[^/!?][^>]*[^/]>$/.test(line) && !/<\/[^>]+>$/.test(line)) depth++;
      return out;
    })
    .join('\n');
}

const XML_NS = 'http://www.w3.org/XML/1998/namespace';

/** A DOM element (`Element` here is the engine's type). */
type DomElement = Document['documentElement'];

/**
 * XML indented for reading and editing, at the person's request (the parts panel's Format, Monaco's
 * Format Document): a line break and indentation before each child of an element that holds only
 * elements, the whitespace that was there replaced. An element holding text (a run's `w:t`, a data
 * node's value, mixed content), one with no children, and everything under `xml:space="preserve"`
 * are left as they are, so no text changes; whitespace alone between elements is not kept. The XML
 * declaration stays. Throws, with the parser's message, when the text does not parse.
 */
export function formatXml(text: string, indent = '  '): string {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  const error = doc.getElementsByTagName('parsererror')[0];
  if (error) throw new Error((error.textContent ?? '').trim() || 'not well-formed');
  const isText = (n: Node): boolean => n.nodeType === Node.TEXT_NODE || n.nodeType === Node.CDATA_SECTION_NODE;
  const walk = (el: DomElement, depth: number): void => {
    if (el.getAttributeNS(XML_NS, 'space') === 'preserve') return;
    const kids = Array.from(el.childNodes);
    if (kids.some((n) => isText(n) && (n.nodeValue ?? '').trim() !== '')) return;
    const others = kids.filter((n) => !isText(n));
    if (others.length === 0) return;
    for (const n of kids) if (isText(n)) el.removeChild(n);
    for (const n of others) {
      el.insertBefore(doc.createTextNode(`\n${indent.repeat(depth + 1)}`), n);
      if (n.nodeType === Node.ELEMENT_NODE) walk(n as DomElement, depth + 1);
    }
    el.appendChild(doc.createTextNode(`\n${indent.repeat(depth)}`));
  };
  walk(doc.documentElement, 0);
  const body = new XMLSerializer().serializeToString(doc);
  const declaration = /^\s*(<\?xml\s[^?]*\?>)/.exec(text)?.[1];
  return declaration && !body.startsWith('<?xml') ? `${declaration}\n${body}` : body;
}

/** A start tag's attributes, each with its name and the span it takes in the tag. */
const ATTRIBUTE = /([^\s=/>]+)\s*=\s*("[^"]*"|'[^']*')/g;

/** The markup of an XML text: comments, CDATA, processing instructions and declarations, start tags (quoted values may hold `>`), end tags. */
const MARKUP = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<![^>]*>|<\/[^>]*>|<[^\s/>!?][^\s/>]*(?:\s+[^\s=/>]+\s*=\s*(?:"[^"]*"|'[^']*'))*\s*\/?>/g;

/**
 * Where an offset in an XML text is, as the element it is in (or whose tag it is on): its position among its parent's
 * element children at each level from the root, one-based, so `/*[1]/*[2]` selects it in the part the text is; and the
 * attribute whose name or value it is on, if any (the Data tab's XML view, ED-005 section 12.64). Undefined outside the
 * root element. The text is assumed to parse.
 */
export function elementPathAt(text: string, offset: number): { path: number[]; attribute?: string } | undefined {
  const stack: { index: number; children: number }[] = [];
  let roots = 0;
  const current = (): { path: number[] } | undefined => (stack.length > 0 ? { path: stack.map((s) => s.index) } : undefined);
  MARKUP.lastIndex = 0;
  for (let m = MARKUP.exec(text); m; m = MARKUP.exec(text)) {
    const tag = m[0];
    const start = m.index;
    const end = start + tag.length;
    // In the text before this markup: the element the text is in.
    if (offset < start) return current();
    if (tag.startsWith('</')) {
      if (offset < end) return current();
      stack.pop();
      continue;
    }
    if (tag.startsWith('<!') || tag.startsWith('<?')) {
      if (offset < end) return current();
      continue;
    }
    const parent = stack[stack.length - 1];
    const index = parent ? ++parent.children : ++roots;
    const path = [...stack.map((s) => s.index), index];
    if (offset < end) {
      ATTRIBUTE.lastIndex = 0;
      for (let a = ATTRIBUTE.exec(tag); a; a = ATTRIBUTE.exec(tag)) {
        if (offset >= start + a.index && offset < start + a.index + a[0].length) return { path, attribute: a[1]! };
      }
      return { path };
    }
    if (!tag.endsWith('/>')) stack.push({ index, children: 0 });
  }
  return current();
}

/** An element of an XML text as the outline reads it: its name, its attributes, the lines of its start and end tags, its element children. */
export interface XmlOutline {
  name: string;
  /** The name after its prefix. */
  local: string;
  attrs: Record<string, string>;
  /** The line of the start tag, one-based. */
  line: number;
  /** The line of the end tag; the start tag's own for an empty element or one closed on its line. */
  endLine: number;
  children: XmlOutline[];
}

/**
 * The elements of an XML text with the lines their tags are on (reveal codes in a table, the folding of every XML
 * editor): the text is assumed to parse, as `elementPathAt` assumes. Comments, CDATA sections, processing instructions
 * and declarations are passed over.
 */
export function outlineOf(text: string): XmlOutline[] {
  const roots: XmlOutline[] = [];
  const stack: XmlOutline[] = [];
  let line = 1;
  let scanned = 0;
  // The line an offset is on; the offsets come in order, so the text is read once.
  const lineAt = (offset: number): number => {
    for (; scanned < offset; scanned++) if (text.charCodeAt(scanned) === 10) line++;
    return line;
  };
  MARKUP.lastIndex = 0;
  for (let m = MARKUP.exec(text); m; m = MARKUP.exec(text)) {
    const tag = m[0];
    if (tag.startsWith('</')) {
      const el = stack.pop();
      if (el) el.endLine = lineAt(m.index);
      continue;
    }
    if (tag.startsWith('<!') || tag.startsWith('<?')) continue;
    const name = /^<([^\s/>]+)/.exec(tag)![1]!;
    const attrs: Record<string, string> = {};
    ATTRIBUTE.lastIndex = 0;
    for (let a = ATTRIBUTE.exec(tag); a; a = ATTRIBUTE.exec(tag)) attrs[a[1]!] = a[2]!.slice(1, -1);
    const at = lineAt(m.index);
    const el: XmlOutline = { name, local: name.slice(name.indexOf(':') + 1), attrs, line: at, endLine: at, children: [] };
    (stack[stack.length - 1]?.children ?? roots).push(el);
    if (!tag.endsWith('/>')) stack.push(el);
  }
  return roots;
}

/** The regions an XML editor folds: every element whose end tag is on a later line than its start tag, start line to end line. */
export function foldingRangesOf(text: string): { start: number; end: number }[] {
  const out: { start: number; end: number }[] = [];
  const walk = (els: XmlOutline[]): void => {
    for (const el of els) {
      if (el.endLine > el.line) out.push({ start: el.line, end: el.endLine });
      walk(el.children);
    }
  };
  walk(outlineOf(text));
  return out;
}

/** One table on the way down to the caret: which of its parent's tables, which row, and the grid column the caret's cell starts at. */
export interface TableStep {
  table: number;
  row: number;
  /** Absent when the selection is of the row (a cell selection's anchor sits before its cell). */
  column?: number;
}

/**
 * The tables the caret is in, outermost first, each as its index among its parent's tables (0 for the block itself), the caret's row and the grid
 * column of the caret's cell (through `TableMap`, since a vertical merge's continuation is no node of its row, and a
 * merged cell spans several columns). Empty outside a table.
 */
export function tableChainOf(state: EditorState): TableStep[] {
  const $from = state.selection.$from;
  const chain: TableStep[] = [];
  for (let d = 1; d <= $from.depth; d++) {
    const node = $from.node(d);
    if (node.type.spec['tableRole'] !== 'table') continue;
    if (d + 1 > $from.depth) break;
    // The block's XML is the table's own when the table is the block; inside a control or a cell, which of its tables.
    const parent = $from.node(d - 1);
    let table = 0;
    if (d > 1) for (let i = 0; i < $from.index(d - 1); i++) if (parent.child(i).type.spec['tableRole'] === 'table') table++;
    const row = $from.index(d);
    const cellPos = d + 2 <= $from.depth ? $from.before(d + 2) : $from.nodeAfter?.type.spec['tableRole'] === 'cell' ? $from.pos : undefined;
    const step: TableStep = { table, row };
    if (cellPos !== undefined) step.column = TableMap.get(node).findCell(cellPos - $from.start(d)).left;
    chain.push(step);
    d += 2;
  }
  return chain;
}

/** Where a table's XML the caret is: the lines of the caret's cell (or row), and the start lines of the rows and cells around it to fold. */
export interface XmlFocus {
  line: number;
  endLine: number;
  fold: number[];
}

const valueOf = (el: XmlOutline | undefined): number => {
  if (!el) return 1;
  const key = Object.keys(el.attrs).find((k) => k === 'w:val' || k.endsWith(':val') || k === 'val');
  const n = key ? Number(el.attrs[key]) : NaN;
  return Number.isFinite(n) && n > 0 ? n : 1;
};
const childNamed = (el: XmlOutline, local: string): XmlOutline | undefined => el.children.find((c) => c.local === local);
/** The elements named `local` under `els`, in order, not looking inside the elements named in `stop` (a nested table's rows are its own). */
const collect = (els: XmlOutline[], local: string, stop: string[]): XmlOutline[] => {
  const out: XmlOutline[] = [];
  const walk = (list: XmlOutline[]): void => {
    for (const el of list) {
      if (el.local === local) out.push(el);
      else if (!stop.includes(el.local)) walk(el.children);
    }
  };
  walk(els);
  return out;
};

/**
 * The caret's cell in the marshalled XML of its block (reveal codes in a table, 2026-10-15): the rows other than the
 * caret's and, in its row, the cells other than its own are folded, and the cell's lines are marked. A row wrapped in
 * a control is found through it; a cell's grid column is its `w:gridSpan`s summed from the row's `w:gridBefore`; a
 * nested table is the step's table among the cell's. What the XML cannot give (a row or cell not found) ends the
 * chain with what was found.
 */
export function cellFocusOf(outline: XmlOutline[], chain: TableStep[]): XmlFocus | undefined {
  let scope = outline;
  let focus: XmlOutline | undefined;
  const fold: number[] = [];
  for (const step of chain) {
    const table = collect(scope, 'tbl', [])[step.table];
    if (!table) break;
    const rows = collect(table.children, 'tr', ['tbl', 'tc']);
    const row = rows[step.row];
    if (!row) break;
    for (const r of rows) if (r !== row && r.endLine > r.line) fold.push(r.line);
    focus = row;
    if (step.column === undefined) break;
    const cells = collect(row.children, 'tc', ['tbl', 'tr', 'tc']);
    const trPr = childNamed(row, 'trPr');
    let at = trPr && childNamed(trPr, 'gridBefore') ? valueOf(childNamed(trPr, 'gridBefore')) : 0;
    let cell: XmlOutline | undefined;
    for (const c of cells) {
      if (at === step.column) cell = c;
      const tcPr = childNamed(c, 'tcPr');
      at += tcPr ? valueOf(childNamed(tcPr, 'gridSpan')) : 1;
    }
    if (!cell) break;
    for (const c of cells) if (c !== cell && c.endLine > c.line) fold.push(c.line);
    focus = cell;
    scope = cell.children;
  }
  return focus ? { line: focus.line, endLine: focus.endLine, fold } : undefined;
}

/** The chain of nodes from the document down to the caret's block and inline container. */
export interface Crumb {
  label: string;
  node: PMNode;
  pos: number;
  depth: number;
}

function labelOf(node: PMNode): string {
  switch (node.type.name) {
    case 'paragraph': return 'p';
    case 'table': return 'tbl';
    case 'table_row': return 'tr';
    case 'table_cell': return 'tc';
    case 'sdtBlock': case 'sdtInline': return `sdt${sdtAlias(node.attrs['sdtPr'])}`;
    case 'opaqueBlock': return String(node.attrs['kind']);
    default: return node.type.name;
  }
}

function sdtAlias(sdtPr: unknown): string {
  const items = (sdtPr as { rPrOrAliasOrLock?: { name?: { localPart?: string }; value?: { val?: unknown } }[] } | null)?.rPrOrAliasOrLock ?? [];
  const alias = items.find((i) => i.name?.localPart === 'alias')?.value?.val ?? items.find((i) => i.name?.localPart === 'tag')?.value?.val;
  return alias ? ` (${String(alias)})` : '';
}

/** Breadcrumbs from the state's selection: `body > sdt (alias) > p`, plus `r` when the caret is in a run. */
export function crumbsOf(state: EditorState): { crumbs: Crumb[]; run: string | null } {
  const $from = state.selection.$from;
  const crumbs: Crumb[] = [];
  for (let d = 1; d <= $from.depth; d++) {
    const node = $from.node(d);
    crumbs.push({ label: labelOf(node), node, pos: $from.before(d), depth: d });
  }
  if ($from.depth === 0) {
    const node = state.doc.nodeAt($from.pos);
    if (node) crumbs.push({ label: labelOf(node), node, pos: $from.pos, depth: 0 });
  }
  const around = $from.nodeAfter ?? $from.nodeBefore;
  const run = around?.marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'] as string | undefined
    ?? (around && 'runSource' in around.attrs ? (around.attrs['runSource'] as string | null) : null);
  return { crumbs, run: run ?? null };
}

/** The block-level node holding the caret, with its position. */
export function blockAt(state: EditorState): { node: PMNode; pos: number; index: number } | undefined {
  const $from = state.selection.$from;
  if ($from.depth === 0) {
    const node = state.doc.nodeAt($from.pos);
    return node ? { node, pos: $from.pos, index: $from.index(0) } : undefined;
  }
  return { node: $from.node(1), pos: $from.before(1), index: $from.index(0) };
}

/** Every top-level block the selection touches, first to last (ED-001 section 3.5: the selected element or elements). */
export function blocksIn(state: EditorState): { nodes: PMNode[]; from: number; to: number } {
  const { $from, $to } = state.selection;
  const from = $from.index(0);
  let to = $to.index(0);
  // A selection ending exactly at a block's start does not include that block.
  if (to > from && $to.depth > 0 && $to.parentOffset === 0 && $to.depth === 1) to--;
  if (to >= state.doc.childCount) to = state.doc.childCount - 1;
  const nodes: PMNode[] = [];
  for (let i = from; i <= to; i++) nodes.push(state.doc.child(i));
  return { nodes, from, to };
}

/** The run element under the caret inside a rebuilt paragraph element, by the caret's text offset. */
export function runUnderCaret(session: EditorDocument, state: EditorState, block: Element): { run: Element; index: number } | undefined {
  const $from = state.selection.$from;
  // The nearest paragraph up the chain (the caret may sit inside an inline content control).
  let depth = $from.depth;
  while (depth >= 1 && $from.node(depth).type.name !== 'paragraph') depth--;
  if (depth < 1) return undefined;
  const paragraph = $from.node(depth);
  const start = $from.start(depth);
  const target = $from.pos - start;
  // The caret's offset in the paragraph's text, counting text and the characters atoms stand for,
  // through inline controls, which contribute their content and no characters of their own.
  let offset = 0;
  paragraph.descendants((child, childPos) => {
    if (childPos >= target) return false;
    if (child.isText) { offset += Math.min(child.nodeSize, target - childPos); return false; }
    if (child.type.name === 'tab' || child.type.name === 'hardBreak') { offset += 1; return false; }
    return true;
  });
  // The paragraph element inside the block (the block may be an sdt around it).
  const paragraphElement = depth === 1 ? block : findParagraph(block, $from, depth);
  const segments = segmentsOf((paragraphElement ?? block).value as object);
  let runIndex = -1;
  let lastRun: object | undefined;
  for (const seg of segments) {
    if (seg.run !== lastRun) { runIndex++; lastRun = seg.run; }
    if (offset >= seg.start && offset <= seg.end) {
      const el = seg.runOwner[seg.runIndex];
      return el ? { run: el, index: runIndex } : undefined;
    }
  }
  const last = segments[segments.length - 1];
  const el = last ? last.runOwner[last.runIndex] : undefined;
  void session;
  return el ? { run: el, index: runIndex } : undefined;
}

/** The n-th paragraph element inside a block, by the caret's index path through block-level controls. */
function findParagraph(block: Element, $from: import('prosemirror-model').ResolvedPos, depth: number): Element | undefined {
  let el: Element | undefined = block;
  for (let d = 1; d < depth && el; d++) {
    const index = $from.index(d);
    const value = el.value as { sdtContent?: { content?: Element[] }; content?: Element[] };
    const content: Element[] = (value.sdtContent?.content ?? value.content ?? []).filter((e) => typeNameOf(e) !== undefined && e.name.localPart !== 'proofErr' && e.name.localPart !== 'bookmarkStart' && e.name.localPart !== 'bookmarkEnd');
    el = content[index];
  }
  return el;
}

export async function xmlOf(element: Element): Promise<string> {
  // Shown, never parsed: a character XML 1.0 excludes is shown as held, not refused (objects 0.8.0, CR-020).
  return pretty(await marshalString(element, { checkCharacters: false }));
}

const COMMON_DECLARATIONS = [
  ['w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'],
  ['r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'],
  ['w14', 'http://schemas.microsoft.com/office/word/2010/wordml'],
  ['w15', 'http://schemas.microsoft.com/office/word/2012/wordml'],
  ['wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'],
  ['a', 'http://schemas.openxmlformats.org/drawingml/2006/main'],
  ['pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture'],
  ['mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006'],
  ['m', 'http://schemas.openxmlformats.org/officeDocument/2006/math'],
] as const;

export class XmlSyntaxError extends Error {
  constructor(message: string, readonly line: number | undefined) {
    super(message);
    this.name = 'XmlSyntaxError';
  }
}

function wrapFragment(xml: string, expected: 'block' | 'run'): string {
  const text = xml.trim().replace(/^<\?xml[^>]*\?>\s*/, '');
  const decls = COMMON_DECLARATIONS.map(([p, ns]) => `xmlns:${p}="${ns}"`).join(' ');
  // The wrapper opens on the first line of the text, so the parser's line numbers are the editor's.
  return expected === 'block' ? `<w:body ${decls}>${text}</w:body>` : `<w:p ${decls}>${text}</w:p>`;
}

/**
 * The browser parser's verdict on edited XML: null when well-formed, else the message and the
 * line it names (Chromium: "error on line 7 at column 10: ..."; Firefox names the line in its
 * text too). The unmarshaller would otherwise be handed the parser's error element and complain
 * about an unknown xhtml element, which tells the user nothing.
 */
export function xmlSyntaxError(xml: string, expected: 'block' | 'run' = 'block'): { message: string; line: number | undefined } | null {
  if (typeof DOMParser === 'undefined') return null;
  const doc = new DOMParser().parseFromString(wrapFragment(xml, expected), 'application/xml');
  const error = doc.getElementsByTagName('parsererror')[0];
  if (!error) return null;
  const raw = (error.textContent ?? 'not well-formed XML').replace(/\s+/g, ' ').trim();
  const m = /line (\d+)/i.exec(raw);
  const line = m ? Number(m[1]) : undefined;
  const message = raw.replace(/^This page contains the following errors:\s*/i, '').replace(/Below is a rendering of the page up to the first error\.?/i, '').trim();
  return { message, line };
}

/**
 * Parse edited XML into one element. The pane's text carries the declarations `marshalString`
 * wrote; a fragment typed without them gets the common wordprocessing prefixes declared on a
 * wrapper. The unmarshaller's error is the message on a bad edit.
 */
export async function parseElements(xml: string, expected: 'block' | 'run'): Promise<Element[]> {
  const syntax = xmlSyntaxError(xml, expected);
  if (syntax) throw new XmlSyntaxError(syntax.message, syntax.line);
  const wrapped = wrapFragment(xml, expected);
  const parsed = await unmarshalString(wrapped);
  const items = (parsed.value as { content?: Element[] }).content ?? [];
  const elements = items.filter((e) => typeNameOf(e) !== undefined);
  if (elements.length === 0) throw new Error(`expected at least one ${expected === 'block' ? 'block-level element' : 'run'}`);
  if (expected === 'run' && elements.length !== 1) throw new Error(`expected one run, found ${elements.length}`);
  return elements;
}

export async function parseElement(xml: string, expected: 'block' | 'run'): Promise<Element> {
  const elements = await parseElements(xml, expected);
  if (elements.length !== 1) throw new Error(`expected one ${expected === 'block' ? 'block-level element' : 'run'}, found ${elements.length}`);
  return elements[0]!;
}
