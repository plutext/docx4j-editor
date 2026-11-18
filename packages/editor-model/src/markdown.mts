// SPDX-License-Identifier: GPL-3.0-or-later
//
// Markdown in (ED-005 sections 8.4 and 12.66): parsed with mdast into blocks, then built with the
// engine's builders into the elements Word would make. Three callers share it:
//
// - the agent (`docs/developer/agent.md` section 2.6) reads it **strictly**: a paragraph, headings 1 to 6 by
//   the document's style names, bold, italic and strikethrough as direct formatting, bulleted and
//   numbered lists (nested by indentation) through the document's numbering, a GFM table (its first
//   row the header), and a link; anything else is `MarkdownUnsupported`, naming the construct, which
//   the companion answers as `markdown.unsupported`;
// - paste ("Format as Markdown") and Open of a `.md` file read it **leniently**: inline code takes
//   Word's `HTML Code` character style and a code block's lines `HTML Preformatted`, a block quote's
//   paragraphs `Quote`, a reference link resolves through its definition, `<br>` is a line break, and
//   what has no place (an image, other HTML, a rule, a link to a file) is dropped and counted, its
//   text kept where it has any.
//
// The parts a build needs (a style, a numbering definition, a link's relationship) come from the
// caller (`BuildContext`), so a dry run can name them without writing them.
import { fromMarkdown } from 'mdast-util-from-markdown';
import { gfmTableFromMarkdown } from 'mdast-util-gfm-table';
import { gfmStrikethroughFromMarkdown } from 'mdast-util-gfm-strikethrough';
import { gfmTable } from 'micromark-extension-gfm-table';
import { gfmStrikethrough } from 'micromark-extension-gfm-strikethrough';
import type { Definition, List, ListItem, Nodes, PhrasingContent, Root, RootContent, Table as MdastTable } from 'mdast';
import { p, r, tbl, W_NS, linkParents, WordprocessingMLPackage, type Element, type PageSizePaper } from '@docx4j/core-ts';
import { ensureStylesIn } from './knownstyles.mjs';
import { runItemsFor } from './excluded.mjs';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import * as el from '@docx4j/generated-objects-ts/modules/org_docx4j_wml.el';

/** A piece of a paragraph's text with its formatting. A `\n` in `text` is a line break, a `\t` a tab. */
export interface Inline {
  text: string;
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  /** Inline code (lenient only): Word's `HTML Code` character style. */
  code?: boolean;
  /** The link's address. */
  link?: string;
}

/** A list item's place: its nesting level and the list it belongs to (one per Markdown list and kind). */
export interface ListPlace {
  ordered: boolean;
  level: number;
  /** Which list: items of one key share a numbering definition. */
  key: string;
}

export type MdBlock =
  | {
    kind: 'paragraph';
    heading?: number;
    list?: ListPlace;
    /** In a block quote (lenient only): Word's `Quote`. */
    quote?: boolean;
    /** A line of a code block (lenient only): Word's `HTML Preformatted`. */
    code?: boolean;
    inlines: Inline[];
  }
  | { kind: 'table'; rows: Inline[][][] };

/** The style ids the lenient blocks take, Word's built-ins (the engine splices them since core-ts 0.4.0). */
export const MARKDOWN_STYLES = { code: 'HTMLCode', codeBlock: 'HTMLPreformatted', quote: 'Quote' } as const;

/**
 * An address a pasted link keeps (Markdown's, and HTML's since ED-005 section 12.67): the web's and
 * mail's. A place in the source document (`#name`), a file or another scheme has no target here.
 */
export function webAddress(url: string): string | undefined {
  const scheme = /^([a-z][a-z0-9+.-]*):/i.exec(url.trim())?.[1]?.toLowerCase();
  return scheme === 'http' || scheme === 'https' || scheme === 'mailto' ? url.trim() : undefined;
}

/** A construct outside the agent's subset, in a strict read. */
export class MarkdownUnsupported extends Error {
  constructor(readonly construct: string) {
    super(`${construct} is not in the Markdown subset`);
    this.name = 'MarkdownUnsupported';
  }
}

/** What a read made: the blocks, and in a lenient read what was dropped, by kind, with how many times. */
export interface MarkdownRead {
  blocks: MdBlock[];
  dropped: Record<string, number>;
}

export interface ReadOptions {
  /** Paste's and Open's reading (above); the agent's is strict. */
  lenient?: boolean;
}

/** The walk's state: the mode, the definitions a reference link resolves through, the count of what was dropped. */
class Reader {
  readonly dropped: Record<string, number> = {};
  private readonly definitions = new Map<string, string>();
  private lists = 0;

  constructor(readonly lenient: boolean, root: Root) {
    if (lenient) collectDefinitions(root, this.definitions);
  }

  /** In a strict read, the refusal; in a lenient one, the construct counted under `kind`. */
  drop(construct: string, kind: string): void {
    if (!this.lenient) throw new MarkdownUnsupported(construct);
    this.dropped[kind] = (this.dropped[kind] ?? 0) + 1;
  }

  /** A link's address, when it is the web's or mail's; any other, in a lenient read, undefined (the text kept, the link counted). */
  linkAddress(url: string): string | undefined {
    if (webAddress(url) !== undefined) return url;
    const scheme = /^([a-z][a-z0-9+.-]*):/i.exec(url)?.[1]?.toLowerCase();
    this.drop(scheme ? `a link to a ${scheme}: address` : `a link to ${JSON.stringify(url)} (only http, https and mailto)`, 'links');
    return undefined;
  }

  /** The inlines of phrasing content, formatting carried down. */
  inlines(nodes: readonly PhrasingContent[], format: Omit<Inline, 'text'> = {}): Inline[] {
    const out: Inline[] = [];
    for (const node of nodes) {
      switch (node.type) {
        // A newline inside a text node is a soft break, a space as CommonMark renders it (a hard-wrapped
        // source stays one paragraph); only `break` and `<br>` are line breaks.
        case 'text': out.push({ ...format, text: node.value.replace(/\r?\n/g, ' ') }); break;
        case 'strong': out.push(...this.inlines(node.children, { ...format, bold: true })); break;
        case 'emphasis': out.push(...this.inlines(node.children, { ...format, italic: true })); break;
        case 'delete': out.push(...this.inlines(node.children, { ...format, strike: true })); break;
        case 'link': out.push(...this.linked(node.children, node.url, format)); break;
        case 'break': out.push({ ...format, text: '\n' }); break;
        case 'inlineCode':
          if (!this.lenient) throw new MarkdownUnsupported('inline code');
          out.push({ ...format, code: true, text: node.value });
          break;
        case 'linkReference': {
          const url = this.lenient ? this.definitions.get(node.identifier) : undefined;
          if (url === undefined) throw new MarkdownUnsupported('a reference-style link');
          out.push(...this.linked(node.children, url, format));
          break;
        }
        case 'image': case 'imageReference': this.drop('an image', 'images'); break;
        case 'html':
          // A `<br>` is a line break (GFM tables have no other); any other tag goes, the text between kept.
          if (this.lenient && /^<br\s*\/?>$/i.test(node.value.trim())) out.push({ ...format, text: '\n' });
          else this.drop('HTML', 'HTML');
          break;
        case 'footnoteReference': this.drop('a footnote', 'footnotes'); break;
        default: this.drop(`Markdown's ${(node as { type: string }).type}`, (node as { type: string }).type);
      }
    }
    return out;
  }

  private linked(children: readonly PhrasingContent[], url: string, format: Omit<Inline, 'text'>): Inline[] {
    const address = this.linkAddress(url);
    return this.inlines(children, address === undefined ? format : { ...format, link: address });
  }

  /** A list's items as paragraphs: nested lists one level deeper, a second paragraph in an item a line break. */
  listBlocks(list: List, level: number, top: number, quote: boolean, out: MdBlock[]): void {
    const place: ListPlace = { ordered: list.ordered === true, level: Math.min(level, 8), key: `${top}:${list.ordered === true ? 'numbered' : 'bullet'}` };
    for (const item of list.children as ListItem[]) {
      let paragraph: Extract<MdBlock, { kind: 'paragraph' }> | undefined;
      for (const child of item.children) {
        if (child.type === 'paragraph') {
          if (paragraph) paragraph.inlines.push({ text: '\n' }, ...this.inlines(child.children));
          else { paragraph = { kind: 'paragraph', list: place, inlines: this.inlines(child.children) }; out.push(paragraph); }
        } else if (child.type === 'list') {
          if (!paragraph) { paragraph = { kind: 'paragraph', list: place, inlines: [] }; out.push(paragraph); }
          this.listBlocks(child, level + 1, top, quote, out);
        } else if (this.lenient) {
          // A code block or a quote in an item follows it, outside the list.
          if (!paragraph) { paragraph = { kind: 'paragraph', list: place, inlines: [] }; out.push(paragraph); }
          this.block(child, quote, out);
        } else {
          throw new MarkdownUnsupported(blockName(child));
        }
      }
      if (!paragraph) out.push({ kind: 'paragraph', list: place, inlines: [] });
    }
  }

  /** One block of the document (or of a quote), appended to `out`. */
  block(node: RootContent, quote: boolean, out: MdBlock[]): void {
    const q = quote ? { quote: true } : {};
    switch (node.type) {
      case 'paragraph': out.push({ kind: 'paragraph', ...q, inlines: this.inlines(node.children) }); return;
      case 'heading': out.push({ kind: 'paragraph', heading: node.depth, inlines: this.inlines(node.children) }); return;
      case 'list': this.listBlocks(node, 0, this.lists++, quote, out); return;
      case 'table': out.push({ kind: 'table', rows: (node as MdastTable).children.map((row) => row.children.map((cell) => this.inlines(cell.children))) }); return;
      case 'definition': if (this.lenient) return; break;
      case 'code':
        if (!this.lenient) break;
        for (const line of node.value.split(/\r\n|\r|\n/)) out.push({ kind: 'paragraph', code: true, inlines: line === '' ? [] : [{ text: line }] });
        return;
      case 'blockquote':
        if (!this.lenient) break;
        for (const child of node.children) this.block(child, true, out);
        return;
      case 'thematicBreak': if (this.lenient) { this.drop('a thematic break', 'rules'); return; } break;
      case 'html': if (this.lenient) { this.drop('HTML', 'HTML'); return; } break;
      default: break;
    }
    if (this.lenient) this.drop(blockName(node), node.type);
    else throw new MarkdownUnsupported(blockName(node));
  }
}

function blockName(node: RootContent): string {
  switch (node.type) {
    case 'thematicBreak': return 'a thematic break';
    case 'code': return 'a code block';
    case 'blockquote': return 'a block quote';
    case 'html': return 'HTML';
    case 'definition': return 'a link reference definition';
    case 'footnoteDefinition': return 'a footnote';
    case 'yaml': return 'front matter';
    default: return `Markdown's ${node.type}`;
  }
}

/** The link reference definitions anywhere in the tree, the first of a label winning (CommonMark). */
function collectDefinitions(node: Nodes, into: Map<string, string>): void {
  if (node.type === 'definition') {
    // mdast normalises a label (CommonMark's case fold, whitespace collapsed) into `identifier`, a reference's as a definition's.
    const key = (node as Definition).identifier;
    if (!into.has(key)) into.set(key, (node as Definition).url);
    return;
  }
  if ('children' in node) for (const child of node.children as Nodes[]) collectDefinitions(child, into);
}

/** Markdown as blocks: strictly the agent's subset (anything else `MarkdownUnsupported`), or leniently, with what was dropped. */
export function readMarkdown(markdown: string, options: ReadOptions = {}): MarkdownRead {
  const root = fromMarkdown(markdown, {
    extensions: [gfmTable(), gfmStrikethrough({ singleTilde: false })],
    mdastExtensions: [gfmTableFromMarkdown(), gfmStrikethroughFromMarkdown()],
  });
  const reader = new Reader(options.lenient === true, root);
  const blocks: MdBlock[] = [];
  for (const node of root.children) reader.block(node, false, blocks);
  return { blocks, dropped: reader.dropped };
}

/** Plain text (`edit`'s `text`): a paragraph per line, nothing interpreted. */
export function plainBlocks(text: string): MdBlock[] {
  return text.split(/\r\n|\r|\n/).map((line) => ({ kind: 'paragraph' as const, inlines: line === '' ? [] : [{ text: line }] }));
}

/** The words of an inline list, for a comment's content or a check. */
export function inlineText(list: readonly Inline[]): string {
  return list.map((i) => i.text).join('');
}

/**
 * Whether plain text reads as Markdown, for the offer after a plain paste (ED-005 section 12.66): a
 * line that opens a heading, a list item, a quote, a fence or a table row, or a span of bold, code or
 * a link. A cheap test of the text alone, never a parse; a false offer costs one status line.
 */
export function looksLikeMarkdown(text: string): boolean {
  if (/^(#{1,6}\s+\S|\s{0,3}(?:[-*+]|\d{1,9}[.)])\s+\S|>\s?\S|```|~~~|\|.*\|\s*$)/m.test(text)) return true;
  return /\*\*[^*\n]+\*\*|__[^_\n]+__|`[^`\n]+`|\[[^\]\n]+\]\([^)\s]+\)/.test(text);
}

// ---------------------------------------------------------------------------------------------
// Building
// ---------------------------------------------------------------------------------------------

/** What a build needs from the document: each part it touches is the caller's to write, or, in a dry run, to name. */
export interface BuildContext {
  /** The style a plain paragraph takes (the target's, section 2.6), or undefined for Normal. */
  paragraphStyle: string | undefined;
  /** The style id of a heading level, the style added from docx4j's defaults when the document lacks it. */
  headingStyle(level: number): Promise<string>;
  /** The `w:numId` of a list, a definition made for it once per key. */
  numId(place: ListPlace): Promise<string>;
  /** The relationship id of a link's address. */
  linkId(address: string): Promise<string>;
  /** The link's character style (`Hyperlink`), added when the document lacks it. */
  linkStyle(): Promise<string>;
  /** A new table's style, or undefined for none. */
  tableStyle(): Promise<string | undefined>;
  /** The section's text width in twips, which a table spans; A4 with 2.54 cm margins (9,026) when absent. */
  textWidth?: number;
  /** A lenient block's style (`MARKDOWN_STYLES`), added when the document lacks it; undefined leaves the text plain. Only a lenient read needs it. */
  builtInStyle?(id: string): Promise<string | undefined>;
}

/** A hyperlink holder around runs, as the engine's `Range.hyperlink` makes one (`w:history`, as Word writes). */
function hyperlink(runs: Element[], id: string): Element {
  const value = { TYPE_NAME: 'org_docx4j_wml.P.Hyperlink', id, history: true, content: runs } as unknown as wml.P.Hyperlink;
  return { name: { namespaceURI: W_NS, localPart: 'hyperlink' }, value } as Element;
}

async function runOf(inline: Inline, style: string | undefined, ctx: BuildContext): Promise<Element> {
  const code = inline.code && style === undefined ? await ctx.builtInStyle?.(MARKDOWN_STYLES.code) : undefined;
  const runStyle = style ?? code;
  const run = r('', { ...(runStyle ? { style: runStyle } : {}), ...(inline.bold ? { bold: true } : {}), ...(inline.italic ? { italic: true } : {}), ...(inline.strike ? { strikeThrough: true } : {}) });
  run.value.content = runItemsFor(inline.text) as wml.R['content'];
  return run;
}

/** A paragraph's content: runs, and a hyperlink around the runs of one link. */
async function runsOf(list: readonly Inline[], ctx: BuildContext): Promise<Element[]> {
  const out: Element[] = [];
  let i = 0;
  while (i < list.length) {
    const inline = list[i]!;
    if (inline.link === undefined) {
      if (inline.text !== '') out.push(await runOf(inline, undefined, ctx));
      i++;
      continue;
    }
    const address = inline.link;
    const runs: Element[] = [];
    const style = await ctx.linkStyle();
    for (; i < list.length && list[i]!.link === address; i++) {
      const piece = list[i]!;
      if (piece.text !== '') runs.push(await runOf(piece, style, ctx));
    }
    if (runs.length > 0) out.push(hyperlink(runs, await ctx.linkId(address)));
  }
  return out;
}

async function paragraphOf(block: Extract<MdBlock, { kind: 'paragraph' }>, ctx: BuildContext): Promise<Element<wml.P>> {
  // A list item takes its look from the numbering, as the editor's list commands write it (`numPr` alone).
  let style: string | undefined;
  if (block.heading !== undefined) style = await ctx.headingStyle(block.heading);
  else if (block.code) style = await ctx.builtInStyle?.(MARKDOWN_STYLES.codeBlock);
  else if (block.list) style = undefined;
  else if (block.quote) style = await ctx.builtInStyle?.(MARKDOWN_STYLES.quote);
  else style = ctx.paragraphStyle;
  const para = p(await runsOf(block.inlines, ctx), style ? { style } : {});
  if (block.list) {
    const pPr = (para.value.pPr ??= { TYPE_NAME: 'org_docx4j_wml.PPr' } as wml.PPr);
    pPr.numPr = { TYPE_NAME: 'org_docx4j_wml.PPrBase.NumPr', ilvl: { val: block.list.level }, numId: { val: Number(await ctx.numId(block.list)) } } as unknown as wml.PPrBase.NumPr;
  }
  return para;
}

/**
 * A Markdown table's column widths: the text width shared in proportion to each column's longest cell, between
 * twelve and sixty characters' worth, so a column of long identifiers gets room and a column of numbers little, and
 * the sum is the text width to the twip (ED-005 section 12.92; equal shares had put a long column in a fifth).
 */
export function tableColumnWidths(rows: readonly (readonly (readonly Inline[])[])[], columns: number, textWidth: number): number[] {
  const weights: number[] = [];
  for (let c = 0; c < columns; c++) {
    const longest = Math.max(0, ...rows.map((row) => (row[c] ?? []).reduce((n, inline) => n + inline.text.length, 0)));
    weights.push(Math.min(60, Math.max(12, longest)));
  }
  const total = weights.reduce((a, b) => a + b, 0);
  const widths = weights.map((w) => Math.floor((w / total) * textWidth));
  widths[widths.length - 1]! += textWidth - widths.reduce((a, b) => a + b, 0);
  return widths;
}

async function tableOf(block: Extract<MdBlock, { kind: 'table' }>, ctx: BuildContext): Promise<Element> {
  const width = Math.max(1, ...block.rows.map((row) => row.length));
  const style = await ctx.tableStyle();
  const widths = tableColumnWidths(block.rows, width, ctx.textWidth ?? 9026);
  const table = tbl(block.rows.map(() => new Array<string>(width).fill('')), { widths, ...(style ? { style } : {}) });
  const rows = (table.value.content ?? []) as Element<wml.Tr>[];
  for (let ri = 0; ri < rows.length; ri++) {
    const tr = rows[ri]!.value;
    const cells = (tr.content ?? []) as Element<wml.Tc>[];
    for (let ci = 0; ci < cells.length; ci++) {
      const cell = block.rows[ri]![ci] ?? [];
      cells[ci]!.value.content = [p(await runsOf(cell, { ...ctx, paragraphStyle: undefined }))] as wml.Tc['content'];
    }
    // The first row is the header, repeated on each page as Word marks one.
    if (ri === 0) ((tr.trPr ??= { TYPE_NAME: 'org_docx4j_wml.TrPr' } as wml.TrPr).cnfStyleOrDivIdOrGridBefore ??= []).push(el.tblHeader({}) as never);
  }
  linkParents(table, undefined);
  return table;
}

/** The elements of the blocks, in order. */
export async function buildBlocks(blocks: readonly MdBlock[], ctx: BuildContext): Promise<Element[]> {
  const out: Element[] = [];
  for (const block of blocks) out.push(block.kind === 'paragraph' ? await paragraphOf(block, ctx) : await tableOf(block, ctx));
  return out;
}

// ---------------------------------------------------------------------------------------------
// A document of its own (Open of a `.md` file)
// ---------------------------------------------------------------------------------------------

const HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

/** A Markdown file as a new document: the bytes to open, and what the reading dropped. */
export interface MarkdownDocument {
  bytes: Uint8Array;
  dropped: Record<string, number>;
}

/**
 * A Markdown file read leniently into a new document (ED-005 section 12.66): the engine's
 * `createPackage` in the page size, as New makes one, its body the blocks built with the engine's
 * builders as the agent's are, and the styles, lists and links they need written into its parts.
 * Saved to bytes, so it opens as any file does: no edit to undo, nothing unsaved.
 */
export async function markdownDocument(markdown: string, options: { pageSize?: PageSizePaper; margins?: number } = {}): Promise<MarkdownDocument> {
  const { blocks, dropped } = readMarkdown(markdown, { lenient: true });
  const pkg = await WordprocessingMLPackage.createPackage(options.pageSize ? { pageSize: options.pageSize } : {});
  // The margins a side, in twips: `createPackage`'s 2.54 cm unless the caller's preference says otherwise
  // (the application's option for new documents, ED-007 section 8.1 item 15). Set before the body is built,
  // so the tables' grid and the fit that follows the open see the text width the document will have.
  const margins = options.margins ?? 1440;
  const ensured = async (id: string): Promise<string | undefined> => {
    await ensureStylesIn(pkg, [id]);
    return pkg.propertyResolverOrUndefined?.getStyle(id) ? id : undefined;
  };
  const lists = new Map<string, string>();
  let firstNumbered: string | undefined;
  const ctx: BuildContext = {
    paragraphStyle: undefined,
    headingStyle: async (level) => (await ensured(`Heading${level}`)) ?? 'Normal',
    numId: async (place) => {
      const made = lists.get(place.key);
      if (made !== undefined) return made;
      // A second numbered list starts again at 1, over the first's definition, as paste and the agent do.
      const numId = place.ordered && firstNumbered !== undefined ? pkg.numbering.restart(firstNumbered) : await pkg.numbering.newList({ bullet: !place.ordered });
      if (place.ordered && firstNumbered === undefined) firstNumbered = numId;
      lists.set(place.key, numId);
      return numId;
    },
    linkId: async (address) => pkg.getMainDocumentPart().getRelationshipsPart(true)!.addExternalRelationship(HYPERLINK, address).id,
    linkStyle: async () => (await ensured('Hyperlink')) ?? 'Hyperlink',
    tableStyle: async () => ensured('TableGrid'),
    // The page's text width: its width less the margins a side.
    textWidth: ({ LETTER: 12240, LEGAL: 12240, A3: 16838, A4: 11906, A5: 8391, B4JIS: 14570 }[options.pageSize ?? 'A4'] ?? 11906) - 2 * margins,
    builtInStyle: ensured,
  };
  const main = pkg.getMainDocumentPart();
  const document = await main.getContents();
  const body = document.body!;
  if (margins !== 1440) {
    const pgMar = (body.sectPr!.pgMar ??= { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 708, footer: 708, gutter: 0 });
    pgMar.top = pgMar.right = pgMar.bottom = pgMar.left = margins;
  }
  body.content = (await buildBlocks(blocks, ctx)) as typeof body.content;
  linkParents(document, undefined);
  return { bytes: await pkg.save(), dropped };
}
