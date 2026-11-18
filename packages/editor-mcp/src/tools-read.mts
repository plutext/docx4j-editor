// SPDX-License-Identifier: GPL-3.0-or-later
//
// The reading tools (docs/developer/agent.md section 3.2): the outline, find (text or XPath), the document as
// Markdown with addresses and as plain text, the comments, and the tracked changes as the review panel
// lists them. Each reads the room's document as it is now, the person's edits included (`readTree`), and
// presents its text as data (section 2.12). The outline, find's text search, read and extract_text read
// through the engine's text model with the specification's options (section 2.13, ED-008; `textmodel.mts`).
import { writeFile } from 'node:fs/promises';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { Node as PMNode } from 'prosemirror-model';
import { Paragraph, Table, bindSnapshot, boundsOf, officeMatches, MarkdownExporter, PlainText, type Element, type SearchSegment, type TextBlock, type TextResult, type TextRun } from '@docx4j/core-ts';
import * as el from '@docx4j/generated-objects-ts/modules/org_docx4j_wml.el';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { trackedChangesOf, preservedChangesOf, preservedGroupsOf, preservedTargetOf, referenceIdOf, isDeleted, XPATH_NAMESPACES, type TreeBindings, type TrackedChangeInfo, type CommentThread, type PreservedGroup } from '@docx4j/editor-model';
import { ToolError } from './errors.mjs';
import { writableUnderRoots } from './paths.mjs';
import { addressSchema, paraIdOf, resolveBlock, type AddressInput } from './address.mjs';
import { answer, documentArg, inRoom, ok, searchOptions, searchShape, type ToolHost, type ToolRoom } from './kit.mjs';
import { SCOPE_SHAPE, SERIALISATION_SHAPE, textOptionsOf, readText, blocksIn, blocksAt, resultOf, warningsOf, warningText, type ReadWarning } from './textmodel.mjs';

/** A block's address from a node of the document the tree was exported from (the bridge of ED-005 section 12.27). */
export function addressOfNode(b: TreeBindings, node: PMNode): string | undefined {
  const element = b.elementOf(node);
  return element ? b.body.addressOf({ element, container: [] }) : undefined;
}

/** The address of the paragraph (or table) a position is in or before. */
export function addressAt(b: TreeBindings, pos: number): string | undefined {
  const $pos = b.doc.resolve(Math.max(0, Math.min(pos, b.doc.content.size)));
  const after = $pos.nodeAfter;
  if (after && (after.type.name === 'paragraph' || after.type.name === 'table')) {
    const a = addressOfNode(b, after);
    if (a) return a;
  }
  for (let d = $pos.depth; d > 0; d--) {
    const node = $pos.node(d);
    if (node.type.name === 'paragraph' || node.type.name === 'table') {
      const a = addressOfNode(b, node);
      if (a) return a;
    }
  }
  return undefined;
}

/** Whether an address lies in (or is) another. */
function within(address: string | undefined, scope: string | undefined): boolean {
  if (scope === undefined) return true;
  return address !== undefined && (address === scope || address.startsWith(`${scope}/`));
}

/** The ordinal an address resolves to, for filtering by it. */
function scopeOf(b: TreeBindings, address: AddressInput | undefined): string | undefined {
  if (address === undefined) return undefined;
  const { block } = resolveBlock(b.body, address);
  return b.body.addressOf(block);
}

/** A result's text held to the inline cap (section 2.9): written to `output_path` when given, else cut. */
async function capped(host: ToolHost, text: string, maxChars: number | undefined, outputPath: string | undefined, overwrite: boolean | undefined): Promise<{ text?: string; truncated: boolean; chars: number; path?: string }> {
  const cap = Math.min(maxChars ?? host.maxInlineChars, host.maxInlineChars);
  if (text.length <= cap) return { text, truncated: false, chars: text.length };
  if (outputPath !== undefined) {
    const w = await writableUnderRoots(host.roots, outputPath);
    if (w.exists && overwrite !== true) throw new ToolError('path.exists', `${outputPath} exists`, 'pass overwrite: true, or choose another name', { path: w.path });
    await writeFile(w.path, text);
    return { truncated: false, chars: text.length, path: w.path };
  }
  return { text: text.slice(0, cap), truncated: true, chars: text.length };
}

// ---------------------------------------------------------------- tracked changes

/** One tracked change as `get_tracked_changes` lists it. */
export interface ListedChange {
  /** What `accept_tracked_changes` and `reject_tracked_changes` take; absent for what only "all" resolves. */
  key?: string;
  kind: string;
  author: string | null;
  date: string | null;
  address: string | null;
  text: string;
  /** Word shows it and the editor keeps it as it is (the review panel's "In Word only"). */
  shown_in_word_only?: true;
}

function changeKind(c: TrackedChangeInfo): string {
  if (c.move !== undefined) return 'move';
  if (c.row) return c.type === 'Added' ? 'row_insertion' : 'row_deletion';
  if (c.type === 'Added') return 'insertion';
  if (c.type === 'Deleted') return 'deletion';
  return c.pieces[0]?.what === 'properties' ? 'paragraph_properties' : 'formatting';
}

const PRESERVED_KINDS: Record<PreservedGroup['kind'], string> = {
  tableProperties: 'table_properties', tblPrChange: 'table_properties', tblGridChange: 'table_properties', trPrChange: 'table_properties', tcPrChange: 'table_properties',
  cellIns: 'cell_insertion', cellDel: 'cell_deletion', cellMerge: 'cell_merge', sectPrChange: 'section_properties', markRPrChange: 'paragraph_mark_formatting', numberingChange: 'numbering',
  // A revision inside an opaque atom (core-ts 0.6.0; ED-008 section 8.8 item 10): the engine lists and resolves it, the projection shows it as the atom.
  mathIns: 'equation_insertion', mathDel: 'equation_deletion', holderIns: 'content_control_insertion', holderDel: 'content_control_deletion',
};

/** The key a revision shown in Word only goes by, which the resolving tools read back. */
export const PRESERVED_KEY = 'preserved:';

/** The document's tracked changes as the review panel lists them: a move's two halves one entry; the revisions shown in Word only among them. */
export function listChanges(room: ToolRoom, b: TreeBindings): ListedChange[] {
  const changes = trackedChangesOf(b.doc);
  const keys = room.session.changeKeys(changes);
  const out: { at: number; change: ListedChange }[] = [];
  const moves = new Set<string>();
  changes.forEach((c, i) => {
    if (c.move !== undefined) {
      if (moves.has(c.move)) return;
      moves.add(c.move);
      // The move listed at its destination, the text as it now stands.
      const destination = changes.find((d) => d.move === c.move && d.type === 'Added') ?? c;
      out.push({ at: c.from, change: { key: keys[changes.indexOf(destination)]!, kind: 'move', author: destination.author, date: destination.date, address: addressAt(b, destination.from) ?? null, text: destination.text } });
      return;
    }
    out.push({ at: c.from, change: { key: keys[i]!, kind: changeKind(c), author: c.author, date: c.date, address: addressAt(b, c.from) ?? null, text: c.text } });
  });
  const groups = preservedGroupsOf(b.doc, preservedChangesOf(b.doc, (id) => room.session.ctx.ids.get(id)?.value));
  for (const g of groups) {
    const target = preservedTargetOf(b.doc, g, groups);
    out.push({ at: g.pos, change: { ...(target ? { key: PRESERVED_KEY + JSON.stringify(target) } : {}), kind: PRESERVED_KINDS[g.kind], author: g.author, date: g.date, address: addressAt(b, g.pos) ?? null, text: '', shown_in_word_only: true } });
  }
  return out.sort((a, b2) => a.at - b2.at).map((e) => e.change);
}

// ---------------------------------------------------------------- comments

interface Anchor { paragraph: PMNode; text: string }

/** Where each comment is anchored in the body: its first paragraph and its text as accepted. */
function commentAnchorsIn(b: TreeBindings, ids: Parameters<typeof referenceIdOf>[1]): Map<string, Anchor> {
  const out = new Map<string, Anchor>();
  b.doc.descendants((node) => {
    if (node.type.name !== 'paragraph') return true;
    node.descendants((child) => {
      const here = new Set<string>();
      for (const m of child.marks) if (m.type.name === 'comment') here.add(String(m.attrs['id']));
      const ref = referenceIdOf(child, ids);
      if (ref !== undefined) here.add(ref);
      for (const id of here) {
        let a = out.get(id);
        if (!a) { a = { paragraph: node, text: '' }; out.set(id, a); }
        if (child.isText && !isDeleted(child) && child.marks.some((m) => m.type.name === 'comment' && String(m.attrs['id']) === id)) a.text += child.text;
      }
      return true;
    });
    return false;
  });
  return out;
}

/** Plain text as a Markdown paragraph that reads back as the same text: Markdown's characters escaped, and a block marker at its start. */
export function escapeMarkdown(text: string): string {
  return text.replace(/[\\*_~`[\]<>]/g, (c) => `\\${c}`)
    .replace(/^(\s*)(#{1,6}(?=\s|$)|[-+](?=\s)|\d+(?=[.)]\s))/, (_m, space: string, marker: string) => `${space}\\${marker}`);
}

/** A comment's text as Markdown: its paragraphs, each escaped. */
function commentMarkdown(room: ToolRoom, thread: CommentThread): string {
  const doc = room.session.stories.currentDoc(thread.record);
  const paragraphs: string[] = [];
  doc.forEach((block) => { if (block.type.name === 'paragraph') paragraphs.push(escapeMarkdown(block.textContent)); });
  return paragraphs.join('\n\n');
}

// ---------------------------------------------------------------- the tools

// ---------------------------------------------------------------- the text model's readings

/** A table's style, from its `w:tbl`. */
function tableStyleOf(block: TextBlock): string {
  return ((block.source as wml.Tbl | undefined)?.tblPr as { tblStyle?: { val?: string } } | undefined)?.tblStyle?.val ?? '';
}

/** A block's paragraph id: its `w:p`'s (a joined block's first, REQ-405), as the contract spells it. */
function blockParaId(block: TextBlock): string | undefined {
  const p = block.source as wml.P | undefined;
  return p && (p as { TYPE_NAME?: string }).TYPE_NAME === 'org_docx4j_wml.P' ? paraIdOf(p) : block.paraId?.toUpperCase();
}

const KINDS: Record<TextBlock['kind'], string> = { PARAGRAPH: 'paragraph', TABLE: 'table', CHART: 'chart' };

/**
 * One block as `outline` lists it: the block model trimmed (ED-008 section 3.1), its text cut. The text is the
 * block's in the view (its text runs: no note number, no label), what find searches and a span counts in; under
 * markup, the block's line as `extract_text` writes it, so that the marks say what is inserted and what deleted.
 */
function outlineEntry(result: TextResult, block: TextBlock): Record<string, unknown> {
  const kind = KINDS[block.kind];
  const paraId = blockParaId(block);
  const text = result.options.view === 'markup' ? PlainText.of(resultOf(result, [block])).replace(/\n+$/, '') : block.text;
  const entry: Record<string, unknown> = { address: block.address, ...(paraId ? { paraId } : {}), style: kind === 'table' ? tableStyleOf(block) : block.style ?? 'Normal', kind, text: text.slice(0, 80) };
  if (kind === 'table') entry['rows'] = block.rows.length;
  if (block.story !== 'MAIN') entry['story'] = block.story.toLowerCase();
  if (block.joined.length > 0) entry['joined'] = [...block.joined];
  if (block.hidden) entry['hidden'] = true;
  if (block.placeholder) entry['placeholder'] = true;
  if (block.decorative) entry['decorative'] = true;
  if (block.revisions > 0) entry['revisions'] = block.revisions;
  if (block.numbering) entry['numbering'] = { level: block.numbering.level, ...(block.numbering.format ? { format: block.numbering.format } : {}), ...(block.numbering.value !== undefined ? { value: block.numbering.value } : {}) };
  if (block.listLabel) entry['list_label'] = block.listLabel;
  return entry;
}

/** A paragraph block's text as find searches it (its text runs and field codes, the block's `text`), each run's place in it. */
function searchTextOf(block: TextBlock): { text: string; segments: SearchSegment[]; pieces: { run: TextRun; start: number }[] } {
  let text = '';
  const segments: SearchSegment[] = [];
  const pieces: { run: TextRun; start: number }[] = [];
  for (const run of block.runs) {
    if (run.kind === 'IMAGE') { segments.push({ text: '', start: text.length, object: true, item: { name: { localPart: 'drawing' } } }); continue; }
    if (run.kind !== 'TEXT' && run.kind !== 'FIELD_CODE') continue;
    if (run.text === '') continue;
    pieces.push({ run, start: text.length });
    segments.push({ text: run.text, start: text.length, item: { name: { localPart: 't' } } });
    text += run.text;
  }
  return { text, segments, pieces };
}

/**
 * A hit's span in the paragraph's text as the edit tools count it (section 2.3), under the accepted view alone
 * (P2): the hit's runs all in one paragraph, each placed in its `Paragraph.text` (`officeOffset`), one after
 * another there. Undefined for a hit in text `Paragraph.text` does not hold or across two paragraphs.
 */
function spanOf(b: TreeBindings, pieces: readonly { run: TextRun; start: number }[], start: number, end: number): { address: string; paraId?: string; offset: number; length: number } | undefined {
  if (end <= start) return undefined;
  const covered = pieces.filter((p) => p.start < end && p.start + p.run.text.length > start);
  const first = covered[0];
  if (!first || first.run.officeOffset === undefined) return undefined;
  const paragraph = first.run.source?.paragraph;
  if (!paragraph) return undefined;
  for (let i = 1; i < covered.length; i++) {
    const prev = covered[i - 1]!.run;
    const next = covered[i]!.run;
    if (next.source?.paragraph !== paragraph || next.officeOffset === undefined || next.officeOffset !== prev.officeOffset! + prev.text.length) return undefined;
  }
  const address = b.body.addressOf({ element: { value: paragraph } as unknown as Element, container: [] });
  if (address === undefined) return undefined;
  const paraId = paraIdOf(paragraph);
  return { address, ...(paraId ? { paraId } : {}), offset: first.run.officeOffset + (start - first.start), length: end - start };
}

/** The ordinal an address names in a reading: one the model has (a note's, a text box's among them), else the body's resolution. */
function addressIn(b: TreeBindings, result: TextResult, input: AddressInput, warnings: ReadWarning[]): string {
  if (typeof input === 'string' && blocksAt(result, input).length > 0) return input;
  const { block, warnings: w } = resolveBlock(b.body, input);
  for (const message of w) warnings.push({ kind: 'address.first_match', message });
  return b.body.addressOf(block) ?? String(input);
}

/** What a reading at an address (or of the whole document) serialises: the result itself, or one of the blocks there. */
function chosenAt(b: TreeBindings, result: TextResult, address: AddressInput | undefined, warnings: ReadWarning[]): { result: TextResult; at?: string; whole: boolean } {
  if (address === undefined) return { result, whole: true };
  const at = addressIn(b, result, address, warnings);
  const blocks = blocksAt(result, at);
  if (blocks.length === 0) warnings.push({ kind: 'address.no_text', message: `the block at ${at} has no text in this view and scope` });
  return { result: resultOf(result, blocks), at, whole: false };
}

/** The address comment before a block's Markdown (section 2.6). */
function marker(address: string, paraId: string | undefined): string {
  return `<!-- ${address}${paraId ? ` paraId=${paraId}` : ''} -->`;
}

/** The tool's own arguments, beside the text model's options. */
const OWN = {
  outline: ['document', 'budget', 'max_chars'],
  find: ['document', 'text', 'xpath', 'match_case', 'match_whole_word', 'match_wildcards', 'match_regexp', 'limit', 'context'],
  read: ['document', 'address', 'max_chars', 'output_path', 'overwrite'],
  extract_text: ['document', 'address', 'max_chars', 'output_path', 'overwrite'],
} as const;

const VIEWS_NOTE = 'Reads through the text model the other docx4j servers share, with the text extraction specification\'s options (view, include and the rest; an option the engine does not implement is arguments.invalid).';

// ---------------------------------------------------------------- the tools

export function registerReadTools(server: McpServer, host: ToolHost): void {
  server.registerTool('outline', {
    title: 'Outline',
    description: `The document's blocks in reading order, at every depth (a table, then its cells' paragraphs; text boxes after their anchor; notes after the body; headers, footers and comments where included), each with its address, paragraph id, style, kind and the first 80 characters of its text in the view (accepted by default). A paragraph joined to the next across a deleted mark (accepted) is one block at its first paragraph's address, joined naming the others; the edit tools take each paragraph by its own address, and if_text is its own text. Start here: the addresses are what the other tools take. ${VIEWS_NOTE}`,
    inputSchema: z.object({
      document: documentArg,
      budget: z.number().int().min(1).optional().describe('how many blocks at most (default 400)'),
      max_chars: z.number().int().min(100).optional(),
      ...SCOPE_SHAPE,
    }).passthrough(),
  }, (args) => answer(() => inRoom(host, args.document, (room) => {
    const options = textOptionsOf('outline', args, OWN.outline);
    return readText(room, options, (result) => {
      const budget = args.budget ?? 400;
      const blocks: Record<string, unknown>[] = [];
      let more = false;
      for (const block of blocksIn(result.blocks)) {
        if (blocks.length >= budget) { more = true; break; }
        blocks.push(outlineEntry(result, block));
      }
      const cap = Math.min(args.max_chars ?? host.maxInlineChars, host.maxInlineChars);
      let truncated = more;
      while (blocks.length > 0 && JSON.stringify(blocks).length > cap) { blocks.pop(); truncated = true; }
      return ok({ blocks, truncated, ...(options.warnings && result.warnings.length > 0 ? { warnings: warningsOf(result) } : {}) });
    });
  })));

  server.registerTool('find', {
    title: 'Find',
    description: `Every match of text (Word's search: match_case, match_whole_word, match_wildcards, match_regexp) in the paragraphs' text in the view (accepted by default), each with its paragraph's address, a snippet, and under accepted alone the offset and length in the paragraph's text (a span for set_format, delete, comment; none for a hit in a note, a header, a comment or across a joined paragraph's mark); or, with xpath, the paragraphs and tables an XPath over the document's XML selects. ${VIEWS_NOTE}`,
    inputSchema: z.object({
      document: documentArg,
      text: z.string().min(1).optional(),
      xpath: z.string().min(1).optional(),
      ...searchShape,
      limit: z.number().int().min(1).optional().describe('default 100'),
      context: z.number().int().min(0).optional().describe('characters of snippet each side (default 40)'),
      ...SCOPE_SHAPE,
    }).passthrough(),
  }, (args) => answer(() => inRoom(host, args.document, async (room) => {
    if ((args.text === undefined) === (args.xpath === undefined)) throw new ToolError('arguments.invalid', 'give text or xpath, one of them', 'text searches the text; xpath selects in the XML', { argument: 'text' });
    const limit = args.limit ?? 100;
    const context = args.context ?? 40;
    const hits: Record<string, unknown>[] = [];
    let total = 0;
    if (args.text !== undefined) {
      const query = args.text;
      const options = textOptionsOf('find', args, OWN.find);
      const search = searchOptions(args);
      const accepted = options.view === 'accepted';
      return readText(room, options, (result, b) => {
        for (const block of blocksIn(result.blocks)) {
          if (block.kind !== 'PARAGRAPH') continue;
          const { text, segments, pieces } = searchTextOf(block);
          if (text === '') continue;
          let matches;
          try { matches = officeMatches(segments, query, search); } catch (e) { throw new ToolError('arguments.invalid', `the search does not parse: ${(e as Error).message}`, 'check the pattern', { argument: 'text' }); }
          for (const m of matches) {
            total++;
            if (hits.length >= limit) continue;
            const from = Math.max(0, m.start - context);
            const to = Math.min(text.length, m.end + context);
            const snippet = `${from > 0 ? '…' : ''}${text.slice(from, to)}${to < text.length ? '…' : ''}`;
            const span = accepted ? spanOf(b, pieces, m.start, m.end) : undefined;
            const paraId = span?.paraId ?? blockParaId(block);
            hits.push({ address: span?.address ?? block.address, ...(paraId ? { paraId } : {}), ...(span ? { offset: span.offset, length: span.length } : {}), snippet, ...(block.story !== 'MAIN' ? { story: block.story.toLowerCase() } : {}) });
          }
        }
        return ok({ hits, total, truncated: total > hits.length, ...(accepted ? {} : { spans: `none under view ${options.view}: offset and length count in the accepted text` }) });
      });
    }
    for (const name of Object.keys(SCOPE_SHAPE)) {
      if ((args as Record<string, unknown>)[name] !== undefined) throw new ToolError('arguments.invalid', `xpath selects in the document's XML, and ${name} is an option of the text`, 'drop it, or search with text', { argument: name });
    }
    textOptionsOf('find', args, OWN.find);
    return room.session.readTree(async (b) => {
      // An XPath over the tree as it is, never handed to the part (ED-005 section 3.5).
      const snapshot = await bindSnapshot(el.document(b.document) as never);
      const engine = b.pkg.xpathEngine;
      await engine.ready(snapshot.root);
      let nodes;
      try { nodes = engine.select(args.xpath!, snapshot.root, { ...XPATH_NAMESPACES }); } catch (e) { throw new ToolError('arguments.invalid', `the XPath does not parse: ${(e as Error).message}`, 'XPath 1.0 over the document\'s XML, w: bound to WordprocessingML', { argument: 'xpath' }); }
      const seen = new Set<object>();
      const addresses: string[] = [];
      for (const bound of boundsOf(nodes, snapshot)) {
        // The paragraph or table the selected object is, or is in: the first the body addresses. A text box's
        // paragraphs are addressed through their anchor (`body/16/textbox:0/0`, the engine's Choice branch, core-ts
        // 0.6.0 CR-012 D4); the Fallback's twin of the same box is unaddressed, so its climb reaches the paragraph
        // holding the box, and that hit is folded into the box's own (ED-008 section 8.8 item 15).
        let v = bound.object as { TYPE_NAME?: string; PARENT?: object } | undefined;
        let element: Element | undefined;
        let address: string | undefined;
        let unaddressed = false;
        for (let depth = 0; v && depth < 64 && address === undefined; depth++, v = v.PARENT as typeof v) {
          if (v.TYPE_NAME !== 'org_docx4j_wml.P' && v.TYPE_NAME !== 'org_docx4j_wml.Tbl') continue;
          element = { value: v } as unknown as Element;
          address = b.body.addressOf({ element, container: [] });
          if (address === undefined) unaddressed = true;
        }
        if (!element || address === undefined || seen.has(element.value as object)) continue;
        if (unaddressed && addresses.some((a) => a.startsWith(`${address}/textbox:`))) continue;
        seen.add(element.value as object);
        addresses.push(address);
        const block = element.value as { TYPE_NAME?: string };
        total++;
        if (hits.length >= limit) continue;
        const isParagraph = block.TYPE_NAME === 'org_docx4j_wml.P';
        const text = isParagraph ? new Paragraph(element as Element<wml.P>, [], b.body).text : new Table(element as Element<wml.Tbl>, [], b.body).text;
        const paraId = isParagraph ? paraIdOf(block as wml.P) : undefined;
        hits.push({ address, ...(paraId ? { paraId } : {}), kind: isParagraph ? 'paragraph' : 'table', snippet: text.length > 2 * context + 1 ? `${text.slice(0, 2 * context)}…` : text });
      }
      return ok({ hits, total, truncated: total > hits.length });
    });
  })));

  server.registerTool('read', {
    title: 'Read as Markdown',
    description: `The document, or the block at address, as Markdown in the view (accepted by default; CriticMarkup under markup), with each block's address in an HTML comment before it (<!-- body/3 paraId=5A2B1C3D -->, the paragraph id where the file has one). The Markdown is the text model's: GFM tables (a cell's paragraphs <br> apart), notes as [^n] with their definitions after the body, headings from outline levels, metadata as YAML front matter under metadata: true. ${VIEWS_NOTE}`,
    inputSchema: z.object({
      document: documentArg,
      address: addressSchema.optional(),
      max_chars: z.number().int().min(100).optional(),
      output_path: z.string().min(1).optional().describe('where a result past the inline cap is written'),
      overwrite: z.boolean().optional(),
      ...SCOPE_SHAPE,
      ...SERIALISATION_SHAPE,
    }).passthrough(),
  }, (args) => answer(async () => {
    const read = await inRoom(host, args.document, (room) => {
      const options = textOptionsOf('read', args, OWN.read);
      return readText(room, options, async (result, b) => {
        const warnings: ReadWarning[] = options.warnings ? warningsOf(result) : [];
        const chosen = chosenAt(b, result, args.address, warnings);
        // A top-level block keeps its place in the whole export (a list item its number); a cell's paragraph is written alone.
        const nested = !chosen.whole && chosen.result.blocks.some((x) => !result.blocks.includes(x));
        const exported = await MarkdownExporter.blocks(nested || chosen.whole ? chosen.result : result, b.textPackage);
        const wanted = new Set(chosen.result.blocks);
        const blocks: { address: string; paraId?: string; style: string; kind: string; markdown: string }[] = [];
        let markdown = chosen.whole ? exported.frontMatter : '';
        for (const { block, markdown: md } of exported.blocks) {
          if (!wanted.has(block) || md.trim() === '') continue;
          const paraId = blockParaId(block);
          const kind = KINDS[block.kind];
          blocks.push({ address: block.address, ...(paraId ? { paraId } : {}), style: kind === 'table' ? tableStyleOf(block) : block.style ?? 'Normal', kind, markdown: md.replace(/\n+$/, '') });
          markdown += `${marker(block.address, paraId)}\n${md}`;
        }
        return { markdown: markdown.replace(/\n+$/, ''), blocks, warnings };
      });
    });
    const out = await capped(host, read.markdown, args.max_chars, args.output_path, args.overwrite);
    const structured: Record<string, unknown> = { ...(out.text !== undefined ? { markdown: out.text, ...(out.truncated ? {} : { blocks: read.blocks }) } : { path: out.path }), warnings: read.warnings, truncated: out.truncated, chars: out.chars };
    return ok(structured, out.text !== undefined ? `${out.text}${out.truncated ? `\n\n(truncated: ${out.chars} characters in all)` : ''}${read.warnings.length > 0 ? `\n\nwarnings: ${read.warnings.map(warningText).join('; ')}` : ''}` : `Written to ${out.path} (${out.chars} characters)`);
  }));

  server.registerTool('extract_text', {
    title: 'Extract plain text',
    description: `The document, or the block at address, as plain text in the view (accepted by default): one line per paragraph, a table's rows a line each, [n] at a note's reference and the notes after the body, CriticMarkup under markup ({++inserted++}, {--deleted--}, {>>comment<<}). The vocabulary's extract_text, shared by name with docx4j-mcp and the Python server. ${VIEWS_NOTE}`,
    inputSchema: z.object({
      document: documentArg,
      address: addressSchema.optional(),
      max_chars: z.number().int().min(100).optional(),
      output_path: z.string().min(1).optional().describe('where a result past the inline cap is written'),
      overwrite: z.boolean().optional(),
      ...SCOPE_SHAPE,
      ...SERIALISATION_SHAPE,
    }).passthrough(),
  }, (args) => answer(async () => {
    const read = await inRoom(host, args.document, (room) => {
      const options = textOptionsOf('extract_text', args, OWN.extract_text);
      return readText(room, options, (result, b) => {
        const warnings: ReadWarning[] = options.warnings ? warningsOf(result) : [];
        const chosen = chosenAt(b, result, args.address, warnings);
        return { text: PlainText.of(chosen.result), warnings };
      });
    });
    const out = await capped(host, read.text, args.max_chars, args.output_path, args.overwrite);
    const structured: Record<string, unknown> = { ...(out.text !== undefined ? { text: out.text } : { path: out.path }), warnings: read.warnings, truncated: out.truncated, chars: out.chars };
    return ok(structured, out.text !== undefined ? `${out.text}${out.truncated ? `\n\n(truncated: ${out.chars} characters in all)` : ''}${read.warnings.length > 0 ? `\n\nwarnings: ${read.warnings.map(warningText).join('; ')}` : ''}` : `Written to ${out.path} (${out.chars} characters)`);
  }));

  server.registerTool('get_comments', {
    title: 'Comments',
    description: 'The comment threads (or those anchored at address), each with its id, author, date, whether it is resolved, the address and words it is anchored on, its text as Markdown, and its replies.',
    inputSchema: { document: documentArg, address: addressSchema.optional() },
  }, (args) => answer(() => inRoom(host, args.document, async (room) => {
    await room.session.stories.comments();
    return room.session.readTree((b) => {
      const scope = scopeOf(b, args.address);
      const anchors = commentAnchorsIn(b, room.session.ctx.ids);
      const view = (t: CommentThread): Record<string, unknown> => {
        const anchor = anchors.get(t.id);
        return {
          id: t.id, author: t.author, date: t.date?.toISOString() ?? null, resolved: t.done,
          address: anchor ? addressOfNode(b, anchor.paragraph) ?? null : null, anchor: anchor?.text ?? '',
          text: commentMarkdown(room, t), replies: t.replies.map(view),
        };
      };
      const threads = room.session.stories.commentThreads(b.doc).map(view).filter((t) => within((t['address'] as string | null) ?? undefined, scope));
      return ok({ threads });
    });
  })));

  server.registerTool('get_tracked_changes', {
    title: 'Tracked changes',
    description: 'The tracked changes as the review panel lists them: each with its key (what accept_tracked_changes and reject_tracked_changes take), kind, author, date, address and text. A move is one entry; a revision Word shows that the editor does not draw is marked shown_in_word_only.',
    inputSchema: { document: documentArg, address: addressSchema.optional(), author: z.string().optional() },
  }, (args) => answer(() => inRoom(host, args.document, (room) => room.session.readTree((b) => {
    const scope = scopeOf(b, args.address);
    const changes = listChanges(room, b).filter((c) => within(c.address ?? undefined, scope) && (args.author === undefined || c.author === args.author));
    return ok({ changes });
  }))));
}

