// SPDX-License-Identifier: GPL-3.0-or-later
//
// Paste from Word (ED-001 section 11 item 6; ED-003 section 8.3, step F2; section 9 item 33).
// Browsers withhold Word's OOXML flavour, so what arrives is `text/html` in Word's dialect:
// `<h1>` to `<h6>` for the heading styles, `<p class=MsoNormal>` and `MsoTitle`, `<b>`, `<i>`,
// `<u>`, `<s>`, `<sup>`, `<sub>` and their `span` style forms, `<br>`, a tab as a span with
// `mso-tab-count`, lists as `<ul>` and `<ol>` (Word 15) or as `MsoListParagraph` paragraphs
// with `mso-list:l0 level1 lfo1` and a label span marked `mso-list:Ignore`, tables as
// `<table class=MsoTableGrid>`, and a great deal else (fonts, sizes, colours, spacing,
// `mso-` properties). The converter keeps what the outline names and, since F5 (section 8.6),
// the direct formatting an element states: alignment, indents, space before, the font, size,
// colour and highlight of a run; it drops the rest and counts it, so the status line can say
// what was dropped. `margin-bottom` and `line-height` stay dropped: Word writes them on every
// paragraph from the source's style, and keeping them would turn style inheritance into direct
// formatting. Since ED-005 section 12.67 a link to the web or to mail is kept, in Word's look, its
// relationship made when the paste lands, and a terminal's HTML (Konsole's "Copy text as HTML":
// coloured spans in a monospace font, no paragraphs) is pasted as its plain text. Since section 12.121 a plain
// paste that is one web address goes in as a link whose text is the address, and since section 12.127 a
// picture on the clipboard (a screenshot, a browser's Copy image) goes in as Insert picture puts one. Pure over its
// inputs: a DOM walk to an intermediate form the tests pin, then nodes of the schema once the list
// definitions exist.
import type { Schema, Node as PMNode, Mark } from 'prosemirror-model';
import { Fragment, Slice } from 'prosemirror-model';
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { W_NS } from '@docx4j/core-ts';
import { documentStyles, storyOf, looksLikeMarkdown, webAddress, insertLinkedText, HYPERLINK_STYLE, commandNamed, taggedDispatch, type EditorDocument, type StyleEntry } from '@docx4j/editor-model';
import { nextOffer, type MarkdownOffer } from './markdown.js';
import { EditorClipboardSerializer, editorSliceOf, pasteSlice, transformPastedSlice } from './clipboard.js';
import { t, hasMessage } from '../app/i18n.js';

// ---------------------------------------------------------------- the intermediate form

// The converter's pure half is the engine's since core-ts 0.4.0 (ED-005 proposal 44, section 1.11): the markup to
// paste blocks, the clipboard's fragment, the terminal test and the style lookup come from `@docx4j/core-ts/html`;
// what stays here turns the blocks into the schema's nodes and pastes them through a view.
export { convertHtml, clipboardMarkup, isTerminalHtml, styleLookupOf } from '@docx4j/core-ts/html';
export type { PasteRun, PasteInline, PasteList, PasteParagraph, PasteCell, PasteTable, PasteBlock, PasteReport, PasteResult, StyleLookup } from '@docx4j/core-ts/html';
import { convertHtml, clipboardMarkup, isTerminalHtml, styleLookupOf, type PasteRun, type PasteInline, type PasteParagraph, type PasteTable, type PasteBlock, type PasteReport, type PasteResult, type StyleLookup } from '@docx4j/core-ts/html';

export async function styleLookupFor(session: EditorDocument): Promise<StyleLookup> {
  const resolver = session.pkg.propertyResolverOrUndefined;
  return styleLookupOf(await documentStyles(session), resolver?.getStyle('TableGrid') ? ['TableGrid'] : []);
}

// ---------------------------------------------------------------- the clipboard text

/** Word's `CF_HTML` header (`Version:`, `StartHTML:` and the rest) precedes `<html` in a saved capture; a browser hands the markup alone. */
export function isEditorClipboard(html: string): boolean {
  return html.includes('data-pm-slice');
}

// ---------------------------------------------------------------- the walk

export type ListIds = Map<string, string>;

/** The relationships a paste's links need: a relationship id per address. */
export type LinkIds = Map<string, string>;

function marksFor(run: PasteRun, schema: Schema, links: LinkIds): Mark[] {
  const marks: Mark[] = [];
  const rId = run.link !== undefined ? links.get(run.link) : undefined;
  // A link's text takes Word's look, the `Hyperlink` character style in place of any other, as `range.hyperlink` gives it.
  if (rId !== undefined) marks.push(schema.marks['link']!.create({ rId }), schema.marks['runStyle']!.create({ id: HYPERLINK_STYLE }));
  else if (run.style) marks.push(schema.marks['runStyle']!.create({ id: run.style }));
  if (run.bold) marks.push(schema.marks['bold']!.create());
  if (run.italic) marks.push(schema.marks['italic']!.create());
  if (run.underline) marks.push(schema.marks['underline']!.create({ val: 'Single' }));
  if (run.strike) marks.push(schema.marks['strike']!.create({ double: false }));
  if (run.vertAlign) marks.push(schema.marks['vertAlign']!.create({ val: run.vertAlign }));
  // A link's colour is its style's, as `range.hyperlink` leaves no direct colour on a link.
  const color = rId !== undefined ? undefined : run.color;
  if (run.fontName !== undefined || run.fontSize !== undefined || color !== undefined || run.highlight !== undefined) {
    marks.push(schema.marks['font']!.create({ name: run.fontName ?? null, size: run.fontSize ?? null, color: color ?? null, highlight: run.highlight ?? null }));
  }
  return marks;
}

function paragraphNode(p: PasteParagraph, schema: Schema, lists: ListIds, links: LinkIds): PMNode {
  const content: PMNode[] = [];
  for (const inline of p.inlines) {
    if ('br' in inline) content.push(schema.nodes['hardBreak']!.create());
    else if ('tab' in inline) content.push(schema.nodes['tab']!.create());
    else if (inline.text !== '') content.push(schema.text(inline.text, marksFor(inline, schema, links)));
  }
  const numId = p.list ? lists.get(p.list.key) : undefined;
  return schema.nodes['paragraph']!.create({
    pStyle: p.style,
    ...(p.list && numId ? { numPr: { numId, ilvl: p.list.level } } : {}),
    ...(p.jc ? { jc: p.jc } : {}),
    ...(p.ind ? { ind: { start: p.ind.start ?? null, end: p.ind.end ?? null, firstLine: p.ind.firstLine ?? null, hanging: p.ind.hanging ?? null } } : {}),
    ...(p.spacing ? { spacing: { before: p.spacing.before, after: null, line: null, lineRule: null } } : {}),
  }, content);
}

function tableNode(t: PasteTable, schema: Schema, lists: ListIds, links: LinkIds): PMNode {
  const columns = Math.max(1, ...t.rows.map((r) => r.reduce((n, c) => n + c.colspan, 0)));
  // Column widths from the first row that says, else the section's text width shared equally.
  const widths: number[] = Array.from({ length: columns }, () => Math.floor(9026 / columns));
  const first = t.rows[0];
  if (first && first.every((c) => c.width !== undefined && c.colspan === 1) && first.length === columns) first.forEach((c, i) => { widths[i] = c.width!; });
  const rows = t.rows.map((cells, r) => schema.nodes['table_row']!.create(r === 0 && t.headerRow ? { trPr: HEADER_ROW_TRPR } : null, cells.map((c, i) => {
    const colwidth = Array.from({ length: c.colspan }, (_, k) => (widths[i + k] ?? widths[0]!) / 15);
    const blocks = c.blocks.map((b) => (b.kind === 'table' ? tableNode(b, schema, lists, links) : paragraphNode(b, schema, lists, links)));
    return schema.nodes['table_cell']!.create({ colspan: c.colspan, rowspan: c.rowspan, colwidth }, blocks);
  })));
  const tblPr: Record<string, unknown> = { TYPE_NAME: 'org_docx4j_wml.TblPr', tblW: { w: 0, type: 'auto' } };
  if (t.style) tblPr['tblStyle'] = { val: t.style };
  return schema.nodes['table']!.create({ tblPr, tblGrid: widths }, rows);
}

/** A first row repeated at the top of each page, as the header-row command writes it. */
const HEADER_ROW_TRPR = { TYPE_NAME: 'org_docx4j_wml.TrPr', cnfStyleOrDivIdOrGridBefore: [{ name: { namespaceURI: W_NS, localPart: 'tblHeader' }, value: { TYPE_NAME: 'org_docx4j_wml.BooleanDefaultTrue' } }] };

/** The blocks as nodes, the list keys resolved to the definitions made for them and the links to their relationships. */
export function toNodes(blocks: readonly PasteBlock[], schema: Schema, lists: ListIds, links: LinkIds = new Map()): PMNode[] {
  return blocks.map((b) => (b.kind === 'table' ? tableNode(b, schema, lists, links) : paragraphNode(b, schema, lists, links)));
}

// ---------------------------------------------------------------- the paste itself

export interface PasteOptions {
  /** The status line's report after a paste. */
  onPaste?: (report: PasteReport, text: string) => void;
  /**
   * A plain paste that reads as Markdown made an offer in `view` (the status line's Format as
   * Markdown), or the offer of `from` went (null): this person edited, the view went.
   */
  onMarkdownOffer?: (view: EditorView | null, from: EditorView | undefined) => void;
}

/** The message a dropped kind is reported under (the report's keys are pinned in the conversions, so they stay English). */
const DROPPED_MESSAGES: Record<string, string> = { 'tables in a note or header': 'tablesInStory', 'styles the defaults lack': 'stylesMissing', HTML: 'html' };

/** What was dropped, by kind, as a list for the status bar ('' for nothing). */
export function describeDropped(dropped: Record<string, number>): string {
  return Object.entries(dropped).map(([kind, count]) => {
    const id = DROPPED_MESSAGES[kind] ?? kind;
    return hasMessage(`paste.dropped.${id}`) ? t(`paste.dropped.${id}`, { count }) : t('paste.dropped.other', { kind, count });
  }).join(', ');
}

/** One line for the status bar, in the application's language. */
export function describePaste(report: PasteReport): string {
  const kept: string[] = [];
  if (report.paragraphs > 0) kept.push(t('paste.paragraphs', { count: report.paragraphs }));
  if (report.lists > 0) kept.push(t('paste.lists', { count: report.lists }));
  if (report.tables > 0) kept.push(t('paste.tables', { count: report.tables }));
  const dropped = describeDropped(report.dropped);
  const parts = [kept.length > 0 ? t('paste.pasted', { what: kept.join(', ') }) : t('paste.pastedNothing')];
  if (report.stylesWanted.length > 0) parts.push(t('paste.stylesAdded', { count: report.stylesWanted.length, styles: report.stylesWanted.join(', ') }));
  if (dropped !== '') parts.push(t('paste.dropped', { what: dropped }));
  return parts.join(t('paste.separator'));
}

/**
 * Paste clipboard HTML into a view as one transaction (one undo step): the conversion, then
 * `pasteConverted`.
 */
export async function pasteHtml(view: EditorView, session: EditorDocument, html: string): Promise<PasteReport> {
  return (await pasteConverted(view, session, convertHtml(html, await styleLookupFor(session))))!;
}

/**
 * Converted blocks into a view as one transaction (one undo step): the styles, list definitions
 * and link relationships the paste needs (part writes through the session, not undone, as C2
 * decided), then the nodes replacing the selection, or `range`, with ProseMirror's own fitting.
 * In a story (a note, a header) tables are dropped and counted, since the notes hold none. A
 * `range` is asked for when the transaction is made, after the parts; null from it pastes nothing,
 * and the answer is null.
 */
/** What a conversion needs of the document before its nodes are built: the styles spliced in, a list definition per key, a relationship per link; the blocks kept (a note refuses tables) and the report. */
export async function prepareConverted(view: EditorView, session: EditorDocument, converted: PasteResult): Promise<{ kept: PasteBlock[]; lists: ListIds; links: LinkIds; report: PasteReport }> {
  const { blocks, report } = converted;
  let kept = blocks;
  if (storyOf(view.state) !== undefined && blocks.some((b) => b.kind === 'table')) {
    kept = blocks.filter((b) => b.kind !== 'table');
    report.dropped['tables in a note or header'] = report.tables;
    report.tables = 0;
  }
  // A built-in style the document lacks (a Heading 2 pasted into a document that never had one)
  // is spliced in from docx4j's defaults, as a note's styles are (section 11.11); one the defaults
  // lack either leaves its paragraphs plain.
  if (report.stylesWanted.length > 0) {
    await session.ensureStyles(report.stylesWanted);
    const resolver = session.pkg.propertyResolverOrUndefined;
    const missing = new Set(report.stylesWanted.filter((id) => !resolver?.getStyle(id)));
    report.stylesWanted = report.stylesWanted.filter((id) => !missing.has(id));
    if (missing.size > 0) {
      report.dropped['styles the defaults lack'] = missing.size;
      const strip = (bs: readonly PasteBlock[]): void => {
        for (const b of bs) {
          if (b.kind === 'paragraph') {
            if (b.style && missing.has(b.style)) b.style = null;
            for (const i of b.inlines) if ('text' in i && i.style && missing.has(i.style)) delete i.style;
          } else for (const row of b.rows) for (const c of row) strip(c.blocks);
        }
      };
      strip(kept);
    }
  }
  // One definition per list key: bullets may share one; each numbered list restarts.
  const lists: ListIds = new Map();
  let firstNumbered: string | undefined;
  const keys = new Map<string, boolean>();
  const collect = (bs: readonly PasteBlock[]): void => { for (const b of bs) { if (b.kind === 'paragraph') { if (b.list) keys.set(b.list.key, b.list.bullet); } else for (const row of b.rows) for (const c of row) collect(c.blocks); } };
  collect(kept);
  for (const [key, bullet] of keys) {
    if (bullet) lists.set(key, await session.newList(true, view.state.doc));
    else if (firstNumbered === undefined) { firstNumbered = await session.newList(false, view.state.doc); lists.set(key, firstNumbered); }
    else lists.set(key, await session.restartList(firstNumbered));
  }
  // A relationship per address, and Word's link style, as `range.hyperlink`'s prepare half makes them.
  const links: LinkIds = new Map();
  const addresses = new Set<string>();
  const collectLinks = (bs: readonly PasteBlock[]): void => { for (const b of bs) { if (b.kind === 'paragraph') { for (const i of b.inlines) if ('text' in i && i.link !== undefined) addresses.add(i.link); } else for (const row of b.rows) for (const c of row) collectLinks(c.blocks); } };
  collectLinks(kept);
  // A note's or a header's part has relationships of its own, which a link there would need: its text is kept, the link counted.
  if (addresses.size > 0 && storyOf(view.state) !== undefined) {
    let count = 0;
    const unlink = (bs: readonly PasteBlock[]): void => { for (const b of bs) { if (b.kind === 'paragraph') { for (const i of b.inlines) if ('text' in i && i.link !== undefined) { delete i.link; count++; } } else for (const row of b.rows) for (const c of row) unlink(c.blocks); } };
    unlink(kept);
    report.dropped['links'] = (report.dropped['links'] ?? 0) + count;
    addresses.clear();
  }
  if (addresses.size > 0) {
    await session.ensureStyles([HYPERLINK_STYLE]);
    for (const address of addresses) links.set(address, session.addHyperlink(address));
  }
  return { kept: kept as PasteBlock[], lists, links, report };
}

export async function pasteConverted(view: EditorView, session: EditorDocument, converted: PasteResult, range?: () => { from: number; to: number } | null, meta?: Record<string, unknown>, onInserted?: (inserted: { from: number; to: number }) => void): Promise<PasteReport | null> {
  const { kept, lists, links, report } = await prepareConverted(view, session, converted);
  if ((view as { isDestroyed?: boolean }).isDestroyed) return report;
  const target = range ? range() : { from: view.state.selection.from, to: view.state.selection.to };
  if (target === null) return null;
  const nodes = toNodes(kept, view.state.schema, lists, links);
  if (nodes.length > 0) {
    const slice = Slice.maxOpen(Fragment.from(nodes));
    const tr = range ? replaceWholeBlocks(view.state.tr, target.from, target.to, slice) : view.state.tr.replaceSelection(slice);
    for (const [key, value] of Object.entries(meta ?? {})) tr.setMeta(key, value);
    view.dispatch(tr.scrollIntoView());
    // The range the blocks occupy after the replacement: the replaced range's edges mapped outwards.
    onInserted?.({ from: tr.mapping.map(target.from, -1), to: tr.mapping.map(target.to, 1) });
  }
  return report;
}

/**
 * A range replaced by blocks, a side of it on a paragraph's edge widened to the whole paragraph and
 * that side of the slice closed: the first block then keeps its own style (a heading pasted into an
 * empty paragraph stays a heading) where an open slice would pour it into the paragraph there.
 */
function replaceWholeBlocks(tr: Transaction, from: number, to: number, slice: Slice): Transaction {
  const $from = tr.doc.resolve(from);
  const $to = tr.doc.resolve(to);
  const startsBlock = $from.parent.isTextblock && $from.parentOffset === 0;
  const endsBlock = $to.parent.isTextblock && $to.parentOffset === $to.parent.content.size;
  const fitted = new Slice(slice.content, startsBlock ? 0 : slice.openStart, endsBlock ? 0 : slice.openEnd);
  return tr.replaceRange(startsBlock ? $from.before() : from, endsBlock ? $to.after() : to, fitted);
}

// ---------------------------------------------------------------- a pasted address

/**
 * A plain paste that is one web address goes in as a link, its text the address (Jason, 2026-11-12; ED-005 section 12.121):
 * the text starts with `http://` or `https://` (in any case) and holds no whitespace, a line end included.
 */
export function pastedAddress(text: string): string | undefined {
  return /^https?:\/\/\S+$/i.test(text) ? text : undefined;
}

/**
 * The address over the selection as a link in Word's look (`insertLinkedText`), one transaction, its relationship and the
 * `Hyperlink` style made first as `range.hyperlink`'s prepare half makes them (part writes, not undone). The caret is left
 * after the link, where what is typed next takes the formatting the address had before the link's look.
 */
export async function pasteAddress(view: EditorView, session: EditorDocument, address: string): Promise<void> {
  await session.ensureStyles([HYPERLINK_STYLE]);
  const rId = session.addHyperlink(address);
  if ((view as { isDestroyed?: boolean }).isDestroyed) return;
  const { tr, typing } = insertLinkedText(view.state, address, rId);
  view.dispatch(tr.setMeta('uiEvent', 'paste').scrollIntoView());
  if (view.state.selection.empty) view.dispatch(view.state.tr.setStoredMarks(typing));
}

// ---------------------------------------------------------------- a pasted picture

/** The types that go in as they are: `range.insertInlinePicture`'s list, which Word reads. */
const WORD_PICTURE_TYPES = new Set(['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/tiff', 'image/svg+xml', 'image/emf', 'image/x-emf', 'image/wmf', 'image/x-wmf']);

/**
 * The pictures a paste is (Jason, 2026-11-13; ED-005 section 12.127): the clipboard's image files, when the HTML beside them,
 * if any, holds no text. A screenshot comes as a file alone, a browser's Copy image as the file and the markup of one `<img>`.
 * HTML with text (a document's text and pictures, a spreadsheet's cells beside a bitmap of them) is pasted as HTML, its
 * pictures dropped and counted as before; the editor's own clipboard carries no file.
 */
export function clipboardPictures(files: readonly File[], html: string): File[] {
  const pictures = files.filter((f) => f.type.startsWith('image/'));
  if (pictures.length === 0) return [];
  if (html !== '' && (new DOMParser().parseFromString(html, 'text/html').body.textContent ?? '').trim() !== '') return [];
  return pictures;
}

/** A picture's bytes as Word reads them: as they are in a type Word reads, else (WebP, AVIF, an icon) decoded by the browser and encoded as PNG. */
async function pictureBytes(file: File): Promise<Uint8Array> {
  if (WORD_PICTURE_TYPES.has(file.type)) return new Uint8Array(await file.arrayBuffer());
  const bitmap = await createImageBitmap(file);
  try {
    const canvas = document.createElement('canvas');
    canvas.width = bitmap.width;
    canvas.height = bitmap.height;
    canvas.getContext('2d')!.drawImage(bitmap, 0, 0);
    const png = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/png'));
    if (png === null) throw new Error('not encoded as PNG');
    return new Uint8Array(await png.arrayBuffer());
  } finally {
    bitmap.close();
  }
}

function base64Of(bytes: Uint8Array): string {
  let binary = '';
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

/**
 * Pictures into a view at its selection, each through `range.insertInlinePicture` as Insert picture puts one: the image part
 * and its relationship made first, then the image node at its natural size scaled down to the text width, tracked when
 * tracking is on; the first replaces the selection and the next goes after it. Where the command is not enabled (a note, a
 * header, a selection that is not in text) nothing goes in. The message for the status line when a picture did not go in,
 * else null.
 */
export async function pastePictures(view: EditorView, pictures: readonly File[]): Promise<string | null> {
  const spec = commandNamed('range.insertInlinePicture')!;
  if (spec.isEnabled && !spec.isEnabled(view.state, undefined)) return t('paste.picture.notHere');
  const dispatch = taggedDispatch(spec.name, (tr) => view.dispatch(tr.setMeta('uiEvent', 'paste')));
  for (const file of pictures) {
    let base64: string;
    try {
      base64 = base64Of(await pictureBytes(file));
    } catch {
      return t('paste.picture.unreadable', { type: file.type });
    }
    if ((view as { isDestroyed?: boolean }).isDestroyed) return null;
    const prepared = await spec.prepare!(view.state, { base64 });
    if ((view as { isDestroyed?: boolean }).isDestroyed) return null;
    if (!spec.run(prepared)(view.state, dispatch, view)) return t('paste.picture.notHere');
  }
  return null;
}

/** The report a paste that failed is told with. */
const NOTHING_PASTED: PasteReport = { paragraphs: 0, lists: 0, tables: 0, dropped: {}, stylesWanted: [] };

export const pasteKey = new PluginKey<MarkdownOffer | null>('docx4j-paste');

/**
 * The paste handler as a plugin: Word's and any other HTML through the converter, the editor's own
 * clipboard and plain text left to ProseMirror. A plain text that reads as Markdown leaves an offer
 * in the plugin's state (`markdownOfferOf`), told to `onMarkdownOffer` as it comes and goes.
 */
export function pastePlugin(session: EditorDocument, options: PasteOptions = {}): Plugin<MarkdownOffer | null> {
  // The plain text of the paste under way, for the state to take from ProseMirror's own paste transaction.
  let pending: string | null = null;
  // A terminal's paste handed back to ProseMirror as its plain text: the HTML is not looked at again.
  let plainOnly = false;
  // The slice being copied, for the serializer to write beside ProseMirror's HTML (ED-010 section 3.1).
  let copied: Slice | null = null;
  const serializer = new EditorClipboardSerializer(session.schema, session.clipboardToken, () => copied);
  return new Plugin<MarkdownOffer | null>({
    key: pasteKey,
    state: {
      init: () => null,
      apply: (tr, value) => nextOffer(tr, value, pending),
    },
    view: (editorView) => ({
      update(view: EditorView, prev: EditorState): void {
        const now = pasteKey.getState(view.state) ?? null;
        const before = pasteKey.getState(prev) ?? null;
        if ((now === null) !== (before === null) || (now !== null && before !== null && now.text !== before.text)) options.onMarkdownOffer?.(now ? view : null, view);
      },
      destroy(): void { options.onMarkdownOffer?.(null, editorView); },
    }),
    props: {
      transformCopied: (slice) => { copied = slice; return slice; },
      clipboardSerializer: serializer,
      // A dragged copy, and a slice ProseMirror parsed itself, through the paste rules (ED-010 section 3.4).
      transformPasted: (slice, view) => transformPastedSlice(slice, view, session),
      handlePaste(view: EditorView, event: ClipboardEvent): boolean {
        const html = plainOnly ? '' : event.clipboardData?.getData('text/html') ?? '';
        // A picture on the clipboard, with no text beside it, goes in as Insert picture puts one (section 12.127).
        const pictures = plainOnly ? [] : clipboardPictures(Array.from(event.clipboardData?.files ?? []), html);
        if (pictures.length > 0) {
          pending = null;
          void pastePictures(view, pictures).then((message) => { if (message !== null) options.onPaste?.(NOTHING_PASTED, message); }, (e: unknown) => options.onPaste?.(NOTHING_PASTED, t('paste.failed', { error: (e as Error).message })));
          return true;
        }
        if (html === '') {
          const text = event.clipboardData?.getData('text/plain') ?? '';
          // One web address is a link (section 12.121). In a note or a header it stays text, as a link pasted there in HTML does
          // (the story's part would need a relationship of its own); over a cell or a node selection it is ProseMirror's paste.
          const address = pastedAddress(text);
          if (address !== undefined && storyOf(view.state) === undefined && view.state.selection instanceof TextSelection) {
            pending = null;
            void pasteAddress(view, session, address).catch((e: unknown) => options.onPaste?.(NOTHING_PASTED, t('paste.failed', { error: (e as Error).message })));
            return true;
          }
          pending = looksLikeMarkdown(text) ? text : null;
          queueMicrotask(() => { pending = null; });
          return false;
        }
        // The editor's own clipboard from this session: the slice rebuilt from its JSON (ED-010 section 3.2);
        // another session's falls through to ProseMirror's reading of the HTML.
        const own = editorSliceOf(html, session.clipboardToken);
        if (own !== null) {
          void pasteSlice(view, session, own).then((done) => { if (!done) view.pasteHTML(html); });
          return true;
        }
        if (isEditorClipboard(html)) return false;
        if (isTerminalHtml(html)) {
          // ProseMirror's own paste of the plain text (with the Markdown offer, when it reads as Markdown);
          // returning false here would paste ProseMirror's reading of the HTML instead.
          plainOnly = true;
          try { view.pasteText(event.clipboardData?.getData('text/plain') ?? '', event); } finally { plainOnly = false; }
          return true;
        }
        void pasteHtml(view, session, html).then((report) => options.onPaste?.(report, describePaste(report)), (e: unknown) => options.onPaste?.(NOTHING_PASTED, t('paste.failed', { error: (e as Error).message })));
        return true;
      },
    },
  });
}

/** The paste options among a surface's, as `exactOptionalPropertyTypes` wants them. */
export function pasteOptionsOf(options: PasteOptions): PasteOptions {
  return { ...(options.onPaste ? { onPaste: options.onPaste } : {}), ...(options.onMarkdownOffer ? { onMarkdownOffer: options.onMarkdownOffer } : {}) };
}

/** The offer a plain paste left, while it stands. */
export function markdownOfferOf(state: EditorState): MarkdownOffer | null {
  return pasteKey.getState(state) ?? null;
}

/** The state's paste plugin, for tests. */
export function pastePluginOf(state: EditorState): Plugin | undefined {
  return state.plugins.find((p) => p.spec.key === pasteKey);
}
