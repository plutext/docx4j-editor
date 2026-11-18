// SPDX-License-Identifier: GPL-3.0-or-later
//
// Markdown paste (ED-005 section 12.66; Jason's decision 2026-10-06). A plain-text paste stays plain;
// when its text reads as Markdown (`looksLikeMarkdown`), the status line offers "Format as Markdown",
// which replaces what the paste inserted with the Markdown read leniently (the model's `readMarkdown`,
// the agent's reading with code, quotes and the rest given a place) as one undoable step: Undo gives
// the plain text back. The offer is the paste plugin's state, the inserted range mapped through a
// peer's changes and withdrawn at this person's next edit; the replacement goes through
// `pasteConverted`, so the styles, lists and links it needs are made as a paste makes them, and
// tracking records it as any paste.
import type { EditorState, Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { ySyncPluginKey } from 'y-prosemirror';
import { readMarkdown, isDeleted, heldText, MARKDOWN_STYLES, type EditorDocument, type Inline, type MarkdownRead, type MdBlock } from '@docx4j/editor-model';
import { pasteConverted, styleLookupFor, type PasteBlock, type PasteCell, type PasteInline, type PasteParagraph, type PasteReport, type PasteResult, type StyleLookup } from './paste.js';

// ---------------------------------------------------------------- the conversion

/** Markdown's blocks as paste's: Word's built-in styles by the document's names, a style it lacks wanted, as a Word paste's headings are. */
export function markdownToPaste(read: MarkdownRead, lookup: StyleLookup): PasteResult {
  const report: PasteReport = { paragraphs: 0, lists: 0, tables: 0, dropped: { ...read.dropped }, stylesWanted: [] };
  const wanted = new Set<string>();
  const listKeys = new Set<string>();
  const builtIn = (name: string, id: string): string => {
    const have = lookup.idFor(name) ?? (lookup.hasId(id) ? id : undefined);
    if (have) return have;
    wanted.add(id);
    return id;
  };
  const heading = (level: number): string => lookup.idFor(`heading ${level}`) ?? (lookup.hasId(`Heading${level}`) ? `Heading${level}` : (wanted.add(`Heading${level}`), `Heading${level}`));
  const inlinesOf = (list: readonly Inline[]): PasteInline[] => {
    const out: PasteInline[] = [];
    for (const inline of list) {
      for (const piece of inline.text.split(/(\n|\t)/)) {
        if (piece === '\n') out.push({ br: true });
        else if (piece === '\t') out.push({ tab: true });
        else if (piece !== '') {
          out.push({
            text: piece,
            ...(inline.bold ? { bold: true } : {}), ...(inline.italic ? { italic: true } : {}), ...(inline.strike ? { strike: true } : {}),
            ...(inline.code ? { style: builtIn('HTML Code', MARKDOWN_STYLES.code) } : {}),
            ...(inline.link !== undefined ? { link: inline.link } : {}),
          });
        }
      }
    }
    return out;
  };
  const paragraphOf = (b: Extract<MdBlock, { kind: 'paragraph' }>): PasteParagraph => {
    report.paragraphs++;
    let style: string | null = null;
    if (b.heading !== undefined) style = heading(b.heading);
    else if (b.code) style = builtIn('HTML Preformatted', MARKDOWN_STYLES.codeBlock);
    else if (b.quote && !b.list) style = builtIn('Quote', MARKDOWN_STYLES.quote);
    if (b.list) listKeys.add(b.list.key);
    return { kind: 'paragraph', style, ...(b.list ? { list: { key: b.list.key, level: b.list.level, bullet: !b.list.ordered } } : {}), inlines: inlinesOf(b.inlines) };
  };
  const blocks: PasteBlock[] = read.blocks.map((b) => {
    if (b.kind === 'paragraph') return paragraphOf(b);
    report.tables++;
    const width = Math.max(1, ...b.rows.map((row) => row.length));
    const rows = b.rows.map((row) => Array.from({ length: width }, (_, i): PasteCell => {
      report.paragraphs++;
      return { blocks: [{ kind: 'paragraph', style: null, inlines: inlinesOf(row[i] ?? []) }], colspan: 1, rowspan: 1 };
    }));
    return { kind: 'table', rows, style: lookup.hasId('TableGrid') ? 'TableGrid' : null, headerRow: true };
  });
  report.lists = listKeys.size;
  report.stylesWanted = [...wanted];
  return { blocks, report };
}

// ---------------------------------------------------------------- the offer

/** What a plain paste inserted, while it can still be formatted: its range now, and the clipboard's text. */
export interface MarkdownOffer {
  from: number;
  to: number;
  text: string;
}

/** The transaction meta of a Format as Markdown, which ends the offer. */
export const FORMAT_MARKDOWN_META = 'docx4j-format-markdown';

/**
 * The offer after a transaction: made by the paste of `pending` (the plain text `handlePaste` saw
 * read as Markdown), mapped through a peer's change, withdrawn by any other change. The paste's own
 * step is the first: the tracker appends its steps after it (deleted text put back before the
 * insertion, which the mapping steps over).
 */
export function nextOffer(tr: Transaction, value: MarkdownOffer | null, pending: string | null): MarkdownOffer | null {
  if (tr.getMeta('uiEvent') === 'paste' && pending !== null && tr.docChanged) {
    let from = Infinity;
    let to = -Infinity;
    tr.steps[0]!.getMap().forEach((_a, _b, start, end) => { from = Math.min(from, start); to = Math.max(to, end); });
    if (!(from < to)) return null;
    const rest = tr.mapping.slice(1);
    return { from: rest.map(from, 1), to: rest.map(to, -1), text: pending };
  }
  if (value === null || !tr.docChanged) return value;
  if (tr.getMeta(FORMAT_MARKDOWN_META)) return null;
  if ((tr.getMeta(ySyncPluginKey) as { isChangeOrigin?: boolean } | undefined)?.isChangeOrigin) {
    const from = tr.mapping.map(value.from, 1);
    const to = tr.mapping.map(value.to, -1);
    return from < to ? { ...value, from, to } : null;
  }
  return null;
}

/** The text in a range as the content API reads it (deleted text skipped), its whitespace dropped: a paste's paragraphs lose the blank lines. */
function wordsBetween(doc: PMNode, from: number, to: number): string {
  let out = '';
  doc.nodesBetween(from, to, (node, pos) => {
    if (node.isText && !isDeleted(node)) out += node.text!.slice(Math.max(0, from - pos), Math.max(0, to - pos));
    return true;
  });
  return out.replace(/\s+/g, '');
}

/** Whether the offer's range still holds the text the paste put there (a peer may have typed inside it). */
export function offerHolds(doc: PMNode, offer: MarkdownOffer): boolean {
  return offer.to <= doc.content.size && wordsBetween(doc, offer.from, offer.to) === heldText(offer.text).replace(/\s+/g, '');
}

/**
 * Format as Markdown: the offer's range replaced with its text read leniently, one transaction and
 * one undo step; null when the offer is gone or its text was changed. The paste plugin's state is
 * read through `offerOf`, so this works on any surface that has the plugin.
 */
export async function formatAsMarkdown(view: EditorView, session: EditorDocument, offerOf: (state: EditorState) => MarkdownOffer | null): Promise<(PasteReport & { inserted?: { from: number; to: number } }) | null> {
  const offer = offerOf(view.state);
  if (!offer || !offerHolds(view.state.doc, offer)) return null;
  const converted = markdownToPaste(readMarkdown(offer.text, { lenient: true }), await styleLookupFor(session));
  const range = (): { from: number; to: number } | null => {
    const now = offerOf(view.state);
    if (!now || !offerHolds(view.state.doc, now)) return null;
    // Its own undo step, never joined to the paste's (within Yjs's 500 ms capture) or to typing.
    session.undoManager.stopCapturing();
    return { from: now.from, to: now.to };
  };
  let inserted: { from: number; to: number } | undefined;
  const report = await pasteConverted(view, session, converted, range, { [FORMAT_MARKDOWN_META]: true }, (r) => { inserted = r; });
  // The range the formatted blocks occupy, for the tables in it to be fit (ED-007 section 8.1 item 14).
  return report && inserted ? { ...report, inserted } : report;
}
