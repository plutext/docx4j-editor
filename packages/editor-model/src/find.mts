// SPDX-License-Identifier: GPL-3.0-or-later
//
// Find and replace (ED-005 sections 3.2 and 3.3): the document as Word finds it. Each paragraph's
// text is built as the content API reads it since core-ts 0.6.0 (`Paragraph.text`, Office JS's own
// character stream in Word's default view, check 48; ED-008 section 8.8 items 1 and 12): text, a tab
// U+0009, a line break U+000B, a page break U+000C, a column break U+000E, a carriage return U+000D, a
// note reference U+0002, the hyphens U+001E and U+001F, any symbol "(", an equation its drawn text, a
// picture a place of no text; hidden text and a control's placeholder are not in it, nor deleted text.
// With it a map from its UTF-16 offsets back to ProseMirror positions. The engine's `officeMatches`
// runs over it (Word's Find codes, a picture a plain query does not cross, an optional hyphen passed
// over), with its `searchPattern` precompiled for the usual paragraph of plain text, so the syntax
// (Word's wildcards, regular expressions through `matchRegExp`) and the semantics (case, whole words)
// are the engine's and `body.search` finds what this finds; `find.test.mts` holds the two to that. A
// hit runs across a paragraph mark where Word's does (ED-008 section 8.8 item 11): a wildcard (`?`,
// `*`, a class) crosses one and `^13` is one (check 59, CR-002 section 32 K7), and a plain query's
// `^p` matches one (check 52, CR-014 phase 3). A mark a hit may hold is a paragraph's own, never a
// cell's or a row's end, a deleted mark (which joins its paragraphs) or anything across a table's
// edge, so the paragraphs of a hit are in one cell or in none; under wildcards the document's last
// paragraph's mark is out, as the engine's `Body.search` reads its Content range. Such a hit is
// found and selected, never replaced: the engine's `replaceText` refuses it before changing anything,
// and so do `find.replace` and `find.replaceAll`. A hit that would split a grapheme cluster is not
// produced (ED-001 section 3.7). Replace follows `Range.insertText(text, 'Replace')`: the replacement
// takes the first replaced character's formatting, the text in the hit goes with every atom that
// reads as text (a note reference among them, check 52), and an atom the engine reads no text from
// (a picture, a field character) stays where it was.
import { searchPattern, matchesOf, officeMatches, expandReplacement, textOfView, officeItemText, mathTextOf, hiddenRunsOf, paragraphMarkPieces, type MarkedSegments, type MarkedMatch, type SearchOptions, type SearchSegment, type HiddenRuns, type Element } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { Mark, Node as PMNode } from 'prosemirror-model';
import { Selection, TextSelection, type Transaction } from 'prosemirror-state';
import type { SourceIds } from './ids.mjs';
import { isDeleted } from './revisions.mjs';
import type { Revision } from './schema.mjs';
import { readControl } from './sdtpr.mjs';

/**
 * What a paragraph's text is read under beside the source elements: the resolver that says which
 * runs a style hides (the engine's `hiddenRunsOf`; without it a run's own `w:vanish` alone counts).
 */
export interface FindContext {
  resolver?: { isVanishedByStyles(paragraphStyleId: string | undefined, runStyleId: string | undefined): boolean } | undefined;
}

/** Word's find options, and a regular expression as a fourth (the engine's `matchRegExp`, exclusive with wildcards). */
export interface FindOptions {
  matchCase?: boolean;
  matchWholeWord?: boolean;
  matchWildcards?: boolean;
  matchRegExp?: boolean;
}

/**
 * One hit: where it is, what it matched as the content API reads it (a paragraph mark `\r`), and the
 * match with its groups. One that runs across a paragraph mark or ends with one has `mark`: its `to`
 * is then past the last paragraph's end (`hitSelection` selects into the next paragraph, as Word
 * selects the pilcrow), and it is never replaced.
 */
export interface Hit {
  from: number;
  to: number;
  text: string;
  match: RegExpExecArray;
  mark?: true;
}

/** The hits of one top-level block at positions relative to the block's start; a hit with `mark` may end past the block. */
export interface BlockHits {
  pos: number;
  size: number;
  hits: Hit[];
}

/**
 * The selection of a hit: its text, and across its mark into the next paragraph's start where it
 * holds one (Word's selection of a found pilcrow), since `to` is then no text position.
 */
export function hitSelection(doc: PMNode, hit: Pick<Hit, 'from' | 'to'>): Selection {
  let $from = doc.resolve(hit.from);
  let $to = doc.resolve(hit.to);
  if (!$to.parent.inlineContent) $to = (Selection.findFrom($to, 1, true) ?? Selection.findFrom($to, -1, true))?.$head ?? $to;
  if (!$from.parent.inlineContent) $from = (Selection.findFrom($from, 1, true) ?? Selection.findFrom($from, -1, true))?.$head ?? $from;
  return TextSelection.between($from, $to);
}

/** The options as the engine takes them: only the ones set, so a line written from them reads as a person would write it. */
export function searchOptionsOf(options: FindOptions): SearchOptions {
  const out: SearchOptions = {};
  if (options.matchCase) out.matchCase = true;
  if (options.matchWholeWord) out.matchWholeWord = true;
  if (options.matchWildcards) out.matchWildcards = true;
  if (options.matchRegExp) out.matchRegExp = true;
  return out;
}

const NO_BREAK_HYPHEN = '\u001E';
const SOFT_HYPHEN = '\u001F';
const NOTE_REFERENCE = '\u0002';

/**
 * The text the content API reads from an inline atom, Office JS's characters (the engine's
 * `officeItemText`, check 48): a tab, the breaks, a carriage return, a positional tab "0", the
 * hyphens, any symbol "(", a note reference U+0002, an equation its drawn text (`mathTextOf`), a
 * simple field's result, a stray `w:delText`; nothing for a picture, a field character, a bookmark,
 * a comment marker, a hidden deletion. `ids` finds the element of a field or an equation.
 */
export function atomText(node: PMNode, ids?: SourceIds): string {
  switch (node.type.name) {
    case 'tab': return '\t';
    case 'hardBreak': return node.attrs['kind'] === 'page' ? '\u000C' : node.attrs['kind'] === 'column' ? '\u000E' : '\u000B';
    case 'footnoteRef': return NOTE_REFERENCE;
    case 'opaqueInline': {
      const kind = node.attrs['kind'] as string;
      if (kind === 'noBreakHyphen') return NO_BREAK_HYPHEN;
      if (kind === 'softHyphen') return SOFT_HYPHEN;
      if (kind === 'sym') return '(';
      if (kind === 'cr') return '\u000D';
      if (kind === 'ptab') return '0';
      if (kind === 'footnoteRef' || kind === 'endnoteRef' || kind === 'footnoteReference' || kind === 'endnoteReference') return NOTE_REFERENCE;
      const id = node.attrs['sourceId'] as string | null;
      const el = id && ids ? ids.get(id) : undefined;
      if (!el) return '';
      if (kind === 'fldSimple') return textOfView(el.value as object);
      if (kind === 'delText') return String((el.value as { value?: unknown }).value ?? '');
      if (kind === 'oMath' || kind === 'oMathPara') return mathTextOf(el.value);
      return officeItemText(el as Element) ?? '';
    }
    default: return '';
  }
}

/** The element name the engine's search tells a note reference's kind by (`^f`, `^e`). */
function localPartOf(node: PMNode): string {
  if (node.isText) return 't';
  if (node.type.name === 'footnoteRef') return node.attrs['kind'] === 'endnote' ? 'endnoteReference' : 'footnoteReference';
  if (node.type.name === 'opaqueInline') return String(node.attrs['kind']);
  return node.type.name;
}

/** Whether a control node shows its placeholder, which reads as no text (core-ts 0.6.0, CR-012 section 18). */
function showsPlaceholder(node: PMNode): boolean {
  return (node.type.name === 'sdtInline' || node.type.name === 'sdtBlock') && readControl(node.attrs['sdtPr']).showingPlaceholder;
}

/**
 * Which of a paragraph's runs are hidden (a run's own `w:vanish`, else its styles' through the
 * context's resolver), as a predicate on a text node by its run's source: a run with no source (typed
 * here) is not hidden. Undefined when nothing can be hidden (no source paragraph, or no `ids`).
 */
function hiddenTextOf(paragraph: PMNode, ids: SourceIds | undefined, context: FindContext | undefined): ((node: PMNode) => boolean) | undefined {
  const id = paragraph.attrs['sourceId'] as string | null | undefined;
  const el = id && ids ? ids.get(id) : undefined;
  if (!el) return undefined;
  let hidden: HiddenRuns | undefined;
  return (node) => {
    const source = node.marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'] as string | null | undefined;
    const run = source ? ids!.get(source) : undefined;
    if (!run) return false;
    hidden ??= hiddenRunsOf(context?.resolver, el.value as wml.P);
    return hidden(run.value as wml.R);
  };
}

/** A piece of a paragraph's text: a text node's, or an atom's (possibly empty), at a position; a picture `object`. */
interface Piece {
  start: number;
  end: number;
  pos: number;
  size: number;
  text: boolean;
  object?: boolean;
  item: string;
}

/** A paragraph's text as the content API reads it, and where each piece of it is. */
export interface ParagraphText {
  text: string;
  pieces: Piece[];
}

/**
 * The paragraph's text, `base` being the position of its content's start. Deleted content (a
 * tracked deletion, E4.c) is not in it, as it is not in the engine's accepted view: a hit never
 * covers deleted text, and text either side of a deletion reads as one.
 */
export function paragraphText(paragraph: PMNode, base: number, ids?: SourceIds, context?: FindContext): ParagraphText {
  let text = '';
  const pieces: Piece[] = [];
  const hidden = hiddenTextOf(paragraph, ids, context);
  paragraph.descendants((node, rel) => {
    if (isDeleted(node) || (hidden !== undefined && node.isText && hidden(node))) {
      pieces.push({ start: text.length, end: text.length, pos: base + rel, size: node.nodeSize, text: false, item: 'hidden' });
      return false;
    }
    if (node.isText) {
      const s = node.text ?? '';
      pieces.push({ start: text.length, end: text.length + s.length, pos: base + rel, size: s.length, text: true, item: 't' });
      text += s;
      return false;
    }
    if (node.type.name === 'sdtInline') {
      // an inline control's runs are the paragraph's (the engine's SdtRun holder); one showing its placeholder reads as nothing
      if (!showsPlaceholder(node)) return true;
      pieces.push({ start: text.length, end: text.length, pos: base + rel, size: node.nodeSize, text: false, item: 'sdt' });
      return false;
    }
    if (node.type.name === 'image') {
      pieces.push({ start: text.length, end: text.length, pos: base + rel, size: node.nodeSize, text: false, object: true, item: 'drawing' });
      return false;
    }
    const s = atomText(node, ids);
    pieces.push({ start: text.length, end: text.length + s.length, pos: base + rel, size: node.nodeSize, text: false, item: localPartOf(node) });
    text += s;
    return false;
  });
  return { text, pieces };
}

/**
 * The accepted text between two positions (ED-008 proposal P4): each paragraph's text as `paragraphText` reads it, so deleted
 * and hidden content is out and an atom reads as the engine reads it (a tab `\t`, a break `\n`), cut to the range, the
 * paragraphs joined by a newline. What the ask mode tells the model of a selection, so that a deletion still pending in the
 * document is never read as text it holds.
 */
export function acceptedTextBetween(doc: PMNode, from: number, to: number, ids?: SourceIds, context?: FindContext): string {
  const out: string[] = [];
  doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name !== 'paragraph') return true;
    const { text, pieces } = paragraphText(node, pos + 1, ids, context);
    let s = '';
    for (const p of pieces) {
      if (p.end === p.start) continue;
      if (p.text) {
        const a = Math.max(from, p.pos);
        const b = Math.min(to, p.pos + p.size);
        if (b > a) s += text.slice(p.start + (a - p.pos), p.start + (b - p.pos));
      } else if (p.pos >= from && p.pos < to) {
        s += text.slice(p.start, p.end);
      }
    }
    out.push(s);
    return false;
  });
  return out.join('\n');
}

/** The pieces as the engine's search reads them: text, and a picture as a place of no text (`object`). */
function segmentsOf(pieces: readonly Piece[], text: string): SearchSegment[] {
  const out: SearchSegment[] = [];
  for (const p of pieces) {
    if (p.object) out.push({ text: '', start: p.start, object: true, item: { name: { localPart: p.item } } });
    else if (p.end > p.start) out.push({ text: text.slice(p.start, p.end), start: p.start, item: { name: { localPart: p.item } } });
  }
  return out;
}

/** The position of offset `offset` as a hit's start: the character it points at (an atom with text is covered whole). */
function startPos(pieces: readonly Piece[], offset: number): number {
  for (const p of pieces) {
    if (p.end <= offset || p.start === p.end) continue;
    return p.text ? p.pos + (offset - p.start) : p.pos;
  }
  return pieces.length > 0 ? pieces[pieces.length - 1]!.pos + pieces[pieces.length - 1]!.size : 0;
}

/** The position of offset `offset` as a hit's end: after the character before it. */
function endPos(pieces: readonly Piece[], offset: number): number {
  for (let i = pieces.length - 1; i >= 0; i--) {
    const p = pieces[i]!;
    if (p.start >= offset || p.start === p.end) continue;
    return p.text ? p.pos + (offset - p.start) : p.pos + p.size;
  }
  return pieces.length > 0 ? pieces[0]!.pos : 0;
}

/** Whether a text needs the grapheme check at all: nothing below U+0300 combines or pairs. */
function mayCluster(text: string): boolean {
  for (let i = 0; i < text.length; i++) if (text.charCodeAt(i) >= 0x300) return true;
  return false;
}

// UPSTREAM-24: the engine's `wildcardMatchesAcross` (core-ts 0.7.0) segments every stretch between control characters
// with `Intl.Segmenter`, quadratic in V8, so one long paragraph still costs about a second at 80,000 characters; the copy
// below differs in one line (`rawMatches`), the finder checking a hit's ends per paragraph instead. Measured at the 0.7.0
// install over the 255-page body, the same hits: 82 against 95 ms for a whole search, 5 against 8 ms after a keystroke.
// It goes when the engine's reads a long text in bounded time. `paragraphMarkPieces` and the types are the engine's.

const OBJECT = '\uFFFC';

/** The engine's: a wildcard query's hits over consecutive paragraphs, each followed by its mark where `mark` says (check 59). */
function wildcardMatchesAcross(pieces: readonly MarkedSegments[], query: string, options: SearchOptions): MarkedMatch[] {
  const pattern = searchPattern(query, { ...options, matchWildcards: true });
  let text = '';
  const starts: [number, number][] = [];
  const ends: [number, number][] = [];
  const marks: boolean[] = [];
  const put = (c: string, k: number, start: number, end: number, mark: boolean): void => { text += c; starts.push([k, start]); ends.push([k, end]); marks.push(mark); };
  pieces.forEach((piece, k) => {
    for (const seg of piece.segments) {
      if (seg.hidden) continue;
      if (seg.object) {
        if (piece.from < seg.start && seg.start < piece.to) put(OBJECT, k, seg.start, seg.start, false);
        continue;
      }
      for (let i = 0; i < seg.text.length; i++) {
        const at = seg.start + i;
        if (at >= piece.from && at < piece.to) put(seg.text[i]!, k, at, at + 1, false);
      }
    }
    if (piece.mark) put('\r', k, piece.length, piece.length, true);
  });
  return rawMatches(text, pattern).map((match) => {
    const a = match.index;
    const b = a + match[0].length;
    return { first: starts[a]![0], start: starts[a]![1], last: ends[b - 1]![0], end: ends[b - 1]![1], mark: marks[b - 1]!, match };
  });
}

/**
 * The engine's `matchesOf` less its grapheme filter: every non-empty match of a global pattern. The
 * engine segments the whole text with `Intl.Segmenter`, which is quadratic in V8 (80,000 characters
 * take a second, a 255-page body's half a million never return), so the finder checks a hit's ends
 * against its own paragraphs' boundaries instead (`onBoundaries`).
 */
function rawMatches(text: string, pattern: RegExp): RegExpExecArray[] {
  const out: RegExpExecArray[] = [];
  pattern.lastIndex = 0;
  for (let m = pattern.exec(text); m !== null; m = pattern.exec(text)) {
    if (m[0].length === 0) { pattern.lastIndex++; continue; }
    out.push(m);
  }
  return out;
}

/** A paragraph of a document in the finder's walk: its content's start, the cell it is in (the cell's position, -1 for none), whether its own mark may be matched, and whether it reads as no text (inside a control showing its placeholder). */
interface Para {
  node: PMNode;
  base: number;
  cell: number;
  mark: boolean;
  empty: boolean;
}

/**
 * Every paragraph of a document in order, as the engine's `Body.paragraphs` lists them: through
 * tables and block controls, never into an atom. A paragraph's mark may be matched where it is the
 * paragraph's own: not a cell's end (the cell's last paragraph as Word reads it, through block
 * controls: `endCellsWithParagraph`'s rule) and not deleted (`markText` is then nothing, and the
 * paragraph reads on into the next).
 */
function paragraphsOf(doc: PMNode): Para[] {
  const out: Para[] = [];
  const cells: { pos: number; end: number; last: PMNode | null }[] = [];
  let placeholderEnd = -1;
  doc.descendants((node, pos) => {
    while (cells.length > 0 && pos >= cells[cells.length - 1]!.end) cells.pop();
    if (node.type.name === 'paragraph') {
      const cell = cells[cells.length - 1];
      const deleted = !!(node.attrs['markRevision'] as Revision | null)?.deletion;
      out.push({ node, base: pos + 1, cell: cell ? cell.pos : -1, mark: !deleted && cell?.last !== node, empty: pos < placeholderEnd });
      return false;
    }
    if (node.isAtom) return false;
    if (node.type.name === 'table_cell') cells.push({ pos, end: pos + node.nodeSize, last: lastParagraphOf(node) });
    else if (showsPlaceholder(node) && pos >= placeholderEnd) placeholderEnd = pos + node.nodeSize;
    return true;
  });
  return out;
}

/** The cell's last paragraph as Word reads it, through block controls at its end; null when it ends in a table. */
function lastParagraphOf(cell: PMNode): PMNode | null {
  let last = cell.lastChild;
  while (last?.type.name === 'sdtBlock') last = last.lastChild;
  return last?.type.name === 'paragraph' ? last : null;
}

/** A hit over a group of paragraphs, by paragraph index within the group and offset in its text; `mark` where it ends with the last one's mark. */
type GroupMatch = MarkedMatch;

/** A group's hits memoised against the paragraph nodes and marks they were found over. */
interface GroupMemo {
  nodes: PMNode[];
  marks: boolean[];
  matches: GroupMatch[];
}

/** One paragraph's part in a `^p` search: where the first piece ends at the text's end, whether each middle piece is the whole text, where the last piece from the start ends. */
interface MarkSearch {
  starts: number[];
  whole: boolean[];
  endAt: number | undefined;
}

let segmenter: Intl.Segmenter | undefined;

/** The grapheme boundaries of a text, as offsets. */
function boundaries(text: string): Set<number> {
  const out = new Set<number>([text.length]);
  for (const { index } of (segmenter ??= new Intl.Segmenter(undefined, { granularity: 'grapheme' })).segment(text)) out.add(index);
  return out;
}

const EMPTY_TEXT: ParagraphText = { text: '', pieces: [] };

/** Why a hit holding a paragraph mark is not replaced: the engine's `replaceText` refuses it before changing anything (CR-002 section 32, CR-014 phase 3). */
export const HOLDS_MARK = "a hit holds a paragraph mark, which Word's Find matches and whose replacement is not measured; it is found and selected, never replaced";

/**
 * A query compiled once, with its hits memoised per top-level block node (ED-005 section 3.6): a
 * block a transaction did not touch is the same node object afterwards, so its hits are read
 * back rather than searched again. The pattern is the engine's; an invalid expression or
 * wildcard pattern throws from the constructor, with the engine's message.
 */
export class Finder {
  readonly pattern: RegExp;
  private readonly memo = new WeakMap<PMNode, Hit[]>();
  /** Blocks searched afresh (the benchmark and the tests read it); with `across`, the groups of paragraphs searched afresh. */
  misses = 0;

  /** Whether the query needs the engine's full reading: a plain one holding a Find code (`^t`, `^f`, …). */
  private readonly coded: boolean;

  /**
   * How a hit may run across paragraphs: a wildcard query whose pattern can take a mark (`?`, `*`, a
   * negated class, `^13`) crosses one as Word's does, a plain query's `^p` matches one; undefined where
   * every hit is one paragraph's (the engine runs every wildcard query across, and finds the same for
   * one that cannot match `\r`, which keeps the memo per block).
   */
  readonly across: 'wildcards' | 'marks' | undefined;
  /** Whether the wildcard pattern can take a mark (`?`, `*`, a negated class, `^13`); one that cannot still crosses a deleted mark's join, so a document with one is searched across too (`joins`). */
  private readonly crossing: boolean;
  private readonly markPieces: string[] | undefined;
  /** Per top-level block, whether a paragraph in it has a deleted mark (its text runs into the next paragraph's). */
  private readonly joins = new WeakMap<PMNode, boolean>();
  /** Each paragraph's text at base 0 (its content's start), and the grapheme boundaries of a text, memoised by node. */
  private readonly texts = new WeakMap<PMNode, ParagraphText>();
  private readonly bounds = new WeakMap<PMNode, Set<number> | null>();
  /** A group's hits by its first paragraph node, a `^p` search's parts by paragraph node, a document's hits by its node. */
  private readonly groups = new WeakMap<PMNode, GroupMemo>();
  private readonly markSearches = new WeakMap<PMNode, MarkSearch>();
  private readonly docs = new WeakMap<PMNode, Hit[]>();

  constructor(readonly query: string, readonly options: FindOptions = {}, private readonly ids?: SourceIds, private readonly context?: FindContext) {
    this.pattern = searchPattern(query, searchOptionsOf(options));
    this.coded = !options.matchRegExp && !options.matchWildcards && query.includes('^');
    this.markPieces = options.matchRegExp || options.matchWildcards ? undefined : paragraphMarkPieces(query);
    this.across = options.matchWildcards ? 'wildcards' : this.markPieces ? 'marks' : undefined;
    this.crossing = this.across === 'marks' || (this.across === 'wildcards' && /\[\\s\\S\]|\[\^|\\r|\r/.test(this.pattern.source));
  }

  /**
   * Whether this document is searched across paragraphs: `^p`, or a wildcard pattern that can take a
   * mark, or any wildcard query over a document with a deleted mark, whose paragraphs read as one
   * (the engine runs every wildcard query across; one that cannot match `\r` finds the same per
   * paragraph, which keeps the memo per block). The blocks' memo is filled on the way.
   */
  acrossIn(doc: PMNode): boolean {
    if (!this.across) return false;
    if (this.crossing) return true;
    let joined = false;
    doc.forEach((block) => { if (!joined && this.hasJoin(block)) joined = true; });
    return joined;
  }

  private hasJoin(block: PMNode): boolean {
    let j = this.joins.get(block);
    if (j === undefined) {
      j = false;
      if (block.type.name === 'paragraph') j = !!(block.attrs['markRevision'] as Revision | null)?.deletion;
      else if (!block.isAtom) block.descendants((node) => { if (j) return false; if (node.type.name === 'paragraph') { j = !!(node.attrs['markRevision'] as Revision | null)?.deletion; return false; } return !node.isAtom; });
      this.joins.set(block, j);
    }
    return j;
  }

  /** The hits of one paragraph, `base` its content's start. */
  inParagraph(paragraph: PMNode, base: number): Hit[] {
    if (this.query === '') return [];
    const { text, pieces } = paragraphText(paragraph, base, this.ids, this.context);
    if (text === '') return [];
    // The usual paragraph, text alone: the precompiled pattern. Otherwise the engine's reading, which a plain
    // query does not cross a picture in, passes an optional hyphen over, and takes Word's Find codes in.
    const plain = !this.coded && !pieces.some((p) => p.object) && (!(this.options.matchRegExp || this.options.matchWildcards) ? !text.includes(SOFT_HYPHEN) : true);
    const matches = plain
      ? matchesOf(text, this.pattern).map((m) => ({ start: m.index, end: m.index + m[0].length, match: m }))
      : officeMatches(segmentsOf(pieces, text), this.query, searchOptionsOf(this.options));
    if (matches.length === 0) return [];
    const bounds = mayCluster(text) ? boundaries(text) : undefined;
    const out: Hit[] = [];
    for (const m of matches) {
      const s = m.start;
      const e = m.end;
      if (e === s) continue;   // a wildcard class matching a picture's place: nothing to select or replace
      if (bounds && (!bounds.has(s) || !bounds.has(e))) continue;
      out.push({ from: startPos(pieces, s), to: endPos(pieces, e), text: text.slice(s, e), match: m.match });
    }
    return out;
  }

  /** Whether a block's hits are memoised: it was searched before, and no transaction has changed it since. */
  has(block: PMNode): boolean {
    return this.memo.has(block);
  }

  /** The hits of one top-level block, at positions relative to the block's start. */
  inBlock(block: PMNode): Hit[] {
    let hits = this.memo.get(block);
    if (hits) return hits;
    this.misses++;
    hits = [];
    if (block.type.name === 'paragraph') hits = this.inParagraph(block, 1);
    else if (!block.isAtom && !showsPlaceholder(block)) {
      const found: Hit[] = [];
      block.descendants((node, rel) => {
        if (node.type.name !== 'paragraph') return !node.isAtom && !showsPlaceholder(node);
        found.push(...this.inParagraph(node, rel + 2));
        return false;
      });
      hits = found;
    }
    this.memo.set(block, hits);
    return hits;
  }

  /** Every hit in a document, in order. */
  all(doc: PMNode): Hit[] {
    if (this.acrossIn(doc)) return this.acrossAll(doc);
    const out: Hit[] = [];
    doc.forEach((block, offset) => {
      for (const h of this.inBlock(block)) out.push(offset === 0 ? h : { ...h, from: h.from + offset, to: h.to + offset });
    });
    return out;
  }

  /**
   * The hits by the top-level block each starts in, the blocks with none left out, with the ranges of
   * the blocks searched afresh (a view patches those). Searched across paragraphs (`across` true, see
   * `acrossIn`), a block's hits depend on its neighbours: every block is then fresh, and a view draws
   * the list whole.
   */
  byBlock(doc: PMNode): { blocks: BlockHits[]; changed: [number, number][]; across: boolean } {
    const blocks: BlockHits[] = [];
    const changed: [number, number][] = [];
    if (this.acrossIn(doc)) {
      const hits = this.acrossAll(doc);
      let i = 0;
      doc.forEach((block, pos) => {
        changed.push([pos, pos + block.nodeSize]);
        const end = pos + block.nodeSize;
        const own: Hit[] = [];
        while (i < hits.length && hits[i]!.from < end) { const h = hits[i++]!; own.push({ ...h, from: h.from - pos, to: h.to - pos }); }
        if (own.length > 0) blocks.push({ pos, size: block.nodeSize, hits: own });
      });
      return { blocks, changed, across: true };
    }
    doc.forEach((block, pos) => {
      const fresh = !this.has(block);
      const hits = this.inBlock(block);
      if (fresh) changed.push([pos, pos + block.nodeSize]);
      if (hits.length > 0) blocks.push({ pos, size: block.nodeSize, hits });
    });
    return { blocks, changed, across: false };
  }

  /** A paragraph's text at base 0, memoised by node; nothing for one inside a control showing its placeholder. */
  private textOf(p: Para): ParagraphText {
    if (p.empty) return EMPTY_TEXT;
    let t = this.texts.get(p.node);
    if (!t) { t = paragraphText(p.node, 0, this.ids, this.context); this.texts.set(p.node, t); }
    return t;
  }

  /** Whether an offset of a paragraph's text is on a grapheme boundary (memoised per node; null where no text can cluster). */
  private onBoundary(p: Para, text: string, offset: number): boolean {
    let b = this.bounds.get(p.node);
    if (b === undefined) { b = mayCluster(text) ? boundaries(text) : null; this.bounds.set(p.node, b); }
    return b === null || b.has(offset);
  }

  /**
   * The hits of a document whose query runs across paragraphs: the paragraphs in order, grouped into
   * runs in one cell or in none (a hit never crosses a table's edge), each group searched as the
   * engine's `Range.searchWildcards` or `searchMarks` does and memoised against its nodes, so a
   * keystroke searches its group again and no other.
   */
  private acrossAll(doc: PMNode): Hit[] {
    const done = this.docs.get(doc);
    if (done) return done;
    const paras = paragraphsOf(doc);
    const out: Hit[] = [];
    const last = paras.length - 1;
    for (let from = 0; from <= last;) {
      let to = from;
      while (to < last && paras[to + 1]!.cell === paras[from]!.cell) to++;
      const group = paras.slice(from, to + 1);
      // Under wildcards the document's last mark is out (the engine's Body.search reads its Content range); `^p` takes it (Whole).
      const marks = group.map((p, k) => p.mark && !(this.across === 'wildcards' && from + k === last));
      for (const m of this.groupMatches(group, marks)) {
        const a = group[m.first]!;
        const b = group[m.last]!;
        const ta = this.textOf(a);
        const tb = this.textOf(b);
        if (!m.mark && m.first === m.last && m.start === m.end) continue;   // a class on a picture's place: nothing to select
        if (!this.onBoundary(a, ta.text, m.start) || !this.onBoundary(b, tb.text, m.end)) continue;
        let text = '';
        for (let k = m.first; k <= m.last; k++) {
          const t = this.textOf(group[k]!).text;
          text += t.slice(k === m.first ? m.start : 0, k === m.last ? m.end : t.length);
          if (k < m.last && marks[k]) text += '\r';
        }
        if (m.mark) text += '\r';
        const hit: Hit = { from: a.base + startPos(ta.pieces, m.start), to: m.mark ? b.base + b.node.content.size + 1 : b.base + endPos(tb.pieces, m.end), text, match: m.match };
        if (m.mark || m.first !== m.last) hit.mark = true;
        out.push(hit);
      }
      from = to + 1;
    }
    this.docs.set(doc, out);
    return out;
  }

  /** A group's matches, memoised against its paragraph nodes and their marks. */
  private groupMatches(group: readonly Para[], marks: readonly boolean[]): GroupMatch[] {
    const first = group[0]!.node;
    const memo = this.groups.get(first);
    if (memo && memo.nodes.length === group.length && memo.nodes.every((n, k) => n === group[k]!.node) && memo.marks.every((m, k) => m === marks[k])) return memo.matches;
    this.misses++;
    const matches = this.across === 'wildcards' ? this.wildcardsOver(group, marks) : this.marksOver(group, marks);
    this.groups.set(first, { nodes: group.map((p) => p.node), marks: [...marks], matches });
    return matches;
  }

  /** The engine's `searchWildcards` over one group. */
  private wildcardsOver(group: readonly Para[], marks: readonly boolean[]): GroupMatch[] {
    const pieces: MarkedSegments[] = group.map((p, k) => {
      const t = this.textOf(p);
      return { segments: segmentsOf(t.pieces, t.text), from: 0, to: t.text.length, length: t.text.length, mark: marks[k]! };
    });
    return wildcardMatchesAcross(pieces, this.query, searchOptionsOf(this.options));
  }

  /**
   * The engine's `searchMarks` over one group (check 52 case 2): a hit starts where the first piece
   * matches to a paragraph's end, takes each middle piece as a whole paragraph's text, and ends where
   * the last piece matches from the next paragraph's start; a query ending with `^p` ends with that
   * paragraph's mark, which the hit holds. Each paragraph's part is memoised by node (`markSearch`).
   */
  private marksOver(group: readonly Para[], marks: readonly boolean[]): GroupMatch[] {
    const pieces = this.markPieces!;
    const n = pieces.length - 1;
    const endsWithMark = pieces[n] === '';
    const last = group.length - 1;
    const out: GroupMatch[] = [];
    for (let i = 0; i <= last; i++) {
      const end = i + n - (endsWithMark ? 1 : 0);
      if (end > last) break;
      let ok = true;
      for (let j = i; j < i + n && ok; j++) ok = marks[j]!;
      for (let j = 1; j < n && ok; j++) ok = this.markSearch(group[i + j]!).whole[j]!;
      if (!ok) continue;
      const starts = this.markSearch(group[i]!).starts;
      if (starts.length === 0) continue;
      const endParagraph = group[end]!;
      const tEnd = this.textOf(endParagraph).text;
      let to: number;
      if (endsWithMark) to = tEnd.length;
      else {
        const at = this.markSearch(endParagraph).endAt;
        if (at === undefined) continue;
        to = at;
      }
      const matched = group.slice(i, end + 1).map((p, k) => { const t = this.textOf(p).text; return t.slice(k === 0 ? starts[0]! : 0, k === end - i ? to : t.length); }).join('\r') + (endsWithMark ? '\r' : '');
      out.push({ first: i, start: starts[0]!, last: end, end: to, mark: endsWithMark, match: Object.assign([matched] as string[], { index: 0, input: matched }) as RegExpExecArray });
    }
    return out;
  }

  /** One paragraph's part in the `^p` search, memoised by node. */
  private markSearch(p: Para): MarkSearch {
    let m = this.markSearches.get(p.node);
    if (m) return m;
    const pieces = this.markPieces!;
    const n = pieces.length - 1;
    const t = this.textOf(p);
    const segments = segmentsOf(t.pieces, t.text);
    const options = searchOptionsOf(this.options);
    const len = t.text.length;
    const matches = (q: string): { start: number; end: number }[] => officeMatches(segments, q, options, 0, len);
    const starts = pieces[0] === '' ? [len] : matches(pieces[0]!).filter((h) => h.end === len).map((h) => h.start);
    const whole = pieces.map((q, j) => (j === 0 || j === n ? false : q === '' ? len === 0 : matches(q).some((h) => h.start === 0 && h.end === len)));
    const endAt = pieces[n] === '' ? undefined : matches(pieces[n]!).find((h) => h.start === 0)?.end;
    m = { starts, whole, endAt };
    this.markSearches.set(p.node, m);
    return m;
  }

  /** The replacement for one hit: the text itself, or with its group references expanded (`$1`, Word's `\1`; CR-002 section 26). */
  replacementFor(hit: Hit, replacement: string): string {
    return expandReplacement(replacement, hit.match, searchOptionsOf(this.options));
  }

  /**
   * Replace hits of this query in a transaction over the document they were found in, last to
   * first so the earlier ones keep their positions, as the engine's `replaceText` does. Each is
   * `Range.insertText(replacement, 'Replace')`: the text and the atoms with text inside the hit
   * go, the atoms with none and deleted content stay, and the replacement goes in at the hit's
   * start with the first replaced character's marks. The number replaced.
   */
  replace(tr: Transaction, hits: readonly Hit[], replacement: string): number {
    if (hits.some((h) => h.mark)) throw new Error(HOLDS_MARK);
    for (let i = hits.length - 1; i >= 0; i--) this.replaceOne(tr, hits[i]!, this.replacementFor(hits[i]!, replacement));
    return hits.length;
  }

  private replaceOne(tr: Transaction, hit: Hit, text: string): void {
    const doc = tr.doc;
    const first = doc.nodeAt(hit.from);
    let marks: readonly Mark[] | undefined = first?.isText ? first.marks : undefined;
    const doomed: [number, number][] = [];
    doc.nodesBetween(hit.from, hit.to, (node, pos) => {
      // Deleted content inside a hit stays, as the engine's `splice` works in the accepted view.
      if (isDeleted(node)) return false;
      if (node.isText) { doomed.push([Math.max(pos, hit.from), Math.min(pos + node.nodeSize, hit.to)]); return false; }
      if (!node.isInline) return true;
      if (node.type.name === 'sdtInline') return true;
      if (pos >= hit.from && pos + node.nodeSize <= hit.to && atomText(node, this.ids) !== '') doomed.push([pos, pos + node.nodeSize]);
      return false;
    });
    for (let i = doomed.length - 1; i >= 0; i--) tr.delete(doomed[i]![0], doomed[i]![1]);
    if (text === '') return;
    if (!marks) {
      // The hit began with a tab or a symbol: the engine extends the text before it, else the text after.
      const $at = tr.doc.resolve(hit.from);
      marks = $at.nodeBefore?.isText ? $at.nodeBefore.marks : $at.nodeAfter?.isText ? $at.nodeAfter.marks : [];
    }
    tr.insert(hit.from, tr.doc.type.schema.text(text, marks));
  }
}

/** Every hit of a query in a document (a `Finder` for one use). */
export function findIn(doc: PMNode, query: string, options: FindOptions = {}, ids?: SourceIds, context?: FindContext): Hit[] {
  return new Finder(query, options, ids, context).all(doc);
}

/** Whether a replacement refers to a group of the match, which Office JS's `insertText` cannot express. */
export function hasGroupReference(replacement: string, options: FindOptions): boolean {
  if (options.matchRegExp) return /\$[$&1-9]/.test(replacement);
  if (options.matchWildcards) return /\\[1-9]/.test(replacement);
  return false;
}
