// SPDX-License-Identifier: GPL-3.0-or-later
//
// The cases of ED-003 section 7.2 (every command with an `api` line, over `tracked-changes.docx`)
// and the comparison both suites run: `api.test.mts` with tracking off, `agreement.test.mts` with
// it on (ED-005 section 5.8). A case runs twice, through the registry and as its line inside
// `runOnTree`, and the two main parts must agree; a known difference must keep differing.
import { expect } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, commandNamed, runCommand, runCommandAsync, stateFor, selectionSpec, selectionBinding, type CommandSpec } from '../src/index.mjs';

/** Positions of block `i`: the start of its content and the end of its content. */
export function block(session: EditorDocument, i: number): { from: number; to: number } {
  const doc = session.doc;
  let pos = 0;
  for (let k = 0; k < i; k++) pos += doc.child(k).nodeSize;
  return { from: pos + 1, to: pos + 1 + doc.child(i).content.size };
}

/**
 * The main part's XML after a save, marshalled (a command that leaves the document as it was
 * keeps the part's source bytes, where a tree run has marshalled it), with the rsids and the app
 * part out of the comparison.
 */
export async function mainXmlOf(session: EditorDocument): Promise<string> {
  const saved = await session.save({ selfCheck: false });
  const main = (await WordprocessingMLPackage.load(saved.bytes)).getMainDocumentPart();
  await main.getContents();
  return (await main.getXml()).replace(/ w:rsid\w*="[^"]*"/g, "");
}

/** The line evaluated as the console would: an async function over the bindings, inside a tree run. */
export async function runLine(session: EditorDocument, line: string, selection: { from: number; to?: number }): Promise<void> {
  const spec = selectionSpec(stateFor(session, selection));
  const fn = new Function('selection', 'body', 'pkg', `return (async () => { ${line} })();`) as (selection: unknown, body: unknown, pkg: unknown) => Promise<void>;
  const result = await session.runOnTree(({ body, pkg }) => fn(selectionBinding(body, spec), body, pkg));
  if (result.error) throw result.error;
}

/** The first differing region of two strings, with some context each side. */
export function diff(a: string, b: string): [string, string] {
  let i = 0;
  while (i < a.length && i < b.length && a[i] === b[i]) i++;
  const from = Math.max(0, i - 120);
  return [a.slice(from, i + 200), b.slice(from, i + 200)];
}

export interface Case {
  name: string;
  args: unknown;
  selection: { from: number; to?: number };
  /** A difference the check tolerates, with the reason recorded (ED-003 section 11, ED-005 section 1): both exports normalized by it. */
  tolerate?: { normalize: (xml: string) => string; reason: string };
  /**
   * A difference known and recorded, where the engine is the wrong side and a fix is asked of
   * core-ts: the case must differ (so the entry goes when the engine agrees), and nothing else may.
   */
  known?: string;
  /** The same with tracking on (ED-005 section 5.8): a difference between the editor's tracker and the engine's, known and recorded. */
  knownTracked?: string;
  /** Commands run first, through the registry in both documents (over the case's selection unless given), so the case starts from an edited state. */
  before?: { name: string; args: unknown; selection?: { from: number; to?: number } }[];
}

/**
 * Adjacent runs with the same properties as one run: Word writes case F's linked text as one run,
 * the engine as two with identical properties, the editor as one (ED-005 section 1.7).
 */
export const RUNS_MERGED = {
  normalize: (xml: string): string => {
    let out = xml.replace(/<w:t xml:space="preserve">/g, '<w:t>');
    const pair = /<w:r>(<w:rPr>(?:(?!<\/w:rPr>).)*<\/w:rPr>)?<w:t>([^<]*)<\/w:t><\/w:r><w:r>\1<w:t>([^<]*)<\/w:t><\/w:r>/;
    for (let m = pair.exec(out); m; m = pair.exec(out)) out = out.replace(m[0], `<w:r>${m[1] ?? ''}<w:t>${m[2]}${m[3]}</w:t></w:r>`);
    return out;
  },
  reason: 'ED-005 section 1.7',
};

/** A selection over `tracked-changes.docx`: block 4 is "Some text." in one run, block 6 a plain paragraph. */
/** The elements are the engine's, in its order (`excluded-characters.test.mts` holds them to it); the export gives a made break a run of its own, as it does any atom, where the engine keeps one run. */
const A_BREAKS_RUN = 'a made break is in a run of its own (ED-005 section 12.114)';
/** Row shading is untracked in the editor, where Office JS records a cell's w:tcPrChange (ED-005 section 12.130). */
const SHADING_UNTRACKED = "row.shading is untracked: the model holds a cell's property revision only as shown in Word, where the engine records a w:tcPrChange as Word does (ED-005 section 12.130)";
/** A vertical merge: the editor shades the merged cell whole, Office JS the row's own w:tcs alone (ED-005 section 12.130). */
const SHADING_MERGE = "the editor shades a merged cell whole from any row it reaches, writing every w:tc of it, where Office JS writes the row's own w:tcs, never a restart in another row (ED-005 section 12.130, core-ts CR-002 section 67.3): the editor's choice, Jason, 2026-11-18";
/** A caret in a vertically merged cell: in every row the cell spans for the editor, in its first for Office JS. */
const SHADING_MERGED_CARET = "a caret in a vertically merged cell is in every row the cell spans for the editor (prosemirror-tables' selectedRect), so it shades them all, where Office JS's parentRow is the cell's first row; and the merge rule besides (ED-005 section 12.130)";

/** A caret in the first paragraph of the first cell that starts in row `row` of the table at block `i`. */
function rowCaret(session: EditorDocument, i: number, row: number): number {
  const table = session.doc.child(i);
  let pos = block(session, i).from;
  for (let r = 0; r < row; r++) pos += table.child(r).nodeSize;
  return pos + 3;   // into the row, its first cell, its first paragraph
}

export function cases(session: EditorDocument): Case[] {
  const b13 = block(session, 13);   // "Centred ", centred by its own w:jc over Normal
  const b4 = block(session, 4);
  const b6 = block(session, 6);
  const word = { from: b4.from + 5, to: b4.from + 9 };   // "text"
  const whole = { from: b4.from, to: b6.to };             // three blocks
  // Jason Harrop's insertion "An insertion" and deletion " A deletion", for accept and reject (ED-005 section 5.5).
  const within = (text: string): { from: number; to: number } => {
    let at = -1;
    session.doc.descendants((n, pos) => { if (at < 0 && n.isText && n.text!.includes(text)) at = pos + n.text!.indexOf(text); return at < 0; });
    return { from: at, to: at + text.length };
  };
  const insertion = within('An insertion');
  return [
    // Accept and reject (ED-005 section 5.5): the engine's Range.getTrackedChanges and Body.acceptAll, rejectAll.
    { name: 'trackedChanges.acceptAll', args: undefined, selection: { from: b4.from } },
    { name: 'trackedChanges.rejectAll', args: undefined, selection: { from: b4.from } },
    { name: 'trackedChange.accept', args: undefined, selection: insertion },
    { name: 'trackedChange.reject', args: undefined, selection: insertion },
    { name: 'font.bold', args: { value: true }, selection: word },
    { name: 'font.bold', args: { value: 'toggle' }, selection: word },
    { name: 'font.italic', args: { value: true }, selection: word },
    { name: 'font.strikeThrough', args: { value: true }, selection: word },
    { name: 'font.doubleStrikeThrough', args: { value: true }, selection: word },
    { name: 'font.subscript', args: { value: true }, selection: word },
    { name: 'font.superscript', args: { value: true }, selection: word },
    { name: 'font.underline', args: { value: 'Single' }, selection: word },
    { name: 'font.underline', args: { value: 'Double' }, selection: word },
    { name: 'font.name', args: { value: 'Arial' }, selection: word },
    { name: 'font.size', args: { value: 14 }, selection: word },
    { name: 'font.color', args: { value: '#FF0000' }, selection: word },
    { name: 'font.highlightColor', args: { value: 'yellow' }, selection: word },
    { name: 'font.bold', args: { value: true }, selection: whole },
    { name: 'paragraph.style', args: { value: 'Heading1' }, selection: { from: b4.from } },
    { name: 'paragraph.style', args: { value: 'Heading2' }, selection: whole },
    { name: 'paragraph.alignment', args: { value: 'Centered' }, selection: { from: b4.from } },
    { name: 'paragraph.alignment', args: { value: 'Justified' }, selection: whole },
    // Left over a left-aligned style: the w:jc removed, as Word does (ED-005 section 12.44), and the w:pPr it empties with
    // it, by the engine too since 0.3.2 (ED-005 section 12.50's difference, gone).
    { name: 'paragraph.alignment', args: { value: 'Left' }, selection: { from: b13.from } },
    { name: 'paragraph.leftIndent', args: { value: 36 }, selection: { from: b4.from } },
    { name: 'paragraph.rightIndent', args: { value: 18 }, selection: { from: b4.from } },
    { name: 'paragraph.firstLineIndent', args: { value: 18 }, selection: { from: b4.from } },
    { name: 'paragraph.firstLineIndent', args: { value: -18 }, selection: { from: b4.from } },
    { name: 'paragraph.spaceBefore', args: { value: 12 }, selection: { from: b4.from } },
    { name: 'paragraph.spaceAfter', args: { value: 6 }, selection: { from: b4.from } },
    { name: 'paragraph.lineSpacing', args: { value: 24 }, selection: { from: b4.from } },
    // The table at block 7 (a vertical merge, a horizontal one): a column after the first cell's, two before it; the
    // rows covered by the merge get cells too (core-ts 0.4.0's insertColumns, ED-005 section 1.11).
    { name: 'table.addColumns', args: { location: 'After' }, selection: { from: block(session, 7).from + 3 } },
    { name: 'table.addColumns', args: { location: 'Before', count: 2 }, selection: { from: block(session, 7).from + 3 } },
    // Row shading, Office JS's TableRow.shadingColor (core-ts CR-002 section 67, ED-005 section 12.130), over the same table:
    // its first row, one cell with shading of its own, set and cleared; its last row, a horizontal merge, by bare lower-case
    // hex; a caret in the vertically merged cell (rows 2 and 3), and one in row 3's own cell, which the merge reaches.
    { name: 'row.shading', args: { color: '#FFFF00' }, selection: { from: rowCaret(session, 7, 0) }, knownTracked: SHADING_UNTRACKED },
    { name: 'row.shading', args: { color: null }, selection: { from: rowCaret(session, 7, 0) }, knownTracked: SHADING_UNTRACKED },
    { name: 'row.shading', args: { color: '00b0f0' }, selection: { from: rowCaret(session, 7, 3) }, knownTracked: SHADING_UNTRACKED },
    { name: 'row.shading', args: { color: '#D9D9D9' }, selection: { from: rowCaret(session, 7, 1) }, known: SHADING_MERGED_CARET, knownTracked: SHADING_UNTRACKED },
    { name: 'row.shading', args: { color: '#D9D9D9' }, selection: { from: rowCaret(session, 7, 2) }, known: SHADING_MERGE, knownTracked: SHADING_UNTRACKED },
    { name: 'paragraph.increaseIndent', args: undefined, selection: { from: b4.from } },
    { name: 'paragraph.increaseIndent', args: undefined, selection: whole },
    { name: 'insertText', args: { text: 'words' }, selection: word },
    { name: 'insertText', args: { text: 'at the caret' }, selection: { from: b4.from + 5 } },
    // The characters XML 1.0 excludes, as the engine's writers write them (ED-005 section 12.114): a break each, the rest dropped.
    { name: 'insertText', args: { text: 'one\u000Btwo' }, selection: word, known: A_BREAKS_RUN, knownTracked: A_BREAKS_RUN },
    { name: 'insertText', args: { text: 'one\u000Ctwo' }, selection: word, known: A_BREAKS_RUN, knownTracked: A_BREAKS_RUN },
    { name: 'insertText', args: { text: 'one\u000Etwo' }, selection: word, known: A_BREAKS_RUN, knownTracked: A_BREAKS_RUN },
    { name: 'insertText', args: { text: 'one\u0002two\u0007' }, selection: word },
    { name: 'range.hyperlink', args: { address: 'https://www.docx4java.org/' }, selection: word },
    // A link's look (core-ts 0.2.1, CR-002 section 25): the Hyperlink style over another character style, the direct colour gone, bold kept.
    { name: 'range.hyperlink', args: { address: 'https://www.docx4java.org/' }, selection: word, before: [{ name: 'font.style', args: { value: 'Strong' } }, { name: 'font.color', args: { value: '#FF0000' } }, { name: 'font.bold', args: { value: true } }] },
    { name: 'range.hyperlink', args: { subAddress: 'top' }, selection: { from: word.from - 2, to: word.to + 1 }, before: [{ name: 'font.color', args: { value: '#00FF00' } }] },
    // A range over part of a link, as Word does it (core-ts 0.2.2, check 16; ED-005 section 1.6):
    // F, crossing a link's edge, takes the whole link in; G, inside a link, leaves the rest of it plain; H, removing part, removes all.
    { name: 'range.hyperlink', args: { address: 'https://f2.example/' }, selection: { from: b4.from, to: b4.from + 7 }, before: [{ name: 'range.hyperlink', args: { address: 'https://f1.example/' }, selection: word }], tolerate: RUNS_MERGED },
    // Tracked, G and H take the link's look off text whose tracked change gave it: the change is dropped, its formatting being
    // back to `before` (section 5.4), as Word does (core-ts check 24); the engine agrees since 0.3.0 (settleRPrChange).
    { name: 'range.hyperlink', args: { address: 'https://g2.example/' }, selection: { from: b4.from + 2, to: b4.from + 7 }, before: [{ name: 'range.hyperlink', args: { address: 'https://g1.example/' }, selection: { from: b4.from, to: b4.to } }] },
    { name: 'range.hyperlink', args: {}, selection: { from: b4.from + 2, to: b4.from + 7 }, before: [{ name: 'range.hyperlink', args: { address: 'https://h1.example/' }, selection: { from: b4.from, to: b4.to } }] },
    // Find and replace (ED-005 section 3.3): Office JS's search and insertText, the engine's replaceText for a group reference.
    { name: 'find.replace', args: { text: 'text', replacement: 'words' }, selection: word },
    { name: 'find.replace', args: { text: 'text', replacement: 'words', matchWholeWord: true, matchCase: true }, selection: word },
    { name: 'find.replaceAll', args: { text: 'Some', replacement: 'Any' }, selection: { from: b4.from } },
    { name: 'find.replaceAll', args: { text: 'indented', replacement: 'moved', matchWholeWord: true }, selection: { from: b4.from } },
    // A `w:t` shortened ("Font styles " to "Font X") keeps `xml:space="preserve"` only where its text needs it (core-ts 0.2.1; ED-005 section 12.1 item 3).
    { name: 'find.replaceAll', args: { text: 'styles Aerial', replacement: 'X' }, selection: { from: b4.from } },
    { name: 'find.replaceAll', args: { text: 'short para', replacement: 'X' }, selection: { from: b4.from }, knownTracked: "the engine's tracked Replace writes its insertion right after the first deleted piece, before a text-less item inside or at the start of the hit (a bookmark marker, a w:lastRenderedPageBreak run), where the editor writes it after the whole hit; the two orders are Word's either way (ED-008 section 8.8 item 9) and Word's own placement around such an item is unmeasured (core-ts 0.6.0)" },
    { name: 'find.replaceAll', args: { text: 'aaa ', replacement: '' }, selection: { from: b4.from } },
    { name: 'find.replaceAll', args: { text: 't?xt', replacement: 'TXT', matchWildcards: true }, selection: { from: b4.from } },
    // "An" inside Jason Harrop's insertion: with tracking off Word makes the replacement plain, splitting the insertion (core-ts check 21),
    // and so does the engine since 0.3.0, which merges the "W" into the plain run before it where the editor keeps a run of its own.
    // Tracked, Word splits his w:ins around Author C's (check 21); the engine 0.3.0 still nests C's inside his, though fix A is in it:
    // reported to core-ts (ED-005 section 12.26 item 4).
    { name: 'find.replaceAll', args: { text: '<[A-Z][a-z]@>', replacement: 'W', matchWildcards: true, matchCase: true }, selection: { from: b4.from }, knownTracked: "the engine's tracked Replace writes its insertion right after the first deleted piece, before a text-less item inside or at the start of the hit (a bookmark marker, a w:lastRenderedPageBreak run), where the editor writes it after the whole hit; the two orders are Word's either way (ED-008 section 8.8 item 9) and Word's own placement around such an item is unmeasured (core-ts 0.6.0)", tolerate: { ...RUNS_MERGED, reason: 'ED-005 section 12.26 item 4' } },
    { name: 'find.replaceAll', args: { text: 'b+', replacement: 'B', matchRegExp: true }, selection: { from: b4.from } },
    { name: 'find.replaceAll', args: { text: '(Some) (text)', replacement: '$2 $1', matchRegExp: true }, selection: { from: b4.from } },
    { name: 'find.replaceAll', args: { text: '(Left) (indent)', replacement: '\\2 \\1', matchWildcards: true }, selection: { from: b4.from } },
    // Comments (ED-005 section 7.3): made on a selection, then replied to, resolved and deleted through the engine's views over the tree.
    { name: 'range.insertComment', args: { text: 'On "text"' }, selection: word },
    { name: 'comment.reply', args: { id: '0', text: 'A reply' }, selection: { from: b4.from }, before: [{ name: 'range.insertComment', args: { text: 'On "text"' }, selection: word }] },
    { name: 'comment.setResolved', args: { id: '0', value: true }, selection: { from: b4.from }, before: [{ name: 'range.insertComment', args: { text: 'On "text"' }, selection: word }] },
    // The engine's delete leaves split the runs its insertComment split; the editor writes the text as one run again, as the file had it (ED-005 section 12.20).
    { name: 'comment.delete', args: { id: '0' }, selection: { from: b4.from }, before: [{ name: 'range.insertComment', args: { text: 'On "text"' }, selection: word }], tolerate: { ...RUNS_MERGED, reason: 'ED-005 section 12.20' } },
  ];
}


/** How a comparison runs: tracking on or off for both sides, and what else levels them. */
export interface CompareOptions {
  /** Set up each document before the case (tracking on, an author, a fixed date). */
  prepare?: (session: EditorDocument) => void | Promise<void>;
  /** The known difference of a case in this comparison. */
  known: (c: Case) => string | undefined;
  /** What both exports lose before they are compared (revision ids with tracking on). */
  normalize?: (xml: string) => string;
}

/** Every case compared: the findings (empty when all agree) and how many ran. */
export async function compareLines(bytes: Uint8Array, options: CompareOptions): Promise<{ findings: string[]; checked: number }> {
  const probe = await EditorDocument.open(bytes);
  const findings: string[] = [];
  let checked = 0;
  for (const c of cases(probe)) {
    const spec = commandNamed(c.name) as CommandSpec<unknown> | undefined;
    expect(spec, c.name).toBeDefined();
    const line = spec!.api!(c.args);
    expect(line, `${c.name} has no api form for ${JSON.stringify(c.args)}`).toBeDefined();
    const viaCommand = await EditorDocument.open(bytes);
    const viaLine = await EditorDocument.open(bytes);
    for (const session of [viaCommand, viaLine]) await options.prepare?.(session);
    for (const b of c.before ?? []) {
      for (const session of [viaCommand, viaLine]) {
        const r = await runCommandAsync(session, b.name, b.args, b.selection ?? c.selection);
        expect(r.ok, `${b.name} before ${c.name}`).toBe(true);
      }
    }
    // Both halves a command may have around its transaction: `prepare` (a part write) and `settle` (a tree run).
    const r = await runCommandAsync(viaCommand, c.name, c.args, c.selection);
    expect(r.ok, `${c.name} did not apply`).toBe(true);
    await runLine(viaLine, line!, c.selection);
    let [a, b] = await Promise.all([mainXmlOf(viaCommand), mainXmlOf(viaLine)]);
    if (options.normalize) { a = options.normalize(a); b = options.normalize(b); }
    if (c.tolerate) { a = c.tolerate.normalize(a); b = c.tolerate.normalize(b); }
    const known = options.known(c);
    if (known !== undefined) { if (a === b) findings.push(`${c.name} ${JSON.stringify(c.args)}: now agrees with the engine; remove its known difference (${known})`); checked++; continue; }
    if (a !== b) findings.push(`${c.name} ${JSON.stringify(c.args)}: ${line}\n  command: ${diff(a, b)[0]}\n  line:    ${diff(a, b)[1]}`);
    checked++;
  }
  return { findings, checked };
}
