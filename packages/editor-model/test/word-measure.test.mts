// SPDX-License-Identifier: GPL-3.0-or-later
//
// E4.g step G5 (ED-005 section 12.33): the three measurements from Jason's Word session the editor
// did not yet follow, each held to the saves Word made from `measure-base.mjs`'s file
// (`fixtures/review/measure-word2010.docx`, `measure-word15.docx`). The editor starts from the same
// file Word started from, does what the section asked of Word, and must leave what Word left.
import { describe, test, expect } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, runCommandAsync, trackedChangesOf, type ParagraphAttrs } from '../src/index.mjs';
import { sameValue } from '../src/json.mjs';
import { fixture } from './helpers.mjs';
import { buildMeasureBase } from './fixtures/build/measure-base.mjs';

const WORDS = ['review/measure-word2010.docx', 'review/measure-word15.docx'] as const;

/** The first top-level paragraph whose text is `text`, from `after` on, and where it is. */
function paragraph(doc: PMNode, text: string, after = 0): { node: PMNode; pos: number; index: number } {
  let found: { node: PMNode; pos: number; index: number } | undefined;
  doc.forEach((node, pos, index) => { if (!found && index >= after && node.type.name === 'paragraph' && node.textContent === text) found = { node, pos, index }; });
  if (!found) throw new Error(`no paragraph "${text}"`);
  return found;
}

/** Markers of where Word laid out or checked the text, which it writes as it saves: nothing a comment's placement decides. */
const LAYOUT = new Set(['lastRenderedPageBreak', 'proofErr', 'bookmarkStart', 'bookmarkEnd']);

/** A paragraph's content as Word left it: its text, and its atoms by kind (a comment's range markers included). */
function inline(node: PMNode): string[] {
  const out: string[] = [];
  node.forEach((c) => { const kind = c.isText ? `text:${c.text}` : String(c.attrs['kind'] ?? c.type.name); if (!LAYOUT.has(kind)) out.push(kind); });
  return out;
}

async function saved(session: EditorDocument): Promise<Uint8Array> {
  const result = await session.save();
  expect(result.selfCheck?.equal, JSON.stringify(result.selfCheck?.difference)).toBe(true);
  return result.bytes;
}

/**
 * The paragraph after the first "- - -" in a main part's XML (the measurements' empty paragraph), as
 * the elements a comment places there: its range markers and its reference run, in order.
 */
function afterDashes(xml: string): string[] {
  const paragraphs = xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? [];
  const i = paragraphs.findIndex((p) => p.includes('- - -'));
  const next = paragraphs[i + 1] ?? '';
  return [...next.matchAll(/<w:(commentRangeStart|commentRangeEnd|commentReference)\b/g)].map((m) => m[1]!);
}

async function mainXml(bytes: Uint8Array): Promise<string> {
  const reopened = await EditorDocument.open(bytes);
  return reopened.main.getXml();
}

describe('the measurements in Word 2010 and Word 15, followed (ED-005 section 12.33)', () => {
  test('Ctrl+L with tracking off on a paragraph centred over a left-aligned style: the w:jc removed, as both Words write it', async () => {
    const session = await EditorDocument.open(await buildMeasureBase());
    const text = 'This paragraph is centred by direct formatting and then aligned left.';
    expect(paragraph(session.doc, text).node.attrs['jc']).toBe('center');
    expect(runCommand(session, 'paragraph.alignment', { value: 'Left' }, { from: paragraph(session.doc, text).pos + 1 }).ok).toBe(true);
    const mine = paragraph(session.doc, text).node.attrs as ParagraphAttrs;
    expect(mine.jc).toBeNull();
    for (const name of WORDS) {
      const theirs = paragraph((await EditorDocument.open(await fixture(name))).doc, text).node.attrs as ParagraphAttrs;
      for (const k of ['pStyle', 'jc', 'ind', 'spacing', 'bidi', 'numPr'] as const) expect(mine[k], `${name}: ${k}`).toEqual(theirs[k]);
    }
    const xml = await mainXml(await saved(session));
    const written = new RegExp(`<w:p\\b(?:(?!</w:p>).)*${text.replace(/\./g, '\\.')}`, 's').exec(xml)?.[0] ?? '';
    expect(written).not.toBe('');
    expect(written).not.toContain('<w:jc');
    // A value the style does not give is written.
    expect(runCommand(session, 'paragraph.alignment', { value: 'Right' }, { from: paragraph(session.doc, text).pos + 1 }).ok).toBe(true);
    expect(paragraph(session.doc, text).node.attrs['jc']).toBe('right');
  });

  test('New Comment in an empty paragraph: its reference run alone, with no range markers, as both Words write it', async () => {
    const session = await EditorDocument.open(await buildMeasureBase());
    session.trackingAuthor = 'Author A';
    const dashes = paragraph(session.doc, '- - -');
    const empty = paragraph(session.doc, '', dashes.index + 1);
    expect(empty.index).toBe(dashes.index + 1);
    expect((await runCommandAsync(session, 'range.insertComment', { text: 'M3c' }, { from: empty.pos + 1 })).ok).toBe(true);
    // The markers are read from the XML: the projection makes a comment's range a mark on its text, so
    // with no text they would not show in the model on either side.
    const mine = afterDashes(await mainXml(await saved(session)));
    expect(mine).toEqual(['commentReference']);
    for (const name of WORDS) expect(afterDashes(await mainXml(await fixture(name))), name).toEqual(mine);
  });

  test('a heading\'s paragraph mark deleted with tracking on: the next paragraph takes the heading\'s style, recording its own, as Word 15 writes it; accepted, the joined line is a heading; rejected, both as they were', async () => {
    const heading = 'A heading joined to the next paragraph';
    const body = 'The body paragraph after the heading.';
    const deleteTheMark = async (): Promise<EditorDocument> => {
      const session = await EditorDocument.open(await buildMeasureBase());
      session.trackingMode = 'TrackAll';
      session.trackingAuthor = 'Author A';
      const h = paragraph(session.doc, heading);
      const end = h.pos + h.node.nodeSize - 1;
      // Delete at the heading's end: its mark, from the end of its text to the start of the next paragraph's.
      expect(runCommand(session, 'insertText', { text: '' }, { from: end, to: end + 2 }).ok).toBe(true);
      return session;
    };
    const session = await deleteTheMark();
    const h = paragraph(session.doc, heading);
    const next = session.doc.child(h.index + 1).attrs as ParagraphAttrs;
    expect((h.node.attrs as ParagraphAttrs).markRevision?.deletion?.author).toBe('Author A');
    expect(h.node.attrs['pStyle']).toBe('Heading2');
    expect(next.pStyle).toBe('Heading2');
    expect(next.pPrChange?.author).toBe('Author A');

    const word = (await EditorDocument.open(await fixture('review/measure-word15.docx'))).doc;
    const wh = paragraph(word, heading);
    const theirs = word.child(wh.index + 1).attrs as ParagraphAttrs;
    expect((wh.node.attrs as ParagraphAttrs).markRevision?.deletion?.author).toBe('Author A');
    for (const k of ['pStyle', 'jc', 'ind', 'spacing', 'bidi', 'numPr'] as const) expect(next[k], k).toEqual(theirs[k]);
    expect(theirs.pPrChange?.author).toBe('Author A');
    expect(sameValue(next.pPrChange?.before, theirs.pPrChange?.before)).toBe(true);
    // The listing: the mark's deletion and the next paragraph's properties, as two changes.
    expect(trackedChangesOf(session.doc).filter((c) => c.author === 'Author A').map((c) => c.type)).toEqual(['Deleted', 'Formatted']);
    await saved(session);

    // Accepted: the two paragraphs joined, a heading throughout.
    expect((await runCommandAsync(session, 'trackedChanges.acceptAll', {})).ok).toBe(true);
    const joined = paragraph(session.doc, heading + body);
    expect(joined.node.attrs['pStyle']).toBe('Heading2');
    // Rejected: both as they were.
    const again = await deleteTheMark();
    expect((await runCommandAsync(again, 'trackedChanges.rejectAll', {})).ok).toBe(true);
    const back = paragraph(again.doc, heading);
    expect((back.node.attrs as ParagraphAttrs).markRevision).toBeNull();
    const after = again.doc.child(back.index + 1).attrs as ParagraphAttrs;
    expect(after.pStyle).toBeNull();
    expect(after.pPrChange).toBeNull();
  });
});
