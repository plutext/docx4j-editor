// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 8.3 (step F2): paste from Word over the corpus Jason captured from Word 15 on
// the reference machine (`test/fixtures/paste/results/*.html`), the converted form pinned beside
// each capture (`UPDATE_PASTE=1` regenerates after a deliberate change), and the paste through a
// view as one undo step.
import { describe, test, expect, vi } from 'vitest';
import { readFile, writeFile } from 'node:fs/promises';
import { deflateSync } from 'node:zlib';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { TextSelection } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { unzipSync, strFromU8 } from 'fflate';
import { EditorDocument, hyperlinkTarget } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { Slice } from 'prosemirror-model';
import { convertHtml, styleLookupFor, styleLookupOf, clipboardMarkup, describePaste, pasteHtml, isEditorClipboard, isTerminalHtml, markdownOfferOf, pastedAddress, clipboardPictures, type PasteParagraph, type PasteTable, type PasteRun } from '../src/view/paste.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');
const corpus = join(fixtures, 'paste', 'results');

/**
 * A capture as text. `Get-Clipboard | Out-File` on the reference machine wrote the clipboard's
 * UTF-8 bytes through a Latin-1 reading and out as UTF-8 again, so a non-breaking space arrived
 * as "Â" followed by one; the reading is undone here, since a browser hands the editor the text
 * decoded once. The converter itself never sees the doubled form.
 */
async function capture(name: string): Promise<string> {
  const bytes = await readFile(join(corpus, `${name}.html`));
  let text = bytes.toString('utf8');
  text = text.replace(/^\uFEFF/, '');
  // The clipboard's UTF-8 read as Windows-1252 and written out as UTF-8 again: a non-breaking space
  // arrived as A-circumflex followed by one, and a Japanese font name in the head as a run of such pairs.
  if (/[\u00c2\u00c3][\u0080-\u00bf\u20ac\u201a\u0192\u201e\u2026\u2020\u2021\u02c6\u2030\u0160\u2039\u0152\u017d\u2018\u2019\u201c\u201d\u2022\u2013\u2014\u02dc\u2122\u0161\u203a\u0153\u017e\u0178]/.test(text)) text = fromCp1252(text);
  return text;
}

/** The Windows-1252 characters above 0x7F that are not Latin-1, by their byte. */
const CP1252: Record<string, number> = { '\u20ac': 0x80, '\u201a': 0x82, '\u0192': 0x83, '\u201e': 0x84, '\u2026': 0x85, '\u2020': 0x86, '\u2021': 0x87, '\u02c6': 0x88, '\u2030': 0x89, '\u0160': 0x8a, '\u2039': 0x8b, '\u0152': 0x8c, '\u017d': 0x8e, '\u2018': 0x91, '\u2019': 0x92, '\u201c': 0x93, '\u201d': 0x94, '\u2022': 0x95, '\u2013': 0x96, '\u2014': 0x97, '\u02dc': 0x98, '\u2122': 0x99, '\u0161': 0x9a, '\u203a': 0x9b, '\u0153': 0x9c, '\u017e': 0x9e, '\u0178': 0x9f };

/** Text that was UTF-8 bytes read as Windows-1252: the bytes back, decoded once. */
function fromCp1252(text: string): string {
  const bytes = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) { const c = text.charCodeAt(i); bytes[i] = c < 0x100 ? c : (CP1252[text[i]!] ?? 0x3f); }
  return Buffer.from(bytes).toString('utf8');
}

async function lookup() {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'paste', 'paste-samples.docx'))));
  return styleLookupFor(session);
}

const NAMES = ['headings', 'inline', 'bullets', 'numbering', 'table', 'dropped', 'all'];

describe('paste from Word: the corpus', () => {
  for (const name of NAMES) {
    test(`${name}.html converts as pinned`, async () => {
      const result = convertHtml(await capture(name), await lookup());
      const path = join(corpus, `${name}.expected.json`);
      const json = JSON.stringify(result, null, 2) + '\n';
      if (process.env['UPDATE_PASTE']) { await writeFile(path, json); return; }
      let expected: string;
      try { expected = await readFile(path, 'utf8'); } catch { throw new Error(`no expectation for ${name}: run with UPDATE_PASTE=1 and commit it`); }
      expect(JSON.parse(json)).toEqual(JSON.parse(expected));
    });
  }

  test('the CF_HTML header is skipped and the fragment markers are harmless; the editor\'s own clipboard is left to ProseMirror', async () => {
    const text = await capture('headings');
    expect(text.startsWith('Version:')).toBe(true);
    // The fragment Word marked, and nothing of the head or the tail.
    expect(clipboardMarkup(text).trim().startsWith('<h1>')).toBe(true);
    expect(clipboardMarkup(text)).not.toContain('<body');
    expect(isEditorClipboard('<p data-pm-slice="1 1 []">x</p>')).toBe(true);
    expect(isEditorClipboard(clipboardMarkup(text))).toBe(false);
  });

  test('headings: the levels by style name, plain paragraphs as the default', async () => {
    const { blocks, report } = convertHtml(await capture('headings'), await lookup());
    const styles = blocks.map((b) => (b as PasteParagraph).style);
    expect(styles).toEqual(['Heading1', 'Heading2', null, 'Heading3', null]);
    expect((blocks[2] as PasteParagraph).inlines).toEqual([{ text: 'Body text under the second-level heading, one plain sentence.' }]);
    expect(report).toEqual({ paragraphs: 5, lists: 0, tables: 0, dropped: {}, stylesWanted: [] });
  });

  test('inline: the six marks each on its own run, two together, a line break and a tab', async () => {
    const { blocks } = convertHtml(await capture('inline'), await lookup());
    const marks = blocks[1] as PasteParagraph;
    expect(marks.inlines).toEqual([
      { text: 'This paragraph has ' }, { text: 'bold', bold: true }, { text: ', ' }, { text: 'italic', italic: true }, { text: ', ' },
      { text: 'underlined', underline: true }, { text: ', ' }, { text: 'struck through', strike: true }, { text: ', a super' },
      { text: 'script', vertAlign: 'superscript' }, { text: ' and a sub' }, { text: 'script', vertAlign: 'subscript' }, { text: ' run, then ' },
      { text: 'bold italic together', bold: true, italic: true }, { text: '.' },
    ]);
    const breaks = blocks[2] as PasteParagraph;
    expect(breaks.inlines).toEqual([
      { text: 'A line break follows this text' }, { br: true },
      { text: 'and this is the same paragraph after it, then a tab' }, { tab: true },
      { text: 'and text after the tab.' },
    ]);
  });

  test('bullets and numbering: two levels, one list each, the paragraph after outside it', async () => {
    for (const [name, bullet] of [['bullets', true], ['numbering', false]] as const) {
      const { blocks, report } = convertHtml(await capture(name), await lookup());
      const items = blocks.slice(1, 6) as PasteParagraph[];
      expect(items.map((p) => p.list?.level)).toEqual([0, 0, 1, 1, 0]);
      expect(items.every((p) => p.list?.bullet === bullet)).toBe(true);
      expect(new Set(items.map((p) => p.list?.key)).size).toBe(1);
      expect((blocks[6] as PasteParagraph).list).toBeUndefined();
      expect(report.lists).toBe(1);
    }
  });

  test('table: three by three with the grid style, cell widths from Word, a paragraph per cell', async () => {
    const { blocks, report } = convertHtml(await capture('table'), await lookup());
    const table = blocks[2] as PasteTable;
    expect(table.kind).toBe('table');
    expect(table.style).toBe('TableGrid');
    expect(table.rows.map((r) => r.map((c) => (c.blocks[0] as PasteParagraph).inlines.map((i) => ('text' in i ? i.text : '')).join('')))).toEqual([['Name', 'Quantity', 'Price'], ['Widget', '2', '3.50'], ['Gadget', '10', '12.00']]);
    expect(table.rows[0]![0]!.width).toBe(3009);   // 150.45pt
    expect(report.tables).toBe(1);
  });

  test('direct formatting is kept where the element states it (section 8.6, F5): the alignment, the indent, the font, size, colour and highlight; nothing dropped', async () => {
    const { blocks, report } = convertHtml(await capture('dropped'), await lookup());
    const centred = blocks[1] as PasteParagraph;
    expect(centred.inlines.map((i) => ('text' in i ? i.text : '')).join('')).toContain('keeps the text and drops the rest');
    // `margin-left:36.0pt;text-align:center` on the paragraph; `font-size:14.0pt;font-family:"Arial",sans-serif;color:red` on its run.
    expect(centred.jc).toBe('center');
    expect(centred.ind).toEqual({ start: 720 });
    expect(centred.spacing).toBeUndefined();
    expect(centred.inlines[0]).toEqual({ text: centred.inlines.map((i) => ('text' in i ? i.text : '')).join(''), fontName: 'Arial', fontSize: 14, color: '#FF0000' });
    // `background:yellow;mso-highlight:yellow` on the run Word wrote (the word and the full stop after it, as the capture has it).
    const highlighted = blocks[2] as PasteParagraph;
    expect(highlighted.inlines).toEqual([{ text: 'This one has a ' }, { text: 'highlighted word.', highlight: '#FFFF00' }]);
    expect(report.dropped).toEqual({});
    expect(describePaste(report)).toBe('Pasted 3 paragraphs');
    expect(report.stylesWanted).toEqual([]);
  });

  test('margin-bottom and line-height stay dropped and counted: Word writes them on every paragraph from the source\'s style', async () => {
    const { report } = convertHtml(await capture('table'), await lookup());
    expect(report.dropped).toEqual({ spacing: 9 });   // the nine cell paragraphs, `margin-bottom:0cm;line-height:normal` on each
    expect(describePaste(report)).toContain('dropped spacing (9)');
  });

  test('the unit and colour readings: lengths to twips, a hanging indent, named and hex colours, generic families and unknown highlights counted', async () => {
    const parse = (html: string) => convertHtml(`<html><body>${html}</body></html>`, { idFor: () => undefined, hasId: () => false });
    const p = (html: string) => parse(html).blocks[0] as PasteParagraph;
    expect(p('<p style="margin-left:1cm;margin-right:0.5in;text-indent:-18pt;margin-top:6pt;text-align:justify">x</p>')).toMatchObject({ jc: 'both', ind: { start: 567, end: 720, hanging: 360 }, spacing: { before: 120 } });
    expect(p('<p align="right" style="text-indent:36px">x</p>')).toMatchObject({ jc: 'right', ind: { firstLine: 540 } });
    expect(p('<p><span style="color:#1F4E79;font-size:12px;font-family:Calibri">x</span></p>').inlines[0]).toEqual({ text: 'x', color: '#1F4E79', fontSize: 9, fontName: 'Calibri' });
    expect(p('<p><span style="color:#abc;background:#FFFF00">x</span></p>').inlines[0]).toEqual({ text: 'x', color: '#AABBCC' });   // a hex background is not one of Word's highlight names
    expect(p('<p><span style="color:windowtext">x</span></p>').inlines[0]).toEqual({ text: 'x', color: 'auto' });
    const generic = parse('<p style="text-align:middle"><span style="font-family:sans-serif;background:#FFFF00;color:hotpink;font-size:large">x</span></p>');
    expect((generic.blocks[0] as PasteParagraph).inlines[0]).toEqual({ text: 'x' });
    expect(generic.report.dropped).toEqual({ alignment: 1, fonts: 1, sizes: 1, colours: 1, highlights: 1 });
  });

  test('all: the whole document in one paste, the title by its class inside Word\'s border div', async () => {
    const { blocks, report } = convertHtml(await capture('all'), await lookup());
    expect((blocks[0] as PasteParagraph).style).toBe('Title');
    expect(report.lists).toBe(2);
    expect(report.tables).toBe(1);
    expect(blocks.filter((b) => b.kind === 'paragraph' && (b as PasteParagraph).list).length).toBe(10);
    expect(report.dropped).toEqual({ spacing: 9 });
    expect(blocks.some((b) => b.kind === 'paragraph' && (b as PasteParagraph).jc === 'center')).toBe(true);
  });
});

describe('paste through a view', () => {
  test('the lists get definitions, the table its nodes, everything in one undo step', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'tracked-changes.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    // At the end of block 4 ("Some text.").
    let pos = 0;
    for (let k = 0; k < 4; k++) pos += editor.state.doc.child(k).nodeSize;
    const end = pos + 1 + editor.state.doc.child(4).content.size;
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, end)));
    const before = editor.state.doc.childCount;
    const countTables = (doc: { forEach(f: (n: { type: { name: string } }) => void): void }): number => { let n = 0; doc.forEach((node) => { if (node.type.name === 'table') n++; }); return n; };
    const tablesBefore = countTables(editor.state.doc);   // the fixture has a table of its own
    const report = await pasteHtml(editor.view, session, await capture('all'));
    expect(report.lists).toBe(2);
    const doc = editor.state.doc;
    expect(doc.childCount).toBeGreaterThanOrEqual(before + 30);
    const paragraphs: { style: string | null; numId: string | null; ilvl: number | null; text: string }[] = [];
    doc.forEach((n) => { if (n.type.name === 'paragraph') paragraphs.push({ style: n.attrs['pStyle'] as string | null, numId: (n.attrs['numPr'] as { numId: string } | null)?.numId ?? null, ilvl: (n.attrs['numPr'] as { ilvl: number } | null)?.ilvl ?? null, text: n.textContent }); });
    const bullets = paragraphs.filter((p) => p.text.includes('bullet') && p.numId !== null);
    const numbers = paragraphs.filter((p) => p.text.includes('item') && p.numId !== null);
    expect(bullets.length).toBe(5);
    expect(numbers.length).toBe(5);
    expect(new Set(bullets.map((p) => p.numId)).size).toBe(1);
    expect(new Set(numbers.map((p) => p.numId)).size).toBe(1);
    expect(bullets[0]!.numId).not.toBe(numbers[0]!.numId);
    expect(numbers.map((p) => p.ilvl)).toEqual([0, 0, 1, 1, 0]);
    expect(paragraphs.some((p) => p.style === 'Heading1' && p.text.startsWith('1. Headings'))).toBe(true);
    // tracked-changes.docx has no Heading 2: the paste spliced it in from the defaults (section 11.11's route) and the paragraph carries it.
    expect(report.stylesWanted).toContain('Heading2');
    expect(paragraphs.some((p) => p.style === 'Heading2' && p.text === 'A second-level heading')).toBe(true);
    expect(session.pkg.propertyResolverOrUndefined?.getStyle('Heading2')).toBeDefined();
    expect(countTables(doc)).toBe(tablesBefore + 1);
    // The direct formatting reaches the nodes (section 8.6): the centred paragraph's attributes, the run's font mark.
    let centred: PMNode | undefined;
    doc.forEach((n) => { if (n.type.name === 'paragraph' && n.textContent.includes('keeps the text and drops the rest')) centred = n; });
    expect(centred?.attrs['jc']).toBe('center');
    expect(centred?.attrs['ind']).toEqual({ start: 720, end: null, firstLine: null, hanging: null });
    const font = centred?.firstChild?.marks.find((m) => m.type.name === 'font');
    expect(font?.attrs).toEqual({ name: 'Arial', size: 14, color: '#FF0000', highlight: null });
    // The paste is one undo step; the definitions are not undone (C2's rule).
    editor.run('history.undo');
    expect(editor.state.doc.childCount).toBe(before);
    // The export carries the paste (the reload after a save needs the browser or Node: jsdom's
    // serializer writes a namespace twice, so the browser suite does the round trip).
    editor.run('history.redo');
    const exported = session.exportTree(editor.state.doc);
    expect(exported.changed).toBe(true);
    const names = exported.content.map((el) => el.name.localPart);
    expect(names.filter((n) => n === 'tbl').length).toBe(tablesBefore + 1);
    expect(names.filter((n) => n === 'p').length).toBeGreaterThan(before);
    // The export carries the formatting as w:jc, w:ind and the run's rPr (the highlight by its name).
    const exportedCentred = exported.content.find((el) => JSON.stringify(el).includes('keeps the text and drops the rest')) as unknown as { value: { pPr?: { jc?: { val: string }; ind?: { left?: number; start?: number } }; content: { value: { rPr?: { rFonts?: { ascii?: string }; sz?: { val: number }; color?: { val: string } } } }[] } };
    expect(exportedCentred.value.pPr?.jc?.val).toBe('center');
    expect(exportedCentred.value.pPr?.ind?.start ?? exportedCentred.value.pPr?.ind?.left).toBe(720);
    const rPr = exportedCentred.value.content.find((c) => c.value.rPr)?.value.rPr;
    expect(rPr?.rFonts?.ascii).toBe('Arial');
    expect(rPr?.sz?.val).toBe(28);
    expect(rPr?.color?.val).toBe('FF0000');
    const exportedHighlight = exported.content.find((el) => JSON.stringify(el).includes('highlighted word')) as unknown as { value: { content: { value: { rPr?: { highlight?: { val: string } } } }[] } };
    expect(exportedHighlight.value.content.map((c) => c.value.rPr?.highlight?.val).filter(Boolean)).toEqual(['yellow']);
  });

  test('a note refuses a table and says so; the paragraphs go in', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'footnote.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const { EditableStoryView } = await import('../src/view/StoryView.js');
    const notes = await session.stories.footnotes();
    const note = notes[0]!;
    const host = document.createElement('div');
    document.body.appendChild(host);
    const story = new EditableStoryView(host, session, note, { id: 'test-note', stylesheet: editor });
    const report = await pasteHtml(story.view, session, await capture('table'));
    expect(report.tables).toBe(0);
    expect(report.dropped['tables in a note or header']).toBe(1);
    expect(story.state.doc.textContent).toContain('A three-by-three table with a header row:');
    let tables = 0;
    story.state.doc.forEach((n) => { if (n.type.name === 'table') tables++; });
    expect(tables).toBe(0);
    story.destroy();
  });
});

/** Konsole's capture (`fixtures/paste/terminal/`): its `text/html`, or the same copy's `text/plain`. */
async function terminal(kind: 'html' | 'txt'): Promise<string> {
  return readFile(join(fixtures, 'paste', 'terminal', `konsole-26.04.${kind}`), 'utf8');
}

/** A paste event carrying the given flavours, as a browser hands one to the view. */
function clipboard(data: Record<string, string>): ClipboardEvent {
  return { clipboardData: { getData: (type: string) => data[type] ?? '' }, preventDefault: () => {} } as unknown as ClipboardEvent;
}

/** A view over a small document, the caret in a new empty paragraph at the end. */
async function emptyParagraphAtEnd(name = 'HelloWordOnline.docx'): Promise<{ session: EditorDocument; editor: DocxEditorView }> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));
  const container = document.createElement('div');
  document.body.appendChild(container);
  const editor = new DocxEditorView(container, session);
  const end = editor.state.doc.content.size;
  const tr = editor.state.tr.insert(end, editor.state.schema.nodes['paragraph']!.create());
  editor.view.dispatch(tr.setSelection(TextSelection.create(tr.doc, end + 1)));
  return { session, editor };
}

const texts = (doc: PMNode): string[] => { const out: string[] = []; doc.forEach((n) => { if (n.type.name === 'paragraph') out.push(n.textContent); }); return out; };

describe("a terminal's HTML (ED-005 section 12.67)", () => {
  test("Konsole's capture is a terminal's; none of Word's captures is, nor a document's HTML", async () => {
    expect(isTerminalHtml(await terminal('html'))).toBe(true);
    for (const name of NAMES) expect(isTerminalHtml(await capture(name)), name).toBe(false);
    expect(isTerminalHtml('<span style="font-family:Arial">words</span>')).toBe(false);
    expect(isTerminalHtml('<div style="font-family:monospace">a block</div>')).toBe(false);
    expect(isTerminalHtml('<span style="font-family:monospace">mono</span> and not')).toBe(false);
    expect(isTerminalHtml('<span style="font-family:monospace"> </span>')).toBe(false);
    // The converter alone gathers the capture's loose spans into one paragraph (core-ts 58b23fb, CR-005 section 8.1
    // item 4; at 0.4.0 it made a paragraph of each span), which is why the paste takes the plain text instead.
    const { blocks } = convertHtml(await terminal('html'), styleLookupOf([]));
    const texts = blocks.map((b) => (b as PasteParagraph).inlines.map((i) => (i as PasteRun).text).join(''));
    expect(texts).toHaveLength(1);
    expect(texts[0]).toContain('--list-types');
  });

  test('pasted as its plain text: a paragraph per line, in no colour', async () => {
    const { editor } = await emptyParagraphAtEnd();
    const before = texts(editor.state.doc);
    const event = clipboard({ 'text/html': await terminal('html'), 'text/plain': await terminal('txt') });
    const handled = editor.view.someProp('handlePaste', (f) => f(editor.view, event, Slice.empty));
    expect(handled).toBe(true);
    const after = texts(editor.state.doc);
    // The copy ends with its last line's end, which ProseMirror's plain paste makes an empty paragraph, as for any text.
    expect(after.slice(before.length - 1)).toEqual(['wl-paste --list-types', 'STRING', 'text/plain;charset=utf-8', 'TEXT', 'text/plain', 'UTF8_STRING', '']);
    let marked = false;
    editor.state.doc.descendants((n) => { if (n.isText && n.marks.some((m) => m.type.name === 'font')) marked = true; return true; });
    expect(marked).toBe(false);
    // Not Markdown, so no offer.
    expect(markdownOfferOf(editor.state)).toBeNull();
  });

  test("a terminal's Markdown gets the offer", async () => {
    const { editor } = await emptyParagraphAtEnd();
    // Konsole's shape (the capture's), around a list.
    const run = (text: string): string => `<span style="color:#000000;background-color:#ffffff;">${text}</span>`;
    const html = `<html><body><span style="font-family:monospace">${run('- one')}<br>${run('- two')}<br></span></body></html>`;
    editor.view.someProp('handlePaste', (f) => f(editor.view, clipboard({ 'text/html': html, 'text/plain': '- one\n- two\n' }), Slice.empty));
    expect(markdownOfferOf(editor.state)?.text).toBe('- one\n- two\n');
  });
});

describe('links in HTML (ED-005 section 12.67)', () => {
  const HTML = '<p>See <a href="https://example.com/a"><span style="color:#0563C1;text-decoration:underline">the page</span></a>, <a href="mailto:a@example.com">mail</a>, <a href="#_Toc1">a heading</a> and <a name="_GoBack"></a>a mark.</p>';

  test("a web or mail link is kept; a link into its source is dropped and counted, its text kept; a bookmark is no link", () => {
    const { blocks, report } = convertHtml(HTML, styleLookupOf([]));
    const inlines = (blocks[0] as PasteParagraph).inlines as PasteRun[];
    expect(inlines.filter((i) => i.link).map((i) => [i.text, i.link])).toEqual([['the page', 'https://example.com/a'], ['mail', 'mailto:a@example.com']]);
    expect(inlines.map((i) => i.text).join('')).toBe('See the page, mail, a heading and a mark.');
    expect(report.dropped).toEqual({ links: 1 });
  });

  test("pasted: a relationship per address, Word's look (the Hyperlink style, no direct colour)", async () => {
    const { session, editor } = await emptyParagraphAtEnd();
    const report = await pasteHtml(editor.view, session, HTML);
    expect(report.dropped).toEqual({ links: 1 });
    const linked: { text: string; rId: string; style: string | undefined; font: boolean }[] = [];
    editor.state.doc.descendants((n) => {
      const link = n.marks.find((m) => m.type.name === 'link');
      if (n.isText && link) linked.push({ text: n.text!, rId: link.attrs['rId'] as string, style: n.marks.find((m) => m.type.name === 'runStyle')?.attrs['id'] as string | undefined, font: n.marks.some((m) => m.type.name === 'font') });
      return true;
    });
    expect(linked.map((l) => [l.text, l.style, l.font])).toEqual([['the page', 'Hyperlink', false], ['mail', 'Hyperlink', false]]);
    expect(linked[0]!.rId).not.toBe(linked[1]!.rId);
    expect(linked.every((l) => /^rId/.test(l.rId))).toBe(true);
    expect(session.pkg.propertyResolverOrUndefined?.getStyle('Hyperlink')).toBeDefined();
  });
});

describe("a ChatGPT answer (ED-005 section 12.101)", () => {
  const chatgpt = join(fixtures, 'paste', 'chatgpt');
  // The engine's walk at 0.4.0 made a paragraph of each word in the note after a table (a bare span holding a span
  // per word); core-ts 58b23fb gathers loose inline content into one paragraph (CR-005 section 8.1 item 4), and the
  // editor waited for the release that carries it (Jason, 2026-10-15: no fallback). Pinned to the engine's own conversion.
  test('converts as the engine pins it: the note after the table one paragraph', async () => {
    const html = await readFile(join(chatgpt, 'chatgpt-answer.html'), 'utf8');
    expect(isTerminalHtml(html)).toBe(false);
    const result = convertHtml(html, styleLookupOf([]));
    expect(JSON.parse(JSON.stringify(result))).toEqual(JSON.parse(await readFile(join(chatgpt, 'chatgpt-answer.expected.json'), 'utf8')));
  });
});

describe('a web address pasted as plain text is a link, its text the address (ED-005 section 12.121)', () => {
  const ADDRESS = 'https://example.com/docs/__init__?q=1#top';

  /** The text nodes of the document's last paragraph: their text, link target, character style, and whether they are bold or coloured. */
  function lastParagraph(session: EditorDocument, doc: PMNode): { text: string; target: string | null; style: string | null; bold: boolean; colour: unknown }[] {
    const out: { text: string; target: string | null; style: string | null; bold: boolean; colour: unknown }[] = [];
    doc.lastChild!.forEach((n) => {
      if (!n.isText) return;
      const link = n.marks.find((m) => m.type.name === 'link');
      out.push({
        text: n.text!,
        target: link ? hyperlinkTarget(session, link.attrs['rId'] as string) : null,
        style: (n.marks.find((m) => m.type.name === 'runStyle')?.attrs['id'] as string | undefined) ?? null,
        bold: n.marks.some((m) => m.type.name === 'bold'),
        colour: n.marks.find((m) => m.type.name === 'font')?.attrs['color'] ?? null,
      });
    });
    return out;
  }

  /** A plain paste through the view's handler, as ProseMirror makes it; whether the handler took it. */
  function pastePlain(editor: { view: DocxEditorView['view'] }, text: string): boolean {
    return editor.view.someProp('handlePaste', (f) => f(editor.view, clipboard({ 'text/plain': text }), Slice.empty)) === true;
  }

  const linked = async (editor: DocxEditorView): Promise<void> => {
    await vi.waitFor(() => { let any = false; editor.state.doc.descendants((n) => { if (n.marks.some((m) => m.type.name === 'link')) any = true; return !any; }); if (!any) throw new Error('not linked yet'); });
  };

  test('the rule: http:// or https:// in any case, then no whitespace at all', () => {
    for (const yes of [ADDRESS, 'http://x', 'HTTPS://EXAMPLE.COM/A', 'https://example.com/a(b)"c"<d>']) expect(pastedAddress(yes), yes).toBe(yes);
    for (const no of ['https://', 'http:/x', 'ftp://example.com', 'mailto:a@example.com', 'www.example.com', ' https://example.com', 'https://example.com\n', 'https://example.com/a b', 'https://example.com\tx', 'see https://example.com', 'https://example.com\u00a0']) expect(pastedAddress(no), JSON.stringify(no)).toBeUndefined();
  });

  test("at a caret: a link to the address in Word's look, the relationship made, no Markdown offer; what is typed next is plain", async () => {
    const { session, editor } = await emptyParagraphAtEnd();
    expect(session.pkg.propertyResolverOrUndefined?.getStyle('Hyperlink')).toBeUndefined();
    expect(pastePlain(editor, ADDRESS)).toBe(true);
    await linked(editor);
    expect(lastParagraph(session, editor.state.doc)).toEqual([{ text: ADDRESS, target: ADDRESS, style: 'Hyperlink', bold: false, colour: null }]);
    expect(session.pkg.propertyResolverOrUndefined?.getStyle('Hyperlink')).toBeDefined();
    // The address reads as Markdown (`__init__`), but a link is what it became.
    expect(markdownOfferOf(editor.state)).toBeNull();
    expect(editor.state.selection.empty && editor.state.selection.from === editor.state.doc.content.size - 1).toBe(true);
    editor.view.dispatch(editor.state.tr.insertText(' and on.'));
    expect(lastParagraph(session, editor.state.doc)).toEqual([
      { text: ADDRESS, target: ADDRESS, style: 'Hyperlink', bold: false, colour: null },
      { text: ' and on.', target: null, style: null, bold: false, colour: null },
    ]);
  });

  test("over a selection in bold, coloured text: the selection replaced, the link bold in the link's colour, typing bold and coloured again", async () => {
    const { session, editor } = await emptyParagraphAtEnd();
    const schema = editor.state.schema;
    const marks = [schema.marks['bold']!.create(), schema.marks['font']!.create({ name: null, size: null, color: 'FF0000', highlight: null })];
    const at = editor.state.selection.from;
    let tr = editor.state.tr.insert(at, schema.text('Read THIS now.', marks));
    tr = tr.setSelection(TextSelection.create(tr.doc, at + 5, at + 9));
    editor.view.dispatch(tr);
    expect(pastePlain(editor, 'http://example.org')).toBe(true);
    await linked(editor);
    expect(lastParagraph(session, editor.state.doc)).toEqual([
      { text: 'Read ', target: null, style: null, bold: true, colour: 'FF0000' },
      { text: 'http://example.org', target: 'http://example.org', style: 'Hyperlink', bold: true, colour: null },
      { text: ' now.', target: null, style: null, bold: true, colour: 'FF0000' },
    ]);
    expect(editor.state.storedMarks?.map((m) => m.type.name).sort()).toEqual(['bold', 'font']);
  });

  test("inside a link, by range.hyperlink's rule (Word, core-ts check 16): the old link goes from the text around, the address is the new link", async () => {
    const { session, editor } = await emptyParagraphAtEnd();
    expect(pastePlain(editor, 'https://one.example')).toBe(true);
    await linked(editor);
    const end = editor.state.selection.from;
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, end - '.example'.length)));
    expect(pastePlain(editor, 'https://two.example')).toBe(true);
    await vi.waitFor(() => expect(editor.state.doc.lastChild!.textContent).toBe('https://onehttps://two.example.example'));
    expect(lastParagraph(session, editor.state.doc)).toEqual([
      { text: 'https://one', target: null, style: null, bold: false, colour: null },
      { text: 'https://two.example', target: 'https://two.example', style: 'Hyperlink', bold: false, colour: null },
      { text: '.example', target: null, style: null, bold: false, colour: null },
    ]);
    expect(editor.state.storedMarks).toEqual([]);
  });

  test('not an address: text around it, a line end, another scheme: ProseMirror pastes the text', async () => {
    const { editor } = await emptyParagraphAtEnd();
    for (const text of ['see https://example.com', 'https://example.com\n', 'ftp://example.com', 'https://example.com and more']) expect(pastePlain(editor, text), text).toBe(false);
  });

  test("in a note it stays text: the note's part would need a relationship of its own", async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'footnote.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const { EditableStoryView } = await import('../src/view/StoryView.js');
    const host = document.createElement('div');
    document.body.appendChild(host);
    const story = new EditableStoryView(host, session, (await session.stories.footnotes())[0]!, { id: 'test-note', stylesheet: editor });
    story.view.dispatch(story.state.tr.setSelection(TextSelection.atEnd(story.state.doc)));
    story.view.pasteText(ADDRESS, clipboard({ 'text/plain': ADDRESS }));
    expect(story.state.doc.textContent).toContain(ADDRESS);
    let link = false;
    story.state.doc.descendants((n) => { if (n.marks.some((m) => m.type.name === 'link')) link = true; return true; });
    expect(link).toBe(false);
  });

  test("with tracking on, the link's text is this author's insertion, which the save writes inside the hyperlink", async () => {
    const { session, editor } = await emptyParagraphAtEnd();
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author A';
    expect(pastePlain(editor, ADDRESS)).toBe(true);
    await linked(editor);
    const node = editor.state.doc.lastChild!.firstChild!;
    expect(node.text).toBe(ADDRESS);
    expect(node.marks.find((m) => m.type.name === 'insertion')?.attrs['author']).toBe('Author A');
    const rId = node.marks.find((m) => m.type.name === 'link')!.attrs['rId'] as string;
    // this fixture's main part is word/document22.xml
    const files = unzipSync((await session.save({ selfCheck: false })).bytes);
    const xml = strFromU8(Object.entries(files).find(([name]) => /^word\/document\d*\.xml$/.test(name))![1]);
    const hyperlink = new RegExp(`<w:hyperlink [^>]*r:id="${rId}"[^>]*>(.*?)</w:hyperlink>`, 's').exec(xml)?.[1];
    expect(hyperlink).toMatch(/^<w:ins [^>]*w:author="Author A"[^>]*><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/><\/w:rPr><w:t>([^<]*)<\/w:t><\/w:r><\/w:ins>$/);
    expect(/<w:t>([^<]*)<\/w:t>/.exec(hyperlink!)![1]).toBe(ADDRESS);
  });
});

describe('a picture on the clipboard goes in as Insert picture puts one (ED-005 section 12.127)', () => {
  // jsdom's Blob has no arrayBuffer(), which every browser's has.
  Blob.prototype.arrayBuffer ??= function (this: Blob): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(reader.result as ArrayBuffer); reader.onerror = () => reject(reader.error); reader.readAsArrayBuffer(this); });
  };

  function pngBytes(w: number, h: number): Buffer {
    const crcTable = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
    const crc = (buf: Buffer): number => { let c = 0xffffffff; for (const b of buf) c = crcTable[(c ^ b) & 0xff]! ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
    const chunk = (type: string, data: Buffer): Buffer => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type, 'ascii'), data]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([len, td, c]); };
    const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
    const raw = Buffer.alloc((1 + w * 3) * h);
    return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
  }
  const png = (w = 120, h = 60): File => new File([Uint8Array.from(pngBytes(w, h))], 'image.png', { type: 'image/png' });

  /** A paste event as a browser hands one over: its flavours and its files. */
  function withFiles(data: Record<string, string>, files: File[]): ClipboardEvent {
    return { clipboardData: { getData: (type: string) => data[type] ?? '', files }, preventDefault: () => {} } as unknown as ClipboardEvent;
  }
  const paste = (view: DocxEditorView['view'], data: Record<string, string>, files: File[]): boolean =>
    view.someProp('handlePaste', (f) => f(view, withFiles(data, files), Slice.empty)) === true;
  const imagesIn = (doc: PMNode): PMNode[] => { const out: PMNode[] = []; doc.descendants((n) => { if (n.type.name === 'image') out.push(n); return true; }); return out; };

  test('the rule: image files, with no text in the HTML beside them', () => {
    const picture = png();
    const pdf = new File(['%PDF-1.7'], 'a.pdf', { type: 'application/pdf' });
    expect(clipboardPictures([picture], '')).toEqual([picture]);
    expect(clipboardPictures([pdf, picture], '')).toEqual([picture]);
    // A browser's Copy image: the file and the markup of one <img>.
    expect(clipboardPictures([picture], "<meta charset='utf-8'><img src=\"https://example.org/a.png\" alt=\"A picture\"/>")).toEqual([picture]);
    // Markup holding a picture in a paragraph, conditional comments and a no-break space, but no text.
    expect(clipboardPictures([picture], '<html><head><style>p { margin: 0 }</style></head><body><!--StartFragment--><p class=MsoNormal><span><!--[if gte vml 1]><v:shape></v:shape><![endif]--><img src="file:///C:/clip_image001.png"></span><o:p>&nbsp;</o:p></p><!--EndFragment--></body></html>')).toEqual([picture]);
    // Text beside the picture (a document's text and pictures, a spreadsheet's cells and a bitmap of them): HTML.
    expect(clipboardPictures([picture], '<p>Total</p><img src="x.png">')).toEqual([]);
    expect(clipboardPictures([pdf], '')).toEqual([]);
    expect(clipboardPictures([], '<img src="x.png">')).toEqual([]);
  });

  test('at a caret: an image from a new media part at its natural size, the caret after it', async () => {
    const { session, editor } = await emptyParagraphAtEnd();
    expect(imagesIn(editor.state.doc)).toHaveLength(0);
    expect(paste(editor.view, {}, [png(120, 60)])).toBe(true);
    await vi.waitFor(() => { if (imagesIn(editor.state.doc).length === 0) throw new Error('not pasted yet'); });
    const [image] = imagesIn(editor.state.doc);
    expect(editor.state.doc.lastChild!.firstChild).toBe(image);
    // 120 by 60 pixels at 96 dpi.
    expect([image!.attrs['width'], image!.attrs['height']]).toEqual([90, 45]);
    expect(editor.state.selection.empty && editor.state.selection.from === editor.state.doc.content.size - 1).toBe(true);
    const files = unzipSync((await session.save({ selfCheck: false })).bytes);
    expect(Object.keys(files).filter((name) => name.startsWith('word/media/'))).toEqual(['word/media/image1.png']);
    const xml = strFromU8(Object.entries(files).find(([name]) => /^word\/document\d*\.xml$/.test(name))![1]);
    expect(xml).toMatch(/<wp:inline[\s\S]*<wp:extent cx="1143000" cy="571500"\/>[\s\S]*<a:blip r:embed="rId\d+"/);
  });

  test('over a selection, with tracking on: the selection replaced, the picture this author\'s insertion', async () => {
    const { session, editor } = await emptyParagraphAtEnd();
    const at = editor.state.selection.from;
    let tr = editor.state.tr.insertText('before XX after', at);
    tr = tr.setSelection(TextSelection.create(tr.doc, at + 7, at + 9));
    editor.view.dispatch(tr);
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author A';
    expect(paste(editor.view, {}, [png()])).toBe(true);
    await vi.waitFor(() => { if (imagesIn(editor.state.doc).length === 0) throw new Error('not pasted yet'); });
    const [image] = imagesIn(editor.state.doc);
    expect((image!.attrs['revision'] as { insertion?: { author: string } } | null)?.insertion?.author).toBe('Author A');
    const deleted: string[] = [];
    editor.state.doc.lastChild!.forEach((n) => { if (n.isText && n.marks.some((m) => m.type.name === 'deletion')) deleted.push(n.text!); });
    expect(deleted).toEqual(['XX']);
  });

  test('beside text in HTML: the HTML is pasted, the picture dropped and counted', async () => {
    const { editor } = await emptyParagraphAtEnd();
    const html = '<p>Total</p><p><img src="file:///C:/clip_image001.png"></p>';
    expect(paste(editor.view, { 'text/html': html, 'text/plain': 'Total' }, [png()])).toBe(true);
    await vi.waitFor(() => { if (!texts(editor.state.doc).includes('Total')) throw new Error('not pasted yet'); });
    expect(imagesIn(editor.state.doc)).toHaveLength(0);
  });

  test('in a note: nothing goes in, and the status line says why', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'footnote.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const { EditableStoryView } = await import('../src/view/StoryView.js');
    const note = (await session.stories.footnotes())[0]!;
    const host = document.createElement('div');
    document.body.appendChild(host);
    const said: string[] = [];
    const story = new EditableStoryView(host, session, note, { id: 'test-note', stylesheet: editor, onPaste: (_report, text) => said.push(text) });
    story.view.dispatch(story.state.tr.setSelection(TextSelection.atEnd(story.state.doc)));
    const before = story.state.doc;
    expect(paste(story.view, {}, [png()])).toBe(true);
    await vi.waitFor(() => { if (said.length === 0) throw new Error('not said yet'); });
    expect(said).toEqual(["A picture can't be pasted here: put the caret in the document's text, outside notes, headers and footers"]);
    expect(story.state.doc.eq(before)).toBe(true);
  });

  test('a type Word does not read, where the browser cannot decode it (jsdom): not pasted, and said so', async () => {
    const said: string[] = [];
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'HelloWordOnline.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session, { onPaste: (_report, text) => said.push(text) });
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.atEnd(editor.state.doc)));
    expect(paste(editor.view, {}, [new File([new Uint8Array([0x52, 0x49, 0x46, 0x46])], 'a.webp', { type: 'image/webp' })])).toBe(true);
    await vi.waitFor(() => { if (said.length === 0) throw new Error('not said yet'); });
    expect(said).toEqual(['The picture on the clipboard could not be read (image/webp)']);
    expect(imagesIn(editor.state.doc)).toHaveLength(0);
  });
});
