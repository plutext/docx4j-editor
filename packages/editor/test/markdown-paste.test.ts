// SPDX-License-Identifier: GPL-3.0-or-later
//
// Markdown paste (ED-005 section 12.66): a plain paste that reads as Markdown stays plain and leaves
// an offer; Format as Markdown replaces what it inserted as one undo step, Word's built-in styles
// spliced in where the document lacks them; the offer goes at this person's next edit.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, readMarkdown } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { markdownOfferOf, styleLookupOf, type PasteParagraph, type PasteTable } from '../src/view/paste.js';
import { formatAsMarkdown, markdownToPaste, offerHolds } from '../src/view/markdown.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

/** A paste event as a browser hands one over: `text/plain` alone, as from a terminal. */
function plainPaste(text: string): ClipboardEvent {
  return { clipboardData: { getData: (type: string) => (type === 'text/plain' ? text : '') }, preventDefault: () => {} } as unknown as ClipboardEvent;
}

async function editorOver(name: string, offers: (EditorView | null)[]): Promise<{ session: EditorDocument; editor: DocxEditorView }> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));
  const container = document.createElement('div');
  document.body.appendChild(container);
  const editor = new DocxEditorView(container, session, { onMarkdownOffer: (view) => offers.push(view) });
  // A caret in a new empty paragraph at the end.
  const end = editor.state.doc.content.size;
  const tr = editor.state.tr.insert(end, editor.state.schema.nodes['paragraph']!.create());
  editor.view.dispatch(tr.setSelection(TextSelection.create(tr.doc, end + 1)));
  return { session, editor };
}

const paragraphs = (doc: PMNode): PMNode[] => { const out: PMNode[] = []; doc.forEach((n) => { if (n.type.name === 'paragraph') out.push(n); }); return out; };

const MARKDOWN = '# Plan\n\nRun `npm ci` and read [the guide](https://example.com/guide).\n\n- one\n- two\n\n```\nnpm test\n```\n\n> Quoted.\n';

describe('the conversion', () => {
  test("Markdown's blocks as paste's, the built-ins the document lacks wanted", () => {
    const lookup = styleLookupOf([{ id: 'Heading1', name: 'heading 1' } as never], []);
    const { blocks, report } = markdownToPaste(readMarkdown(MARKDOWN + '\n| a | b |\n|---|---|\n| 1 |\n', { lenient: true }), lookup);
    expect(blocks.filter((b): b is PasteParagraph => b.kind === 'paragraph').map((b) => [b.style, b.list?.bullet ?? null])).toEqual([
      ['Heading1', null], [null, null], [null, true], [null, true], ['HTMLPreformatted', null], ['Quote', null],
    ]);
    expect((blocks[1] as PasteParagraph).inlines).toEqual([
      { text: 'Run ' }, { text: 'npm ci', style: 'HTMLCode' }, { text: ' and read ' }, { text: 'the guide', link: 'https://example.com/guide' }, { text: '.' },
    ]);
    // A short row is filled out; the first row is the header.
    const table = blocks[6] as PasteTable;
    expect(table.rows.map((r) => r.length)).toEqual([2, 2]);
    expect(table.headerRow).toBe(true);
    expect(report).toMatchObject({ lists: 1, tables: 1, dropped: {} });
    expect(report.stylesWanted.sort()).toEqual(['HTMLCode', 'HTMLPreformatted', 'Quote']);
  });
});

describe('the offer and Format as Markdown', () => {
  test('a plain paste that reads as Markdown stays plain and leaves an offer; Format replaces it as one undo step', async () => {
    const offers: (EditorView | null)[] = [];
    const { session, editor } = await editorOver('HelloWordOnline.docx', offers);
    const before = paragraphs(editor.state.doc).length;
    editor.view.pasteText(MARKDOWN, plainPaste(MARKDOWN));
    expect(paragraphs(editor.state.doc).some((p) => p.textContent === '# Plan')).toBe(true);
    expect(offers.at(-1)).toBe(editor.view);
    const offer = markdownOfferOf(editor.state)!;
    expect(offer.text).toBe(MARKDOWN);
    expect(offerHolds(editor.state.doc, offer)).toBe(true);

    const report = await formatAsMarkdown(editor.view, session, markdownOfferOf);
    expect(report).toMatchObject({ paragraphs: 6, lists: 1, tables: 0, dropped: {} });
    expect(markdownOfferOf(editor.state)).toBeNull();
    expect(offers.at(-1)).toBeNull();
    const made = paragraphs(editor.state.doc).slice(before - 1);
    expect(made.map((p) => [p.attrs['pStyle'], p.attrs['numPr'] ? 'list' : null, p.textContent])).toEqual([
      ['Heading1', null, 'Plan'],
      [null, null, 'Run npm ci and read the guide.'],
      [null, 'list', 'one'],
      [null, 'list', 'two'],
      ['HTMLPreformatted', null, 'npm test'],
      ['Quote', null, 'Quoted.'],
    ]);
    const marksOf = (text: string): string[] => {
      let found: string[] = [];
      made[1]!.forEach((n) => { if (n.text === text) found = n.marks.map((m) => (m.type.name === 'runStyle' ? `runStyle:${m.attrs['id']}` : m.type.name)); });
      return found;
    };
    expect(marksOf('npm ci')).toEqual(['runStyle:HTMLCode']);
    expect(marksOf('the guide')).toEqual(['runStyle:Hyperlink', 'link']);
    // Word's built-ins the document lacked are in its styles part now.
    for (const id of ['HTMLCode', 'HTMLPreformatted', 'Quote']) expect(session.pkg.propertyResolverOrUndefined?.getStyle(id), id).toBeDefined();

    // Undo gives the plain text back, not the state before the paste.
    editor.run('history.undo');
    expect(paragraphs(editor.state.doc).some((p) => p.textContent === '# Plan')).toBe(true);
    expect(paragraphs(editor.state.doc).some((p) => p.attrs['pStyle'] === 'HTMLPreformatted')).toBe(false);
  });

  test('the offer goes at the next edit, and text that is not Markdown makes none', async () => {
    const offers: (EditorView | null)[] = [];
    const { session, editor } = await editorOver('HelloWordOnline.docx', offers);
    editor.view.pasteText('Just a sentence.', plainPaste('Just a sentence.'));
    expect(markdownOfferOf(editor.state)).toBeNull();
    expect(offers).toEqual([]);
    editor.view.pasteText('- a\n- b', plainPaste('- a\n- b'));
    expect(markdownOfferOf(editor.state)).not.toBeNull();
    editor.view.dispatch(editor.state.tr.insertText('x'));
    expect(markdownOfferOf(editor.state)).toBeNull();
    expect(offers.at(-1)).toBeNull();
    expect(await formatAsMarkdown(editor.view, session, markdownOfferOf)).toBeNull();
  });

  test("in a note, a link keeps its text and is counted (the note's part would need a relationship of its own)", async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'footnote.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const { EditableStoryView } = await import('../src/view/StoryView.js');
    const note = (await session.stories.footnotes())[0]!;
    const host = document.createElement('div');
    document.body.appendChild(host);
    const offers: (EditorView | null)[] = [];
    const story = new EditableStoryView(host, session, note, { id: 'test-note', stylesheet: editor, onMarkdownOffer: (view) => offers.push(view) });
    story.view.dispatch(story.state.tr.setSelection(TextSelection.atEnd(story.state.doc)));
    const text = 'See **[the guide](https://example.com)** and `npm ci`.';
    story.view.pasteText(text, plainPaste(text));
    expect(offers.at(-1)).toBe(story.view);
    const report = await formatAsMarkdown(story.view, session, markdownOfferOf);
    expect(report?.dropped).toEqual({ links: 1 });
    let linked = false;
    let bold = '';
    story.state.doc.descendants((n) => {
      if (n.marks.some((m) => m.type.name === 'link')) linked = true;
      if (n.isText && n.marks.some((m) => m.type.name === 'bold')) bold += n.text;
      return true;
    });
    expect(linked).toBe(false);
    expect(bold).toBe('the guide');
    story.destroy();
  });

  test('with tracking on, the formatted text is the paste: one insertion, the plain text gone for good', async () => {
    const offers: (EditorView | null)[] = [];
    const { session, editor } = await editorOver('HelloWordOnline.docx', offers);
    await editor.runAsync('document.setChangeTrackingMode', { mode: 'TrackAll' });
    editor.view.pasteText('## Two\n\n**bold**', plainPaste('## Two\n\n**bold**'));
    const report = await formatAsMarkdown(editor.view, session, markdownOfferOf);
    expect(report).not.toBeNull();
    const deleted: string[] = [];
    const inserted: string[] = [];
    editor.state.doc.descendants((n) => {
      if (n.isText && n.marks.some((m) => m.type.name === 'deletion')) deleted.push(n.text!);
      if (n.isText && n.marks.some((m) => m.type.name === 'insertion')) inserted.push(n.text!);
      return true;
    });
    expect(deleted).toEqual([]);
    expect(inserted).toEqual(['Two', 'bold']);
    expect(paragraphs(editor.state.doc).some((p) => p.attrs['pStyle'] === 'Heading2' && p.textContent === 'Two')).toBe(true);
  });
});

