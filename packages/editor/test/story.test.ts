// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 7.1: a story renders read-only with the body's node views, beside a body view
// over the same session, and reading it leaves the export untouched.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { EditorDocument } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { StoryView, EditableStoryView } from '../src/view/StoryView.js';
import { surfaces, type EditorSurface } from '../src/view/surface.js';
import { TextSelection } from 'prosemirror-state';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

// jsdom lays nothing out: ProseMirror's scroll-into-view after a typing command reads rectangles a
// Range has not got there, so both answer empty, as the other jsdom suites' commands never scroll.
const zeroRect = (): DOMRect => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON: () => ({}) }) as DOMRect;
if (typeof Range !== 'undefined') {
  Range.prototype.getClientRects ??= () => ([] as unknown as DOMRectList);
  Range.prototype.getBoundingClientRect ??= zeroRect;
}
if (typeof Element !== 'undefined') Element.prototype.getClientRects ??= () => ([] as unknown as DOMRectList);

/** A selection without `scrollIntoView` (a typing command scrolls anyway, hence the stubs above). */
function caret(surface: EditorSurface, from: number, to: number = from): void {
  surface.view.dispatch(surface.state.tr.setSelection(TextSelection.create(surface.state.doc, from, to)));
}

describe('StoryView', () => {
  test('a footnote renders its paragraphs read-only; the body export stays unchanged', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'footnote.docx'))));
    const body = new DocxEditorView(Object.assign(document.createElement('div'), {}), session);
    const notes = await session.stories.footnotes();
    const container = document.createElement('div');
    document.body.appendChild(container);
    const story = new StoryView(container, session, notes[0]!, { label: 'Footnote 1' });
    expect(container.classList.contains('story-view')).toBe(true);
    const pm = container.querySelector('.ProseMirror')!;
    expect(pm.getAttribute('contenteditable')).toBe('false');
    expect(pm.getAttribute('aria-label')).toBe('Footnote 1');
    expect(pm.querySelectorAll('p').length).toBe(notes[0]!.doc.childCount);
    expect(pm.textContent).toContain('THIRD');
    // The note's own reference mark is an opaque inline the stylesheet hides in a story.
    expect(pm.querySelector('.opaque-inline.kind-footnoteRef')).not.toBeNull();
    // No footnote numbering, page boundaries or section classes inside a story.
    expect(pm.querySelector('.page-boundary')).toBeNull();
    expect(pm.querySelector('p[class*="sect-"]')).toBeNull();
    story.update(notes[1]!);
    expect(pm.textContent).toContain('Footnote 1');
    expect(session.exportTree().changed).toBe(false);
    story.destroy();
    body.destroy();
  });
});

describe('EditableStoryView (ED-003 section 6.4)', () => {
  test('a keystroke in a footnote reaches its fragment and the session\'s undo; the body export stays unchanged; the toolbar\'s surface follows focus', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'footnote.docx'))));
    const activated: string[] = [];
    const bodyEl = document.createElement('div');
    document.body.appendChild(bodyEl);
    const body = new DocxEditorView(bodyEl, session, { headersAndFooters: false, onActivate: (s) => activated.push(s.id) });
    const notes = await session.stories.footnotes();
    const note = notes[0]!;
    const container = document.createElement('div');
    document.body.appendChild(container);
    const story = new EditableStoryView(container, session, note, { label: 'Footnote 1', stylesheet: body, onActivate: (s) => activated.push(s.id) });
    const pm = container.querySelector('.ProseMirror')!;
    expect(pm.getAttribute('contenteditable')).toBe('true');
    expect(pm.getAttribute('role')).toBe('textbox');
    expect(pm.getAttribute('aria-label')).toBe('Footnote 1');
    expect(surfaces.get('footnote:' + note.id)).toBe(story);
    expect(session.stories.hasFragment(note)).toBe(true);
    expect(session.undoManager.canUndo()).toBe(false);
    // A command through the surface: the fragment changes, the session's manager has a step, the body does not.
    const end = story.state.doc.firstChild!.nodeSize - 1;
    caret(story, end);
    expect(story.run('insertText', { text: ' (typed)' })).toBe(true);
    expect(session.stories.currentDoc(note).firstChild!.textContent).toContain('(typed)');
    expect(session.undoManager.canUndo()).toBe(true);
    expect(session.exportTree().changed).toBe(false);
    expect(story.formatting().font.bold).toBe(false);
    // Bold through the surface applies to the note.
    caret(story, end, end + 8);
    expect(story.run('font.bold', { value: 'toggle' })).toBe(true);
    expect(story.formatting().font.bold).toBe(true);
    // Undo through the keymap's command reaches the same manager.
    story.run('history.undo');
    expect(story.formatting().font.bold).toBe(false);
    story.run('history.undo');
    expect(session.stories.currentDoc(note).eq(note.doc)).toBe(true);
    // Focus reports the surface.
    story.focus();
    body.focus();
    expect(activated).toEqual(['footnote:' + note.id, 'body']);
    // Sections and page breaks refuse in a story; a line break does not.
    expect(story.run('section.insertBreak')).toBe(false);
    expect(story.run('insertBreak', { type: 'Page' })).toBe(false);
    expect(story.run('insertBreak', { type: 'Line' })).toBe(true);
    story.run('history.undo');
    // Destroying a story view leaves the session's undo manager alive (UPSTREAM-4).
    story.destroy();
    expect(surfaces.has('footnote:' + note.id)).toBe(false);
    expect(body.run('insertText', { text: 'x' })).toBe(true);
    expect(session.undoManager.canUndo()).toBe(true);
    body.destroy();
  });

  test('two views over one header converge, and the part is written once', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'headers.docx'))));
    const bodyEl = document.createElement('div');
    document.body.appendChild(bodyEl);
    const body = new DocxEditorView(bodyEl, session, { headersAndFooters: false });
    const header = (await session.stories.headers()).find((h) => h.kind === 'header' && h.doc.textContent.length > 0)!;
    const a = new EditableStoryView(document.createElement('div'), session, header, { id: 'a', stylesheet: body });
    const b = new EditableStoryView(document.createElement('div'), session, header, { id: 'b', stylesheet: body });
    caret(a, 1);
    a.run('insertText', { text: 'Both: ' });
    expect(b.state.doc.firstChild!.textContent.startsWith('Both: ')).toBe(true);
    const saved = await session.save({ selfCheck: false });   // the reload needs the node project (ED-003 section 11.8)
    expect(saved.changed).toBe(true);
    expect([...session.touchedParts]).toEqual([header.part.partName.toString()]);
    a.destroy();
    b.destroy();
    body.destroy();
  });

  test('the boundary widgets keep their header views through a body edit that mutes them, and the stylesheet covers the stories\' styles', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'odd-even-first-page.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    await until(() => container.querySelectorAll('.boundary-story').length >= 3);
    const before = container.querySelector('.page-boundary .boundary-header .ProseMirror')!;
    expect(before.getAttribute('contenteditable')).toBe('true');
    expect(container.querySelector('.page-boundary')!.classList.contains('stale')).toBe(false);
    // An edit at the top of the body: the boundary below turns stale in place, its header view still the same element.
    caret(editor, 1);
    editor.run('insertText', { text: 'Edited ' });
    expect(container.querySelector('.page-boundary')!.classList.contains('stale')).toBe(true);
    expect(container.querySelector('.page-boundary .boundary-label')!.textContent).toContain('approximate');
    expect(container.querySelector('.page-boundary .boundary-header .ProseMirror')).toBe(before);
    // The stylesheet has a rule for the header style once the header parts were read.
    const css = document.querySelector('style[data-docx4j-editor="document"]')!.textContent ?? '';
    expect(css).toContain('.s-Header');
    editor.destroy();
  });

  test('a page-break marker inside a table cell draws its boundary as a row between the table\'s rows, never inside the cell (ED-003 section 11.14)', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'large', 'anon-250.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const rows = container.querySelectorAll('tr.page-boundary');
    expect(rows.length).toBeGreaterThan(0);
    expect(container.querySelectorAll('td > .page-boundary, td > div.page-boundary').length).toBe(0);
    const first = rows[0]!;
    expect(first.parentElement!.tagName).toBe('TBODY');
    const td = first.firstElementChild as HTMLTableCellElement;
    expect(td.tagName).toBe('TD');
    // The band spans the table: as many columns as the table's own rows have cells for.
    const table = first.closest('table')!;
    const widest = Math.max(...Array.from(table.querySelectorAll(':scope > tbody > tr:not(.page-boundary)')).map((tr) => Array.from(tr.children).reduce((n, c) => n + ((c as HTMLTableCellElement).colSpan || 1), 0)));
    expect(td.colSpan).toBe(widest);
    expect(first.querySelector('.boundary-label')!.textContent).toMatch(/^Page \d+/);
    await until(() => first.querySelectorAll('.boundary-story').length >= 1);
    expect(first.querySelector('.boundary-footer .story-view, .boundary-header .story-view')).not.toBeNull();
    editor.destroy();
  }, 30_000);

  test('page numbers from w:pgNumType: the label, the PAGE field and the variant follow the shown number', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'notes-numbering.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const labels = [...container.querySelectorAll('.page-boundary .boundary-label')].map((e) => e.textContent);
    // Section 1 starts at 5 (one page); section 2 restarts at 1 in lower roman, with a marker for its second page.
    expect(labels).toEqual(['Page i', 'Page ii']);
    editor.destroy();
  });
});

async function until(check: () => boolean, ms = 3000): Promise<void> {
  const start = Date.now();
  while (!check()) {
    if (Date.now() - start > ms) throw new Error('timed out');
    await new Promise((r) => setTimeout(r, 20));
  }
}

describe('headers and footers at the page boundaries', () => {
  test('the first page shows the first-page header, an even page the even one, a boundary the footer above and the header below', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'odd-even-first-page.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    await until(() => container.querySelectorAll('.boundary-story').length >= 3);
    const start = container.querySelector('.page-start')!;
    expect(start.querySelector('.boundary-header .boundary-story')!.className).toContain('variant-first');
    expect(start.querySelector('.boundary-header .story-view')!.textContent).toContain('First page');
    const boundary = container.querySelector('.page-boundary')!;   // one marker: page 2 starts, an even page
    expect(boundary.querySelector('.boundary-footer .story-view')!.textContent).toContain('My footer');
    expect(boundary.querySelector('.boundary-header .boundary-story')!.className).toContain('variant-even');
    expect(boundary.querySelector('.boundary-header .story-view')!.textContent).toContain('Even header');
    expect(container.querySelector('.page-end .boundary-footer .story-view')!.textContent).toContain('My footer');
    expect(session.exportTree().changed).toBe(false);
    editor.destroy();
    expect(container.querySelectorAll('.boundary-story').length).toBe(0);
  });

  test('a page-starting section break is a boundary; a later section without references inherits ("same as previous")', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'sections-multiple.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    await until(() => container.querySelectorAll('.boundary-story').length >= 2);
    const boundaries = container.querySelectorAll('.page-boundary');
    expect(boundaries.length).toBe(1);   // no markers; one next-page section break
    expect(boundaries[0]!.textContent).toContain('Page 2');
    const header = boundaries[0]!.querySelector('.boundary-header .story-view');
    expect(header).not.toBeNull();
    expect(header!.textContent).toContain('Section 2');
    editor.destroy();
  });

  test('PAGE and NUMPAGES in a footer show the page the markers give, simple and complex fields alike', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'large', 'anon-250.docx'))));
    const [footer] = await session.stories.headers();
    const container = document.createElement('div');
    document.body.appendChild(container);
    const story = new StoryView(container, session, footer!, { fields: { PAGE: '7', NUMPAGES: '255' } });
    // What shows: the text less the hidden cached results (jsdom lays nothing out, so textContent includes them).
    const shown = container.querySelector('.ProseMirror')!.cloneNode(true) as HTMLElement;
    for (const hidden of shown.querySelectorAll('.field-hidden')) hidden.remove();
    expect(shown.textContent).toContain('7 / 255');
    expect(shown.textContent).not.toContain('338');
    expect(container.querySelector('.field-hidden')!.textContent).toBe('338');   // the complex PAGE field's cached result
    // Without known values the cached results show.
    const plain = new StoryView(document.createElement('div'), session, footer!);
    expect(plain.container.querySelector('.ProseMirror')!.textContent).toContain('338 / ');
    story.destroy();
    plain.destroy();
  }, 30_000);
});
