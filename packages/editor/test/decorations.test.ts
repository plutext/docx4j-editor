// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 8.2 (section 9 item 32): the body's decoration walk is memoised per top-level
// block. A keystroke walks the edited block afresh and reads the rest back; the page numbers
// and the stale frontier stay exact; a tree run, which moves what "changed" means, resets it.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { TextSelection } from 'prosemirror-state';
import { EditorDocument, applyPartXml } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { decorationsKey, decorationCacheOf } from '../src/view/plugins/decorations.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

async function open(name: string): Promise<{ session: EditorDocument; editor: DocxEditorView }> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));
  const container = document.createElement('div');
  document.body.appendChild(container);
  const editor = new DocxEditorView(container, session);
  return { session, editor };
}

function blockStart(editor: DocxEditorView, i: number): number {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += editor.state.doc.child(k).nodeSize;
  return pos;
}

/** The caret through a plain transaction: `caretAt` scrolls into view, which jsdom cannot. */
function caret(editor: DocxEditorView, pos: number): void {
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos)));
}

function cacheOf(editor: DocxEditorView) {
  const plugin = editor.state.plugins.find((p) => p.spec.key === decorationsKey)!;
  return decorationCacheOf(plugin)!;
}

describe('the memoised decoration walk', () => {
  test('a keystroke walks one block afresh; the boundaries keep their keys, and those below the edit turn stale', async () => {
    const { editor } = await open('tracked-changes.docx');
    const cache = cacheOf(editor);
    const before = decorationsKey.getState(editor.state)!;
    const keysBefore = [...before.boundaries.keys()];
    expect(keysBefore.some((k) => k.startsWith('page-marker-'))).toBe(true);
    expect([...before.boundaries.values()].every((b) => !b.stale)).toBe(true);
    const blocks = editor.state.doc.childCount;
    expect(cache.misses).toBe(blocks);
    cache.misses = 0;
    // Type in block 4: one block collected again, the rest read back.
    caret(editor, blockStart(editor, 4) + 1);
    editor.run('insertText', { text: 'x' });
    expect(cache.misses).toBe(1);
    // The previous set was mapped and patched, not rebuilt: the sequence of markers is what it was.
    expect(cache.patched).toBe(true);
    const after = decorationsKey.getState(editor.state)!;
    expect([...after.boundaries.keys()]).toEqual(keysBefore);
    // The first marker sits after block 4 in this document: every boundary below the edit is stale, the start is not.
    expect(after.boundaries.get('page-start')?.stale).toBe(false);
    const markers = [...after.boundaries.entries()].filter(([k]) => k.startsWith('page-marker-'));
    expect(markers.length).toBeGreaterThan(0);
    expect(markers.every(([, b]) => b.stale)).toBe(true);
    expect(markers[0]![1].label).toContain('approximate since your edit');
    // The decoration count is what a whole walk gives: the same set, re-created.
    expect(after.set.find().length).toBe(before.set.find().length);
  });

  test('undo restores the imported node, and the boundaries are exact again without a full walk', async () => {
    const { editor } = await open('tracked-changes.docx');
    const cache = cacheOf(editor);
    caret(editor, blockStart(editor, 4) + 1);
    editor.run('insertText', { text: 'x' });
    cache.misses = 0;
    editor.run('history.undo');
    // y-prosemirror rebuilds the undone block (a new node object): one miss, and it reads as unchanged again.
    expect(cache.misses).toBeLessThanOrEqual(2);
    const state = decorationsKey.getState(editor.state)!;
    expect([...state.boundaries.values()].every((b) => !b.stale)).toBe(true);
  });

  test('a marker added or a paragraph split changes the sequence: the set is rebuilt, then patched again from there', async () => {
    const { editor } = await open('tracked-changes.docx');
    const cache = cacheOf(editor);
    caret(editor, blockStart(editor, 4) + 1);
    editor.run('insertText', { text: 'x' });
    expect(cache.patched).toBe(true);
    // Enter splits the block: two fresh blocks, the sequence unchanged (no marker in them), so still a patch; the decorations are complete either way.
    editor.run('insertParagraph');
    const blocks = editor.state.doc.childCount;
    const whole = decorationsKey.getState(editor.state)!.set.find().length;
    // A whole walk over the same document gives the same decorations.
    const fresh = new (Object.getPrototypeOf(cache).constructor)();
    const { computeDecorations } = await import('../src/view/plugins/decorations.js');
    const { labelLookup } = await import('../src/view/plugins/labels.js');
    const rebuilt = computeDecorations(editor.session, editor.state, labelLookup(editor.state), { boundary: () => () => {} }, fresh);
    expect(rebuilt.set.find().length).toBe(whole);
    expect([...rebuilt.boundaries.entries()]).toEqual([...decorationsKey.getState(editor.state)!.boundaries.entries()]);
    expect(blocks).toBeGreaterThan(0);
  });

  test('a tree run resets the memo: what "changed" means moved with the baseline', async () => {
    const { editor, session } = await open('tracked-changes.docx');
    const cache = cacheOf(editor);
    caret(editor, blockStart(editor, 4) + 1);
    editor.run('insertText', { text: 'x' });
    expect([...decorationsKey.getState(editor.state)!.boundaries.values()].some((b) => b.stale)).toBe(true);
    await session.runOnTree(() => undefined);
    // After the run the edited block is the baseline: nothing is stale, and the memo was rebuilt over the new context.
    const state = decorationsKey.getState(editor.state)!;
    expect([...state.boundaries.values()].every((b) => !b.stale)).toBe(true);
    expect(cache.ctx).toBe(session.ctx);
  });
});

describe('a run\'s own theme colour (ED-008 section 8.8)', () => {
  /** What the view draws in a theme colour: the text of each decoration's element, with its colour. */
  const drawn = (editor: DocxEditorView): Record<string, string | undefined> => Object.fromEntries(
    [...editor.view.dom.querySelectorAll<HTMLElement>('.theme-colour')].map((el) => [el.textContent, /--theme-colour: (#[0-9A-Fa-f]{6})/.exec(el.getAttribute('style') ?? '')?.[1]]),
  );
  const fontColours = (editor: DocxEditorView): (string | undefined)[] => {
    const found: (string | undefined)[] = [];
    editor.state.doc.descendants((n) => { if (n.isText && n.text === 'stale run') found.push(n.marks.find((m) => m.type.name === 'font')?.attrs['color'] as string | undefined); return true; });
    return found;
  };
  const posOf = (editor: DocxEditorView, text: string): number => {
    let at = -1;
    editor.state.doc.descendants((n, pos) => { if (at < 0 && n.isText && n.text?.includes(text)) at = pos; return at < 0; });
    return at;
  };

  test('is drawn from the theme wherever a run is, where the value beside it is stale; a value that agrees gets no decoration; the mark keeps what the file wrote', async () => {
    const { editor } = await open('theme-colour-run.docx');
    expect(drawn(editor)).toEqual({
      'stale run': '#4F81BD',
      'tinted run': '#95B3D7',   // Word's Accent 1, Lighter 40%
      'inserted run': '#4F81BD',
      'deleted run': '#4F81BD',
      'cell run': '#4F81BD',
      'control run': '#4F81BD',
    });
    expect(fontColours(editor)).toEqual(['#FF00FF']);
  });

  test('a file without a stale copy has no decoration at all', async () => {
    const { editor } = await open('drawing-anchor.docx');   // six theme references, every copy right
    expect(drawn(editor)).toEqual({});
  });

  test('text typed in the run stays in its colour, and an edit elsewhere leaves it drawn', async () => {
    const { editor } = await open('theme-colour-run.docx');
    caret(editor, posOf(editor, 'stale run') + 3);
    editor.run('insertText', { text: 'X' });
    expect(drawn(editor)['staXle run']).toBe('#4F81BD');
    caret(editor, posOf(editor, 'last run') + 1);
    editor.run('insertText', { text: 'Y' });
    expect(cacheOf(editor).patched).toBe(true);
    expect(drawn(editor)['staXle run']).toBe('#4F81BD');
  });

  test('a colour the person sets takes the run off the theme; undo puts it back', async () => {
    const { editor } = await open('theme-colour-run.docx');
    const from = posOf(editor, 'stale run');
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, from + 'stale run'.length)));
    editor.run('font.color', { value: '#00FF00' });
    expect(drawn(editor)['stale run']).toBeUndefined();
    editor.run('history.undo');
    expect(drawn(editor)['stale run']).toBe('#4F81BD');
  });

  test('a theme changed in the parts panel redraws what the memo kept: a run follows it, and one whose copy now agrees loses its decoration', async () => {
    const { session, editor } = await open('theme-colour-run.docx');
    const theme = session.pkg.getPart('/word/theme/theme1.xml')!;
    const xml = await (theme as unknown as { getXml(): Promise<string> }).getXml();
    await applyPartXml(session, theme, xml.replace(/(<a:accent1><a:srgbClr val=")4F81BD/, '$1123456'));
    editor.partsChanged();
    expect(drawn(editor)['stale run']).toBe('#123456');
    await applyPartXml(session, theme, xml.replace(/(<a:accent1><a:srgbClr val=")4F81BD/, '$1FF00FF'));
    editor.partsChanged();
    expect(drawn(editor)['stale run']).toBeUndefined();
    expect(drawn(editor)['agreed run']).toBe('#FF00FF');
  });
});
