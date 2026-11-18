// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 5.6 step C3: the toolbar renders from the model; pressed state follows the
// effective formatting; a button runs its command through the view.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { act, Profiler } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { TextSelection } from 'prosemirror-state';
import { EditorDocument, documentStyles, documentFonts } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { Toolbar } from '../src/app/Toolbar.js';
import { consoleLine } from '../src/app/Palette.js';
import { commandNamed } from '@docx4j/editor-model';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

async function mount(name: string): Promise<{ editor: DocxEditorView; root: Root; host: HTMLElement }> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));
  const container = document.createElement('div');
  document.body.appendChild(container);
  const editor = new DocxEditorView(container, session);
  const styles = await documentStyles(session);
  const fonts = await documentFonts(session);
  const host = document.createElement('div');
  document.body.appendChild(host);
  const root = createRoot(host);
  await act(async () => { root.render(<Toolbar editor={editor} styles={styles} fonts={fonts} />); });
  return { editor, root, host };
}

function blockStart(editor: DocxEditorView, i: number): number {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += editor.state.doc.child(k).nodeSize;
  return pos;
}

describe('Toolbar', () => {
  test('bold is pressed on a heading (effective) and not on a plain paragraph; the picker shows the style', async () => {
    const { editor, host, root } = await mount('tracked-changes.docx');
    const bold = host.querySelector<HTMLButtonElement>('button[data-command="font.bold"]')!;
    const picker = host.querySelector<HTMLSelectElement>('[data-testid="style-picker"]')!;
    // The caret starts in the Title paragraph; move it into the heading (block 2).
    await act(async () => { editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, blockStart(editor, 2) + 1))); });
    const styleBold = editor.session.pkg.propertyResolver.getEffectiveRPr('Heading1');
    expect(bold.getAttribute('aria-pressed')).toBe(String(styleBold?.b !== undefined && styleBold.b.val !== false));
    expect(picker.value).toBe('Heading1');
    await act(async () => { editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, blockStart(editor, 4) + 1))); });
    expect(bold.getAttribute('aria-pressed')).toBe('false');
    expect(picker.value).toBe('');
    root.unmount();
    editor.destroy();
  });

  test('a button runs its command: alignment centred writes w:jc', async () => {
    const { editor, host, root } = await mount('tracked-changes.docx');
    await act(async () => { editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, blockStart(editor, 4) + 1))); });
    const centre = host.querySelector<HTMLButtonElement>('button[aria-label="Centre"]')!;
    expect(centre.getAttribute('aria-pressed')).toBe('false');
    await act(async () => { centre.click(); });
    expect(centre.getAttribute('aria-pressed')).toBe('true');
    expect(editor.state.doc.child(4).attrs['jc']).toBe('center');
    expect(editor.session.exportTree().changed).toBe(true);
    root.unmount();
    editor.destroy();
  });

  test('the font box resolves a theme font through the selector', async () => {
    const { editor, host, root } = await mount('tracked-changes.docx');
    const font = host.querySelector<HTMLSelectElement>('[data-testid="font-name"]')!;
    expect(font.value).not.toBe('');
    root.unmount();
    editor.destroy();
  });
});

describe('the toolbar renders from a memoised reading (ED-003 section 8.2)', () => {
  test('a keystroke that changes no reading re-renders nothing; a formatting change does', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'tracked-changes.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const styles = await documentStyles(session);
    const fonts = await documentFonts(session);
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    let renders = 0;
    const onRender = (): void => { renders++; };
    await act(async () => { root.render(<Profiler id="toolbar" onRender={onRender}><Toolbar editor={editor} styles={styles} fonts={fonts} /></Profiler>); });
    await act(async () => { editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, blockStart(editor, 4) + 3))); });
    // The first keystroke enables Undo, a reading change of its own; the ones after it change nothing.
    await act(async () => { editor.run('insertText', { text: '0' }); });
    const settled = renders;
    await act(async () => { editor.run('insertText', { text: 'a' }); });
    await act(async () => { editor.run('insertText', { text: 'b' }); });
    expect(renders).toBe(settled);
    await act(async () => { editor.run('font.bold', { value: true }); });
    expect(renders).toBeGreaterThan(settled);   // the formatting, then the undo button's enabled state as Yjs records the step
    expect(host.querySelector('button[data-command="font.bold"]')!.getAttribute('aria-pressed')).toBe('true');
    root.unmount();
  });
});

describe('the palette\'s console line', () => {
  test('an editor-only command is editor.run; a portable one with its arguments is its api line; placeholders fall back to editor.run', () => {
    expect(consoleLine(commandNamed('history.undo')!, {})).toBe('editor.run("history.undo");');
    expect(consoleLine(commandNamed('paragraph.alignment')!, { value: 'Centered' })).toBe('for (const p of selection.paragraphs) p.alignment = "Centered";');
    expect(consoleLine(commandNamed('font.size')!, { value: '14' })).toBe('selection.font.size = 14;');
    // A cleared value has no api form: the editor's line.
    expect(consoleLine(commandNamed('font.size')!, { value: '0' })).toBe('editor.run("font.size", { value: 0 });');
    expect(consoleLine(commandNamed('paragraph.alignment')!, {})).toMatch(/^editor\.run\("paragraph\.alignment", \{ value: undefined \/\* .*Centered.* \*\/ \}\);$/);
  });
});
