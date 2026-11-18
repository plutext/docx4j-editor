// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 5.6 (step C5) in the view: the flow of a measurement. jsdom lays nothing out,
// so every tab reads as starting at the column's edge and takes the first stop's width; what is
// checked is when a measurement runs (at mount, after a change in a paragraph with a tab, after
// a stylesheet refresh) and that its widths reach the atoms as node decorations.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { TextSelection } from 'prosemirror-state';
import { EditorDocument } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { tabStopsKey } from '../src/view/plugins/tabstops.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');
const tick = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 30));

async function open(name: string): Promise<{ session: EditorDocument; editor: DocxEditorView }> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));
  const container = document.createElement('div');
  document.body.appendChild(container);
  return { session, editor: new DocxEditorView(container, session, { headersAndFooters: false }) };
}

const measured = (editor: DocxEditorView): number => tabStopsKey.getState(editor.state)!.measured;

describe('tab widths in the view', () => {
  test('measured once at mount, again after a change in a paragraph with a tab, not after one elsewhere, and after a stylesheet refresh', async () => {
    const { editor } = await open('headers-footers.docx');
    expect(measured(editor)).toBe(0);
    await tick();
    expect(measured(editor)).toBe(1);
    const tabs = editor.container.querySelectorAll('.ProseMirror .tab');
    expect(tabs.length).toBeGreaterThan(0);
    // The first stop from the edge: the default interval, 720 twips, 48px.
    expect((tabs[0] as HTMLElement).style.inlineSize).toBe('48px');
    // A paragraph with a tab, and one without.
    const doc = editor.state.doc;
    const find = (withTab: boolean): number => {
      let found = -1;
      doc.descendants((node, pos) => {
        if (found >= 0) return false;
        if (node.type.name !== 'paragraph') return true;
        let has = false;
        node.forEach((c) => { if (c.type.name === 'tab') has = true; });
        if (has === withTab && node.childCount > 0) found = pos + 1;
        return false;
      });
      return found;
    };
    const plain = find(false);
    const tabbed = find(true);
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(doc, plain)));
    editor.run('insertText', { text: 'x' });
    await tick();
    expect(measured(editor)).toBe(1);
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, tabbed < plain ? tabbed : tabbed + 1)));
    editor.run('insertText', { text: 'y' });
    await tick();
    expect(measured(editor)).toBe(2);
    // A tab typed into a paragraph that had none: measured too.
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, plain)));
    editor.run('insertTab');
    await tick();
    expect(measured(editor)).toBe(3);
    expect((editor.container.querySelector('.ProseMirror p .tab') as HTMLElement).style.inlineSize).toBe('48px');
    editor.refreshCss();
    await tick();
    expect(measured(editor)).toBe(4);
    editor.destroy();
  });

  test('a stop set on a paragraph (paragraph.tabStops, ED-003 section 8.5) measures it again, and the tab takes the stop\'s width', async () => {
    const { editor } = await open('headers-footers.docx');
    await tick();
    expect(measured(editor)).toBe(1);
    let tabbed = -1;
    editor.state.doc.descendants((node, pos) => {
      if (tabbed >= 0) return false;
      if (node.type.name !== 'paragraph') return true;
      node.forEach((c) => { if (c.type.name === 'tab') tabbed = pos + 1; });
      return false;
    });
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, tabbed)));
    expect(editor.run('paragraph.tabStops', { stops: [{ position: 100, alignment: 'right', leader: 'dot' }] })).toBe(true);
    await tick();
    expect(measured(editor)).toBe(2);
    // jsdom reads every tab as starting at the edge: the first tab of that paragraph is as wide as the stop, 100pt, 133.33px, with its leader.
    const tab = editor.view.nodeDOM(tabbed) as HTMLElement;
    expect(tab.style.inlineSize).toBe('133.33px');
    expect(tab.dataset['leader']).toBe('dot');
    editor.destroy();
  });

  test('a document without a tab never measures', async () => {
    const { editor } = await open('tracked-changes.docx');
    await tick();
    expect(measured(editor)).toBe(0);
    expect(tabStopsKey.getState(editor.state)!.pending).toBeNull();
    editor.destroy();
  });
});
