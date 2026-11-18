// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 8.6 (step F5): the read-only elements grouped by kind with the selection
// filter, and the folding sections with their state remembered.
import { describe, test, expect, beforeEach } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { TextSelection } from 'prosemirror-state';
import { EditorDocument } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { DiagnosticList, groupByKind, sourceIdsInSelection } from '../src/app/panels/DiagnosticList.js';
import { Fold, loadFold, saveFold } from '../src/app/panels/Fold.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

async function open(name: string): Promise<{ session: EditorDocument; editor: DocxEditorView }> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));
  const container = document.createElement('div');
  document.body.appendChild(container);
  return { session, editor: new DocxEditorView(container, session) };
}

describe('the diagnostics list', () => {
  test('groups the entries by kind, the most frequent first', async () => {
    const { session } = await open('loadAndSave.docx');
    const groups = groupByKind(session.diagnostics.opaque);
    expect(groups.map((g) => [g.kind, g.entries.length])).toEqual([['fldChar', 15], ['instrText', 5], ['bookmarkEnd', 4], ['bookmarkStart', 4], ['commentReference', 1], ['drawing', 1], ['oMathPara', 1]]);
    expect(groups.reduce((n, g) => n + g.entries.length, 0)).toBe(session.diagnostics.opaque.length);
  });

  test('renders a row per kind with its count, opened to the addresses; the filter keeps what the selection touches', async () => {
    const { session, editor } = await open('loadAndSave.docx');
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    await act(async () => { root.render(<DiagnosticList editor={editor} entries={session.diagnostics.opaque} />); });
    const rows = host.querySelectorAll<HTMLDetailsElement>('details.diagnostic-kind');
    expect(rows.length).toBe(7);
    expect(rows[0]!.dataset['kind']).toBe('fldChar');
    expect(rows[0]!.querySelector('summary')!.textContent).toBe('fldChar 15');
    expect(rows[0]!.querySelectorAll('li').length).toBe(15);
    expect(rows[0]!.open).toBe(false);
    // An address is a button that selects the node, named by its kind and address. The anchored text
    // box is the drawing inside its kept mc:AlternateContent's selected branch (ED-009).
    const drawing = host.querySelector<HTMLButtonElement>('details[data-kind="drawing"] button.address')!;
    expect(drawing.getAttribute('aria-label')).toMatch(/^drawing at body\//);
    await act(async () => { drawing.click(); });
    const entry = session.diagnostics.opaque.find((e) => e.kind === 'drawing')!;
    expect(sourceIdsInSelection(editor.state).has(entry.sourceId)).toBe(true);
    // The filter: the selection touches the drawing and the run around it, not the fields elsewhere.
    const filter = host.querySelector<HTMLInputElement>('[data-testid="diagnostic-list-in-selection"]')!;
    await act(async () => { filter.click(); });
    const shown = Array.from(host.querySelectorAll<HTMLDetailsElement>('details.diagnostic-kind')).map((d) => d.dataset['kind']);
    expect(shown).toContain('drawing');
    expect(shown).not.toContain('fldChar');
    expect(shown).not.toContain('oMathPara');
    // A caret at a bookmark's edge counts the marker beside it, which cannot be selected by hand.
    const bookmark = session.diagnostics.opaque.find((e) => e.kind === 'bookmarkStart')!;
    let at = -1;
    editor.state.doc.descendants((node, pos) => { if (node.attrs['sourceId'] === bookmark.sourceId) at = pos; return at < 0; });
    expect(at).toBeGreaterThan(0);
    await act(async () => { editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at))); });
    const after = Array.from(host.querySelectorAll<HTMLDetailsElement>('details.diagnostic-kind')).map((d) => d.dataset['kind']);
    expect(after).toContain('bookmarkStart');
    expect(host.querySelector('details[data-kind="bookmarkStart"]')!.querySelectorAll('li').length).toBe(1);
    // Off again: everything.
    await act(async () => { filter.click(); });
    expect(host.querySelectorAll('details.diagnostic-kind').length).toBe(7);
    root.unmount();
    editor.destroy();
  });
});

describe('a folding section', () => {
  beforeEach(() => { localStorage.clear(); });

  test('opens by default, mounts as it was left, and saves a toggle', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    await act(async () => { root.render(<Fold id="test.section" title="Console" meta="types loaded" testId="fold-test"><p>body</p></Fold>); });
    const section = host.querySelector('[data-testid="fold-test"]')!;
    expect(section.tagName).toBe('SECTION');
    expect(section.getAttribute('aria-label')).toBe('Console');
    const details = section.querySelector('details')!;
    expect(details.open).toBe(true);
    expect(details.querySelector('summary > h2')!.textContent).toBe('Console types loaded');
    // Closed by the user: the browser toggles the element and reports it; the state is saved.
    await act(async () => { details.open = false; details.dispatchEvent(new Event('toggle')); });
    expect(loadFold('test.section')).toBe(false);
    root.unmount();
    // Mounted again: closed, as it was left; a section without a saved state takes its default.
    const again = createRoot(host);
    await act(async () => { again.render(<><Fold id="test.section" title="Console"><p>body</p></Fold><Fold id="test.other" level={3} title="Binding" defaultOpen={false}><p>b</p></Fold></>); });
    const both = host.querySelectorAll('details');
    expect(both[0]!.open).toBe(false);
    expect(both[1]!.open).toBe(false);
    expect(both[1]!.closest('.fold')!.tagName).toBe('DIV');   // a group is not a region
    saveFold('test.other', true);
    expect(loadFold('test.other', false)).toBe(true);
    again.unmount();
  });
});
