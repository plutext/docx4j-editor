// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 4.4: after an Apply to a part the look depends on, the body view's
// `partsChanged` walks the labels again and redraws the decorations, the document unchanged; a
// part edit arriving from a peer does the same through `onPartsReceived`.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { EditorDocument, applyPartXml } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { labelWalksOf } from '../src/view/plugins/labels.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

describe('partsChanged', () => {
  test('a numbering Apply: the labels walked again and drawn in the new format', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'tracked-changes.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session, { headersAndFooters: false });
    const labels = (): string[] => Array.from(container.querySelectorAll('.list-label')).map((el) => el.textContent ?? '');
    expect(labels()).toContain('1)');
    const walks = labelWalksOf(editor.state);
    const part = session.main.numberingDefinitionsPart!;
    await applyPartXml(session, part, (await part.getXml()).replace(/w:numFmt w:val="decimal"/g, 'w:numFmt w:val="upperRoman"'));
    // Nothing redraws until the view is told.
    expect(labels()).toContain('1)');
    editor.partsChanged();
    expect(labelWalksOf(editor.state)).toBe(walks + 1);
    expect(labels()).toContain('I)');
    expect(labels()).not.toContain('1)');
    editor.destroy();
  });
});
