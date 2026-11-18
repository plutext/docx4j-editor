// SPDX-License-Identifier: GPL-3.0-or-later
//
// A language change reaching what is already on screen (ED-003 sections 11.22 and 12.5): the
// toolbar re-renders through `useT`, the view's widgets and node views are made again through
// `relocalise()`, the document's own text is untouched, and a key the language lacks shows in
// English. French is the language used; the expectations are read from its file, so a review
// that rewords a message does not break the test.
import { afterEach, describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { EditorDocument, documentStyles, documentFonts } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import { Toolbar } from '../src/app/Toolbar.js';
import { setLanguage, language, loadMessages, format, t, isDraft, availableLanguages } from '../src/app/i18n.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

afterEach(async () => { await setLanguage('en'); });

describe('a language change on screen', () => {
  test('every shipped language loads and is marked as a draft until reviewed', async () => {
    const tags = availableLanguages().filter((tag) => tag !== 'en');
    expect(tags.length).toBe(19);
    for (const tag of tags) {
      const messages = await loadMessages(tag);
      expect(messages?.['toolbar.bold'], tag).toBeTruthy();
      expect(isDraft(tag), tag).toBe(true);
    }
  });

  test('the toolbar, the section break widget and the view\'s name follow; the document\'s text does not', async () => {
    const fr = (await loadMessages('fr'))!;
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'sections-multiple.docx'))));
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = new DocxEditorView(container, session);
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    const [styles, fonts] = await Promise.all([documentStyles(session), documentFonts(session)]);
    await act(async () => { root.render(<Toolbar editor={editor} styles={styles} fonts={fonts} />); });

    const bold = (): string => host.querySelector('button[data-command="font.bold"]')!.getAttribute('aria-label')!;
    const breakText = (): string => container.querySelector('.section-break')?.textContent ?? '';
    const pm = container.querySelector('.ProseMirror')!;
    const text = editor.state.doc.textContent;
    expect(bold()).toBe('Bold');
    expect(breakText()).toBe('Section Break (Next Page)');
    expect(pm.getAttribute('aria-label')).toBe('Document');
    expect(pm.getAttribute('dir')).toBe('ltr');

    await act(async () => { await setLanguage('fr'); });
    expect(language()).toBe('fr');
    expect(bold()).toBe(fr['toolbar.bold']);
    expect(breakText()).toBe(format(fr['view.sectionBreak']!, { type: fr['view.sectionBreak.nextPage'] }, 'fr'));
    expect(pm.getAttribute('aria-label')).toBe(fr['view.document']);
    expect(editor.state.doc.textContent).toBe(text);

    await act(async () => { await setLanguage('en'); });
    expect(bold()).toBe('Bold');
    expect(breakText()).toBe('Section Break (Next Page)');
    await act(async () => { root.unmount(); });
    editor.destroy();
  });

  test('a key the language lacks shows in English, formatted as English', async () => {
    const fr = (await loadMessages('fr'))!;
    await setLanguage('fr');
    const saved = fr['banner.summary.opaque'];
    delete fr['banner.summary.opaque'];
    try {
      expect(t('banner.summary.opaque', { count: 1 })).toBe('1 element shown read-only');
      expect(t('banner.summary.opaque', { count: 2 })).toBe('2 elements shown read-only');
    } finally {
      fr['banner.summary.opaque'] = saved!;
    }
    expect(t('banner.summary.opaque', { count: 2 })).toBe(format(saved!, { count: 2 }, 'fr'));
  });
});
