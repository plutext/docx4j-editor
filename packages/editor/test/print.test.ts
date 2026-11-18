// SPDX-License-Identifier: GPL-3.0-or-later
//
// Printing (ED-006 level 1): the page rules a document's sections make (section 3.3).
import { describe, test, expect } from 'vitest';
import { schema } from '@docx4j/editor-model';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { printPageRules, printSections } from '../src/view/print.js';

const A4 = { pgSz: { w: 11906, h: 16838 }, pgMar: { top: 1440, right: 1440, bottom: 1440, left: 1440, gutter: 0 } };
const A4_LANDSCAPE = { pgSz: { w: 16838, h: 11906 }, pgMar: A4.pgMar };
const LETTER_GUTTER = { pgSz: { w: 12240, h: 15840 }, pgMar: { top: 1080, right: 1080, bottom: 1080, left: 1080, gutter: 720 } };
const sect = (page: object, type?: string): wml.SectPr => ({ ...page, ...(type ? { type: { val: type } } : {}) }) as unknown as wml.SectPr;
const lines = (css: string): string[] => css.split('\n').map((l) => l.trim()).filter((l) => l !== '' && l !== '}' && !l.startsWith('@media'));

describe('printPageRules', () => {
  test('one section: one page, every block on it', () => {
    expect(lines(printPageRules([sect(A4)]))).toEqual([
      '@page docx-s0 { size: 595.3pt 841.9pt; margin: 72pt 18pt 72pt 18pt; }',
      '.docx-editor.body > .ProseMirror, .print-notes { padding-inline: 54pt; }',
      '.docx-editor.body > .ProseMirror > * { page: docx-s0; }',
      '.print-notes { page: docx-s0; }',
    ]);
  });

  test('portrait, landscape, the same landscape after a continuous break, Letter with a gutter', () => {
    expect(lines(printPageRules([sect(A4), sect(A4_LANDSCAPE, 'nextPage'), sect(A4_LANDSCAPE, 'continuous'), sect(LETTER_GUTTER, 'nextPage')]))).toEqual([
      '@page docx-s0 { size: 595.3pt 841.9pt; margin: 72pt 36pt 72pt 36pt; }',
      '@page docx-s1 { size: 841.9pt 595.3pt; margin: 72pt 36pt 72pt 36pt; }',
      '@page docx-s3 { size: 612pt 792pt; margin: 54pt 18pt 54pt 54pt; }',
      '.docx-editor.body > .ProseMirror, .print-notes { padding-inline: 36pt; }',
      '.docx-editor.body > .ProseMirror > * { page: docx-s0; }',
      '.docx-editor.body > .ProseMirror > .section-break[data-section="0"] ~ * { page: docx-s1; }',
      // The continuous section reuses the page, so no break is forced between them.
      '.docx-editor.body > .ProseMirror > .section-break[data-section="1"] ~ * { page: docx-s1; }',
      '.docx-editor.body > .ProseMirror > .section-break[data-section="2"] ~ * { page: docx-s3; }',
      '.print-notes { page: docx-s3; }',
    ]);
  });

  test('a next-page break between two sections of the same page breaks after the first one; odd and even ask for a side', () => {
    const css = lines(printPageRules([sect(A4), sect(A4, 'nextPage'), sect(A4, 'oddPage'), sect(A4_LANDSCAPE, 'evenPage'), sect(A4_LANDSCAPE)]));
    expect(css.filter((l) => l.startsWith('@page'))).toHaveLength(2);
    expect(css.filter((l) => l.includes('break-after'))).toEqual([
      '.docx-editor.body > .ProseMirror > :has(+ .section-break[data-section="0"]) { break-after: page; }',
      '.docx-editor.body > .ProseMirror > :has(+ .section-break[data-section="1"]) { break-after: right; }',
      '.docx-editor.body > .ProseMirror > :has(+ .section-break[data-section="2"]) { break-after: left; }',
      // No w:type: a next-page start, Word's default, on the same page as the one before.
      '.docx-editor.body > .ProseMirror > :has(+ .section-break[data-section="3"]) { break-after: page; }',
    ]);
  });

  test('mirror margins swap the inside and outside on a left-hand page; a gutter at the top goes on the top', () => {
    expect(lines(printPageRules([sect(LETTER_GUTTER)], { mirrorMargins: true, gutterAtTop: false })).slice(0, 2)).toEqual([
      '@page docx-s0 { size: 612pt 792pt; margin: 54pt 18pt 54pt 54pt; }',
      '@page docx-s0:left { margin-left: 18pt; margin-right: 54pt; }',
    ]);
    expect(lines(printPageRules([sect(LETTER_GUTTER)], { mirrorMargins: false, gutterAtTop: true }))[0]).toBe('@page docx-s0 { size: 612pt 792pt; margin: 90pt 18pt 54pt 18pt; }');
  });

  test("Word's defaults for what a section does not say; a negative top or bottom margin taken as its size", () => {
    expect(lines(printPageRules([null]))[0]).toBe('@page docx-s0 { size: 595.3pt 841.9pt; margin: 72pt 18pt 72pt 18pt; }');
    expect(lines(printPageRules([sect({ pgSz: A4.pgSz, pgMar: { ...A4.pgMar, top: -1440, bottom: -720 } })]))[0]).toBe('@page docx-s0 { size: 595.3pt 841.9pt; margin: 72pt 18pt 36pt 18pt; }');
  });

  describe('the margins give way to padding: Chromium cuts what stands outside them, and the text stays put', () => {
    /** The left and right a page's rule and the padding add up to, in points. */
    const stands = (css: string, index = 0): { left: number; right: number } => {
      const m = /margin: \S+ (\S+)pt \S+ (\S+)pt;/.exec(lines(css).filter((l) => l.startsWith('@page') && !l.includes(':left'))[index]!)!;
      const padding = Number(/padding-inline: (\S+)pt;/.exec(css)![1]);
      return { left: Number(m[2]) + padding, right: Number(m[1]) + padding };
    };

    test('every section of a document keeps its own margins, the padding being one for all', () => {
      const css = printPageRules([sect(A4), sect(LETTER_GUTTER, 'nextPage')]);
      expect(stands(css, 0)).toEqual({ left: 72, right: 72 });
      expect(stands(css, 1)).toEqual({ left: 90, right: 54 });
    });

    test('a margin that is no round number is not left with a tail', () => {
      const css = printPageRules([sect({ pgSz: A4.pgSz, pgMar: { ...A4.pgMar, left: 1417, right: 1417 } })]);
      expect(lines(css)[0]).toBe('@page docx-s0 { size: 595.3pt 841.9pt; margin: 72pt 18pt 72pt 18pt; }');
      expect(lines(css)[1]).toBe('.docx-editor.body > .ProseMirror, .print-notes { padding-inline: 52.85pt; }');
      // Unequal sides, where what is left of a margin is itself no round number.
      const uneven = printPageRules([sect({ pgSz: A4.pgSz, pgMar: { ...A4.pgMar, left: 1440, right: 1417 } })]);
      expect(lines(uneven)[0]).toBe('@page docx-s0 { size: 595.3pt 841.9pt; margin: 72pt 18pt 72pt 19.15pt; }');
    });

    test('a left-hand page swaps its margins and stays clear of the edge on both sides', () => {
      const [right, left] = lines(printPageRules([sect(LETTER_GUTTER)], { mirrorMargins: true, gutterAtTop: false }));
      expect(right).toContain('margin: 54pt 18pt 54pt 54pt');   // 54 + 36 on the outside, 90 on the inside
      expect(left).toBe('@page docx-s0:left { margin-left: 18pt; margin-right: 54pt; }');
    });

    test("a margin already at the paper's edge gives nothing, and none goes negative", () => {
      const narrow = printPageRules([sect({ pgSz: A4.pgSz, pgMar: { ...A4.pgMar, left: 200, right: 1440 } })]);
      expect(lines(narrow)[0]).toBe('@page docx-s0 { size: 595.3pt 841.9pt; margin: 72pt 72pt 72pt 10pt; }');
      expect(lines(narrow)[1]).toContain('padding-inline: 0pt');
    });
  });
});

test("printSections: the paragraph-level sections then the final one, from the view's own document", () => {
  const para = (text: string, sectPr: object | null = null) => schema.nodes['paragraph']!.create({ sectPr }, schema.text(text));
  const doc = schema.nodes['doc']!.create({ finalSectPr: LETTER_GUTTER }, [para('a'), para('b', A4_LANDSCAPE), para('c')]);
  expect(printSections(doc).map((s) => s?.pgSz?.w)).toEqual([16838, 12240]);
});

describe('breaks within the flow (section 3.4)', () => {
  test('a paragraph states its own breaks and keeps as 1 or 0; a style turns on what it has', async () => {
    const { paragraphPagination, paginationDeclarations, rowPagination } = await import('../src/view/print.js');
    expect(paragraphPagination({ pageBreakBefore: {}, keepNext: { val: false }, keepLines: { val: true } } as never)).toEqual({ pageBreakBefore: '1', keepNext: '0', keepLines: '1' });
    expect(paragraphPagination(undefined)).toEqual({});
    expect(paginationDeclarations({ keepNext: {}, keepLines: {}, pageBreakBefore: { val: false }, widowControl: { val: false } } as never)).toEqual(['break-after: avoid', 'break-inside: avoid', 'orphans: 1', 'widows: 1']);
    // Widow control on, or unsaid, is the browsers' default already.
    expect(paginationDeclarations({ widowControl: {} } as never)).toEqual([]);
    const item = (local: string, value: object = { TYPE_NAME: 'org_docx4j_wml.BooleanDefaultTrue' }) => ({ name: { localPart: local }, value });
    expect(rowPagination({ cnfStyleOrDivIdOrGridBefore: [item('tblHeader'), item('cantSplit', { val: false })] })).toEqual({ cantSplit: false, header: true });
    expect(rowPagination(null)).toEqual({ cantSplit: false, header: false });
  });
});

describe('over print-pages.docx', () => {
  const open = async () => {
    const { readFile } = await import('node:fs/promises');
    const { fileURLToPath } = await import('node:url');
    const { dirname, join } = await import('node:path');
    const { EditorDocument } = await import('@docx4j/editor-model');
    const { DocxEditorView } = await import('../src/view/DocxEditorView.js');
    const fixture = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures', 'print-pages.docx');
    const session = await EditorDocument.open(new Uint8Array(await readFile(fixture)));
    const container = document.createElement('div');
    document.body.appendChild(container);
    return { session, editor: new DocxEditorView(container, session) };
  };

  test("the stylesheet prints a heading style's keep with next; the paragraphs and rows carry their own", async () => {
    const { session, editor } = await open();
    const { documentCss } = await import('../src/view/css.js');
    const css = documentCss(session, editor.state.doc);
    expect(css).toMatch(/@media print \{[^]*\.docx-editor \.s-Heading1 \{[^}]*break-after: avoid/);
    const p = (text: string): HTMLElement => [...editor.view.dom.querySelectorAll('p')].find((e) => e.textContent?.startsWith(text)) as HTMLElement;
    expect(p('This paragraph starts a page').dataset['pageBreakBefore']).toBe('1');
    expect(p('This paragraph keeps with the next').dataset['keepNext']).toBe('1');
    expect(p('A heading that does not keep').dataset['keepNext']).toBe('0');
    expect(p('A paragraph with widow control off').dataset['widowControl']).toBe('0');
    expect(p('The next paragraph').dataset['keepNext']).toBeUndefined();
    const rows = [...editor.view.dom.querySelectorAll('tr')] as HTMLElement[];
    expect(rows.filter((r) => r.dataset['header'] === '1').map((r) => r.textContent?.slice(0, 17))).toEqual(['Header oneHeader ', 'Second header row']);
    expect(rows.filter((r) => r.dataset['cantSplit'] === '1')).toHaveLength(1);
    editor.destroy();
  });

  test('at beforeprint the table gets a copy of its two header rows, outside the content; at afterprint it goes; the document is untouched', async () => {
    const { editor } = await open();
    const before = editor.state.doc;
    window.dispatchEvent(new Event('beforeprint'));
    const thead = editor.view.dom.querySelector('table > thead.print-header') as HTMLElement;
    expect(thead.dataset['rows']).toBe('2');
    expect(thead.querySelectorAll('tr')).toHaveLength(2);
    // ProseMirror reads nothing from the copy: give its observer a turn.
    await new Promise((r) => setTimeout(r, 20));
    expect(editor.state.doc.eq(before)).toBe(true);
    window.dispatchEvent(new Event('afterprint'));
    expect(editor.view.dom.querySelector('thead.print-header')).toBeNull();
    await new Promise((r) => setTimeout(r, 20));
    expect(editor.state.doc.eq(before)).toBe(true);
    editor.destroy();
  });
});
