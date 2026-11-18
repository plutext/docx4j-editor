// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 5.6 step C2: the generated stylesheet.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { EditorDocument, listLabels } from '@docx4j/editor-model';
import { documentCss, fontStack, paragraphDeclarations, runDeclarations, pageGeometry } from '../src/view/css.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');
const open = async (name: string) => EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));

describe('colours as Word draws them (ED-008 section 8.8 item 13; core-ts 0.6.0, CR-008 phase 2)', () => {
  test('a style\'s theme colour is drawn from the theme, not the cached w:val beside it: check 53\'s Heading 1', async () => {
    const session = await open('theme-colour-cache.docx');
    const css = documentCss(session);
    const heading = css.split('\n').find((r) => r.startsWith('.docx-editor .s-Heading1 '));
    expect(heading).toBeDefined();
    expect(heading).toContain('color: #0F4761');
    expect(heading).not.toContain('365F91');
    // Without a resolver the written value is drawn, as before.
    expect(runDeclarations({ color: { val: '365F91', themeColor: 'accent1', themeShade: 'BF' } } as never)).toContain('color: #365F91');
    expect(runDeclarations({ color: { val: '365F91', themeColor: 'accent1', themeShade: 'BF' } } as never, undefined, session.pkg.propertyResolver)).toContain('color: #0F4761');
  });
});

describe('font stacks', () => {
  test('a document font gets its metric-compatible substitutes and its generic', () => {
    expect(fontStack('Calibri')).toBe('"Calibri", "Carlito", "Liberation Sans", sans-serif');
    expect(fontStack('Times New Roman')).toMatch(/^"Times New Roman", "Tinos", "Liberation Serif", serif$/);
    expect(fontStack('Courier New')).toMatch(/monospace$/);
    expect(fontStack('Some Unknown Font')).toBe('"Some Unknown Font", sans-serif');
  });
});

describe('declarations', () => {
  test('paragraph properties use logical properties only', () => {
    const decl = paragraphDeclarations({ jc: { val: 'right' }, ind: { left: 720, hanging: 360 }, spacing: { before: 120, after: 240, line: 276, lineRule: 'auto' }, bidi: {} });
    expect(decl).toContain('text-align: end');
    expect(decl).toContain('padding-inline-start: 36pt');   // w:left is where the lines sit; the hanging first line is text-indent's
    expect(decl).toContain('text-indent: -18pt');
    expect(decl).toContain('margin-block-start: 6pt');
    expect(decl).toContain('line-height: 1.150');
    expect(decl).toContain('direction: rtl');
    expect(decl.join(';')).not.toMatch(/margin-left|padding-left|text-align: left|text-align: right/);
  });
  test('run properties', () => {
    const decl = runDeclarations({ b: {}, i: {}, sz: { val: 28 }, u: { val: 'single' }, color: { val: 'FF0000' }, highlight: { val: 'yellow' } }, 'Arial');
    expect(decl).toContain('font-weight: bold');
    expect(decl).toContain('font-size: 14pt');
    expect(decl).toContain('text-decoration: underline');
    expect(decl).toContain('color: #FF0000');
    expect(decl).toContain('background-color: #FFFF00');
    expect(decl[0]).toMatch(/^font-family: "Arial", "Arimo"/);
  });
  test('page geometry has Word defaults', () => {
    expect(pageGeometry(undefined)).toEqual({ width: 595.3, height: 841.9, top: 72, right: 72, bottom: 72, left: 72 });
  });
});

describe('documentCss', () => {
  test('a rule per style in use, from the effective properties; the heading is bold when its style says so', async () => {
    const session = await open('tracked-changes.docx');
    const css = documentCss(session);
    expect(css).toMatch(/\.docx-editor \{ inline-size: [\d.]+pt; padding-inline: /);
    expect(css).toContain('.docx-editor .s-Heading1 {');
    expect(css).toContain('.docx-editor .s-Title {');
    const heading = /\.s-Heading1 \{ ([^}]*)\}/.exec(css)![1]!;
    const styleRPr = session.pkg.propertyResolver.getEffectiveRPr('Heading1');
    if (styleRPr?.b !== undefined && styleRPr.b.val !== false) expect(heading).toContain('font-weight: bold');
    expect(heading).toMatch(/font-family: "/);
    expect(css).not.toMatch(/margin-left|padding-left/);
  });
  test('list labels come from the emulator for the numbering fixture', async () => {
    const session = await open('numbering-stories.docx');
    const labels = listLabels(session);
    expect(labels.size).toBeGreaterThan(0);
    for (const l of labels.values()) expect(typeof l.listString).toBe('string');
  });
});

describe('the paragraph defaults', () => {
  test('a new document renders Word\'s 10pt after each paragraph and 1.15 lines: the defaults go on every paragraph, not on the editor', async () => {
    const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
    const session = await EditorDocument.open(await WordprocessingMLPackage.createPackage({ pageSize: 'A4' }));
    const css = documentCss(session);
    const p = /\.docx-editor p \{ ([^}]*)\}/.exec(css)![1]!;
    expect(p).toContain('margin-block: 0');
    expect(p).toContain('margin-block-end: 10pt');
    expect(p).toContain('line-height: 1.150');
    const editor = /\.docx-editor \{ ([^}]*)\}/g;
    let m: RegExpExecArray | null;
    while ((m = editor.exec(css)) !== null) expect(m[1]).not.toContain('margin-block');
  });
});
