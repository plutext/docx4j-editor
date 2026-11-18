// SPDX-License-Identifier: GPL-3.0-or-later
// @vitest-environment node
// (A package saved under jsdom cannot be reloaded there, ED-003 section 11.2 item 18; these helpers touch no DOM.)
//
// ED-003 section 8's "New documents and templates" (section 11.3): a new document is the
// engine's package in the chosen page size with the default styles and opens as any file does;
// a template is told by its main part's content type, and a document made from it keeps every
// part with the main part's content type changed to the document's (`.docm` for a `.dotm`).
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { ContentTypes, ZipPartStore, WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, documentStyles } from '@docx4j/editor-model';
import { newDocument, templateKind, documentFromTemplate, documentFromMarkdown, nextDocumentName, loadMargins, MARGIN_PRESETS } from '../src/app/newdoc.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

/** A fixture turned into a template of the given kind, as Word's Save As would (the main part's content type). */
async function asTemplate(name: string, kind: 'dotx' | 'dotm'): Promise<Uint8Array> {
  const pkg = await WordprocessingMLPackage.load(new Uint8Array(await readFile(join(fixtures, name))));
  const main = pkg.getMainDocumentPart();
  const type = kind === 'dotm' ? ContentTypes.WORDPROCESSINGML_TEMPLATE_MACROENABLED : ContentTypes.WORDPROCESSINGML_TEMPLATE;
  main.contentType = type;
  pkg.contentTypeManager.addOverrideContentType(main.partName, type);
  return pkg.save();
}

describe('new documents', () => {
  test('a new document has the chosen page size, the default styles and an empty paragraph to type in; it is named as Word names them', async () => {
    const file = await newDocument('LETTER');
    expect(file.name).toMatch(/^Document\d+\.docx$/);
    const session = await EditorDocument.open(file.bytes);
    expect(session.doc.childCount).toBeGreaterThanOrEqual(1);
    expect(session.doc.child(0).type.name).toBe('paragraph');
    const sectPr = session.doc.attrs['finalSectPr'] as { pgSz?: { w?: number; h?: number } } | null;
    expect(sectPr?.pgSz).toMatchObject({ w: 12240, h: 15840 });
    const styles = await documentStyles(session);
    expect(styles.some((s) => s.id === 'Heading1')).toBe(true);
    const a4 = await EditorDocument.open((await newDocument('A4')).bytes);
    expect((a4.doc.attrs['finalSectPr'] as { pgSz?: { w?: number } }).pgSz?.w).toBe(11907);
    expect(templateKind(session)).toBeUndefined();
    const n1 = nextDocumentName('docx');
    const n2 = nextDocumentName('docm');
    expect(Number(n2.match(/\d+/)![0])).toBe(Number(n1.match(/\d+/)![0]) + 1);
    expect(n2.endsWith('.docm')).toBe(true);
  });

  test('the margins option: Narrow gives New and a Markdown open Word\'s 1.27 cm on the four sides, Normal the engine\'s 2.54 cm; without storage the option is Normal (ED-007 section 8.1 item 15)', async () => {
    expect(MARGIN_PRESETS).toEqual({ normal: 1440, narrow: 720 });
    expect(loadMargins()).toBe('normal');   // no localStorage here
    const margins = async (bytes: Uint8Array): Promise<Record<string, number> | undefined> => ((await EditorDocument.open(bytes)).doc.attrs['finalSectPr'] as { pgMar?: Record<string, number> }).pgMar;
    expect(await margins((await newDocument('LETTER', 'narrow')).bytes)).toMatchObject({ top: 720, right: 720, bottom: 720, left: 720, header: 708, footer: 708, gutter: 0 });
    expect(await margins((await newDocument('LETTER', 'normal')).bytes)).toMatchObject({ top: 1440, right: 1440, bottom: 1440, left: 1440 });
    expect(await margins((await newDocument('LETTER')).bytes)).toMatchObject({ top: 1440, left: 1440 });
    const md = { name: 'notes.md', bytes: new TextEncoder().encode('# Notes\n\n| a | b |\n|---|---|\n| 1 | 2 |\n'), lastModified: 0 };
    const narrow = await documentFromMarkdown(md, 'LETTER', 'narrow');
    expect(narrow.file.name).toBe('notes.docx');
    expect(await margins(narrow.file.bytes)).toMatchObject({ top: 720, left: 720 });
    // the table's grid fills the text width those margins leave, so the fit after the open works over it
    const grid = (await EditorDocument.open(narrow.file.bytes)).doc.child(1).attrs['tblGrid'] as number[];
    expect(grid.reduce((a, b) => a + b, 0)).toBe(12240 - 1440);
    expect(await margins((await documentFromMarkdown(md, 'LETTER')).file.bytes)).toMatchObject({ top: 1440, left: 1440 });
  });

  test('a template is told by its main part; a document made from it keeps every part and changes only the content type (.docm for a .dotm)', async () => {
    for (const kind of ['dotx', 'dotm'] as const) {
      const template = await asTemplate('HelloWordOnline.docx', kind);
      const session = await EditorDocument.open(template);
      expect(templateKind(session)).toBe(kind);
      const file = await documentFromTemplate(session);
      expect(file.name).toMatch(kind === 'dotm' ? /^Document\d+\.docm$/ : /^Document\d+\.docx$/);
      const made = await WordprocessingMLPackage.load(file.bytes);
      expect(made.getMainDocumentPart().contentType).toBe(kind === 'dotm' ? ContentTypes.WORDPROCESSINGML_DOCUMENT_MACROENABLED : ContentTypes.WORDPROCESSINGML_DOCUMENT);
      expect(templateKind(await EditorDocument.open(file.bytes))).toBeUndefined();
      // Every part but the content types is byte for byte the template's.
      const a = new ZipPartStore(template);
      const b = new ZipPartStore(file.bytes);
      expect([...b.partNames()].sort()).toEqual([...a.partNames()].sort());
      for (const name of a.partNames()) {
        if (name === '[Content_Types].xml' || name.endsWith('.rels')) continue;
        expect(Buffer.from(b.loadSync(name)).equals(Buffer.from(a.loadSync(name))), name).toBe(true);
      }
      expect(new TextDecoder().decode(b.loadSync('[Content_Types].xml'))).toContain(kind === 'dotm' ? 'document.macroEnabled.main+xml' : 'wordprocessingml.document.main+xml');
    }
  });
});
