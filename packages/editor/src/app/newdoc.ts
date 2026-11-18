// SPDX-License-Identifier: GPL-3.0-or-later
//
// New documents and templates (ED-003 section 8, "New documents and templates"): a new document
// from the engine's `createPackage` (docx4j's default styles, one section in the chosen page size
// with 2.54 cm margins), named `Document1.docx` onwards as Word does, with no file behind it until
// Save As or Download; the page size asked once and kept as an option. A Word template opened
// (`.dotx`, `.dotm`, told by the main part's content type) is offered for editing as it is, or as
// the basis of a new document: the same parts with the main part's content type changed to the
// document's, `.docm` for a macro-enabled template, which is what Word does with File > New. A
// Markdown file opened is a new document made from it (ED-005 section 12.66): `name.docx`, with no
// file behind it, so a save never writes over the Markdown. The margins of a new document, New's and a
// Markdown file's alike, are an option too (Word's Normal, or Narrow; ED-007 section 8.1 item 15: set
// before the open, so a Markdown table is fit to the text width the document will keep), never a
// template's, whose layout is its own.
import { WordprocessingMLPackage, ContentTypes, type PageSizePaper } from '@docx4j/core-ts';
import { markdownDocument, type EditorDocument } from '@docx4j/editor-model';
import type { OpenedFile } from './files.js';

export type { PageSizePaper };

/** The sizes offered, A4 and Letter first (Jason, 2026-09-25); each labelled by `pageSize.<id>` in the message files. */
export const PAGE_SIZES: readonly { id: PageSizePaper }[] = [
  { id: 'A4' }, { id: 'LETTER' }, { id: 'A3' }, { id: 'A5' }, { id: 'LEGAL' }, { id: 'B4JIS' },
];

const PAGE_SIZE_KEY = 'docx4j-editor.pageSize';

export function isPageSize(v: unknown): v is PageSizePaper {
  return PAGE_SIZES.some((s) => s.id === v);
}

/** The page size for new documents, once chosen; undefined until then (the first New asks). */
export function loadPageSize(): PageSizePaper | undefined {
  try {
    const v = localStorage.getItem(PAGE_SIZE_KEY);
    return isPageSize(v) ? v : undefined;
  } catch { return undefined; }
}

export function savePageSize(size: PageSizePaper | undefined): void {
  try {
    if (size === undefined) localStorage.removeItem(PAGE_SIZE_KEY); else localStorage.setItem(PAGE_SIZE_KEY, size);
  } catch { /* ignore */ }
}

/** Word's margin presets in twips a side: Normal 2.54 cm (1"), Narrow 1.27 cm; the header and footer distances are left alone. */
export const MARGIN_PRESETS = { normal: 1440, narrow: 720 } as const;
export type MarginPreset = keyof typeof MARGIN_PRESETS;

const MARGINS_KEY = 'docx4j-editor.margins';

export function isMarginPreset(v: unknown): v is MarginPreset {
  return typeof v === 'string' && v in MARGIN_PRESETS;
}

/** The margins for new documents: Normal until Narrow is chosen in Options. */
export function loadMargins(): MarginPreset {
  try {
    const v = localStorage.getItem(MARGINS_KEY);
    return isMarginPreset(v) ? v : 'normal';
  } catch { return 'normal'; }
}

export function saveMargins(preset: MarginPreset): void {
  try {
    if (preset === 'normal') localStorage.removeItem(MARGINS_KEY); else localStorage.setItem(MARGINS_KEY, preset);
  } catch { /* ignore */ }
}

/** The package's one section given the preset's margins on its four sides, where the engine's `createPackage` put Normal's. */
async function withMargins(pkg: WordprocessingMLPackage, preset: MarginPreset): Promise<void> {
  if (preset === 'normal') return;
  const body = (await pkg.getMainDocumentPart().getContents()).body;
  const pgMar = body?.sectPr?.pgMar;
  if (!pgMar) return;
  pgMar.top = pgMar.right = pgMar.bottom = pgMar.left = MARGIN_PRESETS[preset];
}

let counter = 0;

/** `Document1.docx`, `Document2.docx`, ... for the life of the page, as Word numbers them. */
export function nextDocumentName(extension: 'docx' | 'docm'): string {
  counter += 1;
  return `Document${counter}.${extension}`;
}

/** A new document: the engine's `createPackage` in the page size and the kept margins, saved to bytes so it opens as any file does. */
export async function newDocument(pageSize: PageSizePaper, margins: MarginPreset = 'normal'): Promise<OpenedFile> {
  const pkg = await WordprocessingMLPackage.createPackage({ pageSize });
  await withMargins(pkg, margins);
  return { name: nextDocumentName('docx'), bytes: await pkg.save(), lastModified: Date.now() };
}

export type TemplateKind = 'dotx' | 'dotm';

/** Whether the open document is a Word template, from its main part's content type. */
export function templateKind(session: EditorDocument): TemplateKind | undefined {
  const type = session.pkg.getMainDocumentPart().contentType;
  if (type === ContentTypes.WORDPROCESSINGML_TEMPLATE) return 'dotx';
  if (type === ContentTypes.WORDPROCESSINGML_TEMPLATE_MACROENABLED) return 'dotm';
  return undefined;
}

/**
 * A new document based on the open template: every part as it is, the main part's content type
 * the document's (macro-enabled for a `.dotm`). The session's package is changed in place and
 * saved; the caller opens the bytes as a new session and lets this one go.
 */
export async function documentFromTemplate(session: EditorDocument): Promise<OpenedFile> {
  const kind = templateKind(session);
  if (!kind) throw new Error('the open document is not a template');
  const main = session.pkg.getMainDocumentPart();
  const type = kind === 'dotm' ? ContentTypes.WORDPROCESSINGML_DOCUMENT_MACROENABLED : ContentTypes.WORDPROCESSINGML_DOCUMENT;
  main.contentType = type;
  session.pkg.contentTypeManager.addOverrideContentType(main.partName, type);
  return { name: nextDocumentName(kind === 'dotm' ? 'docm' : 'docx'), bytes: await session.pkg.save(), lastModified: Date.now() };
}

/** A Markdown file as a new document, in the kept page size (the engine's default before one is chosen) and margins, and what its reading dropped. A handle it came with becomes the document's source, never its own (ED-011 section 3.8). */
export async function documentFromMarkdown(file: OpenedFile, pageSize: PageSizePaper | undefined, margins: MarginPreset = 'normal'): Promise<{ file: OpenedFile; dropped: Record<string, number> }> {
  const text = new TextDecoder('utf-8').decode(file.bytes).replace(/^\uFEFF/, '');
  const { bytes, dropped } = await markdownDocument(text, { ...(pageSize ? { pageSize } : {}), margins: MARGIN_PRESETS[margins] });
  const source = file.handle ? { source: { handle: file.handle, name: file.name, lastModified: file.lastModified } } : {};
  return { file: { name: file.name.replace(/\.(md|markdown)$/i, '') + '.docx', bytes, lastModified: Date.now(), ...source }, dropped };
}
