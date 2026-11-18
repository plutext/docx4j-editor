// SPDX-License-Identifier: GPL-3.0-or-later
//
// What the style picker and the font boxes list (ED-002 section 5.4): the document's paragraph
// and character styles by name, quick styles first, and its fonts. Read privately, so listing
// costs no part its round trip.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { EditorDocument } from './session.mjs';

export interface StyleEntry {
  id: string;
  name: string;
  type: 'paragraph' | 'character';
  /** `w:qFormat`: a quick style, listed first as Word's gallery does. */
  quick: boolean;
  isDefault: boolean;
}

/** The document's styles, quick ones first, then by name; the default paragraph style first of all. */
export async function documentStyles(session: EditorDocument): Promise<StyleEntry[]> {
  const part = session.main.styleDefinitionsPart;
  if (!part) return [];
  let styles: wml.Styles;
  try {
    styles = await part.readContents();
  } catch {
    return [];
  }
  const out: StyleEntry[] = [];
  for (const s of styles.style ?? []) {
    if (!s.styleId) continue;
    const type = s.type === 'character' ? 'character' : s.type === 'paragraph' || s.type === undefined ? 'paragraph' : undefined;
    if (!type) continue;
    if (s.hidden !== undefined && s.hidden.val !== false) continue;
    out.push({ id: s.styleId, name: s.name?.val ?? s.styleId, type, quick: s.qFormat !== undefined && s.qFormat.val !== false, isDefault: s._default === true });
  }
  return out.sort((a, b) => Number(b.isDefault) - Number(a.isDefault) || Number(b.quick) - Number(a.quick) || a.name.localeCompare(b.name));
}

/** Word's common font list, after the document's own. */
export const COMMON_FONTS = ['Aptos', 'Arial', 'Calibri', 'Cambria', 'Consolas', 'Courier New', 'Garamond', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana'];

/**
 * The document's fonts, then the common list: the names the runs set directly, the fonts of the
 * styles in use and the default, resolved through the run-font selector where a style names a
 * theme font. Computed over the projection and the resolver, not the engine's story walk, which
 * unmarshals the header and footer parts and would cost them their round trip.
 */
export async function documentFonts(session: EditorDocument): Promise<string[]> {
  const inUse = new Set<string>();
  const resolver = session.pkg.propertyResolverOrUndefined;
  const selector = session.main.runFontSelectorOrUndefined;
  const styleIds = new Set<string | undefined>([resolver?.defaultParagraphStyleId]);
  session.doc.descendants((node) => {
    if (node.type.name === 'paragraph' && node.attrs['pStyle']) styleIds.add(node.attrs['pStyle'] as string);
    for (const m of node.marks) {
      if (m.type.name === 'font' && typeof m.attrs['name'] === 'string' && m.attrs['name']) inUse.add(m.attrs['name']);
      if (m.type.name === 'runStyle' && typeof m.attrs['id'] === 'string') styleIds.add(m.attrs['id']);
    }
    return true;
  });
  if (resolver) {
    for (const id of styleIds) {
      try {
        const rPr = id === undefined ? resolver.getEffectiveRPr(undefined, undefined) : resolver.getEffectiveRPr(id);
        const pPr = id === undefined ? undefined : resolver.getEffectivePPr(id);
        const name = selector ? selector.spans(pPr, rPr, 'a', { rPrIsEffective: true })[0]?.documentFont : rPr?.rFonts?.ascii;
        if (name) inUse.add(name);
      } catch { /* a style the resolver cannot answer for */ }
    }
  }
  const own = [...inUse].sort((a, b) => a.localeCompare(b));
  return [...own, ...COMMON_FONTS.filter((f) => !inUse.has(f))];
}
