// SPDX-License-Identifier: GPL-3.0-or-later
//
// Embedded fonts (ED-001 section 11 item 7; ED-003 section 8.4, step F3; section 9 item 34): the
// font table names, per `w:font`, the parts that embed its regular, bold, italic and bold italic
// faces, each with a `w:fontKey`. An `.odttf` part is obfuscated per ECMA-376 part 2 section
// 8.4 (docx4j's `ObfuscatedFontPart`): the key's GUID as sixteen bytes in reverse order, XORed
// over the first thirty-two bytes of the font, twice over. A `.ttf` part is the font as it is.
// Everything here reads privately: the font table through `readContents`, the parts' bytes as
// they are in the container, so a document with embedded fonts round-trips byte for byte. The
// application turns the faces into `FontFace`s for the session and nothing is ever written; a
// subset Word made needs one field zeroed on the copy before a browser will take it (`repairFont`).
import { ObfuscatedFontPart, TrueTypeFontPart, BinaryPart, type Part } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { EditorDocument } from './session.mjs';

export interface EmbeddedFace {
  /** The `w:font/@w:name`: the family the runs name. */
  family: string;
  weight: 'normal' | 'bold';
  style: 'normal' | 'italic';
  /** The font file's bytes, deobfuscated where the part was obfuscated. */
  bytes: Uint8Array;
  /** Whether the part was an `.odttf` (obfuscated) rather than a plain `.ttf`. */
  obfuscated: boolean;
  /** `w:subsetted` on the embed element, when Word said the font holds the characters used only. */
  subsetted: boolean;
  /** What `repairFont` changed in `bytes` for the browser's parser, one line each; empty for a font it accepts as it is. The part is never changed. */
  repairs: string[];
  partName: string;
}

/**
 * UPSTREAM-5: Word's subsetter (2010 and Microsoft 365 alike) writes a `cmap` format 12 subtable
 * whose `language` field is not zero (0x08A00000 on the fixture's upright faces, 0x0A050000 on
 * the italics, the same from either Word; the specification requires zero outside the Macintosh
 * platform), and the OpenType Sanitizer in Chromium and Firefox refuses the whole font for it
 * ("cmap: format 12 subtable language should be zero"), so a `FontFace` over Word's subset fails
 * to load. Zeroes the field on a copy, for the 32-bit-language formats (8, 10, 12, 13) on the
 * Unicode and Windows platforms, and says what it did; the whole-font case has zero already and
 * comes back as it was. `docs/developer/upstream-limitations.md` item 5.
 */
export function repairFont(bytes: Uint8Array): { bytes: Uint8Array; repairs: string[] } {
  const repairs: string[] = [];
  if (bytes.length < 12) return { bytes, repairs };
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let cmap = -1;
  const count = view.getUint16(4);
  for (let i = 0; i < count && 12 + 16 * i + 16 <= bytes.length; i++) {
    const at = 12 + 16 * i;
    if (String.fromCharCode(...bytes.slice(at, at + 4)) === 'cmap') { cmap = view.getUint32(at + 8); break; }
  }
  if (cmap < 0 || cmap + 4 > bytes.length) return { bytes, repairs };
  const patches: number[] = [];
  const subtables = view.getUint16(cmap + 2);
  for (let i = 0; i < subtables && cmap + 12 + 8 * i <= bytes.length; i++) {
    const platform = view.getUint16(cmap + 4 + 8 * i);
    const encoding = view.getUint16(cmap + 6 + 8 * i);
    const at = cmap + view.getUint32(cmap + 8 + 8 * i);
    if (platform === 1 || at + 12 > bytes.length) continue;
    const format = view.getUint16(at);
    if (format !== 8 && format !== 10 && format !== 12 && format !== 13) continue;
    const language = view.getUint32(at + 8);
    if (language === 0) continue;
    patches.push(at + 8);
    repairs.push(`cmap (${platform},${encoding}) format ${format} language field zeroed`);   // the value is fixed per source face (the same for the italic from Word 2010 and Microsoft 365, over different subsets): derived from the source font somehow, not zero
  }
  if (patches.length === 0) return { bytes, repairs };
  const out = new Uint8Array(bytes);
  const outView = new DataView(out.buffer, out.byteOffset, out.byteLength);
  for (const at of patches) outView.setUint32(at, 0);
  return { bytes: out, repairs };
}

/** The sixteen bytes of a `w:fontKey` GUID (`{F1239FE5-5DF2-409A-8D50-4E519C08D5B6}`), in the order the key applies them: reversed. */
export function fontKeyBytes(fontKey: string): Uint8Array {
  const hex = fontKey.replace(/[{}-]/g, '');
  if (!/^[0-9a-fA-F]{32}$/.test(hex)) throw new Error(`not a font key: ${fontKey}`);
  const out = new Uint8Array(16);
  for (let i = 0; i < 16; i++) out[15 - i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16);
  return out;
}

/** The obfuscation and its inverse (the same XOR): the key over bytes 0 to 15 and again over 16 to 31. */
export function deobfuscateFont(bytes: Uint8Array, fontKey: string): Uint8Array {
  const key = fontKeyBytes(fontKey);
  const out = new Uint8Array(bytes);
  for (let i = 0; i < 32 && i < out.length; i++) out[i] = out[i]! ^ key[i % 16]!;
  return out;
}

const FACES: { element: 'embedRegular' | 'embedBold' | 'embedItalic' | 'embedBoldItalic'; weight: EmbeddedFace['weight']; style: EmbeddedFace['style'] }[] = [
  { element: 'embedRegular', weight: 'normal', style: 'normal' },
  { element: 'embedBold', weight: 'bold', style: 'normal' },
  { element: 'embedItalic', weight: 'normal', style: 'italic' },
  { element: 'embedBoldItalic', weight: 'bold', style: 'italic' },
];

/** The faces the document embeds, read privately; none when it has no font table or no embed elements. */
export async function embeddedFaces(session: EditorDocument): Promise<EmbeddedFace[]> {
  const table = session.main.fontTablePart;
  if (!table) return [];
  const fonts: wml.Fonts = await table.readContents();
  const rels = table.getRelationshipsPart();
  const out: EmbeddedFace[] = [];
  for (const font of fonts.font ?? []) {
    const family = font.name;
    if (!family) continue;
    for (const face of FACES) {
      const rel = font[face.element] as (wml.FontRel & { id?: string }) | undefined;
      if (!rel?.id) continue;
      const part: Part | undefined = rels?.getPart(rel.id);
      if (!(part instanceof BinaryPart)) continue;
      const raw = await part.getBytes();
      const obfuscated = part instanceof ObfuscatedFontPart || (!(part instanceof TrueTypeFontPart) && /\.odttf$/i.test(part.partName.toString()));
      const { bytes, repairs } = repairFont(obfuscated && rel.fontKey ? deobfuscateFont(raw, rel.fontKey) : raw);   // UPSTREAM-5
      out.push({ family, weight: face.weight, style: face.style, bytes, obfuscated, subsetted: rel.subsetted === true, repairs, partName: part.partName.toString() });
    }
  }
  return out;
}

/** True when the bytes start as a TrueType or OpenType font does. */
export function looksLikeFont(bytes: Uint8Array): boolean {
  const tag = String.fromCharCode(...bytes.slice(0, 4));
  return tag === '\0\0\0' || tag === 'OTTO' || tag === 'true' || tag === 'ttcf';
}
