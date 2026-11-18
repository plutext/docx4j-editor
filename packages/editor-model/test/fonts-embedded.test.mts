// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 8.4 (step F3): the faces a document embeds, deobfuscated per ECMA-376 part 2
// section 8.4, over `fonts-embedded.docx` (Word 2010, the entire Noto Sans in four faces) and
// `fonts-embedded-subset.docx` (Word 2010, the characters used only: section 11.18 item 4); the
// reader leaves the document byte for byte.
import { describe, test, expect } from 'vitest';
import { readFile, access } from 'node:fs/promises';
import { EditorDocument, embeddedFaces, deobfuscateFont, fontKeyBytes, looksLikeFont, repairFont } from '../src/index.mjs';
import { fixture, comparePackages } from './helpers.mjs';

const NOTO: Record<string, string> = {
  'normal/normal': '/usr/share/fonts/noto/NotoSans-Regular.ttf',
  'bold/normal': '/usr/share/fonts/noto/NotoSans-Bold.ttf',
  'normal/italic': '/usr/share/fonts/noto/NotoSans-Italic.ttf',
  'bold/italic': '/usr/share/fonts/noto/NotoSans-BoldItalic.ttf',
};

/** The TrueType table directory: tag to `[offset, length]`. */
function tables(font: Uint8Array): Map<string, [number, number]> {
  const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
  const out = new Map<string, [number, number]>();
  const n = view.getUint16(4);
  for (let i = 0; i < n; i++) {
    const at = 12 + 16 * i;
    out.set(String.fromCharCode(...font.slice(at, at + 4)), [view.getUint32(at + 8), view.getUint32(at + 12)]);
  }
  return out;
}

/** Whether the font's `cmap` maps the character to a glyph, read from its format 4 (BMP) subtables: what decides whether a browser draws the character with this face or falls back. */
function covers(font: Uint8Array, char: string): boolean {
  const code = char.codePointAt(0)!;
  const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
  const [cmap] = tables(font).get('cmap')!;
  const count = view.getUint16(cmap + 2);
  for (let i = 0; i < count; i++) {
    const sub = cmap + view.getUint32(cmap + 4 + 8 * i + 4);
    if (view.getUint16(sub) !== 4) continue;
    const segX2 = view.getUint16(sub + 6);
    const ends = sub + 14, starts = ends + segX2 + 2, deltas = starts + segX2, rangeOffsets = deltas + segX2;
    for (let s = 0; s < segX2 / 2; s++) {
      const start = view.getUint16(starts + 2 * s), end = view.getUint16(ends + 2 * s);
      if (code < start || code > end) continue;
      const rangeOffset = view.getUint16(rangeOffsets + 2 * s);
      let glyph: number;
      if (rangeOffset === 0) glyph = (code + view.getInt16(deltas + 2 * s)) & 0xffff;
      else {
        glyph = view.getUint16(rangeOffsets + 2 * s + rangeOffset + 2 * (code - start));
        if (glyph !== 0) glyph = (glyph + view.getInt16(deltas + 2 * s)) & 0xffff;
      }
      if (glyph !== 0) return true;
    }
  }
  return false;
}

describe('embedded fonts', () => {
  test('the key\'s GUID is applied in reverse, and the XOR is its own inverse', () => {
    const key = fontKeyBytes('{F1239FE5-5DF2-409A-8D50-4E519C08D5B6}');
    expect([...key.slice(0, 4)]).toEqual([0xb6, 0xd5, 0x08, 0x9c]);
    const font = new Uint8Array(40).map((_, i) => i);
    const once = deobfuscateFont(font, '{F1239FE5-5DF2-409A-8D50-4E519C08D5B6}');
    expect(once[0]).toBe(0xb6);
    expect(once[16]).toBe(16 ^ 0xb6);
    expect(once[32]).toBe(32);
    expect([...deobfuscateFont(once, '{F1239FE5-5DF2-409A-8D50-4E519C08D5B6}')]).toEqual([...font]);
    expect(() => fontKeyBytes('nonsense')).toThrow();
  });

  test('fonts-embedded.docx embeds Noto Sans in four faces, each a TrueType file once deobfuscated, equal to the installed font where it is', async () => {
    const bytes = await fixture('fonts-embedded.docx');
    const session = await EditorDocument.open(bytes);
    const faces = await embeddedFaces(session);
    expect(faces.map((f) => `${f.family} ${f.weight}/${f.style}`).sort()).toEqual(['Noto Sans bold/italic', 'Noto Sans bold/normal', 'Noto Sans normal/italic', 'Noto Sans normal/normal']);
    for (const f of faces) {
      expect(f.obfuscated).toBe(true);
      expect(f.partName).toMatch(/^\/word\/fonts\/font\d\.odttf$/);
      expect(looksLikeFont(f.bytes)).toBe(true);
      const installed = NOTO[`${f.weight}/${f.style}`]!;
      let present = true;
      try { await access(installed); } catch { present = false; }
      if (present) expect(Buffer.from(f.bytes).equals(await readFile(installed))).toBe(true);   // the whole file, when this machine has Noto Sans
      expect(f.subsetted).toBe(false);
      expect(f.repairs).toEqual([]);   // the whole font's cmap is as Noto Sans ships it
    }
    // Reading the faces touched nothing: the untouched round trip holds, font parts included.
    const saved = await session.save();
    expect(saved.changed).toBe(false);
    expect(comparePackages(bytes, saved.bytes)).toEqual([]);
  });

  // The same source saved by Word 2010 and by Microsoft 365 Word (the latter with its Aptos embeddings stripped by `build/fonts-embedded-strip.mjs`): the Noto Sans subsets, over 13 distinct characters in four faces.
  // UPSTREAM-5: Word writes a non-zero value in the format 12 subtables' language field (0x08A00000 on the upright faces, 0x0A050000 on the italics, from either Word), which the browsers' sanitizer refuses; the reader zeroes it on the copy, once per subtable, and says so. The Microsoft 365 machine's Noto Sans regular and bold were version 2.001, which has no format 12 subtable, so those two need nothing.
  const REPAIRED = ['cmap (0,4) format 12 language field zeroed', 'cmap (3,10) format 12 language field zeroed'];
  test.each([
    { name: 'fonts-embedded-subset.docx', repairs: { 'normal/normal': REPAIRED, 'bold/normal': REPAIRED, 'normal/italic': REPAIRED, 'bold/italic': REPAIRED } },
    { name: 'fonts-embedded-subset-m365.docx', repairs: { 'normal/normal': [], 'bold/normal': [], 'normal/italic': REPAIRED, 'bold/italic': REPAIRED } },
  ])('$name embeds the characters used only: subsetted, the outlines a fraction of the whole font, repaired where Word wrote the cmap language field, the letters used covered and a letter not used absent, the document untouched', async ({ name, repairs }) => {
    const bytes = await fixture(name);
    const session = await EditorDocument.open(bytes);
    const faces = await embeddedFaces(session);
    expect(faces.map((f) => `${f.family} ${f.weight}/${f.style}`).sort()).toEqual(['Noto Sans bold/italic', 'Noto Sans bold/normal', 'Noto Sans normal/italic', 'Noto Sans normal/normal']);
    const whole = await embeddedFaces(await EditorDocument.open(await fixture('fonts-embedded.docx')));
    for (const f of faces) {
      expect(f.obfuscated).toBe(true);
      expect(f.subsetted).toBe(true);
      expect(looksLikeFont(f.bytes)).toBe(true);
      expect(f.repairs, `${f.weight}/${f.style}`).toEqual(repairs[`${f.weight}/${f.style}` as keyof typeof repairs]);
      expect(repairFont(f.bytes).repairs).toEqual([]);
      // Word's subset keeps the glyph ids (maxp, loca, hmtx as they were) and empties the outlines it does not need: glyf is kilobytes against the whole font's hundreds of kilobytes.
      const [, glyf] = tables(f.bytes).get('glyf')!;
      const [, wholeGlyf] = tables(whole.find((w) => w.weight === f.weight && w.style === f.style)!.bytes).get('glyf')!;
      expect(glyf).toBeLessThan(wholeGlyf / 50);
      expect(f.bytes.length).toBeLessThan(200_000);
      // The source set "regular", "bold", "italic", "bold italic" and "abc" in the faces (13 distinct characters over the four, under Word's threshold of 32): `x` was never used, so no face maps it and a browser falls back for it.
      expect(covers(f.bytes, 'x')).toBe(false);
      expect(covers(whole.find((w) => w.weight === f.weight && w.style === f.style)!.bytes, 'x')).toBe(true);
    }
    const regular = faces.find((f) => f.weight === 'normal' && f.style === 'normal')!;
    for (const c of 'regular' + 'abc') expect(covers(regular.bytes, c), c).toBe(true);
    const bold = faces.find((f) => f.weight === 'bold' && f.style === 'normal')!;
    for (const c of 'bold') expect(covers(bold.bytes, c), c).toBe(true);
    expect(covers(bold.bytes, 'a')).toBe(false);   // used in the regular face only
    const saved = await session.save();
    expect(saved.changed).toBe(false);
    expect(comparePackages(bytes, saved.bytes)).toEqual([]);
  });

  test('a document without a font table or without embeds reports no faces', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    expect(await embeddedFaces(session)).toEqual([]);
  });
});
