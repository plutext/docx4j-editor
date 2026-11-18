// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 7.4: macros in the document as a custom XML part in the editor's namespace.
import { describe, test, expect } from 'vitest';
import { ZipPartStore } from '@docx4j/core-ts';
import { EditorDocument, MACROS_NAMESPACE, macrosXml, readDocumentMacros, writeDocumentMacros } from '../src/index.mjs';
import { fixture, comparePackages } from './helpers.mjs';

const ONE = { name: 'Number the headings', script: "for (const p of body.paragraphs) if (p.style === 'Heading 1') p.insertText('1. ', 'Start');" };
const TWO = { name: 'Bold <the> "selection"', script: 'selection.font.bold = true; // & done' };

describe('document macros', () => {
  test('a document without macros reads none, and stays byte for byte after the reader ran', async () => {
    const bytes = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(bytes);
    expect(await readDocumentMacros(session)).toEqual([]);
    const saved = await session.save();
    expect(saved.changed).toBe(false);
    expect(comparePackages(bytes, saved.bytes)).toEqual([]);
  });

  test('macros written are read back after a save and a reload; the part is in the namespace, marked, and Word-ignorable', async () => {
    const bytes = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(bytes);
    await writeDocumentMacros(session, [ONE, TWO]);
    expect(session.touchedParts.size).toBeGreaterThan(0);
    const saved = await session.save();
    expect(saved.changed).toBe(true);
    const names = [...new ZipPartStore(saved.bytes).partNames()];
    const partName = [...session.touchedParts].find((n) => n.includes('customXml'))!;
    expect(names).toContain(partName.replace(/^\//, ''));
    const xml = new TextDecoder().decode(new ZipPartStore(saved.bytes).loadSync(partName.replace(/^\//, '')));
    expect(xml).toContain(`xmlns="${MACROS_NAMESPACE}"`);
    expect(xml).toContain('Bold &lt;the&gt; &quot;selection&quot;');
    // Every other part is as it was, apart from the parts a custom XML part brings (its properties part, the regenerated relationships) and app.xml.
    const diffs = comparePackages(bytes, saved.bytes, ['docProps/app.xml']).filter((d) => d.reason !== 'added');
    expect(diffs).toEqual([]);
    const reloaded = await EditorDocument.open(saved.bytes);
    expect(await readDocumentMacros(reloaded)).toEqual([ONE, TWO]);
    // The reopened document, untouched, round-trips with the part in it.
    const again = await reloaded.save();
    expect(again.changed).toBe(false);
    expect(comparePackages(saved.bytes, again.bytes)).toEqual([]);
  });

  test('the list is rewritten in place, and the part goes with the last macro', async () => {
    const bytes = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(bytes);
    await writeDocumentMacros(session, [ONE]);
    await writeDocumentMacros(session, [ONE, TWO]);
    expect((await readDocumentMacros(session)).map((m) => m.name)).toEqual([ONE.name, TWO.name]);
    expect(session.pkg.customXmlParts.getByNamespace(MACROS_NAMESPACE)).toHaveLength(1);
    await writeDocumentMacros(session, []);
    expect(session.pkg.customXmlParts.getByNamespace(MACROS_NAMESPACE)).toHaveLength(0);
    const saved = await session.save();
    const names = [...new ZipPartStore(saved.bytes).partNames()];
    expect(names.some((n) => n.includes('customXml/item'))).toBe(false);
  });

  test('the XML escapes names and scripts', () => {
    const xml = macrosXml([TWO]);
    expect(xml).toContain('name="Bold &lt;the&gt; &quot;selection&quot;"');
    expect(xml).toContain('selection.font.bold = true; // &amp; done');
    expect(macrosXml([])).toContain(`<macros xmlns="${MACROS_NAMESPACE}"></macros>`);
  });
});
