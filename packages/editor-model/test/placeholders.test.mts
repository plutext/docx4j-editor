// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-008 section 8.8 item 14: the editor's own placeholders in Word's form. A control made at a caret, or a
// bound control written an empty value, shows Word's default placeholder for its kind, names the default entry
// in `w:placeholder/w:docPart` after `w:id` and before `w:showingPlcHdr`, and the save writes the entry into the
// glossary document where the file has none, through the engine (core-ts CR-002 section 63, D9), so a control
// the person makes and one the agent makes look alike. A kind Word gives no entry (a check box, a picture) names
// none. A file control keeps whatever form it has.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, runCommand, readControl, sdtItems, writeBoundValue, stateFor, schema, DEFAULT_PLACEHOLDERS, defaultPlaceholderFor, placeholderFor, PLACEHOLDER_TEXT, PLACEHOLDER_STYLE } from '../src/index.mjs';
import type { Node as PMNode } from 'prosemirror-model';
import { fixture, mainXml, comparePackages } from './helpers.mjs';

function controls(doc: PMNode): PMNode[] {
  const out: PMNode[] = [];
  doc.descendants((n) => { if (n.type.name === 'sdtInline' || n.type.name === 'sdtBlock') out.push(n); return true; });
  return out;
}

describe("the editor's own placeholders take Word's form (ED-008 section 8.8 item 14)", () => {
  test("a control made at a caret: Word's text for its kind, the default entry named in Word's place, and the entry written to the glossary at save", async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const at = { from: 3 };
    expect(runCommand(session, 'contentControl.insert', { kind: 'RichText' }, at).ok).toBe(true);
    expect(runCommand(session, 'contentControl.insert', { kind: 'DatePicker' }, { from: 40 }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.insert', { kind: 'DropDownList' }, { from: 80 }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.insert', { kind: 'CheckBox' }, { from: 120 }).ok).toBe(true);
    const made = controls(session.doc).filter((c) => c.attrs['sourceId'] === null);
    const byKind = Object.fromEntries(made.map((c) => [readControl(c.attrs['sdtPr']).kind, c]));
    expect(byKind['RichText']!.textContent).toBe(PLACEHOLDER_TEXT);
    expect(PLACEHOLDER_TEXT).toBe('Click or tap here to enter text.');
    expect(byKind['DatePicker']!.textContent).toBe('Click or tap to enter a date.');
    expect(byKind['DropDownList']!.textContent).toBe('Choose an item.');
    for (const [kind, entry] of [['RichText', DEFAULT_PLACEHOLDERS.text], ['DatePicker', DEFAULT_PLACEHOLDERS.date], ['DropDownList', DEFAULT_PLACEHOLDERS.list]] as const) {
      const reading = readControl(byKind[kind]!.attrs['sdtPr']);
      expect(reading.showingPlaceholder).toBe(true);
      expect(reading.placeholderEntry).toBe(entry.name);
      // Word's order: after w:id, before w:showingPlcHdr.
      const names = sdtItems(byKind[kind]!.attrs['sdtPr']).map((i) => i.name.localPart);
      expect(names.indexOf('placeholder')).toBeGreaterThan(names.indexOf('id'));
      expect(names.indexOf('placeholder')).toBeLessThan(names.indexOf('showingPlcHdr'));
      // The run is in the Placeholder Text style.
      byKind[kind]!.descendants((n) => { if (n.isText) expect(n.marks.some((m) => m.type.name === 'runStyle' && m.attrs['id'] === PLACEHOLDER_STYLE)).toBe(true); return true; });
    }
    // A check box has no default entry: the bare form, with no reference (the engine's rule).
    expect(readControl(byKind['CheckBox']!.attrs['sdtPr']).placeholderEntry).toBeUndefined();
    expect(defaultPlaceholderFor(byKind['CheckBox']!.attrs['sdtPr'])).toBeUndefined();

    // The save: the glossary document part made with the three entries, each one paragraph of its text in Placeholder Text.
    const before = await fixture('tracked-changes.docx');
    const { bytes, changed } = await session.save();
    expect(changed).toBe(true);
    const xml = await mainXml(bytes);
    expect(xml).toContain(`<w:placeholder><w:docPart w:val="${DEFAULT_PLACEHOLDERS.text.name}"/></w:placeholder><w:showingPlcHdr/>`);
    expect(xml).toContain(`<w:docPart w:val="${DEFAULT_PLACEHOLDERS.date.name}"/>`);
    expect(xml).toContain(`<w:docPart w:val="${DEFAULT_PLACEHOLDERS.list.name}"/>`);
    // The parts: the main part, the styles (Placeholder Text), app.xml; the glossary document, its rels and its styles added.
    const differences = comparePackages(before, bytes, ['word/document.xml', 'word/styles.xml', 'docProps/app.xml']);
    expect(differences.filter((d) => d.reason !== 'added')).toEqual([]);
    expect(differences.filter((d) => d.reason === 'added').map((d) => d.name).sort()).toEqual(['word/glossary/_rels/document.xml.rels', 'word/glossary/document.xml', 'word/glossary/styles.xml']);
    const saved = await WordprocessingMLPackage.load(bytes);
    const glossary = saved.getMainDocumentPart().glossaryDocumentPart;
    expect(glossary).toBeDefined();
    await saved.glossaryEntries.load(glossary);
    expect(saved.glossaryEntries.textOf(DEFAULT_PLACEHOLDERS.text.name)).toBe(PLACEHOLDER_TEXT);
    expect(saved.glossaryEntries.textOf(DEFAULT_PLACEHOLDERS.date.name)).toBe('Click or tap to enter a date.');
    expect(saved.glossaryEntries.textOf(DEFAULT_PLACEHOLDERS.list.name)).toBe('Choose an item.');
    // The engine reads the control as showing its placeholder, with no text (CR-012 section 18).
    const body = await saved.getBody();
    expect(body.text).not.toContain('Click or tap');

    // Saved again from the saved file after an edit: the glossary holds the entries, nothing is required twice.
    const again = await EditorDocument.open(bytes);
    expect(runCommand(again, 'insertText', { text: 'x' }, { from: 1 }).ok).toBe(true);
    const second = await again.save();
    expect(comparePackages(bytes, second.bytes, ['word/document.xml', 'docProps/app.xml'])).toEqual([]);
  });

  test("a bound control written an empty value shows its kind's default placeholder and names the entry; a value clears the reference's flag, not the name", async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: 3, to: 7 }).ok).toBe(true);
    const state = stateFor(session);
    const control = controls(state.doc).find((c) => c.attrs['sourceId'] === null)!;
    let pos = -1;
    state.doc.descendants((n, p) => { if (n === control) pos = p; return pos < 0; });
    const tr = state.tr;
    const sdtPr = writeBoundValue(tr, pos, control.attrs['sdtPr'], '');
    const reading = readControl(sdtPr);
    expect(reading.showingPlaceholder).toBe(true);
    expect(reading.placeholderEntry).toBe(DEFAULT_PLACEHOLDERS.text.name);
    expect(tr.doc.nodeAt(pos)!.textContent).toBe(PLACEHOLDER_TEXT);
    const tr2 = state.apply(tr).tr;
    const written = writeBoundValue(tr2, pos, sdtPr, 'Ann');
    expect(readControl(written).showingPlaceholder).toBe(false);
    expect(readControl(written).placeholderEntry).toBe(DEFAULT_PLACEHOLDERS.text.name);   // Word keeps the name; the flag says what shows
    expect(tr2.doc.nodeAt(pos)!.textContent).toBe('Ann');
  });

  test('placeholderFor: the properties and the text per kind, a kind without an entry the bare text', () => {
    const date = placeholderFor({ TYPE_NAME: 'org_docx4j_wml.SdtPr', rPrOrAliasOrLock: [{ name: { namespaceURI: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', localPart: 'date' }, value: { TYPE_NAME: 'org_docx4j_wml.CTSdtDate' } }] });
    expect(date.text).toBe('Click or tap to enter a date.');
    expect(readControl(date.sdtPr).placeholderEntry).toBe(DEFAULT_PLACEHOLDERS.date.name);
    const picture = placeholderFor({ TYPE_NAME: 'org_docx4j_wml.SdtPr', rPrOrAliasOrLock: [{ name: { namespaceURI: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', localPart: 'picture' }, value: { TYPE_NAME: 'org_docx4j_wml.CTEmpty' } }] });
    expect(picture.text).toBe(PLACEHOLDER_TEXT);
    expect(readControl(picture.sdtPr).placeholderEntry).toBeUndefined();
    expect(readControl(picture.sdtPr).showingPlaceholder).toBe(true);
    expect(schema.nodes['sdtInline']).toBeDefined();
  });
});
