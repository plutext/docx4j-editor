// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 4.7 in the model (step B5): the preview is the bind step on a clone of the
// package holding the document as the edit layer would export it now: the bound values in
// place, the placeholder for an empty value (REQ-024), an unsaved edit shown, an edited data
// part reflected, the live document and package untouched.
import { describe, test, expect } from 'vitest';
import { JSDOM } from 'jsdom';
import { EditorDocument, previewInstance, listControls, runCommand, PLACEHOLDER_TEXT, type SdtGroup } from '../src/index.mjs';
import { fixture, comparePackages } from './helpers.mjs';

function blockStart(session: EditorDocument, i: number): number {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += session.doc.child(k).nodeSize;
  return pos;
}

describe('previewInstance', () => {
  test('invoice.docx: the values in place on a copy, the groups carried, the document and the package untouched; an edited data part and an unsaved edit show', async () => {
    const original = await fixture('invoice.docx');
    const session = await EditorDocument.open(original);
    await session.template.load({ engine: false });
    const preview = await previewInstance(session);
    expect(preview.result).toMatchObject({ bound: 3, updated: 3, skipped: 0 });
    expect(preview.doc.child(3).type.name).toBe('sdtBlock');
    expect(preview.doc.child(3).textContent).toBe('Joe Bloggs');
    const row = preview.doc.child(5).child(1);
    expect(row.child(0).textContent).toBe('apples');
    expect(row.child(1).textContent).toBe('$20');
    // The preview's groups come from its own projection, with the same properties.
    const groupId = (row.attrs['sdt'] as SdtGroup[])[0]!.groupId;
    expect(preview.controls(groupId)?.form).toBe('row');
    const controls = listControls(preview.doc, preview.controls);
    expect(controls.map((c) => [c.form, c.tag.role])).toEqual([['block', 'none'], ['row', 'repeat'], ['cell', 'none'], ['cell', 'none'], ['block', 'condition'], ['block', 'condition']]);
    expect(session.template.verdictFor(controls[5]!.sdtPr)).toMatchObject({ kept: false });
    // Nothing of it touched the session: ids of its own, the export unchanged, the save byte for byte.
    expect(preview.ctx.ids).not.toBe(session.ctx.ids);
    expect(session.exportTree().changed).toBe(false);
    const saved = await session.save();
    expect(saved.changed).toBe(false);
    expect(comparePackages(original, saved.bytes)).toEqual([]);
    // The data part edited: the next preview shows the new value; an unsaved edit to the document shows too.
    session.template.dataPart.part!.selectSingleNode('/invoice/customer/name')!.text = 'Jane Doe';
    session.template.invalidate();
    const start = blockStart(session, 2) + 1;
    expect(runCommand(session, 'insertText', { text: 'NEW ' }, { from: start }).ok).toBe(true);
    const again = await previewInstance(session);
    expect(again.doc.child(3).textContent).toBe('Jane Doe');
    expect(again.doc.child(2).textContent.startsWith('NEW ')).toBe(true);
    expect(session.doc.child(3).textContent).toBe('Joe Bloggs');   // the live document keeps its own text
  });

  test('an empty value shows the placeholder (REQ-024); a plain document previews as itself', async () => {
    const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
    session.pkg.customXmlParts.add('<data><name></name><city>Perth</city></data>');
    await session.template.load({ engine: false });
    const start = blockStart(session, 0) + 1;
    expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: start, to: start + 3 }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/name' }, { from: start + 2 }).ok).toBe(true);
    const preview = await previewInstance(session);
    expect(preview.result).toMatchObject({ bound: 1 });
    expect(preview.doc.child(0).child(0).textContent).toBe('Click or tap here to enter text.');   // the engine's bind pass writes Word's placeholder since core-ts 0.6.0 (CR-002 section 63.22), as the editor's own does since ED-008 section 8.8 item 14 (PLACEHOLDER_TEXT)
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/city' }, { from: start + 2 }).ok).toBe(true);
    expect((await previewInstance(session)).doc.child(0).child(0).textContent).toBe('Perth');
  });

  // ED-005 section 12.117: an XHTML-bound control converted as the processor converts it. core-ts check 37's template, whose two
  // controls bound through od:xpath entries were never filled; Node has no DOMParser, so the parser is jsdom's (a browser's by default).
  test('check37a-template.docx: XHTML-bound controls converted, the list labelled from the clone\'s numbering, the document untouched', async () => {
    const original = await fixture('xhtml/check37a-template.docx');
    const session = await EditorDocument.open(original);
    await session.template.load();
    const preview = await previewInstance(session, { html: { parser: (html: string) => new JSDOM(html).window.document } });
    expect(preview.result).toMatchObject({ bound: 3, updated: 3, skipped: 0 });
    // The run-level control keeps the first paragraph's inlines; the processor's note says what it dropped.
    expect(preview.result.notes).toEqual(['od:xpath=x2&od:ContentType=application/xhtml+xml: dropped blocks after the first, in a run-level control (1)']);
    const paragraphs: { style: string | null; label: string | null; text: string }[] = [];
    preview.doc.descendants((n) => {
      if (n.type.name !== 'paragraph') return true;
      paragraphs.push({ style: (n.attrs['pStyle'] as string | null) ?? null, label: preview.labels.get(n)?.listString ?? null, text: n.textContent });
      return false;
    });
    expect(paragraphs.slice(2, 12)).toEqual([
      { style: 'Heading1', label: null, text: 'Delivery terms' },
      { style: null, label: null, text: 'Goods are delivered within 14 days of the order, see the terms.' },
      { style: null, label: '\u2022', text: 'Standard delivery is free.' },
      { style: null, label: '\u2022', text: 'Express delivery costs extra.' },
      { style: null, label: null, text: 'Service' },
      { style: null, label: null, text: 'Days' },
      { style: null, label: null, text: 'Standard' },
      { style: null, label: null, text: '14' },
      { style: null, label: null, text: 'Express' },
      { style: null, label: null, text: '2' },
    ]);
    expect(paragraphs[12]!.text).toBe('2. A run-level control (should show the first paragraph\'s text after the bind): [First paragraph in italics');
    // The labels are the clone's: the file has no numbering part, so the session's own walk labels nothing.
    expect(session.main.numberingDefinitionsPart).toBeUndefined();
    expect(preview.numbering).toBeDefined();
    // Nothing reached the session: the document's controls keep their placeholders, the package is the file's.
    expect(session.doc.textContent).toContain('XHTML placeholder (block-level)');
    const saved = await session.save();
    expect(saved.changed).toBe(false);
    expect(comparePackages(original, saved.bytes)).toEqual([]);
  });
});
