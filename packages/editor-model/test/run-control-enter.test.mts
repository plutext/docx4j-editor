// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.123 (Jason, 2026-11-12, tmp/testcase_sdt-replicates.docx): Enter never splits a run-level control. Split,
// the new paragraph held an empty copy with the control's `w:id`, tag and binding. At a control's inner edge Enter splits
// outside it; in its middle, a plain text control that allows line breaks takes one, and any other control takes nothing.
import { describe, test, expect } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, readControl } from '../src/index.mjs';
import { fixture, mainXml } from './helpers.mjs';

/** The run-level controls of the document, in order: their kind, `w:id`, text and the paragraph they are in. */
function controls(doc: PMNode): { kind: string; id: number | undefined; text: string; paragraph: number }[] {
  const out: { kind: string; id: number | undefined; text: string; paragraph: number }[] = [];
  doc.forEach((block, _pos, paragraph) => block.descendants((n) => {
    if (n.type.name === 'sdtInline') { const r = readControl(n.attrs['sdtPr']); out.push({ kind: r.kind, id: r.id, text: n.textContent, paragraph }); }
    return true;
  }));
  return out;
}

/** The start of the first top-level paragraph's content and that paragraph. */
function firstParagraph(session: EditorDocument): { from: number; node: PMNode } {
  let found: { from: number; node: PMNode } | undefined;
  session.doc.forEach((node, pos) => { if (!found && node.type.name === 'paragraph' && node.textContent.length > 10) found = { from: pos + 1, node }; });
  return found!;
}

/** Each run-level control's place: where it starts (before it) and where its content starts and ends. */
function places(doc: PMNode): { before: number; start: number; end: number }[] {
  const out: { before: number; start: number; end: number }[] = [];
  doc.descendants((n, pos) => { if (n.type.name === 'sdtInline') out.push({ before: pos, start: pos + 1, end: pos + 1 + n.content.size }); return true; });
  return out;
}

/** A session whose first long paragraph holds a control of `kind` over its last `n` characters (or its first, with `atStart`). */
async function withControl(kind: string, options: { atStart?: boolean; n?: number; multiLine?: boolean } = {}): Promise<{ session: EditorDocument; paragraph: number }> {
  const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
  const { from, node } = firstParagraph(session);
  const n = options.n ?? 4;
  const range = options.atStart ? { from, to: from + n } : { from: from + node.content.size - n, to: from + node.content.size };
  expect(runCommand(session, 'contentControl.insert', { kind }, range).ok).toBe(true);
  if (options.multiLine) expect(runCommand(session, 'contentControl.properties', { multiLine: true }, { from: places(session.doc)[0]!.start + 1 }).ok).toBe(true);
  let paragraph = 0;
  session.doc.forEach((b, _p, i) => { if (b.textContent === node.textContent) paragraph = i; });
  return { session, paragraph };
}

describe('Enter and a run-level control (ED-005 section 12.123)', () => {
  test("at the control's inner end, the end of the paragraph: a new empty paragraph after it, the control whole and alone in the file", async () => {
    const { session, paragraph } = await withControl('PlainText');
    const before = controls(session.doc);
    expect(runCommand(session, 'insertParagraph', undefined, { from: places(session.doc)[0]!.end }).ok).toBe(true);
    expect(controls(session.doc)).toEqual(before);
    expect(session.doc.child(paragraph + 1).type.name).toBe('paragraph');
    expect(session.doc.child(paragraph + 1).childCount).toBe(0);
    const xml = await mainXml((await session.save()).bytes);
    expect(xml.match(/<w:sdt>/g)).toHaveLength(1);
  });

  test("at the control's inner start, the start of the paragraph: a new empty paragraph before it, the control moving with its text", async () => {
    const { session, paragraph } = await withControl('PlainText', { atStart: true });
    const text = session.doc.child(paragraph).textContent;
    expect(runCommand(session, 'insertParagraph', undefined, { from: places(session.doc)[0]!.start }).ok).toBe(true);
    expect(session.doc.child(paragraph).childCount).toBe(0);
    expect(session.doc.child(paragraph + 1).textContent).toBe(text);
    expect(controls(session.doc)).toHaveLength(1);
    expect(controls(session.doc)[0]!.paragraph).toBe(paragraph + 1);
  });

  test("the control's text selected to its end: the text deleted, the paragraph split after the control", async () => {
    const { session, paragraph } = await withControl('PlainText', { n: 6 });
    const place = places(session.doc)[0]!;
    expect(runCommand(session, 'insertParagraph', undefined, { from: place.start + 2, to: place.end }).ok).toBe(true);
    expect(controls(session.doc).map((c) => [c.text.length, c.paragraph])).toEqual([[2, paragraph]]);
    expect(session.doc.child(paragraph + 1).childCount).toBe(0);
  });

  test('in the middle of a plain text control: nothing, as Word takes no paragraph mark there; with w:multiLine, a line break inside it', async () => {
    const plain = await withControl('PlainText');
    const doc = plain.session.doc;
    expect(runCommand(plain.session, 'insertParagraph', undefined, { from: places(doc)[0]!.start + 2 }).ok).toBe(false);
    expect(plain.session.doc.eq(doc)).toBe(true);

    const multi = await withControl('PlainText', { multiLine: true });
    const count = multi.session.doc.childCount;
    expect(runCommand(multi.session, 'insertParagraph', undefined, { from: places(multi.session.doc)[0]!.start + 2 }).ok).toBe(true);
    expect(multi.session.doc.childCount).toBe(count);
    const control = multi.session.doc.child(multi.paragraph).lastChild!;
    expect(control.type.name).toBe('sdtInline');
    expect(control.content.content.map((n) => (n.isText ? n.text : n.type.name))).toEqual([expect.any(String), 'hardBreak', expect.any(String)]);
  });

  test('in the middle of a rich text control: nothing (what Word does there is not measured)', async () => {
    const { session } = await withControl('RichText');
    const doc = session.doc;
    expect(runCommand(session, 'insertParagraph', undefined, { from: places(doc)[0]!.start + 2 }).ok).toBe(false);
    expect(session.doc.eq(doc)).toBe(true);
  });

  test('nested: at the inner end of a control that ends its outer control, the split is after both', async () => {
    const { session, paragraph } = await withControl('RichText', { n: 8 });
    const outer = places(session.doc)[0]!;
    expect(runCommand(session, 'contentControl.insert', { kind: 'PlainText' }, { from: outer.end - 3, to: outer.end }).ok).toBe(true);
    const before = controls(session.doc);
    expect(before.map((c) => c.kind)).toEqual(['RichText', 'PlainText']);
    expect(runCommand(session, 'insertParagraph', undefined, { from: places(session.doc)[1]!.end }).ok).toBe(true);
    expect(controls(session.doc)).toEqual(before);
    expect(session.doc.child(paragraph + 1).childCount).toBe(0);
  });
});
