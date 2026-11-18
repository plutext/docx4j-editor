// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 6.1: decision 20 as runOnTree. A script edits the tree through the content
// API; the edit layer follows as one transaction; untouched blocks keep their nodes; undo takes
// the whole run back; a pending edit-layer change survives; a throwing script is reported.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage, textOf as textOfElement } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { EditorDocument, runCommand } from '../src/index.mjs';
import { fixture, comparePackages, plain } from './helpers.mjs';

async function reloadBody(bytes: Uint8Array): Promise<wml.Body> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  return (await pkg.getMainDocumentPart().readContents()).body!;
}

function withoutProofErr(value: unknown): unknown {
  if (Array.isArray(value)) return value.filter((v) => !(v && typeof v === 'object' && (v as { name?: { localPart?: string } }).name?.localPart === 'proofErr')).map(withoutProofErr);
  if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value as object).map(([k, v]) => [k, withoutProofErr(v)]));
  return value;
}

function block(session: EditorDocument, i: number): { from: number; to: number } {
  const doc = session.doc;
  let pos = 0;
  for (let k = 0; k < i; k++) pos += doc.child(k).nodeSize;
  return { from: pos + 1, to: pos + 1 + doc.child(i).content.size };
}

describe('runOnTree', () => {
  test('a script inserts, edits, deletes and restyles; the screen follows; one undo step; the export is right', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const before = [...session.body.content!];
    // Snapshots, not references: a script edits the original elements in place (that is the
    // point of the content API), so `before[i]` itself changes under the run.
    const snapshots = before.map((el) => withoutProofErr(plain(el)));
    const untouched = session.doc.child(2);
    const result = await session.runOnTree(({ body }) => {
      const p = body.paragraphs[4]!;   // "Some text."
      p.insertText(' via the API', 'End');
      const inserted = p.insertParagraph('Inserted by a script', 'After');
      inserted.style = 'Heading 2';
      body.paragraphs[1]!.delete();
      body.paragraphs[0]!.alignment = 'Right';
      return body.paragraphs.length;
    });
    expect(result.error).toBeUndefined();
    expect(result.changed).toBe(true);
    expect(session.treeLive).toBe(true);
    const doc = session.doc;
    expect(doc.childCount).toBe(before.length);   // one deleted, one inserted
    expect(doc.child(0).attrs['jc']).toBe('right');
    expect(doc.child(3).textContent).toBe('Some text. via the API');
    expect(doc.child(4).textContent).toBe('Inserted by a script');
    expect(doc.child(4).attrs['pStyle']).toBe('Heading2');
    // The heading the script did not touch is the very same node (same id, same element).
    expect(doc.child(1).eq(untouched)).toBe(true);
    expect(doc.child(1).attrs['sourceId']).toBe(untouched.attrs['sourceId']);
    // The screen is the baseline now: nothing is "changed" on export, yet the part is live.
    expect(session.exportTree().changed).toBe(false);
    const saved = await session.save();
    expect(comparePackages(original, saved.bytes, [session.main.partName.storeName, 'docProps/app.xml']).filter((d) => d.name !== 'docProps/app.xml')).toEqual([]);
    const after = (await reloadBody(saved.bytes)).content!;
    expect(textOfElement(after[3]!)).toBe('Some text. via the API');
    expect((after[4]!.value as wml.P).pPr?.pStyle?.val).toBe('Heading2');
    expect(withoutProofErr(plain(after[1]))).toEqual(snapshots[2]);
    // One undo step takes the whole run back, and the export is the original again.
    expect(session.undoManager.canUndo()).toBe(true);
    runCommand(session, 'history.undo');
    expect(session.doc.eq(session.imported)).toBe(false);   // the baseline is the post-run tree
    expect(session.doc.childCount).toBe(before.length);
    expect(session.doc.child(1).textContent).toBe(textOfElement(before[1]!));
    const restored = await session.save();
    const back = (await reloadBody(restored.bytes)).content!;
    // A block the script deleted comes back rebuilt from its element, so Word's proofErr marks
    // (never projected, ED-002 section 10.1 item 5) are gone from it; everything else is equal.
    snapshots.forEach((snapshot, i) => expect(withoutProofErr(plain(back[i]))).toEqual(snapshot));
  });

  test('a pending edit-layer change reaches the tree before the script runs', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const { from, to } = block(session, 4);
    runCommand(session, 'font.bold', { value: true }, { from, to });
    runCommand(session, 'insertText', { text: '!' }, { from: to });
    const seen = await session.runOnTree(({ body }) => ({ text: body.paragraphs[4]!.text, bold: body.paragraphs[4]!.font.bold }));
    expect(seen.value).toEqual({ text: 'Some text.!', bold: true });
    expect(seen.changed).toBe(false);
    expect(session.doc.child(4).textContent).toBe('Some text.!');
    const saved = await session.save();
    const after = (await reloadBody(saved.bytes)).content!;
    expect(textOfElement(after[4]!)).toBe('Some text.!');
  });

  test('a throwing script is reported, the tree is re-imported as it is, and undo takes it back', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    const result = await session.runOnTree(({ body }) => {
      body.paragraphs[4]!.insertText(' partly', 'End');
      throw new Error('boom');
    });
    expect((result.error as Error).message).toBe('boom');
    expect(result.changed).toBe(true);
    expect(session.doc.child(4).textContent).toBe('Some text. partly');
    runCommand(session, 'history.undo');
    expect(session.doc.child(4).textContent).toBe('Some text.');
    const saved = await session.save();
    expect(textOfElement((await reloadBody(saved.bytes)).content![4]!)).toBe('Some text.');
  });

  test('a run that changes nothing still makes the part live: the main part re-marshals and the self-check holds', async () => {
    const original = await fixture('invoice.docx');
    const session = await EditorDocument.open(original);
    const result = await session.runOnTree(({ body }) => body.text.length);
    expect(typeof result.value).toBe('number');
    expect(result.changed).toBe(false);
    expect(session.doc.eq(session.imported)).toBe(true);
    const saved = await session.save();
    expect(saved.selfCheck?.equal).toBe(true);
    const diffs = comparePackages(original, saved.bytes);
    expect(diffs.map((d) => d.name).sort()).toEqual(['docProps/app.xml', 'word/document.xml']);
  });

  test('the diagnostics follow a run that adds content: a table is projected (E2.a), a bookmark is opaque', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const before = session.diagnostics.opaque.length;
    await session.runOnTree(({ body }) => { body.insertTable(2, 2, 'End', [['a', 'b'], ['c', 'd']]); });
    expect(session.diagnostics.opaque.length).toBe(before);
    expect(session.doc.lastChild!.type.name).toBe('table');
    expect(session.doc.lastChild!.childCount).toBe(2);
    await session.runOnTree(async ({ body }) => { await body.insertXml('<w:p><w:bookmarkStart w:id="99" w:name="x"/><w:r><w:t>y</w:t></w:r><w:bookmarkEnd w:id="99"/></w:p>', 'End'); });
    expect(session.diagnostics.opaque.length).toBe(before + 2);
  });
});
