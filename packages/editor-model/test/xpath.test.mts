// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 3.5: Select by XPath. Over every fixture and a set of expressions, the hits are
// the objects the engine's CR-006 selection returns for the same part (`pkg.selectObjects`, by
// their paths), less those above the blocks (the root, the body, the final section), and each is
// placed where its element projects: a paragraph at a paragraph, a table at a table, a run at its
// text. After an edit the expression sees the edited text, placed at the edited block; the
// document's bytes are untouched by a select; a notes or header part is selected over as well.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, runCommand, type XPathHit } from '../src/index.mjs';
import { fixture, allFixtures, comparePackages } from './helpers.mjs';

const EXPRESSIONS = [
  '//w:p', '//w:p[w:pPr/w:pStyle]', '//w:r[w:rPr/w:b]', '//w:tbl', '//w:tr', '//w:tc', '//w:sdt', '//w:hyperlink',
  "//w:t[contains(., 'e')]", '//w:t/text()', '//w:pStyle/@w:val', '//w:bookmarkStart', '//w:ins', '//w:del', '//w:drawing', '//w:footnoteReference',
];

/** A path the editor places nowhere: the root, the body element, the body's final section and what is in it. */
const aboveBlocks = (path: string): boolean => /^\/w:document\[1\](\/w:body\[1\])?$/.test(path) || path.startsWith('/w:document[1]/w:body[1]/w:sectPr[');

/** What a hit's name says its node should be, where it is placed exactly. */
const NODE_OF: Record<string, string[]> = {
  'w:p': ['paragraph'], 'w:tbl': ['table'], 'w:tr': ['table_row'], 'w:tc': ['table_cell'], 'w:sdt': ['sdtBlock', 'sdtInline', 'table_row', 'table_cell'],
};

describe('Select by XPath (ED-005 section 3.5)', () => {
  test('every fixture: the engine\'s objects, each placed where its element projects', async () => {
    const findings: string[] = [];
    let hits = 0;
    let compared = 0;
    for (const name of await allFixtures()) {
      const bytes = await fixture(name);
      const session = await EditorDocument.open(bytes);
      const pkg = await WordprocessingMLPackage.load(bytes);
      for (const xpath of EXPRESSIONS) {
        const ours = await session.selectXPath(xpath);
        const theirs = (await pkg.selectObjects(xpath)).map((b) => b.path).filter((p) => !aboveBlocks(p));
        compared++;
        hits += ours.length;
        if (JSON.stringify(ours.map((h) => h.path)) !== JSON.stringify(theirs)) findings.push(`${name} ${xpath}: engine ${theirs.length} ${JSON.stringify(theirs.slice(0, 3))}, editor ${ours.length} ${JSON.stringify(ours.slice(0, 3).map((h) => h.path))}`);
        const doc = session.doc;
        for (const h of ours) {
          if (!h.exact) { findings.push(`${name} ${xpath}: ${h.path} not placed exactly in an unedited document`); continue; }
          const node = doc.nodeAt(h.pos);
          const want = NODE_OF[h.name];
          // A text box's paragraph projects as part of its drawing, an atom; a merge's continuation cell as the cell it folds into.
          const inTextBox = /\/w:(drawing|pict)\[/.test(h.path) && node?.isAtom === true;
          const inMerge = node?.type.name === 'table_cell' && node.attrs['merged'] !== null;
          if (want && !want.includes(node?.type.name ?? '') && !inTextBox && !inMerge) findings.push(`${name} ${xpath}: ${h.path} placed on ${node?.type.name}`);
          if (h.name === 'w:r' && !(node?.isText || node?.attrs['runSource'] !== undefined)) findings.push(`${name} ${xpath}: ${h.path} placed on ${node?.type.name}`);
          if (doc.resolve(Math.min(h.pos, doc.content.size)).index(0) !== h.block) findings.push(`${name} ${xpath}: ${h.path} block ${h.block}`);
        }
      }
    }
    expect(compared).toBeGreaterThan(500);
    expect(hits).toBeGreaterThan(3000);
    expect(findings).toEqual([]);
  });

  test('after an edit the expression sees the text on screen, placed at the edited block; a select leaves the bytes alone', async () => {
    const bytes = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(bytes);
    const untouched = await session.save();
    expect(comparePackages(bytes, untouched.bytes)).toEqual([]);
    await session.selectXPath('//w:p');
    const saved = await session.save();
    expect(saved.changed).toBe(false);
    expect(comparePackages(bytes, saved.bytes)).toEqual([]);
    let b4 = 0;
    for (let k = 0; k < 4; k++) b4 += session.doc.child(k).nodeSize;
    runCommand(session, 'insertText', { text: 'Zebra ' }, { from: b4 + 1 });
    const found: XPathHit[] = await session.selectXPath("//w:t[contains(., 'Zebra')]");
    expect(found.length).toBe(1);
    expect(found[0]!.block).toBe(4);
    expect(found[0]!.exact).toBe(false);
    expect(found[0]!.pos).toBe(b4);
    expect(found[0]!.snippet).toBe('Zebra Some text.');
    // The edited document saves as it would have without the select.
    const after = await session.save();
    expect(after.changed).toBe(true);
  });

  test('a header part and a notes part; an invalid expression throws', async () => {
    const session = await EditorDocument.open(await fixture('headers.docx'));
    const parts = await session.stories.selectableParts();
    expect(parts).toContain('/word/header1.xml');
    const inHeader = await session.selectXPath('//w:p', { part: '/word/header1.xml' });
    expect(inHeader.length).toBeGreaterThan(0);
    expect(inHeader[0]!.story?.key).toBe('header:/word/header1.xml');
    expect(inHeader[0]!.exact).toBe(true);
    const notes = await EditorDocument.open(await fixture('notes-numbering.docx'));
    const inNotes = await notes.selectXPath("//w:p[contains(., 'Footnote 3')]", { part: '/word/footnotes.xml' });
    expect(inNotes.map((h) => h.story?.key)).toEqual(['footnote:3']);
    await expect(session.selectXPath('//w:p[')).rejects.toThrow();
    await expect(session.selectXPath('//w:p', { part: '/word/styles.xml' })).rejects.toThrow(/not a part/);
  });
});
