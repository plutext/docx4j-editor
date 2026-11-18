// SPDX-License-Identifier: GPL-3.0-or-later
//
// Line-ending characters in `w:t` survive a re-marshal (@docx4j/jsonix 3.3.0, reached through the
// object model's range): the NEL (U+0085) in tracked-changes.docx's "and here it continues" used
// to become U+000A on unmarshalling, which an edited save then wrote. Core-ts's parity goldens
// caught it; this holds the editor's side.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage, textOf as textOfElement } from '@docx4j/core-ts';
import { EditorDocument, runCommand } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

const NEL = String.fromCharCode(0x85);

describe('line-ending characters', () => {
  test('the NEL in tracked-changes.docx is kept on screen, through an edited save, and on reload', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    let index = -1;
    session.doc.forEach((node, _pos, i) => { if (index < 0 && node.textContent.includes('here it continues')) index = i; });
    expect(index).toBeGreaterThanOrEqual(0);
    expect(session.doc.child(index).textContent).toContain(NEL);
    // An edit elsewhere re-marshals the whole main part; the character is what the file had.
    runCommand(session, 'insertText', { text: '!' }, { from: 1 });
    const saved = await session.save();
    const pkg = await WordprocessingMLPackage.load(saved.bytes);
    const body = (await pkg.getMainDocumentPart().readContents()).body!;
    const text = textOfElement(body.content![index]!);
    expect(text).toContain(NEL);
    expect(text).not.toContain(NEL + '\n');   // not doubled into a line feed either
  });
});
