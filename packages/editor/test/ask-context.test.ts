// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 7.5: the context builders and the request they make.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { EditorDocument, stateFor, documentStyles } from '@docx4j/editor-model';
import type { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { outlineOf, stylesLine, apiDeclarations, buildContext, buildRequest, chatPrompt, DECLARATIONS_CAP } from '../src/app/ask-context.js';
import { stylesInUse } from '../src/view/css.js';
import type { EditorSurface } from '../src/view/surface.js';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

/** A surface for the builders: the state and the session's export, nothing else. */
function surfaceOver(session: EditorDocument, state: EditorState): EditorSurface {
  return { state, session, exportBlock: (n: PMNode) => session.exportBlock(n) } as unknown as EditorSurface;
}

function block(session: EditorDocument, i: number): { from: number; to: number } {
  const doc = session.doc;
  let pos = 0;
  for (let k = 0; k < i; k++) pos += doc.child(k).nodeSize;
  return { from: pos + 1, to: pos + 1 + doc.child(i).content.size };
}

describe('the ask context', () => {
  test('the outline lists the headings by style, indented by level', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'tracked-changes.docx'))));
    const outline = outlineOf(session.doc);
    expect(outline).toContain('Heading1: This is style Heading 1');
    expect(outline.split('\n').every((l) => /^\s*(Heading\d|Title|Subtitle): /.test(l))).toBe(true);
  });

  test('the styles line puts the styles in use first and names each with its type', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'tracked-changes.docx'))));
    const line = stylesLine(await documentStyles(session), stylesInUse(session, session.doc), 5);
    const lines = line.split('\n');
    expect(lines).toHaveLength(5);
    expect(lines[0]).toMatch(/\(.*, (paragraph|character), in use\)$/);
  });

  test('the declarations carry the content API and the shim, comments stripped, under the cap', async () => {
    const d = await apiDeclarations();
    expect(d.text).toContain('// model/content/Body.d.mts');
    expect(d.text).toContain('insertParagraph(');
    expect(d.text).toContain('// office-js/index.d.mts');
    expect(d.text).not.toContain('/**');
    expect(d.text.length).toBeLessThanOrEqual(DECLARATIONS_CAP);
    expect(d.capped).toBe(false);
  });

  test('the request carries the selection as text and XML, and the contract in the system prompt', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'tracked-changes.docx'))));
    const b4 = block(session, 4);
    const ctx = await buildContext(session, surfaceOver(session, stateFor(session, { from: b4.from + 5, to: b4.from + 9 })));
    expect(ctx.blocks).toBe(1);
    expect(ctx.selectionText).toBe('text');
    expect(ctx.selectionXml).toContain('<w:t>Some text.</w:t>');
    const req = buildRequest('claude-sonnet-5', 'Make it a heading', ctx);
    expect(req.system).toContain('fenced block tagged wml');
    expect(req.system).toContain('Heading1: This is style Heading 1');
    expect(req.prompt.startsWith('Make it a heading')).toBe(true);
    expect(req.prompt).toContain('Selected text:\ntext');
    expect(req.prompt).toContain('Selected XML (1 block)');
    expect(req.model).toBe('claude-sonnet-5');
  });
});

describe('the request for a chat (ED-003 section 11.25)', () => {
  test('one message: the system prompt, a rule, then the question with the selection', () => {
    const message = chatPrompt({ model: 'claude-sonnet-5', system: 'CONTRACT', prompt: 'QUESTION', maxTokens: 4096 });
    expect(message).toBe('CONTRACT\n\n---\n\nQUESTION');
  });

  // ED-008 proposal P4: the ask mode tells the model the accepted text, as the find bar reads it. A deletion still pending
  // in the document is not text it holds, in the outline or in the selection.
  test('the selection\'s text is the accepted text: a deletion pending in the document is left out (the outline reads through the same paragraphText)', async () => {
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, 'revisions', 'revisions-word15.docx'))));
    const doc = session.doc;
    let at: { from: number; to: number; raw: string } | undefined;
    doc.descendants((node, pos) => {
      if (at || node.type.name !== 'paragraph') return !at;
      if (node.textContent.startsWith('Cows graze in the green field')) at = { from: pos + 1, to: pos + 1 + node.content.size, raw: node.textContent };
      return false;
    });
    expect(at!.raw).toBe('Cows graze in the green field beside the river.');
    const ctx = await buildContext(session, surfaceOver(session, stateFor(session, { from: at!.from, to: at!.to })));
    expect(ctx.selectionText).toBe('Cows graze in the field beside the river.');
    // A selection that ends inside the deletion reads up to it.
    const part = await buildContext(session, surfaceOver(session, stateFor(session, { from: at!.from, to: at!.from + 'Cows graze in the gr'.length })));
    expect(part.selectionText).toBe('Cows graze in the ');
  });
});

