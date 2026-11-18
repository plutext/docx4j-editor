// SPDX-License-Identifier: GPL-3.0-or-later
//
// Markdown in (ED-005 section 12.66): the reading the agent, paste and Open share. The agent's is
// strict and refuses as it always has (its tools' tests hold the companion's answers); paste's and
// Open's are lenient. Word's built-in styles the engine cannot splice come from `knownstyles.mts`
// (UPSTREAM-20). A `.md` file opens as a document the engine built.
import { describe, expect, test } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, readMarkdown, looksLikeMarkdown, markdownDocument, ensureStylesIn, MarkdownUnsupported, tableColumnWidths } from '../src/index.mjs';

const strictRefusal = (markdown: string): string | undefined => {
  try { readMarkdown(markdown); return undefined; } catch (e) { if (e instanceof MarkdownUnsupported) return e.construct; throw e; }
};

describe('the strict reading (the agent)', () => {
  test('the subset reads as before', () => {
    const { blocks, dropped } = readMarkdown('# Title\n\nSome **bold** and *it* and ~~gone~~ [here](https://example.com).\n\n- one\n  - nested\n1. first\n\n| a | b |\n|---|---|\n| 1 | 2 |\n');
    expect(dropped).toEqual({});
    expect(blocks.map((b) => (b.kind === 'paragraph' ? [b.heading ?? null, b.list?.level ?? null] : 'table'))).toEqual([[1, null], [null, null], [null, 0], [null, 1], [null, 0], 'table']);
    const para = blocks[1]!;
    expect(para.kind === 'paragraph' && para.inlines).toEqual([
      { text: 'Some ' }, { bold: true, text: 'bold' }, { text: ' and ' }, { italic: true, text: 'it' }, { text: ' and ' }, { strike: true, text: 'gone' }, { text: ' ' }, { link: 'https://example.com', text: 'here' }, { text: '.' },
    ]);
  });

  test('a hard-wrapped paragraph is one paragraph: a soft break is a space, a hard break a line break (Jason, 2026-10-12)', () => {
    const { blocks } = readMarkdown('Drafted on request ("Yes;\nplease draft\nCR-036"), **two\nwords** then.\n\nhard  \nbreak\\\nhere\n');
    expect(blocks.map((b) => b.kind === 'paragraph' && b.inlines)).toEqual([
      [{ text: 'Drafted on request ("Yes; please draft CR-036"), ' }, { bold: true, text: 'two words' }, { text: ' then.' }],
      [{ text: 'hard' }, { text: '\n' }, { text: 'break' }, { text: '\n' }, { text: 'here' }],
    ]);
  });

  test('anything else is refused, naming the construct', () => {
    expect(strictRefusal('Use `npm`.')).toBe('inline code');
    expect(strictRefusal('```\ncode\n```')).toBe('a code block');
    expect(strictRefusal('> quoted')).toBe('a block quote');
    expect(strictRefusal('![alt](x.png)')).toBe('an image');
    expect(strictRefusal('[a][r]\n\n[r]: https://example.com')).toBe('a reference-style link');
    expect(strictRefusal('a\n\n---\n\nb')).toBe('a thematic break');
    expect(strictRefusal('[file](notes.md)')).toBe('a link to "notes.md" (only http, https and mailto)');
    expect(strictRefusal('a <b>x</b>')).toBe('HTML');
  });
});

describe('the lenient reading (paste and Open)', () => {
  test('code, quotes, reference links and <br> have a place', () => {
    const { blocks, dropped } = readMarkdown('Run `npm test` [now][go].\n\n```ts\nconst a = 1;\n\n\tb();\n```\n\n> A quote\n> continued\n\n| a |\n|---|\n| x<br>y |\n\n[go]: https://example.com/go\n', { lenient: true });
    expect(dropped).toEqual({});
    expect(blocks[0]).toEqual({ kind: 'paragraph', inlines: [{ text: 'Run ' }, { code: true, text: 'npm test' }, { text: ' ' }, { link: 'https://example.com/go', text: 'now' }, { text: '.' }] });
    expect(blocks.slice(1, 4)).toEqual([
      { kind: 'paragraph', code: true, inlines: [{ text: 'const a = 1;' }] },
      { kind: 'paragraph', code: true, inlines: [] },
      { kind: 'paragraph', code: true, inlines: [{ text: '\tb();' }] },
    ]);
    expect(blocks[4]).toEqual({ kind: 'paragraph', quote: true, inlines: [{ text: 'A quote continued' }] });
    expect(blocks[5]).toEqual({ kind: 'table', rows: [[[{ text: 'a' }]], [[{ text: 'x' }, { text: '\n' }, { text: 'y' }]]] });
  });

  test('what has no place is dropped and counted, its text kept', () => {
    const { blocks, dropped } = readMarkdown('See ![a diagram](d.png) and [the notes](notes.md) <kbd>Ctrl</kbd>.\n\n---\n\n<div>block</div>\n', { lenient: true });
    expect(dropped).toEqual({ images: 1, links: 1, HTML: 3, rules: 1 });
    expect(blocks).toHaveLength(1);
    const first = blocks[0]!;
    expect(first.kind === 'paragraph' && first.inlines.map((i) => i.text).join('')).toBe('See  and the notes Ctrl.');
  });

  test('a code block in a list item follows the item', () => {
    const { blocks } = readMarkdown('- step\n\n  ```\n  npm ci\n  ```\n- next\n', { lenient: true });
    expect(blocks.map((b) => b.kind === 'paragraph' && [b.list?.level ?? null, b.code ?? false])).toEqual([[0, false], [null, true], [0, false]]);
  });
});

describe('looksLikeMarkdown', () => {
  test.each([
    '# Heading', '## Two\ntext', '- item', '* item', '1. first', '2) second', '> quoted', '```\ncode\n```', '| a | b |', 'Some **bold** words', 'Run `npm ci` first', 'Read [this](https://x.y)',
  ])('%j reads as Markdown', (text) => expect(looksLikeMarkdown(text)).toBe(true));
  test.each([
    'Hello world.', 'A sentence with * one star and 3.5 numbers.', 'Two lines\nof plain text', '#hashtag', 'C# and F#', 'costs 5 - 10 euros',
  ])('%j does not', (text) => expect(looksLikeMarkdown(text)).toBe(false));
});

describe("Word's built-in styles the engine lacks (UPSTREAM-20)", () => {
  test('spliced with what they link to, once', async () => {
    const pkg = await WordprocessingMLPackage.createPackage();
    expect(await ensureStylesIn(pkg, ['HTMLCode', 'HTMLPreformatted', 'Quote', 'Heading5', 'Heading1'])).toEqual(['HTMLCode', 'HTMLPreformattedChar', 'HTMLPreformatted', 'QuoteChar', 'Quote', 'Heading5Char', 'Heading5']);
    expect(await ensureStylesIn(pkg, ['HTMLCode', 'Quote'])).toEqual([]);
    const resolver = pkg.propertyResolverOrUndefined!;
    expect(resolver.getStyle('HTMLPreformatted')?.name?.val).toBe('HTML Preformatted');
    // Heading 5's definition names no list (KnownStyles.xml's numPr is dropped).
    expect(resolver.getStyle('Heading5')?.pPr?.numPr).toBeUndefined();
  });

  test("the engine's own definitions win, and an id neither has is skipped", async () => {
    const pkg = await WordprocessingMLPackage.createPackage();
    expect(await ensureStylesIn(pkg, ['FootnoteText', 'NoSuchStyle'])).toEqual(['FootnoteTextChar', 'FootnoteText']);
  });

  test("a document's own definition is left as it is", async () => {
    const pkg = await WordprocessingMLPackage.createPackage();
    await ensureStylesIn(pkg, ['Quote']);
    const quote = pkg.propertyResolverOrUndefined!.getStyle('Quote')!;
    quote.uiPriority = { val: 1 } as typeof quote.uiPriority;
    expect(await ensureStylesIn(pkg, ['Quote'])).toEqual([]);
    expect(pkg.propertyResolverOrUndefined!.getStyle('Quote')?.uiPriority?.val).toBe(1);
  });
});

describe('a .md file opened', () => {
  const blocksOf = (doc: PMNode): PMNode[] => { const out: PMNode[] = []; doc.forEach((n) => out.push(n)); return out; };

  test('as a document the engine built, with the styles, lists and links it needs', async () => {
    const { bytes, dropped } = await markdownDocument('# Notes\n\nRun `npm ci` and see [docs](https://example.com).\n\n- a\n- b\n\n1. x\n\n```\nnpm test\n```\n\n> said\n\n###### Six\n\n| h |\n|---|\n| c |\n', { pageSize: 'A4' });
    expect(dropped).toEqual({});
    const session = await EditorDocument.open(bytes);
    const blocks = blocksOf(session.doc);
    expect(blocks.map((b) => (b.type.name === 'paragraph' ? b.attrs['pStyle'] ?? (b.attrs['numPr'] ? 'list' : null) : b.type.name))).toEqual(['Heading1', null, 'list', 'list', 'list', 'HTMLPreformatted', 'Quote', 'Heading6', 'table']);
    const code: string[] = [];
    let link: Record<string, unknown> | undefined;
    blocks[1]!.descendants((n) => {
      if (n.marks.some((m) => m.type.name === 'runStyle' && m.attrs['id'] === 'HTMLCode')) code.push(n.text ?? '');
      link ??= n.marks.find((m) => m.type.name === 'link')?.attrs;
      return true;
    });
    expect(code).toEqual(['npm ci']);
    expect(link?.['rId']).toBeTruthy();
    // The bullets and the numbered list are two definitions.
    expect(blocks[2]!.attrs['numPr'].numId).toBe(blocks[3]!.attrs['numPr'].numId);
    expect(blocks[4]!.attrs['numPr'].numId).not.toBe(blocks[2]!.attrs['numPr'].numId);
    // Saved unedited, it keeps its bytes.
    const saved = await session.save();
    expect(saved.changed).toBe(false);
  });
});

describe('a table\'s columns are sized by their longest cell (ED-005 section 12.92)', () => {
  test('the weights are the longest cell per column, between twelve and sixty characters, the widths summing to the text width', () => {
    const rows = [[[{ text: 'Phase' }], [{ text: 'What' }], [{ text: 'Effort' }]], [[{ text: '0' }], [{ text: 'A'.repeat(200) }], [{ text: '2-3' }]]];
    const widths = tableColumnWidths(rows, 3, 9026);
    expect(widths.reduce((a, b) => a + b, 0)).toBe(9026);
    expect(widths[1]!).toBeGreaterThan(widths[0]! * 4);   // 60 against 12 and 12
    expect(Math.abs(widths[2]! - widths[0]!)).toBeLessThan(3);   // the remainder on the last column alone
    expect(tableColumnWidths([[[{ text: 'a' }], [{ text: 'b' }]]], 2, 9026)).toEqual([4513, 4513]);
  });
  test('a Markdown table gets that grid, on the page size\'s text width', async () => {
    const md = '| Phase | What | Effort |\n|---|---|---|\n| 0 | ' + 'ChartStylePart/ChartColorStylePart/ChartExSpacePart content and more words here' + ' | 2-3 |\n';
    const doc = await markdownDocument(md, { pageSize: 'LETTER' });
    const session = await EditorDocument.open(doc.bytes);
    const table = session.doc.child(0);
    expect(table.type.name).toBe('table');
    const grid = table.attrs['tblGrid'] as number[];
    expect(grid.reduce((a, b) => a + b, 0)).toBe(12240 - 2880);
    expect(grid[1]!).toBeGreaterThan(grid[0]! * 4);
  });
  test('with narrower margins asked (the application\'s option, ED-007 section 8.1 item 15) the section has them and the grid fills the wider text width', async () => {
    const md = '| a | b |\n|---|---|\n| 1 | 2 |\n';
    const doc = await markdownDocument(md, { pageSize: 'LETTER', margins: 720 });
    const session = await EditorDocument.open(doc.bytes);
    const sectPr = session.doc.attrs['finalSectPr'] as { pgMar?: Record<string, number> };
    expect(sectPr.pgMar).toMatchObject({ top: 720, right: 720, bottom: 720, left: 720, header: 708, footer: 708 });
    const grid = session.doc.child(0).attrs['tblGrid'] as number[];
    expect(grid.reduce((a, b) => a + b, 0)).toBe(12240 - 1440);
    // the default is the engine's 2.54 cm
    const plain = await EditorDocument.open((await markdownDocument(md, { pageSize: 'LETTER' })).bytes);
    expect((plain.doc.attrs['finalSectPr'] as { pgMar?: Record<string, number> }).pgMar).toMatchObject({ top: 1440, left: 1440 });
  });
});

