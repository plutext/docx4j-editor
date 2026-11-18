// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 7.3: the portability lint over pasted Office JS.
import { describe, test, expect } from 'vitest';
import { lintOfficeJs } from '../src/app/office-js-lint.js';

const ADDIN = `await Word.run(async (context) => {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load('items/text,items/style');
  await context.sync();
  for (const p of paragraphs.items) if (p.style === 'Heading 1') p.insertText(' (reviewed)', 'End');
  await context.sync();
});`;

describe('the Office JS portability lint', () => {
  test('a script that loads and syncs before it reads is clean', () => {
    expect(lintOfficeJs(ADDIN).warnings).toEqual([]);
  });

  test('a read of a loadable property without a load is reported with its line, and the load is offered before the loop', () => {
    const code = `await Word.run(async (ctx) => {
  const paragraphs = ctx.document.body.paragraphs;
  for (const p of paragraphs.items) {
    if (p.style === 'Heading 1') console.log(p.text);
  }
});`;
    const r = lintOfficeJs(code);
    expect(r.warnings.map((w) => [w.line, w.object, w.property])).toEqual([[4, 'p', 'style'], [4, 'p', 'text']]);
    expect(r.warnings[0]!.message).toContain("load('items/style')");
    expect(r.insertion).toEqual({ line: 3, text: "  paragraphs.load('items/style,items/text');\n  await ctx.sync();\n" });
  });

  test('a load without a sync before the read is reported as such; a write needs no load; a call is not a read', () => {
    const code = `await Word.run(async (context) => {
  const first = context.document.body.paragraphs.getFirst();
  first.load('text');
  console.log(first.text);
  first.style = 'Heading 1';
  first.insertText('x', 'End');
});`;
    const r = lintOfficeJs(code);
    expect(r.warnings.map((w) => w.line)).toEqual([4]);
    expect(r.warnings[0]!.message).toContain('sync()');
    expect(r.insertion).toBeUndefined();
  });

  test('a script without Word.run is left alone, and the console bindings are not Office JS objects', () => {
    expect(lintOfficeJs("const p = body.paragraphs[0]; p.text = 'x'; console.log(p.text);")).toEqual({ warnings: [] });
  });

  test('a loop over a dotted collection is attributed to it, and the insertion loads that collection', () => {
    const code = `await Word.run(async (context) => {
  for (const p of context.document.body.paragraphs.items) console.log(p.text);
});`;
    const r = lintOfficeJs(code);
    expect(r.warnings[0]!.message).toContain("load('items/text')");
    expect(r.insertion).toEqual({ line: 2, text: "  context.document.body.paragraphs.load('items/text');\n  await context.sync();\n" });
  });

  test('a variable taken from a collection by index inherits the collection\'s items loads', () => {
    const code = `await Word.run(async (context) => {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load('items/text');
  await context.sync();
  const p = paragraphs.items[0];
  return p.text;
});`;
    expect(lintOfficeJs(code).warnings).toEqual([]);
  });
});
