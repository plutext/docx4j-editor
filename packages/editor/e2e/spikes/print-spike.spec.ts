// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-006 step P0: the CSS level 1's printing rests on, measured before any of it is written. Each case is
// a static page shaped as the editor's DOM is (a `.pm` root of sibling blocks, a `.section-break` widget
// between sections, a `span.page-break` inside a paragraph). In Chromium the case is printed to PDF
// (`page.pdf` with `preferCSSPageSize`), and poppler's `pdfinfo` and `pdftotext` read each page's size
// and text. In Firefox, which Playwright cannot print to PDF, the support of each property is read
// through `CSS.supports` and the parsed rules, and the silent print-to-file preferences are tried. It
// reports; the findings go in ED-006 section 9. Run alone:
//   npx playwright test -c playwright.print-spike.config.ts
// Needs poppler (`pdfinfo`, `pdftotext`) on the host; not part of any suite.
import { test, expect, type Page } from '@playwright/test';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, existsSync, statSync, unlinkSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

function report(line: string): void {
  console.log(`spike [${test.info().project.name}] ${line}`);
}

const BASE = `
  body { margin: 0; font-family: 'DejaVu Sans', sans-serif; font-size: 11pt; line-height: 14pt; }
  p, h1 { margin: 0; }
  h1 { font-size: 14pt; line-height: 20pt; }
  .fill { display: block; }
  .section-break { display: none; }
`;

interface Case {
  name: string;
  css: string;
  body: string;
  /** What to read back: the pages' sizes, and on which page each probe's text lies. */
  probes: string[];
}

/** A paragraph of `n` lines, each `L<k>` on its own line. */
const lines = (prefix: string, n: number): string => Array.from({ length: n }, (_, i) => `${prefix}${i + 1}`).join('<br>');

const CASES: Case[] = [
  {
    name: 'A named pages by sibling rule (A4 portrait, A4 landscape, Letter)',
    css: `@page s0 { size: 595pt 842pt; margin: 72pt; } @page s1 { size: 842pt 595pt; margin: 72pt; } @page s2 { size: 612pt 792pt; margin: 54pt; }
      .pm > * { page: s0; }
      .pm > .section-break[data-section="0"] ~ * { page: s1; }
      .pm > .section-break[data-section="1"] ~ * { page: s2; }`,
    body: `<p>AAA-one</p><p>AAA-two</p><div class="section-break" data-section="0"></div><p>BBB-one</p><div class="section-break" data-section="1"></div><p>CCC-one</p>`,
    probes: ['AAA-one', 'AAA-two', 'BBB-one', 'CCC-one'],
  },
  {
    name: 'A2 a reused name across a continuous break: no page break',
    css: `@page s0 { size: 595pt 842pt; margin: 72pt; }
      .pm > * { page: s0; }
      .pm > .section-break[data-section="0"] ~ * { page: s0; }`,
    body: `<p>SAME-one</p><div class="section-break" data-section="0" data-break="continuous"></div><p>SAME-two</p>`,
    probes: ['SAME-one', 'SAME-two'],
  },
  {
    name: 'A3 a reused name with a next-page break: break-before on the adjacent block',
    css: `@page s0 { size: 595pt 842pt; margin: 72pt; }
      .pm > * { page: s0; }
      .pm > .section-break[data-section="0"] ~ * { page: s0; }
      .pm > .section-break[data-break="nextPage"] + * { break-before: page; }`,
    body: `<p>NEXT-one</p><div class="section-break" data-section="0" data-break="nextPage"></div><p>NEXT-two</p>`,
    probes: ['NEXT-one', 'NEXT-two'],
  },
  {
    name: 'B forced break on a sizeless block inside a paragraph (break-before)',
    css: `@page { size: 595pt 842pt; margin: 72pt; }
      .page-break { display: block; break-before: page; block-size: 0; overflow: hidden; font-size: 0; border: 0; margin: 0; }`,
    body: `<p>BEFORE-brk <span class="page-break" contenteditable="false">Page break</span> AFTER-brk</p>`,
    probes: ['BEFORE-brk', 'AFTER-brk', 'Page break'],
  },
  {
    name: 'B2 forced break inside a paragraph (break-after on the block)',
    css: `@page { size: 595pt 842pt; margin: 72pt; }
      .page-break { display: block; break-after: page; block-size: 0; overflow: hidden; font-size: 0; }`,
    body: `<p>BEFORE-brk2 <span class="page-break" contenteditable="false">Page break</span> AFTER-brk2</p>`,
    probes: ['BEFORE-brk2', 'AFTER-brk2'],
  },
  {
    name: 'C keep with next: break-after: avoid moves the heading with its paragraph',
    css: `@page { size: 595pt 842pt; margin: 72pt; }
      .fill { block-size: 640pt; } .keep { break-after: avoid; } .whole { break-inside: avoid; }`,
    body: `<div class="fill">filler</div><h1 class="keep">KEEP-heading</h1><p class="whole">${lines('KEEPBODY', 10)}</p>`,
    probes: ['KEEP-heading', 'KEEPBODY1'],
  },
  {
    name: 'C0 control: the same without break-after: avoid',
    css: `@page { size: 595pt 842pt; margin: 72pt; }
      .fill { block-size: 640pt; } .whole { break-inside: avoid; }`,
    body: `<div class="fill">filler</div><h1>FREE-heading</h1><p class="whole">${lines('FREEBODY', 10)}</p>`,
    probes: ['FREE-heading', 'FREEBODY1'],
  },
  {
    name: 'D orphans 2: one line of room at the foot',
    css: `@page { size: 595pt 842pt; margin: 72pt; }
      .fill { block-size: 684pt; } .para { orphans: 2; widows: 2; }`,
    body: `<div class="fill">filler</div><p class="para">${lines('ORPH', 6)}</p>`,
    probes: ['ORPH1', 'ORPH2', 'ORPH6'],
  },
  {
    name: 'D0 control: orphans 1',
    css: `@page { size: 595pt 842pt; margin: 72pt; }
      .fill { block-size: 684pt; } .para { orphans: 1; widows: 1; }`,
    body: `<div class="fill">filler</div><p class="para">${lines('ORPHONE', 6)}</p>`,
    probes: ['ORPHONE1', 'ORPHONE2', 'ORPHONE6'],
  },
  {
    name: 'D2 widows 2: room for five of six lines',
    css: `@page { size: 595pt 842pt; margin: 72pt; }
      .fill { block-size: 628pt; } .para { orphans: 2; widows: 2; }`,
    body: `<div class="fill">filler</div><p class="para">${lines('WID', 6)}</p>`,
    probes: ['WID4', 'WID5', 'WID6'],
  },
  {
    name: 'E header rows as table-header-group inside tbody',
    css: `@page { size: 595pt 842pt; margin: 72pt; }
      table { border-collapse: collapse; inline-size: 100%; } td { border: 1px solid #000; }
      tr.hdr { display: table-header-group; }`,
    body: `<table><tbody><tr class="hdr"><td>HDRA</td><td>HDRB</td></tr>${Array.from({ length: 120 }, (_, i) => `<tr><td>ROW${i + 1}</td><td>x</td></tr>`).join('')}</tbody></table>`,
    probes: ['HDRA', 'ROW1', 'ROW60', 'ROW120'],
  },
  {
    name: 'E0 control: header rows in a thead',
    css: `@page { size: 595pt 842pt; margin: 72pt; }
      table { border-collapse: collapse; inline-size: 100%; } td { border: 1px solid #000; }`,
    body: `<table><thead><tr><td>THDA</td><td>THDB</td></tr></thead><tbody>${Array.from({ length: 120 }, (_, i) => `<tr><td>TROW${i + 1}</td><td>x</td></tr>`).join('')}</tbody></table>`,
    probes: ['THDA', 'TROW1', 'TROW60', 'TROW120'],
  },
  {
    name: 'E3 a print-only thead copy of the header rows, the originals hidden in print',
    css: `@page { size: 595pt 842pt; margin: 72pt; }
      table { border-collapse: collapse; inline-size: 100%; } td { border: 1px solid #000; }
      thead.print-header { display: none; }
      @media print { thead.print-header { display: table-header-group; } tbody tr.hdr { display: none; } }`,
    body: `<table><thead class="print-header"><tr><td>CPYA</td><td>CPYB</td></tr></thead><tbody><tr class="hdr"><td>CPYA</td><td>CPYB</td></tr>${Array.from({ length: 120 }, (_, i) => `<tr><td>CROW${i + 1}</td><td>x</td></tr>`).join('')}</tbody></table>`,
    probes: ['CPYA', 'CROW1', 'CROW60', 'CROW120'],
  },
  {
    name: 'E2 a row that cannot split (break-inside: avoid on tr)',
    css: `@page { size: 595pt 842pt; margin: 72pt; }
      .fill { block-size: 660pt; } table { border-collapse: collapse; } td { border: 1px solid #000; }
      tr.cant { break-inside: avoid; }`,
    body: `<div class="fill">filler</div><table><tbody><tr class="cant"><td>${lines('CANT', 6)}</td></tr></tbody></table>`,
    probes: ['CANT1', 'CANT6'],
  },
  {
    name: 'F a right-hand break after one page: a blank page or not',
    css: `@page { size: 595pt 842pt; margin: 72pt; }
      .right { break-before: right; }`,
    body: `<p>RIGHT-before</p><p class="right">RIGHT-after</p>`,
    probes: ['RIGHT-before', 'RIGHT-after'],
  },
  {
    name: 'G margin boxes: a string and the page counters (for the record)',
    css: `@page { size: 595pt 842pt; margin: 72pt; @top-center { content: "MARGINBOX-TOP"; } @bottom-center { content: "PG " counter(page) " OF " counter(pages); } }
      .fill { block-size: 900pt; }`,
    body: `<div class="fill">MB-body</div>`,
    probes: ['MARGINBOX-TOP', 'PG 1 OF 2', 'PG 2 OF 2'],
  },
];

function pageHtml(c: Case): string {
  return `<!doctype html><html><head><meta charset="utf-8"><style>${BASE}${c.css}</style></head><body><div class="docx-editor body"><div class="pm">${c.body}</div></div></body></html>`;
}

/** Whether a page's text holds the probe as a whole word (ROW1 is not in ROW10). */
const holds = (text: string, probe: string): boolean => new RegExp(`(^|[^A-Za-z0-9])${probe.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}($|[^A-Za-z0-9])`).test(text);

/** On which pages each probe lies, and which repeat. */
function placed(c: Case, pages: { size: string; text: string }[]): string {
  const where = c.probes.map((probe) => `${probe}@${pages.findIndex((p) => holds(p.text, probe)) + 1 || '-'}`);
  const repeated = c.probes.map((probe) => [probe, pages.filter((p) => holds(p.text, probe)).length] as const).filter(([, n]) => n > 1).map(([probe, n]) => `${probe}x${n}`);
  return `${pages.length} pages [${pages.map((p) => p.size).join(', ')}] ${where.join(' ')}${repeated.length ? ` repeated: ${repeated.join(' ')}` : ''}`;
}

/** Each page's size in points and text, from the PDF. */
function readPdf(path: string): { size: string; text: string }[] {
  const info = execFileSync('pdfinfo', ['-f', '1', '-l', '999', path], { encoding: 'utf8' });
  const sizes = [...info.matchAll(/^Page\s+\d+ size:\s+([\d.]+) x ([\d.]+) pts/gm)].map((m) => `${Math.round(Number(m[1]))}x${Math.round(Number(m[2]))}`);
  return sizes.map((size, i) => ({ size, text: execFileSync('pdftotext', ['-f', String(i + 1), '-l', String(i + 1), '-layout', path, '-'], { encoding: 'utf8' }) }));
}

async function supports(page: Page): Promise<Record<string, unknown>> {
  return page.evaluate(() => {
    const out: Record<string, unknown> = {};
    for (const [p, v] of [['page', 's1'], ['break-after', 'avoid'], ['break-before', 'right'], ['break-before', 'page'], ['break-inside', 'avoid'], ['orphans', '2'], ['widows', '2'], ['print-color-adjust', 'exact'], ['-webkit-print-color-adjust', 'exact'], ['display', 'table-header-group']] as const) out[`${p}: ${v}`] = CSS.supports(p, v);
    const style = document.createElement('style');
    style.textContent = '@page named { size: 842pt 595pt; @top-center { content: "x"; } }';
    document.head.appendChild(style);
    const rule = style.sheet?.cssRules[0] as (CSSRule & { selectorText?: string; cssRules?: CSSRuleList; style?: CSSStyleDeclaration }) | undefined;
    out['@page parsed'] = rule?.constructor.name;
    out['@page named selector'] = rule?.selectorText;
    out['@page size'] = rule?.style?.getPropertyValue('size');
    out['@page margin rules'] = rule?.cssRules ? rule.cssRules.length : 'none';
    out['CSSMarginRule'] = typeof (window as unknown as { CSSMarginRule?: unknown }).CSSMarginRule;
    return out;
  });
}

test('the print CSS level 1 needs', async ({ page, browserName }) => {
  test.setTimeout(120_000);
  await page.setContent('<!doctype html><html><head></head><body></body></html>');
  report(`supports ${JSON.stringify(await supports(page))}`);
  if (browserName !== 'chromium') {
    report('pagination: through silent printing (the next test); Playwright prints to PDF in Chromium only');
    return;
  }
  const dir = mkdtempSync(join(tmpdir(), 'print-spike-'));
  for (const [i, c] of CASES.entries()) {
    await page.setContent(pageHtml(c));
    await page.emulateMedia({ media: 'print' });
    const path = join(dir, `case-${i}.pdf`);
    await page.pdf({ path, preferCSSPageSize: true, printBackground: true });
    const pages = readPdf(path);
    report(`${c.name}: ${placed(c, pages)}`);
  }
  expect(existsSync(dir)).toBe(true);
});

test('Firefox: every case printed silently to a PDF file', async ({ page, browserName }) => {
  test.skip(browserName !== 'firefox', 'Firefox only');
  test.setTimeout(180_000);
  const target = process.env['PRINT_SPIKE_FIREFOX_PDF'] ?? '';
  for (const c of CASES) {
    if (existsSync(target)) unlinkSync(target);
    await page.setContent(pageHtml(c));
    // window.print() returns at once under the silent preferences; the file appears when the job ends.
    await page.evaluate(() => { window.print(); });
    const deadline = Date.now() + 20_000;
    let size = -1;
    // Written when its size has held for half a second.
    while (Date.now() < deadline) {
      await page.waitForTimeout(250);
      if (!existsSync(target)) continue;
      const now = statSync(target).size;
      if (now > 0 && now === size) break;
      size = now;
    }
    if (!existsSync(target)) { report(`${c.name}: no file written`); continue; }
    report(`${c.name}: ${placed(c, readPdf(target))}`);
  }
});
