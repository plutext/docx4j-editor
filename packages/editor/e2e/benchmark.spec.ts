// SPDX-License-Identifier: GPL-3.0-or-later
//
// The benchmark in the browser (ED-002 section 7.5; ED-001 decision 10): the 255-page document
// opened through the file input, the marks the app leaves (`docx4j:open`, `docx4j:ready`,
// `docx4j:painted`), the footnote panel's time over `footnote.docx`, a find of "um" and typing
// with the find bar open (ED-005 section 3.6), typing with tracking on (section 5.4), and an edited save (ED-015). Printed and written to `benchmark.json` beside
// this package, which CI uploads; never asserted.
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { test, expect, type Page } from '@playwright/test';
import { createCompanion } from '@docx4j/editor-mcp';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { buildRevisionScale } from '../../editor-model/test/fixtures/build/revisions-scale.mjs';
import { fixturePath } from './helpers.js';

const OUT = join(dirname(fileURLToPath(import.meta.url)), '..', 'benchmark.json');

// No trace: a trace takes a snapshot of the whole DOM at each action, about 0.7 s of the page's main thread on the 255-page
// document, which the figures measured with the editor's own work until ED-005 section 12.38 (E4.g step G3) found it.
test.use({ trace: 'off' });

/** The report so far (each test adds its part), or an empty one. */
async function readReport(): Promise<Record<string, unknown>> {
  try { return JSON.parse(await readFile(OUT, 'utf8')) as Record<string, unknown>; } catch { return {}; }
}

/** A document opened from bytes through the file input, to painted. */
async function openBytes(page: Page, name: string, bytes: Uint8Array): Promise<void> {
  await page.goto('/');
  await page.getByTestId('open').setInputFiles({ name, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(bytes) });
  await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor({ timeout: 120_000 });
  await page.waitForFunction(() => performance.getEntriesByName('docx4j:painted').length > 0, undefined, { timeout: 120_000 });
}

interface Figures {
  document: string;
  openToReadyMs: number;
  readyToPaintedMs: number;
  openToPaintedMs: number;
  footnotesMs?: number;
  /** The open path phase by phase (ED-003 section 8.2): the time from the previous mark, in order. */
  breakdownMs?: Record<string, number>;
}

/** The order the open path's marks fall in; each figure is the time since the previous one. */
const OPEN_MARKS = ['docx4j:open', 'docx4j:store', 'docx4j:zip', 'docx4j:unmarshalled', 'docx4j:ids', 'docx4j:resolved', 'docx4j:projected', 'docx4j:loaded', 'docx4j:session', 'docx4j:styles', 'docx4j:ready', 'docx4j:painted'];

interface Typing {
  keystrokes: number;
  dispatchMedianMs: number;
  dispatchP90Ms: number;
  dispatchMaxMs: number;
  /** With a peer typing (ED-004 section 5): the transactions y-prosemirror applied from the room, timed apart from the keystrokes. */
  remote?: { transactions: number; medianMs: number; p90Ms: number; maxMs: number };
  /** With a peer typing: the cursor plugin's redraws after the peer's caret moved, timed apart too. */
  carets?: { transactions: number; medianMs: number; p90Ms: number; maxMs: number };
  /** Local transactions that change no document (the plugins' own after a remote change: measured tab widths, labels). */
  other?: { transactions: number; medianMs: number; p90Ms: number; maxMs: number };
}

/** Thirty keystrokes in a plain paragraph near the middle of the document, the view's dispatch timed from inside the page (ED-003 section 8.2). */
async function typing(page: Page): Promise<Typing> {
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { type: { name: string }; content: { size: number }; nodeSize: number } } }; caretAt(pos: number): void; view: { dispatch: (tr: unknown) => void } } }).docx4jEditor;
    const doc = v.state.doc;
    let pos = 0;
    const target = Math.floor(doc.childCount / 2);
    for (let i = 0; i < doc.childCount; i++) {
      const node = doc.child(i);
      if (i >= target && node.type.name === 'paragraph' && node.content.size > 5) { v.caretAt(pos + 2); break; }
      pos += node.nodeSize;
    }
    const view = v.view;
    const original = view.dispatch.bind(view);
    const times: number[] = [];
    const remote: number[] = [];
    const carets: number[] = [];
    const other: number[] = [];
    const w = window as unknown as { __dispatchTimes: number[]; __remoteTimes: number[]; __caretTimes: number[]; __otherTimes: number[] };
    w.__dispatchTimes = times;
    w.__remoteTimes = remote;
    w.__caretTimes = carets;
    w.__otherTimes = other;
    // A transaction y-prosemirror made from the room carries its sync meta with `isChangeOrigin`;
    // a redraw of the peers' carets carries the cursor plugin's meta; the rest are the keystrokes.
    view.dispatch = (tr: unknown) => {
      const t0 = performance.now();
      original(tr);
      const meta = (k: string): { isChangeOrigin?: boolean; awarenessUpdated?: boolean } | undefined => (tr as { getMeta(k: string): { isChangeOrigin?: boolean; awarenessUpdated?: boolean } | undefined }).getMeta(k);
      const docChanged = (tr as { docChanged: boolean }).docChanged;
      // The presence plugin's redraw carries its own meta (`docx4j-presence$`, view/presence.ts).
      const list = meta('y-sync$')?.isChangeOrigin === true ? remote : meta('docx4j-presence$') !== undefined ? carets : docChanged ? times : other;
      list.push(performance.now() - t0);
    };
  });
  await page.keyboard.type('the quick brown fox jumps over', { delay: 20 });
  const [times, remote, carets, other] = await page.evaluate(() => { const w = window as unknown as { __dispatchTimes: number[]; __remoteTimes: number[]; __caretTimes: number[]; __otherTimes: number[] }; return [w.__dispatchTimes, w.__remoteTimes, w.__caretTimes, w.__otherTimes]; });
  const quantiles = (values: number[]): ((q: number) => number) => {
    const sorted = [...values].sort((a, b) => a - b);
    return (q) => Math.round((sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * q))] ?? 0) * 10) / 10;
  };
  const at = quantiles(times!);
  const out: Typing = { keystrokes: times!.length, dispatchMedianMs: at(0.5), dispatchP90Ms: at(0.9), dispatchMaxMs: at(1) };
  if (remote!.length > 0) { const r = quantiles(remote!); out.remote = { transactions: remote!.length, medianMs: r(0.5), p90Ms: r(0.9), maxMs: r(1) }; }
  if (carets!.length > 0) { const c = quantiles(carets!); out.carets = { transactions: carets!.length, medianMs: c(0.5), p90Ms: c(0.9), maxMs: c(1) }; }
  if (other!.length > 0) { const o = quantiles(other!); out.other = { transactions: other!.length, medianMs: o(0.5), p90Ms: o(0.9), maxMs: o(1) }; }
  return out;
}

/**
 * Find (ED-005 section 3.6): the bar opened over the document, "um" typed, the time of the
 * transaction that set the query on the body (the search of every block and the highlights
 * drawn), and the count the bar shows. The bar stays open for the typing figure that follows. "um" is as common in the
 * anonymised document (3,636 times, case aside) as "the" was in the specification it stands in for (3,772).
 */
async function findCommon(page: Page): Promise<{ query: string; count: string; setQueryMs: number }> {
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { view: { dispatch: (tr: unknown) => void } } }).docx4jEditor.view;
    const original = v.dispatch.bind(v);
    const w = window as unknown as { __findMs?: number };
    v.dispatch = (tr: unknown) => {
      const t0 = performance.now();
      original(tr);
      const meta = (tr as { getMeta(k: string): { finder?: unknown } | undefined }).getMeta('docx4j-find');
      if (meta && meta.finder) w.__findMs = performance.now() - t0;
    };
  });
  await page.keyboard.press('Control+f');
  await page.getByTestId('find-query').fill('um');
  await expect(page.getByTestId('find-count')).toHaveText(/ of /, { timeout: 60_000 });
  const count = (await page.getByTestId('find-count').textContent()) ?? '';
  const setQueryMs = Math.round((await page.evaluate(() => (window as unknown as { __findMs?: number }).__findMs ?? -1)) * 10) / 10;
  return { query: 'um', count, setQueryMs };
}

/**
 * Select by XPath (ED-005 section 3.5): `//w:p[w:pPr/w:pStyle]` over the open document, the export
 * and the marshal included, with the package's engine as the Edit perspective has it (the default
 * one, the browser's native `document.evaluate`, which the XPath tab uses too), then through the
 * tab with its "XPath 3.1" option, FontoXPath.
 */
async function selectByXPath(page: Page): Promise<{ xpath: string; hits: number; nativeMs: number; fontoMs: number }> {
  const xpath = '//w:p[w:pPr/w:pStyle]';
  const native = await page.evaluate(async (x) => {
    const s = (window as unknown as { docx4jEditor: { session: { selectXPath(x: string): Promise<unknown[]> } } }).docx4jEditor.session;
    const t0 = performance.now();
    const hits = await s.selectXPath(x);
    return { hits: hits.length, ms: Math.round(performance.now() - t0) };
  }, xpath);
  await page.getByTestId('perspective-developer').click();
  await page.getByTestId('tab-xpath').click();
  const panel = page.getByTestId('panel-xpath');
  await panel.getByTestId('xpath-31').check();
  await panel.getByTestId('xpath-expression').fill(xpath);
  const t0 = Date.now();
  await panel.getByTestId('xpath-run').click();
  await expect(panel.getByTestId('xpath-count')).toContainText('hits', { timeout: 180_000 });
  const fontoMs = Date.now() - t0;
  await page.getByTestId('perspective-edit').click();
  return { xpath, hits: native.hits, nativeMs: native.ms, fontoMs };
}

async function marks(page: Page): Promise<Record<string, number>> {
  return page.evaluate(() => Object.fromEntries(performance.getEntriesByType('mark').filter((m) => m.name.startsWith('docx4j:')).map((m) => [m.name, m.startTime])));
}

async function openAndMeasure(page: Page, name: string, extraMark?: string, headersAndFooters = true): Promise<Figures> {
  await page.goto('/');
  await page.evaluate((on) => localStorage.setItem('docx4j-editor.headersAndFooters', String(on)), headersAndFooters);
  await page.goto('/');
  await page.getByTestId('open').setInputFiles(fixturePath(name));
  await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor({ timeout: 120_000 });
  await page.waitForFunction(() => performance.getEntriesByName('docx4j:painted').length > 0, undefined, { timeout: 120_000 });
  if (extraMark) await page.waitForFunction((m) => performance.getEntriesByName(m).length > 0, extraMark, { timeout: 60_000 });
  const m = await marks(page);
  const out: Figures = {
    document: name,
    openToReadyMs: Math.round(m['docx4j:ready']! - m['docx4j:open']!),
    readyToPaintedMs: Math.round(m['docx4j:painted']! - m['docx4j:ready']!),
    openToPaintedMs: Math.round(m['docx4j:painted']! - m['docx4j:open']!),
  };
  if (extraMark && m[extraMark] !== undefined) out.footnotesMs = Math.round(m[extraMark] - m['docx4j:ready']!);
  const breakdown: Record<string, number> = {};
  let previous: number | undefined;
  for (const name of OPEN_MARKS) {
    const t = m[name];
    if (t === undefined) continue;
    if (previous !== undefined) breakdown[name.slice('docx4j:'.length)] = Math.round(t - previous);
    previous = t;
  }
  out.breakdownMs = breakdown;
  return out;
}

/** The Review perspective opened (ED-005 section 6.5): from the click to the panel's count, its chunk loaded and its first listing made. */
async function reviewOpen(page: Page): Promise<number> {
  const t0 = await page.evaluate(() => performance.now());
  await page.getByTestId('perspective-review').click();
  await page.getByTestId('review-count').waitFor();
  return Math.round((await page.evaluate(() => performance.now())) - t0);
}

/** Each view of the tracked changes chosen (ED-005 section 6.4): from the class set to two frames later, its style recalculation and layout done. */
async function markupSwitch(page: Page): Promise<Record<string, number>> {
  const out: Record<string, number> = {};
  for (const view of ['accepted', 'original', 'all']) {
    out[view] = await page.evaluate(async (v) => {
      const e = (window as unknown as { docx4jEditor: { markupView: string } }).docx4jEditor;
      const t0 = performance.now();
      e.markupView = v;
      await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
      return Math.round(performance.now() - t0);
    }, view);
  }
  return out;
}

/**
 * The Print command (ED-006 section 3.8): from the button to `print`, which is replaced so no dialog opens. The
 * fonts awaited, the images decoded, the notes laid out, the page rules written.
 */
async function printPrepare(page: Page): Promise<{ toPrintMs: number; notes: number; pageRulesChars: number }> {
  await page.evaluate(() => {
    const w = window as unknown as { printed?: { at: number; notes: number; rules: number }; printStart?: number };
    delete w.printed;
    window.print = () => {
      w.printed = { at: performance.now(), notes: document.querySelectorAll('.print-note').length, rules: document.querySelector('style[data-docx4j-editor="print-pages"]')?.textContent?.length ?? 0 };
      window.dispatchEvent(new Event('afterprint'));
    };
    w.printStart = performance.now();
    (document.querySelector('[data-testid="print"]') as HTMLButtonElement).click();
  });
  await page.waitForFunction(() => (window as unknown as { printed?: unknown }).printed !== undefined, undefined, { timeout: 60_000 });
  return page.evaluate(() => {
    const w = window as unknown as { printed: { at: number; notes: number; rules: number }; printStart: number };
    return { toPrintMs: Math.round(w.printed.at - w.printStart), notes: w.printed.notes, pageRulesChars: w.printed.rules };
  });
}

test('benchmark: the 255-page document and the footnote panel (reported, never gating)', async ({ page }) => {
  test.setTimeout(300_000);
  const large = await openAndMeasure(page, 'large/anon-250.docx');
  const boundaries = await page.locator('.docx-editor .page-boundary').count();
  const typed = await typing(page);
  // Typing with tracking on (ED-005 section 5.4): every keystroke goes through the tracker.
  const tracking = (mode: string): Promise<boolean> => page.evaluate((value) => (window as unknown as { docx4jEditor: { runAsync(n: string, a: unknown): Promise<boolean> } }).docx4jEditor.runAsync('document.setChangeTrackingMode', { value }), mode);
  await tracking('TrackAll');
  const typedTracked = await typing(page);
  // The Review perspective (ED-005 section 6.5): the panel's first listing, typing with tracking on while it is open, each view's switch.
  const reviewPanelOpenMs = await reviewOpen(page);
  const typedTrackedReview = await typing(page);
  const switched = await markupSwitch(page);
  await page.getByTestId('perspective-edit').click();
  await tracking('Off');
  const found = await findCommon(page);
  const typedFinding = await typing(page);
  const selected = await selectByXPath(page);
  const printed = await printPrepare(page);
  const largeBare = await openAndMeasure(page, 'large/anon-250.docx', undefined, false);
  const notes = await openAndMeasure(page, 'footnote.docx', 'docx4j:footnotes');
  const report = {
    when: new Date().toISOString(),
    userAgent: await page.evaluate(() => navigator.userAgent),
    aspirationMs: 2000,
    large: { ...large, pageBoundaries: boundaries },
    typing: typed,
    typingTracked: typedTracked,
    reviewPanelOpenMs,
    typingTrackedReview: typedTrackedReview,
    markupSwitchMs: switched,
    find: found,
    typingWithFindBar: typedFinding,
    selectByXPath: selected,
    print: printed,
    largeWithoutHeadersAndFooters: largeBare,
    footnotes: notes,
  };
  console.log(`benchmark ${large.document}: open→ready ${large.openToReadyMs} ms, ready→painted ${large.readyToPaintedMs} ms, open→painted ${large.openToPaintedMs} ms (${boundaries} page boundaries)`);
  console.log(`benchmark ${large.document} open breakdown: ${Object.entries(large.breakdownMs ?? {}).map(([k, v]) => `${k} ${v}`).join(', ')}`);
  console.log(`benchmark ${large.document} typing: ${typed.keystrokes} keystrokes, dispatch median ${typed.dispatchMedianMs} ms, p90 ${typed.dispatchP90Ms} ms, max ${typed.dispatchMaxMs} ms`);
  console.log(`benchmark ${large.document} typing with tracking on: ${typedTracked.keystrokes} keystrokes, dispatch median ${typedTracked.dispatchMedianMs} ms, p90 ${typedTracked.dispatchP90Ms} ms, max ${typedTracked.dispatchMaxMs} ms`);
  console.log(`benchmark ${large.document} the Review perspective: the panel open ${reviewPanelOpenMs} ms; typing with tracking on, the panel open: dispatch median ${typedTrackedReview.dispatchMedianMs} ms, p90 ${typedTrackedReview.dispatchP90Ms} ms, max ${typedTrackedReview.dispatchMaxMs} ms; the views: ${Object.entries(switched).map(([k, v]) => `${k} ${v} ms`).join(', ')}`);
  console.log(`benchmark ${large.document} find "${found.query}": ${found.count}, the query set in ${found.setQueryMs} ms; typing with the bar open: dispatch median ${typedFinding.dispatchMedianMs} ms, p90 ${typedFinding.dispatchP90Ms} ms, max ${typedFinding.dispatchMaxMs} ms`);
  console.log(`benchmark ${large.document} select ${selected.xpath}: ${selected.hits} hits, the browser's own engine ${selected.nativeMs} ms, FontoXPath (the XPath tab's 3.1 option) ${selected.fontoMs} ms`);
  console.log(`benchmark ${large.document} Print: ${printed.toPrintMs} ms from the button to print (${printed.notes} notes laid out, page rules ${printed.pageRulesChars} characters)`);
  console.log(`benchmark ${large.document} without headers and footers: open→ready ${largeBare.openToReadyMs} ms, ready→painted ${largeBare.readyToPaintedMs} ms, open→painted ${largeBare.openToPaintedMs} ms`);
  console.log(`benchmark ${notes.document}: footnote panel ${notes.footnotesMs ?? '?'} ms after ready`);
  await writeFile(OUT, JSON.stringify(report, null, 2) + '\n');
  expect(large.openToPaintedMs).toBeGreaterThan(0);
});

test('benchmark: typing in the 255-page document while two peers type elsewhere in it (ED-004 sections 5 and 6.3; reported, never gating)', async ({ browser }) => {
  test.setTimeout(300_000);
  const withRelay = async (): Promise<Page> => {
    const context = await browser.newContext();
    await context.addInitScript(() => { try { localStorage.setItem('docx4j-editor.relay', '/relay'); } catch { /* ignore */ } });
    return context.newPage();
  };
  const a = await withRelay();
  const b = await withRelay();
  const c = await withRelay();
  await a.goto('/');
  await a.getByTestId('open').setInputFiles(fixturePath('large/anon-250.docx'));
  await a.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor({ timeout: 120_000 });
  const shareStart = Date.now();
  await a.getByTestId('share').click();
  const link = await a.getByTestId('share-link').inputValue({ timeout: 120_000 });
  const shareMs = Date.now() - shareStart;
  await a.keyboard.press('Escape');
  const joinStart = Date.now();
  await b.goto(link);
  await b.getByTestId('join-accept').click();
  await b.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor({ timeout: 120_000 });
  await expect(b.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 120_000 });
  const joinMs = Date.now() - joinStart;
  await c.goto(link);
  await c.getByTestId('join-accept').click();
  await expect(c.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 120_000 });
  // Two peers type, one near the start and one near the end, while this page types near the middle.
  await b.evaluate(() => (window as unknown as { docx4jEditor: { caretAt(p: number): void } }).docx4jEditor.caretAt(3));
  await c.evaluate(() => { const v = (window as unknown as { docx4jEditor: { state: { doc: { content: { size: number } } }; caretAt(p: number): void } }).docx4jEditor; v.caretAt(v.state.doc.content.size - 3); });
  const peers = Promise.all([
    b.keyboard.type('a colleague typing near the start of the document meanwhile', { delay: 25 }),
    c.keyboard.type('another colleague typing near its end at the same time', { delay: 25 }),
  ]);
  const typed = await typing(a);
  await peers;
  let report: Record<string, unknown> = {};
  try { report = JSON.parse(await readFile(OUT, 'utf8')) as Record<string, unknown>; } catch { /* the first test did not run */ }
  report['typingWithPeer'] = { ...typed, shareMs, joinMs };
  console.log(`benchmark typing with two peers: ${typed.keystrokes} keystrokes, dispatch median ${typed.dispatchMedianMs} ms, p90 ${typed.dispatchP90Ms} ms; ${typed.remote?.transactions ?? 0} remote transactions, median ${typed.remote?.medianMs ?? '?'} ms, p90 ${typed.remote?.p90Ms ?? '?'} ms; ${typed.carets?.transactions ?? 0} caret redraws, median ${typed.carets?.medianMs ?? '?'} ms; ${typed.other?.transactions ?? 0} other, median ${typed.other?.medianMs ?? '?'} ms; share ${shareMs} ms, join ${joinMs} ms`);
  await writeFile(OUT, JSON.stringify(report, null, 2) + '\n');
  expect(typed.keystrokes).toBeGreaterThan(0);
});


// ED-005 section 9.4 item 1 (E4.g step G3): the revision-scale documents (`revisions-scale.mjs`, built here from the 255-page
// document) at 200 and 1,000 revisions, in the Review perspective: the panel's first listing, typing with tracking on with the
// panel open, and each view's switch.
test('benchmark: the revision-scale documents in the Review perspective (ED-005 section 9.4; reported, never gating)', async ({ page }) => {
  test.setTimeout(600_000);
  const scale: Record<string, unknown> = {};
  for (const count of [200, 1000]) {
    await openBytes(page, `revisions-${count}.docx`, await buildRevisionScale(count));
    const reviewPanelOpenMs = await reviewOpen(page);
    const listed = await page.getByTestId('review-count').textContent();
    await page.evaluate(() => (window as unknown as { docx4jEditor: { runAsync(n: string, a: unknown): Promise<boolean> } }).docx4jEditor.runAsync('document.setChangeTrackingMode', { value: 'TrackAll' }));
    const typedTrackedReview = await typing(page);
    const markupSwitchMs = await markupSwitch(page);
    scale[String(count)] = { reviewPanelOpenMs, listed, typingTrackedReview: typedTrackedReview, markupSwitchMs };
    console.log(`benchmark ${count} revisions: the panel open ${reviewPanelOpenMs} ms (${listed}); typing with tracking on, the panel open: dispatch median ${typedTrackedReview.dispatchMedianMs} ms, p90 ${typedTrackedReview.dispatchP90Ms} ms, max ${typedTrackedReview.dispatchMaxMs} ms; the views: ${Object.entries(markupSwitchMs).map(([k, v]) => `${k} ${v} ms`).join(', ')}`);
  }
  const report = await readReport();
  report['revisionScale'] = scale;
  await writeFile(OUT, JSON.stringify(report, null, 2) + '\n');
});

/** The transactions y-prosemirror applies from the room while `during` runs, timed inside the page, and the time until `done` holds. */
async function remoteWhile(page: Page, during: () => Promise<unknown>, done: string): Promise<{ transactions: number; medianMs: number; maxMs: number; totalMs: number; untilShownMs: number }> {
  await page.evaluate(() => {
    const view = (window as unknown as { docx4jEditor: { view: { dispatch: (tr: unknown) => void } } }).docx4jEditor.view;
    const w = window as unknown as { __remote: number[]; __original?: (tr: unknown) => void };
    w.__remote = [];
    const original = w.__original ?? view.dispatch.bind(view);
    w.__original = original;
    view.dispatch = (tr: unknown) => {
      const t0 = performance.now();
      original(tr);
      if ((tr as { getMeta(k: string): { isChangeOrigin?: boolean } | undefined }).getMeta('y-sync$')?.isChangeOrigin === true) w.__remote.push(performance.now() - t0);
    };
  });
  const t0 = Date.now();
  await during();
  await page.waitForFunction((text) => (window as unknown as { docx4jEditor: { state: { doc: { textContent: string } } } }).docx4jEditor.state.doc.textContent.includes(text), done, { timeout: 60_000 });
  const untilShownMs = Date.now() - t0;
  const times = await page.evaluate(() => (window as unknown as { __remote: number[] }).__remote);
  const sorted = [...times].sort((a, b) => a - b);
  const r = (v: number) => Math.round(v * 10) / 10;
  return { transactions: times.length, medianMs: r(sorted[Math.floor(sorted.length / 2)] ?? 0), maxMs: r(sorted[sorted.length - 1] ?? 0), totalMs: r(times.reduce((a, b) => a + b, 0)), untilShownMs };
}

// ED-005 section 9.4 item 3 (E4.g step G3): an agent writing in the 255-page document, the companion in this process and its
// relay in a worker of its own (section 12.35). From the page: an agent's one-paragraph edit and a Replace over two hundred
// paragraphs as the page applies them; typing while the agent makes a call a second; and two pages' typing reaching each
// other while a call runs.
test('benchmark: an agent writing in the 255-page document (ED-005 section 9.4 item 3; reported, never gating)', async ({ browser }) => {
  test.setTimeout(600_000);
  const root = await mkdtemp(join(tmpdir(), 'bench-agent-root-'));
  const data = await mkdtemp(join(tmpdir(), 'bench-agent-rooms-'));
  await writeFile(join(root, 'large.docx'), await readFile(fixturePath('large/anon-250.docx')));
  const companion = createCompanion({ roots: [root], port: 0, app: join(dirname(fileURLToPath(import.meta.url)), '..', 'dist'), data, log: () => undefined });
  const contexts = [await browser.newContext(), await browser.newContext()];
  try {
    await companion.listen();
    const room = await companion.open('large.docx');
    const [a, b] = await Promise.all(contexts.map((c) => c.newPage()));
    const joinStart = Date.now();
    for (const page of [a!, b!]) {
      await page.goto(room.link);
      await page.getByTestId('join-accept').click();
      await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor({ timeout: 120_000 });
      await expect(page.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 120_000 });
    }
    const joinMs = Date.now() - joinStart;
    const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
    await companion.mcp().connect(serverSide);
    const client = new Client({ name: 'benchmark', version: '1' });
    await client.connect(clientSide);
    const call = async (name: string, args: Record<string, unknown>): Promise<{ ms: number; isError: boolean }> => {
      const t0 = Date.now();
      const result = await client.callTool({ name, arguments: { document: room.id, ...args } }) as { isError?: boolean };
      return { ms: Date.now() - t0, isError: result.isError === true };
    };
    // The page's cost of an agent's call: a paragraph near the middle replaced, then "lorem" replaced
    // wherever it is: in 227 paragraphs by a count of their text, as "document" was in 228 of the specification's.
    const blocks = ((await client.callTool({ name: 'outline', arguments: { document: room.id, budget: 100_000 } })) as unknown as { structuredContent: { blocks: { address: string; kind: string; text: string }[] } }).structuredContent.blocks;
    const middle = blocks.filter((x) => x.kind === 'paragraph' && x.text.length > 30)[Math.floor(blocks.length / 4)]!;
    let editCall = { ms: 0, isError: false };
    const edit = await remoteWhile(a!, async () => { editCall = await call('edit', { address: middle.address, text: 'The agent rewrote this paragraph.' }); }, 'The agent rewrote this paragraph.');
    let replaceCall = { ms: 0, isError: false };
    const replace = await remoteWhile(a!, async () => { replaceCall = await call('replace_text', { find: 'lorem', replace: 'record' }); }, 'record');
    // Typing on page A while the agent inserts a paragraph at the end once a second.
    let writing = true;
    const agent = (async () => { let i = 0; while (writing) { await call('insert', { location: 'End', text: `The agent's paragraph ${i++}.` }); await new Promise((r) => setTimeout(r, 1000)); } })();
    const typedWithAgent = await typing(a!);
    writing = false;
    await agent;
    // Two pages' typing reaching each other while a call runs: A types on its own timer; B notes each keystroke's arrival.
    await b!.evaluate(() => {
      const w = window as unknown as { docx4jEditor: { state: { doc: { textContent: string } } }; __seen: Record<string, number> };
      w.__seen = {};
      const timer = setInterval(() => { const t = w.docx4jEditor.state.doc.textContent; for (let i = 0; i < 10; i++) if (w.__seen[i] === undefined && t.includes(`[k${i}]`)) w.__seen[i] = Date.now(); if (Object.keys(w.__seen).length === 10) clearInterval(timer); }, 2);
    });
    await a!.evaluate(() => {
      const v = (window as unknown as { docx4jEditor: { view: { state: { tr: { insertText(t: string, pos: number): unknown } }; dispatch(tr: unknown): void } } }).docx4jEditor.view;
      const w = window as unknown as { __sent: Record<string, number> };
      w.__sent = {};
      let i = 0;
      const timer = setInterval(() => { v.dispatch(v.state.tr.insertText(`[k${i}]`, 1)); w.__sent[i] = Date.now(); if (++i === 10) clearInterval(timer); }, 60);
    });
    const duringCall = await call('replace_text', { find: 'record', replace: 'lorem' });
    await b!.waitForFunction(() => Object.keys((window as unknown as { __seen: Record<string, number> }).__seen).length === 10, undefined, { timeout: 60_000 });
    const sent = await a!.evaluate(() => (window as unknown as { __sent: Record<string, number> }).__sent);
    const seen = await b!.evaluate(() => (window as unknown as { __seen: Record<string, number> }).__seen);
    const latencies = Object.keys(sent).map((k) => seen[k]! - sent[k]!).sort((x, y) => x - y);
    const report = await readReport();
    report['agentWriting'] = {
      joinMs, edit: { callMs: editCall.ms, ...edit }, replaceAll: { callMs: replaceCall.ms, ...replace },
      typingWithAgent: typedWithAgent,
      peerLatencyDuringCall: { callMs: duringCall.ms, medianMs: latencies[5], maxMs: latencies[9] },
    };
    await writeFile(OUT, JSON.stringify(report, null, 2) + '\n');
    console.log(`benchmark agent writing (255 pages): two pages joined in ${joinMs} ms`);
    console.log(`benchmark agent writing: an edit of one paragraph, the call ${editCall.ms} ms, shown on the page after ${edit.untilShownMs} ms, ${edit.transactions} remote transactions, median ${edit.medianMs} ms, max ${edit.maxMs} ms`);
    console.log(`benchmark agent writing: "lorem" replaced everywhere, the call ${replaceCall.ms} ms, shown after ${replace.untilShownMs} ms, ${replace.transactions} remote transactions, max ${replace.maxMs} ms, total ${replace.totalMs} ms`);
    console.log(`benchmark agent writing: typing while the agent writes once a second, dispatch median ${typedWithAgent.dispatchMedianMs} ms, p90 ${typedWithAgent.dispatchP90Ms} ms, max ${typedWithAgent.dispatchMaxMs} ms; ${typedWithAgent.remote?.transactions ?? 0} remote transactions, median ${typedWithAgent.remote?.medianMs ?? '?'} ms, max ${typedWithAgent.remote?.maxMs ?? '?'} ms`);
    console.log(`benchmark agent writing: page A's keystrokes reaching page B while a ${duringCall.ms} ms call ran: median ${latencies[5]} ms, max ${latencies[9]} ms`);
    await client.close();
  } finally {
    for (const c of contexts) await c.close();
    await companion.close();
    await rm(root, { recursive: true, force: true });
    await rm(data, { recursive: true, force: true });
  }
});

// ED-014 P20 (S1.4): in VS Code the store lives in memory and each hot-exit backup encodes it whole, through the page's `backup`
// message (VS Code asks about a second after each change, S0). Over 200 ms on the 255-page document, typed in, the backup
// becomes incremental, as a later step.
test('benchmark: a VS Code backup of the 255-page document, typed in (ED-014 P20; reported, never gating)', async ({ page }) => {
  test.setTimeout(300_000);
  await page.addInitScript(() => {
    const w = window as unknown as Record<string, unknown>;
    const posted: unknown[] = [];
    w['__posted'] = posted;
    w['acquireVsCodeApi'] = () => ({ postMessage: (m: unknown) => { posted.push(m); } });
  });
  await page.goto('/');
  await page.getByTestId('opening').waitFor();
  const bytes = await readFile(fixturePath('large/anon-250.docx'));
  await page.evaluate((b64) => {
    window.dispatchEvent(new MessageEvent('message', { data: { type: 'open', name: 'anon-250.docx', bytes: Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)) } }));
  }, bytes.toString('base64'));
  await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor({ timeout: 120_000 });
  await page.waitForFunction(() => performance.getEntriesByName('docx4j:painted').length > 0, undefined, { timeout: 120_000 });
  const typed = await typing(page);
  // The message is handled as it is dispatched, so the time around the dispatch is the page's whole answer.
  const backup = await page.evaluate(() => {
    const posted = (window as unknown as { __posted: { type: string; update?: Uint8Array }[] }).__posted;
    const times: number[] = [];
    let size = 0;
    for (let id = 1; id <= 11; id++) {
      posted.length = 0;
      const start = performance.now();
      window.dispatchEvent(new MessageEvent('message', { data: { type: 'backup', id } }));
      times.push(performance.now() - start);
      size = posted.find((m) => m.type === 'backup')?.update?.byteLength ?? 0;
    }
    times.sort((a, b) => a - b);
    const round = (ms: number): number => Math.round(ms * 10) / 10;
    return { medianMs: round(times[5]!), maxMs: round(times[10]!), kB: Math.round(size / 1024) };
  });
  const report = await readReport();
  report['backupMs'] = { ...backup, keystrokes: typed.keystrokes, thresholdMs: 200 };
  console.log(`benchmark backupMs (255 pages, ${typed.keystrokes} keystrokes typed): the store encoded whole for VS Code's hot exit, median ${backup.medianMs} ms, max ${backup.maxMs} ms, ${backup.kB} kB (ED-014 P20: incremental over 200 ms)`);
  await writeFile(OUT, JSON.stringify(report, null, 2) + '\n');
  expect(backup.kB).toBeGreaterThan(0);
});

// ED-015 (S1): an edited save of the 255-page document, the export alone and with the self-check the application always runs,
// timed in the page. A keystroke before each, so the main part is rebuilt and marshalled both times. The page takes no input
// while a save runs, so these are also how long a person waits after Save.
test('benchmark: an edited save of the 255-page document (ED-015; reported, never gating)', async ({ page }) => {
  test.setTimeout(300_000);
  await openAndMeasure(page, 'large/anon-250.docx');
  const typeOne = async (key: string): Promise<void> => {
    await page.evaluate(() => {
      const v = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { type: { name: string }; content: { size: number }; nodeSize: number } } }; caretAt(pos: number): void } }).docx4jEditor;
      const doc = v.state.doc;
      let pos = 0;
      for (let i = 0; i < doc.childCount; i++) {
        const node = doc.child(i);
        if (i >= doc.childCount / 2 && node.type.name === 'paragraph' && node.content.size > 5) { v.caretAt(pos + 2); break; }
        pos += node.nodeSize;
      }
    });
    await page.keyboard.type(key);
  };
  const save = (selfCheck: boolean): Promise<{ ms: number; changed: boolean }> => page.evaluate(async (check) => {
    const s = (window as unknown as { docx4jEditor: { session: { save(o: { selfCheck?: boolean }): Promise<{ changed: boolean }> } } }).docx4jEditor.session;
    const t0 = performance.now();
    const r = await s.save(check ? {} : { selfCheck: false });
    return { ms: Math.round(performance.now() - t0), changed: r.changed };
  }, selfCheck);
  await typeOne('x');
  const unchecked = await save(false);
  await typeOne('y');
  const checked = await save(true);
  const report = await readReport();
  report['saveEditedMs'] = unchecked.ms;
  report['saveEditedCheckedMs'] = checked.ms;
  console.log(`benchmark saveEditedMs (255 pages, one keystroke): the export alone ${unchecked.ms} ms; saveEditedCheckedMs, with the self-check, ${checked.ms} ms (ED-015)`);
  await writeFile(OUT, JSON.stringify(report, null, 2) + '\n');
  expect(unchecked.changed && checked.changed).toBe(true);
});
