// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 6.4, step D2's spike: hidden content next to the caret, before any of the views'
// CSS is written. Deleted text is hidden with one injected rule, as the As accepted view will hide
// it, and each case records what the browser and ProseMirror do, with no rule of the editor's to
// help: the selection read back, typing at the junction, the arrows across it and down from a
// paragraph that ends hidden, scrolling to a position inside hidden text, an IME composition at
// the junction (Chromium: the DevTools protocol), a peer's caret inside hidden text. It reports;
// the findings go in ED-005 section 12. Run with `playwright.spike.config.ts` (Chromium, Firefox).
import { test, expect, type Page } from '@playwright/test';
import { openFixture } from '../helpers.js';

const HIDE = '.docx-editor del.deletion { display: none !important; }';

interface Found { from: number; to: number; text: string; block: number }

type Win = {
  docx4jEditor: {
    state: { doc: PMDoc; selection: { from: number; to: number } };
    view: { dispatch(tr: unknown): void; state: { tr: { setSelection(s: unknown): { scrollIntoView(): unknown } } } };
    caretAt(pos: number, to?: number): void;
    session: { trackingAuthor: string | undefined };
    runAsync(n: string, a: unknown): Promise<boolean>;
  };
};
type PMNode = { firstChild?: PMNode | null; isText: boolean; text?: string; nodeSize: number; marks: { type: { name: string } }[]; type: { name: string }; textContent: string; childCount: number; child(i: number): PMNode; descendants(f: (n: PMNode, p: number) => boolean | void): void; forEach(f: (n: PMNode, p: number, i: number) => void): void; lastChild: PMNode | null };
type PMDoc = PMNode & { content: { size: number } };

function report(line: string): void {
  console.log(`spike [${test.info().project.name}] ${line}`);
}

/** The first deleted text with visible text on both sides in its paragraph (the junction), or the last one in the document. */
async function hiddenJunction(page: Page, which: 'first' | 'last' = 'first'): Promise<Found> {
  return page.evaluate((w) => {
    const doc = (window as unknown as Win).docx4jEditor.state.doc;
    const out: { from: number; to: number; text: string; block: number }[] = [];
    doc.forEach((block, offset, index) => {
      if (block.type.name !== 'paragraph') return;
      const kids: { n: PMNode; p: number }[] = [];
      block.forEach((n, p) => kids.push({ n, p: offset + 1 + p }));
      kids.forEach((k, i) => {
        const del = k.n.isText && k.n.marks.some((m) => m.type.name === 'deletion');
        const visible = (x: { n: PMNode } | undefined): boolean => !!x && x.n.isText && !x.n.marks.some((m) => m.type.name === 'deletion');
        if (del && visible(kids[i - 1]) && visible(kids[i + 1])) out.push({ from: k.p, to: k.p + k.n.nodeSize, text: k.n.text!, block: index });
      });
    });
    return w === 'first' ? out[0]! : out[out.length - 1]!;
  }, which);
}

/** A paragraph's text as accepted (deleted text left out). */
async function accepted(page: Page, block: number): Promise<string> {
  return page.evaluate((b) => {
    let s = '';
    (window as unknown as Win).docx4jEditor.state.doc.child(b).descendants((n) => { if (n.isText && !n.marks.some((m) => m.type.name === 'deletion')) s += n.text; return true; });
    return s;
  }, block);
}

const selection = (page: Page): Promise<{ from: number; to: number }> => page.evaluate(() => { const s = (window as unknown as Win).docx4jEditor.state.selection; return { from: s.from, to: s.to }; });

/** Where the DOM selection is: the element holding its anchor, and whether that element is displayed. */
const domSelection = (page: Page): Promise<string> => page.evaluate(() => {
  const sel = document.getSelection();
  const node = sel?.anchorNode;
  const el = node?.nodeType === 3 ? node.parentElement : (node as Element | null);
  const rect = sel && sel.rangeCount > 0 ? sel.getRangeAt(0).getBoundingClientRect() : undefined;
  return `${el?.tagName.toLowerCase() ?? 'none'}${el?.className ? '.' + String(el.className).split(' ').join('.') : ''} offset ${sel?.anchorOffset} ${el && getComputedStyle(el).display === 'none' ? '(hidden)' : el?.closest('del') ? '(inside del)' : ''}, the caret's rectangle ${rect ? `${Math.round(rect.width)}x${Math.round(rect.height)} at ${Math.round(rect.left)},${Math.round(rect.top)}` : 'none'}`;
});

async function trackAs(page: Page, author: string): Promise<void> {
  await page.evaluate(async (name) => {
    const v = (window as unknown as Win).docx4jEditor;
    v.session.trackingAuthor = name;
    await v.runAsync('document.setChangeTrackingMode', { value: 'TrackAll' });
  }, author);
}

async function openHidden(page: Page, name: string): Promise<void> {
  await openFixture(page, name);
  await page.addStyleTag({ content: HIDE });
  await page.locator('.docx-editor .ProseMirror[aria-label="Document"]').focus();
}

test('1. the caret set at either side of hidden text, and inside it: what the browser reads back, and where typing lands', async ({ page }) => {
  await openHidden(page, 'tracked-changes.docx');
  await trackAs(page, 'Author C');
  const h = await hiddenJunction(page);
  report(`hidden "${h.text}" at ${h.from}-${h.to} in block ${h.block}: "${await accepted(page, h.block)}"`);
  for (const [label, key] of [['before it', 'X'], ['after it', 'Y'], ['inside it', 'Z']] as const) {
    const now = await hiddenJunction(page);
    const pos = label === 'before it' ? now.from : label === 'after it' ? now.to : now.from + 3;
    await page.evaluate((p) => (window as unknown as Win).docx4jEditor.caretAt(p), pos);
    await page.waitForTimeout(150);
    const read = await selection(page);
    report(`caret set ${label} (${pos}): ProseMirror reads ${read.from}; the DOM selection is in ${await domSelection(page)}`);
    await page.keyboard.type(key);
    await page.waitForTimeout(100);
    const after = await hiddenJunction(page).catch(() => undefined);
    report(`  typed "${key}": "${await accepted(page, h.block)}"; the selection at ${(await selection(page)).from}; hidden text now "${after?.text ?? '(none)'}" at ${after?.from}-${after?.to}`);
  }
  expect(await accepted(page, h.block)).toContain('X');
});

test('2. ArrowRight and ArrowLeft across hidden text', async ({ page }) => {
  await openHidden(page, 'tracked-changes.docx');
  const h = await hiddenJunction(page);
  await page.evaluate((p) => (window as unknown as Win).docx4jEditor.caretAt(p - 1), h.from);
  await page.waitForTimeout(100);
  const steps: string[] = [`start ${(await selection(page)).from} (hidden ${h.from}-${h.to})`];
  for (const key of ['ArrowRight', 'ArrowRight', 'ArrowRight', 'ArrowLeft', 'ArrowLeft', 'ArrowLeft']) {
    await page.keyboard.press(key);
    await page.waitForTimeout(60);
    steps.push(`${key === 'ArrowRight' ? '→' : '←'} ${(await selection(page)).from}`);
  }
  report(steps.join(', '));
  // Shift+ArrowRight over the junction: what a selection across hidden text holds.
  await page.evaluate((p) => (window as unknown as Win).docx4jEditor.caretAt(p), h.from);
  await page.keyboard.press('Shift+ArrowRight');
  await page.waitForTimeout(60);
  const s = await selection(page);
  report(`Shift+ArrowRight from ${h.from}: ${s.from}-${s.to}`);
});

test('3. ArrowDown and ArrowUp at a paragraph whose end is hidden, and End', async ({ page }) => {
  await openHidden(page, 'revisions/check22/22-deleted.docx');
  // The first paragraph whose last text is deleted, with visible text before it and a paragraph after it.
  const p = await page.evaluate(() => {
    const doc = (window as unknown as Win).docx4jEditor.state.doc;
    let found: { block: number; lastVisible: number; end: number } | undefined;
    doc.forEach((block, offset, index) => {
      if (found || block.type.name !== 'paragraph' || index + 1 >= doc.childCount || doc.child(index + 1).type.name !== 'paragraph') return;
      let last: PMNode | undefined;
      let lastAt = -1;
      let after = false;
      block.forEach((n, p) => { if (n.isText) { last = n; lastAt = offset + 1 + p; after = false; } else if (last) after = after || n.type.name !== 'opaqueInline'; });
      if (!last || after || !last.marks.some((m) => m.type.name === 'deletion')) return;
      const end = offset + block.nodeSize - 1;
      const start = lastAt;
      let visible = '';
      block.descendants((n) => { if (n.isText && !n.marks.some((m) => m.type.name === 'deletion')) visible += n.text; return true; });
      if (visible.length > 0) found = { block: index, lastVisible: start, end };
    });
    return found;
  });
  if (!p) { report('no paragraph ends in deleted text with one after it'); return; }
  const nextStart = await page.evaluate((b) => { let pos = 0; const doc = (window as unknown as Win).docx4jEditor.state.doc; for (let i = 0; i <= b; i++) pos += doc.child(i).nodeSize; return pos + 1; }, p.block);
  await page.evaluate((pos) => (window as unknown as Win).docx4jEditor.caretAt(pos), p.lastVisible);
  await page.waitForTimeout(100);
  await page.keyboard.press('ArrowDown');
  await page.waitForTimeout(100);
  const down = await selection(page);
  report(`block ${p.block}: caret at its visible end ${p.lastVisible} (hidden to ${p.end}); ArrowDown → ${down.from} (the next paragraph starts at ${nextStart})`);
  await page.keyboard.press('ArrowUp');
  await page.waitForTimeout(100);
  report(`  ArrowUp back → ${(await selection(page)).from}`);
  await page.evaluate((pos) => (window as unknown as Win).docx4jEditor.caretAt(pos), p.lastVisible - 3);
  await page.keyboard.press('End');
  await page.waitForTimeout(100);
  report(`  End from ${p.lastVisible - 3} → ${(await selection(page)).from}`);
  await page.keyboard.press('ArrowRight');
  await page.waitForTimeout(100);
  report(`  then ArrowRight → ${(await selection(page)).from}`);
});

test('4. scrolling a selection inside hidden text into view', async ({ page }) => {
  await page.setViewportSize({ width: 1000, height: 420 });
  await openHidden(page, 'revisions/revisions-word15.docx');
  const h = await hiddenJunction(page, 'last');
  await page.evaluate(() => window.scrollTo(0, 0));
  await page.locator('.page-area').evaluate((el) => { el.scrollTop = 0; });
  // The paragraph through the view: the root's children include widgets, so an index would measure another element.
  const where = async (): Promise<string> => page.evaluate((b) => {
    const v = (window as unknown as Win).docx4jEditor as unknown as { state: { doc: PMDoc }; view: { nodeDOM(pos: number): Node | null } };
    let pos = 0;
    for (let k = 0; k < b; k++) pos += v.state.doc.child(k).nodeSize;
    const p = v.view.nodeDOM(pos) as HTMLElement;
    const r = p.getBoundingClientRect();
    return `paragraph "${p.textContent?.slice(0, 20)}" top ${Math.round(r.top)} bottom ${Math.round(r.bottom)} of a ${window.innerHeight}px window`;
  }, h.block);
  report(`before: ${await where()}`);
  const scrollTo = async (pos: number): Promise<void> => page.evaluate((p) => {
    const v = (window as unknown as Win).docx4jEditor;
    const TextSelection = (v.state.selection as unknown as { constructor: { create(doc: unknown, p: number): unknown } }).constructor;
    v.view.dispatch(v.view.state.tr.setSelection(TextSelection.create(v.state.doc, p)).scrollIntoView());
  }, pos);
  await scrollTo(h.from - 1);
  await page.waitForTimeout(200);
  report(`control, a visible position ${h.from - 1} in the same paragraph: ${await where()}`);
  await page.evaluate(() => window.scrollTo(0, 0));
  await page.locator('.page-area').evaluate((el) => { el.scrollTop = 0; });
  await scrollTo(1);
  await page.waitForTimeout(200);
  report(`back to the start: ${await where()}`);
  let error = '';
  try {
    await page.evaluate((pos) => {
      const v = (window as unknown as Win).docx4jEditor;
      const TextSelection = (v.state.selection as unknown as { constructor: { create(doc: unknown, p: number): unknown } }).constructor;
      v.view.dispatch(v.view.state.tr.setSelection(TextSelection.create(v.state.doc, pos)).scrollIntoView());
    }, h.from + 1);
  } catch (e) { error = (e as Error).message; }
  await page.waitForTimeout(200);
  report(`scrolled to ${h.from + 1}, inside hidden "${h.text}": ${error ? `error ${error}` : 'no error'}; after: ${await where()}`);
});

test('5. an IME composition at the junction, tracking on (Chromium)', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'the DevTools protocol is Chromium\'s');
  await openHidden(page, 'tracked-changes.docx');
  await trackAs(page, 'Author C');
  const h = await hiddenJunction(page);
  const cdp = await page.context().newCDPSession(page);
  for (const [label, pos] of [['before it', h.from], ['after it', h.to]] as const) {
    const now = await hiddenJunction(page);
    const at = label === 'before it' ? now.from : now.to;
    await page.evaluate((p) => (window as unknown as Win).docx4jEditor.caretAt(p), at);
    await page.waitForTimeout(100);
    for (const s of ['か', 'かん', 'かんじ']) {
      await cdp.send('Input.imeSetComposition', { text: s, selectionStart: s.length, selectionEnd: s.length });
      await page.waitForTimeout(40);
    }
    await cdp.send('Input.insertText', { text: '漢字' });
    await page.waitForTimeout(250);
    report(`composed at the junction ${label} (${pos}): "${await accepted(page, h.block)}"`);
  }
  expect(await accepted(page, h.block)).toContain('漢字');
});

test('6. a peer\'s caret inside text this page hides (Chromium, two pages in a room)', async ({ browser, browserName }) => {
  test.skip(browserName !== 'chromium', 'one browser is enough for the room');
  const contexts = await Promise.all([browser.newContext(), browser.newContext()]);
  for (const c of contexts) await c.addInitScript(() => { try { localStorage.setItem('docx4j-editor.relay', '/relay'); } catch { /* ignore */ } });
  const [a, b] = await Promise.all(contexts.map((c) => c.newPage()));
  await openFixture(a!, 'tracked-changes.docx');
  await a!.getByTestId('share').click();
  const link = await a!.getByTestId('share-link').inputValue();
  await a!.keyboard.press('Escape');
  await b!.goto(link);
  await b!.getByTestId('join-accept').click();
  await b!.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await a!.addStyleTag({ content: HIDE });
  const h = await hiddenJunction(a!);
  for (const [label, pos] of [['inside it', h.from + 3], ['at its start', h.from], ['at its end', h.to]] as const) {
    await b!.locator('.docx-editor .ProseMirror[aria-label="Document"]').focus();
    await b!.evaluate((p) => (window as unknown as Win).docx4jEditor.caretAt(p), pos);
    await b!.keyboard.press('Shift+ArrowRight');
    await b!.keyboard.press('Shift+ArrowLeft');
    await expect.poll(() => a!.locator('.peer-caret').count(), { timeout: 10_000 }).toBeGreaterThan(0);
    await a!.waitForTimeout(300);
    const caret = a!.locator('.peer-caret').first();
    const seen = await caret.evaluate((el) => {
      const r = el.getBoundingClientRect();
      return `${el.closest('del') ? 'inside the del' : 'outside any del'}, ${r.width}x${r.height} at ${Math.round(r.left)},${Math.round(r.top)}, visible ${el.checkVisibility()}`;
    });
    report(`B's caret ${label} (${pos}), drawn on A: ${seen}`);
  }
  await Promise.all(contexts.map((c) => c.close()));
});

test('7. Backspace and Delete at either side of hidden text, tracking on', async ({ page }) => {
  for (const [label, side, key] of [['Backspace after it', 'to', 'Backspace'], ['Backspace before it', 'from', 'Backspace'], ['Delete before it', 'from', 'Delete'], ['Delete after it', 'to', 'Delete']] as const) {
    await openHidden(page, 'tracked-changes.docx');
    await trackAs(page, 'Author C');
    const h = await hiddenJunction(page);
    const before = await accepted(page, h.block);
    await page.evaluate((p) => (window as unknown as Win).docx4jEditor.caretAt(p), h[side]);
    await page.waitForTimeout(100);
    await page.keyboard.press(key);
    await page.waitForTimeout(150);
    report(`${label} (${h[side]}, hidden ${h.from}-${h.to}): "${before}" → "${await accepted(page, h.block)}"; selection ${(await selection(page)).from}; the DOM selection in ${await domSelection(page)}`);
  }
});

test('8. a paragraph that starts with hidden text: the caret at its start, and after the hidden run', async ({ page }) => {
  await openHidden(page, 'revisions/check22/22-deleted.docx');
  const p = await page.evaluate(() => {
    const doc = (window as unknown as Win).docx4jEditor.state.doc;
    let found: { block: number; start: number; after: number } | undefined;
    doc.forEach((block, offset, index) => {
      if (found || block.type.name !== 'paragraph') return;
      const first = block.firstChild as PMNode | null;
      if (!first?.isText || !first.marks.some((m) => m.type.name === 'deletion')) return;
      let visible = false;
      block.forEach((n) => { if (n.isText && !n.marks.some((m) => m.type.name === 'deletion')) visible = true; });
      if (visible) found = { block: index, start: offset + 1, after: offset + 1 + first.nodeSize };
    });
    return found;
  });
  if (!p) { report('no paragraph starts with deleted text and has visible text'); return; }
  for (const [label, pos] of [['its start', p.start], ['after the hidden run', p.after]] as const) {
    await page.evaluate((x) => (window as unknown as Win).docx4jEditor.caretAt(x), pos);
    await page.waitForTimeout(120);
    report(`block ${p.block}, the caret at ${label} (${pos}): ProseMirror reads ${(await selection(page)).from}; the DOM selection in ${await domSelection(page)}`);
    await page.keyboard.type('Q');
    await page.waitForTimeout(100);
    report(`  typed "Q": "${await accepted(page, p.block)}"`);
  }
});

test('9. a zero-width widget at the end of each hidden run (side -1): where the caret is drawn', async ({ page }) => {
  for (const name of ['revisions/check22/22-deleted.docx', 'tracked-changes.docx']) {
    await openHidden(page, name);
    const runs = await page.evaluate(() => {
      const v = (window as unknown as Win).docx4jEditor as unknown as { view: { state: { doc: PMDoc; plugins: { getState(s: unknown): unknown }[] }; setProps(p: unknown): void } };
      const state = v.view.state;
      // The Decoration classes, borrowed from a set a plugin already holds (the page does not export them).
      let DS: { create(doc: unknown, d: unknown[]): unknown } | undefined;
      let D: { widget(pos: number, dom: () => HTMLElement, spec: unknown): unknown } | undefined;
      for (const p of state.plugins) {
        const st = p.getState(state) as { set?: { constructor: unknown; find(): { constructor: unknown }[] } } | undefined;
        const found = st?.set?.find?.();
        if (found && found.length > 0) { DS = st!.set!.constructor as typeof DS; D = found[0]!.constructor as typeof D; break; }
      }
      if (!DS || !D) return [];
      const ends: number[] = [];
      state.doc.descendants((n, p) => { if (n.isText && n.marks.some((m) => m.type.name === 'deletion')) ends.push(p + n.nodeSize); return true; });
      const widgets = ends.map((pos) => D!.widget(pos, () => { const s = document.createElement('span'); s.className = 'rev-anchor'; s.textContent = '​'; return s; }, { side: -1, marks: [], key: `anchor-${pos}` }));
      const set = DS.create(state.doc, widgets);
      v.view.setProps({ decorations: () => set });
      return ends;
    });
    report(`${name}: ${runs.length} widgets at the ends of hidden runs`);
    for (const pos of runs.slice(0, 4)) {
      await page.evaluate((x) => (window as unknown as Win).docx4jEditor.caretAt(x), pos);
      await page.waitForTimeout(100);
      report(`  caret at ${pos}: ProseMirror reads ${(await selection(page)).from}; the DOM selection in ${await domSelection(page)}`);
    }
    if (runs[0] !== undefined) {
      await page.evaluate((x) => (window as unknown as Win).docx4jEditor.caretAt(x), runs[0]);
      await page.keyboard.type('W');
      await page.waitForTimeout(100);
      report(`  typed "W" at ${runs[0]}: the selection at ${(await selection(page)).from}, "W" at ${await page.evaluate(() => { let at = -1; (window as unknown as Win).docx4jEditor.state.doc.descendants((n, p) => { if (at < 0 && n.isText && n.text!.includes('W')) at = p + n.text!.indexOf('W'); return at < 0; }); return at; })}`);
    }
  }
});

