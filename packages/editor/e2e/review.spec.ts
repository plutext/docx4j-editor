// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 6, E4.d on screen. Step D2: tracked changes drawn in the author's colour
// (underline, strikethrough, a move's green double lines, a formatting change dotted, the change
// bar), and the three views: As accepted hides deletions, a wholly deleted paragraph takes no
// space, a paragraph whose mark is deleted runs on with the next on one line (check 22), and
// Backspace at that join deletes the visible character before it; a selection inside hidden text
// is scrolled to; Original hides insertions and cannot be typed in, and a find hit in inserted text
// switches to All markup; a peer's caret in text this page hides is drawn at its visible edge.
import { test, expect, type Page } from '@playwright/test';
import { openFixture, download, comparePackages, caretInBlock, caretAtBlockEnd } from './helpers.js';
import { ZipPartStore } from '@docx4j/core-ts';

const body = '.docx-editor .ProseMirror[aria-label="Document"]';

type Win = {
  docx4jEditor: {
    state: { doc: PMNode; selection: { from: number; to: number; head: number } };
    view: { dispatch(tr: unknown): void; state: { tr: { setSelection(s: unknown): { scrollIntoView(): unknown } } } };
    caretAt(pos: number, to?: number): void;
    markupView: 'all' | 'accepted' | 'original';
    session: { trackingAuthor: string | undefined };
    runAsync(n: string, a: unknown): Promise<boolean>;
  };
};
type PMNode = { isText: boolean; text?: string; textContent: string; nodeSize: number; marks: { type: { name: string } }[]; type: { name: string }; childCount: number; child(i: number): PMNode; forEach(f: (n: PMNode, p: number, i: number) => void): void; descendants(f: (n: PMNode, p: number) => boolean | void): void };

const setView = (page: Page, view: 'all' | 'accepted' | 'original'): Promise<void> => page.evaluate((v) => { (window as unknown as Win).docx4jEditor.markupView = v; }, view);

async function trackAs(page: Page, author: string): Promise<void> {
  await page.evaluate(async (name) => {
    const v = (window as unknown as Win).docx4jEditor;
    v.session.trackingAuthor = name;
    await v.runAsync('document.setChangeTrackingMode', { value: 'TrackAll' });
  }, author);
}

/** Each top-level paragraph's text as accepted (deleted text left out). */
const acceptedTexts = (page: Page): Promise<string[]> => page.evaluate(() => {
  const out: string[] = [];
  (window as unknown as Win).docx4jEditor.state.doc.forEach((p) => {
    let s = '';
    p.forEach((n) => { if (n.isText && !n.marks.some((m) => m.type.name === 'deletion')) s += n.text; });
    out.push(s);
  });
  return out;
});

/** Where top-level block `i`'s content starts. */
const blockStart = (page: Page, i: number): Promise<number> => page.evaluate((index) => {
  const doc = (window as unknown as Win).docx4jEditor.state.doc;
  let pos = 0;
  for (let k = 0; k < index; k++) pos += doc.child(k).nodeSize;
  return pos + 1;
}, i);

test('All markup: the author\'s colour, underline and strikethrough, a move\'s green double lines, a formatting change dotted, the change bar', async ({ page }) => {
  await openFixture(page, 'revisions/revisions-word15.docx');
  const styles = await page.evaluate((sel) => {
    const root = document.querySelector(sel)!;
    const expected = (author: string): string => {
      let h = 0;
      for (let i = 0; i < author.length; i++) h = (h * 31 + author.charCodeAt(i)) >>> 0;
      const probe = document.createElement('span');
      probe.style.color = `hsl(${h % 360} 65% 42%)`;
      document.body.appendChild(probe);
      const c = getComputedStyle(probe).color;
      probe.remove();
      return c;
    };
    const cs = (q: string): CSSStyleDeclaration => getComputedStyle(root.querySelector(q)!);
    const ins = cs('ins.insertion:not(.move)[data-author="Author A"]');
    const del = cs('del.deletion:not(.move)[data-author="Author B"]');
    const moveTo = cs('ins.insertion.move');
    const moveFrom = cs('del.deletion.move');
    const fc = cs('span.format-change');
    const bar = getComputedStyle(root.querySelector('p.has-change')!, '::before');
    return {
      insColor: ins.color, insExpected: expected('Author A'), insLine: ins.textDecorationLine,
      delColor: del.color, delExpected: expected('Author B'), delLine: del.textDecorationLine,
      moveTo: [moveTo.color, moveTo.textDecorationLine, moveTo.textDecorationStyle], moveFrom: [moveFrom.color, moveFrom.textDecorationLine, moveFrom.textDecorationStyle],
      fc: [fc.textDecorationLine, fc.textDecorationStyle],
      bar: [bar.content, bar.position, bar.borderInlineStartWidth],
    };
  }, body);
  expect(styles.insColor).toBe(styles.insExpected);
  expect(styles.insLine).toContain('underline');
  expect(styles.delColor).toBe(styles.delExpected);
  expect(styles.delLine).toContain('line-through');
  expect(styles.moveTo).toEqual(['rgb(27, 127, 27)', 'underline', 'double']);
  expect(styles.moveFrom).toEqual(['rgb(27, 127, 27)', 'line-through', 'double']);
  expect(styles.fc).toEqual(['underline', 'dotted']);
  expect(styles.bar[0]).toBe('""');
  expect(styles.bar[1]).toBe('absolute');
  // The markers of a move's range between blocks: a rule, no label.
  await expect(page.locator(`${body} .opaque-block.move-marker`).first()).toBeVisible();
  await expect(page.locator(`${body} .opaque-block.move-marker .opaque-label`)).toHaveCount(0);
});

test("a tracked row's cells keep their columns, and its change bar is its first cell's (Jason, 2026-10-17: the bar on the row took a column)", async ({ page }) => {
  await openFixture(page, 'revisions/revisions-word15.docx');
  const rows = await page.evaluate((sel) => {
    const root = document.querySelector(sel)!;
    const widths = (tr: Element): number[] => Array.from(tr.children, (c) => Math.round(c.getBoundingClientRect().width));
    return Array.from(root.querySelectorAll('tr.rev-row-ins, tr.rev-row-del')).map((tr) => {
      const bar = getComputedStyle(tr.children[0]!, '::before');
      return { cells: widths(tr), first: widths(tr.closest('table')!.querySelector('tr')!), rowBefore: getComputedStyle(tr, '::before').content, bar: [bar.content, bar.position] };
    });
  }, body);
  expect(rows.length).toBeGreaterThan(0);
  for (const r of rows) {
    expect(r.cells).toEqual(r.first);
    expect(r.rowBefore).toBe('none');
    expect(r.bar).toEqual(['""', 'absolute']);
  }
});

test('As accepted (check 22): deletions hidden, a wholly deleted paragraph takes no space, a deleted mark\'s paragraph runs on with the next, and Backspace at the join deletes the visible character before it', async ({ page }) => {
  // "One." / "[Two.]¶" / "Three." / "Four.¶" / "Five [six.]¶" / "[Seven ]eight." / "Nine.", the marks marked ¶ deleted.
  await openFixture(page, 'revisions/check22/22-deleted.docx');
  const paragraphs = page.locator(`${body} > p`);
  await expect(paragraphs.nth(1)).toBeVisible();
  await setView(page, 'accepted');
  await expect(page.locator(body)).toHaveClass(/markup-accepted/);
  await expect(page.locator(`${body} del.deletion`).first()).toBeHidden();
  await expect(paragraphs.nth(1)).toBeHidden();
  const tops = await paragraphs.evaluateAll((ps) => ps.map((p) => Math.round(p.getBoundingClientRect().top)));
  // "Four.", "Five " and "eight." on one line; "Three." and "Nine." on lines of their own.
  expect(Math.abs(tops[3]! - tops[4]!)).toBeLessThanOrEqual(2);
  expect(Math.abs(tops[4]! - tops[5]!)).toBeLessThanOrEqual(2);
  expect(tops[2]!).toBeLessThan(tops[3]!);
  expect(tops[6]!).toBeGreaterThan(tops[5]!);
  await expect(page.locator(`${body} .rev-anchor`).first()).toBeAttached();
  // Tracking on: the caret at "eight"'s start, Backspace.
  await trackAs(page, 'Author C');
  const eight = (await blockStart(page, 5)) + 'Seven '.length;
  await page.locator(body).focus();
  await page.evaluate((p) => (window as unknown as Win).docx4jEditor.caretAt(p), eight);
  await page.keyboard.press('Backspace');
  await expect.poll(() => acceptedTexts(page)).toEqual(['One.', '', 'Three.', 'Four.', 'Five', 'eight.', 'Nine.', '']);
  // Typing there goes where the caret is drawn.
  await page.keyboard.type('X');
  await expect.poll(async () => (await acceptedTexts(page)).slice(3, 6).join('|')).toMatch(/^Four\.\|FiveX\|eight\.$|^Four\.\|Five\|Xeight\.$/);
  // All markup again: the deleted text and the pilcrows are back.
  await setView(page, 'all');
  await expect(page.locator(`${body} del.deletion`).first()).toBeVisible();
  await expect(paragraphs.nth(1)).toBeVisible();
});

test('As accepted: a selection put inside hidden text is scrolled to, since the caret moves to where it is drawn', async ({ page }) => {
  await page.setViewportSize({ width: 1000, height: 420 });
  await openFixture(page, 'revisions/revisions-word15.docx');
  await setView(page, 'accepted');
  // The last deleted text with visible text on both sides, and its paragraph's index.
  const h = await page.evaluate(() => {
    const doc = (window as unknown as Win).docx4jEditor.state.doc;
    let found: { pos: number; block: number } | undefined;
    doc.forEach((block, offset, index) => {
      if (block.type.name !== 'paragraph') return;
      const kids: { n: PMNode; p: number }[] = [];
      block.forEach((n, p) => kids.push({ n, p: offset + 1 + p }));
      const del = (k?: { n: PMNode }): boolean => !!k && k.n.isText && k.n.marks.some((m) => m.type.name === 'deletion');
      const vis = (k?: { n: PMNode }): boolean => !!k && k.n.isText && !del(k);
      kids.forEach((k, i) => { if (del(k) && vis(kids[i - 1]) && vis(kids[i + 1])) found = { pos: k.p + 1, block: index }; });
    });
    return found!;
  });
  await page.evaluate(() => window.scrollTo(0, 0));
  await page.locator('.page-area').evaluate((el) => { el.scrollTop = 0; });
  await page.evaluate((pos) => {
    const v = (window as unknown as Win).docx4jEditor;
    const TextSelection = (v.state.selection as unknown as { constructor: { create(doc: unknown, p: number): unknown } }).constructor;
    v.view.dispatch(v.view.state.tr.setSelection(TextSelection.create(v.state.doc, pos)).scrollIntoView());
  }, h.pos);
  // The paragraph through the view (the root's children include widgets, so an index would miss it).
  await expect.poll(() => page.evaluate((b) => {
    const v = (window as unknown as Win).docx4jEditor as unknown as { state: { doc: PMNode }; view: { nodeDOM(pos: number): Node | null } };
    let pos = 0;
    for (let k = 0; k < b; k++) pos += v.state.doc.child(k).nodeSize;
    const r = (v.view.nodeDOM(pos) as HTMLElement).getBoundingClientRect();
    // Scrolled to the nearest edge: the bottom one, to a fraction of a pixel.
    return r.top >= 0 && r.bottom <= window.innerHeight + 1;
  }, h.block)).toBe(true);
  expect(await page.evaluate(() => (window as unknown as Win).docx4jEditor.state.selection.head)).toBe(h.pos - 1);
});

test('Original: insertions hidden, typing does nothing; a find hit in inserted text switches to All markup and says so', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const before = await acceptedTexts(page);
  await setView(page, 'original');
  await expect(page.locator(body)).toHaveClass(/markup-original/);
  await expect(page.locator(body)).toHaveAttribute('contenteditable', 'false');
  await expect(page.locator(`${body} ins.insertion`).first()).toBeHidden();
  await expect(page.locator(`${body} del.deletion`).first()).toBeVisible();
  await page.locator(body).click();
  await page.keyboard.type('Q');
  expect(await acceptedTexts(page)).toEqual(before);
  // "An insertion" is inserted text, hidden here: finding it shows it.
  await page.keyboard.press('Control+f');
  await expect(page.getByTestId('find-bar')).toBeVisible();
  await page.getByTestId('find-query').fill('An insertion');
  await expect(page.locator(body)).not.toHaveClass(/markup-original/);
  await expect(page.getByTestId('status')).toContainText('All markup');
});

test('a peer\'s caret inside text this page hides is drawn at its visible edge (two pages in a room)', async ({ browser }) => {
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
  await setView(a!, 'accepted');
  // B's caret inside the deleted " A deletion", which A hides.
  const inside = await b!.evaluate(() => {
    let pos = -1;
    (window as unknown as Win).docx4jEditor.state.doc.descendants((n, p) => { if (pos < 0 && n.isText && n.marks.some((m) => m.type.name === 'deletion')) pos = p + 3; return pos < 0; });
    return pos;
  });
  await b!.locator(body).focus();
  await b!.evaluate((p) => (window as unknown as Win).docx4jEditor.caretAt(p), inside);
  const caret = a!.locator('.peer-caret').first();
  await expect(caret).toBeAttached({ timeout: 10_000 });
  await expect.poll(() => caret.evaluate((el) => !el.closest('del') && el.checkVisibility()), { timeout: 10_000 }).toBe(true);
  await Promise.all(contexts.map((c) => c.close()));
});

// Step D3: the review panel, in the Review perspective.

/** The panel's entries as the page lists them: kind, type, author and the text shown. */
const panelEntries = (page: Page): Promise<string[]> => page.locator('[data-testid="review-entry"]').evaluateAll((els) => els.map((e) => {
  const el = e as HTMLElement;
  return `${el.dataset['kind']}|${el.dataset['type']}|${el.dataset['author']}|${el.querySelector('.snippet')?.textContent ?? ''}`;
}));

async function openReview(page: Page, name: string): Promise<void> {
  await openFixture(page, name);
  await page.getByTestId('perspective-review').click();
  await expect(page.getByTestId('panel-review')).toBeVisible();
  await expect(page.locator('[data-testid="review-entry"]').first()).toBeVisible();
}

test('the review panel: next and previous, by button and by Alt+PageDown and Alt+PageUp, visit every entry once, a move once with both halves marked', async ({ page }) => {
  await openReview(page, 'revisions/revisions-word15.docx');
  const entries = await panelEntries(page);
  await page.locator(body).focus();
  await page.evaluate(() => (window as unknown as Win).docx4jEditor.caretAt(1));
  const current = (): Promise<string> => page.locator('[data-testid="review-entry"][aria-current="true"]').evaluate((el) => {
    const e = el as HTMLElement;
    return `${e.dataset['kind']}|${e.dataset['type']}|${e.dataset['author']}|${e.querySelector('.snippet')?.textContent ?? ''}`;
  });
  const seen: string[] = [];
  let moves = 0;
  for (let i = 0; i < entries.length; i++) {
    if (i % 2 === 0) await page.getByTestId('review-next').click(); else await page.keyboard.press('Alt+PageDown');
    seen.push(await current());
    if ((await page.locator('[data-testid="review-entry"][aria-current="true"] .type').textContent()) === 'Moved') {
      // A move: both halves drawn as the current change, the selection on the destination.
      const marked = await page.locator(`${body} .current-change`).evaluateAll((els) => new Set(els.map((e) => e.closest('p'))).size);
      expect(marked).toBe(2);
      const inDestination = await page.evaluate(() => {
        const v = (window as unknown as Win).docx4jEditor as unknown as { state: { doc: { nodesBetween(a: number, b: number, f: (n: PMNode) => void): void }; selection: { from: number; to: number } } };
        const { from, to } = v.state.selection;
        let to_ = false;
        let from_ = false;
        v.state.doc.nodesBetween(from, to, (n) => {
          if (!n.isText) return;
          for (const m of n.marks as unknown as { type: { name: string }; attrs: { move: string | null } }[]) {
            if (m.type.name === 'insertion' && m.attrs.move) to_ = true;
            if (m.type.name === 'deletion' && m.attrs.move) from_ = true;
          }
        });
        return to_ && !from_;
      });
      expect(inDestination).toBe(true);
      moves++;
    }
  }
  expect(seen).toEqual(entries);
  expect(moves).toBe(1);
  // Round to the first, and back to the last.
  await page.getByTestId('review-next').click();
  expect(await current()).toBe(entries[0]);
  await expect(page.getByTestId('status')).toContainText('first change');
  await page.keyboard.press('Alt+PageUp');
  expect(await current()).toBe(entries[entries.length - 1]);
});

test('the review panel: Accept and Reject each resolve their own entry and no other, for each kind; the save passes its self-check', async ({ page }) => {
  const kinds = ['Added', 'Deleted', 'Formatted'];
  for (const action of ['accept', 'reject'] as const) {
    for (const kind of kinds) {
      await openReview(page, 'revisions/revisions-word15.docx');
      const before = await panelEntries(page);
      const index = before.findIndex((e) => e.startsWith(`change|${kind}|`));
      const entry = page.locator('[data-testid="review-entry"]').nth(index);
      await entry.getByTestId(`review-${action}`).click();
      await expect.poll(async () => (await panelEntries(page)).length).toBe(before.length - 1);
      const after = await panelEntries(page);
      expect(after, `${action} ${kind}`).toEqual(before.filter((_, i) => i !== index));
      await expect(page.getByTestId('status')).toContainText(action === 'accept' ? 'Accepted 1 change' : 'Rejected 1 change');
    }
  }
  // A move: one entry, and either action resolves both halves.
  await openReview(page, 'revisions/revisions-word15.docx');
  const before = await panelEntries(page);
  const moved = page.locator('[data-testid="review-entry"]').filter({ has: page.locator('.type', { hasText: 'Moved' }) }).first();
  await moved.getByTestId('review-accept').click();
  await expect.poll(async () => (await panelEntries(page)).length).toBe(before.length - 1);
  const saved = await page.evaluate(async () => {
    const r = await (window as unknown as { docx4jEditor: { session: { save(): Promise<{ selfCheck?: { equal: boolean } }> } } }).docx4jEditor.session.save();
    return r.selfCheck?.equal;
  });
  expect(saved).toBe(true);
});

test("the current change's tint goes with its accept or reject (Jason, 2026-10-17: accepted text stayed blue until another entry was selected)", async ({ page }) => {
  for (const action of ['accept', 'reject'] as const) {
    await openReview(page, 'revisions/revisions-word15.docx');
    const entry = page.locator('[data-testid="review-entry"][data-type="Added"]').first();
    await entry.locator('.review-entry-main').click();
    await expect(page.locator(`${body} .current-change`).first()).toBeVisible();
    await entry.getByTestId(`review-${action}`).click();
    await expect(page.getByTestId('status')).toContainText(action === 'accept' ? 'Accepted 1 change' : 'Rejected 1 change');
    await expect(page.locator(`${body} .current-change`)).toHaveCount(0);
  }
});

test('the review panel: Accept all says how many it resolved, the two shown in Word only among them (the engine\'s, ED-005 section 12.26); the stories line; an Original says it is read-only', async ({ page }) => {
  await openReview(page, 'revisions/revisions-word15.docx');
  await expect(page.getByTestId('review-count')).toContainText('shown in Word only');
  await expect(page.getByTestId('review-stories')).toBeVisible();
  await page.getByTestId('review-accept-all').click();
  await expect(page.getByTestId('status')).toHaveText(/^Accepted \d+ changes$/);
  await expect(page.locator('[data-testid="review-entry"][data-kind="change"]')).toHaveCount(0);
  await expect(page.locator('[data-testid="review-entry"][data-kind="preserved"]')).toHaveCount(0);
  await page.getByTestId('review-view-original').check();
  await expect(page.locator(body)).toHaveClass(/markup-original/);
  await expect(page.getByTestId('status')).toContainText('read-only');
  await page.getByTestId('review-view-all').check();
  await expect(page.locator(body)).not.toHaveClass(/markup-original/);
});

// Step D4: the tracking switch, its name dialog and Ctrl+Shift+E; the Review perspective offered by the banner; the tracking hint.

test('the tracking switch: with no name set a name is asked for first; Ctrl+Shift+E turns it off and on; the saved file says so', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const toggle = page.getByTestId('tracking-switch');
  await expect(toggle).toHaveAttribute('aria-pressed', 'false');
  await toggle.click();
  const name = page.getByTestId('tracking-name');
  await expect(name).toBeVisible();
  await expect(name).toHaveValue(/Guest/);
  await name.fill('Reviewer One');
  await name.press('Enter');
  await expect(toggle).toHaveAttribute('aria-pressed', 'true');
  await expect(page.getByTestId('status')).toContainText('tracked as Reviewer One');
  await expect(page.getByTestId('option-name')).toHaveValue('Reviewer One');
  // Typing now is an insertion under the name.
  await page.locator(body).focus();
  await page.evaluate(() => (window as unknown as Win).docx4jEditor.caretAt(2));
  await page.keyboard.type('Z');
  await expect(page.locator(`${body} ins.insertion[data-author="Reviewer One"]`)).toHaveCount(1);
  // Ctrl+Shift+E: off, then on again, the name kept, no dialog.
  await page.keyboard.press('Control+Shift+E');
  await expect(toggle).toHaveAttribute('aria-pressed', 'false');
  await page.keyboard.press('Control+Shift+E');
  await expect(toggle).toHaveAttribute('aria-pressed', 'true');
  await expect(page.getByTestId('tracking-name')).toHaveCount(0);
  const settings = new TextDecoder().decode(new ZipPartStore(await download(page)).loadSync('word/settings.xml'));
  expect(settings).toMatch(/<w:trackRevisions\b/);
  // Cancel leaves tracking off.
  await page.keyboard.press('Control+Shift+E');
  await expect(toggle).toHaveAttribute('aria-pressed', 'false');
});

test("Enter with tracking off by a mark Author A inserted: the new mark goes at the caret, plain and drawn so, before A's, at the paragraph's end and in its middle (Jason's findings, ED-005 sections 12.16 and 12.17; core-ts check 29c and 29e)", async ({ page }) => {
  await openFixture(page, 'revisions/check29/29-input.docx');
  await page.evaluate(async () => {
    const v = (window as unknown as Win).docx4jEditor;
    v.session.trackingAuthor = 'Author B';
    await v.runAsync('document.setChangeTrackingMode', { value: 'Off' });
  });
  const indexOf = (start: string): Promise<number> => page.evaluate((text) => {
    let found = -1;
    (window as unknown as Win).docx4jEditor.state.doc.forEach((p, _pos, index) => { if (found < 0 && p.textContent.startsWith(text)) found = index; });
    return found;
  }, start);
  /** Blocks `i` and `i + 1`: text, whose mark's insertion, and whether the pilcrow is drawn inserted. */
  const marks = (i: number): Promise<{ text: string; author: string | null; drawn: boolean }[]> => page.evaluate((index) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: PMNode }; view: { nodeDOM(pos: number): Node | null } } }).docx4jEditor;
    const doc = v.state.doc;
    let pos = 0;
    for (let k = 0; k < index; k++) pos += doc.child(k).nodeSize;
    const out: { text: string; author: string | null; drawn: boolean }[] = [];
    for (const k of [index, index + 1]) {
      const n = doc.child(k) as PMNode & { attrs: { markRevision: { insertion?: { author: string } } | null } };
      out.push({ text: n.textContent, author: n.attrs.markRevision?.insertion?.author ?? null, drawn: (v.view.nodeDOM(pos) as HTMLElement).classList.contains('rev-mark-ins') });
      pos += n.nodeSize;
    }
    return out;
  }, i);
  await page.locator(body).focus();
  // 29c: the caret just before A's inserted mark, at the paragraph's end; Enter, then "New".
  const c = await indexOf('Cases c and d');
  const cText = (await marks(c))[0]!.text;
  await caretAtBlockEnd(page, c);
  await page.keyboard.press('Enter');
  await page.keyboard.type('New');
  expect(await marks(c)).toEqual([{ text: cText, author: null, drawn: false }, { text: 'New', author: 'Author A', drawn: true }]);
  // 29e: the middle of A's other inserted-mark paragraph, after "Case e: before the split,".
  const e = await indexOf('Case e: before the split');
  await caretInBlock(page, e, 'Case e: before the split,'.length);
  await page.keyboard.press('Enter');
  const split = await marks(e);
  expect(split.map(({ author, drawn }) => ({ author, drawn }))).toEqual([{ author: null, drawn: false }, { author: 'Author A', drawn: true }]);
  expect(split[0]!.text).toBe('Case e: before the split,');
});

test('the tracking switch in a room: a peer turning tracking on shows on the other page, and says so', async ({ browser }) => {
  const contexts = await Promise.all([browser.newContext(), browser.newContext()]);
  for (const c of contexts) await c.addInitScript(() => { try { localStorage.setItem('docx4j-editor.relay', '/relay'); localStorage.setItem('docx4j-editor.name', 'Peer'); } catch { /* ignore */ } });
  const [a, b] = await Promise.all(contexts.map((c) => c.newPage()));
  await openFixture(a!, 'tracked-changes.docx');
  await a!.getByTestId('share').click();
  const link = await a!.getByTestId('share-link').inputValue();
  await a!.keyboard.press('Escape');
  await b!.goto(link);
  await b!.getByTestId('join-accept').click();
  await b!.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await expect(b!.getByTestId('tracking-switch')).toHaveAttribute('aria-pressed', 'false');
  await a!.getByTestId('tracking-switch').click();
  await expect(a!.getByTestId('tracking-switch')).toHaveAttribute('aria-pressed', 'true');
  await expect(b!.getByTestId('tracking-switch')).toHaveAttribute('aria-pressed', 'true', { timeout: 10_000 });
  await expect(b!.getByTestId('status')).toContainText('turned on in this shared document');
  // The page that turned it on says it as its own, not as a peer's.
  await expect(a!.getByTestId('status')).not.toContainText('shared document');
  await Promise.all(contexts.map((c) => c.close()));
});

// ED-005 section 12.53: the state is drawn, not only announced. The tests above read `aria-pressed`, which the switch
// always set while its look never changed (Jason's report): here the computed look itself, and the status bar's indicator.
test('tracking\'s state is visible: the switch drawn pressed with a red dot, the status bar saying it; the indicator toggles too', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const toggle = page.getByTestId('tracking-switch');
  const indicator = page.getByTestId('tracking-indicator');
  const look = (el: typeof toggle) => el.evaluate((e) => ({ background: getComputedStyle(e).backgroundColor, dot: getComputedStyle(e, '::before').backgroundColor }));
  await expect(indicator).toHaveText('Track changes: Off');
  const off = await look(toggle);
  // The indicator turns it on, by way of the name, as the switch does.
  await indicator.click();
  const name = page.getByTestId('tracking-name');
  await name.fill('Reviewer One');
  await name.press('Enter');
  await expect(toggle).toHaveAttribute('aria-pressed', 'true');
  await expect(indicator).toHaveText('Track changes: On');
  const on = await look(toggle);
  expect(on.background).not.toBe(off.background);
  expect(on.dot).toBe('rgb(192, 0, 0)');
  expect(off.dot).not.toBe(on.dot);
  expect((await look(indicator)).dot).toBe('rgb(192, 0, 0)');
  await indicator.click();
  await expect(toggle).toHaveAttribute('aria-pressed', 'false');
  await expect(indicator).toHaveText('Track changes: Off');
  expect((await look(toggle)).background).toBe(off.background);
});

test('the banner offers the Review perspective for a document with tracked changes, and not once in it; comments are counted', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const offer = page.getByTestId('review-offer');
  await expect(offer).toHaveText('Open the Review perspective');
  await offer.click();
  await expect(page.getByTestId('perspective-review')).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByTestId('panel-review')).toBeVisible();
  await expect(offer).toHaveCount(0);
  await openFixture(page, 'comments-two.docx');
  await expect(page.getByTestId('banner')).toContainText('2 unresolved comments');
});

// Jason, 2026-11-28: the banner had counted the file as opened, and stood unchanged through every change accepted and
// comment resolved. It follows the document; once nothing it reported is left it goes, and stays gone.
test('the banner follows the document: changes accepted, comments resolved; gone once it has nothing to say, for good', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const banner = page.getByTestId('banner');
  await expect(banner).toContainText('2 tracked changes');
  await page.getByTestId('perspective-review').click();
  await page.getByTestId('review-accept-all').click();
  await expect(page.getByTestId('review-count')).toHaveText('No tracked changes in the body.');
  // The elements shown read-only are the file's as it opened.
  await expect(banner).not.toContainText('tracked');
  await expect(banner).toContainText('shown read-only');

  await openFixture(page, 'comments-two.docx');
  await expect(banner).toContainText('2 unresolved comments');
  await page.locator('[data-comment-id="0"]').getByTestId('comment-resolve').click();
  await expect(banner).toContainText('1 unresolved comment.');
  await page.locator('[data-comment-id="1"]').getByTestId('comment-resolve').click();
  await expect(banner).not.toContainText('comment');
  await page.locator('[data-comment-id="1"] summary').click();
  await page.locator('[data-comment-id="1"]').getByTestId('comment-reopen').click();
  await expect(banner).toContainText('1 unresolved comment.');

  // Two tracked changes and nothing read-only: accepted, the banner goes, and a change typed after does not bring it back.
  await openFixture(page, 'theme-colour-run.docx');
  await expect(banner).toContainText('2 tracked changes');
  await page.getByTestId('perspective-review').click();
  await page.getByTestId('review-accept-all').click();
  await expect(banner).toHaveCount(0);
  await page.getByTestId('tracking-switch').click();
  const name = page.getByTestId('tracking-name');
  if (await name.isVisible()) { await name.fill('Reviewer One'); await name.press('Enter'); }
  await expect(page.getByTestId('tracking-switch')).toHaveAttribute('aria-pressed', 'true');
  await page.locator(body).focus();
  await page.evaluate(() => (window as unknown as Win).docx4jEditor.caretAt(2));
  await page.keyboard.type('Z');
  await expect(page.getByTestId('review-count')).toHaveText('1 change by 1 person');
  await expect(banner).toHaveCount(0);
});

test('the tracking hint: shown once in Review with tracking off, never switching by itself; "Track my changes" goes through the name dialog', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await page.getByTestId('perspective-review').click();
  const hint = page.getByTestId('review-hint');
  await expect(hint).toBeVisible();
  await expect(page.getByTestId('tracking-switch')).toHaveAttribute('aria-pressed', 'false');
  await page.getByTestId('review-hint-dismiss').click();
  await expect(hint).toHaveCount(0);
  // Once answered, not again in this browser.
  await openFixture(page, 'tracked-changes.docx');
  await expect(page.getByTestId('panel-review')).toBeVisible();
  await expect(hint).toHaveCount(0);
  await page.evaluate(() => { try { localStorage.removeItem('docx4j-editor.review.trackingHint'); } catch { /* ignore */ } });
  await openFixture(page, 'tracked-changes.docx');
  await page.getByTestId('review-hint-track').click();
  await page.getByTestId('tracking-name').fill('Hinted');
  await page.getByTestId('tracking-name').press('Enter');
  await expect(page.getByTestId('tracking-switch')).toHaveAttribute('aria-pressed', 'true');
  await expect(hint).toHaveCount(0);
});

// Step D5: the Word-made fixtures in the Review perspective end to end, and in a room.

/** The body's tracked changes as the model lists them in a saved file (a move's halves as one, as the panel shows them), and its revisions shown in Word only. */
async function modelCounts(bytes: Uint8Array): Promise<{ changes: number; preserved: number }> {
  const { EditorDocument, trackedChangesOf, preservedChangesOf, preservedGroupsOf } = await import('@docx4j/editor-model');
  const session = await EditorDocument.open(bytes);
  const changes = trackedChangesOf(session.doc);
  const moves = new Set(changes.filter((c) => c.move !== undefined).map((c) => c.move));
  const halves = changes.filter((c) => c.move !== undefined).length;
  return { changes: changes.length - halves + moves.size, preserved: preservedGroupsOf(session.doc, preservedChangesOf(session.doc, (id) => session.ctx.ids.get(id)?.value)).length };
}

const panelCounts = async (page: Page): Promise<{ changes: number; preserved: number }> => ({
  changes: await page.locator('[data-testid="review-entry"][data-kind="change"]').count(),
  preserved: await page.locator('[data-testid="review-entry"][data-kind="preserved"]').count(),
});

for (const name of ['revisions-word2010.docx', 'revisions-word15.docx']) {
  test(`${name} in Review: listed as the model lists it, each view drawn, Accept all and Reject all saved and reopened with the counts the model gives`, async ({ page }) => {
    await openReview(page, `revisions/${name}`);
    const original = await (await import('./helpers.js')).fixture(`revisions/${name}`);
    expect(await panelCounts(page)).toEqual(await modelCounts(original));
    // Each view: As accepted draws no deletion, Original no insertion; All markup both.
    for (const view of ['accepted', 'original', 'all'] as const) {
      await page.getByTestId(`review-view-${view}`).check();
      const drawn = await page.evaluate(() => {
        const visible = (sel: string): number => [...document.querySelectorAll(`.docx-editor .ProseMirror[aria-label="Document"] ${sel}`)].filter((e) => (e as HTMLElement).checkVisibility()).length;
        return { del: visible('del.deletion'), ins: visible('ins.insertion') };
      });
      if (view === 'accepted') expect(drawn.del).toBe(0);
      if (view === 'original') expect(drawn.ins).toBe(0);
      if (view === 'all') { expect(drawn.del).toBeGreaterThan(0); expect(drawn.ins).toBeGreaterThan(0); }
    }
    for (const action of ['accept', 'reject'] as const) {
      await openReview(page, `revisions/${name}`);
      await page.getByTestId(`review-${action}-all`).click();
      await expect(page.locator('[data-testid="review-entry"][data-kind="change"]')).toHaveCount(0);
      const saved = await download(page);
      expect(await modelCounts(saved), `${action} all, saved`).toEqual({ changes: 0, preserved: 0 });
      // Reopened, the panel says the same.
      await page.getByTestId('open').setInputFiles({ name, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(saved) });
      await expect(page.getByTestId('panel-review')).toBeVisible();
      await expect.poll(() => panelCounts(page)).toEqual({ changes: 0, preserved: 0 });
    }
  });
}

test('in a room: one page accepts a change the other made, both lists agree, both saves are the same', async ({ browser }) => {
  const contexts = await Promise.all([browser.newContext(), browser.newContext()]);
  for (const c of contexts) await c.addInitScript(() => { try { localStorage.setItem('docx4j-editor.relay', '/relay'); } catch { /* ignore */ } });
  const [a, b] = await Promise.all(contexts.map((c) => c.newPage()));
  await openReview(a!, 'revisions/revisions-word15.docx');
  await a!.getByTestId('share').click();
  const link = await a!.getByTestId('share-link').inputValue();
  await a!.keyboard.press('Escape');
  await b!.goto(link);
  await b!.getByTestId('join-accept').click();
  await b!.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await b!.getByTestId('perspective-review').click();
  await expect(b!.getByTestId('panel-review')).toBeVisible();
  const before = await panelEntries(a!);
  expect(await panelEntries(b!)).toEqual(before);
  const index = before.findIndex((e) => e.startsWith('change|Added|'));
  await b!.locator('[data-testid="review-entry"]').nth(index).getByTestId('review-accept').click();
  const after = before.filter((_, i) => i !== index);
  await expect.poll(() => panelEntries(b!)).toEqual(after);
  await expect.poll(() => panelEntries(a!), { timeout: 10_000 }).toEqual(after);
  const [savedA, savedB] = [await download(a!), await download(b!)];
  expect(comparePackages(savedA, savedB)).toEqual([]);
  await Promise.all(contexts.map((c) => c.close()));
});

// The room repair (ED-005 section 12.15), through the view and a real relay: B, offline, accepts an insertion while A types
// inside it; back online, both pages keep A's typing pending and the rest accepted.
test('in a room: an accept made offline while the other page types inside the change leaves the typing pending, on both pages', async ({ browser }) => {
  const contexts = await Promise.all([browser.newContext(), browser.newContext()]);
  await contexts[0]!.addInitScript(() => { try { localStorage.setItem('docx4j-editor.relay', '/relay'); localStorage.setItem('docx4j-editor.name', 'Author A'); } catch { /* ignore */ } });
  await contexts[1]!.addInitScript(() => { try { localStorage.setItem('docx4j-editor.relay', '/relay'); localStorage.setItem('docx4j-editor.name', 'Author B'); } catch { /* ignore */ } });
  const [a, b] = await Promise.all(contexts.map((c) => c.newPage()));
  await openFixture(a!, 'tracked-changes.docx');
  await a!.getByTestId('share').click();
  const link = await a!.getByTestId('share-link').inputValue();
  await a!.keyboard.press('Escape');
  await b!.goto(link);
  await b!.getByTestId('join-accept').click();
  await b!.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await expect(b!.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 20_000 });
  // A turns tracking on and types an insertion into "Some text.", which B sees.
  await a!.getByTestId('tracking-switch').click();
  await expect(b!.getByTestId('tracking-switch')).toHaveAttribute('aria-pressed', 'true', { timeout: 10_000 });
  const block = await a!.evaluate(() => { const d = (window as unknown as Win).docx4jEditor.state.doc; for (let i = 0; i < d.childCount; i++) if (d.child(i).textContent === 'Some text.') return i; return -1; });
  const startOf = (page: Page): Promise<number> => blockStart(page, block);
  await a!.locator(body).focus();
  await a!.evaluate((p) => (window as unknown as Win).docx4jEditor.caretAt(p), (await startOf(a!)) + 5);
  await a!.keyboard.type('new words ');
  await expect.poll(() => b!.evaluate((i) => (window as unknown as Win).docx4jEditor.state.doc.child(i).textContent, block), { timeout: 10_000 }).toBe('Some new words text.');
  // B offline accepts it from the panel; A types inside it meanwhile.
  await b!.getByTestId('perspective-review').click();
  await contexts[1]!.setOffline(true);
  await expect(b!.getByTestId('room-status')).toHaveAttribute('data-status', 'offline', { timeout: 20_000 });
  await b!.locator('[data-testid="review-entry"][data-author="Author A"]').filter({ hasText: 'new words' }).getByTestId('review-accept').click();
  await a!.evaluate((p) => (window as unknown as Win).docx4jEditor.caretAt(p), (await startOf(a!)) + 9);
  await a!.keyboard.type('more ');
  await contexts[1]!.setOffline(false);
  await expect(b!.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 20_000 });
  const pending = (page: Page): Promise<string> => page.evaluate((i) => {
    const out: string[] = [];
    (window as unknown as Win).docx4jEditor.state.doc.child(i).forEach((n) => { if (n.isText) out.push(`${n.text}${n.marks.some((m) => m.type.name === 'insertion') ? '[ins]' : ''}`); });
    return out.join('|');
  }, block);
  await expect.poll(() => pending(a!), { timeout: 10_000 }).toBe('Some new |more [ins]|words text.');
  await expect.poll(() => pending(b!), { timeout: 10_000 }).toBe('Some new |more [ins]|words text.');
  await Promise.all(contexts.map((c) => c.close()));
});

// ED-005 section 12.113: what a Contents refresh in ONLYOFFICE leaves, listed as a reader needs it.
test('revisions around nothing draw no bar and list nothing, a Contents line is one entry, and the banner counts what the panel lists', async ({ page }) => {
  await openFixture(page, 'authorless-contents.docx');
  await expect(page.getByTestId('banner')).toContainText('8 tracked changes');
  // The two links stand apart: JSX had dropped the line break between them ("ShowOpen the Review perspective").
  await expect(page.getByTestId('banner')).toContainText('Show Open the Review perspective');
  await page.getByTestId('perspective-review').click();
  const entries = page.getByTestId('review-entry');
  await expect(entries).toHaveCount(8);
  await expect(page.getByTestId('review-count')).toHaveText('8 changes by 2 people');
  await expect(entries.nth(0)).toContainText('2→Membership generally→4');
  await expect(entries.nth(0)).toContainText('Unknown author');
  // A paragraph mark with an empty author is the same unknown author.
  await expect(entries.nth(1)).toContainText('Unknown author');
  // The three headings hold only wrappers beside their text: no change bar. The Contents lines above them keep theirs.
  const barred = await page.evaluate(() => [...document.querySelectorAll('.ProseMirror[aria-label="Document"] p')].filter((p) => p.classList.contains('has-change')).map((p) => p.textContent));
  expect(barred.filter((text) => text === 'Membership generally' || text === 'Register of members')).toEqual([]);
  expect(barred.length).toBeGreaterThanOrEqual(3);

  // The unknown author's group clears all of that author's changes, in both forms, and leaves the named one's.
  await page.getByTestId('review-by-author').check();
  await page.locator('[data-testid="review-author"][data-author=""]').getByTestId('review-author-accept').click();
  await expect(entries).toHaveCount(2);
  await expect(page.getByTestId('review-count')).toHaveText('2 changes by 1 person');
  await expect(page.locator('[data-testid="review-author"][data-author=""]')).toHaveCount(0);
});
