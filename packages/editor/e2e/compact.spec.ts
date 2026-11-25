// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-013 S1 to S3: the compact layout for phones (section 6). A phone is a Pixel 7's profile (touch, mobile,
// 412 × 915); the emulator shows no keyboard, so one is mimicked by shrinking the viewport's height, which is what
// `interactive-widget=resizes-content` does on the phone. The bar above the keyboard is tapped, as a finger would.
import { test, expect, devices, type Page } from '@playwright/test';
import { ZipPartStore } from '@docx4j/core-ts';
import { fixturePath, caretInCell, caretInBlock } from './helpers.js';

const { defaultBrowserType: _browser, ...pixel } = devices['Pixel 7'];
void _browser;

/** Opens a fixture through the file input the compact layout keeps (the ribbon the usual helper waits for is not drawn). */
async function openCompact(page: Page, name: string): Promise<void> {
  await page.goto('/');
  await page.getByTestId('open').setInputFiles(fixturePath(name));
  await page.locator('.ProseMirror[aria-label="Document"]').waitFor();
}

type Win = { docx4jEditor: { hasFocus(): boolean; caretAt(from: number, to?: number): void; state: { doc: { child(i: number): { nodeSize: number }; descendants(f: (n: { isText: boolean; marks: { type: { name: string } }[] }) => boolean | void): void } } } };

/** The keyboard mimicked: the viewport's height as `resizes-content` leaves it above Samsung Keyboard (section 1.4). */
async function openKeyboard(page: Page): Promise<void> {
  const size = page.viewportSize()!;
  await page.setViewportSize({ width: size.width, height: Math.round(size.height * 0.5) });
  await expect(page.locator('.app')).toHaveClass(/\bkeyboard-open\b/);
}

/** Download through the menu sheet, and the saved file's part. */
async function savedPart(page: Page, name: string): Promise<string> {
  await page.getByTestId('compact-menu').tap();
  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByTestId('compact-download').tap()]);
  const { readFile } = await import('node:fs/promises');
  return new TextDecoder().decode(new ZipPartStore(new Uint8Array(await readFile(await dl.path()))).loadSync(name));
}

/** Download through the menu sheet: the saved file's bytes. */
async function downloaded(page: Page): Promise<Uint8Array> {
  await page.getByTestId('compact-menu').tap();
  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByTestId('compact-download').tap()]);
  const { readFile } = await import('node:fs/promises');
  return new Uint8Array(await readFile(await dl.path()));
}

/** Opens a saved file's bytes through the compact layout's file input, in place of the document. */
async function reopen(page: Page, bytes: Uint8Array, name: string): Promise<void> {
  await page.getByTestId('open').setInputFiles({ name, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(bytes) });
  await expect(page.getByTestId('name')).toHaveText(name);
}

/**
 * The document's share of the screen (section 6): the height between the top bar and the bar above the keyboard (or the
 * viewport's end), over the screen's height with the keyboard closed.
 */
async function documentShare(page: Page, screenHeight: number): Promise<number> {
  return page.evaluate((full) => {
    const top = document.querySelector('[data-testid="compact-bar"]')?.getBoundingClientRect().bottom ?? 0;
    const bottom = document.querySelector('[data-testid="keyboard-bar"]')?.getBoundingClientRect().top ?? innerHeight;
    return (bottom - top) / full;
  }, screenHeight);
}

/** The document's text nodes carrying a revision mark. */
async function revisedTexts(page: Page): Promise<number> {
  return page.evaluate(() => {
    let n = 0;
    (window as unknown as Win).docx4jEditor.state.doc.descendants((node) => { if (node.isText && node.marks.some((m) => m.type.name === 'insertion' || m.type.name === 'deletion')) n++; });
    return n;
  });
}

/** The body's text width, its padding off, and the page's sizes. */
async function measure(page: Page): Promise<{ text: number; own: number; scrollWidth: number; innerWidth: number; innerHeight: number; scrollHeight: number; bodyTop: number }> {
  return page.evaluate(() => {
    const body = document.querySelector('.docx-editor.body') as HTMLElement;
    const style = getComputedStyle(body);
    const text = body.clientWidth - parseFloat(style.paddingInlineStart) - parseFloat(style.paddingInlineEnd);
    const root = document.scrollingElement!;
    // The document's own text width, which the stylesheet states in points.
    const own = parseFloat(style.getPropertyValue('--text-width')) * 4 / 3;
    return { text, own, scrollWidth: root.scrollWidth, innerWidth, innerHeight, scrollHeight: root.scrollHeight, bodyTop: body.getBoundingClientRect().top };
  });
}

test.describe('a phone in portrait', () => {
  test.use(pixel);

  test('is compact: one bar, no ribbon, the text reflowed to the screen, nothing wider than it, the root scrolling', async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    await expect(page.locator('.app')).toHaveClass(/\bcompact\b/);
    await expect(page.getByTestId('compact-bar')).toBeVisible();
    await expect(page.locator('.app-bar:not(.compact-bar)')).toHaveCount(0);
    await expect(page.getByTestId('toolbar')).toHaveCount(0);
    const m = await measure(page);
    // The document's text is wider than the screen: the screen's width less 12 px either side.
    expect(m.own).toBeGreaterThan(m.innerWidth - 24);
    expect(m.text).toBeCloseTo(m.innerWidth - 24, 0);
    expect(m.scrollWidth).toBeLessThanOrEqual(m.innerWidth);
    // The document starts in the top part of the screen (at least 60% of it the document's), and the root is what scrolls.
    expect(m.bodyTop).toBeLessThan(m.innerHeight * 0.4);
    expect(m.scrollHeight).toBeGreaterThan(m.innerHeight);
    expect(await page.evaluate(() => getComputedStyle(document.querySelector('.page-area')!).overflowY)).toBe('visible');
  });

  test('with the keyboard open the bar drops the name alone', async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    const bar = page.getByTestId('compact-bar');
    await expect(bar.getByTestId('perspective-review')).toBeVisible();
    await expect(bar.getByTestId('name')).toBeVisible();
    const closed = page.viewportSize()!.height;
    expect(await documentShare(page, closed)).toBeGreaterThanOrEqual(0.6);
    await caretInBlock(page, 1);
    await page.setViewportSize({ width: 412, height: 450 });
    await expect(page.locator('.app')).toHaveClass(/\bkeyboard-open\b/);
    // With the keyboard open and the bar above it, at least 30% of the screen is still the document's (section 6).
    await expect(page.getByTestId('keyboard-bar')).toBeVisible();
    expect(await documentShare(page, closed)).toBeGreaterThanOrEqual(0.3);
    await expect(bar.getByTestId('name')).toHaveCount(0);
    await expect(bar.getByTestId('perspective-edit')).toBeVisible();
    await expect(bar.getByTestId('perspective-review')).toBeVisible();
    await expect(bar.getByTestId('compact-menu')).toBeVisible();
    await expect(bar.getByTestId('tracking-switch')).toBeVisible();
    await expect(bar.locator('[data-command="history.undo"]')).toBeVisible();
    await page.setViewportSize({ width: 412, height: 839 });
    await expect(page.locator('.app')).not.toHaveClass(/\bkeyboard-open\b/);
    await expect(bar.getByTestId('perspective-review')).toBeVisible();
  });

  test('the menu is a sheet: Escape closes it and the focus goes back; Full in its Options draws the full layout, kept', async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    // From the keyboard, the focus on the menu button: it comes back there. (A tap leaves the focus where it was, the
    // document's, so that the phone's keyboard comes back with it.)
    await page.getByTestId('compact-menu').focus();
    await page.keyboard.press('Enter');
    const sheet = page.getByTestId('compact-menu-sheet');
    await expect(sheet).toBeVisible();
    await expect(sheet.getByTestId('compact-download')).toBeEnabled();
    await page.keyboard.press('Escape');
    await expect(sheet).toHaveCount(0);
    await expect(page.getByTestId('compact-menu')).toBeFocused();
    await page.getByTestId('compact-menu').click();
    await sheet.getByTestId('compact-options').locator('summary').click();
    await sheet.getByTestId('option-layout').selectOption('full');
    await expect(page.locator('.app')).not.toHaveClass(/\bcompact\b/);
    await expect(page.getByTestId('toolbar')).toBeVisible();
    await page.reload();
    await expect(page.locator('.app')).not.toHaveClass(/\bcompact\b/);
  });

  test('Undo and Redo on the bar act on the document', async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    const text = (): Promise<string> => page.evaluate(() => (window as unknown as { docx4jEditor: { state: { doc: { textContent: string } } } }).docx4jEditor.state.doc.textContent);
    const before = await text();
    await page.evaluate(() => {
      const v = (window as unknown as { docx4jEditor: { view: { state: { tr: { insertText(t: string, p: number): unknown } }; dispatch(tr: unknown): void } } }).docx4jEditor.view;
      v.dispatch(v.state.tr.insertText('Typed ', 1));
    });
    await expect.poll(text).toContain('Typed ');
    const undo = page.getByTestId('compact-bar').locator('[data-command="history.undo"]');
    await expect(undo).toBeEnabled();
    await undo.click();
    await expect.poll(text).toBe(before);
    await page.getByTestId('compact-bar').locator('[data-command="history.redo"]').click();
    await expect.poll(text).toContain('Typed ');
  });

  test('Review draws no side panel', async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    await page.getByTestId('compact-bar').getByTestId('perspective-review').click();
    await expect(page.getByTestId('perspective-review')).toHaveAttribute('aria-selected', 'true');
    await expect(page.locator('.workspace')).not.toHaveClass(/with-panels/);
  });
});

test.describe('the bar above the keyboard, in portrait (S2)', () => {
  test.use(pixel);
  const bar = (page: Page) => page.getByTestId('keyboard-bar');

  test("shows while a surface has the focus and the keyboard is open, in the status bar's place; Done closes it", async ({ page }) => {
    await openCompact(page, 'tables.docx');
    await caretInBlock(page, 0);
    await expect(bar(page)).toHaveCount(0);
    await openKeyboard(page);
    await expect(bar(page)).toBeVisible();
    await expect(page.locator('.app > .status-bar')).toBeHidden();
    // Above the keyboard: its bottom at the viewport's, nothing wider than the screen.
    const box = (await bar(page).boundingBox())!;
    const view = await page.evaluate(() => ({ height: innerHeight, width: innerWidth, scrollWidth: document.scrollingElement!.scrollWidth }));
    expect(box.y + box.height).toBeCloseTo(view.height, 0);
    expect(view.scrollWidth).toBeLessThanOrEqual(view.width);
    // Outside a table, with no selection, in Edit: Done alone. It is the keyboard-hiding symbol, its name in its label.
    await expect(bar(page).getByRole('button')).toHaveCount(1);
    const done = bar(page).getByRole('button', { name: 'Hide the keyboard' });
    await expect(done).toHaveAttribute('data-testid', 'keyboard-done');
    await expect(done).toHaveText('');
    await expect(done.locator('svg')).toBeVisible();
    await done.tap();
    await expect(bar(page)).toHaveCount(0);
    expect(await page.evaluate(() => (window as unknown as Win).docx4jEditor.hasFocus())).toBe(false);
  });

  test('the caret comes back into view above the bar as the keyboard opens', async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    // A caret in the lower half of the screen, which the keyboard then covers.
    const block = await page.evaluate(() => {
      const v = (window as unknown as Win).docx4jEditor as unknown as { view: { state: { doc: { childCount: number; child(i: number): { nodeSize: number; isTextblock: boolean; content: { size: number } } } }; coordsAtPos(p: number): { top: number } } };
      let pos = 0;
      for (let i = 0; i < v.view.state.doc.childCount; i++) {
        const n = v.view.state.doc.child(i);
        if (n.isTextblock && n.content.size > 0 && v.view.coordsAtPos(pos + 1).top > innerHeight * 0.65 && v.view.coordsAtPos(pos + 1).top < innerHeight - 40) return i;
        pos += n.nodeSize;
      }
      return -1;
    });
    expect(block).toBeGreaterThan(0);
    await caretInBlock(page, block);
    await openKeyboard(page);
    await expect(bar(page)).toBeVisible();
    const barTop = (await bar(page).boundingBox())!.y;
    await expect.poll(() => page.evaluate(() => {
      const v = (window as unknown as Win).docx4jEditor as unknown as { view: { state: { selection: { head: number } }; coordsAtPos(p: number): { bottom: number } } };
      return v.view.coordsAtPos(v.view.state.selection.head).bottom;
    })).toBeLessThanOrEqual(barTop);
  });

  test('in a table: Row above, Row below and Delete row act on the table and keep the focus (the keyboard up)', async ({ page }) => {
    await openCompact(page, 'tables.docx');
    const table = page.locator('.docx-editor .ProseMirror[aria-label="Document"] > .tableWrapper > table').nth(0);   // A B / C D
    await expect(table.locator('tr')).toHaveCount(2);
    await caretInCell(page, 1, 0, 0);
    await openKeyboard(page);
    await expect(bar(page).getByTestId('keyboard-group-table')).toBeVisible();
    await bar(page).getByRole('button', { name: 'Add row below' }).tap();
    await expect(table.locator('tr')).toHaveCount(3);
    await bar(page).getByRole('button', { name: 'Add row above' }).tap();
    await expect(table.locator('tr')).toHaveCount(4);
    expect(await page.evaluate(() => (window as unknown as Win).docx4jEditor.hasFocus())).toBe(true);
    await bar(page).getByRole('button', { name: 'Delete row' }).tap();
    await expect(table.locator('tr')).toHaveCount(3);
    await expect(bar(page)).toBeVisible();
  });

  test('a selection offers Comment: the sheet quotes it and asks the name the first time, the focus comes back, and the saved file has the comment with its text', async ({ page }) => {
    await openCompact(page, 'tables.docx');
    await caretInBlock(page, 0);
    await openKeyboard(page);
    await expect(bar(page).getByTestId('keyboard-comment')).toHaveCount(0);
    await page.evaluate(() => { const v = (window as unknown as Win).docx4jEditor; v.caretAt(1, 4); });
    await bar(page).getByTestId('keyboard-comment').tap();
    const sheet = page.getByTestId('comment-sheet');
    await expect(sheet).toBeVisible();
    await expect(sheet.getByTestId('comment-sheet-quote')).toHaveText('Sha');
    await expect(sheet.getByTestId('comment-sheet-text')).toBeFocused();
    await expect(sheet.getByTestId('comment-sheet-post')).toBeDisabled();
    // No name set: the sheet asks it, from the guest's name; no dialog follows (it lost the phone's keyboard, section 9).
    await expect(sheet.getByTestId('comment-sheet-name')).toHaveValue(/Guest/);
    await sheet.getByTestId('comment-sheet-name').fill('Reviewer One');
    await sheet.getByTestId('comment-sheet-text').fill('From the phone');
    await sheet.getByTestId('comment-sheet-post').tap();
    await expect(sheet).toHaveCount(0);
    await expect(page.getByTestId('tracking-name')).toHaveCount(0);
    await expect(page.locator('.docx-editor .ProseMirror[aria-label="Document"] .comment').first()).toBeVisible();
    // The focus is the document's again, so the keyboard and its bar stay.
    await expect(bar(page)).toBeVisible();
    expect(await page.evaluate(() => (window as unknown as Win).docx4jEditor.hasFocus())).toBe(true);
    // The name is kept: the next sheet does not ask it. Cancel makes nothing.
    await page.evaluate(() => { const v = (window as unknown as Win).docx4jEditor; v.caretAt(1, 3); });
    await bar(page).getByTestId('keyboard-comment').tap();
    await expect(page.getByTestId('comment-sheet-name')).toHaveCount(0);
    await page.getByTestId('comment-sheet-cancel').tap();
    await expect(page.getByTestId('comment-sheet')).toHaveCount(0);
    await page.setViewportSize({ width: 412, height: 839 });
    const comments = await savedPart(page, 'word/comments.xml');
    expect(comments).toContain('From the phone');
    expect(comments).toContain('w:author="Reviewer One"');
    expect(comments.match(/<w:comment /g)).toHaveLength(1);
  });

  test('in Review: Next goes to a change, where Accept and Reject resolve it through the four commands', async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    await page.getByTestId('compact-bar').getByTestId('perspective-review').tap();
    // Choosing Review opens the changes' sheet (S3); closed, the document is reviewed from the bar.
    await page.getByTestId('review-sheet').getByTestId('sheet-close').tap();
    await caretInBlock(page, 0);
    await openKeyboard(page);
    const before = await revisedTexts(page);
    expect(before).toBeGreaterThan(0);
    await bar(page).getByTestId('keyboard-next').tap();
    await expect(bar(page).getByTestId('keyboard-accept')).toBeVisible();
    await bar(page).getByTestId('keyboard-accept').tap();
    await expect.poll(() => revisedTexts(page)).toBeLessThan(before);
    const accepted = await revisedTexts(page);
    await bar(page).getByTestId('keyboard-next').tap();
    await bar(page).getByTestId('keyboard-reject').tap();
    await expect.poll(() => revisedTexts(page)).toBeLessThan(accepted);
    expect(await page.evaluate(() => (window as unknown as Win).docx4jEditor.hasFocus())).toBe(true);
  });

  test('in Review: Previous before Accept and Reject, Next after; Previous greyed at the first change, Next at the last, a tap on either keeping the focus', async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    await page.getByTestId('compact-bar').getByTestId('perspective-review').tap();
    await page.getByTestId('review-sheet').getByTestId('sheet-close').tap();
    await caretInBlock(page, 0);
    await openKeyboard(page);
    const group = bar(page).getByTestId('keyboard-group-review');
    const previous = bar(page).getByTestId('keyboard-previous');
    const next = bar(page).getByTestId('keyboard-next');
    const selection = (): Promise<[number, number]> => page.evaluate(() => { const s = (window as unknown as { docx4jEditor: { state: { selection: { from: number; to: number } } } }).docx4jEditor.state.selection; return [s.from, s.to]; });
    const focused = (): Promise<boolean> => page.evaluate(() => (window as unknown as Win).docx4jEditor.hasFocus());
    await expect(previous).toHaveAttribute('aria-disabled', 'true');
    await expect(next).toHaveAttribute('aria-disabled', 'false');
    // The first change: Previous still greyed; the four in order.
    await next.tap();
    await expect(group.getByRole('button')).toHaveText(['Previous', 'Accept', 'Reject', 'Next']);
    await expect(previous).toHaveAttribute('aria-disabled', 'true');
    // A tap on a greyed button does nothing, and the keyboard stays (forced: Playwright will not tap an aria-disabled one).
    const first = await selection();
    await previous.tap({ force: true });
    expect(await selection()).toEqual(first);
    expect(await focused()).toBe(true);
    // On to the last change, where Next is greyed and Previous is not; Previous goes back one.
    let steps = 0;
    while ((await next.getAttribute('aria-disabled')) === 'false') { await next.tap(); steps++; expect(steps).toBeLessThan(50); }
    expect(steps).toBeGreaterThan(0);
    await expect(previous).toHaveAttribute('aria-disabled', 'false');
    const last = await selection();
    await next.tap({ force: true });
    expect(await selection()).toEqual(last);
    expect(await focused()).toBe(true);
    await previous.tap();
    expect(await selection()).not.toEqual(last);
    await expect(next).toHaveAttribute('aria-disabled', 'false');
    await expect(bar(page)).toBeVisible();
  });

  test('the tracking switch and Edit keep the keyboard open; Review takes the focus for its sheet', async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    const top = page.getByTestId('compact-bar');
    const focused = (): Promise<boolean> => page.evaluate(() => (window as unknown as Win).docx4jEditor.hasFocus());
    await caretInBlock(page, 1);
    await openKeyboard(page);
    await expect(bar(page)).toBeVisible();
    // The first time the switch asks the name in a dialog, which takes the focus and gives it back to the document.
    const toggle = top.getByTestId('tracking-switch');
    await toggle.tap();
    const name = page.getByTestId('tracking-name');
    await expect(name).toBeFocused();
    await name.fill('Reviewer One');
    await name.press('Enter');
    await expect(toggle).toHaveAttribute('aria-pressed', 'true');
    await expect.poll(focused).toBe(true);
    await expect(bar(page)).toBeVisible();
    // With a name, the focus never leaves the document (a tap had closed the keyboard mid-sentence, section 9).
    await toggle.tap();
    await expect(toggle).toHaveAttribute('aria-pressed', 'false');
    expect(await focused()).toBe(true);
    await toggle.tap();
    await expect(toggle).toHaveAttribute('aria-pressed', 'true');
    expect(await focused()).toBe(true);
    // Review opens its sheet with the focus, so no keyboard comes up over it; Edit opens nothing and keeps the keyboard.
    await top.getByTestId('perspective-review').tap();
    await expect(page.getByTestId('review-sheet')).toBeVisible();
    expect(await focused()).toBe(false);
    await page.getByTestId('review-sheet').getByTestId('sheet-close').tap();
    await caretInBlock(page, 1);
    await expect(bar(page)).toBeVisible();
    await top.getByTestId('perspective-edit').tap();
    await expect(top.getByTestId('perspective-edit')).toHaveAttribute('aria-selected', 'true');
    expect(await focused()).toBe(true);
    await expect(bar(page)).toBeVisible();
  });
});

test.describe('the sheets, in portrait (S3)', () => {
  test.use(pixel);
  const bar = (page: Page) => page.getByTestId('keyboard-bar');
  const hasFocus = (page: Page) => page.evaluate(() => (window as unknown as Win).docx4jEditor.hasFocus());

  test('choosing Review opens the changes as a sheet, the document blurred; an entry closes it with the caret in the change, where the bar offers Accept', async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    // The open leaves the caret in the document, as on the phone, where the tap on Review then raised the keyboard (S2's note).
    await caretInBlock(page, 0);
    await page.getByTestId('compact-bar').getByTestId('perspective-review').tap();
    const sheet = page.getByTestId('review-sheet');
    await expect(sheet).toBeVisible();
    await expect(page.getByTestId('perspective-review')).toHaveAttribute('aria-selected', 'true');
    expect(await hasFocus(page)).toBe(false);
    // Each entry has Accept and Reject; one Accepted stays in the sheet. (The panel is a chunk of its own: wait for it.)
    const entries = sheet.getByTestId('review-entry');
    await expect(entries.first()).toBeVisible();
    const before = await entries.count();
    expect(before).toBeGreaterThanOrEqual(2);
    await sheet.locator('[data-testid="review-entry"][data-kind="change"]:not([data-type="Formatted"])').first().getByTestId('review-accept').tap();
    await expect(entries).toHaveCount(before - 1);
    await expect(sheet).toBeVisible();
    // An entry: the sheet closes, the caret in its change; with the keyboard up, the bar's Accept resolves it.
    const revised = await revisedTexts(page);
    await entries.first().locator('.review-entry-main').tap();
    await expect(sheet).toHaveCount(0);
    expect(await hasFocus(page)).toBe(true);
    await openKeyboard(page);
    await expect(bar(page).getByTestId('keyboard-accept')).toBeVisible();
    await bar(page).getByTestId('keyboard-accept').tap();
    await expect.poll(() => revisedTexts(page)).toBeLessThan(revised);
  });

  test("the changes' sheet: Accept all asks first; Next closes it on a change; a swipe down its handle closes it", async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    await page.getByTestId('compact-bar').getByTestId('perspective-review').tap();
    const sheet = page.getByTestId('review-sheet');
    await expect(sheet.getByTestId('review-entry').first()).toBeVisible();
    const before = await sheet.getByTestId('review-entry').count();
    await sheet.getByTestId('review-accept-all').tap();
    await expect(sheet.getByTestId('review-confirm')).toContainText(/Accept all \d+ changes\?/);
    await sheet.getByTestId('review-confirm-no').tap();
    await expect(sheet.getByTestId('review-confirm')).toHaveCount(0);
    await expect(sheet.getByTestId('review-entry')).toHaveCount(before);
    // Next: the first change selected, the sheet closed.
    await sheet.getByTestId('review-next').tap();
    await expect(sheet).toHaveCount(0);
    expect(await page.evaluate(() => (window as unknown as { docx4jEditor: { state: { selection: { empty: boolean } } } }).docx4jEditor.state.selection.empty)).toBe(false);
    // Review chosen again opens it again; a swipe down closes it, a short one does not.
    await page.getByTestId('compact-bar').getByTestId('perspective-review').tap();
    await expect(sheet).toBeVisible();
    const handle = (await sheet.locator('.sheet-handle').boundingBox())!;
    const x = handle.x + handle.width / 2;
    const y = handle.y + handle.height / 2;
    await page.mouse.move(x, y);
    await page.mouse.down();
    await page.mouse.move(x, y + 20, { steps: 4 });
    await page.mouse.up();
    await expect(sheet).toBeVisible();
    await page.mouse.move(x, y);
    await page.mouse.down();
    await page.mouse.move(x, y + 150, { steps: 6 });
    await page.mouse.up();
    await expect(sheet).toHaveCount(0);
    // Accept all, confirmed.
    await page.getByTestId('compact-bar').getByTestId('perspective-review').tap();
    await sheet.getByTestId('review-accept-all').tap();
    await sheet.getByTestId('review-confirm-yes').tap();
    await expect(sheet.getByTestId('review-count')).toHaveText(/No tracked changes/);
    await expect.poll(() => revisedTexts(page)).toBe(0);
  });

  test('Comments (n) on commented text opens its thread: read-only, a reply with the name asked, Resolve, and the saved file has both', async ({ page }) => {
    await openCompact(page, 'comments-two.docx');
    // "One paragraph", comment 0 on "paragraph" (positions 5 to 14), by jharrop2; its text keeps its highlight.
    await page.evaluate(() => (window as unknown as Win).docx4jEditor.caretAt(8));
    await expect.poll(() => page.evaluate(() => getComputedStyle(document.querySelector('.docx-editor .comment[data-comment="0"]')!).getPropertyValue('--comment-color').trim())).not.toBe('');
    await openKeyboard(page);
    const comments = bar(page).getByTestId('keyboard-comments');
    await expect(comments).toHaveText('Comments (1)');
    await page.evaluate(() => (window as unknown as Win).docx4jEditor.caretAt(2));
    await expect(comments).toHaveCount(0);
    await page.evaluate(() => (window as unknown as Win).docx4jEditor.caretAt(8));
    await comments.tap();
    const sheet = page.getByTestId('thread-sheet');
    await expect(sheet).toBeVisible();
    await expect(sheet.getByTestId('sheet-thread')).toHaveCount(1);
    await expect(sheet.getByTestId('sheet-thread')).toContainText('jharrop2');
    await expect(sheet.getByTestId('sheet-thread')).toContainText('One comment');
    // Not this page's: no Delete, nothing editable. The field is not focused, so the keyboard goes and the thread can be read.
    await expect(sheet.getByTestId('comment-delete')).toHaveCount(0);
    await expect(sheet.locator('[data-editable="true"]')).toHaveCount(0);
    await expect(sheet.getByTestId('thread-sheet-text')).not.toBeFocused();
    // A reply: no name set, so the field asks it. The keyboard went as the sheet opened and comes back as the field is
    // tapped: the field's Reply stays in view above it, the sheet scrolled to it (on the phone it was under the keyboard).
    await page.setViewportSize({ width: 412, height: 839 });
    await sheet.getByTestId('thread-sheet-text').tap();
    await openKeyboard(page);
    const postBottom = (): Promise<number> => page.evaluate(() => document.querySelector('[data-testid="thread-sheet-post"]')!.getBoundingClientRect().bottom - innerHeight);
    await expect.poll(postBottom).toBeLessThanOrEqual(0);
    await expect(sheet.getByTestId('thread-sheet-name')).toHaveValue(/Guest/);
    await sheet.getByTestId('thread-sheet-name').fill('Reviewer One');
    await sheet.getByTestId('thread-sheet-text').fill('From the phone');
    await sheet.getByTestId('thread-sheet-post').tap();
    const reply = sheet.locator('.comment-reply');
    await expect(reply).toContainText('Reviewer One');
    await expect(reply).toContainText('From the phone');
    await expect(reply.getByTestId('comment-delete')).toHaveCount(1);
    await expect(sheet.getByTestId('thread-sheet-name')).toHaveCount(0);
    await expect(page.getByTestId('tracking-name')).toHaveCount(0);
    // Resolve: marked, and no field (a resolved thread takes no reply).
    await sheet.getByTestId('comment-resolve').tap();
    await expect(sheet.getByTestId('sheet-thread-resolved')).toBeVisible();
    await expect(sheet.getByTestId('thread-sheet-text')).toHaveCount(0);
    await sheet.getByTestId('sheet-close').tap();
    await expect(sheet).toHaveCount(0);
    await page.setViewportSize({ width: 412, height: 839 });
    const part = await savedPart(page, 'word/comments.xml');
    expect(part).toContain('From the phone');
    expect(part).toContain('w:author="Reviewer One"');
    expect(part.match(/<w:comment /g)).toHaveLength(3);
    const extended = await savedPart(page, 'word/commentsExtended.xml');
    expect(extended).toMatch(/w15:done="1"/);
  });
});

test.describe('saved and opened again (section 6)', () => {
  test.use(pixel);
  const bar = (page: Page) => page.getByTestId('keyboard-bar');

  test("a review made on the phone is in the saved file: a change accepted, one rejected, and a comment's reply and resolve read back", async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    const body = page.locator('.ProseMirror[aria-label="Document"]');
    await expect(body).toContainText('An insertion');
    await page.getByTestId('compact-bar').getByTestId('perspective-review').tap();
    const sheet = page.getByTestId('review-sheet');
    const entries = sheet.getByTestId('review-entry');
    await expect(entries).toHaveCount(2);
    // The insertion accepted, the deletion rejected: both texts stay, as plain text.
    await entries.filter({ hasText: 'An insertion' }).getByTestId('review-accept').tap();
    await expect(entries).toHaveCount(1);
    await entries.filter({ hasText: 'A deletion' }).getByTestId('review-reject').tap();
    await expect(sheet.getByTestId('review-count')).toHaveText(/No tracked changes/);
    await sheet.getByTestId('sheet-close').tap();
    await reopen(page, await downloaded(page), 'tracked-changes.docx');
    await expect(body).toContainText('An insertion Followed by A deletion.');
    expect(await revisedTexts(page)).toBe(0);

    await openCompact(page, 'comments-two.docx');
    await page.evaluate(() => (window as unknown as Win).docx4jEditor.caretAt(8));
    await openKeyboard(page);
    await bar(page).getByTestId('keyboard-comments').tap();
    const thread = page.getByTestId('thread-sheet');
    await thread.getByTestId('thread-sheet-text').tap();
    await thread.getByTestId('thread-sheet-name').fill('Reviewer One');
    await thread.getByTestId('thread-sheet-text').fill('Read back');
    await thread.getByTestId('thread-sheet-post').tap();
    await expect(thread.locator('.comment-reply')).toContainText('Read back');
    await thread.getByTestId('comment-resolve').tap();
    await expect(thread.getByTestId('sheet-thread-resolved')).toBeVisible();
    await thread.getByTestId('sheet-close').tap();
    await page.setViewportSize({ width: 412, height: 839 });
    await reopen(page, await downloaded(page), 'comments-two.docx');
    await page.evaluate(() => (window as unknown as Win).docx4jEditor.caretAt(8));
    await openKeyboard(page);
    await bar(page).getByTestId('keyboard-comments').tap();
    await expect(thread.getByTestId('sheet-thread-resolved')).toBeVisible();
    await expect(thread.locator('.comment-reply')).toContainText('Reviewer One');
    await expect(thread.locator('.comment-reply')).toContainText('Read back');
  });
});

/**
 * The Web Share API stubbed as a browser that shares Word documents would have it: `canShare` takes a `.docx`, `share` keeps
 * what it was handed, and the browser poses as Safari, since Send a copy is never offered in Chromium (UPSTREAM-28), which
 * this suite runs in. `fail` names the errors the first calls throw, in turn: `AbortError` is a share sheet closed without a
 * choice, `NotAllowedError` with the tap's activation there the browser refusing the file. `spent`: the first reading of
 * `userActivation` says the activation is gone. `chromium`: no pose, as Chromium is.
 */
async function stubShare(page: Page, options: { fail?: string[]; spent?: boolean; chromium?: boolean } = {}): Promise<void> {
  await page.addInitScript(({ errors, spent, chromium }: { errors: string[]; spent: boolean; chromium: boolean }) => {
    const w = window as unknown as { shared: { name: string; type: string; bytes: number[] }[] };
    w.shared = [];
    if (!chromium) {
      Object.defineProperty(navigator, 'userAgentData', { configurable: true, value: undefined });
      Object.defineProperty(navigator, 'userAgent', { configurable: true, value: 'Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Mobile/15E148 Safari/604.1' });
    }
    if (spent) {
      let first = true;
      Object.defineProperty(navigator, 'userActivation', { configurable: true, get: () => { const isActive = !first; first = false; return { isActive, hasBeenActive: true }; } });
    }
    Object.defineProperty(navigator, 'canShare', { configurable: true, value: (data: { files?: File[] }) => (data.files ?? []).every((f) => f.name.endsWith('.docx')) });
    Object.defineProperty(navigator, 'share', {
      configurable: true,
      value: async (data: { files: File[] }) => {
        const error = errors.shift();
        if (error) throw new DOMException('stubbed', error);
        for (const f of data.files) w.shared.push({ name: f.name, type: f.type, bytes: [...new Uint8Array(await f.arrayBuffer())] });
      },
    });
  }, { errors: options.fail ?? [], spent: options.spent ?? false, chromium: options.chromium ?? false });
}

const shared = (page: Page): Promise<{ name: string; type: string; bytes: number[] }[]> =>
  page.evaluate(() => (window as unknown as { shared: { name: string; type: string; bytes: number[] }[] }).shared);

const SENT = 'A copy of tracked-changes.docx went to the app you chose';

/** Send a copy through the menu sheet, tapped. */
async function sendCopy(page: Page): Promise<void> {
  await page.getByTestId('compact-menu').tap();
  await page.getByTestId('compact-send-copy').tap();
}

/** Whether the palette lists Send a copy (the menu sheet's All commands; the palette closed again). */
async function paletteOffersSend(page: Page): Promise<boolean> {
  await page.getByTestId('compact-menu').tap();
  await page.getByTestId('compact-palette').tap();
  await page.getByTestId('palette').locator('input').fill('send a copy');
  const offered = (await page.getByTestId('palette').locator('[data-action="send-copy"]').count()) > 0;
  await page.keyboard.press('Escape');
  return offered;
}

test.describe('Send a copy, in portrait (S4)', () => {
  test.use(pixel);

  test("the menu's Send a copy hands the saved .docx to the share sheet; the document stays unsaved", async ({ page }) => {
    await stubShare(page);
    await openCompact(page, 'tracked-changes.docx');
    await page.evaluate(() => {
      const v = (window as unknown as { docx4jEditor: { view: { state: { tr: { insertText(t: string, p: number): unknown } }; dispatch(tr: unknown): void } } }).docx4jEditor.view;
      v.dispatch(v.state.tr.insertText('Sent ', 1));
    });
    await expect(page.getByTestId('name')).toHaveText('tracked-changes.docx •');
    await sendCopy(page);
    await expect(page.getByTestId('status')).toHaveText(SENT);
    const files = await shared(page);
    expect(files.map((f) => [f.name, f.type])).toEqual([['tracked-changes.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document']]);
    // The file is the saved document, the edit in it.
    expect(new TextDecoder().decode(new ZipPartStore(Uint8Array.from(files[0]!.bytes)).loadSync('word/document.xml'))).toContain('Sent ');
    // A copy went elsewhere: the file the document came from is not written, so it is still unsaved.
    await expect(page.getByTestId('name')).toHaveText('tracked-changes.docx •');
  });

  test('a share sheet closed without a choice is "Not sent"; a refusal with the tap\'s activation there is said, and Send a copy is offered no more', async ({ page }) => {
    await stubShare(page, { fail: ['AbortError', 'NotAllowedError'] });
    await openCompact(page, 'tracked-changes.docx');
    await sendCopy(page);
    await expect(page.getByTestId('status')).toHaveText('Not sent');
    await sendCopy(page);
    await expect(page.getByTestId('status')).toHaveText('This browser will not send Word documents: use Download, then share the downloaded file');
    expect(await shared(page)).toEqual([]);
    await page.getByTestId('compact-menu').tap();
    await expect(page.getByTestId('compact-download')).toBeVisible();
    await expect(page.getByTestId('compact-send-copy')).toHaveCount(0);
    await page.keyboard.press('Escape');
    expect(await paletteOffersSend(page)).toBe(false);
  });

  test('a spent activation asks a second tap, which sends the copy made', async ({ page }) => {
    await stubShare(page, { spent: true });
    await openCompact(page, 'tracked-changes.docx');
    await sendCopy(page);
    await expect(page.getByTestId('send-ready-body')).toHaveText('The copy of tracked-changes.docx is ready.');
    expect(await shared(page)).toEqual([]);
    await page.getByRole('dialog').getByRole('button', { name: 'Send' }).tap();
    await expect(page.getByTestId('status')).toHaveText(SENT);
    expect((await shared(page)).map((f) => f.name)).toEqual(['tracked-changes.docx']);
    await expect(page.getByRole('dialog')).toHaveCount(0);
  });

  test('the palette has Send a copy too; neither the menu nor the palette offers it in Chromium (UPSTREAM-28) or without the API', async ({ page, context }) => {
    await stubShare(page);
    await openCompact(page, 'tracked-changes.docx');
    expect(await paletteOffersSend(page)).toBe(true);
    for (const setUp of [
      (p: Page) => stubShare(p, { chromium: true }),
      (p: Page) => p.addInitScript(() => { delete (Navigator.prototype as { share?: unknown }).share; delete (Navigator.prototype as { canShare?: unknown }).canShare; }),
    ]) {
      const other = await context.newPage();
      await setUp(other);
      await openCompact(other, 'tracked-changes.docx');
      await other.getByTestId('compact-menu').tap();
      await expect(other.getByTestId('compact-download')).toBeVisible();
      await expect(other.getByTestId('compact-send-copy')).toHaveCount(0);
      await other.keyboard.press('Escape');
      expect(await paletteOffersSend(other)).toBe(false);
      await other.close();
    }
  });
});

test.describe('a phone in landscape', () => {
  test.use({ ...pixel, viewport: { width: 915, height: 412 }, contextOptions: { screen: { width: 915, height: 412 } } });

  test("stays compact, and draws the page's own text width where it fits", async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    await expect(page.locator('.app')).toHaveClass(/\bcompact\b/);
    const m = await measure(page);
    // The document's own text width, narrower than 915 less 24.
    expect(m.own).toBeLessThan(m.innerWidth - 24);
    expect(m.text).toBeCloseTo(m.own, 0);
    expect(m.scrollWidth).toBeLessThanOrEqual(m.innerWidth);
  });

  test('with the keyboard open, the top bar is not drawn', async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    await expect(page.getByTestId('compact-bar')).toBeVisible();
    await page.setViewportSize({ width: 915, height: 150 });
    await expect(page.locator('.app')).toHaveClass(/\blandscape\b/);
    await expect(page.getByTestId('compact-bar')).toHaveCount(0);
  });

  test('with the keyboard open, the bar above it is a strip: Done and one group (P11)', async ({ page }) => {
    await openCompact(page, 'tables.docx');
    await caretInCell(page, 1, 0, 0);
    await page.evaluate(() => { const v = (window as unknown as Win).docx4jEditor; const from = v.state.doc.child(0).nodeSize + 4; v.caretAt(from, from + 1); });
    await page.setViewportSize({ width: 915, height: 150 });
    const bar = page.getByTestId('keyboard-bar');
    await expect(bar).toBeVisible();
    expect((await bar.boundingBox())!.height).toBeLessThanOrEqual(34);
    // A selection in a table: Comment, the first group, alone.
    await expect(bar.getByTestId('keyboard-comment')).toBeVisible();
    await expect(bar.getByTestId('keyboard-group-table')).toHaveCount(0);
    await expect(bar.getByTestId('keyboard-done')).toBeVisible();
  });
});

test.describe('a desktop', () => {
  test('is full by default, and compact when the option says so', async ({ page }) => {
    await page.goto('/');
    await expect(page.locator('.app')).not.toHaveClass(/\bcompact\b/);
    await page.evaluate(() => localStorage.setItem('docx4j-editor.layout', 'compact'));
    await openCompact(page, 'tracked-changes.docx');
    await expect(page.locator('.app')).toHaveClass(/\bcompact\b/);
    await expect(page.getByTestId('compact-bar')).toBeVisible();
    await page.evaluate(() => localStorage.removeItem('docx4j-editor.layout'));
  });
});
