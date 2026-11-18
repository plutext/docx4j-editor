// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 3.9, steps A2 and A3: the find bar in the browser. Ctrl+F with the count and the
// highlights, next and previous, a regular expression with a group in the replacement, replace
// one, replace all and its undo as one step, Escape back to the document on the current hit, a
// Hebrew query in a right-to-left document, the notes and the headers after the body, and two
// pages in a room, one replacing while the other types.
import { test, expect, type Page, type BrowserContext } from '@playwright/test';
import { openFixture, caretInBlock, selectBlock } from './helpers.js';

const blockText = (page: Page, i: number): Promise<string> =>
  page.evaluate((index) => (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc.child(index).textContent, i);

const bodyText = (page: Page): Promise<string> =>
  page.evaluate(() => (window as unknown as { docx4jEditor: { state: { doc: { textContent: string } } } }).docx4jEditor.state.doc.textContent);

const selectedText = (page: Page, surface = 'body'): Promise<string> =>
  page.evaluate((id) => {
    const s = (window as unknown as { docx4jSurfaces: Map<string, { state: { doc: { textBetween(a: number, b: number): string }; selection: { from: number; to: number } } }> }).docx4jSurfaces.get(id)!;
    return s.state.doc.textBetween(s.state.selection.from, s.state.selection.to);
  }, surface);

/** The top-level block indices the body's selection starts and ends in, as [from, to] offsets from the first hit's block. */
const selectionBlocks = (page: Page): Promise<[number, number]> =>
  page.evaluate(() => {
    const { selection } = (window as unknown as { docx4jEditor: { state: { selection: { $from: { index(d: number): number }; $to: { index(d: number): number } } } } }).docx4jEditor.state;
    const a = selection.$from.index(0);
    return [0, selection.$to.index(0) - a];
  });

const undo = (page: Page): Promise<void> => page.evaluate(() => { (window as unknown as { docx4jEditor: { run(n: string): boolean } }).docx4jEditor.run('history.undo'); });

async function openFind(page: Page, replace = false): Promise<void> {
  await caretInBlock(page, 0);
  await page.keyboard.press(replace ? 'Control+h' : 'Control+f');
  await expect(page.getByTestId('find-bar')).toBeVisible();
  await expect(page.getByTestId('find-query')).toBeFocused();
}

test('find, next and previous, highlights, a regular expression with a group, replace one, replace all and its undo, Escape', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await openFind(page);
  await page.getByTestId('find-query').fill('indented');
  await expect(page.getByTestId('find-count')).toHaveText('1 of 20');
  await expect(page.getByTestId('editor').locator('.find-hit')).toHaveCount(20);
  await expect(page.getByTestId('editor').locator('.find-hit.current')).toHaveCount(1);
  expect((await selectedText(page)).toLowerCase()).toBe('indented');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('find-count')).toHaveText('2 of 20');
  await page.keyboard.press('Shift+Enter');
  await expect(page.getByTestId('find-count')).toHaveText('1 of 20');
  await page.keyboard.press('Shift+Enter');
  await expect(page.getByTestId('find-count')).toHaveText('20 of 20');
  await page.keyboard.press('F3');
  await expect(page.getByTestId('find-count')).toHaveText('1 of 20');

  // Match case: the lower-case ones, block 17's first word ("Indented") left out.
  await page.getByTestId('find-match-case').check();
  await expect(page.getByTestId('find-count')).toHaveText(/ of 19$/);
  await page.getByTestId('find-match-case').uncheck();

  // An invalid expression says why and finds nothing.
  await page.getByTestId('find-regexp').check();
  await page.getByTestId('find-query').fill('(Some');
  await expect(page.getByTestId('find-error')).toBeVisible();
  await expect(page.getByTestId('editor').locator('.find-hit')).toHaveCount(0);

  // A regular expression with two groups, swapped in the replacement.
  await page.getByTestId('find-query').fill('(Some) (text)');
  await expect(page.getByTestId('find-count')).toHaveText('1 of 1');
  await page.getByTestId('find-toggle-replace').click();
  await page.getByTestId('find-replacement').fill('$2 $1');
  await page.getByTestId('find-replace-all').click();
  await expect.poll(() => blockText(page, 4)).toBe('text Some.');
  await expect(page.getByTestId('status')).toHaveText('Replaced 1 result');
  await undo(page);
  await expect.poll(() => blockText(page, 4)).toBe('Some text.');

  // Replace one: the current hit replaced, the next one current.
  await page.getByTestId('find-regexp').uncheck();
  await page.getByTestId('find-query').fill('aaa');
  await expect(page.getByTestId('find-count')).toHaveText('1 of 14');
  await page.getByTestId('find-replacement').fill('A');
  await page.getByTestId('find-replace').click();
  await expect(page.getByTestId('find-count')).toHaveText('1 of 13');
  expect(await blockText(page, 19)).toContain('Hanging indent A bbb aaa');

  // Replace all, then one undo step takes all of it back.
  // Found from the selection, which the replace left on the second "aaa", after the first "bbb".
  await page.getByTestId('find-query').fill('bbb');
  await expect(page.getByTestId('find-count')).toHaveText('2 of 14');
  await page.getByTestId('find-replacement').fill('B');
  await page.getByTestId('find-replace-all').click();
  await expect(page.getByTestId('status')).toHaveText('Replaced 14 results');
  await expect(page.getByTestId('find-count')).toHaveText('No results');
  await undo(page);
  await expect(page.getByTestId('find-count')).toHaveText('14 results');
  expect((await bodyText(page)).match(/bbb/g)?.length).toBe(14);

  // Escape: the bar closes and the document has the focus with the current hit selected.
  await page.getByTestId('find-query').fill('Justified');
  await expect(page.getByTestId('find-count')).toHaveText('1 of 1');
  await page.getByTestId('find-query').press('Escape');
  await expect(page.getByTestId('find-bar')).toHaveCount(0);
  await expect(page.locator('.ProseMirror[aria-label="Document"]')).toBeFocused();
  expect(await selectedText(page)).toBe('Justified');
  await expect(page.getByTestId('editor').locator('.find-hit')).toHaveCount(0);

  // The options are remembered; the selection's text fills the query.
  await page.keyboard.press('Control+f');
  await expect(page.getByTestId('find-query')).toHaveValue('Justified');
  await expect(page.getByTestId('find-regexp')).not.toBeChecked();
});

test('a wildcard across a paragraph mark and ^p: the count, the selection into the next paragraph, Replace and Replace all refused with a message (ED-008 section 8.8 item 11)', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await openFind(page, true);
  await page.getByTestId('find-wildcards').check();
  // "?^13?": a character, a paragraph mark, a character: one hit per mark between two paragraphs with text at the join.
  await page.getByTestId('find-query').fill('?^13?');
  await expect(page.getByTestId('find-count')).toHaveText(/^1 of \d+$/);
  expect(await selectedText(page)).toMatch(/^..$/);   // a character of each paragraph: the selection runs into the next one
  expect(await selectionBlocks(page)).toEqual([0, 1]);
  // Replace on it is refused with a message and changes nothing; Replace all too.
  const before = await bodyText(page);
  await page.getByTestId('find-replacement').fill('X');
  await page.getByTestId('find-replace').click();
  await expect(page.getByTestId('status')).toContainText('holds a paragraph mark');
  expect(await bodyText(page)).toBe(before);
  await page.getByTestId('find-replace-all').click();
  expect(await bodyText(page)).toBe(before);
  // A plain query's ^p: the marks, each selected across into the next paragraph.
  await page.getByTestId('find-wildcards').uncheck();
  await page.getByTestId('find-query').fill('text.^p');
  await expect.poll(() => selectedText(page)).toBe('text.');
  await expect(page.getByTestId('find-count')).toHaveText(/^1 of \d+$/);
  expect(await selectionBlocks(page)).toEqual([0, 1]);
  await page.getByTestId('find-replace').click();
  await expect(page.getByTestId('status')).toContainText('holds a paragraph mark');
  expect(await bodyText(page)).toBe(before);
  // Typing in the document with the bar open keeps the count: the hits are searched again across the marks.
  const previous = (await page.getByTestId('find-count').textContent())!;
  await page.getByTestId('find-query').fill('^p');
  await expect(page.getByTestId('find-count')).not.toHaveText(previous);
  const count = await page.getByTestId('find-count').textContent();
  await caretInBlock(page, 4, 2);
  await page.keyboard.type('zz');
  await expect(page.getByTestId('find-count')).toHaveText(count!.replace(/^1 of/, '2 of'));
});

test('Next while the query field holds the focus scrolls to the hit: ProseMirror scrolls only to a DOM selection inside the editor', async ({ page }) => {
  await page.setViewportSize({ width: 1000, height: 420 });
  await openFixture(page, 'tracked-changes.docx');
  await openFind(page);
  await page.getByTestId('find-query').fill('indented');
  await expect(page.getByTestId('find-count')).toHaveText('1 of 20');
  const area = page.locator('.page-area');
  await area.evaluate((el) => { el.scrollTop = 0; });
  // The last hit, far below the window, with no click in the document before it: the DOM
  // selection is still the field's, so ProseMirror's own scroll does nothing.
  await page.getByTestId('find-previous').click();
  await expect(page.getByTestId('find-count')).toHaveText('20 of 20');
  const inView = (): Promise<boolean> => page.evaluate(() => {
    const hit = document.querySelector('.find-hit.current')?.getBoundingClientRect();
    const box = document.querySelector('.page-area')!.getBoundingClientRect();
    return !!hit && hit.top >= box.top && hit.bottom <= box.bottom;
  });
  await expect.poll(inView).toBe(true);
  expect(await area.evaluate((el) => el.scrollTop)).toBeGreaterThan(0);
  // And Next from there wraps to the first, back at the top.
  await page.getByTestId('find-next').click();
  await expect(page.getByTestId('find-count')).toHaveText('1 of 20');
  await expect.poll(inView).toBe(true);
});

test('a Hebrew query in a right-to-left document; Arabic whole words', async ({ page }) => {
  await openFixture(page, 'fonts-hebrew-no-cs.docx');
  await openFind(page);
  await expect(page.getByTestId('find-query')).toHaveAttribute('dir', 'auto');
  await page.getByTestId('find-query').fill('שלום');
  await expect(page.getByTestId('find-count')).toHaveText('1 of 3');
  await expect(page.getByTestId('editor').locator('.find-hit')).toHaveCount(3);
  expect(await selectedText(page)).toBe('שלום');
  await page.getByTestId('find-whole-words').check();
  await page.getByTestId('find-query').fill('نص تجريبي');
  await expect(page.getByTestId('find-count')).toHaveText('1 of 2');
  await page.getByTestId('find-whole-words').uncheck();
});

test('the notes and the headers after the body, in Word\'s order; a hit in a header selects it in its view', async ({ page }) => {
  await openFixture(page, 'headers.docx');
  await openFind(page);
  await page.getByTestId('find-query').fill('header');
  // Body: "Header and footer test"; headers: six with "header" in four ("First page header." twice, "Normal header.", "Just a plain header.").
  await expect(page.getByTestId('find-count')).toHaveText('1 of 5');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('find-count')).toHaveText('2 of 5');
  const active = await page.evaluate(() => {
    const surfaces = (window as unknown as { docx4jSurfaces: Map<string, { id: string; state: { doc: { textBetween(a: number, b: number): string }; selection: { from: number; to: number } }; view: { dom: Element } }> }).docx4jSurfaces;
    const withSelection = [...surfaces.values()].filter((s) => s.id !== 'body' && s.state.selection.to > s.state.selection.from);
    return withSelection.map((s) => ({ id: s.id, text: s.state.doc.textBetween(s.state.selection.from, s.state.selection.to), current: s.view.dom.querySelectorAll('.find-hit.current').length }));
  });
  expect(active.length).toBe(1);
  expect(active[0]!.id).toMatch(/^header:/);
  expect(active[0]!.text).toBe('header');
  expect(active[0]!.current).toBe(1);
  // The body's own hit is no longer current: one current hit on the page, the header's.
  await expect(page.locator('.find-hit.current')).toHaveCount(1);

  // Replace all reaches the headers too, and one undo takes it all back.
  await page.getByTestId('find-toggle-replace').click();
  await page.getByTestId('find-replacement').fill('HEADING');
  await page.getByTestId('find-replace-all').click();
  await expect(page.getByTestId('status')).toHaveText('Replaced 5 results');
  const headerTexts = (): Promise<string[]> => page.evaluate(async () => {
    const editor = (window as unknown as { docx4jEditor: { session: { stories: { headers(): Promise<{ key: string }[]>; currentDoc(s: unknown): { textContent: string } } } } }).docx4jEditor;
    return (await editor.session.stories.headers()).map((h) => editor.session.stories.currentDoc(h).textContent);
  });
  expect((await headerTexts()).filter((t) => t.includes('HEADING')).length).toBe(4);
  await undo(page);
  await expect.poll(async () => (await headerTexts()).filter((t) => t.includes('HEADING')).length).toBe(0);
  expect(await blockText(page, 0)).toMatch(/^Header and footer test/);
});

test('a hit in a footnote: the note\'s view is the active surface with the hit selected', async ({ page }) => {
  await openFixture(page, 'notes-numbering.docx');
  await openFind(page);
  await page.getByTestId('find-query').fill('Footnote 3');
  // None in the body: the count says how many, and Enter goes to the note.
  await expect(page.getByTestId('find-count')).toHaveText('1 result');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('find-count')).toHaveText('1 of 1');
  const note = page.getByTestId('footnotes').locator('li.note').nth(2);
  await expect(note.locator('.find-hit.current')).toHaveCount(1);
  const id = await note.getAttribute('data-note');
  expect(await selectedText(page, id!)).toBe('Footnote 3');
  await page.getByTestId('find-query').press('Escape');
  await expect(note.locator('.ProseMirror')).toBeFocused();
});

/** A context whose pages have the relay option set to the preview server's mount. */
async function withRelay(context: BrowserContext): Promise<void> {
  await context.addInitScript(() => { try { localStorage.setItem('docx4j-editor.relay', '/relay'); } catch { /* ignore */ } });
}

test('two pages in a room: one replaces all while the other types, and both converge', async ({ browser }) => {
  const contextA = await browser.newContext();
  const contextB = await browser.newContext();
  await withRelay(contextA);
  await withRelay(contextB);
  const a = await contextA.newPage();
  const b = await contextB.newPage();
  await openFixture(a, 'tracked-changes.docx');
  await a.getByTestId('share').click();
  const link = await a.getByTestId('share-link').inputValue();
  await a.keyboard.press('Escape');
  await b.goto(link);
  await b.getByTestId('join-accept').click();
  await expect(b.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 20_000 });

  await caretInBlock(a, 1);
  const typing = a.keyboard.type('Typing while they replace. ', { delay: 20 });
  await openFind(b, true);
  await b.getByTestId('find-query').fill('indented');
  await expect(b.getByTestId('find-count')).toHaveText(/of 20$/);
  await b.getByTestId('find-replacement').fill('moved');
  await b.getByTestId('find-replace-all').click();
  await typing;
  await expect(b.getByTestId('status')).toHaveText('Replaced 20 results');
  for (const page of [a, b]) {
    await expect.poll(() => blockText(page, 1)).toMatch(/^Typing while they replace\. /);
    await expect.poll(async () => (await bodyText(page)).match(/indented/gi)?.length ?? 0).toBe(0);
    await expect.poll(async () => (await bodyText(page)).match(/moved/g)?.length ?? 0).toBe(20);
  }
  expect(await bodyText(a)).toBe(await bodyText(b));
  await contextA.close();
  await contextB.close();
});

// ED-005 section 12.108: a new term typed in the box is searched, whichever way the search is asked for.
test('a new term typed over the old one is searched: Ctrl+F or Ctrl+H from the box keeps it, and Enter straight after typing compiles it', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await openFind(page);
  await page.getByTestId('find-query').fill('indented');
  await expect(page.getByTestId('find-count')).toHaveText('1 of 20');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('find-count')).toHaveText('2 of 20');
  // Ctrl+F from the box with the old hit still selected in the document: the typed term stays.
  await page.keyboard.press('Control+a');
  await page.keyboard.type('Justified', { delay: 20 });
  await page.keyboard.press('Control+f');
  await expect(page.getByTestId('find-query')).toHaveValue('Justified');
  await expect(page.getByTestId('find-count')).toHaveText('1 of 1');
  expect(await selectedText(page)).toBe('Justified');
  // Ctrl+H from the box the same, the replacement row opening.
  await page.keyboard.press('Control+a');
  await page.keyboard.type('indented', { delay: 20 });
  await page.keyboard.press('Control+h');
  await expect(page.getByTestId('find-replacement')).toBeVisible();
  await expect(page.getByTestId('find-query')).toHaveValue('indented');
  await expect(page.getByTestId('find-count')).toHaveText('1 of 20');
  // Enter straight after typing: the new term's first hit, never the old term's next.
  await page.getByTestId('find-query').focus();
  await page.keyboard.press('Control+a');
  await page.keyboard.type('Justified', { delay: 20 });
  await page.keyboard.press('Enter');
  expect(await selectedText(page)).toBe('Justified');
  await expect(page.getByTestId('find-count')).toHaveText('1 of 1');
  // From the document, Ctrl+F still brings the selection's text, as before.
  await selectBlock(page, 0);
  await page.keyboard.press('Control+f');
  await expect(page.getByTestId('find-query')).toHaveValue(await blockText(page, 0));
});
