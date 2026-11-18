// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 5.5 test 3: the banner, the palette, the toolbar and the style picker.
import { test, expect } from '@playwright/test';
import { openFixture, download, caretInBlock, caretAtBlockEnd, selectBlock, xmlEditorValue } from './helpers.js';

test('the banner lists what is read-only, and a click selects it', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const banner = page.getByTestId('banner');
  await expect(banner).toContainText('shown read-only');
  // Since E4.c the revisions are content, counted as tracked changes (ED-005 section 12.3 item 4).
  await expect(banner).toContainText('2 tracked changes');
  await page.getByTestId('banner-details').click();
  const list = page.getByTestId('banner-list');
  // Grouped by kind (ED-003 section 8.6): tables are editable since E2.a; the bookmarks stay opaque, one row per kind.
  await expect(list.locator('details[data-kind="bookmarkStart"] > summary')).toHaveText('bookmarkStart 1');
  await list.locator('details[data-kind="bookmarkStart"] > summary').click();
  await list.getByRole('button', { name: /^bookmarkStart at body\// }).first().click();
  // The click selects the item's node, or puts the caret at it when the node cannot be selected (an invisible marker).
  const at = await page.evaluate(() => { const s = (window as unknown as { docx4jEditor: { state: { selection: { $from: { index(d: number): number } } } } }).docx4jEditor.state; return s.selection.$from.index(0); });
  expect(at).toBe(23);
});

test('the palette runs a command with arguments', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const editor = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  await caretInBlock(page, 4);
  await page.keyboard.press('Control+Shift+p');
  const palette = page.getByTestId('palette');
  await expect(palette).toBeVisible();
  await palette.getByLabel('Command', { exact: true }).fill('Alignment');
  await page.keyboard.press('Enter');
  await palette.getByLabel('value').selectOption('Centered');   // an enumerated argument is a drop-down
  await page.keyboard.press('Enter');
  await expect(palette).toBeHidden();
  await expect(editor.locator(':scope > p').nth(4)).toHaveAttribute('style', /text-align: center/);
  await expect(page.getByRole('button', { name: 'Centre', exact: true })).toHaveAttribute('aria-pressed', 'true');
});

test('the toolbar shows bold pressed on a heading and not on a plain paragraph; the picker restyles', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const editor = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  const bold = page.getByRole('button', { name: 'Bold' });
  await caretInBlock(page, 2);   // the Heading1 paragraph
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await expect(page.getByTestId('style-picker')).toHaveValue('Heading1');
  await caretInBlock(page, 4);   // "Some text."
  await expect(bold).toHaveAttribute('aria-pressed', 'false');
  // Word writes only the styles a document uses, so pick one the fixture has.
  await page.getByTestId('style-picker').selectOption('Heading1');
  await expect(editor.locator(':scope > p').nth(4)).toHaveClass(/s-Heading1/);
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  // Ctrl+B on a selected paragraph, then the download carries it. (The selection is made through
  // the view: a Shift+End followed at once by Ctrl+B races the browser's asynchronous selection
  // report, which no person can do.)
  await selectBlock(page, 1);
  await page.keyboard.press('Control+b');
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  const saved = await download(page);
  const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(saved);
  const xml = await pkg.getMainDocumentPart().getXml();
  expect((xml.match(/w:pStyle w:val="Heading1"/g) ?? []).length).toBeGreaterThanOrEqual(3);
});

test('the highlighter: a click applies the pen, a second clears it, the arrow\'s swatches pick a colour, the save carries it', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const editor = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  const pen = page.getByRole('button', { name: 'Highlight', exact: true });
  // The arrow's menu shows the palette as swatches, five to a row, the pen's colour pressed (ED-005 section 12.131).
  const menu = page.getByTestId('highlight-menu');
  const swatch = (name: string) => menu.getByRole('button', { name, exact: true, includeHidden: true });   // a closed menu's too
  await expect(swatch('Yellow')).toHaveAttribute('aria-pressed', 'true');
  await selectBlock(page, 4);   // "Some text."
  await expect(pen).toHaveAttribute('aria-pressed', 'false');
  await pen.click();
  await expect(pen).toHaveAttribute('aria-pressed', 'true');
  const run = editor.locator(':scope > p').nth(4).locator('span').filter({ hasText: 'Some text.' }).first();
  expect(await run.evaluate((el) => getComputedStyle(el).backgroundColor)).toBe('rgb(255, 255, 0)');
  // Pressed, a click clears the highlight, as Word's button does.
  await pen.click();
  await expect(pen).toHaveAttribute('aria-pressed', 'false');
  expect(await run.evaluate((el) => getComputedStyle(el).backgroundColor)).toBe('rgba(0, 0, 0, 0)');
  // A colour picked from the arrow's menu becomes the pen and is applied at once; the menu closes.
  await menu.locator('summary').click();
  await expect(menu.locator('.menu-pane')).toBeVisible();
  expect(await swatch('Yellow').evaluate((el) => getComputedStyle(el).backgroundColor)).toBe('rgb(255, 255, 0)');
  await swatch('Green').click();
  await expect(menu.locator('.menu-pane')).toBeHidden();
  await expect(pen).toHaveAttribute('aria-pressed', 'true');
  await expect(swatch('Green')).toHaveAttribute('aria-pressed', 'true');
  // Drawn in the palette's green, the object model's table (Word's own `#00FF00` from objects 0.6.0, CR-014).
  const { highlightHexValue } = await import('@docx4j/generated-objects-ts/helpers/wml');
  const green = highlightHexValue('green')!;
  const rgb = `rgb(${[1, 3, 5].map((i) => parseInt(green.slice(i, i + 2), 16)).join(', ')})`;
  expect(await run.evaluate((el) => getComputedStyle(el).backgroundColor)).toBe(rgb);
  const saved = await download(page);
  const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
  const pkg = await WordprocessingMLPackage.load(saved);
  const xml = await pkg.getMainDocumentPart().getXml();
  expect(xml).toContain('w:highlight w:val="green"');
  // In the Review perspective's basic toolbar too (Jason, 2026-10-31: a reviewer marks text with it), the pen kept.
  await page.getByTestId('perspective-review').click();
  await expect(page.getByTestId('font-name')).toHaveCount(0);
  await expect(pen).toBeVisible();
  await expect(swatch('Green')).toHaveAttribute('aria-pressed', 'true');
  await expect(pen).toHaveAttribute('aria-pressed', 'true');
});

test('the font colour: a menu of swatches like Row shading\'s, Automatic first and More colours… last; the save carries each (ED-005 section 12.131)', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const editor = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  const menu = page.getByTestId('font-colour');
  const run = editor.locator(':scope > p').nth(4).locator('span').filter({ hasText: 'Some text.' }).first();
  const colourOf = (): Promise<string> => run.evaluate((el) => getComputedStyle(el).color);
  const savedXml = async (): Promise<string> => {
    const { WordprocessingMLPackage } = await import('@docx4j/core-ts');
    return (await WordprocessingMLPackage.load(await download(page))).getMainDocumentPart().getXml();
  };
  await selectBlock(page, 4);   // "Some text."
  await menu.locator('summary').click();
  await expect(menu.locator('.menu-pane')).toBeVisible();
  await expect(menu.locator('.menu-pane > *').first()).toHaveAttribute('data-testid', 'font-colour-auto');   // Automatic first, as in Word
  const red = menu.getByRole('button', { name: 'Red', exact: true, includeHidden: true });
  expect(await red.evaluate((el) => getComputedStyle(el).backgroundColor)).toBe('rgb(255, 0, 0)');
  await red.click();
  await expect(menu.locator('.menu-pane')).toBeHidden();
  await expect.poll(colourOf).toBe('rgb(255, 0, 0)');
  await expect(menu.locator('summary .bar')).toHaveCSS('background-color', 'rgb(255, 0, 0)');
  await expect(red).toHaveAttribute('aria-pressed', 'true');
  expect(await savedXml()).toContain('<w:color w:val="FF0000"/>');
  // More colours…: the browser's picker, applied once a colour is chosen.
  await menu.locator('summary').click();
  await page.getByTestId('font-colour-more').fill('#123456');
  await expect(menu.locator('.menu-pane')).toBeHidden();
  await expect.poll(colourOf).toBe('rgb(18, 52, 86)');
  // Automatic: Word's auto.
  await menu.locator('summary').click();
  await page.getByTestId('font-colour-auto').click();
  await expect(menu.locator('.menu-pane')).toBeHidden();
  await expect.poll(colourOf).not.toBe('rgb(18, 52, 86)');
  expect(await savedXml()).toContain('<w:color w:val="auto"/>');
  // Escape closes it with no change.
  await menu.locator('summary').click();
  await page.keyboard.press('Escape');
  await expect(menu.locator('.menu-pane')).toBeHidden();
});

test('zoom: the status bar\'s control, Ctrl+wheel and the palette scale the page alone, kept per browser; a tab stop and a comment bubble keep their places', async ({ page }) => {
  await openFixture(page, 'headers-footers.docx');
  const editor = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]');
  const pageAndMargin = page.locator('.page-and-margin');
  const level = page.getByTestId('zoom-level');
  const body = '.docx-editor .ProseMirror[aria-label="Document"]';
  const tab = page.locator(`${body} > p`).nth(3).locator('.tab').first();
  const widthOf = async (l: typeof editor): Promise<number> => (await l.boundingBox())!.width;
  const base = await widthOf(editor);
  const toolbarHeight = (await page.getByTestId('toolbar').boundingBox())!.height;
  // The tab's width is a layout size the plugin measured (tabstops.ts); the first tab ends at the default stop.
  await expect.poll(() => tab.evaluate((el) => el.style.inlineSize)).not.toBe('');
  const tabWidth = await tab.evaluate((el) => el.style.inlineSize);
  await expect(level).toHaveText('100%');
  await page.getByTestId('zoom-in').click();
  await expect(level).toHaveText('110%');
  expect(await pageAndMargin.evaluate((el) => getComputedStyle(el).zoom)).toBe('1.1');
  expect(await widthOf(editor)).toBeCloseTo(base * 1.1, 0);
  // The chrome keeps its size.
  expect((await page.getByTestId('toolbar').boundingBox())!.height).toBe(toolbarHeight);
  // Ctrl+wheel over the page steps the page's zoom, not the browser's.
  await editor.hover();
  await page.keyboard.down('Control');
  await page.mouse.wheel(0, -120);
  await page.keyboard.up('Control');
  await expect(level).toHaveText('120%');
  // The percentage returns to 100%.
  await level.click();
  await expect(level).toHaveText('100%');
  await page.getByTestId('zoom-out').click();
  await expect(level).toHaveText('90%');
  // The palette's actions.
  await caretInBlock(page, 3);
  await page.keyboard.press('Control+Shift+p');
  const palette = page.getByTestId('palette');
  await expect(palette).toBeVisible();
  await palette.getByLabel('Command', { exact: true }).fill('Zoom to 100%');
  await page.keyboard.press('Enter');
  await expect(palette).toBeHidden();
  await expect(level).toHaveText('100%');
  // Kept per browser: at 200% the tab's measured width is the same layout size, drawn twice as wide.
  await page.evaluate(() => localStorage.setItem('docx4j-editor.zoom', '2'));
  await openFixture(page, 'headers-footers.docx');
  await expect(level).toHaveText('200%');
  expect(await widthOf(editor)).toBeCloseTo(base * 2, 0);
  await expect.poll(() => tab.evaluate((el) => el.style.inlineSize)).toBe(tabWidth);
  // A comment's bubble sits at its anchor's height under the zoom (CommentsMargin divides the caret's coordinates).
  await openFixture(page, 'comments-two.docx');
  await expect(level).toHaveText('200%');
  const bubble = page.getByTestId('comment-bubble').first();
  const anchor = page.locator(`${body} .kind-commentReference`).first();
  await expect(bubble).toBeVisible();
  const bubbleTop = (await bubble.boundingBox())!.y;
  const anchorTop = (await anchor.boundingBox())!.y;
  expect(Math.abs(bubbleTop - anchorTop)).toBeLessThan(8);
  // A click in the zoomed text puts the caret where the pointer is (ProseMirror's coordinates under CSS zoom).
  const first = page.locator(`${body} > p`).first();
  await first.scrollIntoViewIfNeeded();
  await page.waitForTimeout(600);   // past the double-click interval of the clicks above
  const box = (await first.boundingBox())!;
  const x = box.x + 40;
  const y = box.y + box.height / 2;
  await page.mouse.click(x, y);
  // The view takes the click's selection on the browser's asynchronous report, so the caret is polled.
  const caretLeft = (): Promise<number> => page.evaluate(() => { const v = (window as unknown as { docx4jEditor: { view: { coordsAtPos: (p: number) => { left: number }; state: { selection: { from: number } } } } }).docx4jEditor.view; return v.coordsAtPos(v.state.selection.from).left; });
  await expect.poll(async () => Math.abs((await caretLeft()) - x)).toBeLessThan(16);
  await page.evaluate(() => localStorage.removeItem('docx4j-editor.zoom'));
});

test('a menu\'s pane stays inside the window wherever the wrapped bar puts its button, at any size (ED-005 section 12.112)', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await page.getByTestId('perspective-developer').click();
  await page.getByTestId('panel-xml').waitFor();
  const summary = (name: string) => page.locator(`.panels-menu > summary[aria-label="${name}"]`);
  const expectInside = async (name: string, label: string): Promise<void> => {
    const pane = page.locator('.panels-menu[open] > .menu-pane');
    // The pane is moved when the browser reports the toggle, a task after the click.
    await expect.poll(() => pane.evaluate((el) => {
      const b = el.getBoundingClientRect();
      return b.left >= 0 && b.top >= 0 && b.right <= document.documentElement.clientWidth && b.bottom <= window.innerHeight;
    }), `${name} at ${label}`).toBe(true);
  };
  for (const [width, height] of [[1280, 720], [1000, 720], [900, 720], [700, 500], [420, 400]] as const) {
    await page.setViewportSize({ width, height });
    for (const name of ['Panels', 'Macros', 'Options']) {
      await summary(name).click();
      await expectInside(name, `${width}x${height}`);
      await summary(name).click();
    }
  }
  // The checkbox the bug hid, at the width the suite runs at: in Developer the button starts the bar's second row.
  await page.setViewportSize({ width: 1280, height: 720 });
  await summary('Panels').click();
  await expect(page.getByTestId('option-panel')).toBeInViewport({ ratio: 1 });
  // An open menu follows its button when the window changes.
  await page.setViewportSize({ width: 700, height: 500 });
  await expectInside('Panels', 'resized to 700x500');
  await summary('Panels').click();
  // The last entry of the longest pane can be reached in a short window: the pane scrolls, nothing past the edge.
  await page.setViewportSize({ width: 900, height: 400 });
  await summary('Options').click();
  await expectInside('Options', '900x400');
  const options = page.locator('.options-menu[open] > .menu-pane');
  await options.evaluate((el) => { el.scrollTop = el.scrollHeight; });
  await expect(options.locator('input, select, button').last()).toBeInViewport();
});

test('perspectives: Developer shows the XML pane and diagnostics; Edit hides them; the choice persists', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await expect(page.getByTestId('panel-xml')).toHaveCount(0);
  await page.getByTestId('perspective-developer').click();
  await expect(page.getByTestId('panel-xml')).toBeVisible();
  await caretInBlock(page, 2);
  // The pane's text through the live Monaco instance (ED-005 section 12.99): Monaco draws the lines in view alone.
  await expect.poll(() => xmlEditorValue(page)).toContain('w:pStyle w:val="Heading1"');
  await page.getByTestId('tab-diagnostics').click();
  await expect(page.getByTestId('panel-diagnostics')).toBeVisible();
  await page.getByTestId('tab-codes').click();
  await page.reload();
  await page.getByTestId('open').setInputFiles(new URL('../../editor-model/test/fixtures/tracked-changes.docx', import.meta.url).pathname);
  await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await page.getByTestId('toolbar').waitFor();
  await expect(page.getByTestId('panel-xml')).toBeVisible();
  await page.getByTestId('perspective-edit').click();
  await expect(page.getByTestId('panel-xml')).toHaveCount(0);
});

test('the palette inserts a command in the console, without running it', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await caretInBlock(page, 4);
  await page.keyboard.press('Control+Shift+p');
  const palette = page.getByTestId('palette');
  await palette.getByLabel('Command', { exact: true }).fill('Alignment');
  await page.keyboard.press('Shift+Enter');
  await expect(palette).toBeHidden();
  // The Developer perspective's console comes to the front with the line at the caret; the paragraph is unchanged.
  await expect(page.getByTestId('perspective-developer')).toHaveAttribute('aria-selected', 'true');
  const console = page.getByTestId('panel-console');
  await expect(console.locator('.console-editor .monaco-editor')).toBeVisible();
  await expect(console.locator('.console-editor .view-lines')).toContainText("editor.run(\"paragraph.alignment\"");
  await expect(console.locator('.console-editor .view-lines')).toContainText('value: undefined');
  await expect(page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p').nth(4)).not.toHaveAttribute('style', /text-align/);
  // With the arguments given, the line is the command's portable form (ED-003 section 7.2): the content-API line, which runs anywhere core-ts does.
  await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').focus();
  await page.keyboard.press('Control+Shift+p');
  await palette.getByLabel('Command', { exact: true }).fill('Alignment');
  await expect(palette.locator('li[data-command="paragraph.alignment"] .badge')).toHaveText('API');
  await page.keyboard.press('Enter');
  await palette.getByLabel('value').selectOption('Centered');   // an enumerated argument is a drop-down
  await page.getByTestId('palette-insert-args').click();
  await expect(console.locator('.console-editor .view-lines')).toContainText('p.alignment = "Centered"');
  // An editor-only command says so and inserts the editor's form.
  await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').focus();
  await page.keyboard.press('Control+Shift+p');
  await palette.getByLabel('Command', { exact: true }).fill('Insert tab');
  await expect(palette.locator('li[data-command="insertTab"] .badge')).toHaveText('editor');
  await page.keyboard.press('Shift+Enter');
  await expect(console.locator('.console-editor .view-lines')).toContainText('editor.run("insertTab")');
  await expect(page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p').nth(4)).not.toHaveAttribute('style', /text-align/);
});

test.describe('a narrow window', () => {
  test.use({ viewport: { width: 900, height: 1400 } });

  test('the chrome wraps and never widens the window, so Page Down moves the caret and not the splitter', async ({ page }) => {
    // Jason, 2026-11-03, on a portrait monitor (ED-005 section 12.111): the app bar kept to one line and overflowed the
    // viewport, so the window could scroll sideways, and the caret reveal after Page Down scrolled the whole application,
    // the splitter with it, left and right with the caret's horizontal position.
    await openFixture(page, 'large/anon-250.docx');
    const widths = () => page.evaluate(() => ({ document: document.documentElement.scrollWidth, viewport: document.documentElement.clientWidth, app: (document.querySelector('.app') as HTMLElement).scrollWidth, appBarRight: Math.round(document.querySelector('.app-bar')!.getBoundingClientRect().right) }));
    const before = await widths();
    expect(before.document).toBe(before.viewport);
    expect(before.app).toBe(before.viewport);
    await page.getByTestId('perspective-developer').click();
    await expect(page.locator('.splitter')).toBeVisible();
    const after = await widths();
    expect(after.document).toBe(after.viewport);
    // The toolbar (wrapping already) and the app bar end inside the viewport.
    expect(after.appBarRight).toBeLessThanOrEqual(900);
    const splitterLeft = () => page.locator('.splitter').evaluate((el) => Math.round(el.getBoundingClientRect().left));
    const scrollLeft = () => page.evaluate(() => [document.documentElement.scrollLeft, (document.querySelector('.app') as HTMLElement).scrollLeft, (document.querySelector('.workspace') as HTMLElement).scrollLeft]);
    await caretAtBlockEnd(page, 3);   // a line's end, to the right of the page area's visible part
    const left = await splitterLeft();
    const tops: number[] = [];
    for (let i = 0; i < 4; i++) {
      await page.keyboard.press(i < 3 ? 'PageDown' : 'PageUp');
      await page.waitForTimeout(250);
      expect(await splitterLeft()).toBe(left);
      expect(await scrollLeft()).toEqual([0, 0, 0]);
      tops.push(await page.locator('.page-area').evaluate((el) => el.scrollTop));
    }
    // The keys did scroll the page area.
    expect(tops[1]).toBeGreaterThan(tops[0]!);
    expect(tops[3]).toBeLessThan(tops[2]!);
  });
});
