// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 12.5 item 3: the interface in French and in Arabic, chosen in Options. French:
// the toolbar's Bold, the Macros menu, the palette's Alignment, the paste report in the status
// line, the choice kept across a reload; the document's style names as the document has them.
// Arabic: the chrome right-to-left (`dir` and `lang` on the root, the toolbar starting at the
// right), the document's paragraphs still left-to-right, the side panel's splitter and tab strip
// following the direction. Then back to English. Expectations are read from the message files,
// so a review that rewords a message leaves the test standing.
import { test, expect, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openFixture, caretInBlock, capture } from './helpers.js';

const locales = join(dirname(fileURLToPath(import.meta.url)), '..', 'src', 'locales');
const messages = async (tag: string): Promise<Record<string, string>> => JSON.parse(await readFile(join(locales, `${tag}.json`), 'utf8')) as Record<string, string>;

async function chooseLanguage(page: Page, tag: string): Promise<void> {
  const options = page.locator('details.options-menu');
  if (!(await options.evaluate((d) => (d as HTMLDetailsElement).open))) await options.locator('summary').click();
  await page.getByTestId('option-language').selectOption(tag);
  if (tag !== '') await expect(page.locator('.app')).toHaveAttribute('lang', tag);
  await options.locator('summary').click();
}

/** Paste a capture of the corpus through a clipboard event, as paste.spec.ts does. */
async function pasteCapture(page: Page, name: string): Promise<void> {
  const html = await capture(name);
  await page.evaluate((h) => {
    const dt = new DataTransfer();
    dt.setData('text/html', h);
    dt.setData('text/plain', 'plain');
    document.querySelector('.ProseMirror[role="textbox"]')!.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  }, html);
}

test('French: the toolbar, the menus, the palette and the status line in French; the document\'s style names as they are; kept across a reload', async ({ page }) => {
  const fr = await messages('fr');
  await openFixture(page, 'tracked-changes.docx');
  const styleNames = await page.getByTestId('style-picker').locator('option').allTextContents();
  await chooseLanguage(page, 'fr');
  await expect(page.locator('html')).toHaveAttribute('lang', 'fr');
  await expect(page.locator('.app')).toHaveAttribute('dir', 'ltr');

  await expect(page.locator('button[data-command="font.bold"]')).toHaveAttribute('title', fr['toolbar.bold']!);
  await expect(page.getByTestId('macros-menu')).toHaveText(fr['macros.menu']!);
  await expect(page.getByTestId('new')).toHaveText(fr['app.new']!);
  // The draft notice under the language selector.
  await page.locator('details.options-menu summary').click();
  await expect(page.getByTestId('language-draft')).toHaveText(fr['options.language.draft']!);
  await page.locator('details.options-menu summary').click();

  // The palette's labels are the commands' in French; the name stays the API's.
  await page.getByTestId('palette-button').click();
  const alignment = page.getByTestId('palette').locator('li[data-command="paragraph.alignment"]');
  await expect(alignment.locator('.label')).toHaveText(fr['command.paragraph.alignment']!);
  await expect(alignment.locator('.name')).toHaveText('paragraph.alignment');
  await page.keyboard.press('Escape');

  // The document's words are the document's: the style picker lists the same names.
  expect(await page.getByTestId('style-picker').locator('option').allTextContents()).toEqual(styleNames);

  // The paste report: 9 spacing values dropped, as paste.spec.ts counts them, said in French.
  await caretInBlock(page, 4);
  await pasteCapture(page, 'all');
  await expect(page.getByTestId('status')).toContainText(fr['paste.dropped.spacing']!.replace('{count}', '9'));

  // Kept across a reload, before the first render.
  await page.reload();
  await expect(page.locator('.app')).toHaveAttribute('lang', 'fr');
  await expect(page.getByTestId('new')).toHaveText(fr['app.new']!);
});

test('Arabic: the chrome right-to-left, the document left-to-right; the splitter and the tab strip follow the direction; then English again', async ({ page }) => {
  const ar = await messages('ar');
  await openFixture(page, 'tracked-changes.docx');
  await chooseLanguage(page, 'ar');
  const app = page.locator('.app');
  await expect(app).toHaveAttribute('dir', 'rtl');
  await expect(page.locator('html')).toHaveAttribute('lang', 'ar');
  // A pane hangs the other way from its button here, and stays inside the window just the same (ED-005 section 12.112).
  await page.keyboard.press('Escape');
  for (const width of [1280, 700]) {
    await page.setViewportSize({ width, height: 720 });
    for (let i = 0; i < 3; i++) {
      const menu = page.locator('.panels-menu').nth(i);
      if (await menu.evaluate((el) => (el as HTMLDetailsElement).open)) await menu.locator('> summary').click();
      await menu.locator('> summary').click();
      await expect.poll(() => menu.locator('> .menu-pane').evaluate((el) => {
        const b = el.getBoundingClientRect();
        return b.left >= 0 && b.right <= document.documentElement.clientWidth;
      }), `menu ${i} at ${width}`).toBe(true);
      await menu.locator('> summary').click();
    }
  }
  await page.setViewportSize({ width: 1280, height: 720 });
  await expect(page.locator('button[data-command="font.bold"]')).toHaveAttribute('title', ar['toolbar.bold']!);

  // The toolbar starts at the right: its first control (Undo) is nearer its right edge than its left.
  const toolbar = (await page.getByTestId('toolbar').boundingBox())!;
  const undo = (await page.locator('button[data-command="history.undo"]').boundingBox())!;
  expect(toolbar.x + toolbar.width - (undo.x + undo.width)).toBeLessThan(undo.x - toolbar.x);

  // The document keeps its own direction: the editor is ltr, and a paragraph without w:bidi reads left-to-right.
  const editor = page.locator('.docx-editor.body > .ProseMirror');
  await expect(editor).toHaveAttribute('dir', 'ltr');
  expect(await editor.locator('p').first().evaluate((p) => getComputedStyle(p).direction)).toBe('ltr');

  // The Developer perspective's side panel sits on the left; the arrow that moves the splitter towards the page (Right) widens it.
  await page.getByTestId('perspective-developer').click();
  const panels = page.getByTestId('panels');
  await expect(panels).toBeVisible();
  const page_ = (await page.locator('.page-area').boundingBox())!;
  const panel = (await panels.boundingBox())!;
  expect(panel.x).toBeLessThan(page_.x);
  const splitter = page.getByTestId('splitter');
  const before = Number(await splitter.getAttribute('aria-valuenow'));
  await splitter.focus();
  await page.keyboard.press('ArrowRight');
  await expect(splitter).toHaveAttribute('aria-valuenow', String(before + 16));
  await page.keyboard.press('ArrowLeft');
  await expect(splitter).toHaveAttribute('aria-valuenow', String(before));
  // Dragging the splitter towards the page widens the panel.
  const box = (await splitter.boundingBox())!;
  await page.mouse.move(box.x + box.width / 2, box.y + 40);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width / 2 + 60, box.y + 40, { steps: 4 });
  await page.mouse.up();
  expect(Number(await splitter.getAttribute('aria-valuenow'))).toBeGreaterThan(before + 40);

  // The tab strip: the arrow that points along the reading direction (left) goes to the next tab.
  const first = page.getByTestId('tab-codes');
  await first.focus();
  await page.keyboard.press('ArrowLeft');
  await expect(page.getByTestId('tab-format')).toHaveAttribute('aria-selected', 'true');
  await page.keyboard.press('ArrowRight');
  await expect(first).toHaveAttribute('aria-selected', 'true');

  // Code stays left-to-right in a right-to-left chrome: the reveal-codes XML.
  await caretInBlock(page, 4);
  const xml = page.getByTestId('panel-xml').getByTestId('xml-editor');
  await expect(xml).toBeVisible();
  expect(await xml.evaluate((el) => getComputedStyle(el).direction)).toBe('ltr');

  // Back to the browser's language, English here.
  await chooseLanguage(page, '');
  await expect(app).toHaveAttribute('dir', 'ltr');
  await expect(app).toHaveAttribute('lang', 'en');
  await expect(page.locator('button[data-command="font.bold"]')).toHaveAttribute('title', 'Bold');
});

// Proposal 43 (ED-005 section 9.7): a change's values in the review panel, in French as Word's pane gives them there.
test('French: the review panel\'s values ("Police : Gras", a size in points), from the message file', async ({ page }) => {
  const fr = await messages('fr');
  await openFixture(page, 'review/review-for-word.docx');
  await chooseLanguage(page, 'fr');
  await page.getByTestId('perspective-review').click();
  await expect(page.getByTestId('panel-review')).toBeVisible();
  const font = fr['toolbar.font']!;
  const named = (property: string, value: string): string => fr['review.value.named']!.replace('{property}', property).replace('{value}', value);
  const bold = page.locator('[data-testid="review-entry"][data-type="Formatted"]').filter({ hasText: 'alpha' });
  await expect(bold.locator('.what')).toHaveText(named(font, fr['command.font.bold']!));
  const size = page.locator('[data-testid="review-entry"][data-type="Formatted"]').filter({ hasText: 'bravo' });
  await expect(size.locator('.what')).toHaveText(named(font, fr['review.value.points']!.replace('{size}', '16')));
  await chooseLanguage(page, '');
});

// ED-005 section 9.6 (E4.g step G6): the Review perspective and the comments margin in French and Arabic, the words
// from the message files; in Arabic the panel follows the chrome's direction while the document keeps its own.
test('French and Arabic: the Review perspective and the comments margin translated; in Arabic the panel runs right to left and the document keeps its own direction', async ({ page }) => {
  await openFixture(page, 'review/review-for-word.docx');
  for (const tag of ['fr', 'ar'] as const) {
    const m = await messages(tag);
    await chooseLanguage(page, tag);
    await page.getByTestId('perspective-review').click();
    const panel = page.getByTestId('panel-review');
    await expect(panel).toBeVisible();
    await expect(page.getByTestId('perspective-review')).toHaveText(m['perspective.review']!);
    await expect(panel.locator('h2').first()).toHaveText(m['review.panel']!);
    await expect(panel.getByTestId('review-accept-all')).toHaveText(m['review.acceptAll']!);
    await expect(panel.getByTestId('review-accept').first()).toHaveText(m['review.accept']!);
    await expect(panel.locator('[data-testid="review-entry"][data-type="Replaced"] .type').first()).toHaveText(m['review.type.replaced']!);
    const margin = page.getByTestId('comments-margin');
    await expect(margin).toHaveAttribute('aria-label', m['comments.margin']!);
    await expect(margin.getByTestId('comment-reply').first()).toHaveText(m['comments.reply']!);
    const direction = (selector: string): Promise<string> => page.locator(selector).first().evaluate((el) => getComputedStyle(el).direction);
    expect(await direction('[data-testid="panel-review"]')).toBe(tag === 'ar' ? 'rtl' : 'ltr');
    expect(await direction('.docx-editor .ProseMirror[aria-label]')).toBe('ltr');
  }
  await chooseLanguage(page, '');
});
