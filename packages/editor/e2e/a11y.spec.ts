// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-001 section 11 item 9, as the checklist ED-002 section 5.4 asks for: the editor region and
// the panels labelled, toolbar buttons named with a pressed state, the palette a listbox with
// keyboard navigation, every command reachable without a mouse, focus visible, and the
// benchmark marks in place for E1.e. (No axe: its licence is MPL, which the dependency rule
// keeps out even of the tools.)
import { test, expect } from '@playwright/test';
import { openFixture, caretInBlock, xmlEditorValue, setXmlEditor } from './helpers.js';

test('regions, names and states', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await expect(page.getByRole('textbox', { name: 'Document' })).toBeVisible();
  await expect(page.getByRole('toolbar', { name: 'Formatting' })).toBeVisible();
  await expect(page.getByRole('toolbar', { name: 'File' })).toBeVisible();
  await expect(page.getByRole('tablist', { name: 'Perspective' })).toBeVisible();
  const buttons = page.getByRole('toolbar', { name: 'Formatting' }).locator('button');
  const count = await buttons.count();
  expect(count).toBeGreaterThan(15);
  for (let i = 0; i < count; i++) {
    const b = buttons.nth(i);
    expect((await b.getAttribute('aria-label')) ?? (await b.textContent())?.trim(), `button ${i} has a name`).toBeTruthy();
  }
  const toggles = page.getByRole('toolbar', { name: 'Formatting' }).locator('button[aria-pressed]');
  expect(await toggles.count()).toBeGreaterThan(8);
  for (const select of await page.getByRole('toolbar', { name: 'Formatting' }).locator('select').all()) {
    expect(await select.getAttribute('aria-label')).toBeTruthy();
  }
  await page.getByTestId('perspective-developer').click();
  await expect(page.getByRole('region', { name: 'Reveal codes' })).toBeVisible();
  await expect(page.getByRole('complementary', { name: 'Panels' })).toBeVisible();
  await expect(page.getByRole('tablist', { name: 'Panel tabs' })).toBeVisible();
  await page.getByRole('tab', { name: 'Diagnostics' }).click();
  await expect(page.getByRole('region', { name: 'Diagnostics' })).toBeVisible();
  await expect(page.getByRole('separator', { name: 'Resize the panel' })).toBeVisible();
});

test('the Review perspective (ED-005 section 6.7): the panel named, its buttons named, the views a named group, the tracking switch a toggle', async ({ page }) => {
  await openFixture(page, 'revisions/revisions-word15.docx');
  await page.getByTestId('perspective-review').click();
  await expect(page.getByRole('region', { name: 'Tracked changes' })).toBeVisible();
  await expect(page.getByRole('group', { name: 'Show' })).toBeVisible();
  await expect(page.getByRole('radio', { name: 'As accepted' })).toBeVisible();
  const toggle = page.getByTestId('tracking-switch');
  await expect(toggle).toHaveAttribute('aria-pressed', /true|false/);
  expect((await toggle.textContent())?.trim()).toBe('Track changes');
  for (const b of await page.getByTestId('panel-review').locator('button').all()) {
    expect(((await b.getAttribute('aria-label')) ?? (await b.textContent()))?.trim(), 'a named button').toBeTruthy();
  }
  // The current entry is marked for assistive technology.
  await page.getByTestId('review-next').click();
  await expect(page.locator('[data-testid="review-entry"][aria-current="true"]')).toHaveCount(1);
});

test('the palette is a listbox with keyboard navigation and reaches every command', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await page.getByTestId('palette-button').click();
  const palette = page.getByTestId('palette');
  await expect(palette.getByRole('listbox', { name: 'Commands' })).toBeVisible();
  const options = palette.getByRole('option');
  const listed = await options.count();
  const registered = await page.evaluate(() => (window as unknown as { docx4jEditor: { session: unknown } }).docx4jEditor ? document.querySelectorAll('[data-testid="palette"] [role="option"]').length : 0);
  expect(listed).toBe(registered);
  expect(listed).toBeGreaterThan(40);
  await expect(options.first()).toHaveAttribute('aria-selected', 'true');
  await page.keyboard.press('ArrowDown');
  await expect(options.nth(1)).toHaveAttribute('aria-selected', 'true');
  await page.keyboard.press('Escape');
  await expect(palette).toBeHidden();
  await expect(page.getByRole('textbox', { name: 'Document' })).toBeFocused();
});

test('focus is visible on the toolbar and the editor is reachable by Tab', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  const bold = page.getByRole('button', { name: 'Bold' });
  await bold.focus();
  const outline = await bold.evaluate((el) => getComputedStyle(el).outlineStyle);
  expect(outline).not.toBe('none');
  await page.keyboard.press('Tab');
  await expect(page.locator(':focus')).not.toHaveCount(0);
});

test('the Template perspective (ED-003 section 8.6): tabs and trees named, the tree walked by keys, a node bound from the keyboard, the Macros menu and the ask box named and reachable', async ({ page }) => {
  await openFixture(page, 'invoice.docx');
  await page.getByTestId('offer-template').click();
  const tabs = page.getByRole('tablist', { name: 'Panel tabs' });
  await expect(tabs.getByRole('tab')).toHaveText(['Control', 'Controls', 'Data', 'OpenDoPE', 'Preview']);
  await expect(page.getByRole('region', { name: 'Control' })).toBeVisible();
  // Every control in the toolbar's group is named.
  const group = page.getByRole('group', { name: 'Content controls' });
  await expect(group).toBeVisible();
  for (const c of await group.locator('button, select').all()) expect(await c.getAttribute('aria-label')).toBeTruthy();
  // The Controls tab: a named list; the Control tab's groups fold from the keyboard.
  await page.getByRole('tab', { name: 'Controls' }).click();
  await expect(page.getByRole('list', { name: 'Controls in the body' })).toBeVisible();
  await page.getByRole('tab', { name: 'Control', exact: true }).click();
  await caretInBlock(page, 3, 2);   // the "Customer name" control
  const binding = page.getByTestId('control-binding');
  await binding.locator('> details > summary').focus();
  await page.keyboard.press('Enter');
  await expect(binding.locator('> details')).not.toHaveAttribute('open', '');
  await expect(page.getByTestId('bind-xpath')).toBeHidden();
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('bind-xpath')).toBeVisible();
  await expect(page.getByTestId('bind-apply')).toHaveText('Rebind');   // the drag's alternative: the XPath box and its button (this control is bound already)
  // The Data tab: the data part is a named tree the arrow keys walk.
  await page.getByRole('tab', { name: 'Data' }).click();
  const tree = page.getByRole('tree', { name: /invoice/ });
  await expect(tree).toBeVisible();
  const name = (xpath: string) => page.locator(`[data-xpath="${xpath}"] > .tree-row > .node-name`);
  await name('/invoice[1]').focus();
  await page.keyboard.press('ArrowDown');
  await expect(name('/invoice[1]/customer[1]')).toBeFocused();
  await page.keyboard.press('ArrowLeft');
  await expect(page.locator('[data-xpath="/invoice[1]/customer[1]"]')).toHaveAttribute('aria-expanded', 'false');
  await page.keyboard.press('ArrowRight');
  await expect(page.locator('[data-xpath="/invoice[1]/customer[1]"]')).toHaveAttribute('aria-expanded', 'true');
  await page.keyboard.press('ArrowRight');
  await expect(name('/invoice[1]/customer[1]/name[1]')).toBeFocused();
  await expect(page.locator('[data-xpath="/invoice[1]/customer[1]/name[1]"]')).toHaveAttribute('aria-label', 'name: Joe Bloggs');
  await page.keyboard.press('End');
  await expect(name('/invoice[1]/misc[1]/wantspam[1]')).toBeFocused();
  // Bind from the keyboard: the node's bind button, reached by Tab from its name, binds the control at the caret.
  await name('/invoice[1]/items[1]/item[1]/price[1]').focus();
  const bind = page.getByRole('button', { name: 'Bind to /invoice[1]/items[1]/item[1]/price[1]: the control at the caret, or a new plain text control' });
  await bind.focus();
  await expect(bind).toBeFocused();
  await page.keyboard.press('Enter');
  await expect(page.getByTestId('status')).toHaveText('Bound to /invoice[1]/items[1]/item[1]/price[1]');
  await page.getByRole('tab', { name: 'Control', exact: true }).click();
  await expect(page.getByTestId('cc-binding-selects')).toContainText('$20');
  // The Macros menu is named and opens from the keyboard; the ask box is a named region with a named question box.
  const macros = page.locator('summary[aria-label="Macros"]');
  await macros.focus();
  await page.keyboard.press('Enter');
  await expect(page.locator('.macros-menu > .menu-pane')).toBeVisible();
  await expect(page.locator('.macros-menu > .menu-pane')).toContainText('None yet');   // the list is there, empty for this document
  await page.keyboard.press('Enter');
  await expect(page.locator('.macros-menu > .menu-pane')).toBeHidden();
  await page.getByTestId('perspective-developer').click();
  await expect(page.getByRole('region', { name: 'Ask' })).toBeVisible();
  const question = page.getByRole('textbox', { name: 'Question' });
  await question.focus();
  await expect(question).toBeFocused();
});

test('the benchmark marks are recorded at open and ready', async ({ page }) => {
  await openFixture(page, 'headers-footers.docx');
  const marks = await page.evaluate(() => performance.getEntriesByType('mark').map((m) => m.name).filter((n) => n.startsWith('docx4j:')));
  expect(marks).toContain('docx4j:open');
  expect(marks).toContain('docx4j:ready');
});

// ED-005 section 9.5 (E4.g step G4): the checklist over E4's surfaces that the tests above do not cover.
const DOC = '.docx-editor .ProseMirror[aria-label="Document"]';

test('the find bar: a named search region, its fields labelled, Escape giving the focus back to the document', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await caretInBlock(page, 4);
  await page.keyboard.press('Control+f');
  const bar = page.getByRole('search', { name: 'Find and replace' });
  await expect(bar).toBeVisible();
  await expect(bar.getByRole('textbox', { name: 'Find' })).toBeFocused();
  await expect(bar.getByRole('button', { name: 'Close' })).toBeVisible();
  for (const b of await bar.locator('button').all()) expect(((await b.getAttribute('aria-label')) ?? (await b.textContent()))?.trim(), 'a named button').toBeTruthy();
  await page.keyboard.type('text');
  await page.keyboard.press('Escape');
  await expect(bar).toBeHidden();
  await expect(page.locator(DOC)).toBeFocused();
});

test('the comments margin: each thread a group named for its author, its buttons named and reached by Tab, the body named', async ({ page }) => {
  await openFixture(page, 'comments-two.docx');
  const margin = page.getByRole('complementary', { name: 'Comments' });
  await expect(margin).toBeVisible();
  const threads = margin.getByRole('group', { name: /^Comment by / });
  expect(await threads.count()).toBeGreaterThanOrEqual(2);
  const first = threads.first();
  for (const name of ['Reply', 'Resolve', 'Delete']) await expect(first.getByRole('button', { name, exact: true }).first()).toBeVisible();
  // Tab from the thread's head reaches its buttons, in order.
  await first.locator('.comment-head').first().focus();
  const reached: string[] = [];
  for (let i = 0; i < 6; i++) {
    await page.keyboard.press('Tab');
    reached.push(await page.evaluate(() => (document.activeElement?.textContent ?? '').trim()));
  }
  expect(reached).toEqual(expect.arrayContaining(['Reply', 'Resolve', 'Delete']));
  await expect(first.locator('[data-testid="comment-body"] [aria-label^="Comment by "]').first()).toBeVisible();
});

test('the name dialog: labelled, the focus kept inside it, Escape closing it and giving the focus back', async ({ page }) => {
  await openFixture(page, 'comments-two.docx');
  await caretInBlock(page, 0, 6);
  await page.getByTestId('new-comment').focus();
  await page.getByTestId('new-comment').press('Enter');
  const dialog = page.getByRole('dialog', { name: 'Your name for comments' });
  await expect(dialog).toBeVisible();
  await expect(dialog).toHaveAttribute('aria-modal', 'true');
  await expect(page.getByTestId('tracking-name')).toBeFocused();
  // Tab and Shift+Tab go round inside the dialog.
  for (let i = 0; i < 5; i++) {
    await page.keyboard.press('Tab');
    expect(await dialog.evaluate((d) => d.contains(document.activeElement))).toBe(true);
  }
  for (let i = 0; i < 5; i++) {
    await page.keyboard.press('Shift+Tab');
    expect(await dialog.evaluate((d) => d.contains(document.activeElement))).toBe(true);
  }
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(page.getByTestId('new-comment')).toBeFocused();
});

test('the parts panel: Apply and Revert named, a refusal announced', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await page.getByTestId('perspective-developer').click();
  await page.getByTestId('tab-parts').click();
  await page.getByTestId('parts-tree').getByRole('button', { name: '/word/styles.xml', exact: true }).first().click();
  await expect(page.getByRole('button', { name: 'Apply' })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Revert' })).toBeVisible();
  const text = await xmlEditorValue(page, 'part-xml');
  await setXmlEditor(page, text.replace('<w:name w:val="heading 1"/>', '<w:name w:val="heading 1"/>\n<w:bogus/>'), 'part-xml');
  await page.getByRole('button', { name: 'Apply' }).click();
  // The problems an alert, the outcome a status: both read out as they appear.
  const editor = page.getByTestId('part-editor');
  await expect(editor.getByRole('alert')).toContainText('w:bogus');
  await expect(editor.getByRole('status').filter({ hasText: 'Not applied' })).toBeVisible();
});

test('a replacement\'s entry: its split a named toggle with its state, each Accept and Reject named for what it resolves', async ({ page }) => {
  await openFixture(page, 'review/review-for-word.docx');
  await page.getByTestId('perspective-review').click();
  const pair = page.locator('[data-testid="review-entry"][data-type="Replaced"]', { hasText: '7 days' });
  await expect(pair.getByRole('button', { name: 'Accept: Replaced “7 days” with “14 days”' })).toBeVisible();
  await expect(pair.getByRole('button', { name: 'Reject: Replaced “7 days” with “14 days”' })).toBeVisible();
  const split = pair.getByRole('button', { name: 'Split' });
  await expect(split).toHaveAttribute('aria-expanded', 'false');
  await split.click();
  const halves = page.locator('[data-testid="review-entry"][data-split]');
  await expect(halves).toHaveCount(2);
  await expect(halves.first().getByRole('button', { name: 'Accept: Deleted 7 days' })).toBeVisible();
  await expect(halves.last().getByRole('button', { name: 'Reject: Inserted 14 days' })).toBeVisible();
  const join = halves.first().getByRole('button', { name: 'Join' });
  await expect(join).toHaveAttribute('aria-expanded', 'true');
  await join.click();
  await expect(halves).toHaveCount(0);
});

test('an author\'s pattern: their tracked text and the swatch beside their name carry the same one', async ({ page }) => {
  await openFixture(page, 'revisions/revisions-word15.docx');
  await page.getByTestId('perspective-review').click();
  /** The angles of a background's hatch: the pattern, whatever its colour. */
  const angles = (image: string): string => [...image.matchAll(/repeating-linear-gradient\(\s*(-?\d+deg)/g)].map((m) => m[1]).join(',');
  for (const author of ['Author A', 'Author B']) {
    const text = await page.locator(`${DOC} ins.insertion[data-author="${author}"], ${DOC} del.deletion[data-author="${author}"]`).first().evaluate((el) => getComputedStyle(el).backgroundImage);
    const chip = await page.locator(`[data-testid="review-entry"][data-author="${author}"] .chip`).first().evaluate((el) => getComputedStyle(el).backgroundImage);
    expect(angles(text), author).not.toBe('');
    expect(angles(chip), author).toBe(angles(text));
  }
});
