// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 7.5: the ask mode against the provider double served under /ask.
import { test, expect, type Page } from '@playwright/test';
import { openFixture, caretInBlock } from './helpers.js';

async function askReady(page: Page) {
  await openFixture(page, 'tracked-changes.docx');
  // The double as the provider, a dummy key: Options.
  await page.getByLabel('Options').click();
  await page.getByTestId('option-ask-key').fill('sk-test');
  await page.getByTestId('option-ask-base-url').fill(`${new URL(page.url()).origin}/ask`);
  await page.getByLabel('Options').click();
  await page.getByTestId('perspective-developer').click();
  const box = page.getByTestId('ask-box');
  await expect(box).toBeVisible();
  return box;
}

test('the request is shown before it is sent; a wml answer is previewed and inserted after the selection, or replaces it', async ({ page }) => {
  const box = await askReady(page);
  await caretInBlock(page, 4);
  await box.getByTestId('ask-prompt').fill('Add a heading about the double');
  await box.getByTestId('ask-ask').click();
  const request = box.getByTestId('ask-request');
  await expect(request).toContainText('claude-sonnet-5');
  await expect(request).toContainText('Add a heading about the double');
  await expect(request).toContainText('Some text.');
  await expect(request).toContainText('System prompt');
  await box.getByTestId('ask-send').click();
  const answer = box.getByTestId('ask-answer');
  await expect(answer).toHaveAttribute('data-form', 'wml');
  await expect(box.getByTestId('ask-preview')).toContainText('The double answers: Add a heading about the double');
  await expect(box.getByTestId('ask-preview').locator('p')).toHaveClass(/s-Heading1/);
  const paragraphs = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p');
  await expect(paragraphs.nth(5)).not.toContainText('The double answers');
  await box.getByTestId('ask-insert').click();
  await expect(paragraphs.nth(5)).toHaveText('The double answers: Add a heading about the double');
  await expect(paragraphs.nth(5)).toHaveClass(/s-Heading1/);
  await expect(paragraphs.nth(4)).toHaveText('Some text.');
  await caretInBlock(page, 5);   // focus back in the editor: the buttons had it
  await page.keyboard.press('Control+z');
  await expect(paragraphs.nth(5)).not.toContainText('The double answers');
  // Replace selection.
  await caretInBlock(page, 4);
  await box.getByTestId('ask-replace').click();
  await expect(paragraphs.nth(4)).toHaveText('The double answers: Add a heading about the double');
  await expect(page.getByTestId('status')).toContainText('Replaced 1 block');
});

test('a script answer lands in the console editor and does not run; an explanation is shown; a refusal is said', async ({ page }) => {
  const box = await askReady(page);
  await caretInBlock(page, 4);
  await box.getByTestId('ask-send-at-once').check();
  await box.getByTestId('ask-prompt').fill('Write a script that italicises the selection');
  await box.getByTestId('ask-ask').click();
  await expect(box.getByTestId('ask-answer')).toHaveAttribute('data-form', 'script');
  await expect(box.getByTestId('ask-script-note')).toContainText('has not run');
  await expect(page.getByTestId('panel-console').locator('.console-editor .view-lines')).toContainText('p.font.italic = true');
  await expect(page.getByRole('button', { name: 'Italic' })).toHaveAttribute('aria-pressed', 'false');
  await box.getByTestId('ask-prompt').fill('Explain what w:t does');
  await box.getByTestId('ask-ask').click();
  await expect(box.getByTestId('ask-explanation')).toContainText('a w:t the run\'s text');
  await box.getByTestId('ask-prompt').fill('Please refuse this one');
  await box.getByTestId('ask-ask').click();
  await expect(box.getByTestId('ask-answer')).toHaveAttribute('data-form', 'refusal');
  await expect(box.getByTestId('ask-answer')).toContainText('declined');
  // Clear output takes the answer with it; Clear script empties the editor.
  await page.getByTestId('console-clear').click();
  await expect(box.getByTestId('ask-answer')).toHaveCount(0);
  await page.getByTestId('console-clear-script').click();
  await expect(page.getByTestId('panel-console').locator('.console-editor .view-lines')).not.toContainText('italic');
});

test('a plain-text answer is inserted at the caret through the editor', async ({ page }) => {
  const box = await askReady(page);
  await caretInBlock(page, 4, 4);
  await box.getByTestId('ask-send-at-once').check();
  await box.getByTestId('ask-prompt').fill('plain text please');
  await box.getByTestId('ask-ask').click();
  await expect(box.getByTestId('ask-text')).toContainText("The double's plain answer");
  await box.getByTestId('ask-insert').click();
  const paragraphs = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p');
  await expect(paragraphs.nth(4)).toContainText("SomeThe double's plain answer to: plain text please");
});

test('without a key, a chat instead (ED-003 section 11.25): the prompt copied as one message, the answer pasted back and used as a provider\'s would be', async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await openFixture(page, 'tracked-changes.docx');
  await page.getByTestId('perspective-developer').click();
  const box = page.getByTestId('ask-box');
  await caretInBlock(page, 4);
  // "Send at once" cannot send without a key: the request is shown, with the chat route and no Send.
  await box.getByTestId('ask-send-at-once').check();
  await box.getByTestId('ask-prompt').fill('Add a heading about pasting');
  await box.getByTestId('ask-ask').click();
  await expect(box.getByTestId('ask-no-key')).toContainText('No API key');
  await expect(box.getByTestId('ask-send')).toHaveCount(0);
  await box.getByTestId('ask-copy-prompt').click();
  await expect(page.getByTestId('status')).toContainText('Prompt copied');
  const copied = await page.evaluate(() => navigator.clipboard.readText());
  expect(copied).toContain('Answer in exactly one of these forms');   // the contract, which a chat has no system field for
  expect(copied).toContain('\n\n---\n\nAdd a heading about pasting');
  expect(copied).toContain('Some text.');
  // The answer pasted from the chat: previewed, and inserted only when asked.
  await expect(box.getByTestId('ask-use-pasted')).toBeDisabled();
  await box.getByTestId('ask-paste').fill('Here it is:\n```wml\n<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Pasted from a chat</w:t></w:r></w:p>\n```');
  await box.getByTestId('ask-use-pasted').click();
  await expect(box.getByTestId('ask-answer')).toHaveAttribute('data-form', 'wml');
  await expect(box.getByTestId('ask-preview')).toContainText('Pasted from a chat');
  await expect(page.getByTestId('status')).toContainText('nothing has changed in the document yet');
  const paragraphs = page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"] > p');
  await expect(paragraphs.nth(5)).not.toContainText('Pasted from a chat');
  await box.getByTestId('ask-insert').click();
  await expect(paragraphs.nth(5)).toHaveText('Pasted from a chat');
  await expect(paragraphs.nth(5)).toHaveClass(/s-Heading1/);
  // A script pasted back lands in the console editor, unrun.
  await box.getByTestId('ask-prompt').fill('A script, please');
  await box.getByTestId('ask-ask').click();
  await box.getByTestId('ask-paste').fill('```ts\nconst p = selection.paragraphs[0]; // pasted script\n```');
  await box.getByTestId('ask-use-pasted').click();
  await expect(box.getByTestId('ask-script-note')).toContainText('has not run');
  await expect(page.getByTestId('panel-console').locator('.console-editor .view-lines')).toContainText('pasted script');
});
