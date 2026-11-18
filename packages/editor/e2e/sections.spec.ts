// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Section dialog in Word's groups (ED-005 section 12.93): Page with its A4 and Letter buttons, Margins with
// Word's Normal and Narrow presets, Layout with the column count (a count, no unit), the section break. A preset
// fills the fields; OK runs the section commands once, and the dialog reads the values back when reopened.
import { test, expect } from '@playwright/test';
import { openFixture } from './helpers.js';

test('the dialog is grouped; Narrow and Letter fill the fields and OK applies them; columns is a count', async ({ page }) => {
  await openFixture(page, 'tracked-changes.docx');
  await page.getByRole('button', { name: 'Section settings' }).click();
  const dialog = page.getByRole('dialog');
  await expect(dialog.getByTestId('section-page').locator('legend')).toHaveText('Page');
  await expect(dialog.getByTestId('section-margins').locator('legend')).toHaveText('Margins');
  await expect(dialog.getByTestId('section-layout').locator('legend')).toHaveText('Layout');
  await expect(dialog.getByTestId('section-break').locator('legend')).toHaveText('Section break');
  // A count has no unit; a length has.
  await expect(dialog.getByTestId('section-columns')).toHaveValue('1');
  await expect(dialog.getByTestId('section-columns').locator('xpath=..').locator('.unit')).toHaveCount(0);
  await expect(dialog.getByTestId('section-margin-top').locator('xpath=..').locator('.unit')).toHaveText('pt');
  // The file is A4 with half-inch margins, so A4 and Narrow show pressed.
  await expect(dialog.getByRole('button', { name: 'A4' })).toHaveAttribute('aria-pressed', 'true');
  await expect(dialog.getByRole('button', { name: 'Narrow' })).toHaveAttribute('aria-pressed', 'true');
  await expect(dialog.getByTestId('section-margin-top')).toHaveValue('36');
  await dialog.getByRole('button', { name: 'Normal' }).click();
  await expect(dialog.getByTestId('section-margin-top')).toHaveValue('72');
  await expect(dialog.getByTestId('section-margin-left')).toHaveValue('72');
  await expect(dialog.getByRole('button', { name: 'Normal' })).toHaveAttribute('aria-pressed', 'true');
  await expect(dialog.getByRole('button', { name: 'Narrow' })).toHaveAttribute('aria-pressed', 'false');
  await dialog.getByRole('button', { name: 'Letter' }).click();
  await expect(dialog.getByTestId('section-width')).toHaveValue('612');
  await expect(dialog.getByTestId('section-height')).toHaveValue('792');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toHaveCount(0);
  // Reopened, the dialog reads the section's new values.
  await page.getByRole('button', { name: 'Section settings' }).click();
  await expect(dialog.getByTestId('section-width')).toHaveValue('612');
  await expect(dialog.getByTestId('section-margin-top')).toHaveValue('72');
  await expect(dialog.getByRole('button', { name: 'Letter' })).toHaveAttribute('aria-pressed', 'true');
  await expect(dialog.getByRole('button', { name: 'Normal' })).toHaveAttribute('aria-pressed', 'true');
  // Landscape swaps the paper's sides, and the preset still answers for them.
  await dialog.getByRole('combobox').first().selectOption('landscape');
  await dialog.getByRole('button', { name: 'A4' }).click();
  await expect(dialog.getByTestId('section-width')).toHaveValue('841.9');
  await expect(dialog.getByTestId('section-height')).toHaveValue('595.3');
  await dialog.getByRole('button', { name: 'Cancel' }).click();
});
