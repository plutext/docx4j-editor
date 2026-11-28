// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Markdown launcher (ED-014 section 4.3, P26) in a real VS Code: Open With closes its own tab and opens the conversion
// as an untitled `notes.docx`; the `.md` changed on disk reaches the banner, whose Reload converts it again, with no edits
// and with some; Save writes `notes.docx` and leaves the `.md` as it was, byte for byte.
import { test, expect } from '@playwright/test';
import { readFileSync, writeFileSync } from 'node:fs';
import { VsCode, bodyText, caretAtEndOf, watcherSettles } from './harness.js';
import { mainPart } from './zip.js';

test.describe.configure({ mode: 'serial' });

let vs: VsCode;
const MD = '# Notes\n\nA paragraph with **bold** text.\n\n- one\n- two\n';

test.beforeAll(async () => {
  vs = await VsCode.launch({ profile: 'markdown' });
  writeFileSync(vs.path('notes.md'), MD);
});

test.afterAll(async () => { await vs?.close(); });

test.afterEach(async ({}, testInfo) => {
  // A case that failed keeps what first made the page's document unsaved (ED-011's log), since the profile goes at the
  // next launch: the intermittent failure of section 9.10 called an unedited conversion unsaved.
  if (testInfo.status !== testInfo.expectedStatus) {
    for (const frame of await vs.frames()) {
      const causes = await frame.evaluate(() => localStorage.getItem('docx4j-editor.editCauses')).catch(() => null);
      if (causes) await testInfo.attach('editCauses', { body: causes, contentType: 'application/json' });
    }
  }
  expect(vs.problems, 'content security policy refusals').toEqual([]);
});

test('Open With: the launcher\'s tab closed, and an untitled notes.docx beside the .md, dirty from its open', async () => {
  await vs.exec('vscode.openWith', { $uri: vs.path('notes.md') }, 'docx4j.markdown');
  const frame = await vs.frameFor('notes.docx');
  await expect.poll(() => bodyText(frame)).toContain('A paragraph with bold text.');
  const tabs = await vs.tabs();
  expect(tabs.map((t) => t.label)).toEqual(['notes.docx']);
  expect(tabs[0]).toMatchObject({ viewType: 'docx4j.editor', isDirty: true });
  expect(tabs[0]!.uri).toBe(`untitled:${vs.path('notes.docx')}`);
});

test('the .md changed on disk: the banner, and Reload converting it again, with no edits and with some', async () => {
  const frame = await vs.frameFor('notes.docx');
  await watcherSettles();
  if (process.env['DOCX4J_SUITE_DEBUG']) console.log('before the write:', await frame.getByTestId('name').textContent(), await frame.evaluate(() => localStorage.getItem('docx4j-editor.editCauses')?.slice(-1500)));
  writeFileSync(vs.path('notes.md'), MD.replace('A paragraph', 'A changed paragraph'));
  await expect(frame.getByTestId('disk-offer')).toBeVisible();
  await frame.getByTestId('disk-reload').click();
  await expect.poll(() => bodyText(frame)).toContain('A changed paragraph');
  // With an edit: Reload after the confirmation.
  await caretAtEndOf(frame, 1);
  await vs.page.keyboard.type(' edited', { delay: 20 });
  await expect(frame.getByTestId('name')).toHaveText('notes.docx •');
  await watcherSettles();
  writeFileSync(vs.path('notes.md'), MD.replace('A paragraph', 'A second change to the paragraph'));
  await expect(frame.getByTestId('disk-offer')).toBeVisible();
  await frame.getByTestId('disk-discard').click();
  await frame.locator('[role="dialog"] button[type="submit"]').click();
  await expect.poll(() => bodyText(frame)).toContain('A second change to the paragraph');
  expect(await bodyText(frame)).not.toContain(' edited');
  await watcherSettles();
  writeFileSync(vs.path('notes.md'), MD);
  await expect(frame.getByTestId('disk-offer')).toBeVisible();
  await frame.getByTestId('disk-reload').click();
  await expect.poll(() => bodyText(frame)).toContain('A paragraph with bold text.');
});

test('Save: notes.docx written beside it through VS Code\'s dialog, the .md unchanged byte for byte', async () => {
  await vs.frameFor('notes.docx');
  void vs.exec('workbench.action.files.save').catch(() => undefined);
  const input = vs.page.locator('.quick-input-widget input:not([type="checkbox"])');
  await input.waitFor();
  // VS Code offers the untitled document's own path (section 4.3).
  expect(await input.inputValue()).toBe(vs.path('notes.docx'));
  await vs.page.keyboard.press('Enter');
  // The untitled editor and the file share a label: settled when the one tab of that name is the file's. Before 1.147
  // VS Code asks about the untitled one first (ED-014 section 9.10).
  await vs.afterSaveAs('notes.docx', async () => {
    const tabs = (await vs.tabs()).filter((t) => t.label === 'notes.docx');
    return tabs.length === 1 && tabs[0]!.uri === `file://${vs.path('notes.docx')}`;
  });
  await vs.frameFor('notes.docx');
  expect((await vs.tab('notes.docx'))?.isDirty).toBe(false);
  expect(mainPart(vs.path('notes.docx'))).toContain('A paragraph with ');
  expect(readFileSync(vs.path('notes.md'), 'utf8')).toBe(MD);
});
