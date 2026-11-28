// SPDX-License-Identifier: GPL-3.0-or-later
//
// Open, save, revert and the disk (ED-014 section 4.3) in a real VS Code: the round trip over six fixtures, an edit saved,
// a failed write, another writer, Revert File and the banner's Reload, an edit during a save, and two saves inside one
// export of the 255-page fixture.
import { test, expect } from '@playwright/test';
import { chmodSync, copyFileSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { VsCode, FIXTURES, bodyText, caretAtEndOf, insertInPage, sleep, watcherSettles } from './harness.js';
import { mainPart, partsThatDiffer } from './zip.js';

test.describe.configure({ mode: 'serial' });

const ROUND_TRIP: Record<string, string> = {
  'plain.docx': 'HelloWordOnline.docx',
  'rtl.docx': 'fonts-hebrew-no-cs.docx',
  'japanese.docx': 'fonts-cjk-linebox.docx',
  'comments.docx': 'comments-two.docx',
  'revisions.docx': 'revisions/revisions-word2010.docx',
  'fonts.docx': 'fonts-embedded.docx',
};

let vs: VsCode;

test.beforeAll(async () => {
  vs = await VsCode.launch({
    profile: 'save',
    files: {
      ...ROUND_TRIP,
      'edit.docx': 'HelloWordOnline.docx',
      'readonly.docx': 'HelloWordOnline.docx',
      'other.docx': 'HelloWordOnline.docx',
      'revert.docx': 'HelloWordOnline.docx',
      'large.docx': 'large/anon-250.docx',
      'large-twice.docx': 'large/anon-250.docx',
    },
  });
});

test.afterAll(async () => { await vs?.close(); });

test.afterEach(async () => {
  expect(vs.problems, 'content security policy refusals').toEqual([]);
  await vs.closeAll();
});

/** Waits for a tab's dirty state: longer on the 255-page fixture, whose page may still be busy after its open. */
async function tabIs(label: string, dirty: boolean): Promise<void> {
  const timeout = label.startsWith('large') ? 150_000 : 30_000;
  await expect.poll(async () => (await vs.tab(label))?.isDirty, { message: `${label} ${dirty ? 'dirty' : 'clean'}`, timeout }).toBe(dirty);
}

test('the round trip: six fixtures saved as new names unedited, every part the same but [Content_Types].xml and the .rels', async () => {
  for (const name of Object.keys(ROUND_TRIP)) {
    await vs.open(name);
    const copy = name.replace(/\.docx$/, '-copy.docx');
    await vs.saveAs(vs.path(copy));
    // VS Code opens the target afresh (P25).
    await vs.frameFor(copy);
    expect((await vs.tab(copy))?.viewType).toBe('docx4j.editor');
    expect(partsThatDiffer(vs.path(name), vs.path(copy)), name).toEqual([]);
    await vs.exec('workbench.action.closeAllEditors');
  }
});

test('an edit: typed with real keys, the tab dirty; Save writes it, and the tab and the page are clean', async () => {
  const frame = await vs.open('edit.docx');
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.type(' Typed in VS Code', { delay: 20 });
  await tabIs('edit.docx', true);
  await expect(frame.getByTestId('name')).toHaveText('edit.docx •');
  await vs.page.keyboard.press('Control+s');
  await tabIs('edit.docx', false);
  await expect(frame.getByTestId('name')).toHaveText('edit.docx');
  expect(mainPart(vs.path('edit.docx'))).toContain('Typed in VS Code');
  // The extension's own write is no change on disk.
  await sleep(1000);
  await expect(frame.getByTestId('disk-offer')).toHaveCount(0);
});

test('a failed write: Save fails, the tab stays dirty, and the page says why', async () => {
  chmodSync(vs.path('readonly.docx'), 0o444);
  try {
    const frame = await vs.open('readonly.docx');
    await caretAtEndOf(frame, 0);
    await vs.page.keyboard.type('!', { delay: 20 });
    await tabIs('readonly.docx', true);
    await vs.page.keyboard.press('Control+s');
    await expect.poll(() => vs.notifications()).toContainEqual(expect.stringContaining("Failed to save 'readonly.docx'"));
    await expect(frame.getByTestId('status')).toContainText('Not saved');
    await tabIs('readonly.docx', true);
  } finally {
    chmodSync(vs.path('readonly.docx'), 0o644);
    await vs.exec('notifications.clearAll');
  }
});

test('another writer: the banner; then Save fails and the tab stays dirty; after a revert, the extension\'s own save shows none', async () => {
  const frame = await vs.open('other.docx');
  await watcherSettles();
  copyFileSync(join(FIXTURES, 'fonts-hebrew-no-cs.docx'), vs.path('other.docx'));
  await expect(frame.getByTestId('disk-offer')).toBeVisible();
  await frame.getByTestId('disk-keep').click();
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.type('?', { delay: 20 });
  await tabIs('other.docx', true);
  await vs.page.keyboard.press('Control+s');
  await expect.poll(() => vs.notifications()).toContainEqual(expect.stringContaining('has changed on disk'));
  await tabIs('other.docx', true);
  expect(readFileSync(vs.path('other.docx'))).toEqual(readFileSync(join(FIXTURES, 'fonts-hebrew-no-cs.docx')));
  await vs.exec('notifications.clearAll');
  // Revert File reads the other writer's file; a save after it writes, with no banner.
  await vs.exec('workbench.action.files.revert');
  await tabIs('other.docx', false);
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.type('#', { delay: 20 });
  await vs.page.keyboard.press('Control+s');
  await tabIs('other.docx', false);
  await sleep(1000);
  await expect(frame.getByTestId('disk-offer')).toHaveCount(0);
});

test('Revert File, and the banner\'s Reload, clean and dirty', async () => {
  const frame = await vs.open('revert.docx');
  const original = await bodyText(frame);
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.type('XYZ', { delay: 20 });
  await tabIs('revert.docx', true);
  await vs.exec('workbench.action.files.revert');
  await tabIs('revert.docx', false);
  await expect.poll(() => bodyText(frame)).toBe(original);
  // Clean: the banner's Reload opens the file as it is now.
  await watcherSettles();
  copyFileSync(join(FIXTURES, 'fonts-cjk-linebox.docx'), vs.path('revert.docx'));
  await expect(frame.getByTestId('disk-offer')).toBeVisible();
  await frame.getByTestId('disk-reload').click();
  await expect.poll(() => bodyText(frame)).not.toBe(original);
  const second = await bodyText(frame);
  await tabIs('revert.docx', false);
  // Dirty: Reload after the confirmation, through VS Code's Revert File, which clears its dirty state.
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.type('QQ', { delay: 20 });
  await tabIs('revert.docx', true);
  await watcherSettles();
  copyFileSync(join(FIXTURES, 'HelloWordOnline.docx'), vs.path('revert.docx'));
  await expect(frame.getByTestId('disk-offer')).toBeVisible();
  await frame.getByTestId('disk-discard').click();
  await frame.locator('[role="dialog"] button[type="submit"]').click();
  await tabIs('revert.docx', false);
  await expect.poll(() => bodyText(frame)).toBe(original);
  expect(second).not.toBe(original);
});

/** Counts the file's writes, by its modification time, until `stop` is called. */
function countWrites(path: string): { stop: () => Promise<number> } {
  let last = statSync(path).mtimeMs;
  let writes = 0;
  let running = true;
  const loop = (async () => {
    while (running) {
      await sleep(100);
      const now = statSync(path).mtimeMs;
      if (now !== last) { writes++; last = now; }
    }
  })();
  return { stop: async () => { running = false; await loop; return writes; } };
}

test('an edit during a save, on the 255-page fixture: after the save, the tab is dirty and the page unsaved', async () => {
  const frame = await vs.open('large.docx');
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.type('A', { delay: 20 });
  await tabIs('large.docx', true);
  const before = statSync(vs.path('large.docx')).mtimeMs;
  await vs.page.keyboard.press('Control+s');
  // An edit the page makes while its export runs: queued behind the export, it lands before the write.
  void insertInPage(frame, 'B');
  await expect.poll(() => statSync(vs.path('large.docx')).mtimeMs, { timeout: 180_000 }).not.toBe(before);
  await sleep(1500);
  await tabIs('large.docx', true);
  await expect(frame.getByTestId('name')).toHaveText('large.docx •');
  await expect(frame.getByTestId('status')).toContainText('not saved yet');
  const saved = mainPart(vs.path('large.docx'));
  expect(saved).toContain('A');
  expect(await bodyText(frame)).toContain('AB');
});

test('two Ctrl+S inside one export of the 255-page fixture: the later save\'s bytes written last, nothing shown for the cancelled one', async () => {
  const frame = await vs.open('large-twice.docx');
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.type('CD', { delay: 20 });
  await tabIs('large-twice.docx', true);
  await vs.page.keyboard.press('Control+s');
  await sleep(400);
  // The second press waits for the page, busy with the first export; VS Code then cancels the first save.
  const writes = countWrites(vs.path('large-twice.docx'));
  await vs.page.keyboard.press('Control+s');
  await expect.poll(async () => (await vs.tab('large-twice.docx'))?.isDirty, { timeout: 180_000 }).toBe(false);
  await sleep(3000);
  // One write when VS Code's cancel reaches the extension before the first save's write, as measured (section 9.9); two if
  // the first wrote first. Never a cancelled save writing after the later one: the saves are queued.
  const count = await writes.stop();
  test.info().annotations.push({ type: 'writes', description: String(count) });
  expect(count).toBeGreaterThanOrEqual(1);
  expect(count).toBeLessThanOrEqual(2);
  expect(mainPart(vs.path('large-twice.docx'))).toContain('CD');
  expect(await vs.notifications()).toEqual([]);
  await expect(frame.getByTestId('name')).toHaveText('large-twice.docx');
});
