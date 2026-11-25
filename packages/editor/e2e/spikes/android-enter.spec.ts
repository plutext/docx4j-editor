// SPDX-License-Identifier: GPL-3.0-or-later
// ED-013 S5 spike: Enter as Chrome on Android delivers it, after a Backspace, with tracking off and on. On Android
// ProseMirror leaves Enter's key to the browser, which splits the paragraph in the DOM, and reads the split back
// (`readDOMChange`): an Enter command when the change looks like one, else the parsed change as a replacement. S0 saw
// the replacement once in Gboard, with tracking on, right after a tracked Backspace at the deleted text. Keys go through
// Chromium's DevTools protocol (`Input.dispatchKeyEvent`, the browser's own default actions) under a Pixel 7's profile
// (ProseMirror's `android` and `chrome` paths). Reports; asserts nothing. Run:
// npx playwright test -c playwright.spike.config.ts e2e/spikes/android-enter.spec.ts --project chromium
import { test, devices, type Page, type CDPSession } from '@playwright/test';
import { fixturePath, caretInBlock } from '../helpers.js';

const { defaultBrowserType: _b, ...pixel } = devices['Pixel 7'];
void _b;
test.use(pixel);

/** Opens a fixture through the file input the compact layout keeps (it draws no ribbon, which `openFixture` waits for). */
async function openCompact(page: Page, name: string): Promise<void> {
  await page.goto('/');
  await page.getByTestId('open').setInputFiles(fixturePath(name));
  await page.locator('.ProseMirror[aria-label="Document"]').waitFor();
}

type Entry = { command: unknown; newMark: unknown; head: number; steps: string[] };

async function instrument(page: Page): Promise<void> {
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { view: { dispatch: (tr: unknown) => void } } }).docx4jEditor.view;
    const w = window as unknown as { __log: Entry[] };
    w.__log = [];
    const orig = v.dispatch.bind(v);
    v.dispatch = (tr: unknown): void => {
      const t = tr as { docChanged: boolean; getMeta(k: string): unknown; selection: { head: number }; steps: { toJSON(): { stepType: string; from: number; to: number; slice?: { content?: { type: string; text?: string; content?: { text?: string }[]; marks?: { type: string }[] }[] } } }[] };
      if (t.docChanged) {
        w.__log.push({
          command: t.getMeta('docx4j-command'), newMark: t.getMeta('docx4j-new-mark'), head: t.selection.head,
          steps: t.steps.map((s) => { const j = s.toJSON(); return `${j.stepType} ${j.from}-${j.to} ${JSON.stringify((j.slice?.content ?? []).map((c) => c.text ?? `<${c.type}:${(c.content ?? []).map((x) => x.text ?? '').join('')}>`))}`; }),
        });
      }
      orig(tr);
    };
  });
}

/** Blocks from `index` on: their text with each text node's revision marks, and the caret. */
async function report(page: Page, index: number, count = 2): Promise<{ blocks: string[]; caret: string }> {
  return page.evaluate(([i, n]) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { nodeSize: number; forEach(f: (c: { text?: string; type: { name: string }; marks: { type: { name: string }; attrs: { author?: string } }[] }) => void): void } }; selection: { head: number; $head: { parent: { textContent: string }; parentOffset: number } } } } }).docx4jEditor;
    const blocks: string[] = [];
    for (let k = i; k < Math.min(i + n, v.state.doc.childCount); k++) {
      const parts: string[] = [];
      v.state.doc.child(k).forEach((c) => parts.push(`${c.text ?? `[${c.type.name}]`}${c.marks.filter((m) => m.type.name === 'insertion' || m.type.name === 'deletion').map((m) => `{${m.type.name}:${m.attrs.author}}`).join('')}`));
      blocks.push(parts.join('|'));
    }
    const { $head } = v.state.selection;
    return { blocks, caret: `${JSON.stringify($head.parent.textContent)} @ ${$head.parentOffset}` };
  }, [index, count] as const);
}

async function key(client: CDPSession, name: 'Backspace' | 'Enter'): Promise<void> {
  const code = name === 'Enter' ? 13 : 8;
  await client.send('Input.dispatchKeyEvent', { type: 'rawKeyDown', key: name, code: name, windowsVirtualKeyCode: code, nativeVirtualKeyCode: code });
  if (name === 'Enter') await client.send('Input.dispatchKeyEvent', { type: 'char', key: 'Enter', text: '\r', unmodifiedText: '\r', windowsVirtualKeyCode: 13 });
  await client.send('Input.dispatchKeyEvent', { type: 'keyUp', key: name, code: name, windowsVirtualKeyCode: code, nativeVirtualKeyCode: code });
}

async function blockWith(page: Page, text: string): Promise<number> {
  return page.evaluate((t) => {
    const doc = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc;
    for (let i = 0; i < doc.childCount; i++) if (doc.child(i).textContent.includes(t)) return i;
    return -1;
  }, text);
}

const CASES: { name: string; block: string; at: (text: string) => number; keys: ('Backspace' | 'Enter')[]; tracking: boolean; insertFirst?: string }[] = [
  { name: 'plain text, Enter alone', block: 'Some text.', at: (t) => t.indexOf('text'), keys: ['Enter'], tracking: false },
  { name: 'plain text, Enter alone, tracking', block: 'Some text.', at: (t) => t.indexOf('text'), keys: ['Enter'], tracking: true },
  { name: 'plain text, Backspace then Enter', block: 'Some text.', at: (t) => t.indexOf('text') + 2, keys: ['Backspace', 'Enter'], tracking: false },
  { name: 'plain text, Backspace then Enter, tracking', block: 'Some text.', at: (t) => t.indexOf('text') + 2, keys: ['Backspace', 'Enter'], tracking: true },
  { name: 'plain text, two Backspaces then Enter, tracking', block: 'Some text.', at: (t) => t.indexOf('text') + 3, keys: ['Backspace', 'Backspace', 'Enter'], tracking: true },
  { name: "another author's insertion, Backspace then Enter, tracking", block: 'An insertion', at: (t) => t.indexOf('insertion') + 3, keys: ['Backspace', 'Enter'], tracking: true },
  { name: 'just after a deletion, Enter, tracking', block: 'A deletion', at: (t) => t.indexOf('A deletion') + 'A deletion'.length, keys: ['Enter'], tracking: true },
  { name: 'just before a deletion, Enter, tracking', block: 'A deletion', at: (t) => t.indexOf('A deletion'), keys: ['Enter'], tracking: true },
];

for (const commit of [true, false]) {
  test(`a composition ${commit ? 'committed' : 'still open'}, then Enter, tracking`, async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    await page.evaluate(async () => {
      const v = (window as unknown as { docx4jEditor: { session: { trackingAuthor: string | undefined }; runAsync(n: string, a: unknown): Promise<boolean> } }).docx4jEditor;
      v.session.trackingAuthor = 'Phone';
      await v.runAsync('document.setChangeTrackingMode', { value: 'TrackAll' });
    });
    const index = await blockWith(page, 'Some text.');
    await page.locator('.ProseMirror[aria-label="Document"]').focus();
    await caretInBlock(page, index, 5);
    await page.waitForTimeout(100);
    await instrument(page);
    await page.evaluate(() => {
      const w = window as unknown as { __ev: string[] };
      w.__ev = [];
      const root = document.querySelector('.ProseMirror[aria-label="Document"]')!;
      for (const type of ['keydown', 'beforeinput', 'input', 'compositionstart', 'compositionend']) root.addEventListener(type, (e) => w.__ev.push(`${type} ${(e as InputEvent).inputType ?? (e as KeyboardEvent).key ?? ''} ${(e as InputEvent).isComposing ? 'composing' : ''} ${e.cancelable ? 'C' : ''}`), true);
    });
    const client = await page.context().newCDPSession(page);
    await client.send('Input.imeSetComposition', { text: 'wor', selectionStart: 3, selectionEnd: 3 });
    await page.waitForTimeout(150);
    if (commit) { await client.send('Input.insertText', { text: 'wor' }); await page.waitForTimeout(150); }
    await key(client, 'Enter');
    await page.waitForTimeout(300);
    console.log(`composition ${commit}`, 'EVENTS', JSON.stringify(await page.evaluate(() => (window as unknown as { __ev: string[] }).__ev)));
    console.log(`composition ${commit}`, 'after', JSON.stringify(await report(page, index)));
    console.log(`composition ${commit}`, 'LOG', JSON.stringify(await page.evaluate(() => (window as unknown as { __log: Entry[] }).__log)));
  });
}

for (const c of CASES) {
  test(c.name, async ({ page }) => {
    await openCompact(page, 'tracked-changes.docx');
    if (c.tracking) {
      await page.evaluate(async () => {
        const v = (window as unknown as { docx4jEditor: { session: { trackingAuthor: string | undefined }; runAsync(n: string, a: unknown): Promise<boolean> } }).docx4jEditor;
        v.session.trackingAuthor = 'Phone';
        await v.runAsync('document.setChangeTrackingMode', { value: 'TrackAll' });
      });
    }
    const index = await blockWith(page, c.block);
    const text = await page.evaluate((i) => (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc.child(i).textContent, index);
    await page.locator('.ProseMirror[aria-label="Document"]').focus();
    await caretInBlock(page, index, c.at(text));
    await page.waitForTimeout(100);
    console.log(c.name, 'BEFORE', JSON.stringify(await report(page, index, 1)));
    await instrument(page);
    const client = await page.context().newCDPSession(page);
    for (const k of c.keys) {
      await key(client, k);
      await page.waitForTimeout(250);
      console.log(c.name, `after ${k}`, JSON.stringify(await report(page, index)));
    }
    console.log(c.name, 'LOG', JSON.stringify(await page.evaluate(() => (window as unknown as { __log: Entry[] }).__log)));
  });
}
