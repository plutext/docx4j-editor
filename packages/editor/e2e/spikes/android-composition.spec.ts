// SPDX-License-Identifier: GPL-3.0-or-later
// ED-013 S5 spike: Samsung Keyboard's composition inside a marked run, simulated over Chromium's DevTools
// protocol, with an Android user agent (ProseMirror's `android` and `chrome` paths) and without. Reports the
// steps and carets; asserts nothing. Run: npx playwright test -c playwright.spike.config.ts
// e2e/spikes/android-composition.spec.ts --project chromium
import { test, devices, type Page } from '@playwright/test';
import { openFixture, caretInBlock } from '../helpers.js';

async function blockWith(page: Page, text: string): Promise<number> {
  return page.evaluate((t) => {
    const doc = (window as unknown as { docx4jEditor: { state: { doc: { childCount: number; child(i: number): { textContent: string } } } } }).docx4jEditor.state.doc;
    for (let i = 0; i < doc.childCount; i++) if (doc.child(i).textContent.includes(t)) return i;
    return -1;
  }, text);
}

type Log = { composing: boolean; sel: number; steps: string[] }[];

async function instrument(page: Page): Promise<void> {
  await page.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: { view: { dispatch: (tr: unknown) => void; composing: boolean } } }).docx4jEditor.view;
    const w = window as unknown as { __log: Log };
    w.__log = [];
    const orig = v.dispatch.bind(v);
    v.dispatch = (tr: unknown): void => {
      const t = tr as { docChanged: boolean; selection: { head: number }; steps: { toJSON(): { stepType: string; from: number; to: number; slice?: { content?: { text?: string; marks?: { type: string }[] }[] } } }[] };
      if (t.docChanged) w.__log.push({ composing: v.composing, sel: t.selection.head, steps: t.steps.map((s) => { const j = s.toJSON(); return `${j.stepType} ${j.from}-${j.to} ${JSON.stringify((j.slice?.content ?? []).map((c) => [c.text, (c.marks ?? []).map((m) => m.type)]))}`; }) });
      orig(tr);
    };
  });
}

async function state(page: Page, index: number): Promise<{ text: string; head: number; start: number; nodes: string[] }> {
  return page.evaluate((i) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { textContent: string; nodeSize: number; forEach(f: (n: { text?: string; marks: { type: { name: string } }[] }) => void): void } }; selection: { head: number } } } }).docx4jEditor;
    let start = 0;
    for (let k = 0; k < i; k++) start += v.state.doc.child(k).nodeSize;
    const nodes: string[] = [];
    v.state.doc.child(i).forEach((n) => nodes.push(`${JSON.stringify(n.text)} [${n.marks.map((m) => m.type.name).join(',')}]`));
    return { text: v.state.doc.child(i).textContent, head: v.state.selection.head, start: start + 1, nodes };
  }, index);
}

for (const ua of ['android', 'desktop'] as const) {
  test.describe(ua, () => {
    if (ua === 'android') { const { defaultBrowserType: _b, ...pixel } = devices['Pixel 7']; void _b; test.use(pixel); }
    test(`composition after a word, one character left in its run (${ua})`, async ({ page }) => {
      await openFixture(page, 'tracked-changes.docx');
      const index = await blockWith(page, 'Some text.');
      const before = await state(page, index);
      console.log(ua, 'before', JSON.stringify(before));
      const offset = before.text.indexOf('text.') + 4;
      await caretInBlock(page, index, offset);
      await page.locator('.ProseMirror[aria-label="Document"]').focus();
      await caretInBlock(page, index, offset);
      await instrument(page);
      const client = await page.context().newCDPSession(page);
      for (const text of ['a', 'ab']) {
        await client.send('Input.imeSetComposition', { text, selectionStart: text.length, selectionEnd: text.length });
        await page.waitForTimeout(150);
        const s = await state(page, index);
        const html = await page.evaluate((i) => (document.querySelectorAll('.ProseMirror[aria-label="Document"] > p')[i] as HTMLElement).innerHTML, index);
        console.log(ua, `after "${text}": head ${s.head - s.start} nodes ${JSON.stringify(s.nodes)} dom ${html}`);
      }
      await client.send('Input.insertText', { text: 'ab' });
      await page.waitForTimeout(150);
      await client.send('Input.insertText', { text: 'c' });
      await page.waitForTimeout(300);
      const after = await state(page, index);
      console.log(ua, 'after', JSON.stringify({ text: after.text, head: after.head - after.start, nodes: after.nodes }));
      console.log(ua, 'log', JSON.stringify(await page.evaluate(() => (window as unknown as { __log: Log }).__log), null, 1));
    });
  });
}
