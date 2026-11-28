// SPDX-License-Identifier: GPL-3.0-or-later
//
// The README's screenshot (ED-014 section 4.7, P50), made again by `npm run screenshot -w docx4j-editor` after the root build:
// a document with a word replaced by real keys under tracking, beside its XML in reveal codes, in an installed VS Code
// window of the listing's proportions with its side bars closed. Written to `media/screenshot.png`, outside the VSIX's
// allowlist; the README links it at its address in the public repository. Not part of the suite (`playwright.config.ts`
// leaves it out). `DOCX4J_SCREENSHOT_OUT` writes it elsewhere.
import { test, expect } from '@playwright/test';
import { join } from 'node:path';
import { EXTENSION, VsCode, sleep } from './harness.js';
import { vsixes } from './vsix.js';

const OUT = process.env['DOCX4J_SCREENSHOT_OUT'] ?? join(EXTENSION, 'media', 'screenshot.png');
/** The paragraph and the word replaced in it. */
const BLOCK = 1;
const WORD = 'basic';

test('the README\'s screenshot', async () => {
  const { extension, driver } = vsixes();
  const vs = await VsCode.launch({ profile: 'screenshot', installed: [extension, driver], files: { 'Report.docx': 'tracked-changes.docx' } });
  try {
    await vs.page.setViewportSize({ width: 1440, height: 900 });
    for (const command of ['workbench.action.closeSidebar', 'workbench.action.closeAuxiliaryBar', 'workbench.action.closePanel']) await vs.exec(command).catch(() => undefined);
    const frame = await vs.open('Report.docx');
    await frame.getByTestId('perspective-developer').click();
    // Tracking on asks the person's name the first time.
    await frame.locator('.ProseMirror[aria-label="Document"]').click();
    await vs.page.keyboard.press('Control+Shift+E');
    const name = frame.getByRole('dialog').getByRole('textbox');
    await name.fill('Sam Taylor');
    await name.press('Enter');
    await expect(frame.getByRole('dialog')).toHaveCount(0);
    // The word selected through the view, then replaced with real keys.
    await frame.evaluate(([index, word]) => {
      type Node = { nodeSize: number; isText: boolean; text?: string; descendants(f: (n: Node, pos: number) => boolean | void): void };
      const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): Node } }; caretAt(p: number, to?: number): void; focus(): void } }).docx4jEditor;
      let start = 0;
      for (let k = 0; k < index; k++) start += v.state.doc.child(k).nodeSize;
      let at = -1;
      v.state.doc.child(index).descendants((n, pos) => { if (at < 0 && n.isText && n.text!.includes(word)) at = start + 1 + pos + n.text!.indexOf(word); });
      if (at < 0) throw new Error(`no "${word}" in block ${index}`);
      v.caretAt(at, at + word.length);
      v.focus();
    }, [BLOCK, WORD] as const);
    await vs.page.keyboard.type('essential', { delay: 40 });
    await expect(frame.locator('.docx-editor .ProseMirror ins.insertion', { hasText: 'essential' })).toHaveCount(1);
    await sleep(1500);
    await vs.page.screenshot({ path: OUT });
    console.log(`screenshot: ${OUT}`);
  } finally {
    await vs.close();
  }
});
