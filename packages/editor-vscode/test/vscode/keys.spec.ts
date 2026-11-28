// SPDX-License-Identifier: GPL-3.0-or-later
//
// The keyboard and the clipboard in a real VS Code (ED-014 section 4.4, S4.4): real keys over the DevTools Protocol, the
// document read in the page and VS Code in its workbench page. A key pressed in the document belongs to the page or to
// VS Code, never to both (P32); "docx4j editor: Commands" opens the application's palette (P33); Ctrl+Shift+V and "Paste as
// Plain Text" paste the text alone (P34); Options has no spelling option (P35).
//
// - **The settings:** a custom title bar, the classic menu bar and custom menus, without which the lone-Alt case and the
//   context menu's clicks would observe nothing (section 9.12).
// - **Ctrl+Q is not pressed here:** a broken gate would quit VS Code in the middle of the suite. S4.1 measured it (section
//   9.13 item 1), and the browser suite holds the page's half.
// - **The clipboard** is saved through the driver first and put back after.
import { test, expect, type Frame } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { VsCode, caretAtEndOf, bodyText, sleep } from './harness.js';

test.describe.configure({ mode: 'serial' });

let vs: VsCode;
let clipboard = '';

test.beforeAll(async () => {
  vs = await VsCode.launch({
    profile: 'keys',
    files: { 'a.docx': 'HelloWordOnline.docx', 'b.docx': 'HelloWordOnline.docx', 'hf.docx': 'odd-even-first-page.docx', 'rev.docx': 'revisions/revisions-word15.docx' },
    settings: { 'window.titleBarStyle': 'custom', 'window.menuBarVisibility': 'classic', 'window.menuStyle': 'custom' },
  });
  writeFileSync(vs.path('notes.txt'), 'A text file.\n');
  clipboard = await vs.ctl<string>({ op: 'clipboardRead' });
});
test.afterAll(async () => {
  if (vs) await vs.ctl({ op: 'clipboardWrite', text: clipboard }).catch(() => undefined);
  await vs?.close();
});

interface Workbench {
  sidebar: boolean;
  panel: boolean;
  zoom: number;
  quickInput: string | null;
  status: string;
  menuBarFocused: boolean;
  /** Whether VS Code's focus is on the webview still: a binding that moves it (Show Explorer, the breadcrumbs) moves it away. */
  webviewFocused: boolean;
  windows: number;
}

/** What VS Code shows, read in its workbench page. */
async function workbench(): Promise<Workbench> {
  const state = await vs.page.evaluate(() => {
    const w = document.querySelector('.monaco-workbench')!;
    const quick = document.querySelector('.quick-input-widget') as HTMLElement | null;
    const shown = !!quick && quick.style.display !== 'none';
    return {
      sidebar: !w.classList.contains('nosidebar'),
      panel: !w.classList.contains('nopanel'),
      zoom: window.devicePixelRatio,
      quickInput: shown ? ((quick!.querySelector('input:not([type="checkbox"])') as HTMLInputElement | null)?.value ?? '') : null,
      status: (document.querySelector('.part.statusbar') as HTMLElement | null)?.innerText.replace(/\s+/g, ' ') ?? '',
      menuBarFocused: !!document.activeElement?.closest('.menubar'),
      webviewFocused: document.hasFocus() && document.activeElement?.tagName === 'IFRAME' && document.activeElement.classList.contains('webview'),
    };
  });
  const windows = vs.browser.contexts().flatMap((c) => c.pages()).filter((p) => p.url().includes('workbench')).length;
  return { ...state, windows };
}


type View = { state: { doc: { textContent: string; content: { size: number }; descendants(f: (node: { isText: boolean; text?: string }, pos: number) => boolean | void): void }; selection: { from: number; to: number } }; formatting(): { font: { bold: unknown; italic: unknown }; paragraph: { alignment: unknown } }; caretAt(a: number, b?: number): void; focus(): void; view: { coordsAtPos(p: number): { left: number; top: number; bottom: number } } };
const reading = <T>(frame: Frame, f: (v: View) => T): Promise<T> => frame.evaluate(`(${f.toString()})(window.docx4jEditor)`) as Promise<T>;

/**
 * The document's first word selected through the view, from its first text node (a paragraph may begin with invisible
 * atoms, bookmarks among them), and the point at the word's middle in the workbench page.
 */
async function selectFirstWord(frame: Frame): Promise<{ word: string; x: number; y: number }> {
  const at = await frame.evaluate(() => {
    const v = (window as unknown as { docx4jEditor: View }).docx4jEditor;
    let start = -1;
    let word = '';
    v.state.doc.descendants((node, pos) => {
      if (start >= 0) return false;
      if (node.isText) { start = pos; word = node.text!.split(/\s+/)[0]!; return false; }
      return true;
    });
    v.caretAt(start, start + word.length);
    v.focus();
    const a = v.view.coordsAtPos(start);
    const b = v.view.coordsAtPos(start + word.length);
    return { word, x: (a.left + b.left) / 2, y: (a.top + a.bottom) / 2 };
  });
  const box = (await (await frame.frameElement()).boundingBox())!;
  return { word: at.word, x: box.x + at.x, y: box.y + at.y };
}

/** VS Code's context menu over the page at a point, and one of its items. */
async function contextMenu(x: number, y: number, item: string): Promise<void> {
  await vs.page.mouse.click(x, y, { button: 'right' });
  const entry = vs.page.locator('.context-view .monaco-menu .action-label', { hasText: new RegExp(`^${item}$`) });
  await entry.first().waitFor({ timeout: 10_000 });
  // A menu item takes a mouse-up only 100 ms after the menu is drawn (VS Code's `runOnceToEnableMouseUp`); a click sooner is
  // lost, and the menu stays open.
  await sleep(400);
  await entry.first().click();
  await expect(vs.page.locator('.context-view .monaco-menu')).toHaveCount(0);
}

test('the document\'s keys stay in the page: its change made, VS Code\'s not (P32)', async () => {
  const frame = await vs.open('a.docx');
  const before = await workbench();
  const unchanged = async (key: string): Promise<void> => {
    await sleep(500);
    const now = await workbench();
    expect({ ...now, status: now.status.includes('Waiting for second key') || now.status.includes('Tab Moves Focus') }, key)
      .toEqual({ ...before, status: false, webviewFocused: true });
  };
  const press = async (key: string, check: () => Promise<unknown>, expected: unknown): Promise<void> => {
    await caretAtEndOf(frame, 0);
    await vs.page.keyboard.press(key);
    await expect.poll(check, { message: key }).toEqual(expected);
    await unchanged(key);
  };
  const font = (): Promise<{ bold: unknown; italic: unknown }> => reading(frame, (v) => { const f = v.formatting().font; return { bold: f.bold, italic: f.italic }; });
  const alignment = (): Promise<unknown> => reading(frame, (v) => v.formatting().paragraph.alignment);
  await press('Control+b', font, { bold: true, italic: false });
  await vs.page.keyboard.press('Control+b');
  await press('Control+i', font, { bold: false, italic: true });
  await vs.page.keyboard.press('Control+i');
  await press('Control+e', alignment, 'Centered');
  await press('Control+r', alignment, 'Right');
  await press('Control+j', alignment, 'Justified');
  await press('Control+l', alignment, 'Left');
  // Subscript and superscript, the font grown, the indent and Normal style: VS Code's zoom, breadcrumbs, Tab Moves Focus,
  // Problems and New Window each kept out.
  const text = (): Promise<number> => reading(frame, (v) => v.state.doc.content.size);
  const size = await text();
  for (const key of ['Control+Equal', 'Control+Shift+Equal', 'Control+Shift+Period', 'Control+m', 'Control+Shift+m', 'Control+Shift+n']) {
    await caretAtEndOf(frame, 0);
    await vs.page.keyboard.press(key);
    await unchanged(key);
  }
  expect(await text()).toBe(size);
  // Track Changes, VS Code's Show Explorer.
  const tracking = frame.getByTestId('tracking-switch');
  const was = await tracking.getAttribute('aria-pressed');
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.press('Control+Shift+E');
  // With no name set, a name is asked for first.
  const name = frame.getByTestId('tracking-name');
  await expect(name.or(tracking.and(frame.locator(`[aria-pressed="${was === 'true' ? 'false' : 'true'}"]`)))).toBeVisible();
  if (await name.isVisible()) { await name.fill('Ann'); await name.press('Enter'); }
  await expect(tracking).not.toHaveAttribute('aria-pressed', was ?? '');
  await unchanged('Control+Shift+E');
  await vs.page.keyboard.press('Control+Shift+E');
  await expect(tracking).toHaveAttribute('aria-pressed', was ?? '');
  // Undo and redo, the page's.
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.type('q');
  await expect.poll(() => bodyText(frame)).toContain('Onlineq');
  await vs.page.keyboard.press('Control+z');
  await expect.poll(() => bodyText(frame)).not.toContain('Onlineq');
  await unchanged('Control+z');
  await vs.page.keyboard.press('Control+y');
  await expect.poll(() => bodyText(frame)).toContain('Onlineq');
  await unchanged('Control+y');
  // Select All, and Find.
  await press('Control+a', () => reading(frame, (v) => v.state.selection.to - v.state.selection.from > 20), true);
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.press('Control+f');
  await expect(frame.getByTestId('find-bar')).toBeVisible();
  await vs.page.keyboard.press('Escape');
  await vs.closeAll();
});

test('from a footer: Ctrl+S saves, and Ctrl+B bolds without toggling the side bar (section 9.12)', async () => {
  const frame = await vs.open('hf.docx');
  const footer = frame.locator('.docx-editor .page-boundary .boundary-footer .ProseMirror').first();
  await footer.click();
  await expect(frame.locator('.page-column')).toHaveClass(/editing-header/);
  await sleep(300);
  await vs.page.keyboard.press('End');
  await vs.page.keyboard.type(' typed', { delay: 30 });
  await expect.poll(async () => (await vs.tab('hf.docx'))?.isDirty).toBe(true);
  const before = await workbench();
  await vs.page.keyboard.press('Control+b');
  await sleep(500);
  expect((await workbench()).sidebar).toBe(before.sidebar);
  await vs.page.keyboard.type('bold', { delay: 30 });
  await expect(footer.locator('strong')).toContainText('bold');
  await vs.page.keyboard.press('Control+s');
  await expect.poll(async () => (await vs.tab('hf.docx'))?.isDirty, { timeout: 60_000 }).toBe(false);
  await vs.closeAll();
});

test('Ctrl+K, then an address and Enter in the dialog: the link is made, every editor stays open, and no chord starts', async () => {
  await vs.open('b.docx');
  // Kept, not a preview, so that both tabs stay in the group.
  await vs.exec('workbench.action.keepEditor');
  const frame = await vs.open('a.docx');
  await vs.exec('workbench.action.keepEditor');
  const tabs = (await vs.tabs()).map((t) => t.label).sort();
  expect(tabs).toEqual(['a.docx', 'b.docx']);
  const { word } = await selectFirstWord(frame);
  await vs.page.keyboard.press('Control+k');
  const address = frame.getByTestId('link-address');
  await expect(address).toBeVisible();
  expect((await workbench()).status).not.toContain('Waiting for second key');
  await vs.page.keyboard.type('www.example.com', { delay: 40 });
  await vs.page.keyboard.press('Enter');
  // The new link's text, over however many elements the link's runs draw it in.
  await expect.poll(() => frame.evaluate(() => [...document.querySelectorAll('.ProseMirror[aria-label="Document"] a.link')]
    .filter((a) => (a.getAttribute('data-target') ?? '').includes('example.com')).map((a) => a.textContent).join(''))).toBe(word);
  await sleep(500);
  expect((await vs.tabs()).map((t) => t.label).sort()).toEqual(tabs);
  await vs.closeAll();
});

test('VS Code\'s keys from the document: Ctrl+P, Ctrl+Shift+P and F1 open VS Code\'s, and Ctrl+W closes a clean document', async () => {
  const frame = await vs.open('a.docx');
  for (const [key, prefix] of [['Control+p', ''], ['Control+Shift+P', '>'], ['F1', '>']] as const) {
    await caretAtEndOf(frame, 0);
    await vs.page.keyboard.press(key);
    await expect.poll(async () => (await workbench()).quickInput, { message: key }).toBe(prefix);
    await expect(frame.getByTestId('palette')).toHaveCount(0);
    await vs.page.keyboard.press('Escape');
    await expect.poll(async () => (await workbench()).quickInput).toBeNull();
  }
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.press('Control+w');
  await expect.poll(async () => (await vs.tabs()).length).toBe(0);
});

test('Alt+PageDown with Alt released first: the caret goes to the next change, and the menu bar does not take the focus', async () => {
  const frame = await vs.open('rev.docx');
  await frame.evaluate(() => { const v = (window as unknown as { docx4jEditor: View }).docx4jEditor; v.caretAt(1); v.focus(); });
  await sleep(300);
  await vs.page.keyboard.down('Alt');
  await vs.page.keyboard.down('PageDown');
  await vs.page.keyboard.up('Alt');
  await vs.page.keyboard.up('PageDown');
  await expect.poll(() => reading(frame, (v) => v.state.selection.from)).not.toBe(1);
  await sleep(500);
  expect((await workbench()).menuBarFocused).toBe(false);
  expect((await workbench()).webviewFocused).toBe(true);
});

test('Escape in the document, and in the find bar, hides a notification; Ctrl+Z in the find bar\'s field undoes its typing', async () => {
  const frame = await vs.open('rev.docx');
  const toasts = (): Promise<number> => vs.page.evaluate(() => document.querySelectorAll('.notification-toast').length);
  await vs.ctl({ op: 'notify', text: 'From the suite' });
  await expect.poll(toasts).toBe(1);
  await frame.evaluate(() => (window as unknown as { docx4jEditor: View }).docx4jEditor.focus());
  await vs.page.keyboard.press('Escape');
  await expect.poll(toasts).toBe(0);
  await vs.page.keyboard.press('Control+f');
  const query = frame.getByTestId('find-query');
  await expect(query).toBeFocused();
  // The bar opens with the words at the caret in its field: the undo takes the typing back to them.
  const opened = await query.inputValue();
  await vs.page.keyboard.press('End');
  await vs.page.keyboard.type('abc', { delay: 30 });
  await expect(query).toHaveValue(`${opened}abc`);
  await vs.page.keyboard.press('Control+z');
  await expect(query).toHaveValue(opened);
  await vs.ctl({ op: 'notify', text: 'From the suite again' });
  await expect.poll(toasts).toBe(1);
  await query.focus();
  await vs.page.keyboard.press('Escape');
  await expect.poll(toasts).toBe(0);
  await vs.closeAll();
});

test('"docx4j editor: Commands" opens the page\'s palette, from VS Code\'s palette and by Ctrl+Alt+P, from the document and from the tab (P33)', async () => {
  const frame = await vs.open('a.docx');
  const palette = frame.getByTestId('palette');
  const opened = async (how: string): Promise<void> => {
    await expect(palette, how).toBeVisible();
    await expect.poll(() => frame.evaluate(() => document.hasFocus() && !!document.activeElement?.closest('[data-testid="palette"]')), { message: how }).toBe(true);
    await vs.page.keyboard.press('Escape');
    await expect(palette).toHaveCount(0);
  };
  // VS Code's palette, from the document.
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.press('Control+Shift+P');
  const input = vs.page.locator('.quick-input-widget input:not([type="checkbox"])');
  await input.waitFor();
  await input.fill('>docx4j editor: Commands');
  await expect(vs.page.locator('.quick-input-list .monaco-list-row').first()).toContainText('docx4j editor: Commands');
  await vs.page.keyboard.press('Enter');
  await opened('VS Code\'s palette');
  // Its key, from the document.
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.press('Control+Alt+p');
  await opened('Ctrl+Alt+P in the document');
  // Its key, from the editor's tab.
  await vs.page.locator('.tabs-container .tab', { hasText: 'a.docx' }).first().click();
  await expect.poll(() => frame.evaluate(() => document.hasFocus())).toBe(false);
  await vs.page.keyboard.press('Control+Alt+p');
  await opened('Ctrl+Alt+P on the tab');
  // With a text editor active, VS Code's palette does not list it.
  await vs.exec('vscode.open', { $uri: vs.path('notes.txt') });
  await expect.poll(async () => (await vs.tabs()).find((t) => t.isActive)?.label).toBe('notes.txt');
  void vs.exec('workbench.action.showCommands').catch(() => undefined);
  await input.waitFor();
  await input.fill('>docx4j editor: Commands');
  await sleep(800);
  const rows = await vs.page.locator('.quick-input-list .monaco-list-row').allInnerTexts();
  expect(rows.filter((r) => r.includes('docx4j editor: Commands'))).toEqual([]);
  await vs.page.keyboard.press('Escape');
  await vs.closeAll();
});

test('paste as plain text: Ctrl+Shift+V pastes a bold word plain where Ctrl+V keeps it bold; an address becomes a link; in the find bar; from the context menu (P34)', async () => {
  const frame = await vs.open('a.docx');
  const first = frame.locator('.ProseMirror[aria-label="Document"] > p').first();
  // A bold word copied in the document. The plain paste first, after the paragraph's plain end: pasted text takes the
  // caret's own formatting, as typed text does, and only the clipboard's is left behind.
  const { word } = await selectFirstWord(frame);
  await vs.page.keyboard.press('Control+b');
  await vs.page.keyboard.press('Control+c');
  await expect.poll(() => vs.ctl<string>({ op: 'clipboardRead' })).toBe(word);
  await expect(first.locator('strong')).toHaveCount(1);
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.press('Control+Shift+V');
  await expect.poll(async () => (await first.innerText()).split(word).length - 1).toBe(2);
  await expect(first.locator('strong')).toHaveCount(1);
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.press('Control+v');
  await expect.poll(async () => (await first.innerText()).split(word).length - 1).toBe(3);
  await expect(first.locator('strong')).toHaveCount(2);
  // One web address: a link.
  await vs.ctl({ op: 'clipboardWrite', text: 'https://docx4j.org/' });
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.press('Control+Shift+V');
  await expect(first.locator('a.link')).toHaveAttribute('data-target', 'https://docx4j.org/');
  // The find bar's field.
  await vs.ctl({ op: 'clipboardWrite', text: 'in the field' });
  await caretAtEndOf(frame, 0);
  await vs.page.keyboard.press('Control+f');
  const query = frame.getByTestId('find-query');
  await expect(query).toBeFocused();
  await query.fill('');
  await vs.page.keyboard.press('Control+Shift+V');
  await expect(query).toHaveValue('in the field');
  await vs.page.keyboard.press('Escape');
  // The context menu's Paste as Plain Text, where the caret was.
  await vs.ctl({ op: 'clipboardWrite', text: 'From the menu' });
  const at = await selectFirstWord(frame);
  await caretAtEndOf(frame, 0);
  await contextMenu(at.x, at.y, 'Paste as Plain Text');
  await expect.poll(() => bodyText(frame)).toContain('From the menu');
  await vs.closeAll();
});

test('Cut, Copy and Paste from the context menu reach the page', async () => {
  const frame = await vs.open('a.docx');
  const at = await selectFirstWord(frame);
  await vs.ctl({ op: 'clipboardWrite', text: 'before' });
  await contextMenu(at.x, at.y, 'Copy');
  await expect.poll(() => vs.ctl<string>({ op: 'clipboardRead' })).toBe(at.word);
  const cut = await selectFirstWord(frame);
  await contextMenu(cut.x, cut.y, 'Cut');
  await expect.poll(() => bodyText(frame)).not.toMatch(new RegExp(`^${cut.word}`));
  await vs.ctl({ op: 'clipboardWrite', text: 'Pasted' });
  await caretAtEndOf(frame, 0);
  const end = await frame.evaluate(() => { const v = (window as unknown as { docx4jEditor: View }).docx4jEditor; const c = v.view.coordsAtPos(v.state.selection.from); return { x: c.left - 2, y: (c.top + c.bottom) / 2 }; });
  const box = (await (await frame.frameElement()).boundingBox())!;
  await contextMenu(box.x + end.x, box.y + end.y, 'Paste');
  await expect.poll(() => bodyText(frame)).toContain('Pasted');
  await vs.closeAll();
});

test('Options has no spelling option (P35)', async () => {
  const frame = await vs.open('a.docx');
  await frame.getByLabel('Options').click();
  await expect(frame.getByTestId('option-name')).toBeVisible();
  await expect(frame.getByTestId('option-spellcheck')).toHaveCount(0);
  await vs.closeAll();
});
