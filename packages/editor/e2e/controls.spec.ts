// SPDX-License-Identifier: GPL-3.0-or-later
//
// A content control's own properties on the page (ED-005 section 12.60): drawn as its appearance says outside the
// Template perspective (a bounding box at the caret with its title, start and end tags, nothing), in its own colour;
// its lock refusing an edit, saying why, with what was typed drawn back; typing replacing its placeholder.
import { test, expect, type Page } from '@playwright/test';
import { openFixture } from './helpers.js';

const body = '.docx-editor .ProseMirror[aria-label="Document"]';

type View = { caretAt(pos: number, to?: number): void; run(name: string, args?: unknown): boolean; state: { doc: { child(i: number): { textContent: string } } } };

/** Run in the page with the live view. */
async function inView<T>(page: Page, fn: (v: View) => T): Promise<T> {
  return page.evaluate(`(${fn.toString()})(window.docx4jEditor)`) as Promise<T>;
}

const pseudo = (el: Element, [which, prop]: ['::before' | '::after', 'content' | 'borderTopColor']): string => getComputedStyle(el, which)[prop];

test('a control drawn as its appearance says, in its colour, in Edit and in Template', async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  // "Hello" (after the bookmark's two markers) in a rich text control titled "Greeting".
  expect(await inView(page, (v) => { v.caretAt(3, 8); return v.run('contentControl.insert', { kind: 'RichText', title: 'Greeting' }); })).toBe(true);
  const control = page.locator(`${body} .sdt-inline`);
  await expect(control).toHaveAttribute('data-appearance', 'BoundingBox');
  // Elsewhere: no box. The caret in it: the box, and the tab with its title.
  await inView(page, (v) => v.caretAt(20));
  expect(await control.evaluate((el) => getComputedStyle(el).outlineStyle)).toBe('none');
  await inView(page, (v) => v.caretAt(6));
  await expect(control).toHaveClass(/sdt-current/);
  expect(await control.evaluate((el) => getComputedStyle(el).outlineStyle)).toBe('solid');
  expect(await control.evaluate((el) => getComputedStyle(el, '::before').content)).toBe('"Greeting"');
  // Start and end tags, always, in the control's own colour.
  expect(await inView(page, (v) => v.run('contentControl.properties', { appearance: 'Tags', color: '#C00000' }))).toBe(true);
  await inView(page, (v) => v.caretAt(20));
  await expect(control).toHaveAttribute('data-appearance', 'Tags');
  expect(await control.evaluate(pseudo, ['::before', 'content'] as ['::before', 'content'])).toBe('"Greeting"');
  expect(await control.evaluate(pseudo, ['::after', 'content'] as ['::after', 'content'])).toBe('"Greeting"');
  expect(await control.evaluate(pseudo, ['::before', 'borderTopColor'] as ['::before', 'borderTopColor'])).toBe('rgb(192, 0, 0)');
  // Hidden: nothing, even at the caret.
  await inView(page, (v) => v.caretAt(6));
  expect(await inView(page, (v) => v.run('contentControl.properties', { appearance: 'Hidden' }))).toBe(true);
  await expect(control).toHaveAttribute('data-appearance', 'Hidden');
  expect(await control.evaluate((el) => getComputedStyle(el).outlineStyle)).toBe('none');
  expect(await control.evaluate(pseudo, ['::before', 'content'] as ['::before', 'content'])).toBe('none');
  // The Template perspective honours it too (ED-005 section 12.61): hidden, nothing; tags with the label; a bounding box,
  // Template's authoring view (the role's underline and the label's badge), the caret in it or not.
  await page.getByTestId('perspective-template').click();
  await expect(page.locator('.page-column.show-control-tags')).toHaveCount(1);
  expect(await control.evaluate((el) => getComputedStyle(el).boxShadow)).toBe('none');
  expect(await control.evaluate(pseudo, ['::before', 'content'] as ['::before', 'content'])).toBe('none');
  await inView(page, (v) => v.caretAt(6));
  expect(await inView(page, (v) => v.run('contentControl.properties', { appearance: 'Tags' }))).toBe(true);
  expect(await control.evaluate(pseudo, ['::before', 'content'] as ['::before', 'content'])).toBe('"Greeting"');
  expect(await control.evaluate(pseudo, ['::after', 'content'] as ['::after', 'content'])).toBe('"Greeting"');
  expect(await control.evaluate(pseudo, ['::before', 'borderTopColor'] as ['::before', 'borderTopColor'])).toBe('rgb(192, 0, 0)');
  expect(await control.evaluate((el) => getComputedStyle(el).boxShadow)).toBe('none');
  expect(await inView(page, (v) => v.run('contentControl.properties', { appearance: 'BoundingBox' }))).toBe(true);
  await inView(page, (v) => v.caretAt(20));
  expect(await control.evaluate((el) => getComputedStyle(el).boxShadow)).not.toBe('none');
  expect(await control.evaluate(pseudo, ['::before', 'content'] as ['::before', 'content'])).toBe('"Greeting"');
  expect(await control.evaluate(pseudo, ['::after', 'content'] as ['::after', 'content'])).toBe('none');
});

test('a locked control refuses typing in it, saying why, and the page keeps its text; typing over a placeholder replaces it', async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  expect(await inView(page, (v) => { v.caretAt(3, 8); return v.run('contentControl.insert', { kind: 'RichText' }); })).toBe(true);
  expect(await inView(page, (v) => { v.caretAt(6); return v.run('contentControl.properties', { cannotEdit: true }); })).toBe(true);
  const control = page.locator(`${body} .sdt-inline`);
  await inView(page, (v) => v.caretAt(6));
  await page.keyboard.type('x');
  await page.keyboard.press('Backspace');
  await expect(page.getByTestId('status')).toBeVisible();
  await expect(page.getByTestId('status')).toHaveText('The contents of this content control cannot be edited');
  await expect(control).toHaveText('Hello');
  expect(await inView(page, (v) => v.state.doc.child(0).textContent)).toMatch(/^Hello World/);
  // Typing just after it (its end at 10) is not in it.
  await inView(page, (v) => v.caretAt(10));
  await page.keyboard.type('!');
  expect(await inView(page, (v) => v.state.doc.child(0).textContent)).toMatch(/^Hello! World/);
  // A control showing its placeholder: what is typed replaces it.
  expect(await inView(page, (v) => { v.caretAt(20); return v.run('contentControl.insert', { kind: 'RichText' }); })).toBe(true);
  const placeholder = page.locator(`${body} .sdt-inline`).nth(1);
  await expect(placeholder).toHaveText('Click or tap here to enter text.');   // Word's default entry's text (ED-008 section 8.8 item 14)
  await inView(page, (v) => v.caretAt(22));
  await page.keyboard.type('Hi');
  await expect(placeholder).toHaveText('Hi');
  // The caret inside it at its end: what is typed goes in it; outside, after it, it does not (Word's two places).
  const edges = await inView(page, (v) => {
    let end = -1;
    let seen = 0;
    (v.state.doc as unknown as { child(i: number): { forEach(f: (c: { type: { name: string }; nodeSize: number }, off: number) => void): void } }).child(0).forEach((c, off) => {
      if (c.type.name === 'sdtInline' && ++seen === 2) end = 1 + off + c.nodeSize - 1;
    });
    return end;
  });
  await inView(page, new Function('v', `v.caretAt(${edges})`) as (v: View) => void);
  await page.keyboard.type('!');
  await expect(placeholder).toHaveText('Hi!');
  await inView(page, new Function('v', `v.caretAt(${edges + 2})`) as (v: View) => void);
  await page.keyboard.type('?');
  await expect(placeholder).toHaveText('Hi!');
});

test("a control's borders are caret stops, as Word's are: Right at its end steps out after it, Left at its start before it, and a step from outside goes in at the edge", async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  // A plain text control over the last three characters of the first paragraph, so it ends where the paragraph does.
  const at = await inView(page, (v) => {
    const p = v.state.doc.child(0) as unknown as { content: { size: number }; forEach(f: (c: { type: { name: string }; nodeSize: number }, off: number) => void): void };
    const end = 1 + p.content.size;
    v.caretAt(end - 3, end);
    v.run('contentControl.insert', { kind: 'PlainText' });
    let before = -1;
    let after = -1;
    (v.state.doc.child(0) as unknown as typeof p).forEach((c, off) => { if (c.type.name === 'sdtInline') { before = 1 + off; after = 1 + off + c.nodeSize; } });
    return { before, after, innerStart: before + 1, innerEnd: after - 1 };
  });
  const control = page.locator(`${body} .sdt-inline`);
  const text = await control.textContent();
  // The browser's own moves are read by the view after the press returns, so every read polls.
  const caret = (): Promise<number> => inView(page, (v) => (v.state as unknown as { selection: { from: number } }).selection.from);
  // Inside at its end, where the browser has nowhere further to go: Right steps out after it.
  await inView(page, new Function('v', `v.caretAt(${at.innerEnd})`) as (v: View) => void);
  await page.keyboard.press('ArrowRight');
  await expect.poll(caret).toBe(at.after);
  await page.keyboard.type('!');
  await expect(control).toHaveText(text!);
  expect(await inView(page, (v) => v.state.doc.child(0).textContent)).toMatch(/!$/);
  // Back in at its end, along its text to its start, then out before it.
  await page.keyboard.press('ArrowLeft');
  await page.keyboard.press('ArrowLeft');
  await expect.poll(caret).toBe(at.innerEnd);
  for (let i = 1; i <= 3; i++) {
    await page.keyboard.press('ArrowLeft');
    await expect.poll(caret).toBe(at.innerEnd - i);   // read before the next press: a key sent before the browser's move is reported sees the old caret
  }
  await page.keyboard.press('ArrowLeft');
  await expect.poll(caret).toBe(at.before);
  await page.keyboard.type('?');
  await expect(control).toHaveText(text!);
  expect(await inView(page, (v) => v.state.doc.child(0).textContent)).toMatch(new RegExp(`\\?${text}!$`));
  // From before it (one further now, after the typed character), Right goes in at its start, before the first character.
  await page.keyboard.press('ArrowRight');
  await expect.poll(caret).toBe(at.innerStart + 1);
});

test("Enter at a control's inner end, where the paragraph ends, makes an empty paragraph after it and no copy of the control (ED-005 section 12.123)", async ({ page }) => {
  await openFixture(page, 'HelloWordOnline.docx');
  // Jason's case (tmp/testcase_sdt-replicates.docx): a plain text control ending the paragraph, the caret between it and the mark.
  const innerEnd = await inView(page, (v) => {
    const p = v.state.doc.child(0) as unknown as { content: { size: number }; forEach(f: (c: { type: { name: string }; nodeSize: number }, off: number) => void): void };
    const end = 1 + p.content.size;
    v.caretAt(end - 3, end);
    v.run('contentControl.insert', { kind: 'PlainText' });
    let at = -1;
    (v.state.doc.child(0) as unknown as typeof p).forEach((c, off) => { if (c.type.name === 'sdtInline') at = 1 + off + c.nodeSize - 1; });
    return at;
  });
  const controls = page.locator(`${body} .sdt-inline`);
  await expect(controls).toHaveCount(1);
  const paragraphs = await inView(page, (v) => (v.state.doc as unknown as { childCount: number }).childCount);
  await inView(page, new Function('v', `v.caretAt(${innerEnd})`) as (v: View) => void);
  // Inside the control, where the Word keymap's Enter had refused and ProseMirror's own split the control.
  expect(await inView(page, (v) => (v.state as unknown as { selection: { $from: { parent: { type: { name: string } } } } }).selection.$from.parent.type.name)).toBe('sdtInline');
  await page.keyboard.press('Enter');
  await expect.poll(() => inView(page, (v) => (v.state.doc as unknown as { childCount: number }).childCount)).toBe(paragraphs + 1);
  await expect(controls).toHaveCount(1);
  expect(await inView(page, (v) => v.state.doc.child(1).textContent)).toBe('');
  // Typed in the new paragraph, outside the control.
  await page.keyboard.type('Next');
  expect(await inView(page, (v) => v.state.doc.child(1).textContent)).toBe('Next');
  await expect(controls).toHaveCount(1);
});
