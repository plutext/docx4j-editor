// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 7, E4.e step E4: comments edited in the margin. New Comment on a selection and at
// a caret (the name asked for first), the new comment's body taking the caret; Reply, Resolve,
// Reopen and Delete with undo; one's own comment edited in place and another's not; the saved
// file's comments.
import { test, expect, type Page } from '@playwright/test';
import { ZipPartStore, WordprocessingMLPackage } from '@docx4j/core-ts';
import { openFixture, openBytes, download, fixture, caretInBlock } from './helpers.js';

const body = '.docx-editor .ProseMirror[aria-label="Document"]';

type Win = { docx4jEditor: { caretAt(from: number, to?: number): void; state: { doc: { child(i: number): { nodeSize: number; textContent: string } } } } };

/** Select `[from, to)` of block `i`'s text through the view. */
async function selectInBlock(page: Page, i: number, from: number, to: number): Promise<void> {
  await page.evaluate(([index, a, b]) => {
    const v = (window as unknown as Win).docx4jEditor;
    let pos = 0;
    for (let k = 0; k < index; k++) pos += v.state.doc.child(k).nodeSize;
    v.caretAt(pos + 1 + a, pos + 1 + b);
  }, [i, from, to] as const);
}

const part = async (page: Page, name: string): Promise<string> => new TextDecoder().decode(new ZipPartStore(await download(page)).loadSync(name));

test('New Comment on a selection asks for a name first, puts the caret in the new comment, and the saved file has it', async ({ page }) => {
  await openFixture(page, 'comments-two.docx');
  await page.locator(body).focus();
  await selectInBlock(page, 0, 0, 3);   // "One"
  await page.getByTestId('new-comment').click();
  const name = page.getByTestId('tracking-name');
  await expect(name).toBeVisible();
  await name.fill('Reviewer One');
  await name.press('Enter');
  const bubbles = page.getByTestId('comment-bubble');
  await expect(bubbles).toHaveCount(3);
  const mine = page.locator('[data-testid="comment-bubble"]', { hasText: 'Reviewer One' });
  await expect(mine.locator('[data-testid="comment-body"][data-editable="true"]')).toHaveCount(1);
  // The caret is in the new comment: typing goes there.
  await page.keyboard.type('Why "One"?');
  await expect(mine.locator('[data-testid="comment-body"]')).toContainText('Why "One"?');
  await expect(page.locator('.docx-editor .ProseMirror .comment[data-comment="2"]').first()).toHaveText('One');
  const comments = await part(page, 'word/comments.xml');
  expect(comments).toMatch(/w:author="Reviewer One"/);
  expect(comments).toMatch(/Why "One"\?/);
});

test('New Comment far down a long document keeps the page where it was (Jason, 2026-10-17: it scrolled back to the first comment)', async ({ page }) => {
  // The two-comment fixture with 150 paragraphs after them, the name the fixture's comments carry, so the
  // earlier comments are one's own, editable bodies; then a comment far down with the dialog, and another without.
  const pkg = await WordprocessingMLPackage.load(await fixture('comments-two.docx'));
  const part = pkg.getMainDocumentPart();
  const xml = await part.getXml();
  const more = Array.from({ length: 150 }, (_, i) => `<w:p><w:r><w:t>Paragraph ${i + 3} of the long document.</w:t></w:r></w:p>`).join('');
  const patched = xml.replace('<w:sectPr>', `${more}<w:sectPr>`);
  expect(patched).not.toBe(xml);
  part.setXml(patched);
  await openBytes(page, 'comments-long.docx', await pkg.save());
  await page.locator(body).focus();
  const area = page.locator('.page-area');
  const comment = async (block: number, id: string, named: boolean): Promise<void> => {
    await selectInBlock(page, block, 0, 9);   // "Paragraph"
    const scrolled = await area.evaluate((el) => el.scrollTop);
    expect(scrolled).toBeGreaterThan(1000);
    await page.getByTestId('new-comment').click();
    if (named) {
      const name = page.getByTestId('tracking-name');
      await expect(name).toBeVisible();
      await name.fill('jharrop2');
      await name.press('Enter');
    }
    const mine = page.locator(`[data-testid="comment-bubble"][data-comment-id="${id}"]`);
    await expect(mine.locator('[data-testid="comment-body"][data-editable="true"]')).toHaveCount(1);
    await page.keyboard.type('Here');
    await expect(mine.locator('[data-testid="comment-body"]')).toContainText('Here');
    // The page stayed near the selection, and the bubble sits beside it.
    expect(Math.abs((await area.evaluate((el) => el.scrollTop)) - scrolled), `comment ${id}: scroll`).toBeLessThan(400);
    const anchor = await page.locator(`.docx-editor .ProseMirror .comment[data-comment="${id}"]`).first().boundingBox();
    const bubble = await mine.boundingBox();
    expect(Math.abs(bubble!.y - anchor!.y), `comment ${id}: bubble`).toBeLessThan(60);
  };
  await comment(100, '2', true);
  await comment(140, '3', false);
});

test('Ctrl+Alt+M at a caret comments on the word the caret is in', async ({ page }) => {
  await openFixture(page, 'comments-two.docx');
  await page.locator(body).focus();
  await caretInBlock(page, 0, 6);   // inside "paragraph"
  await page.keyboard.press('Control+Alt+m');
  const name = page.getByTestId('tracking-name');
  await expect(name).toBeVisible();
  await name.fill('Reviewer One');
  await name.press('Enter');
  await expect(page.getByTestId('comment-bubble')).toHaveCount(3);
  const marked = await page.locator('.docx-editor .ProseMirror .comment').evaluateAll((els) => els.filter((e) => (e as HTMLElement).dataset['comment'] === '2').map((e) => e.textContent).join(''));
  expect(marked).toBe('paragraph');
});

test("Reply, Resolve, Reopen and Delete with undo; one's own comment is edited in place, another's is not", async ({ page }) => {
  await openFixture(page, 'comments-two.docx');
  const first = page.locator('[data-comment-id="0"]');
  // Another's comment (jharrop2's): read-only.
  await expect(first.getByTestId('comment-body').first()).toBeVisible();
  await expect(first.getByTestId('comment-body').first()).not.toHaveAttribute('data-editable', 'true');
  // Reply.
  await first.getByTestId('comment-reply').click();
  await first.getByTestId('comment-reply-text').fill('A reply');
  await first.getByTestId('comment-post').click();
  const reply = first.locator('.comment-reply');
  await expect(reply).toContainText('A reply');
  await expect(reply.locator('[data-testid="comment-body"][data-editable="true"]')).toHaveCount(1);
  // One's own reply edited in place.
  await reply.locator('[data-testid="comment-body"] .ProseMirror').click();
  await page.keyboard.press('End');
  await page.keyboard.type(', edited');
  await expect(reply).toContainText('A reply, edited');
  // Resolve and Reopen.
  await first.getByTestId('comment-resolve').click();
  await expect(first).toHaveClass(/done/);
  await first.locator('summary').click();
  await first.getByTestId('comment-reopen').click();
  await expect(first).not.toHaveClass(/done/);
  // Delete the second thread, and undo.
  await page.locator('[data-comment-id="1"]').getByTestId('comment-delete').first().click();
  await expect(page.locator('[data-comment-id="1"]')).toHaveCount(0);
  await page.locator(body).focus();
  await page.keyboard.press('Control+z');
  await expect(page.locator('[data-comment-id="1"]')).toHaveCount(1);
  const ext = await part(page, 'word/commentsExtended.xml');
  expect(ext).toMatch(/w15:paraIdParent=/);
  const comments = await part(page, 'word/comments.xml');
  expect(comments).toMatch(/A reply, edited/);
});

test('in a room: a comment made on one page shows on the other, is replied to there, and both pages save the same comments', async ({ browser }) => {
  const contexts = await Promise.all([browser.newContext(), browser.newContext()]);
  await contexts[0]!.addInitScript(() => { try { localStorage.setItem('docx4j-editor.relay', '/relay'); localStorage.setItem('docx4j-editor.name', 'Author A'); } catch { /* ignore */ } });
  await contexts[1]!.addInitScript(() => { try { localStorage.setItem('docx4j-editor.relay', '/relay'); localStorage.setItem('docx4j-editor.name', 'Author B'); } catch { /* ignore */ } });
  const [a, b] = await Promise.all(contexts.map((c) => c.newPage()));
  await openFixture(a!, 'comments-two.docx');
  await a!.getByTestId('share').click();
  const link = await a!.getByTestId('share-link').inputValue();
  await a!.keyboard.press('Escape');
  await b!.goto(link);
  await b!.getByTestId('join-accept').click();
  await b!.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await expect(b!.getByTestId('room-status')).toHaveAttribute('data-status', 'connected', { timeout: 20_000 });
  // A comments on "One".
  await a!.locator(body).focus();
  await selectInBlock(a!, 0, 0, 3);
  await a!.getByTestId('new-comment').click();
  await a!.keyboard.type('From A');
  const fromA = (page: Page) => page.locator('[data-testid="comment-bubble"]', { hasText: 'From A' });
  await expect(fromA(b!)).toHaveCount(1, { timeout: 10_000 });
  // On B's page it is A's: read-only there. B replies.
  await expect(fromA(b!).getByTestId('comment-body').first()).not.toHaveAttribute('data-editable', 'true');
  await fromA(b!).getByTestId('comment-reply').click();
  await fromA(b!).getByTestId('comment-reply-text').fill('From B');
  await fromA(b!).getByTestId('comment-post').click();
  await expect(fromA(a!).locator('.comment-reply')).toContainText('From B', { timeout: 10_000 });
  await expect(fromA(a!).locator('.comment-reply')).toContainText('Author B');
  // Both pages save the same comments, with small ids.
  const comments = async (page: Page): Promise<string[]> => {
    const xml = new TextDecoder().decode(new ZipPartStore(await download(page)).loadSync('word/comments.xml'));
    return [...xml.matchAll(/<w:comment [^>]*>/g)].map((m) => [/w:id="(\d+)"/.exec(m[0])![1], /w:author="([^"]*)"/.exec(m[0])![1]].join(' ')).sort();
  };
  const [ca, cb] = [await comments(a!), await comments(b!)];
  expect(ca).toEqual(cb);
  expect(ca).toEqual(['0 jharrop2', '1 jharrop2', '2 Author A', '3 Author B']);
  await Promise.all(contexts.map((c) => c.close()));
});

// ED-005 section 12.113: an inline ref re-observed each bubble on every render, a fresh `observe` always reports, and the report
// drew the margin again: every frame, for ever, on any document with a comment, and each draw restarted a long render beside it.
test('the margin is at rest while nothing changes, and still follows what does', async ({ page }) => {
  await page.addInitScript(() => {
    const Real = window.ResizeObserver;
    (window as unknown as { resizeReports: number }).resizeReports = 0;
    window.ResizeObserver = class extends Real {
      constructor(callback: ResizeObserverCallback) {
        super((entries, observer) => { (window as unknown as { resizeReports: number }).resizeReports++; callback(entries, observer); });
      }
    };
  });
  await openFixture(page, 'comments-two.docx');
  const bubbles = page.getByTestId('comment-bubble');
  await expect(bubbles).toHaveCount(2);
  const reports = (): Promise<number> => page.evaluate(() => (window as unknown as { resizeReports: number }).resizeReports);
  await page.waitForTimeout(400);
  const settled = await reports();
  await page.waitForTimeout(600);
  expect(await reports()).toBe(settled);

  const box = async (i: number): Promise<{ top: number; bottom: number }> => {
    const b = (await bubbles.nth(i).boundingBox())!;
    return { top: b.y, bottom: b.y + b.height };
  };
  const anchorDistance = (): Promise<number> => page.evaluate(() => {
    const bubble = document.querySelector('[data-testid="comment-bubble"]')!.getBoundingClientRect();
    const anchor = document.querySelector('.docx-editor .comment')!.getBoundingClientRect();
    return bubble.top - anchor.top;
  });
  // A bubble that grows moves the one after it: the observer still reports a real change.
  await bubbles.first().getByTestId('comment-reply').click();
  await expect.poll(async () => (await box(1)).top >= (await box(0)).bottom).toBe(true);
  await bubbles.first().getByTestId('comment-reply').click();

  // A layout that moves the anchor with no change to the document: the page's growth is enough for the bubble to follow.
  const before = await anchorDistance();
  await page.addStyleTag({ content: '.ProseMirror[aria-label="Document"] > p:first-child { margin-block-start: 300px; }' });
  await expect.poll(anchorDistance).toBeCloseTo(before, 0);
});
