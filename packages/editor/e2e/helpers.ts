// SPDX-License-Identifier: GPL-3.0-or-later
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import type { Page } from '@playwright/test';
import { ZipPartStore } from '@docx4j/core-ts';

export const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

export function fixturePath(name: string): string {
  return join(fixturesDir, name);
}

export async function fixture(name: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(fixturePath(name)));
}

/**
 * Open a fixture through the file input and wait for the editor and the toolbar: the toolbar
 * mounts one render after the view and pushes the page down, so a click measured before it
 * appears lands on the wrong element.
 */
export async function openFixture(page: Page, name: string): Promise<void> {
  await page.goto('/');
  await page.getByTestId('open').setInputFiles(fixturePath(name));
  await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await page.getByTestId('toolbar').waitFor();
}

/** Open bytes (a download, a built file) through the file input, with the same waits as `openFixture`. */
export async function openBytes(page: Page, name: string, bytes: Uint8Array): Promise<void> {
  await page.goto('/');
  await page.getByTestId('open').setInputFiles({ name, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(bytes) });
  await page.getByTestId('editor').locator('.ProseMirror[aria-label="Document"]').waitFor();
  await page.getByTestId('toolbar').waitFor();
}

/** Click Download and return the saved bytes. */
export async function download(page: Page): Promise<Uint8Array> {
  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByTestId('download').click()]);
  const path = await dl.path();
  return new Uint8Array(await readFile(path));
}

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** ED-002 section 3.7's oracle: same entries, byte-identical parts except the regenerated and the expected ones. */
export function comparePackages(before: Uint8Array, after: Uint8Array, expectChanged: string[] = []): { name: string; reason: string }[] {
  const a = new ZipPartStore(before);
  const b = new ZipPartStore(after);
  const an = new Set(a.partNames());
  const bn = new Set(b.partNames());
  const out: { name: string; reason: string }[] = [];
  for (const n of an) if (!bn.has(n)) out.push({ name: n, reason: 'missing' });
  for (const n of bn) if (!an.has(n)) out.push({ name: n, reason: 'added' });
  for (const n of an) {
    if (!bn.has(n) || n === '[Content_Types].xml' || n.endsWith('.rels') || expectChanged.includes(n)) continue;
    if (!bytesEqual(a.loadSync(n), b.loadSync(n))) out.push({ name: n, reason: 'bytes' });
  }
  return out;
}

/**
 * Put the caret at the start of block `i` through the view (exposed on `window` for tests and
 * E1.d's console). Tests that only need a position use this; the typing tests click for real.
 */
export async function caretInBlock(page: Page, i: number, offset = 0): Promise<void> {
  await page.evaluate(([index, off]) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { nodeSize: number } } }; caretAt(pos: number): void } }).docx4jEditor;
    let pos = 0;
    for (let k = 0; k < index; k++) pos += v.state.doc.child(k).nodeSize;
    v.caretAt(pos + 1 + off);
  }, [i, offset] as const);
}

/** Put the caret at the end of block `i` through the view (an End key pressed right after `caretInBlock` races the browser's selection report). */
export async function caretAtBlockEnd(page: Page, i: number): Promise<void> {
  await page.evaluate((index) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { nodeSize: number; content: { size: number } } } }; caretAt(pos: number): void } }).docx4jEditor;
    let pos = 0;
    for (let k = 0; k < index; k++) pos += v.state.doc.child(k).nodeSize;
    v.caretAt(pos + 1 + v.state.doc.child(index).content.size);
  }, i);
}

/** Select the whole text of block `i` through the view. */
export async function selectBlock(page: Page, i: number): Promise<void> {
  await page.evaluate((index) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { nodeSize: number; content: { size: number } } } }; caretAt(pos: number, to?: number): void } }).docx4jEditor;
    let pos = 0;
    for (let k = 0; k < index; k++) pos += v.state.doc.child(k).nodeSize;
    v.caretAt(pos + 1, pos + 1 + v.state.doc.child(index).content.size);
  }, i);
}

/** Select from the start of block `from` to the end of block `to` through the view. */
export async function selectBlocks(page: Page, from: number, to: number): Promise<void> {
  await page.evaluate(([a, b]) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { nodeSize: number; content: { size: number } } } }; caretAt(pos: number, to?: number): void } }).docx4jEditor;
    let pos = 0;
    let start = 0;
    for (let k = 0; k <= b; k++) { if (k === a) start = pos + 1; pos += v.state.doc.child(k).nodeSize; }
    v.caretAt(start, pos - 1);
  }, [from, to] as const);
}

/** Put the caret at the start of the first paragraph of cell (`r`, `c`, by node index) of the table at block `i`. */
export async function caretInCell(page: Page, i: number, r: number, c: number): Promise<void> {
  await page.evaluate(([block, row, col]) => {
    const v = (window as unknown as { docx4jEditor: { state: { doc: { child(i: number): { nodeSize: number; child(i: number): { nodeSize: number; child(i: number): { nodeSize: number } } } } }; caretAt(pos: number): void } }).docx4jEditor;
    let pos = 0;
    for (let k = 0; k < block; k++) pos += v.state.doc.child(k).nodeSize;
    const table = v.state.doc.child(block);
    pos += 1;
    for (let k = 0; k < row; k++) pos += table.child(k).nodeSize;
    const tr = table.child(row);
    pos += 1;
    for (let k = 0; k < col; k++) pos += tr.child(k).nodeSize;
    v.caretAt(pos + 2);
  }, [i, r, c] as const);
}

/** Set the console's script through the live Monaco editor (`window.docx4jConsole`): typing a multi-line script through the keyboard goes through Monaco's auto-closing and auto-indent, which no test wants. */
export async function setConsoleScript(page: Page, script: string): Promise<void> {
  await page.waitForFunction(() => (window as unknown as { docx4jConsole?: unknown }).docx4jConsole !== undefined);
  await page.evaluate((s) => { (window as unknown as { docx4jConsole: { setValue(v: string): void } }).docx4jConsole.setValue(s); }, script);
}

/** The text of a Monaco XML editor (reveal codes, the Styles tab, the parts panel; ED-005 section 3.8) by its test id. */
export async function xmlEditorValue(page: Page, testId = 'xml-editor'): Promise<string> {
  await page.waitForFunction((id) => (window as unknown as { docx4jXmlEditors?: Record<string, unknown> }).docx4jXmlEditors?.[id] !== undefined, testId);
  return page.evaluate((id) => (window as unknown as { docx4jXmlEditors: Record<string, { getValue(): string }> }).docx4jXmlEditors[id]!.getValue(), testId);
}

/** Set a Monaco XML editor's text through the live instance: typed text would go through Monaco's auto-closing. */
export async function setXmlEditor(page: Page, text: string, testId = 'xml-editor'): Promise<void> {
  await page.waitForFunction((id) => (window as unknown as { docx4jXmlEditors?: Record<string, unknown> }).docx4jXmlEditors?.[id] !== undefined, testId);
  await page.evaluate(([id, t]) => { (window as unknown as { docx4jXmlEditors: Record<string, { setValue(v: string): void }> }).docx4jXmlEditors[id as string]!.setValue(t as string); }, [testId, text] as const);
}

/** A capture's text, its doubled encoding undone as the jsdom test does (`paste/README.md`). */
export async function capture(name: string): Promise<string> {
  let text = (await readFile(join(fixturesDir, 'paste', 'results', `${name}.html`))).toString('utf8');
  text = text.replace(/^\uFEFF/, '');
  // The clipboard's UTF-8 read as Windows-1252 and written out as UTF-8 again: a non-breaking space
  // arrived as A-circumflex followed by one, and a Japanese font name in the head as a run of such pairs.
  if (/[\u00c2\u00c3][\u0080-\u00bf\u20ac\u201a\u0192\u201e\u2026\u2020\u2021\u02c6\u2030\u0160\u2039\u0152\u017d\u2018\u2019\u201c\u201d\u2022\u2013\u2014\u02dc\u2122\u0161\u203a\u0153\u017e\u0178]/.test(text)) text = fromCp1252(text);
  return text;
}

/** The Windows-1252 characters above 0x7F that are not Latin-1, by their byte. */
const CP1252: Record<string, number> = { '\u20ac': 0x80, '\u201a': 0x82, '\u0192': 0x83, '\u201e': 0x84, '\u2026': 0x85, '\u2020': 0x86, '\u2021': 0x87, '\u02c6': 0x88, '\u2030': 0x89, '\u0160': 0x8a, '\u2039': 0x8b, '\u0152': 0x8c, '\u017d': 0x8e, '\u2018': 0x91, '\u2019': 0x92, '\u201c': 0x93, '\u201d': 0x94, '\u2022': 0x95, '\u2013': 0x96, '\u2014': 0x97, '\u02dc': 0x98, '\u2122': 0x99, '\u0161': 0x9a, '\u203a': 0x9b, '\u0153': 0x9c, '\u017e': 0x9e, '\u0178': 0x9f };

/** Text that was UTF-8 bytes read as Windows-1252: the bytes back, decoded once. */
function fromCp1252(text: string): string {
  const bytes = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) { const c = text.charCodeAt(i); bytes[i] = c < 0x100 ? c : (CP1252[text[i]!] ?? 0x3f); }
  return Buffer.from(bytes).toString('utf8');
}

/**
 * The pages of a PDF Chromium wrote (`page.pdf`), in order, with their sizes in points rounded: its page
 * objects and their `MediaBox`es are plain in the file (no object streams), the order the root `Kids`'.
 * Text is not read: Chromium encodes it as glyphs in compressed streams (ED-006 section 9.1 used poppler).
 */
export function pdfPages(bytes: Uint8Array): { width: number; height: number }[] {
  const s = Buffer.from(bytes).toString('latin1');
  const objects = new Map<string, string>();
  for (const m of s.matchAll(/(\d+) 0 obj([\s\S]*?)endobj/g)) objects.set(m[1]!, m[2]!);
  const box = (dict: string): number[] | undefined => /\/MediaBox\s*\[([^\]]*)\]/.exec(dict)?.[1]!.trim().split(/\s+/).map(Number);
  const out: { width: number; height: number }[] = [];
  const walk = (id: string, inherited: number[] | undefined): void => {
    const dict = objects.get(id) ?? '';
    const own = box(dict) ?? inherited;
    const kids = /\/Kids\s*\[([^\]]*)\]/.exec(dict);
    if (kids && /\/Type\s*\/Pages/.test(dict)) { for (const k of kids[1]!.matchAll(/(\d+) 0 R/g)) walk(k[1]!, own); return; }
    if (own) out.push({ width: Math.round(own[2]! - own[0]!), height: Math.round(own[3]! - own[1]!) });
  };
  const root = /\/Root (\d+) 0 R/.exec(s)?.[1];
  const pages = root ? /\/Pages (\d+) 0 R/.exec(objects.get(root) ?? '')?.[1] : undefined;
  if (pages) walk(pages, undefined);
  return out;
}
