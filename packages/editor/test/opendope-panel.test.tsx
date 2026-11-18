// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 4.6 in jsdom (step B4): the OpenDoPE tab lists the XPaths entries with their
// values and booleans and the conditions with their verdicts; an entry is added (its expression
// checked by evaluating it, REQ-014), edited and deleted (refused while in use, naming the
// users); a condition is built in the operator-tree editor, a cycle refused (REQ-029), a
// used condition kept; the boolean conversion mode is changed; a plain document gets its XPaths
// part on the first entry.
import { describe, test, expect } from 'vitest';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { EditorDocument } from '@docx4j/editor-model';
import { DocxEditorView } from '../src/view/DocxEditorView.js';
import OpenDopePanel from '../src/app/panels/template/OpenDopePanel.js';
import type { TabContext } from '../src/app/perspectives.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');

async function mount(name: string, prepare?: (session: EditorDocument) => void): Promise<{ session: EditorDocument; editor: DocxEditorView; root: Root; host: HTMLElement; status: string[] }> {
  const session = await EditorDocument.open(new Uint8Array(await readFile(join(fixtures, name))));
  prepare?.(session);
  const container = document.createElement('div');
  document.body.appendChild(container);
  const editor = new DocxEditorView(container, session);
  const host = document.createElement('div');
  document.body.appendChild(host);
  const root = createRoot(host);
  const status: string[] = [];
  const ctx: TabContext = { editor, active: editor, session, perspective: { id: 'template', panel: { open: true, tab: 'opendope' } }, consoleScript: null, onConsoleScript: () => {}, onSaveMacro: () => {}, onStatus: (s) => status.push(s), onTab: () => {} };
  await act(async () => { root.render(<OpenDopePanel {...ctx} />); });
  await until(() => host.querySelector('[data-testid="od-xpaths"]') !== null);
  return { session, editor, root, host, status };
}

async function until(check: () => boolean, ms = 5000): Promise<void> {
  const start = Date.now();
  while (!check()) {
    if (Date.now() - start > ms) throw new Error('timed out');
    await act(async () => { await new Promise((r) => setTimeout(r, 20)); });
  }
}

const q = <T extends Element>(host: HTMLElement, testId: string): T => host.querySelector<T>(`[data-testid="${testId}"]`)!;

async function setValue(el: HTMLInputElement | HTMLSelectElement, value: string): Promise<void> {
  const proto = el instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
  const setter = Object.getOwnPropertyDescriptor(proto, 'value')!.set!;
  await act(async () => {
    setter.call(el, value);
    el.dispatchEvent(new Event(el instanceof HTMLSelectElement ? 'change' : 'input', { bubbles: true }));
  });
}

const click = async (el: Element | null): Promise<void> => { await act(async () => { (el as HTMLElement).click(); }); };

describe('the OpenDoPE tab', () => {
  test('lists the entries and conditions with their evaluations; adds, edits and deletes an entry; refuses a bad expression and a used entry', async () => {
    const { session, host, root, editor, status } = await mount('invoice.docx');
    expect(host.querySelectorAll('[data-testid^="od-entry-x"]:not([data-testid$="-boolean"]):not([data-testid$="-value"]):not([data-testid$="-uses"])').length).toBe(6);
    expect(q(host, 'od-entry-x1').textContent).toContain('Joe Bloggs');
    expect(q(host, 'od-entry-x2').textContent).toContain('the row control');
    // How each entry is used, as badges: x2 a repeat, x1 a binding, x5 a condition's operand.
    expect(q(host, 'od-entry-x2-uses').querySelector('.badge')!.className).toBe('badge badge-repeat');
    expect(q(host, 'od-entry-x2').className).toContain('use-repeat');
    expect(q(host, 'od-entry-x1-uses').querySelector('.badge')!.textContent).toBe('bound');
    expect(q(host, 'od-entry-x5-uses').querySelector('.badge')!.textContent).toBe('condition');
    expect(q(host, 'od-entry-x5-boolean').textContent).toBe('true');
    expect(q(host, 'od-condition-c5-verdict').textContent).toBe('true (kept)');
    expect(q(host, 'od-condition-c6-verdict').textContent).toBe('false (dropped)');
    expect(q(host, 'od-version').textContent).toContain('none declared');
    // A new entry: a malformed expression is refused by evaluation (REQ-014), a good one is written.
    await click(q(host, 'od-entry-new'));
    expect(q<HTMLInputElement>(host, 'od-entry-id').value).toBe('x7');
    await setValue(q<HTMLInputElement>(host, 'od-entry-xpath'), '/invoice/[');
    await click(q(host, 'od-entry-apply'));
    expect(q(host, 'od-entry-check').textContent).toContain('REQ-014');
    expect(session.template.xpaths!.entries).toHaveLength(6);
    await setValue(q<HTMLInputElement>(host, 'od-entry-xpath'), '/invoice/items/item/price');
    await setValue(q<HTMLInputElement>(host, 'od-entry-name'), 'Prices');
    await click(q(host, 'od-entry-apply'));
    await until(() => host.querySelector('[data-testid="od-entry-x7"]') !== null);
    expect(session.template.xpaths!.entries[6]).toMatchObject({ id: 'x7', xpath: '/invoice/items/item/price', storeItemID: session.template.dataPart.part!.id, attrs: { name: 'Prices' } });
    expect(q(host, 'od-entry-x7').textContent).toContain('$20');
    expect(q(host, 'od-entry-x7').textContent).toContain('unused');
    expect(q(host, 'od-entry-x7-value').className).toBe('meta');
    expect([...session.touchedParts]).toContain('/customXml/item4.xml');
    expect(status.at(-1)).toContain('no undo');
    // Edited: the expression changes, the other attributes stay.
    await click(q(host, 'od-entry-x7').querySelector('button.link'));
    await setValue(q<HTMLInputElement>(host, 'od-entry-xpath'), '/invoice/items/item/name');
    await click(q(host, 'od-entry-apply'));
    await until(() => q(host, 'od-entry-x7').textContent!.includes('apples'));
    // An entry that selects nothing stands out (bold orange); a valid expression over a missing element, not an error.
    await click(q(host, 'od-entry-x7').querySelector('button.link'));
    await setValue(q<HTMLInputElement>(host, 'od-entry-xpath'), '/invoice/items/item/colour');
    await click(q(host, 'od-entry-apply'));
    await until(() => q(host, 'od-entry-x7-value')?.textContent === 'selects nothing');
    expect(q(host, 'od-entry-x7-value').className).toBe('selects-nothing');
    await click(q(host, 'od-entry-x7').querySelector('button.link'));
    await setValue(q<HTMLInputElement>(host, 'od-entry-xpath'), '/invoice/items/item/name');
    await click(q(host, 'od-entry-apply'));
    await until(() => q(host, 'od-entry-x7').textContent!.includes('apples'));
    expect(session.template.xpaths!.entries[6]).toMatchObject({ xpath: '/invoice/items/item/name', attrs: { name: 'Prices' } });
    // A used entry stays, its users named; an unused one goes.
    await click(host.querySelector('[aria-label="Delete entry x2"]'));
    expect(q(host, 'od-error').textContent).toContain('the row control');
    expect(session.template.xpaths!.entries).toHaveLength(7);
    await click(host.querySelector('[aria-label="Delete entry x7"]'));
    await until(() => host.querySelector('[data-testid="od-entry-x7"]') === null);
    expect(session.template.xpaths!.entries).toHaveLength(6);
    // The mode changed: re-evaluated in xpath1.
    await setValue(q<HTMLSelectElement>(host, 'od-mode'), 'xpath1');
    expect(session.template.mode).toBe('xpath1');
    await until(() => host.querySelector('[data-testid="od-xpaths"] h3')!.textContent!.includes('xpath1'));
    root.unmount();
    editor.destroy();
  });

  test('builds a condition in the operator tree, refuses a cycle and a delete of a used condition', async () => {
    const { session, host, root, editor } = await mount('invoice.docx');
    await click(q(host, 'od-condition-new'));
    expect(q<HTMLInputElement>(host, 'od-condition-id').value).toBe('c1');
    await setValue(q<HTMLInputElement>(host, 'od-condition-name'), 'no spam');
    // not(x6): the root becomes `not`, its operand picks x6.
    await setValue(q<HTMLSelectElement>(host, 'operand-root-op'), 'not');
    await setValue(q<HTMLSelectElement>(host, 'operand-root-0-id'), 'x6');
    expect(q(host, 'od-condition-form').textContent).toContain('not(x6');
    await click(q(host, 'od-condition-apply'));
    await until(() => host.querySelector('[data-testid="od-condition-c1"]') !== null);
    expect(session.template.conditions!.conditions[2]).toEqual({ id: 'c1', attrs: { name: 'no spam' }, operand: { op: 'not', operand: { op: 'xpathref', id: 'x6' } } });
    expect(q(host, 'od-condition-c1-verdict').textContent).toBe('true (kept)');
    expect([...session.touchedParts]).toContain('/customXml/item1.xml');
    // and(x5, c1) with the operands reordered and one removed.
    await click(q(host, 'od-condition-new'));
    await setValue(q<HTMLSelectElement>(host, 'operand-root-op'), 'and');
    await click(q(host, 'operand-root-add'));
    await setValue(q<HTMLSelectElement>(host, 'operand-root-1-op'), 'conditionref');
    await setValue(q<HTMLSelectElement>(host, 'operand-root-1-id'), 'c1');
    await click(q(host, 'operand-root-add'));
    await click(host.querySelector('[aria-label="Move operand 3 up"]'));
    await click(q(host, 'operand-root-1-remove'));
    await setValue(q<HTMLSelectElement>(host, 'operand-root-0-id'), 'x5');
    expect(q(host, 'od-condition-form').textContent).toContain('and(x5');
    await click(q(host, 'od-condition-apply'));
    await until(() => host.querySelector('[data-testid="od-condition-c2"]') !== null);
    expect(session.template.conditions!.conditions[3]!.operand).toEqual({ op: 'and', operands: [{ op: 'xpathref', id: 'x5' }, { op: 'conditionref', id: 'c1' }] });
    // A cycle: c1 made to refer to c2, which refers to c1 (REQ-029).
    await click(q(host, 'od-condition-c1').querySelector('button.link'));
    await setValue(q<HTMLSelectElement>(host, 'operand-root-op'), 'conditionref');
    await setValue(q<HTMLSelectElement>(host, 'operand-root-id'), 'c2');
    await click(q(host, 'od-condition-apply'));
    expect(q(host, 'od-error').textContent).toContain('cycle');
    expect(session.template.conditions!.conditions[2]!.operand).toEqual({ op: 'not', operand: { op: 'xpathref', id: 'x6' } });
    // c5 is on a block control: it stays.
    await click(host.querySelector('[aria-label="Delete condition c5"]'));
    expect(q(host, 'od-error').textContent).toContain('the block control');
    expect(session.template.conditions!.conditions).toHaveLength(4);
    await click(host.querySelector('[aria-label="Delete condition c2"]'));
    await until(() => host.querySelector('[data-testid="od-condition-c2"]') === null);
    root.unmount();
    editor.destroy();
  });

  test('a plain document gets its XPaths part, at version 3.0 in the xpath2 mode, on the first entry', async () => {
    const { session, host, root, editor } = await mount('HelloWordOnline.docx', (s) => { s.pkg.customXmlParts.add('<data><name>Ann</name></data>'); });
    expect(q(host, 'od-root').textContent).toContain('created when the first entry is added');
    expect(host.querySelector('[data-testid="od-condition-new"]')!.hasAttribute('disabled')).toBe(true);
    await click(q(host, 'od-entry-new'));
    expect(q<HTMLInputElement>(host, 'od-entry-id').value).toBe('x1');
    await setValue(q<HTMLInputElement>(host, 'od-entry-xpath'), '/data/name');
    await click(q(host, 'od-entry-apply'));
    await until(() => host.querySelector('[data-testid="od-entry-x1"]') !== null);
    expect(session.template.xpaths).toMatchObject({ version: '3.0', booleanConversion: 'xpath2', entries: [{ id: 'x1', xpath: '/data/name' }] });
    expect(q(host, 'od-version').textContent).toBe('3.0');
    expect(q(host, 'od-entry-x1').textContent).toContain('Ann');
    expect(host.querySelector('[data-testid="od-condition-new"]')!.hasAttribute('disabled')).toBe(false);
    root.unmount();
    editor.destroy();
  });
});
