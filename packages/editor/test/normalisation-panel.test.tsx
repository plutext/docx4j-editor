// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-008 section 8.8: the Diagnostics tab's normalisation report, mounted under StrictMode. The report itself is
// `normalisation.test.ts`'s and the browser's: a package saved under jsdom cannot be reloaded (ED-003 section 11.2
// item 18), so here it is stubbed, and this holds what the panel does with one.
import { describe, test, expect, vi, beforeEach } from 'vitest';
import { act, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import type { EditorDocument } from '@docx4j/editor-model';
import { normalisationReport, type NormalisationReport } from '../src/app/normalisation.js';
import { NormalisationSection } from '../src/app/panels/NormalisationSection.js';

vi.mock('../src/app/normalisation.js', () => ({ normalisationReport: vi.fn() }));
(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const report = vi.mocked(normalisationReport);
const session = (): EditorDocument => ({}) as EditorDocument;
const FOUND: NormalisationReport = {
  values: 1053,
  parts: [{ partName: '/word/styles.xml', count: 1053, kinds: [
    { kind: 'w:shd/@w:fill', count: 1050, changes: [{ before: 'ffffff', after: 'FFFFFF', count: 1050 }] },
    { kind: 'w:highlight', count: 3, changes: [{ before: 'purple', after: null, count: 3 }] },
  ] }],
  notes: [{ code: 'mce.branches_differ', messages: ['/word/document.xml: the branches differ', '/word/document.xml: they differ again'] }],
};

async function mounted(first: EditorDocument) {
  const host = document.createElement('div');
  document.body.appendChild(host);
  const root = createRoot(host);
  const render = (s: EditorDocument) => act(async () => { root.render(<StrictMode><NormalisationSection session={s} /></StrictMode>); });
  await render(first);
  const button = (): HTMLButtonElement => host.querySelector<HTMLButtonElement>('[data-testid="normalisation-run"]')!;
  const run = () => act(async () => { button().click(); });
  return { host, render, button, run };
}

const q = (host: HTMLElement, id: string): Element | null => host.querySelector(`[data-testid="${id}"]`);

describe('the normalisation report in the Diagnostics tab', () => {
  beforeEach(() => { report.mockReset(); });

  test('is made on the button, never at open; each part\'s kinds fold, with their values and counts', async () => {
    report.mockResolvedValue(FOUND);
    const { host, run } = await mounted(session());
    expect(report).not.toHaveBeenCalled();
    expect(q(host, 'normalisation-report')).toBeNull();
    await run();
    expect(report).toHaveBeenCalledTimes(1);
    expect(q(host, 'normalisation-summary')!.textContent).toBe('Values not in normal form: 1,053');
    expect(host.querySelector('[data-part="/word/styles.xml"]')!.textContent).toBe('/word/styles.xml 1053');
    const fill = host.querySelector<HTMLDetailsElement>('details[data-kind="w:shd/@w:fill"]')!;
    expect(fill.open).toBe(false);
    expect(fill.querySelector('li')!.textContent).toBe('ffffff would be written as FFFFFF 1050');
    expect(host.querySelector('details[data-kind="w:highlight"] li')!.textContent).toBe('purple would be left out 3');
    expect(q(host, 'normalisation-notes')!.textContent).toBe('Notes from reading the file: 2');
    expect(host.querySelectorAll('details[data-kind="mce.branches_differ"] li').length).toBe(2);
  });

  test('a clean file says so and shows no empty heading for what was noted', async () => {
    report.mockResolvedValue({ values: 0, parts: [], notes: [] });
    const { host, run } = await mounted(session());
    await run();
    expect(q(host, 'normalisation-summary')!.textContent).toBe('Every value is in its normal form.');
    expect(q(host, 'normalisation-notes')).toBeNull();
  });

  test('the button waits while the report is made, and a failure is said', async () => {
    let fail!: (e: Error) => void;
    report.mockReturnValue(new Promise((_resolve, reject) => { fail = reject; }));
    const { host, button, run } = await mounted(session());
    await run();
    expect(button().disabled).toBe(true);
    expect(host.querySelector('[role="status"]')!.textContent).toBe('Reading the document…');
    await act(async () => { fail(new Error('no save today')); });
    expect(button().disabled).toBe(false);
    expect(q(host, 'normalisation-failed')!.textContent).toContain('no save today');
  });

  test('a report belongs to the document it read: another one starts without it, and a late answer for the first is not shown', async () => {
    report.mockResolvedValue(FOUND);
    const first = session();
    const { host, render, run } = await mounted(first);
    await run();
    expect(q(host, 'normalisation-report')).not.toBeNull();
    const second = session();
    await render(second);
    expect(q(host, 'normalisation-report')).toBeNull();

    let answer!: (r: NormalisationReport) => void;
    report.mockReturnValue(new Promise((resolve) => { answer = resolve; }));
    await render(first);
    await run();
    await render(second);
    await act(async () => { answer(FOUND); });
    expect(q(host, 'normalisation-report')).toBeNull();
  });
});
