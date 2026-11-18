// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 7.4: the browser's macro list and the trust record.
import { describe, test, expect, beforeEach } from 'vitest';
import { browserMacros, putBrowserMacro, deleteBrowserMacro, isTrusted, trust, scriptHash, documentMacrosEnabled } from '../src/app/macros.js';

describe('the browser macro store', () => {
  beforeEach(() => localStorage.clear());

  test('macros are added, replaced by name, listed and deleted', () => {
    expect(browserMacros()).toEqual([]);
    putBrowserMacro({ name: 'a', script: '1' });
    putBrowserMacro({ name: 'b', script: '2' });
    expect(browserMacros().map((m) => m.name)).toEqual(['a', 'b']);
    putBrowserMacro({ name: 'a', script: '3' });
    expect(browserMacros()).toEqual([{ name: 'b', script: '2' }, { name: 'a', script: '3' }]);
    deleteBrowserMacro('b');
    expect(browserMacros()).toEqual([{ name: 'a', script: '3' }]);
  });

  test('a broken record reads as no macros', () => {
    localStorage.setItem('docx4j-editor.macros', '{not json');
    expect(browserMacros()).toEqual([]);
    localStorage.setItem('docx4j-editor.macros', JSON.stringify([{ name: 1 }, { name: 'ok', script: 'x' }]));
    expect(browserMacros()).toEqual([{ name: 'ok', script: 'x' }]);
  });

  test('a document macro is trusted by its script hash: the same script again asks nothing, a changed one does', async () => {
    expect(await isTrusted('a')).toBe(false);
    await trust('a');
    expect(await isTrusted('a')).toBe(true);
    expect(await isTrusted('a ')).toBe(false);
    expect(await scriptHash('a')).toHaveLength(64);
  });

  test('document macros are on unless the deployment turned them off', () => {
    expect(documentMacrosEnabled()).toBe(true);
  });
});
