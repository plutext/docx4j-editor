// SPDX-License-Identifier: GPL-3.0-or-later
//
// Names and addresses (ED-014 section 4.3): the Word types, the names Save As refuses, untitled names, and a link's
// address as P28 sorts it, Word's forms of a relative and a `file:` link included. And the palette's names of the commands
// the page names (section 4.7, P49), from the manifest as VS Code hands it over.
import { describe, expect, test } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { isMarkdownName, isSaveAsName, isWordName, linkTarget, markdownTarget, paletteNames, untitledName } from '../../src/names.js';
import { PAGE_COMMANDS } from '../../src/shared.js';

describe('the Word types', () => {
  test('the four, whatever their case', () => {
    for (const name of ['a.docx', 'b.DOTX', 'c.docm', 'd.Dotm']) expect(isWordName(name), name).toBe(true);
    for (const name of ['a.doc', 'a.md', 'a.xml', 'docx', '.docx', 'a.docx.bak']) expect(isWordName(name), name).toBe(false);
  });

  test('Save As writes a Word type or .xml (Flat OPC), and refuses any other name (P25)', () => {
    for (const name of ['a.docx', 'a.dotm', 'a.xml', 'A.XML']) expect(isSaveAsName(name), name).toBe(true);
    for (const name of ['notes.md', 'a.txt', 'a', 'a.doc']) expect(isSaveAsName(name), name).toBe(false);
  });

  test('a Markdown file opens as the Word document of its name (P26)', () => {
    expect(isMarkdownName('notes.md')).toBe(true);
    expect(isMarkdownName('notes.Markdown')).toBe(true);
    expect(markdownTarget('notes.md')).toBe('notes.docx');
    expect(markdownTarget('/a/b/read.me.markdown')).toBe('read.me.docx');
    expect(markdownTarget('README')).toBe('README.docx');
  });
});

describe('untitled names (P17)', () => {
  test('the name asked for, else the next number', () => {
    const open = new Set(['Document1.docx', 'Document2.docx', 'notes.docx']);
    const taken = (n: string): boolean => open.has(n);
    expect(untitledName('Document3.docx', taken)).toBe('Document3.docx');
    expect(untitledName('Document1.docx', taken)).toBe('Document3.docx');
    expect(untitledName('notes.docx', taken)).toBe('notes 2.docx');
    open.add('notes 2.docx');
    expect(untitledName('notes.docx', taken)).toBe('notes 3.docx');
  });
});

describe('links (P28)', () => {
  test('http, https and mailto go to VS Code\'s opener as they are', () => {
    expect(linkTarget('https://example.com/a?b#c')).toEqual({ kind: 'external', url: 'https://example.com/a?b#c' });
    expect(linkTarget('HTTP://example.com')).toEqual({ kind: 'external', url: 'HTTP://example.com' });
    expect(linkTarget('mailto:someone@example.com')).toEqual({ kind: 'external', url: 'mailto:someone@example.com' });
  });

  test('any other scheme is refused', () => {
    for (const [address, scheme] of [['command:workbench.action.reloadWindow', 'command:'], ['vscode:extension/x', 'vscode:'], ['javascript:alert(1)', 'javascript:'], ['ms-word:ofe|u|https://x/y.docx', 'ms-word:']] as const) {
      expect(linkTarget(address)).toEqual({ kind: 'refused', scheme });
    }
  });

  test('a relative link, as Word writes it with backslashes and percent escapes, its fragment and query dropped', () => {
    expect(linkTarget('other.docx')).toEqual({ kind: 'relative', path: 'other.docx' });
    expect(linkTarget('..\\Folder\\My%20File.docx')).toEqual({ kind: 'relative', path: '../Folder/My File.docx' });
    expect(linkTarget('./sub/./b.docx#bookmark')).toEqual({ kind: 'relative', path: 'sub/b.docx' });
    expect(linkTarget('#only-a-fragment')).toEqual({ kind: 'refused', scheme: '' });
  });

  test('file: links in Word\'s several forms, normalised', () => {
    expect(linkTarget('file:///C:\\Users\\a\\b.docx')).toEqual({ kind: 'file', uri: 'file:///C:/Users/a/b.docx' });
    expect(linkTarget('file:///c:/Users/a/b.docx')).toEqual({ kind: 'file', uri: 'file:///C:/Users/a/b.docx' });
    expect(linkTarget('file:C:/a.docx')).toEqual({ kind: 'file', uri: 'file:///C:/a.docx' });
    expect(linkTarget('file://server/share/a.docx')).toEqual({ kind: 'file', uri: 'file://server/share/a.docx' });
    // The four-slash UNC form (Fable's review, ED-014 section 9.11 finding 7).
    expect(linkTarget('file:////server/share/a.docx')).toEqual({ kind: 'file', uri: 'file://server/share/a.docx' });
    expect(linkTarget('file:///home/a/b.docx')).toEqual({ kind: 'file', uri: 'file:///home/a/b.docx' });
    expect(linkTarget('file:b.docx')).toEqual({ kind: 'relative', path: 'b.docx' });
  });

  test('an absolute path with no scheme: a drive, a UNC share or a POSIX path', () => {
    expect(linkTarget('c:\\a\\b.docx')).toEqual({ kind: 'path', path: 'C:/a/b.docx' });
    expect(linkTarget('\\\\server\\share\\b.docx')).toEqual({ kind: 'path', path: '//server/share/b.docx' });
    expect(linkTarget('/home/a/b.docx')).toEqual({ kind: 'path', path: '/home/a/b.docx' });
  });
});

describe('the palette\'s names (P49)', () => {
  const root = join(import.meta.dirname, '..', '..');
  /** The manifest as VS Code hands it over in English: every `%key%` replaced by package.nls.json's words. */
  function english(): unknown {
    const nls = JSON.parse(readFileSync(join(root, 'package.nls.json'), 'utf8')) as Record<string, string>;
    const text = readFileSync(join(root, 'package.json'), 'utf8').replace(/"%([^%"]+)%"/g, (_, k: string) => JSON.stringify(nls[k]));
    return JSON.parse(text);
  }

  test('from the English manifest: each command the page names as the page names it in English', () => {
    expect(paletteNames(english(), Object.keys(PAGE_COMMANDS))).toEqual(PAGE_COMMANDS);
  });

  test('from a translated one: a title VS Code translated is { value, original }, a category it left a string (section 9.25)', () => {
    const manifest = { contributes: { commands: [
      { command: 'docx4j.ask.setKey', title: { value: 'Définir la clé API Anthropic…', original: 'Set the Anthropic API Key…' }, category: 'docx4j editor' },
      { command: 'docx4j.join', title: { value: 'Rejoindre une salle…', original: 'Join a Room…' }, category: { value: 'docx4j éditeur', original: 'docx4j editor' } },
      { command: 'docx4j.commands', title: 'Commands', category: 'docx4j editor' },
    ] } };
    expect(paletteNames(manifest, ['docx4j.ask.setKey', 'docx4j.join', 'docx4j.commands'])).toEqual({
      'docx4j.ask.setKey': 'docx4j editor: Définir la clé API Anthropic…',
      'docx4j.join': 'docx4j éditeur: Rejoindre une salle…',
      'docx4j.commands': 'docx4j editor: Commands',
    });
  });

  test('only the commands asked for; one with no words it can read, or an unresolved %key%, left out for the page\'s English', () => {
    const manifest = { contributes: { commands: [
      { command: 'docx4j.ask.setKey', title: '%command.ask.setKey%', category: '%category%' },
      { command: 'docx4j.join', title: { original: 'Join a Room…' } },
      { command: 'docx4j.pasteText', title: 'Paste as Plain Text' },
      { command: 'other.command', title: 'Other' },
    ] } };
    expect(paletteNames(manifest, ['docx4j.ask.setKey', 'docx4j.join', 'docx4j.pasteText'])).toEqual({ 'docx4j.pasteText': 'Paste as Plain Text' });
    expect(paletteNames(undefined, ['docx4j.ask.setKey'])).toEqual({});
    expect(paletteNames({ contributes: {} }, ['docx4j.ask.setKey'])).toEqual({});
  });
});
