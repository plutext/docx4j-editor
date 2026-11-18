// SPDX-License-Identifier: GPL-3.0-or-later
//
// What first marked a document changed (ED-011 section 9.9): the description's parts over a real Yjs document, the
// input follower, and the log's bounds. The wiring into the dirty state is the browser suite's (e2e/disk.spec.ts).
import { afterEach, describe, expect, it, vi } from 'vitest';
import * as Y from 'yjs';
import { describeChanges, describeElement, describeKey, describeOrigin, deepStack, editCauseLog, followInput, logEditCause, stackFrames, type EditCauseEntry } from '../src/app/editcause.js';

/** A store shaped as the editor's: a body of paragraphs (one inside a table) and a package map. */
function store(): { doc: Y.Doc; body: Y.XmlFragment; texts: Y.XmlText[] } {
  const doc = new Y.Doc();
  const body = doc.getXmlFragment('body');
  doc.getMap('package');
  const texts: Y.XmlText[] = [];
  doc.transact(() => {
    for (let i = 0; i < 3; i++) {
      const p = new Y.XmlElement('paragraph');
      const t = new Y.XmlText(`block ${i}`);
      p.insert(0, [t]);
      body.insert(i, [p]);
      texts.push(t);
    }
    const table = new Y.XmlElement('table');
    const row = new Y.XmlElement('table_row');
    const cell = new Y.XmlElement('table_cell');
    const p = new Y.XmlElement('paragraph');
    const t = new Y.XmlText('in a cell');
    p.insert(0, [t]);
    cell.insert(0, [p]);
    row.insert(0, [cell]);
    table.insert(0, [row]);
    body.insert(3, [table]);
    texts.push(t);
  }, 'import');
  return { doc, body, texts };
}

/** The described changes of the one transaction `fn` makes. */
function described(doc: Y.Doc, body: Y.XmlFragment, origin: unknown, fn: () => void): ReturnType<typeof describeChanges> {
  let out: ReturnType<typeof describeChanges> | undefined;
  const listen = (tr: Y.Transaction): void => { out = describeChanges(tr, body); };
  doc.on('afterTransaction', listen);
  doc.transact(fn, origin);
  doc.off('afterTransaction', listen);
  return out!;
}

afterEach(() => { localStorage.clear(); vi.restoreAllMocks(); document.body.innerHTML = ''; });

describe('a first edit, described', () => {
  it('names the origin: a string, a plugin key by its key, the undo manager, a class, nothing', () => {
    const { body } = store();
    expect(describeOrigin('local')).toBe('local');
    expect(describeOrigin({ key: 'y-sync$' })).toBe('y-sync$');
    expect(describeOrigin(new Y.UndoManager(body))).toBe('undo');
    expect(describeOrigin(new (class Provider {})())).toBe('Provider');
    expect(describeOrigin(null)).toBe('none');
  });

  it('names the store entries changed and the body\'s blocks, a cell\'s text by its table', () => {
    const { doc, body, texts } = store();
    expect(described(doc, body, { key: 'y-sync$' }, () => texts[1]!.insert(0, 'x'))).toEqual({ origin: 'y-sync$', local: true, changed: ['body'], blocks: [1] });
    expect(described(doc, body, 'local', () => { texts[2]!.insert(0, 'x'); texts[0]!.insert(0, 'y'); }).blocks).toEqual([0, 2]);
    expect(described(doc, body, 'local', () => texts[3]!.insert(0, 'x')).blocks).toEqual([3]);
    expect(described(doc, body, 'parts', () => doc.getMap('package').set('finalSectPr', 'x'))).toEqual({ origin: 'parts', local: true, changed: ['package'] });
  });

  it('names no block where a block was put in or taken out, the body itself being what changed', () => {
    const { doc, body } = store();
    const added = described(doc, body, 'local', () => body.insert(0, [new Y.XmlElement('paragraph')]));
    expect(added.changed).toEqual(['body']);
    expect(added.blocks).toBeUndefined();
  });

  it('names a key by its name and modifiers, and every character as one word, so nothing typed is kept', () => {
    const key = (k: string, mods: Partial<Record<'ctrlKey' | 'metaKey' | 'altKey' | 'shiftKey', boolean>> = {}) => describeKey({ key: k, ctrlKey: false, metaKey: false, altKey: false, shiftKey: false, ...mods });
    expect(key('a')).toBe('character');
    expect(key('A', { shiftKey: true })).toBe('character');
    expect(key('Enter')).toBe('Enter');
    expect(key('Tab', { shiftKey: true })).toBe('Shift+Tab');
    expect(key('z', { ctrlKey: true })).toBe('Ctrl+Z');
  });

  it('names an element as the editor it is in, else its test id, else its tag and class', () => {
    document.body.innerHTML = '<div class="ProseMirror" aria-label="Document"><p id="in">x</p></div><div data-testid="find-bar"><input id="find"></div><span class="lone extra" id="lone"></span>';
    expect(describeElement(document.getElementById('in'))).toBe('editor:Document');
    expect(describeElement(document.getElementById('find'))).toBe('find-bar');
    expect(describeElement(document.getElementById('lone'))).toBe('span.lone');
    expect(describeElement(null)).toBeUndefined();
  });

  it('follows the last input until stopped', () => {
    document.body.innerHTML = '<div class="ProseMirror" aria-label="Document" id="ed"></div><button data-testid="download" id="dl"></button>';
    const follower = followInput(window);
    expect(follower.seen()).toBeUndefined();
    document.getElementById('ed')!.dispatchEvent(new KeyboardEvent('keydown', { key: 'q', bubbles: true }));
    expect(follower.seen()).toMatchObject({ kind: 'keydown', key: 'character', target: 'editor:Document' });
    document.getElementById('dl')!.dispatchEvent(new Event('pointerdown', { bubbles: true }));
    const seen = follower.seen(performance.now() + 50);
    expect(seen).toMatchObject({ kind: 'pointerdown', target: 'download' });
    expect(seen!.msBefore).toBeGreaterThanOrEqual(50);
    expect(seen!.key).toBeUndefined();
    follower.stop();
    document.getElementById('ed')!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(follower.seen()!.kind).toBe('pointerdown');
  });

  it('takes the stack deeper than V8\'s ten frames, and gives the limit back', () => {
    const limit = (Error as { stackTraceLimit?: number }).stackTraceLimit;
    const down = (n: number): string | undefined => (n === 0 ? deepStack() : down(n - 1));
    const frames = stackFrames(down(25), '')!;
    expect(frames.length).toBeGreaterThan(20);
    expect(frames.length).toBeLessThanOrEqual(30);
    expect((Error as { stackTraceLimit?: number }).stackTraceLimit).toBe(limit);
  });

  it('takes the page\'s address out of each frame and the error line away', () => {
    expect(stackFrames('Error\n    at f (http://localhost:5173/src/a.ts:1:2)\n    at g (http://localhost:5173/src/b.ts:3:4)', 'http://localhost:5173'))
      .toEqual(['at f (/src/a.ts:1:2)', 'at g (/src/b.ts:3:4)']);
    expect(stackFrames(undefined)).toBeUndefined();
  });
});

describe('the log of first edits', () => {
  const entry = (i: number): EditCauseEntry => ({ session: `s${i}`, name: 'a.docx', at: `t${i}`, origin: 'y-sync$', local: true, changed: ['body'], visible: true });

  it('keeps the last twenty, oldest first', () => {
    for (let i = 0; i < 25; i++) logEditCause(entry(i));
    const log = editCauseLog();
    expect(log).toHaveLength(20);
    expect(log[0]!.session).toBe('s5');
    expect(log[19]!.session).toBe('s24');
  });

  it('reads a damaged log as empty and writes over it', () => {
    localStorage.setItem('docx4j-editor.editCauses', '{not json');
    expect(editCauseLog()).toEqual([]);
    logEditCause(entry(1));
    expect(editCauseLog()).toHaveLength(1);
  });

  it('keeps nothing, and fails nothing, where storage is refused', () => {
    vi.spyOn(Storage.prototype, 'setItem').mockImplementation(() => { throw new DOMException('full', 'QuotaExceededError'); });
    expect(() => logEditCause(entry(1))).not.toThrow();
    vi.spyOn(Storage.prototype, 'getItem').mockImplementation(() => { throw new DOMException('denied', 'SecurityError'); });
    expect(editCauseLog()).toEqual([]);
  });
});
