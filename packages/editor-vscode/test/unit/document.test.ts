// SPDX-License-Identifier: GPL-3.0-or-later
//
// The document and its page (ED-014 section 4.3), over a fake file system and a fake page: the requests answered and
// failed when the page goes; `written` after the save resolves; the save's check against another writer; a cancelled
// save and a cancelled backup throwing, with `saveCancelled` posted and the previous backup kept; two saves inside one
// export writing once, the later's bytes; the hand-over kept for its target, taken once, gone after ten seconds; a page
// loaded again opening the latest backup, else the bytes last written.
import { describe, expect, test, vi } from 'vitest';
import { Channel } from '../../src/channel.js';
import { DocxDocument, resumeOf, type DocumentDeps } from '../../src/document.js';
import { Handovers, HANDOVER_MS } from '../../src/handover.js';
import { Requests } from '../../src/requests.js';
import { decodeBackup, type BackupFile } from '../../src/backup.js';
import type { FromPage, ToPage } from '../../src/shared.js';

class Cancelled extends Error { override name = 'Canceled'; }

const bytes = (...b: number[]): Uint8Array => new Uint8Array(b);

/** A file system in memory, keyed by name. */
function fakeIo() {
  const files = new Map<string, Uint8Array>();
  const writes: string[] = [];
  return {
    files,
    writes,
    io: {
      read: async (uri: string) => { const f = files.get(uri); if (!f) throw new Error(`no ${uri}`); return f; },
      readIfAny: async (uri: string) => files.get(uri),
      write: async (uri: string, b: Uint8Array) => { writes.push(uri); files.set(uri, b); },
      delete: async (uri: string) => { files.delete(uri); },
    },
  };
}

/** A page: what it was sent, and a way to answer. */
function fakePage(doc: DocxDocument<string>) {
  const sent: ToPage[] = [];
  const channel = new Channel((m) => sent.push(m));
  doc.channel = channel;
  return {
    sent,
    channel,
    answer: (m: FromPage) => channel.answer(m),
    last: <T extends ToPage['type']>(type: T) => sent.filter((m): m is Extract<ToPage, { type: T }> => m.type === type).at(-1),
  };
}

function makeDoc(over: Partial<DocumentDeps<string>> = {}, init: Partial<ConstructorParameters<typeof DocxDocument<string>>[1]> = {}) {
  const fs = fakeIo();
  fs.files.set('a.docx', bytes(1));
  const later: (() => void)[] = [];
  const changed = vi.fn();
  const deps: DocumentDeps<string> = {
    io: fs.io,
    cancelled: () => new Cancelled(),
    words: { changedOnDisk: (n) => `${n} changed on disk`, noPage: (n) => `${n}: no page` },
    later: (f) => later.push(f),
    changed,
    ...over,
  };
  const doc = new DocxDocument<string>('a.docx', { name: 'a.docx', untitled: false, base: bytes(1), disk: bytes(1), ...init }, deps);
  return { doc, fs, later, changed, runLater: () => { for (const f of later.splice(0)) f(); } };
}

const token = (cancelled = false) => ({ isCancellationRequested: cancelled });

/** The open a page is sent, in its bytes form (not a Join's). */
function opened(doc: DocxDocument<string>): Extract<ToPage, { type: 'open'; bytes: Uint8Array }> {
  const m = doc.openMessage();
  if (!('bytes' in m)) throw new Error(`a Join's open: ${JSON.stringify(m)}`);
  return m;
}
const flush = (): Promise<void> => new Promise((r) => setTimeout(r, 0));

describe('the requests', () => {
  test('answered by id, a stray answer ignored, and every one waiting failed when the page goes', async () => {
    const r = new Requests<string>();
    const posted: number[] = [];
    const a = r.ask((id) => posted.push(id));
    const b = r.ask((id) => posted.push(id));
    expect(posted).toEqual([1, 2]);
    expect(r.answer(99, 'stray')).toBe(false);
    expect(r.answer(1, 'one')).toBe(true);
    await expect(a.answer).resolves.toBe('one');
    r.failAll();
    await expect(b.answer).rejects.toThrow(/gone/);
    expect(r.size).toBe(0);
  });

  test('nothing is posted before ready; ready sends the open first; a page that loads again fails the old one\'s requests', async () => {
    const sent: ToPage[] = [];
    const c = new Channel((m) => sent.push(m));
    c.post({ type: 'disk', bytes: bytes(5), lastModified: 1 });
    const waiting = c.ask((id) => ({ type: 'backup', id })).answer;
    expect(sent).toEqual([]);
    c.ready({ type: 'open', name: 'a.docx', bytes: bytes(1) });
    expect(sent.map((m) => m.type)).toEqual(['open', 'disk', 'backup']);
    c.ready({ type: 'open', name: 'a.docx', bytes: bytes(1) });
    await expect(waiting).rejects.toThrow(/gone/);
    expect(c.loads).toBe(2);
    c.close();
    await expect(c.ask((id) => ({ type: 'backup', id })).answer).rejects.toThrow(/gone/);
  });

  test('an answer to a request of a page that has since loaded again is dropped: the new page\'s ids start again (section 9.23 finding 2)', () => {
    const sent: ToPage[] = [];
    const c = new Channel((m) => sent.push(m));
    c.ready({ type: 'open', name: 'a.docx', bytes: bytes(1) });
    const before = c.replier();
    const now = c.replier();
    now({ type: 'askModels', id: 1, models: [] });
    c.ready({ type: 'open', name: 'a.docx', bytes: bytes(1) });
    before({ type: 'asked', id: 1, cancelled: true });
    const after = c.replier();
    after({ type: 'invited', id: 1 });
    expect(sent.map((m) => m.type)).toEqual(['open', 'askModels', 'open', 'invited']);
    c.close();
    after({ type: 'invited', id: 2 });
    expect(sent).toHaveLength(4);
  });
});

describe('a save', () => {
  test('asks the page, writes, and posts written only in a later task, after VS Code has its promise', async () => {
    const { doc, fs, runLater } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    doc.dirty = true;
    const saving = doc.save(token());
    await flush();
    const ask = page.last('save')!;
    expect(ask).toEqual({ type: 'save', id: 1, name: 'a.docx' });
    page.answer({ type: 'saved', id: 1, bytes: bytes(2) });
    await saving;
    expect([...fs.files.get('a.docx')!]).toEqual([2]);
    expect(doc.dirty).toBe(false);
    expect(page.last('written')).toBeUndefined();
    runLater();
    expect(page.last('written')).toEqual({ type: 'written', id: 1, name: 'a.docx' });
    expect([...doc.disk!]).toEqual([2]);
    // The base stays: the store goes on over the bytes it was opened over.
    expect([...doc.base]).toEqual([1]);
  });

  test('a refusal is thrown, nothing written', async () => {
    const { doc, fs } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    const saving = doc.save(token());
    await flush();
    page.answer({ type: 'saved', id: 1, refused: 'the self-check refused' });
    await expect(saving).rejects.toThrow('the self-check refused');
    expect(fs.writes).toEqual([]);
  });

  test('another writer\'s change fails it, and writeFailed is posted; a file that is gone is written again', async () => {
    const { doc, fs } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    fs.files.set('a.docx', bytes(7));
    const saving = doc.save(token());
    await flush();
    page.answer({ type: 'saved', id: 1, bytes: bytes(2) });
    await expect(saving).rejects.toThrow('a.docx changed on disk');
    expect(page.last('writeFailed')).toEqual({ type: 'writeFailed', id: 1, error: 'a.docx changed on disk' });
    expect([...fs.files.get('a.docx')!]).toEqual([7]);
    fs.files.delete('a.docx');
    const again = doc.save(token());
    await flush();
    page.answer({ type: 'saved', id: 2, bytes: bytes(3) });
    await again;
    expect([...fs.files.get('a.docx')!]).toEqual([3]);
  });

  test('a write that throws posts writeFailed and throws', async () => {
    const { doc } = makeDoc({ io: { ...fakeIo().io, readIfAny: async () => bytes(1), write: async () => { throw new Error('EACCES'); } } });
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    const saving = doc.save(token());
    await flush();
    page.answer({ type: 'saved', id: 1, bytes: bytes(2) });
    await expect(saving).rejects.toThrow('EACCES');
    expect(page.last('writeFailed')).toEqual({ type: 'writeFailed', id: 1, error: 'EACCES' });
  });

  test('without a page, it fails at once', async () => {
    const { doc } = makeDoc();
    await expect(doc.save(token())).rejects.toThrow('a.docx: no page');
  });

  test('cancelled meanwhile: nothing written, saveCancelled posted, VS Code\'s CancellationError thrown', async () => {
    const { doc, fs } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    const t = { isCancellationRequested: false };
    const saving = doc.save(t);
    await flush();
    t.isCancellationRequested = true;
    page.answer({ type: 'saved', id: 1, bytes: bytes(2) });
    await expect(saving).rejects.toBeInstanceOf(Cancelled);
    expect(fs.writes).toEqual([]);
    expect(page.last('saveCancelled')).toEqual({ type: 'saveCancelled', id: 1 });
  });

  test('two saves inside one export: the first cancelled writes nothing, the later writes its own bytes once', async () => {
    const { doc, fs } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    const first = { isCancellationRequested: false };
    const one = doc.save(first);
    await flush();
    // VS Code's new save cancels the one in flight.
    first.isCancellationRequested = true;
    const two = doc.save(token());
    await flush();
    expect(page.sent.filter((m) => m.type === 'save')).toHaveLength(1);
    page.answer({ type: 'saved', id: 1, bytes: bytes(2) });
    await expect(one).rejects.toBeInstanceOf(Cancelled);
    await flush();
    expect(page.sent.filter((m) => m.type === 'save')).toHaveLength(2);
    page.answer({ type: 'saved', id: 2, bytes: bytes(3) });
    await two;
    expect(fs.writes).toEqual(['a.docx']);
    expect([...fs.files.get('a.docx')!]).toEqual([3]);
  });
});

describe('a backup', () => {
  test('the page\'s update and meta, with the base, in the file; kept in memory for a page that loads again', async () => {
    const { doc, fs } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    const writing = doc.writeBackup('backup-1', 'file:///a.docx', token());
    page.answer({ type: 'backup', id: 1, update: bytes(9), meta: { name: 'a.docx' } });
    await writing;
    const file = decodeBackup(fs.files.get('backup-1')!);
    expect(file).toMatchObject({ uri: 'file:///a.docx', name: 'a.docx', meta: { name: 'a.docx' } });
    expect([...file.base]).toEqual([1]);
    expect([...file.update]).toEqual([9]);
    expect(doc.backup?.meta).toEqual({ name: 'a.docx' });
    // A page loaded again opens the backup kept, over the base.
    const reopened = opened(doc);
    expect(reopened).toMatchObject({ type: 'open', name: 'a.docx' });
    expect([...reopened.resume!.update]).toEqual([9]);
  });

  test('cancelled meanwhile: VS Code\'s CancellationError, nothing left at the destination, the backup in memory unchanged', async () => {
    const { doc, fs } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    const t = { isCancellationRequested: false };
    const writing = doc.writeBackup('backup-2', 'file:///a.docx', t);
    t.isCancellationRequested = true;
    page.answer({ type: 'backup', id: 1, update: bytes(9), meta: { name: 'a.docx' } });
    await expect(writing).rejects.toBeInstanceOf(Cancelled);
    expect(fs.files.has('backup-2')).toBe(false);
    expect(doc.backup).toBeUndefined();
  });

  test('an empty answer is a backup of the base alone, which resumes as a plain open', async () => {
    const { doc, fs } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    const writing = doc.writeBackup('backup-3', 'file:///a.docx', token());
    page.answer({ type: 'backup', id: 1, empty: true });
    await writing;
    expect(decodeBackup(fs.files.get('backup-3')!).update.byteLength).toBe(0);
    expect(opened(doc).resume).toBeUndefined();
  });

  test('a save and a revert drop the backup in memory: a page loaded again opens the bytes last written (section 9.8 finding 4)', async () => {
    const { doc, fs, runLater } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    const writing = doc.writeBackup('backup-4', 'file:///a.docx', token());
    page.answer({ type: 'backup', id: 1, update: bytes(9), meta: { name: 'a.docx' } });
    await writing;
    const saving = doc.save(token());
    await flush();
    page.answer({ type: 'saved', id: 2, bytes: bytes(4) });
    await saving;
    runLater();
    const reopened = opened(doc);
    expect(reopened.resume).toBeUndefined();
    expect([...reopened.bytes]).toEqual([4]);
    expect([...doc.base]).toEqual([4]);
    fs.files.set('a.docx', bytes(6));
    await doc.revert();
    expect(page.last('revert')).toMatchObject({ type: 'revert' });
    expect([...doc.base]).toEqual([6]);
    expect(doc.backup).toBeUndefined();
  });
});

describe('Save As and the hand-over (P25)', () => {
  test('onto another file: written there, the page asked for its hand-over under the save\'s id, its place and late store returned', async () => {
    const { doc, fs } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    const saving = doc.saveAs('b.docx', 'b.docx', false, token());
    await flush();
    expect(page.last('save')).toEqual({ type: 'save', id: 1, name: 'b.docx' });
    page.answer({ type: 'saved', id: 1, bytes: bytes(5) });
    await flush();
    expect([...fs.files.get('b.docx')!]).toEqual([5]);
    expect(page.last('handoff')).toEqual({ type: 'handoff', id: 1 });
    page.answer({ type: 'handoff', id: 1, place: { caret: 3 }, resume: { update: bytes(8), meta: { name: 'a.docx' } } });
    const handover = await saving;
    expect(handover).toMatchObject({ place: { caret: 3 }, resume: { meta: { name: 'a.docx' } } });
    expect([...handover!.bytes]).toEqual([5]);
    expect([...handover!.base]).toEqual([1]);
    // No written: the page's session ends with its editor.
    expect(page.last('written')).toBeUndefined();
  });

  test('onto its own file: a save, with written', async () => {
    const { doc, runLater } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    const saving = doc.saveAs('a.docx', 'a.docx', true, token());
    await flush();
    page.answer({ type: 'saved', id: 1, bytes: bytes(5) });
    expect(await saving).toBeUndefined();
    runLater();
    expect(page.last('written')).toEqual({ type: 'written', id: 1, name: 'a.docx' });
  });

  test('a hand-over is kept for its target, taken once, and gone after ten seconds', () => {
    vi.useFakeTimers();
    try {
      const h = new Handovers<string>();
      h.put('b.docx', 'one');
      expect(h.take('b.docx')).toBe('one');
      expect(h.take('b.docx')).toBeUndefined();
      h.put('c.docx', 'two');
      vi.advanceTimersByTime(HANDOVER_MS - 1);
      expect(h.has('c.docx')).toBe(true);
      vi.advanceTimersByTime(1);
      expect(h.take('c.docx')).toBeUndefined();
    } finally {
      vi.useRealTimers();
    }
  });
});

describe('opening', () => {
  test('a Markdown source opens by its name, converted again by the page, until it posts its base', () => {
    const { doc } = makeDoc({}, { name: 'notes.docx', untitled: true, base: new Uint8Array(), source: { uri: 'notes.md', name: 'notes.md', bytes: bytes(35) } });
    doc.disk = undefined;
    expect(doc.openMessage()).toMatchObject({ type: 'open', name: 'notes.md' });
  });

  test('a hand-over\'s place goes with the first open only', () => {
    const { doc } = makeDoc({}, { place: { caret: 1 } });
    expect(opened(doc).place).toEqual({ caret: 1 });
    expect(opened(doc).place).toBeUndefined();
  });
});

describe('a resume from a backup (Fable\'s review, ED-014 section 9.11 finding 1)', () => {
  const file = (over: Partial<BackupFile> = {}): BackupFile => ({ uri: 'a.docx', name: 'a.docx', meta: { name: 'a.docx' }, base: bytes(1), update: bytes(9), ...over });

  test('after a save in the session, disk is the bytes written, and a file that holds them is no change', () => {
    const { init, report } = resumeOf<string>(file({ disk: bytes(2) }), 'a.docx', false, { file: bytes(2) });
    expect([...init.disk!]).toEqual([2]);
    expect([...init.base]).toEqual([1]);
    expect(report).toBeUndefined();
  });

  test('a file another program changed while VS Code was closed is reported, and disk stays what the session knew', () => {
    const { init, report } = resumeOf<string>(file({ disk: bytes(2) }), 'a.docx', false, { file: bytes(7) });
    expect([...init.disk!]).toEqual([2]);
    expect([...report!]).toEqual([7]);
  });

  test('a backup with no disk recorded compares with base, as before', () => {
    expect(resumeOf<string>(file(), 'a.docx', false, { file: bytes(1) }).report).toBeUndefined();
    expect([...resumeOf<string>(file(), 'a.docx', false, { file: bytes(5) }).report!]).toEqual([5]);
  });

  test('a Markdown source changed while closed is reported; one unchanged is not; the name is the URI\'s', () => {
    const md = (b: number[]) => ({ uri: 'notes.md', name: 'notes.md', bytes: bytes(...b) });
    const changed = resumeOf<string>(file({ sourceBytes: bytes(35), source: 'notes.md' }), 'notes.docx', true, { source: md([36]) });
    expect([...changed.report!]).toEqual([36]);
    expect([...changed.init.source!.bytes]).toEqual([35]);
    expect(changed.init.disk).toBeUndefined();
    expect(changed.init.backup!.meta.name).toBe('notes.docx');
    expect(resumeOf<string>(file({ sourceBytes: bytes(35) }), 'notes.docx', true, { source: md([35]) }).report).toBeUndefined();
  });

  test('a backup written after a save records the bytes written, and one of base alone records none', async () => {
    const { doc, fs, runLater } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    const first = doc.writeBackup('b1', 'file:///a.docx', token());
    await flush();
    page.answer({ type: 'backup', id: 1, update: bytes(9), meta: { name: 'a.docx' } });
    await first;
    expect(decodeBackup(fs.files.get('b1')!).disk).toBeUndefined();
    const saving = doc.save(token());
    await flush();
    page.answer({ type: 'saved', id: 2, bytes: bytes(4) });
    await saving;
    runLater();
    const second = doc.writeBackup('b2', 'file:///a.docx', token());
    await flush();
    page.answer({ type: 'backup', id: 3, update: bytes(9), meta: { name: 'a.docx' } });
    await second;
    expect([...decodeBackup(fs.files.get('b2')!).disk!]).toEqual([4]);
  });
});

describe('a revert during a save (Fable\'s review, ED-014 section 9.11 finding 2)', () => {
  test('the save whose export began before the revert writes nothing, and the page hears it cancelled', async () => {
    const { doc, fs } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    const saving = doc.save(token());
    await flush();
    await doc.revert();
    page.answer({ type: 'saved', id: 1, bytes: bytes(2) });
    await expect(saving).rejects.toBeInstanceOf(Cancelled);
    expect(fs.writes).toEqual([]);
    expect(page.last('saveCancelled')).toEqual({ type: 'saveCancelled', id: 1 });
  });

  test('a rename\'s hand-over during a save: the save writes nothing to the old name', async () => {
    const { doc, fs } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    const saving = doc.save(token());
    await flush();
    const handing = doc.handOver();
    await flush();
    page.answer({ type: 'handoff', id: 2, place: { caret: 1 } });
    await handing;
    page.answer({ type: 'saved', id: 1, bytes: bytes(2) });
    await expect(saving).rejects.toBeInstanceOf(Cancelled);
    expect(fs.writes).toEqual([]);
  });
});

describe('a room across a Save As and a rename (P42), and a Join (section 4.5)', () => {
  const room = { room: 'AAAAAAAAAAAAAAAAAAAAAA', token: 'tok', relay: 'ws://127.0.0.1:22642' };
  const inRoom = { update: bytes(8), meta: { name: 'a.docx', room } };

  async function saveAsAnswered(late: boolean) {
    const { doc } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    const saving = doc.saveAs('b.docx', 'b.docx', false, token());
    await flush();
    page.answer({ type: 'saved', id: 1, bytes: bytes(5) });
    await flush();
    page.answer({ type: 'handoff', id: 1, resume: inRoom, late });
    return (await saving)!;
  }

  test('a Save As in a room with no edit after the export: the store goes with it, and the target opens saved', async () => {
    const handover = await saveAsAnswered(false);
    expect(handover.resume?.meta.room).toEqual(room);
    expect(handover.clean).toBe(true);
    // The target's editor, opened from it as the provider does: a resume that opens as saved, in its room.
    const target = new DocxDocument<string>('b.docx', { name: 'b.docx', untitled: false, base: handover.base, disk: handover.bytes, backup: handover.resume!, dirty: !handover.clean }, makeDoc().doc['deps']);
    expect(target.dirty).toBe(false);
    expect(target.openMessage()).toMatchObject({ type: 'open', name: 'b.docx', resume: inRoom, clean: true });
  });

  test('a Save As with an edit after the export: the target opens dirty', async () => {
    const handover = await saveAsAnswered(true);
    expect(handover.clean).toBe(false);
    const target = new DocxDocument<string>('b.docx', { name: 'b.docx', untitled: false, base: handover.base, disk: handover.bytes, backup: handover.resume!, dirty: !handover.clean }, makeDoc().doc['deps']);
    expect(target.openMessage()).not.toHaveProperty('clean');
  });

  test('a page that says no late, from before P42, is taken as late when it sends a store', async () => {
    const { doc } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    const saving = doc.saveAs('b.docx', 'b.docx', false, token());
    await flush();
    page.answer({ type: 'saved', id: 1, bytes: bytes(5) });
    await flush();
    page.answer({ type: 'handoff', id: 1, resume: inRoom });
    expect((await saving)!.clean).toBe(false);
  });

  test('a rename of a clean document in a room hands over its store and opens clean, though the page says late', async () => {
    const { doc } = makeDoc();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    doc.room = room.relay;
    const handing = doc.handOver();
    await flush();
    page.answer({ type: 'handoff', id: 1, resume: inRoom, late: true });
    const handover = await handing;
    expect(handover.clean).toBe(true);
    expect(handover.resume).toEqual(inRoom);
    // A dirty one opens dirty.
    const dirty = makeDoc();
    const page2 = fakePage(dirty.doc);
    page2.channel.ready(dirty.doc.openMessage());
    dirty.doc.changed();
    const handing2 = dirty.doc.handOver();
    await flush();
    page2.answer({ type: 'handoff', id: 1, resume: inRoom, late: true });
    expect((await handing2).clean).toBe(false);
  });

  test('a Join\'s untitled document is dirty from its open, and opens as a Join until the page posts the room\'s base file', () => {
    const doc = new DocxDocument<string>('untitled:Room.docx', { name: 'Room.docx', untitled: true, base: new Uint8Array(), join: room, dirty: true }, makeDoc().doc['deps']);
    expect(doc.dirty).toBe(true);
    expect(doc.openMessage()).toEqual({ type: 'open', name: 'Room.docx', join: room });
    // A page loaded again before the base came joins again, and after it too while there is no backup: the room holds it.
    expect(doc.openMessage()).toEqual({ type: 'open', name: 'Room.docx', join: room });
    doc.base = bytes(4, 5);
    expect(doc.openMessage()).toEqual({ type: 'open', name: 'Room.docx', join: room });
  });

  const joined = () => {
    const fs = makeDoc().fs;
    const deps = { ...makeDoc().doc['deps'], io: fs.io };
    const doc = new DocxDocument<string>('untitled:Room.docx', { name: 'Room.docx', untitled: true, base: new Uint8Array(), join: room, dirty: true }, deps);
    return { doc, fs };
  };
  /** The document VS Code opens from this backup file after a restart. */
  const restored = (file: Uint8Array) => {
    const { init } = resumeOf<string>(decodeBackup(file), 'Room.docx', true, {});
    return new DocxDocument<string>('untitled:Room.docx', init, makeDoc().doc['deps']);
  };

  test('a Join backed up before its page is ready records its room, and a restart joins it again, dirty (Fable\'s review, section 9.19 finding 1)', async () => {
    const { doc, fs } = joined();
    await doc.writeBackup('backup-j', 'untitled:Room.docx', token());
    const file = decodeBackup(fs.files.get('backup-j')!);
    expect(file.meta).toEqual({ name: 'Room.docx', room, joined: true });
    expect(file.update.byteLength).toBe(0);
    const again = restored(fs.files.get('backup-j')!);
    expect(again.dirty).toBe(true);
    expect(again.openMessage()).toEqual({ type: 'open', name: 'Room.docx', join: room });
  });

  test('a Join whose page answers a backup with nothing (joining, or not yet edited) records its room the same way; once it answers with the store, that is resumed', async () => {
    const { doc, fs } = joined();
    const page = fakePage(doc);
    page.channel.ready(doc.openMessage());
    doc.base = bytes(4, 5);
    const writing = doc.writeBackup('backup-k', 'untitled:Room.docx', token());
    page.answer({ type: 'backup', id: 1, empty: true });
    await writing;
    expect(decodeBackup(fs.files.get('backup-k')!).meta).toEqual({ name: 'Room.docx', room, joined: true });
    // Its base posted before the backup: still joined again, never the base file opened out of the room.
    expect(restored(fs.files.get('backup-k')!).openMessage()).toEqual({ type: 'open', name: 'Room.docx', join: room });
    // Edited: the page's store, in its room, resumed over the base.
    const edited = doc.writeBackup('backup-l', 'untitled:Room.docx', token());
    page.answer({ type: 'backup', id: 2, update: bytes(9), meta: { name: 'Room.docx', room, joined: true } });
    await edited;
    const resumed = restored(fs.files.get('backup-l')!).openMessage();
    expect(resumed).toMatchObject({ type: 'open', name: 'Room.docx', resume: { meta: { room, joined: true } } });
    expect([...(resumed as { bytes: Uint8Array }).bytes]).toEqual([4, 5]);
  });
});
