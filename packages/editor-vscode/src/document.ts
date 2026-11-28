// SPDX-License-Identifier: GPL-3.0-or-later
//
// One document (ED-014 section 4.3), with no `vscode` import: its bytes, its dirty state, its save queue, and the rules
// for a save, a Save As's hand-over, a backup and a revert. `provider.ts` gives it VS Code's file system, its
// `CancellationError` and its words; the unit tests give it fakes.
//
// The bytes it holds:
// - `base`: the bytes the page's store was opened over, which a backup is resumed over. A save leaves it, since the store
//   goes on; a page opened afresh (a revert, a page loaded again without a backup) opens over new bytes, which become it.
// - `disk`: the document's own file as last read or written (or being written): the save's check and the watcher compare
//   with it, as ED-011's `changedSince` does. None for an untitled document.
// - `sourceDisk`: for a document made from Markdown (P26), the `.md` as last read, watched until the first save.
import { Channel, type Answer } from './channel.js';
import type { BackupData, BackupMetaData, RoomLinkData, ToPage } from './shared.js';
import { encodeBackup, type BackupFile } from './backup.js';

export interface Io<U> {
  /** The file's bytes; throws if there is none. */
  read(uri: U): Promise<Uint8Array>;
  /** The file's bytes, or undefined if there is none. */
  readIfAny(uri: U): Promise<Uint8Array | undefined>;
  write(uri: U, bytes: Uint8Array): Promise<void>;
  delete(uri: U): Promise<void>;
}

export interface Cancel {
  readonly isCancellationRequested: boolean;
}

export interface DocumentDeps<U> {
  io: Io<U>;
  /** VS Code's `CancellationError`: thrown, it keeps the backup (section 9.8 findings 5 and 7). */
  cancelled(): Error;
  words: {
    /** Another writer changed the file since it was read or written. */
    changedOnDisk(name: string): string;
    /** The page has gone, or never said `ready`. */
    noPage(name: string): string;
  };
  /** In a later task: `written` must reach the page after VS Code has cleared its dirty state. */
  later(f: () => void): void;
  /** A local edit: VS Code's content-change event (P6). */
  changed(): void;
}

/** What a hand-over carries to the target's editor (P25). */
export interface Handover {
  /** The bytes just written. */
  bytes: Uint8Array;
  place?: unknown;
  /** An edit came after the export, or the session is in a room (P42): the store, over `base`. */
  resume?: BackupData;
  base: Uint8Array;
  /**
   * The target opens saved (P42), which the extension decides: after a Save As, unless an edit came after the export; after a
   * rename, if the document was clean. A rename asks under an id no export was made under, so the page's `late` cannot.
   */
  clean: boolean;
}

export interface DocumentInit<U> {
  name: string;
  untitled: boolean;
  base: Uint8Array;
  disk?: Uint8Array;
  source?: { uri: U; name: string; bytes: Uint8Array };
  backup?: BackupData;
  place?: unknown;
  /** Dirty from its open: an untitled document with content, or a hand-over with late edits. */
  dirty?: boolean;
  /** A room to join (P17): the page opens it until it posts the room's base file. */
  join?: RoomLinkData;
}

/** Equal bytes. */
export function sameBytes(a: Uint8Array | undefined, b: Uint8Array | undefined): boolean {
  if (a === b) return true;
  if (!a || !b || a.byteLength !== b.byteLength) return false;
  for (let i = 0; i < a.byteLength; i++) if (a[i] !== b[i]) return false;
  return true;
}

/**
 * A document opened from its backup (ED-014 section 4.3, corrected after Fable's review, section 9.11 finding 1).
 * - **`disk`** is the file as the session last read or wrote it, recorded in the backup. A save made in the session is
 *   then not another writer's, as it would be were `disk` the backup's `base`.
 * - **`report`** is what changed while VS Code was closed, for the banner: the file, or the Markdown source, as it is now.
 *   A backup from before these were recorded compares the file with `base`, and the source not at all.
 */
export function resumeOf<U>(
  file: BackupFile,
  name: string,
  untitled: boolean,
  now: { file?: Uint8Array; source?: { uri: U; name: string; bytes: Uint8Array } },
): { init: DocumentInit<U>; report?: Uint8Array } {
  // After a rename, VS Code hands the old backup to the new URI: the document is named by its own URI.
  const meta = { ...file.meta, name };
  const disk = untitled ? undefined : file.disk ?? file.base;
  const recordedSource = file.sourceBytes;
  const source = now.source ? { ...now.source, bytes: recordedSource ?? now.source.bytes } : undefined;
  let report: Uint8Array | undefined;
  if (disk && now.file && !sameBytes(now.file, disk)) report = now.file;
  if (now.source && recordedSource && !sameBytes(now.source.bytes, recordedSource)) report = now.source.bytes;
  // A joined room with nothing of its own beyond the room (backed up before an edit): joined again (section 9.19 finding 1).
  const join = file.update.byteLength === 0 && file.meta.joined === true && file.meta.room ? file.meta.room : undefined;
  return {
    init: {
      name,
      untitled,
      base: file.base,
      ...(join ? { join } : {}),
      ...(disk ? { disk } : {}),
      ...(source ? { source } : {}),
      ...(file.update.byteLength > 0 ? { backup: { update: file.update, meta } } : {}),
      // VS Code shows a document opened from a backup as dirty.
      dirty: true,
    },
    ...(report ? { report } : {}),
  };
}

/** One at a time, in order. */
class Queue {
  private tail: Promise<unknown> = Promise.resolve();
  run<T>(f: () => Promise<T>): Promise<T> {
    const next = this.tail.then(f, f);
    this.tail = next.catch(() => undefined);
    return next;
  }
}

export class DocxDocument<U> {
  name: string;
  readonly untitled: boolean;
  base: Uint8Array;
  disk: Uint8Array | undefined;
  /** The bytes a save is writing: the watcher's own write is never a change. */
  writing: Uint8Array | undefined;
  source: { uri: U; name: string; bytes: Uint8Array } | undefined;
  dirty: boolean;
  /** The latest backup written, kept in memory: a page that loads again opens from it. A save and a revert drop it. */
  backup: BackupData | undefined;
  /** The place the next open restores, once (a hand-over's). */
  private place: unknown;
  /** The room a Join opens, until the page posts the room's base file (`base`). */
  join: RoomLinkData | undefined;
  /** The relay of the room the page's session is in, as the page last said (`room`); none out of a room. */
  room: string | undefined;
  /** The page, once a panel shows the document. */
  channel: Channel | undefined;
  private readonly saves = new Queue();
  /**
   * Bumped by whatever supersedes a save in flight: a revert, a Markdown source read again, a hand-over at a rename. A save
   * whose export began before it writes nothing (Fable's review, ED-014 section 9.11 finding 2): on the 255-page document
   * the export takes about 33 s, and Revert File in that time would otherwise see the discarded edits written after it.
   */
  private generation = 0;

  constructor(readonly uri: U, init: DocumentInit<U>, private readonly deps: DocumentDeps<U>) {
    this.name = init.name;
    this.untitled = init.untitled;
    this.base = init.base;
    this.disk = init.disk;
    this.source = init.source;
    this.backup = init.backup;
    this.place = init.place;
    this.dirty = init.dirty ?? false;
    this.join = init.join;
  }

  /**
   * What a page opens when it says `ready`. The first page opens what the document was opened with; a page that loads again
   * (moved to another window, Developer: Reload Webviews) opens the latest backup, kept in memory, else the file as last
   * read or written, else the Markdown source, else `base`. The edits since that backup are lost: at most, the typing since
   * the last pause. Whatever it opens over becomes `base`, since the store is new (section 9.8 finding 4).
   */
  openMessage(): Extract<ToPage, { type: 'open' }> {
    // A Join with no backup of its own (none yet, or one taken before an edit): it joins again, the room holding the document
    // (section 4.5; section 9.19 finding 1). Its base file, once posted, is what a later backup is resumed over.
    if (this.join && !this.backup) return { type: 'open', name: this.name, join: this.join };
    const place = this.place;
    this.place = undefined;
    const with_ = place !== undefined ? { place } : {};
    // A resume of a clean document (a hand-over in a room, P42) opens as saved.
    if (this.backup) return { type: 'open', name: this.name, bytes: this.base, resume: this.backup, ...with_, ...(this.dirty ? {} : { clean: true }) };
    if (this.disk) {
      this.base = this.disk;
      return { type: 'open', name: this.name, bytes: this.disk, ...with_ };
    }
    // Converted again by the page, which posts `base` with the package it makes.
    if (this.source) return { type: 'open', name: this.source.name, bytes: this.source.bytes, ...with_ };
    return { type: 'open', name: this.name, bytes: this.base, ...with_ };
  }

  /** A local edit on the page. */
  changed(): void {
    this.dirty = true;
    this.deps.changed();
  }

  private page(): Channel {
    if (!this.channel?.isReady) throw new Error(this.deps.words.noPage(this.name));
    return this.channel;
  }

  /** The page's self-checked bytes, in the form `name` asks for; a refusal thrown, so VS Code keeps the document dirty. */
  private async export(token: Cancel, name: string): Promise<{ id: number; bytes: Uint8Array }> {
    const { id, answer } = this.page().ask((id) => ({ type: 'save', id, name }));
    const out = await answer as Extract<Answer, { type: 'saved' }>;
    if ('refused' in out) throw new Error(out.refused);
    if (token.isCancellationRequested) this.cancel(id);
    return { id, bytes: out.bytes };
  }

  /** A cancelled save writes nothing, tells the page, and throws VS Code's `CancellationError`, so its backup is kept. */
  private cancel(id: number): never {
    this.channel?.post({ type: 'saveCancelled', id });
    throw this.deps.cancelled();
  }

  /**
   * VS Code's Save (section 4.3). In the document's queue, since a new save cancels the one in flight and a cancelled save
   * must never write after a later one. The file is read again first: if another writer changed it since it was read or
   * written, the save fails, as Save in place refuses (`files.ts`); a file that is gone is written again.
   */
  save(token: Cancel): Promise<void> {
    return this.saves.run(async () => {
      if (token.isCancellationRequested) throw this.deps.cancelled();
      const generation = this.generation;
      const { id, bytes } = await this.export(token, this.name);
      if (generation !== this.generation) this.cancel(id);
      const now = await this.deps.io.readIfAny(this.uri);
      if (now !== undefined && this.disk !== undefined && !sameBytes(now, this.disk)) {
        const error = this.deps.words.changedOnDisk(this.name);
        this.channel?.post({ type: 'writeFailed', id, error });
        throw new Error(error);
      }
      if (token.isCancellationRequested || generation !== this.generation) this.cancel(id);
      // A token that fires during the write is not looked at again: the bytes are written, and VS Code, which has moved
      // on to the later save queued behind this one, ignores this one's result. The page's `written` is then right.
      await this.write(this.uri, id, bytes);
      this.disk = bytes;
      this.saved();
      const name = this.name;
      this.deps.later(() => this.channel?.post({ type: 'written', id, name }));
    });
  }

  /**
   * VS Code's Save As (P25), with no check: VS Code's dialog has already asked about replacing. Onto the document's own
   * file it is a save. Onto another, VS Code then opens the target in a new editor and closes this one, so the page is
   * asked for a hand-over: its place, and its store if an edit came after the export or the session is in a room (P42),
   * kept for the target's editor, which opens saved unless an edit came after the export.
   */
  saveAs(target: U, targetName: string, same: boolean, token: Cancel): Promise<Handover | undefined> {
    return this.saves.run(async () => {
      if (token.isCancellationRequested) throw this.deps.cancelled();
      const generation = this.generation;
      const { id, bytes } = await this.export(token, targetName);
      if (token.isCancellationRequested || generation !== this.generation) this.cancel(id);
      await this.write(target, id, bytes);
      if (same) {
        this.disk = bytes;
        this.name = targetName;
        this.saved();
        this.deps.later(() => this.channel?.post({ type: 'written', id, name: targetName }));
        return undefined;
      }
      let place: unknown;
      let resume: BackupData | undefined;
      let late = false;
      try {
        // Under the save's id, which the page's count of edits since the export is kept by.
        const out = await this.page().ask((h) => ({ type: 'handoff', id: h }), id).answer as Extract<Answer, { type: 'handoff' }>;
        place = out.place;
        resume = out.resume;
        late = out.late ?? resume !== undefined;
      } catch {
        // The page has gone: the file is written, and the target's editor opens it without a place.
      }
      return { bytes, base: this.base, ...(place !== undefined ? { place } : {}), ...(resume ? { resume } : {}), clean: !late };
    });
  }

  /**
   * A rename or a move of the file while dirty, or while in a room (ED-014 sections 9.9 and 4.5): VS Code reopens the document
   * at its new URI from the file, clean, with no backup. The page is asked for a hand-over as a Save As asks for one, under an
   * id no export was made under, so the page gives its store with its place; the new editor resumes it over this document's
   * base, saved if this document was (P42).
   */
  async handOver(): Promise<Handover> {
    this.generation++;
    const clean = !this.dirty;
    const out = await this.page().ask((id) => ({ type: 'handoff', id })).answer as Extract<Answer, { type: 'handoff' }>;
    return {
      bytes: this.disk ?? this.base,
      base: this.base,
      ...(out.place !== undefined ? { place: out.place } : {}),
      ...(out.resume ? { resume: out.resume } : {}),
      clean,
    };
  }

  private async write(uri: U, id: number, bytes: Uint8Array): Promise<void> {
    this.writing = bytes;
    try {
      await this.deps.io.write(uri, bytes);
    } catch (e) {
      const error = (e as Error).message ?? String(e);
      this.channel?.post({ type: 'writeFailed', id, error });
      throw e;
    } finally {
      this.writing = undefined;
    }
  }

  /** VS Code clears its dirty state and the extension host deletes the backup once a save resolves (section 9.8 finding 4). */
  private saved(): void {
    this.dirty = false;
    this.backup = undefined;
  }

  /**
   * VS Code's hot exit (P8, P20). If the token fired meanwhile, VS Code's `CancellationError` is thrown, so the previous
   * backup stays (section 9.8 finding 5). Before a page is ready, the backup kept in memory (a resume's) is written again.
   */
  async writeBackup(destination: U, uri: string, token: Cancel): Promise<void> {
    let update: Uint8Array = new Uint8Array();
    let meta: BackupMetaData = { name: this.name, ...(this.source ? { source: this.source.name } : {}) };
    let kept: BackupData | undefined;
    // What the file and the source were as the session knew them, for the resume to compare with (section 9.11 finding 1).
    const disk = this.disk && !sameBytes(this.disk, this.base) ? this.disk : undefined;
    const sourceBytes = this.source?.bytes;
    if (this.channel?.isReady) {
      const out = await this.channel.ask((id) => ({ type: 'backup', id })).answer as Extract<Answer, { type: 'backup' }>;
      if (!('empty' in out)) kept = { update: out.update, meta: out.meta };
    } else {
      kept = this.backup;
    }
    if (token.isCancellationRequested) throw this.deps.cancelled();
    if (kept) { update = kept.update; meta = kept.meta; }
    // A Join the page has not finished, or not edited, answers with nothing: the room is recorded, so that a restart joins it
    // again rather than opening a new, empty document (Fable's review, section 9.19 finding 1).
    else if (this.join) meta = { ...meta, room: this.join, joined: true };
    const file = encodeBackup({
      uri,
      name: this.name,
      ...(this.source ? { source: String(this.source.uri) } : {}),
      meta,
      base: this.base,
      update,
      ...(disk ? { disk } : {}),
      ...(sourceBytes ? { sourceBytes } : {}),
    });
    await this.deps.io.write(destination, file);
    if (token.isCancellationRequested) {
      // The workbench will not store it, so nothing would delete it.
      await this.deps.io.delete(destination).catch(() => undefined);
      throw this.deps.cancelled();
    }
    // An `empty` answer (nothing beyond the bytes) resumes as a plain open, with no `resume`.
    this.backup = kept;
  }

  /** VS Code's Revert File, and the banner's Reload of a clean document: the file read again and opened in place. */
  async revert(): Promise<void> {
    this.generation++;
    const bytes = await this.deps.io.read(this.uri);
    this.base = bytes;
    this.disk = bytes;
    this.backup = undefined;
    this.dirty = false;
    this.channel?.post({ type: 'revert', bytes });
  }

  /** The banner's Reload of a document made from Markdown: the source read again and converted in place. */
  async revertToSource(): Promise<void> {
    if (!this.source) return;
    this.generation++;
    const bytes = await this.deps.io.read(this.source.uri);
    this.source = { ...this.source, bytes };
    this.channel?.post({ type: 'revert', bytes });
  }
}
