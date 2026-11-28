// SPDX-License-Identifier: GPL-3.0-or-later
//
// The custom editors (ED-014 section 4.3): `DocxEditorProvider` for the four Word types, and the Markdown launcher (P26).
// VS Code's calls go to the document (`document.ts`), and the page's messages come here. Every message of
// `host-protocol.ts` has its half here; the companion's are passed to `agent.ts` (section 4.5).
import * as vscode from 'vscode';
import { Channel } from './channel.js';
import { DocxDocument, resumeOf, sameBytes, type DocumentDeps, type DocumentInit, type Handover } from './document.js';
import { Handovers } from './handover.js';
import { decodeBackup } from './backup.js';
import { baseName, isSaveAsName, linkTarget, markdownTarget, untitledName } from './names.js';
import { writeNotice, writePage, type App } from './panel.js';
import type { OptionsStore } from './options.js';
import { parseRoomLink, type FromPage, type RoomLinkData, type ToPage } from './shared.js';
import type { Agent } from './agent.js';
import { agentWords } from './agent.js';
import type { Ask } from './ask.js';
import { connectFor, joinOf, pageReaches } from './companion.js';

export const WORD_VIEW_TYPE = 'docx4j.editor';
export const MARKDOWN_VIEW_TYPE = 'docx4j.markdown';

/** A byte array as a message or the file system gives it: a `Uint8Array` of its own, never a Node `Buffer` (S0). */
export function ownBytes(value: unknown): Uint8Array {
  if (value instanceof Uint8Array) return new Uint8Array(value);
  if (value instanceof ArrayBuffer) return new Uint8Array(value.slice(0));
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength));
  const json = value as { type?: unknown; data?: unknown } | null;
  if (json && json.type === 'Buffer' && Array.isArray(json.data)) return Uint8Array.from(json.data as number[]);
  return new Uint8Array();
}

async function readBytes(uri: vscode.Uri): Promise<Uint8Array> {
  return ownBytes(await vscode.workspace.fs.readFile(uri));
}

async function readIfAny(uri: vscode.Uri): Promise<Uint8Array | undefined> {
  try { return await readBytes(uri); } catch (e) {
    if (e instanceof vscode.FileSystemError && e.code === 'FileNotFound') return undefined;
    throw e;
  }
}

class VsDocx extends DocxDocument<vscode.Uri> implements vscode.CustomDocument {
  /** The bytes last reported to the page as a change on disk, so one change is reported once. */
  reported: Uint8Array | undefined;
  /** Opened from a backup: what changed on disk while VS Code was closed, reported once the page has said `ready`. */
  reportOnOpen: Uint8Array | undefined;
  /** VS Code must be told it is dirty: an untitled document with content, a hand-over's late edits. */
  dirtyAtOpen = false;
  /** The origins the pages posted with `ready`, for the companion's allowance (S5). */
  readonly origins = new Set<string>();
  /** The last backup file written, which VS Code leaves behind when a Save As replaces the editor (section 9.9). */
  lastBackup: vscode.Uri | undefined;
  /** The panel showing the document, beside its channel: what Commands and Paste as Plain Text reveal (section 4.4). */
  panel: vscode.WebviewPanel | undefined;
  /** The companion's relay its page's policy names (section 4.5): a page written before the ports moved cannot reach the new one. */
  writtenRelay: string | undefined;
  private watcher: vscode.Disposable | undefined;
  private readonly onDispose: () => void;

  constructor(uri: vscode.Uri, init: DocumentInit<vscode.Uri>, deps: DocumentDeps<vscode.Uri>, onDispose: () => void) {
    super(uri, init, deps);
    this.onDispose = onDispose;
  }

  watch(watcher: vscode.Disposable): void {
    this.watcher?.dispose();
    this.watcher = watcher;
  }

  dispose(): void {
    this.watcher?.dispose();
    this.channel?.close();
    this.onDispose();
  }
}

/** What waits for an untitled document's `openCustomDocument` (P17). */
interface Waiting {
  bytes: Uint8Array;
  source?: { uri: vscode.Uri; name: string; bytes: Uint8Array };
  /** A room to join (section 4.5). */
  join?: RoomLinkData;
}

const words = {
  changedOnDisk: (name: string) => vscode.l10n.t('{0} has changed on disk since it was opened or saved: save it under another name, or reload it.', name),
  noPage: (name: string) => vscode.l10n.t('The editor for {0} is not ready.', name),
};

export class DocxEditorProvider implements vscode.CustomEditorProvider<VsDocx> {
  private readonly changes = new vscode.EventEmitter<vscode.CustomDocumentContentChangeEvent<VsDocx>>();
  readonly onDidChangeCustomDocument = this.changes.event;
  /** Every document of this editor, by its URI. */
  readonly documents = new Map<string, VsDocx>();
  private readonly waiting = new Map<string, Waiting>();
  private readonly handovers = new Handovers<Handover>();

  constructor(private readonly app: App, private readonly options: OptionsStore, private readonly agent?: Agent, private readonly ask?: Ask) {
    // Every page hears the companion's address and its relay's when it starts (section 4.5, P38).
    agent?.onDidStart((running) => this.toEveryPage({ type: 'companion', url: running.url, relay: running.relay }));
    // The Ask box (section 4.6): the key set or removed, and VS Code's models changed, told to every page.
    ask?.onDidChangeKey((stored) => this.toEveryPage({ type: 'askKeyState', stored }));
    ask?.onDidChangeModels((models) => this.toEveryPage({ type: 'askModels', models }));
  }

  private toEveryPage(message: ToPage): void {
    for (const doc of this.documents.values()) doc.channel?.post(message);
  }

  /** Whether a page is in a room on this relay: the companion is then restarted at once if it exits (section 4.5). */
  roomsOn(relay: string): boolean {
    return [...this.documents.values()].some((d) => d.room === relay);
  }

  private deps(document: () => VsDocx): DocumentDeps<vscode.Uri> {
    return {
      io: {
        read: readBytes,
        readIfAny,
        write: (uri, bytes) => Promise.resolve(vscode.workspace.fs.writeFile(uri, bytes)),
        delete: (uri) => Promise.resolve(vscode.workspace.fs.delete(uri)),
      },
      cancelled: () => new vscode.CancellationError(),
      words,
      later: (f) => { setTimeout(f, 0); },
      changed: () => this.changes.fire({ document: document() }),
    };
  }

  async openCustomDocument(uri: vscode.Uri, context: vscode.CustomDocumentOpenContext): Promise<VsDocx> {
    const key = uri.toString();
    const name = baseName(uri.path);
    const untitled = uri.scheme === 'untitled';
    const init = await this.initOf(uri, name, untitled, context.backupId);
    let doc: VsDocx | undefined;
    doc = new VsDocx(uri, init.init, this.deps(() => doc!), () => { if (this.documents.get(key) === doc) this.documents.delete(key); });
    doc.reportOnOpen = init.report;
    doc.dirtyAtOpen = init.dirtyAtOpen;
    this.documents.set(key, doc);
    if (init.note) void vscode.window.showWarningMessage(init.note);
    // A backup's room on the companion's relay: the companion started now, so that the room answers within the page's wait.
    const room = init.init.backup?.meta.room?.relay ?? init.init.join?.relay;
    if (room !== undefined && this.agent && room === this.agent.keptRelay()) void this.agent.start().catch(() => undefined);
    return doc;
  }

  private async initOf(uri: vscode.Uri, name: string, untitled: boolean, backupId: string | undefined): Promise<{ init: DocumentInit<vscode.Uri>; report?: Uint8Array; dirtyAtOpen: boolean; note?: string }> {
    let note: string | undefined;
    if (backupId !== undefined) {
      try {
        const file = decodeBackup(await readBytes(vscode.Uri.parse(backupId)));
        const sourceUri = file.source !== undefined ? vscode.Uri.parse(file.source) : undefined;
        const sourceBytes = sourceUri ? await readIfAny(sourceUri) : undefined;
        const fileNow = untitled ? undefined : await readIfAny(uri);
        // Named and watched by its own URI, which after a rename is not the backup header's.
        const { init, report } = resumeOf(file, name, untitled, {
          ...(fileNow ? { file: fileNow } : {}),
          ...(sourceUri && sourceBytes ? { source: { uri: sourceUri, name: baseName(sourceUri.path), bytes: sourceBytes } } : {}),
        });
        return { init, ...(report ? { report } : {}), dirtyAtOpen: false };
      } catch (e) {
        note = vscode.l10n.t('The unsaved edits to {0} could not be restored: {1}', name, (e as Error).message);
      }
    }
    if (untitled) {
      const waiting = this.waiting.get(uri.toString());
      this.waiting.delete(uri.toString());
      // A joined room is content too, or VS Code keeps no backup of it and a restart opens it empty (section 9.16 finding 5).
      const content = (waiting?.bytes.byteLength ?? 0) > 0 || waiting?.source !== undefined || waiting?.join !== undefined;
      return {
        init: { name, untitled, base: waiting?.bytes ?? new Uint8Array(), ...(waiting?.source ? { source: waiting.source } : {}), ...(waiting?.join ? { join: waiting.join } : {}), dirty: content },
        // An untitled document with content is dirty from its open, as an untitled text file is (section 9.8 finding 3).
        dirtyAtOpen: content,
        ...(note ? { note } : {}),
      };
    }
    const handover = this.handovers.take(uri.toString());
    if (handover) {
      if (handover.resume) {
        // An edit came after the export (section 9.8 finding 1), or the session is in a room (P42): a resume over the old
        // document's base, named as the target, saved unless the hand-over says otherwise.
        const { source: _source, ...meta } = handover.resume.meta;
        return {
          init: { name, untitled, base: handover.base, disk: handover.bytes, backup: { update: handover.resume.update, meta: { ...meta, name } }, ...(handover.place !== undefined ? { place: handover.place } : {}), dirty: !handover.clean },
          dirtyAtOpen: !handover.clean,
        };
      }
      return { init: { name, untitled, base: handover.bytes, disk: handover.bytes, ...(handover.place !== undefined ? { place: handover.place } : {}) }, dirtyAtOpen: false };
    }
    const bytes = await readBytes(uri);
    return { init: { name, untitled, base: bytes, disk: bytes }, dirtyAtOpen: false, ...(note ? { note } : {}) };
  }

  async resolveCustomEditor(document: VsDocx, panel: vscode.WebviewPanel): Promise<void> {
    const channel = new Channel((message) => { void panel.webview.postMessage(message); });
    document.channel?.close();
    document.channel = channel;
    document.panel = panel;
    // The companion's relay named in the policy whether or not it runs (P37), and a room's on another relay (section 4.5).
    const can = this.agent?.canRun();
    const relay = can?.ok ? await this.agent!.relay().catch(() => undefined) : undefined;
    document.writtenRelay = relay;
    const roomRelay = document.join?.relay ?? document.backup?.meta.room?.relay ?? document.room;
    writePage(panel.webview, this.app, {
      options: this.options.all(),
      language: vscode.env.language,
      trusted: vscode.workspace.isTrusted,
      ...(can?.ok ? { agent: this.agent!.name() } : {}),
      // Whether a key is stored, never the key (P45).
      ...(this.ask ? { askKey: await this.ask.keyStored().catch(() => false) } : {}),
      commands: this.app.commands,
    }, connectFor(relay, roomRelay));
    panel.webview.onDidReceiveMessage((m: FromPage) => this.onMessage(document, panel, channel, m));
    panel.onDidDispose(() => {
      // A request still running for this page is cancelled: no answer reaches a closed editor (section 9.20 item 5).
      this.ask?.service.pageGone(channel);
      channel.close();
      if (document.channel === channel) document.channel = undefined;
      if (document.panel === panel) document.panel = undefined;
    });
    if (document.dirtyAtOpen) {
      document.dirtyAtOpen = false;
      this.changes.fire({ document });
    }
    this.watch(document);
  }

  saveCustomDocument(document: VsDocx, token: vscode.CancellationToken): Promise<void> {
    return document.save(token);
  }

  async saveCustomDocumentAs(document: VsDocx, target: vscode.Uri, token: vscode.CancellationToken): Promise<void> {
    const name = baseName(target.path);
    // A Word document saved as `notes.md` would open in VS Code's text editor (P25).
    if (!isSaveAsName(name)) throw new Error(vscode.l10n.t('{0} is not a Word document\'s name: save it as .docx, .dotx, .docm or .dotm, or as .xml for Flat OPC.', name));
    const same = target.toString() === document.uri.toString();
    // VS Code would show the open editor, stale (section 9.8 finding 6).
    if (!same && this.documents.has(target.toString())) throw new Error(vscode.l10n.t('{0} is open in another editor: close it first, or save under another name.', name));
    const handover = await document.saveAs(target, name, same, token);
    if (!handover) return;
    this.handovers.put(target.toString(), handover);
    // VS Code deletes a backup after a save and a revert, but not after a Save As that replaces the editor (measured,
    // section 9.9): the last one would stay in the extension's storage for good. The document ends here, so it goes.
    const stale = document.lastBackup;
    document.lastBackup = undefined;
    if (stale) void Promise.resolve(vscode.workspace.fs.delete(stale)).catch(() => undefined);
  }

  /**
   * Before a rename or a move in the explorer or through a workspace edit (section 9.9): VS Code reopens a renamed custom
   * editor from the file at its new URI, clean, and gives it no backup (`CustomEditorInput.rename` names no old resource),
   * so a dirty document's edits would be lost with no prompt. Each dirty document the rename takes, the file itself or one
   * in a folder renamed, hands its store over to its new URI, as a Save As does (P25).
   */
  async beforeRename(files: readonly { readonly oldUri: vscode.Uri; readonly newUri: vscode.Uri }[]): Promise<void> {
    for (const { oldUri, newUri } of files) {
      const from = oldUri.toString();
      const inside = from.endsWith('/') ? from : `${from}/`;
      for (const doc of [...this.documents.values()]) {
        const key = doc.uri.toString();
        // A document in a room is handed over clean or dirty, so that it stays in its room (P42).
        if (doc.untitled || (!doc.dirty && doc.room === undefined) || (key !== from && !key.startsWith(inside))) continue;
        const target = key === from ? newUri.toString() : `${newUri.toString().replace(/\/$/, '')}/${key.slice(inside.length)}`;
        try {
          this.handovers.put(target, await doc.handOver());
        } catch {
          // The page has gone: VS Code reopens the file as it is on disk.
        }
      }
    }
  }

  revertCustomDocument(document: VsDocx): Promise<void> {
    document.reported = undefined;
    return document.revert();
  }

  async backupCustomDocument(document: VsDocx, context: vscode.CustomDocumentBackupContext, token: vscode.CancellationToken): Promise<vscode.CustomDocumentBackup> {
    const { destination } = context;
    await document.writeBackup(destination, document.uri.toString(), token);
    document.lastBackup = destination;
    return {
      id: destination.toString(),
      delete: () => {
        if (document.lastBackup?.toString() === destination.toString()) document.lastBackup = undefined;
        void Promise.resolve(vscode.workspace.fs.delete(destination)).catch(() => undefined);
      },
    };
  }

  // ---- The commands that reach the page (ED-014 section 4.4).

  /**
   * The document of the active editor group's active tab, if this editor shows it. Through the tab, not the focus: VS Code's
   * palette and the editor's tab take the focus from the page, and a custom editor's context keys follow the active editor.
   */
  activeDocument(): VsDocx | undefined {
    const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
    if (!(input instanceof vscode.TabInputCustom) || input.viewType !== WORD_VIEW_TYPE) return undefined;
    return this.documents.get(input.uri.toString());
  }

  /** The page shown with the focus, and the message posted: it waits in the channel until the page has said `ready`. */
  private toPage(document: VsDocx, message: ToPage): void {
    document.panel?.reveal(undefined, false);
    document.channel?.post(message);
  }

  /** "docx4j editor: Commands" (P33): the application's palette, its field focused. Nothing with no document of this editor active. */
  showCommands(): void {
    const document = this.activeDocument();
    if (document) this.toPage(document, { type: 'palette' });
  }

  /**
   * "Paste as Plain Text" from VS Code's menus (P34): the clipboard's text, pasted where the page's focus was, which the frame
   * keeps across VS Code's context menu. The page's own Ctrl+Shift+V reads the clipboard itself.
   */
  async pasteText(): Promise<void> {
    const document = this.activeDocument();
    if (!document) return;
    const text = await vscode.env.clipboard.readText();
    if (text !== '') this.toPage(document, { type: 'pasteText', text });
  }

  // ---- The page's messages.

  private onMessage(document: VsDocx, panel: vscode.WebviewPanel, channel: Channel, m: FromPage): void {
    if (!m || typeof m !== 'object' || typeof m.type !== 'string') return;
    // A message from a page this document no longer shows.
    if (document.channel !== channel) return;
    switch (m.type) {
      case 'ready': {
        if (typeof m.origin === 'string') {
          document.origins.add(m.origin);
          // Kept for every start of the companion, and allowed at its relay now if it runs (P38).
          if (this.agent?.canRun().ok) void this.agent.allowOrigin(m.origin).catch(() => undefined);
        }
        // A page that loads again is a new page: the old one's requests are cancelled (section 9.20 item 5).
        if (channel.loads > 0) this.ask?.service.pageGone(channel);
        channel.ready(document.openMessage());
        const running = this.agent?.running;
        if (running) channel.post({ type: 'companion', url: running.url, relay: running.relay });
        if (document.reportOnOpen) { const bytes = document.reportOnOpen; document.reportOnOpen = undefined; void this.reportChanged(document, bytes).catch(() => undefined); }
        break;
      }
      case 'companionStart': void this.companionStart(document, channel, m.id); break;
      case 'invite': void this.invite(channel, m.id, m.room, m.token); break;
      case 'room': document.room = typeof m.relay === 'string' ? m.relay : undefined; break;
      // A Join keeps its room: a backup the page answers with nothing still records it (section 9.19 finding 1).
      case 'base': document.base = ownBytes(m.bytes); break;
      case 'changed': document.changed(); break;
      case 'saved':
      case 'backup':
      case 'handoff':
        // Every byte array as one of its own, a hand-over's store included.
        channel.answer({
          ...m,
          ...('bytes' in m ? { bytes: ownBytes(m.bytes) } : {}),
          ...('update' in m ? { update: ownBytes(m.update) } : {}),
          ...('resume' in m && m.resume ? { resume: { ...m.resume, update: ownBytes(m.resume.update) } } : {}),
        } as FromPage);
        break;
      case 'revertRequest': void this.revertRequest(document, panel).catch((e) => this.warn(e)); break;
      case 'untitled':
        if ('bytes' in m) void this.openUntitled(m.name, ownBytes(m.bytes), this.folderOf(document)).catch((e) => this.warn(e));
        else if ('join' in m) { const link = joinOf(m.join); if (link) void this.openJoin(link, this.folderOf(document)).catch((e) => this.warn(e)); }
        break;
      case 'openExternal': void this.follow(document, m.url).catch((e) => this.warn(e)); break;
      case 'download': void this.download(document, ownBytes(m.bytes), m.name).catch((e) => this.warn(e)); break;
      case 'status': void vscode.window.showWarningMessage(m.text); break;
      case 'options': void this.shareOptions(m, channel).catch((e) => this.warn(e)); break;
      // The Ask box (section 4.6, P44 to P46).
      case 'askModels': if (typeof m.id === 'number') void this.askModels(channel, m.id); break;
      case 'ask': if (typeof m.id === 'number') void this.askFor(channel, m.id, m.route, m.request); break;
      case 'askCancel': if (typeof m.id === 'number') this.ask?.service.cancel(channel, m.id); break;
      case 'askKey': this.ask?.action(m.action); break;
    }
  }

  // ---- The Ask box (section 4.6).

  // Each answer goes through the channel's `replier`, taken before the first await: a page that has loaded again since
  // starts its ids again, and an answer meant for the old page would settle one of the new page's (section 9.23 finding 2).
  private async askModels(channel: Channel, id: number): Promise<void> {
    const reply = channel.replier();
    const models = this.ask ? await this.ask.service.models() : [];
    reply({ type: 'askModels', id, models });
  }

  private async askFor(channel: Channel, id: number, route: unknown, request: unknown): Promise<void> {
    const reply = channel.replier();
    if (!this.ask) { reply({ type: 'asked', id, failure: { kind: 'other', message: 'the Ask box is not available here' } }); return; }
    reply(await this.ask.service.ask(channel, id, route, request));
  }

  private warn(e: unknown): void {
    void vscode.window.showWarningMessage((e as Error)?.message ?? String(e));
  }

  // ---- The companion (section 4.5).

  /** Invite and Share ask for the companion: started, the page's origin allowed, its address and its relay's answered. */
  private async companionStart(document: VsDocx, channel: Channel, id: number): Promise<void> {
    const reply = channel.replier();
    if (!this.agent) { reply({ type: 'companion', id, error: agentWords.cannotRun('remote') }); return; }
    try {
      for (const origin of document.origins) await this.agent.allowOrigin(origin);
      const running = await this.agent.start();
      // A page written before the ports moved names the old relay port in its policy, and cannot reach the new one.
      if (!pageReaches(document.writtenRelay, running.relay)) { reply({ type: 'companion', id, error: agentWords.reopen() }); return; }
      reply({ type: 'companion', id, url: running.url, relay: running.relay });
    } catch (e) {
      reply({ type: 'companion', id, error: (e as Error).message });
    }
  }

  /** The agent invited into the page's room: the person has admitted it in this workspace (P40). */
  private async invite(channel: Channel, id: number, room: string, token: string): Promise<void> {
    const reply = channel.replier();
    try {
      if (!this.agent) throw new Error(agentWords.cannotRun('remote'));
      await this.agent.invite(room, token);
      reply({ type: 'invited', id });
    } catch (e) {
      reply({ type: 'invited', id, error: (e as Error).message });
    }
  }

  /**
   * "docx4j editor: Join a Room…" (P17, section 4.5): a link asked for, checked as the page's Join checks one, and opened in an
   * untitled editor of its own. A link to the companion's own relay starts the companion first.
   */
  async joinCommand(): Promise<void> {
    const text = await vscode.window.showInputBox({
      title: vscode.l10n.t('Join a Room'),
      prompt: vscode.l10n.t('Paste the room\'s link, as Share gives it.'),
      placeHolder: 'http://127.0.0.1:…/#room=…&token=…&relay=…',
      ignoreFocusOut: true,
      validateInput: (value) => (value.trim() === '' || parseRoomLink(value.trim()) ? undefined : vscode.l10n.t('This is not a room\'s link: it names a room, a token and a relay.')),
    });
    if (text === undefined || text.trim() === '') return;
    const link = parseRoomLink(text.trim());
    if (!link) return;
    if (this.agent && link.relay === this.agent.keptRelay()) {
      try { await this.agent.start(); } catch (e) { this.warn(e); return; }
    }
    await this.openJoin(link, this.folderOf(undefined));
  }

  /** A room in an untitled `Room.docx` of its own: the page joins it, and posts the room's base file (section 4.5). */
  async openJoin(link: RoomLinkData, folder: vscode.Uri | undefined): Promise<void> {
    const open = (n: string): boolean => [...this.documents.values()].some((d) => d.untitled && d.name === n)
      || [...this.waiting.keys()].some((k) => baseName(vscode.Uri.parse(k).path) === n);
    const free = untitledName('Room.docx', open);
    const path = folder && folder.scheme === 'file' ? vscode.Uri.joinPath(folder, free).path : free;
    const uri = vscode.Uri.from({ scheme: 'untitled', path });
    this.waiting.set(uri.toString(), { bytes: new Uint8Array(), join: link });
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, WORD_VIEW_TYPE);
    } finally {
      this.waiting.delete(uri.toString());
    }
  }

  /** After a resume: the file, or the Markdown source, changed while VS Code was closed (`resumeOf`). */
  private async reportChanged(document: VsDocx, bytes: Uint8Array): Promise<void> {
    document.reported = bytes;
    const watched = document.untitled ? document.source?.uri : document.uri;
    const stat = watched ? await vscode.workspace.fs.stat(watched) : undefined;
    document.channel?.post({ type: 'disk', bytes, lastModified: stat?.mtime ?? Date.now() });
  }

  /**
   * The banner's Reload (P18). A dirty document goes through VS Code's Revert File, which clears its dirty state; a clean
   * one VS Code would not revert, so the extension does. An untitled document (a Markdown conversion) Revert File leaves
   * alone (section 9.8 finding 2): its source is read and converted again, and VS Code's dirty state stays, rightly.
   */
  private async revertRequest(document: VsDocx, panel: vscode.WebviewPanel): Promise<void> {
    document.reported = undefined;
    if (document.untitled) { await document.revertToSource(); return; }
    if (document.dirty) {
      panel.reveal(undefined, false);
      await vscode.commands.executeCommand('workbench.action.files.revert');
      return;
    }
    await document.revert();
  }

  private async shareOptions(message: ToPage & { type: 'options' }, from: Channel): Promise<void> {
    const changed = await this.options.apply(message);
    if (!changed) return;
    for (const doc of this.documents.values()) if (doc.channel && doc.channel !== from) doc.channel.post(changed);
  }

  // ---- Untitled documents (P17).

  /** The folder a document's new documents and downloads go in: its own, else the first workspace folder's. */
  folderOf(document: VsDocx | undefined): vscode.Uri | undefined {
    if (document && !document.untitled) return vscode.Uri.joinPath(document.uri, '..');
    // An untitled document made beside a file has that file's folder as its path's.
    if (document?.untitled && document.uri.path.startsWith('/')) return vscode.Uri.joinPath(vscode.Uri.file(document.uri.path), '..');
    return vscode.workspace.workspaceFolders?.[0]?.uri;
  }

  /** A new editor for these bytes, or for a Markdown source (P17, P26): an untitled document whose first Save is a Save As. */
  async openUntitled(name: string, bytes: Uint8Array, folder: vscode.Uri | undefined, source?: Waiting['source']): Promise<void> {
    const open = (n: string): boolean => [...this.documents.values()].some((d) => d.untitled && d.name === n)
      || [...this.waiting.keys()].some((k) => baseName(vscode.Uri.parse(k).path) === n);
    const free = untitledName(name, open);
    const path = folder && folder.scheme === 'file' ? vscode.Uri.joinPath(folder, free).path : free;
    const uri = vscode.Uri.from({ scheme: 'untitled', path });
    this.waiting.set(uri.toString(), { bytes, ...(source ? { source } : {}) });
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, WORD_VIEW_TYPE);
    } finally {
      this.waiting.delete(uri.toString());
    }
  }

  // ---- The routes the page asks for.

  /** Ctrl+click on a link (P28). */
  private async follow(document: VsDocx, address: string): Promise<void> {
    const target = linkTarget(address);
    let uri: vscode.Uri;
    switch (target.kind) {
      case 'external':
        await vscode.env.openExternal(vscode.Uri.parse(target.url));
        return;
      case 'refused':
        void vscode.window.showWarningMessage(vscode.l10n.t('The link was not followed: a document\'s {0} links are not opened.', target.scheme || address));
        return;
      case 'file': uri = vscode.Uri.parse(target.uri); break;
      case 'path': uri = vscode.Uri.file(target.path); break;
      case 'relative':
        if (document.untitled) {
          void vscode.window.showWarningMessage(vscode.l10n.t('The link was not followed: {0} is relative, and an unsaved document has no folder.', address));
          return;
        }
        uri = vscode.Uri.joinPath(document.uri, '..', target.path);
        break;
    }
    let stat: vscode.FileStat;
    try { stat = await vscode.workspace.fs.stat(uri); } catch {
      void vscode.window.showWarningMessage(vscode.l10n.t('The link was not followed: {0} was not found.', uri.fsPath || uri.toString()));
      return;
    }
    if (stat.type & vscode.FileType.Directory) {
      void vscode.window.showWarningMessage(vscode.l10n.t('The link was not followed: {0} is a folder.', uri.fsPath || uri.toString()));
      return;
    }
    await vscode.commands.executeCommand('vscode.open', uri);
  }

  /** A download the page starts (the self-check's diagnosis file): VS Code's save dialog at the document's folder (S0). */
  private async download(document: VsDocx, bytes: Uint8Array, name: string): Promise<void> {
    const folder = this.folderOf(document);
    const uri = await vscode.window.showSaveDialog({ ...(folder ? { defaultUri: vscode.Uri.joinPath(folder, baseName(name)) } : {}), saveLabel: vscode.l10n.t('Save') });
    if (uri) await vscode.workspace.fs.writeFile(uri, bytes);
  }

  // ---- The watcher (ED-011 in the extension, P9).

  /**
   * The document's file, or its Markdown source while it has one. A change or a creation is read after 100 ms of quiet and
   * reported if it differs from the bytes last read or written and from any being written: the extension's own writes
   * never report (ED-011 section 9.11). A deletion reports nothing; the next Save writes the file again.
   */
  private watch(document: VsDocx): void {
    const target = document.untitled ? document.source?.uri : document.uri;
    if (!target || target.scheme === 'untitled') return;
    const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(vscode.Uri.joinPath(target, '..'), baseName(target.path)));
    let timer: ReturnType<typeof setTimeout> | undefined;
    const check = (): void => {
      clearTimeout(timer);
      timer = setTimeout(() => void (async () => {
        const bytes = await readIfAny(target).catch(() => undefined);
        if (!bytes) return;
        const known = document.untitled ? document.source?.bytes : document.disk;
        if (sameBytes(bytes, known) || sameBytes(bytes, document.writing) || sameBytes(bytes, document.reported)) return;
        document.reported = bytes;
        const stat = await vscode.workspace.fs.stat(target);
        document.channel?.post({ type: 'disk', bytes, lastModified: stat.mtime });
      })(), 100);
    };
    watcher.onDidChange(check);
    watcher.onDidCreate(check);
    document.watch({ dispose: () => { clearTimeout(timer); watcher.dispose(); } });
  }
}

/**
 * The Markdown editor as a launcher (P26): Open With on a `.md` opens its conversion as an untitled `notes.docx` beside it
 * in the Word editor, and closes its own tab. Read-only, so VS Code never saves or backs it up.
 */
export class MarkdownLauncher implements vscode.CustomReadonlyEditorProvider {
  constructor(private readonly word: DocxEditorProvider) {}

  openCustomDocument(uri: vscode.Uri): vscode.CustomDocument {
    return { uri, dispose: () => undefined };
  }

  async resolveCustomEditor(document: vscode.CustomDocument, panel: vscode.WebviewPanel): Promise<void> {
    writeNotice(panel.webview, vscode.l10n.t('Opening as a Word document…'));
    const name = baseName(document.uri.path);
    const bytes = await readBytes(document.uri);
    const folder = document.uri.scheme === 'file' ? vscode.Uri.joinPath(document.uri, '..') : this.word.folderOf(undefined);
    await this.word.openUntitled(markdownTarget(name), new Uint8Array(), folder, { uri: document.uri, name, bytes });
    const tab = vscode.window.tabGroups.all.flatMap((g) => g.tabs).find((t) => t.input instanceof vscode.TabInputCustom
      && t.input.viewType === MARKDOWN_VIEW_TYPE && t.input.uri.toString() === document.uri.toString());
    if (tab) await vscode.window.tabGroups.close(tab);
    else panel.dispose();
  }
}
