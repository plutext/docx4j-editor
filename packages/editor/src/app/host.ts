// SPDX-License-Identifier: GPL-3.0-or-later
//
// The host seam (ED-014 section 3.2, designed in section 4.1): what the application asks of whatever
// it runs in, and what that host may ask of the application. Two hosts: a browser tab (`browserHost`,
// today's behaviour, moved) and VS Code's webview (`vscodeHost`, the page half of the extension's
// protocol, `host-protocol.ts`). The application calls `host()` and never asks which host it is in:
// what a host leaves out is a feature it does not have (`has`).
//
// In a browser one tab holds one document and replaces it in place. In VS Code one editor is one file,
// so the routes sort three ways (section 4.1): Reload and Revert stay in place; New, From template and
// a room joined go to an editor of their own (P17, `openDocument`); Open, Close, Save and Save As are
// VS Code's own.
import { download as downloadFile, type OpenedFile } from './files.js';
import type { RoomLink } from './room.js';
import type { EditCause } from './editcause.js';
import {
  HOST_META, CHANGED_INTERVAL, isSharedOption,
  type FromPage, type ToPage, type HostBoot, type OptionsMessage, type AskRequestData, type AskAnswerData, type AskFailureData, type HostModelData,
} from './host-protocol.js';
import { installKeyGate } from './keygate.js';

/**
 * What a host leaves in the application or takes out (ED-014 section 3.8). The browser has every one but `clipboardText`.
 * VS Code has `clipboardText`, and `documentMacros` only in a trusted workspace (P27), and none of the others.
 * - `clipboardText`: the page's own paste as plain text (P34), where the host cancels the browser's Ctrl+Shift+V.
 * - `spelling`: the "Check spelling as you type" option, which VS Code's windows, with no spell checker, cannot honour (P35).
 * - `companionOptions`: the relay and companion options, which VS Code's extension sets itself (P41).
 * - `askOptions`: the Ask box's key and address in Options, which VS Code keeps itself, the key in its SecretStorage and the
 *   address in a setting (ED-014 section 4.6, P45).
 */
export type HostFeature =
  | 'open' | 'new' | 'close' | 'saveInPlace' | 'saveAs' | 'download' | 'print' | 'sendCopy' | 'join'
  | 'recovery' | 'diskInTab' | 'layoutChoice' | 'startPage' | 'appKeys' | 'unloadGuard' | 'documentMacros'
  | 'clipboardText' | 'spelling' | 'companionOptions' | 'askOptions';

/** What a hot-exit backup holds beside the store's update: enough to open it again as it was (ED-014 section 4.1). */
export interface BackupMeta {
  name: string;
  /** The Markdown file the document was made from, whose conversion the backup is over (ED-011 section 3.8). */
  source?: string;
  /** The room the session was in: rejoined only if its relay answers in time. */
  room?: RoomLink;
  /** This peer joined the room, so its bytes are the room's base file. */
  joined?: true;
  /** What first made the document unsaved (ED-011 section 9.9). */
  firstEdit?: EditCause;
}

/** A hot-exit backup: `Y.encodeStateAsUpdate` of the store, and its meta. */
export interface HostBackup {
  update: Uint8Array;
  meta: BackupMeta;
}

/** A backup to resume over the bytes it was taken over. */
export type HostResume = HostBackup;

/** A room to join: it has no bytes, only its link. */
export interface JoinSource {
  name: string;
  join: RoomLink;
}

/** How a host opens a document: a hot-exit backup to resume, a place to restore (a Save As's hand-over, P25). */
export interface HostOpenHow {
  resume?: HostResume;
  /** The page's `KeptPlace`, as another page kept it: restored once open, as a reload keeps the caret. */
  place?: unknown;
  /** The resume opens as saved: a hand-over in a room with nothing after the save (P42). */
  clean?: boolean;
}

/**
 * The companion as a host starts it for the page (VS Code, ED-014 section 4.5, P38): the page asks the host, and reaches the
 * relay alone directly. A browser's page finds its own companion by its probe.
 */
export interface HostCompanion {
  /** The agent's name, from the boot: Invite's label before the companion runs. */
  readonly agent: string;
  /** The companion started if need be: its address and its relay's. Rejects with words to show. */
  start(): Promise<{ url: string; relay: string }>;
  /** The agent invited into a room on the companion's relay, which the page has shared. Rejects with why not. */
  invite(link: RoomLink): Promise<void>;
}

/** A model the host offers (VS Code's `vscode.lm`), known by its vendor and its id. */
export type HostModel = HostModelData;

/** Why the host's request has no answer: the page words it by its kind (`describeFailure`). */
export class AskFailure extends Error {
  constructor(readonly failure: AskFailureData) {
    super(failure.message);
    this.name = 'AskFailure';
  }
}

/**
 * The Ask box's requests, where the host sends them (VS Code, ED-014 section 4.6, P44): the page holds no key and names no
 * outside address. A browser's page sends its own.
 */
export interface HostAsk {
  /** Whether an Anthropic key is stored: from the boot, then kept current by the host's pushes. */
  keyStored(): boolean;
  onKeyStored(listener: (stored: boolean) => void): () => void;
  /**
   * VS Code's models: asked at the Ask box's first show, then kept current by the host's pushes. `fresh` asks again (the list
   * opened, a model gone); otherwise the list last heard.
   */
  models(fresh?: boolean): Promise<HostModel[]>;
  onModels(listener: (models: HostModel[]) => void): () => void;
  /** Sent by the host. Rejects with an `AskFailure`; a cancelled request rejects with an `AbortError`. */
  send(route: 'vscode' | 'key', request: AskRequestData, signal: AbortSignal): Promise<AskAnswerData>;
  /** The Options dialog's buttons: the key's commands, and the address's setting. */
  key(action: 'set' | 'remove' | 'settings'): void;
}

/** The application, as a host calls it. App implements it and attaches it at mount. */
export interface HostPage {
  /** Empty bytes are a new document, made as New makes one; a Markdown name is converted; with `resume`, a hot-exit backup over the bytes. */
  open(source: OpenedFile | JoinSource, how?: HostOpenHow): Promise<void>;
  /** Export with the self-check in the form the name asks for (Flat OPC for `.xml`). A refusal is shown by the page. */
  save(name: string): Promise<{ bytes: Uint8Array } | { refused: string }>;
  /**
   * The host wrote the file: the document is saved under this name. `clean` is false when an edit came after the export,
   * which the page keeps as unsaved.
   */
  written(name: string, clean: boolean): void;
  /** The host could not write the file: the document stays unsaved. */
  writeFailed(error: string): void;
  /** The store's state and the session's room; undefined with no document open. Never waits on the model's queue. */
  backup(): HostBackup | undefined;
  /**
   * A Save As wrote another file, which the host opens afresh (P25): the page stops taking input, for good, and gives its
   * place (its `KeptPlace`, undefined with no view).
   */
  handoff(): { place?: unknown };
  /** Open these bytes again in place, the caret's place kept (ED-011's `keepPlace`): the file's, or a Markdown source's, converted again. */
  revert(bytes: Uint8Array): Promise<void>;
  /** The file changed on disk: the banner, compared with these bytes (a Markdown source's, for a document made from one). */
  diskChanged(bytes: Uint8Array, lastModified: number): void;
  showPalette(): void;
  /** "Paste as Plain Text" from VS Code's menus: the text pasted where the page's focus is, as the page's own key pastes it (P34). */
  pasteText(text: string): void;
  /** The companion's address and its relay's, which win over the options (ED-014 section 4.5). */
  setCompanion(url: string, relay: string): void;
  /** The browser's Reload: the page reads its own handle again (ED-011 section 3.5); `discarding` after the confirmation. */
  reloadFromDisk(discarding: boolean): Promise<void>;
}

export interface Host {
  readonly kind: 'browser' | 'vscode';
  has(feature: HostFeature): boolean;
  /** The application attached; VS Code's host then says `ready`, once per page. Returns the detach. */
  attach(page: HostPage): () => void;
  /** A local edit: VS Code's dirty state and its backups follow. */
  changed(): void;
  /** A document the page made, or a room to join: in place in a browser, in an editor of its own in VS Code (P17). */
  openDocument(file: OpenedFile | JoinSource): Promise<void>;
  /** The banner's Reload: the browser reads its handle again, VS Code reverts the document (P18). */
  reload(discarding?: boolean): void;
  openExternal(url: string): void;
  download(bytes: Uint8Array, name: string): void;
  /** The address a room link is built on: the page's own in a browser, the companion's served copy in VS Code (S5). */
  siteAddress(): string;
  /** The bytes of a document the page made where the host had none (a new document, a room's base file): VS Code keeps them as the document's base. */
  base(bytes: Uint8Array): void;
  /** Something the page said that the host says too, since the page may stand empty (an open that failed, or fell back). */
  status(text: string): void;
  /** The host's display language, which the page takes unless the person chose one (P27): none in a browser, whose own the page reads. */
  language(): string | undefined;
  /**
   * The name the host's command palette shows for one of its commands (VS Code's, ED-014 P49), in the host's language,
   * which may not be the page's; none in a browser.
   */
  commandName(id: string): string | undefined;
  /** The companion, where the host starts it for the page (VS Code); undefined where the page finds its own (a browser). */
  companion(): HostCompanion | undefined;
  /** The session entered a room on this relay, or left its room (none): VS Code's extension keeps it for a rename's hand-over (P42). */
  room(relay: string | undefined): void;
  /** The Ask box's requests, where the host sends them (VS Code); undefined where the page sends its own (a browser). */
  ask(): HostAsk | undefined;
}

/** The page's own address without its fragment, which a room link replaces. */
export function pageAddress(): string {
  return location.href.replace(/#.*$/, '');
}

/** A browser tab: every feature but the page's own plain paste (the browser's Ctrl+Shift+V works), the page's own routes. */
export function browserHost(): Host {
  let page: HostPage | undefined;
  return {
    kind: 'browser',
    has: (feature) => feature !== 'clipboardText',
    attach(p) {
      page = p;
      return () => { if (page === p) page = undefined; };
    },
    changed: () => undefined,
    async openDocument(file) {
      if (!page) throw new Error('the application is not attached');
      await page.open(file);
    },
    reload(discarding = false) { void page?.reloadFromDisk(discarding); },
    openExternal(url) { window.open(url, '_blank', 'noopener'); },
    download: (bytes, name) => downloadFile(bytes, name),
    siteAddress: pageAddress,
    base: () => undefined,
    status: () => undefined,
    language: () => undefined,
    commandName: () => undefined,
    companion: () => undefined,
    room: () => undefined,
    ask: () => undefined,
  };
}

/** VS Code's webview API, as `acquireVsCodeApi()` returns it: once per page. */
export interface VsCodeApi {
  postMessage(message: unknown): void;
}

/** A byte array as a message carries it: a `Uint8Array`, or, posted by mistake, a Node `Buffer` as VS Code serializes one (S0). */
function asBytes(value: unknown): Uint8Array {
  if (value instanceof Uint8Array) return value;
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  const json = value as { type?: unknown; data?: unknown } | null;
  if (json && json.type === 'Buffer' && Array.isArray(json.data)) return Uint8Array.from(json.data as number[]);
  return new Uint8Array();
}

/** The boot's command names, each a string; none where there are none. */
function commandsOf(value: unknown): { commands?: Record<string, string> } {
  if (!value || typeof value !== 'object') return {};
  const commands: Record<string, string> = {};
  for (const [id, name] of Object.entries(value as Record<string, unknown>)) if (typeof name === 'string' && name.trim() !== '') commands[id] = name;
  return Object.keys(commands).length > 0 ? { commands } : {};
}

/** The boot data the extension wrote into the page (`HOST_META`), or undefined where there is none or it does not read. */
export function readHostBoot(doc: Document | undefined = globalThis.document): HostBoot | undefined {
  const content = doc?.querySelector(`meta[name="${HOST_META}"]`)?.getAttribute('content');
  if (!content) return undefined;
  try {
    const value = JSON.parse(content) as { options?: unknown; language?: unknown; trusted?: unknown; agent?: unknown; askKey?: unknown; commands?: unknown };
    if (!value.options || typeof value.options !== 'object') return undefined;
    const options: Record<string, string> = {};
    for (const [key, v] of Object.entries(value.options as Record<string, unknown>)) if (typeof v === 'string') options[key] = v;
    return {
      options,
      ...(typeof value.language === 'string' && value.language !== '' ? { language: value.language } : {}),
      ...(typeof value.trusted === 'boolean' ? { trusted: value.trusted } : {}),
      ...(typeof value.agent === 'string' && value.agent.trim() !== '' ? { agent: value.agent } : {}),
      ...(typeof value.askKey === 'boolean' ? { askKey: value.askKey } : {}),
      ...(commandsOf(value.commands)),
    };
  } catch {
    return undefined;
  }
}

/** VS Code's host, with a stop that puts the page's storage back as it was (for tests, which make more than one). */
export interface VsCodeHost extends Host {
  stop(): void;
}

/**
 * VS Code's webview (ED-014 section 4.1): the page half of the protocol in `host-protocol.ts`. None of the features. Messages
 * that come before the application attaches wait in a queue; `ready` is posted once per page, though React's development
 * StrictMode attaches twice. The options (P19): the boot data fills the page's `localStorage` before the application reads
 * it, and the page's own writes of a shared key are posted, seen through `Storage.prototype`; the host's writes go round
 * that, so they are not posted back.
 */
export function vscodeHost(api: VsCodeApi, win: Window = window): VsCodeHost {
  const bootData = readHostBoot(win.document);
  const boot0 = (): HostBoot | undefined => bootData;
  let page: HostPage | undefined;
  let readySent = false;
  const queue: ToPage[] = [];
  /** The page's local edits, counted, and the count at each export the host has not yet heard written: an edit between is a dirty document. */
  let edits = 0;
  const exportedAt = new Map<number, number>();
  let lastChanged = -Infinity;
  let changedTimer: ReturnType<typeof setTimeout> | undefined;
  let companion: string | undefined;
  const post = (message: FromPage): void => api.postMessage(message);
  /** The answers waited for, by request id: the companion's (ED-014 section 4.5) and the Ask box's (section 4.6). */
  let nextId = 1;
  type Answer = Extract<ToPage, { type: 'companion' | 'invited' | 'asked' | 'askModels' }>;
  const answers = new Map<number, { resolve: (m: Answer) => void }>();
  const request = (make: (id: number) => FromPage, onId?: (id: number) => void): Promise<Answer> => {
    const id = nextId++;
    onId?.(id);
    return new Promise((resolve) => { answers.set(id, { resolve }); post(make(id)); });
  };
  // ---- The Ask box (ED-014 section 4.6): the key's state from the boot and its pushes, VS Code's models as last heard.
  let keyStored = boot0()?.askKey === true;
  const keyListeners = new Set<(stored: boolean) => void>();
  let models: HostModelData[] | undefined;
  let modelsAsked: Promise<HostModelData[]> | undefined;
  const modelListeners = new Set<(models: HostModelData[]) => void>();
  const heardModels = (list: HostModelData[]): void => {
    models = list;
    for (const listener of modelListeners) listener(list);
  };
  /** The relay the session's room is on, as last posted (`room`). */
  let roomRelay: string | undefined;

  const postChanged = (): void => { lastChanged = Date.now(); post({ type: 'changed' }); };
  /** At most one `changed` per interval: the first at once, the rest as one at the interval's end. */
  const changedSoon = (): void => {
    if (changedTimer !== undefined) return;
    const wait = lastChanged + CHANGED_INTERVAL - Date.now();
    if (wait <= 0) { postChanged(); return; }
    changedTimer = setTimeout(() => { changedTimer = undefined; postChanged(); }, wait);
  };

  // ---- The options (P19).
  let storage: Storage | undefined;
  try { storage = win.localStorage; } catch { storage = undefined; }
  const proto = (win as unknown as { Storage?: { prototype: Storage } }).Storage?.prototype;
  const setItem = proto?.setItem;
  const removeItem = proto?.removeItem;
  /** Writes of the host's own (the boot fill, another editor's change), round the wrapper. */
  const hostSet = (key: string, value: string): void => { try { if (storage && setItem) setItem.call(storage, key, value); } catch { /* a full or refused storage keeps what it had */ } };
  const hostRemove = (key: string): void => { try { if (storage && removeItem) removeItem.call(storage, key); } catch { /* ignore */ } };
  const boot = boot0();
  // A key typed into the page's Options before S6 (S1 to S5's builds): the page never holds one now (P45). Unshared, so
  // nothing is posted.
  hostRemove('docx4j-editor.ask.key');
  if (boot && storage) {
    // The host's copy is the truth: a shared key it does not hold was removed in another editor.
    for (let i = storage.length - 1; i >= 0; i--) {
      const key = storage.key(i);
      if (key !== null && isSharedOption(key) && !(key in boot.options)) hostRemove(key);
    }
    for (const [key, value] of Object.entries(boot.options)) if (isSharedOption(key)) hostSet(key, value);
  }
  let pending: { set: Record<string, string>; remove: Set<string> } | undefined;
  const flushOptions = (): void => {
    const out = pending;
    pending = undefined;
    if (!out) return;
    const set = Object.keys(out.set).length > 0 ? { set: out.set } : {};
    const remove = out.remove.size > 0 ? { remove: [...out.remove] } : {};
    post({ type: 'options', ...set, ...remove });
  };
  const queueOption = (key: string, value: string | null): void => {
    if (!pending) { pending = { set: {}, remove: new Set() }; queueMicrotask(flushOptions); }
    if (value === null) { delete pending.set[key]; pending.remove.add(key); } else { pending.remove.delete(key); pending.set[key] = value; }
  };
  if (proto && setItem && removeItem) {
    proto.setItem = function (this: Storage, key: string, value: string): void {
      const k = String(key);
      const before = this === storage && isSharedOption(k) ? this.getItem(k) : undefined;
      setItem.call(this, key, value);
      if (before !== undefined && before !== String(value)) queueOption(k, String(value));
    };
    proto.removeItem = function (this: Storage, key: string): void {
      const k = String(key);
      const before = this === storage && isSharedOption(k) ? this.getItem(k) : null;
      removeItem.call(this, key);
      if (before !== null) queueOption(k, null);
    };
  }
  const applyOptions = (message: OptionsMessage): void => {
    for (const [key, value] of Object.entries(message.set ?? {})) if (isSharedOption(key) && typeof value === 'string') hostSet(key, value);
    for (const key of message.remove ?? []) if (isSharedOption(key)) hostRemove(key);
  };

  // ---- The messages.
  const save = async (to: HostPage, id: number, name: string): Promise<void> => {
    const at = edits;
    try {
      const out = await to.save(name);
      if ('bytes' in out) { exportedAt.set(id, at); post({ type: 'saved', id, bytes: out.bytes }); }
      else post({ type: 'saved', id, refused: out.refused });
    } catch (e) {
      post({ type: 'saved', id, refused: (e as Error).message ?? String(e) });
    }
  };
  /**
   * The messages in the order they came. An open and a revert hold the ones after them until the document is open: a save,
   * a change on disk or a hand-over that comes while the page is still opening waits for it (ED-014 section 4.3, "requests
   * wait until the page has opened its document"). Anything else is handled at once. A backup never waits, since VS Code
   * asks for one at any time and the page answers from what it holds.
   */
  const jobs: (() => Promise<unknown> | void)[] = [];
  let held = false;
  const pump = (): void => {
    while (!held && jobs.length > 0) {
      let out: Promise<unknown> | void;
      try { out = jobs.shift()!(); } catch { out = undefined; }
      if (out) {
        held = true;
        out.then(() => undefined, () => undefined).then(() => { held = false; pump(); });
      }
    }
  };
  const inOrder = (job: () => Promise<unknown> | void): void => { jobs.push(job); pump(); };
  const deliver = (to: HostPage, m: ToPage): void => {
    if (m.type === 'backup') {
      const backup = to.backup();
      post(backup ? { type: 'backup', id: m.id, update: backup.update, meta: backup.meta } : { type: 'backup', id: m.id, empty: true });
      return;
    }
    if (m.type === 'options') { applyOptions(m); return; }
    inOrder(() => handle(to, m));
  };
  const handle = (to: HostPage, m: Exclude<ToPage, { type: 'backup' } | OptionsMessage>): Promise<unknown> | void => {
    switch (m.type) {
      case 'open': {
        const how: HostOpenHow = {
          ...('resume' in m && m.resume ? { resume: { update: asBytes(m.resume.update), meta: m.resume.meta as BackupMeta } } : {}),
          ...('place' in m && m.place !== undefined ? { place: m.place } : {}),
          ...('clean' in m && m.clean === true && 'resume' in m && m.resume ? { clean: true } : {}),
        };
        // A failure is the page's to say: it shows it and posts `status`.
        return 'join' in m
          ? to.open({ name: m.name, join: m.join })
          : to.open({ name: m.name, bytes: asBytes(m.bytes), lastModified: Date.now() }, Object.keys(how).length > 0 ? how : undefined);
      }
      case 'save': void save(to, m.id, m.name); break;
      case 'written': {
        // Clean only if no edit came after this save's export (section 4.1); else the page stays unsaved and says so again.
        const clean = exportedAt.get(m.id) === edits;
        exportedAt.delete(m.id);
        to.written(m.name, clean);
        // A `changed` the interval was holding is spent either way: its edits came before the export. Left to fire after a
        // clean write, it made VS Code's document dirty while the page called it saved (Fable's review, section 9.4 item 1).
        clearTimeout(changedTimer);
        changedTimer = undefined;
        if (!clean) postChanged();
        break;
      }
      case 'writeFailed': exportedAt.delete(m.id); to.writeFailed(m.error); break;
      // Nothing was written: the export is forgotten, and the page says nothing (section 9.8 finding 7).
      case 'saveCancelled': exportedAt.delete(m.id); break;
      case 'handoff': {
        // The store goes with the place if an edit came after this save's export (P25, section 9.8 finding 1), or if the session
        // is in a room, which the target's editor rejoins (P42). An export the host never heard of (a rename's) is `late`: a
        // resume of an unedited store is only a dirty document, and the extension decides a rename's clean state itself.
        const late = exportedAt.get(m.id) !== edits;
        exportedAt.delete(m.id);
        // The edits since go with the store, or there were none: no `changed` for an editor that is closing.
        clearTimeout(changedTimer);
        changedTimer = undefined;
        const { place } = to.handoff();
        const resume = late || roomRelay !== undefined ? to.backup() : undefined;
        post({ type: 'handoff', id: m.id, ...(place !== undefined ? { place } : {}), ...(resume ? { resume: { update: resume.update, meta: resume.meta } } : {}), late });
        break;
      }
      case 'revert': return to.revert(asBytes(m.bytes));
      case 'disk': to.diskChanged(asBytes(m.bytes), m.lastModified); break;
      case 'companion': if ('url' in m) to.setCompanion(m.url, m.relay); break;
      case 'invited': break;
      case 'palette': to.showPalette(); break;
      case 'pasteText': to.pasteText(m.text); break;
      // The Ask box's pushes: they wait behind an open as the other messages do (section 4.6).
      case 'askModels': heardModels(Array.isArray(m.models) ? m.models : []); break;
      case 'askKeyState':
        keyStored = m.stored === true;
        for (const listener of keyListeners) listener(keyStored);
        break;
      case 'asked': break;
    }
  };
  const onMessage = (event: MessageEvent): void => {
    const m = event.data as ToPage | null;
    if (!m || typeof m !== 'object' || typeof (m as { type?: unknown }).type !== 'string') return;
    // The options are the storage's, applied whether or not the application has attached.
    if (m.type === 'options') { applyOptions(m); return; }
    // An answer to the page's own request goes to its caller at once, never behind an open in the queue.
    if ((m.type === 'companion' || m.type === 'invited' || m.type === 'asked' || m.type === 'askModels') && typeof m.id === 'number' && answers.has(m.id)) {
      const waiting = answers.get(m.id)!;
      answers.delete(m.id);
      if (m.type === 'companion' && 'url' in m) companion = m.url;
      waiting.resolve(m);
      if (m.type !== 'companion' || !('url' in m)) return;
    } else if (m.type === 'companion' && 'url' in m) companion = m.url;
    // An answer nobody waits for (a Cancel deleted its entry): dropped.
    else if ((m.type === 'asked' || m.type === 'invited') || (m.type === 'askModels' && typeof m.id === 'number')) return;
    // In the order they came: one that arrives while the queue waits for its drain goes behind it.
    if (page && queue.length === 0) deliver(page, m); else queue.push(m);
  };
  win.addEventListener('message', onMessage);

  const askHost: HostAsk = {
    keyStored: () => keyStored,
    onKeyStored(listener) { keyListeners.add(listener); return () => { keyListeners.delete(listener); }; },
    models(fresh = false) {
      if (!fresh && models) return Promise.resolve(models);
      if (!fresh && modelsAsked) return modelsAsked;
      const asked = request((id) => ({ type: 'askModels', id })).then((out) => {
        const list = out.type === 'askModels' && Array.isArray(out.models) ? out.models : [];
        heardModels(list);
        return list;
      });
      modelsAsked = asked;
      void asked.finally(() => { if (modelsAsked === asked) modelsAsked = undefined; });
      return asked;
    },
    onModels(listener) { modelListeners.add(listener); return () => { modelListeners.delete(listener); }; },
    send(route, ask, signal) {
      return new Promise<AskAnswerData>((resolve, reject) => {
        const abort = (): DOMException => new DOMException('The request was cancelled', 'AbortError');
        if (signal.aborted) { reject(abort()); return; }
        let askId = 0;
        // The Cancel deletes the entry before it posts `askCancel`, so an `asked` that crosses it is dropped.
        const onAbort = (): void => {
          if (answers.delete(askId)) post({ type: 'askCancel', id: askId });
          reject(abort());
        };
        signal.addEventListener('abort', onAbort, { once: true });
        void request((id) => ({ type: 'ask', id, route, request: ask }), (id) => { askId = id; }).then((out) => {
          signal.removeEventListener('abort', onAbort);
          if (out.type !== 'asked') { reject(new AskFailure({ kind: 'other', message: 'no answer' })); return; }
          if ('answer' in out) resolve(out.answer);
          else if ('failure' in out) reject(new AskFailure(out.failure));
          else reject(abort());
        });
      });
    },
    key(action) { post({ type: 'askKey', action }); },
  };

  return {
    kind: 'vscode',
    // Document macros only in a trusted workspace, as the boot says when the page was written (P27); the page's own plain
    // paste, since the host page cancels Ctrl+Shift+V (P34).
    has: (feature) => feature === 'clipboardText' || (feature === 'documentMacros' && boot?.trusted === true),
    attach(p) {
      page = p;
      if (!readySent) { readySent = true; post({ type: 'ready', origin: win.location.origin }); }
      // After the attach's commit: React's StrictMode detaches and attaches again at once, and the queue goes to the page attached last.
      queueMicrotask(() => { if (page !== p) return; for (const m of queue.splice(0)) deliver(p, m); });
      return () => { if (page === p) page = undefined; };
    },
    changed() { edits++; changedSoon(); },
    async openDocument(file) {
      post('join' in file ? { type: 'untitled', name: file.name, join: file.join } : { type: 'untitled', name: file.name, bytes: file.bytes });
    },
    reload() { post({ type: 'revertRequest' }); },
    openExternal(url) { post({ type: 'openExternal', url }); },
    download(bytes, name) { post({ type: 'download', bytes, name }); },
    siteAddress() {
      // The companion serves its own copy of the editor at its root (S5); a `vscode-webview://` address is no use to a peer.
      if (companion) { try { return new URL('/', companion).href; } catch { /* below */ } }
      return pageAddress();
    },
    base(bytes) { post({ type: 'base', bytes }); },
    status(text) { post({ type: 'status', text }); },
    language: () => boot?.language,
    commandName: (id) => boot?.commands?.[id],
    companion() {
      const agent = boot?.agent;
      if (agent === undefined) return undefined;
      return {
        agent,
        async start() {
          const out = await request((id) => ({ type: 'companionStart', id }));
          if (out.type !== 'companion' || !('url' in out)) throw new Error(out.type === 'companion' && 'error' in out ? out.error : 'the companion did not start');
          return { url: out.url, relay: out.relay };
        },
        async invite(link) {
          const out = await request((id) => ({ type: 'invite', id, room: link.room, token: link.token }));
          if (out.type === 'invited' && out.error !== undefined) throw new Error(out.error);
        },
      };
    },
    ask: () => askHost,
    room(relay) {
      if (relay === roomRelay) return;
      roomRelay = relay;
      post(relay !== undefined ? { type: 'room', relay } : { type: 'room' });
    },
    stop() {
      win.removeEventListener('message', onMessage);
      clearTimeout(changedTimer);
      if (proto && setItem && removeItem) { proto.setItem = setItem; proto.removeItem = removeItem; }
    },
  };
}

let chosen: Host | undefined;
let initialised = false;
/** The key gate's removal, while VS Code's host has one installed. */
let removeGate: (() => void) | undefined;

/**
 * Chosen once, in `main.tsx`, before the language loads: the VS Code host when `acquireVsCodeApi` exists (it may be acquired
 * once per page), else the browser's. A test passes its fake. VS Code's host comes with the key gate on the page's
 * `document` (`keygate.ts`, ED-014 section 4.4); the browser's has none.
 */
export function initHost(api?: VsCodeApi): Host {
  if (initialised && chosen && api === undefined) return chosen;
  const acquire = (globalThis as { acquireVsCodeApi?: () => VsCodeApi }).acquireVsCodeApi;
  const given = api ?? (typeof acquire === 'function' ? acquire() : undefined);
  (chosen as Partial<VsCodeHost> | undefined)?.stop?.();
  removeGate?.();
  removeGate = undefined;
  chosen = given ? vscodeHost(given) : browserHost();
  if (given && typeof document !== 'undefined') removeGate = installKeyGate(document);
  initialised = true;
  return chosen;
}

/** The host chosen; read at render and in handlers, never at module load. A page that never chose (a jsdom test) is a browser's. */
export function host(): Host {
  chosen ??= browserHost();
  return chosen;
}

/** For tests: the host put away, a VS Code host's storage wrapper and key gate with it. */
export function resetHost(): void {
  (chosen as Partial<VsCodeHost> | undefined)?.stop?.();
  removeGate?.();
  removeGate = undefined;
  chosen = undefined;
  initialised = false;
}
