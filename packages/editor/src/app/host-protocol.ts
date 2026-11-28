// SPDX-License-Identifier: GPL-3.0-or-later
//
// The messages between the VS Code extension and the application in its webview (ED-014 section 4.1: the page side of
// section 3.2's protocol, as P18 revises it). No imports, so the extension (`packages/editor-vscode`, S3) shares the file
// as it is. Byte arrays travel as plain `Uint8Array`s, never a Node `Buffer`, which VS Code's serializer turns into a
// plain object (S0, section 9.1 item 10).

/** The meta the extension writes into the page it serves: boot data, read before the language loads (P19). */
export const HOST_META = 'docx4j-host';

/** The boot data, the meta's content as JSON. */
export interface HostBoot {
  /** The options the host keeps for every editor of the extension: each shared key with its value (P19). */
  options: Record<string, string>;
  /** VS Code's display language (`env.language`): the page's, unless the person chose one in Options (P27). */
  language?: string;
  /** Whether the workspace is trusted when the page is written: document macros are offered only then (P27). */
  trusted?: boolean;
  /**
   * The agent's name, where the extension can start the companion for this page (ED-014 section 4.5, P38): Invite is
   * labelled with it before the companion runs. Absent where it cannot run (Restricted Mode, a remote or virtual workspace).
   */
  agent?: string;
  /**
   * Whether an Anthropic API key is stored in VS Code's SecretStorage (ED-014 section 4.6, P45): the key itself never
   * enters the page. Kept current by `askKeyState`.
   */
  askKey?: boolean;
  /**
   * The names VS Code's palette shows for the extension's commands the page names (ED-014 section 4.7, P49), by id:
   * `<category>: <title>`, as VS Code translated them for its own language, which may not be the page's.
   */
  commands?: Record<string, string>;
}

/**
 * The extension's commands the page names, with their names in English: the page's words where the boot names none
 * (`test/unit/names.test.ts` holds them to the manifest).
 */
export const PAGE_COMMANDS: Readonly<Record<string, string>> = {
  'docx4j.ask.setKey': 'docx4j editor: Set the Anthropic API Key…',
};

/** The Ask box's request (ED-003 section 7.5): shown before it is sent, and sent as it was shown (ED-014 section 4.6). */
export interface AskRequestData {
  /** An Anthropic model's id, or, on VS Code's route, a VS Code model's key, `<vendor>/<id>`. */
  model: string;
  system: string;
  prompt: string;
  maxTokens: number;
}

/** An answer: its text, the model that gave it, and what the provider says beside it. */
export interface AskAnswerData {
  text: string;
  model: string;
  stopReason: string | null;
  /** The refusal's category and explanation when the model declined (`stop_reason: 'refusal'`). */
  refusal?: { category: string | null; explanation: string | null };
  usage?: { input: number; output: number };
}

/**
 * Why a request has no answer, worded by the page (section 4.6). The key's route: `key` (401), `rateLimit` (429),
 * `connection`, `api`. VS Code's route: `noPermissions` (the person did not allow it), `blocked` (the provider's limit),
 * `notFound` (the model is gone), `tooLong` (over the model's input, with both figures). Either: `noKey`, `other`.
 */
export interface AskFailureData {
  kind: 'key' | 'rateLimit' | 'connection' | 'api' | 'noPermissions' | 'blocked' | 'notFound' | 'tooLong' | 'noKey' | 'other';
  message: string;
  status?: number;
  /** `tooLong`: the request's tokens and the model's limit. */
  tokens?: number;
  limit?: number;
}

/** A model VS Code offers (`vscode.lm`), known by its vendor and its id, since an id is unique only within its vendor (section 9.20). */
export interface HostModelData {
  vendor: string;
  id: string;
  version: string;
  name: string;
  family: string;
  maxInputTokens: number;
}

/** A VS Code model's key, as `ask.model` keeps it after `vscode:` and the `ask` message names it. */
export function modelKey(model: Pick<HostModelData, 'vendor' | 'id'>): string {
  return `${model.vendor}/${model.id}`;
}

/** A room, as a link names it (`RoomLink` in the application). */
export interface RoomLinkData {
  room: string;
  token: string;
  relay: string;
}

const ROOM_ID = /^[A-Za-z0-9_-]{16,128}$/;

/**
 * The room a link or a location's fragment names, when it names one well formed: `<site>#room=<id>&token=<token>&relay=<ws
 * address>`. Here, with no imports, so that the extension's Join a Room… checks a link as the page's Join does (ED-014
 * section 4.5); `room.ts` re-exports it.
 */
export function parseRoomLink(text: string): RoomLinkData | undefined {
  const at = text.indexOf('#');
  const params = new URLSearchParams(at >= 0 ? text.slice(at + 1) : text);
  const room = params.get('room');
  const token = params.get('token');
  const relay = params.get('relay');
  if (!room || !token || !relay || !ROOM_ID.test(room)) return undefined;
  try {
    const url = new URL(relay);
    if (url.protocol !== 'ws:' && url.protocol !== 'wss:') return undefined;
  } catch {
    return undefined;
  }
  return { room, token, relay };
}

/** What a backup holds beside the store's update. */
export interface BackupMetaData {
  name: string;
  /** The Markdown file the document was made from. */
  source?: string;
  room?: RoomLinkData;
  joined?: true;
  /** What first made the document unsaved (ED-011 section 9.9): opaque to the extension, kept and given back. */
  firstEdit?: unknown;
}

/** A backup: `Y.encodeStateAsUpdate` of the store, and its meta. The extension writes both with the document's base bytes. */
export interface BackupData {
  update: Uint8Array;
  meta: BackupMetaData;
}

/** From the extension to the page. A message that comes before the application attaches waits in a queue. */
export type ToPage =
  /**
   * Open this document. Empty bytes: a new one, its page size asked as New asks. A Markdown name is converted. With `resume`,
   * a hot-exit backup over the bytes. With `place`, the caret and the view a Save As's hand-over carried (P25), restored once
   * open: opaque to the extension, the page's `KeptPlace`. With `clean`, the resume opens as saved (P42): a Save As in a room
   * with no edit after the export, or a rename of a clean document in a room, which the extension decides.
   */
  | { type: 'open'; name: string; bytes: Uint8Array; resume?: BackupData; place?: unknown; clean?: boolean }
  /** Join this room, in an untitled editor (P17). */
  | { type: 'open'; name: string; join: RoomLinkData }
  /** VS Code's Save asked for the self-checked bytes, in the form the name asks for (Flat OPC for `.xml`). */
  | { type: 'save'; id: number; name: string }
  /** The save's second phase, after `saveCustomDocument` has resolved: the document is saved under this name. */
  | { type: 'written'; id: number; name: string }
  | { type: 'writeFailed'; id: number; error: string }
  /**
   * The save was cancelled (a later save came, or VS Code stopped waiting): nothing was written, and the page forgets the
   * export, saying nothing (section 9.8 finding 7).
   */
  | { type: 'saveCancelled'; id: number }
  /**
   * A Save As wrote its target, which VS Code opens afresh in a new editor (P25): the page stops taking input and answers
   * with its place, and its store if an edit came after the export.
   */
  | { type: 'handoff'; id: number }
  /** Hot exit asks for the session's state, about a second after each change (S0). */
  | { type: 'backup'; id: number }
  /** Open these bytes again: the file's (or a Markdown source's), the caret's place kept. */
  | { type: 'revert'; bytes: Uint8Array }
  /** The file changed on disk (ED-011's banner); for a document made from Markdown, the `.md`'s bytes. */
  | { type: 'disk'; bytes: Uint8Array; lastModified: number }
  /**
   * The companion's address and its relay's (ED-014 section 4.5, P38): to every page when the companion starts, and to a page
   * at its `ready` while it runs. With `id`, the answer to `companionStart`.
   */
  | { type: 'companion'; id?: number; url: string; relay: string }
  /** Why the companion cannot run for this page, the answer to `companionStart`: words to show. */
  | { type: 'companion'; id: number; error: string }
  /** The answer to `invite`: the agent joined the room, or why not. */
  | { type: 'invited'; id: number; error?: string }
  /** "docx4j editor: Commands": the application's palette (section 3.4). */
  | { type: 'palette' }
  /**
   * "Paste as Plain Text" from VS Code's menus (ED-014 section 4.4, P34): the clipboard's text, read by the extension, pasted
   * where the page's focus is. The page's own Ctrl+Shift+V reads the clipboard itself.
   */
  | { type: 'pasteText'; text: string }
  /**
   * VS Code's models (ED-014 section 4.6, P46). With `id`, the answer to the page's `askModels`; without, the list again
   * to every page when VS Code's models change.
   */
  | { type: 'askModels'; id?: number; models: HostModelData[] }
  /** The answer to `ask`, or why there is none, or that it was cancelled. */
  | { type: 'asked'; id: number; answer: AskAnswerData }
  | { type: 'asked'; id: number; failure: AskFailureData }
  | { type: 'asked'; id: number; cancelled: true }
  /** The Anthropic key was set or removed, by a command in any window: to every page (P45). */
  | { type: 'askKeyState'; stored: boolean }
  /** Another editor's option change, passed on (P19). */
  | OptionsMessage;

/** From the page to the extension. */
export type FromPage =
  /** The application attached; once per page. `origin` is the page's, for the companion's allowance (section 3.6). */
  | { type: 'ready'; origin: string }
  /** The bytes a backup is resumed over, where they are not the bytes the extension posted: a new document, a Markdown file's conversion, a room's base file. */
  | { type: 'base'; bytes: Uint8Array }
  /** A local edit, at most one per `CHANGED_INTERVAL`; posted again after `written` when an edit came after the export. */
  | { type: 'changed' }
  | { type: 'saved'; id: number; bytes: Uint8Array }
  /** The self-check (or the export) refused: the page shows why, and the extension fails VS Code's save. */
  | { type: 'saved'; id: number; refused: string }
  | { type: 'backup'; id: number; update: Uint8Array; meta: BackupMetaData }
  /**
   * No document is open, nor one being resumed from a backup: the page holds nothing beyond the bytes. While a backup is being
   * resumed (its room waited for, say), the page answers with that backup, so a backup asked then never loses its edits.
   */
  | { type: 'backup'; id: number; empty: true }
  /**
   * The answer to `handoff` (P25): the page's place (its `KeptPlace`, opaque to the extension), and, if an edit came after the
   * export or the session is in a room (P42), the store as a backup's update and meta, which the target's editor resumes over
   * the old document's base. `late`: an edit came after the export (always, for an id no export was made under).
   */
  | { type: 'handoff'; id: number; place?: unknown; resume?: BackupData; late?: boolean }
  /** The banner's Reload: the extension reverts, through `workbench.action.files.revert` when the document is dirty. */
  | { type: 'revertRequest' }
  /** Open a new editor with this document, or this room (P17). */
  | { type: 'untitled'; name: string; bytes: Uint8Array }
  | { type: 'untitled'; name: string; join: RoomLinkData }
  | { type: 'openExternal'; url: string }
  /** A download, through a save dialog (S0: a webview's own waits for Electron's and writes nothing). */
  | { type: 'download'; bytes: Uint8Array; name: string }
  /** Something the page said that the extension says too: an open that failed, or fell back. */
  | { type: 'status'; text: string }
  /** Invite and Share ask for the companion (ED-014 section 4.5): started if need be, the page's origin allowed. */
  | { type: 'companionStart'; id: number }
  /** The agent invited into a room on the companion's relay, which the page has shared. */
  | { type: 'invite'; id: number; room: string; token: string }
  /** The session entered a room on this relay; with none, it left its room. */
  | { type: 'room'; relay?: string }
  /** The Ask box shows for the first time, or its model is gone: VS Code's models, please (P46). */
  | { type: 'askModels'; id: number }
  /**
   * Send this request (P44): `vscode` through a VS Code model, the request's model its key (`<vendor>/<id>`); `key` to the
   * Anthropic API with the key the extension keeps, the request's model an Anthropic model's id.
   */
  | { type: 'ask'; id: number; route: 'vscode' | 'key'; request: AskRequestData }
  /** The person's Cancel. The extension ignores an id it does not know. */
  | { type: 'askCancel'; id: number }
  /** The Options dialog's buttons (P45): the key's commands, and the address's setting. */
  | { type: 'askKey'; action: 'set' | 'remove' | 'settings' }
  | OptionsMessage;

/** An option changed: posted by the page, and passed on by the extension to its other editors (P19). */
export interface OptionsMessage {
  type: 'options';
  set?: Record<string, string>;
  remove?: string[];
}

/** At most one `changed` per this many milliseconds (section 4.1). */
export const CHANGED_INTERVAL = 250;

/** The application's options, in `localStorage` under this prefix. */
export const OPTION_PREFIX = 'docx4j-editor.';

/** The keys kept per page, never shared (P19): recovery's pointer, a log, and the Ask box's key, which S6 keeps in VS Code's SecretStorage. */
export const UNSHARED_OPTIONS: readonly string[] = ['docx4j-editor.session', 'docx4j-editor.editCauses', 'docx4j-editor.ask.key'];

/** Whether an option is kept by the host for every editor of the extension. */
export function isSharedOption(key: string): boolean {
  return key.startsWith(OPTION_PREFIX) && !UNSHARED_OPTIONS.includes(key);
}
