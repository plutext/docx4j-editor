// SPDX-License-Identifier: GPL-3.0-or-later
//
// Files (ED-001 section 11 item 1; ED-002 section 5.4): open by picker or drop, Download always,
// Save As through the File System Access API where the browser has it, Save in place through
// the handle after the self-check and a modification-time check. Save in place is behind a
// flag until the Word 2010 check of the prefixed app.xml has passed (ED-002 section 10.1 item 9). A file named `.xml` is
// written as Flat OPC, Word's "Word XML Document" (ED-005 section 12.122): Save As offers it, and Save keeps to it.
import { EditorDocument, SelfCheckError, type SaveFormat, type SaveResult } from '@docx4j/editor-model';
import { t } from './i18n.js';

/**
 * A failed save in the engine's words, each cause after what it caused: its outer error says only which part
 * failed, and why is on the cause (ED-005 section 12.114).
 */
export function reasonOf(e: unknown): string {
  const words: string[] = [];
  for (let error = e as Error | undefined; error instanceof Error; error = error.cause as Error | undefined) words.push(error.message);
  return words.length > 0 ? words.join(': ') : String(e);
}

export const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
export const DOTX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.template';
export const DOCM_TYPE = 'application/vnd.ms-word.document.macroEnabled.12';
export const DOTM_TYPE = 'application/vnd.ms-word.template.macroEnabled.12';
/** Flat OPC's type: an XML document (Word's own "Word XML Document" is one). */
export const FLAT_OPC_TYPE = 'application/xml';
/** What Open accepts: documents and templates, macro-enabled or not (a template opened is offered as the basis of a new document, newdoc.ts), and Markdown, opened as a new document (ED-005 section 12.66). */
export const OPEN_ACCEPT = '.docx,.dotx,.docm,.dotm,.md,.markdown';

/**
 * The form a save writes, by the file's name (ED-005 section 12.122): `.xml` is Flat OPC, anything else the zip. The name the
 * person chose in Save As decides it, not the picker's type, so the file is what its name says; a later Save keeps to it.
 */
export function formatOfName(name: string): SaveFormat {
  return /\.xml$/i.test(name) ? 'flatOpc' : 'docx';
}

/** A document's name as Flat OPC: its extension `.xml`. */
export function flatOpcName(name: string): string {
  return `${name.replace(/\.(docx|docm|dotx|dotm|md|markdown|xml)$/i, '')}.xml`;
}

/** Whether a file is Markdown, by its name: it opens as a new document made from it (`documentFromMarkdown`). */
export function isMarkdownName(name: string): boolean {
  return /\.(md|markdown)$/i.test(name);
}

type FileHandle = FileSystemFileHandle;

interface PickerWindow {
  showOpenFilePicker?: (options?: unknown) => Promise<FileHandle[]>;
  showSaveFilePicker?: (options?: unknown) => Promise<FileHandle>;
}

/** Documents and templates, described in the current language: what Open reads and Save As writes. */
const documentTypes = (): unknown[] => [
  { description: t('file.picker.document'), accept: { [DOCX_TYPE]: ['.docx'], [DOCM_TYPE]: ['.docm'] } },
  { description: t('file.picker.template'), accept: { [DOTX_TYPE]: ['.dotx'], [DOTM_TYPE]: ['.dotm'] } },
];

/** Flat OPC, which Save As writes and Open does not read (ED-005 section 12.122). */
const flatOpcType = (): unknown => ({ description: t('file.picker.flatOpc'), accept: { [FLAT_OPC_TYPE]: ['.xml'] } });

/** What Save As writes: never Markdown, since a save is a document's package (ED-011 section 3.8); Flat OPC first when asked for. */
const saveTypes = (flatOpc: boolean): unknown[] => (flatOpc ? [flatOpcType(), ...documentTypes()] : [...documentTypes(), flatOpcType()]);

/** What Open reads: documents, templates and Markdown. */
const pickerTypes = (): unknown[] => [
  ...documentTypes(),
  { description: t('file.picker.markdown'), accept: { 'text/markdown': ['.md', '.markdown'] } },
];

export function hasFileSystemAccess(): boolean {
  const w = window as unknown as PickerWindow;
  return typeof w.showOpenFilePicker === 'function' && typeof w.showSaveFilePicker === 'function';
}

/** The Markdown file a document was made from (ED-011 section 3.8): read and watched, never written. */
export interface SourceFile {
  handle: FileHandle;
  name: string;
  /** The Markdown file's modification time when it was read, where the document's own is the conversion's. */
  lastModified: number;
}

export interface OpenedFile {
  name: string;
  bytes: Uint8Array;
  handle?: FileHandle;
  /** The file's modification time when it was opened, for the last-writer check. */
  lastModified: number;
  /** A document made from a Markdown file that came with a handle. `handle` stays unset, so nothing saves over it. */
  source?: SourceFile;
}

/** Open through the browser's picker (a handle, so Save in place is possible) or return null when the user cancelled. */
export async function pickFile(): Promise<OpenedFile | null> {
  const w = window as unknown as PickerWindow;
  if (!w.showOpenFilePicker) return null;
  try {
    const [handle] = await w.showOpenFilePicker({ types: pickerTypes(), multiple: false });
    if (!handle) return null;
    return await readHandle(handle);
  } catch (e) {
    if ((e as Error).name === 'AbortError') return null;
    throw e;
  }
}

/** The file behind a handle as it is now: its bytes and modification time (the picker's open; ED-011's Reload). */
export async function readHandle(handle: FileHandle): Promise<OpenedFile> {
  const file = await handle.getFile();
  return { name: file.name, bytes: new Uint8Array(await file.arrayBuffer()), handle, lastModified: file.lastModified };
}

/** A dropped file, or the file input's: its bytes, and the drop's handle where the browser gave one (`handleOfDrop`). */
export async function fromFile(file: File, handle?: FileHandle): Promise<OpenedFile> {
  return { name: file.name, bytes: new Uint8Array(await file.arrayBuffer()), lastModified: file.lastModified, ...(handle ? { handle } : {}) };
}

/**
 * The handle of a dropped file, where the browser gives one (Chromium's `getAsFileSystemHandle`, with read
 * permission from the drop itself; Firefox and Safari have none, and the file is then not watched; ED-011
 * section 3.1). Call it within the drop event: the items are gone after it.
 */
export function handleOfDrop(dt: DataTransfer): Promise<FileHandle | undefined> {
  const item = dt.items[0] as (DataTransferItem & { getAsFileSystemHandle?: () => Promise<FileSystemHandle | null> }) | undefined;
  if (!item || item.kind !== 'file' || typeof item.getAsFileSystemHandle !== 'function') return Promise.resolve(undefined);
  return item.getAsFileSystemHandle().then((h) => (h && h.kind === 'file' ? (h as FileHandle) : undefined), () => undefined);
}

export type SaveOutcome =
  | { kind: 'saved'; result: SaveResult; how: 'download' | 'handle' }
  | { kind: 'self-check-failed'; error: SelfCheckError; diagnosis: Uint8Array; format: SaveFormat }
  | { kind: 'changed-on-disk'; lastModified: number }
  | { kind: 'cancelled' };

/** The bytes of a download, typed by the name's form; never overwrites anything. */
export function download(bytes: Uint8Array, name: string): void {
  const blob = new Blob([bytes as BlobPart], { type: formatOfName(name) === 'flatOpc' ? FLAT_OPC_TYPE : DOCX_TYPE });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** Export in a form with the self-check; on a mismatch, the bytes for diagnosis come without it, in that form. */
async function exportChecked(session: EditorDocument, format: SaveFormat): Promise<{ result: SaveResult } | { error: SelfCheckError; diagnosis: Uint8Array; format: SaveFormat }> {
  try {
    return { result: await session.save({ format }) };
  } catch (e) {
    if (e instanceof SelfCheckError) return { error: e, diagnosis: (await session.save({ selfCheck: false, format })).bytes, format };
    throw e;
  }
}

/** A failed self-check as an outcome. */
const failed = (out: { error: SelfCheckError; diagnosis: Uint8Array; format: SaveFormat }): SaveOutcome => ({ kind: 'self-check-failed', error: out.error, diagnosis: out.diagnosis, format: out.format });

export async function saveAsDownload(session: EditorDocument, name: string): Promise<SaveOutcome> {
  const out = await exportChecked(session, formatOfName(name));
  if ('error' in out) return failed(out);
  download(out.result.bytes, name);
  return { kind: 'saved', result: out.result, how: 'download' };
}

/**
 * Save As through the picker, or a download without one; returns the new handle with the outcome. `flatOpc` suggests the
 * name as Flat OPC, with its type first (the palette's Save as Flat OPC XML). The export and its self-check run before the
 * picker, which may make the file; when the name chosen asks for the other form, they run again in that form.
 */
export async function saveAs(session: EditorDocument, name: string, options: { flatOpc?: boolean } = {}): Promise<SaveOutcome & { handle?: FileHandle; name?: string }> {
  const w = window as unknown as PickerWindow;
  const suggested = options.flatOpc ? flatOpcName(name) : name;
  if (!w.showSaveFilePicker) return saveAsDownload(session, suggested);
  let out = await exportChecked(session, formatOfName(suggested));
  if ('error' in out) return failed(out);
  let handle: FileHandle;
  try {
    handle = await w.showSaveFilePicker({ suggestedName: suggested, types: saveTypes(options.flatOpc === true) });
  } catch (e) {
    if ((e as Error).name === 'AbortError') return { kind: 'cancelled' };
    throw e;
  }
  if (formatOfName(handle.name) !== formatOfName(suggested)) {
    out = await exportChecked(session, formatOfName(handle.name));
    if ('error' in out) return failed(out);
  }
  await writeHandle(handle, out.result.bytes);
  return { kind: 'saved', result: out.result, how: 'handle', handle, name: handle.name };
}

/**
 * Save in place: refuse when the file changed on disk since it was opened (Word or someone else
 * wrote it), otherwise export with the self-check and write through the handle.
 */
export async function saveInPlace(session: EditorDocument, handle: FileHandle, openedAt: number): Promise<SaveOutcome & { lastModified?: number }> {
  // A handle another tab stored (ED-011 section 3.6, P5), or a drop's read-only one: the permission asked here,
  // from the click, where Chromium grants it; refused, the save is cancelled and Save As remains.
  if (!(await ensureWritable(handle))) return { kind: 'cancelled' };
  const current = await handle.getFile();
  if (current.lastModified !== openedAt) return { kind: 'changed-on-disk', lastModified: current.lastModified };
  const out = await exportChecked(session, formatOfName(handle.name));
  if ('error' in out) return failed(out);
  await writeHandle(handle, out.result.bytes);
  const after = await handle.getFile();
  return { kind: 'saved', result: out.result, how: 'handle', lastModified: after.lastModified };
}

type PermissionHandle = FileHandle & {
  queryPermission?: (d: { mode: 'read' | 'readwrite' }) => Promise<PermissionState>;
  requestPermission?: (d: { mode: 'read' | 'readwrite' }) => Promise<PermissionState>;
};

/** Read and write permission on the handle, asked when not already granted; true when it is. */
async function ensureWritable(handle: FileHandle): Promise<boolean> {
  const h = handle as PermissionHandle;
  if (!h.queryPermission || !h.requestPermission) return true;
  try {
    if ((await h.queryPermission({ mode: 'readwrite' })) === 'granted') return true;
    return (await h.requestPermission({ mode: 'readwrite' })) === 'granted';
  } catch { return false; }
}

async function writeHandle(handle: FileHandle, bytes: Uint8Array): Promise<void> {
  const writable = await handle.createWritable();
  await writable.write(bytes as BlobPart);
  await writable.close();
}

/** A name for the diagnosis download of a failed self-check, in the form the save was written in. */
export function diagnosisName(name: string, format: SaveFormat = 'docx'): string {
  return `${name.replace(/\.(docx|xml)$/i, '')}-selfcheck-failed.${format === 'flatOpc' ? 'xml' : 'docx'}`;
}
