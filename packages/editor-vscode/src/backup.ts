// SPDX-License-Identifier: GPL-3.0-or-later
//
// The hot-exit backup's file (ED-014 section 4.3), with no `vscode` import:
//   1. the 8 bytes `D4JBKUP1`;
//   2. the header's length, 4 bytes, big-endian;
//   3. the header, JSON in UTF-8: `{ uri, name, source?, meta, baseLength, updateLength, diskLength?, sourceLength? }`;
//   4. `base`, the bytes the page's store was opened over;
//   5. the update, `Y.encodeStateAsUpdate` of the store; empty for a backup of `base` alone;
//   6. `disk`, the document's file as last read or written, only where it differs from `base` (a save in the session);
//   7. the Markdown source's bytes as last read, for a document made from one.
// A resume compares the file with `disk` and the source with its bytes, so a change made while VS Code was closed shows
// the banner, and the session's own save does not (Fable's review, ED-014 section 9.11 finding 1).
// VS Code gives each backup a new destination, so a file is always written whole, never over the last.
import type { BackupMetaData } from './shared.js';

const MAGIC = new TextEncoder().encode('D4JBKUP1');

export interface BackupFile {
  /** The document's URI when the backup was written: a check only, since after a rename VS Code hands the backup to the new URI. */
  uri: string;
  name: string;
  /** The Markdown file the document was made from, watched again after a restart (P26). */
  source?: string;
  meta: BackupMetaData;
  base: Uint8Array;
  update: Uint8Array;
  /** The file as last read or written, where it is not `base`. */
  disk?: Uint8Array;
  /** The Markdown source as last read. */
  sourceBytes?: Uint8Array;
}

/** A file this format cannot read. */
export class BackupFormatError extends Error {
  override name = 'BackupFormatError';
}

export function encodeBackup(backup: BackupFile): Uint8Array {
  const header = new TextEncoder().encode(JSON.stringify({
    uri: backup.uri,
    name: backup.name,
    ...(backup.source !== undefined ? { source: backup.source } : {}),
    meta: backup.meta,
    baseLength: backup.base.byteLength,
    updateLength: backup.update.byteLength,
    ...(backup.disk ? { diskLength: backup.disk.byteLength } : {}),
    ...(backup.sourceBytes ? { sourceLength: backup.sourceBytes.byteLength } : {}),
  }));
  const extra = (backup.disk?.byteLength ?? 0) + (backup.sourceBytes?.byteLength ?? 0);
  const out = new Uint8Array(MAGIC.length + 4 + header.length + backup.base.byteLength + backup.update.byteLength + extra);
  out.set(MAGIC, 0);
  new DataView(out.buffer).setUint32(MAGIC.length, header.length);
  let at = MAGIC.length + 4;
  out.set(header, at);
  at += header.length;
  out.set(backup.base, at);
  at += backup.base.byteLength;
  out.set(backup.update, at);
  at += backup.update.byteLength;
  if (backup.disk) { out.set(backup.disk, at); at += backup.disk.byteLength; }
  if (backup.sourceBytes) out.set(backup.sourceBytes, at);
  return out;
}

export function decodeBackup(bytes: Uint8Array): BackupFile {
  if (bytes.byteLength < MAGIC.length + 4 || MAGIC.some((b, i) => bytes[i] !== b)) throw new BackupFormatError('not a docx4j editor backup');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const headerLength = view.getUint32(MAGIC.length);
  let at = MAGIC.length + 4;
  if (at + headerLength > bytes.byteLength) throw new BackupFormatError('the backup is cut short in its header');
  let header: { uri?: unknown; name?: unknown; source?: unknown; meta?: unknown; baseLength?: unknown; updateLength?: unknown; diskLength?: unknown; sourceLength?: unknown };
  try {
    header = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes.subarray(at, at + headerLength))) as typeof header;
  } catch {
    throw new BackupFormatError('the backup\'s header does not read');
  }
  at += headerLength;
  const { uri, name, source, meta, baseLength, updateLength } = header;
  const diskLength = header.diskLength ?? 0;
  const sourceLength = header.sourceLength ?? 0;
  if (typeof uri !== 'string' || typeof name !== 'string' || (source !== undefined && typeof source !== 'string')) throw new BackupFormatError('the backup\'s header is not a backup\'s');
  if (!meta || typeof meta !== 'object' || typeof (meta as { name?: unknown }).name !== 'string') throw new BackupFormatError('the backup\'s meta is not a backup\'s');
  const lengths = [baseLength, updateLength, diskLength, sourceLength];
  if (lengths.some((n) => !Number.isSafeInteger(n) || (n as number) < 0)) throw new BackupFormatError('the backup\'s lengths are not lengths');
  const [b, u, d, so] = lengths as number[];
  if (at + b! + u! + d! + so! !== bytes.byteLength) throw new BackupFormatError('the backup is cut short, or has more than its header says');
  // Copies of their own, so the file's buffer can go.
  const take = (n: number): Uint8Array => { const out = bytes.slice(at, at + n); at += n; return out; };
  const base = take(b!);
  const update = take(u!);
  const disk = header.diskLength !== undefined ? take(d!) : undefined;
  const sourceBytes = header.sourceLength !== undefined ? take(so!) : undefined;
  return {
    uri, name,
    ...(source !== undefined ? { source: source as string } : {}),
    meta: meta as BackupMetaData,
    base, update,
    ...(disk ? { disk } : {}),
    ...(sourceBytes ? { sourceBytes } : {}),
  };
}
