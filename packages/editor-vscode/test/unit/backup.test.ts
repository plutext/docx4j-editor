// SPDX-License-Identifier: GPL-3.0-or-later
//
// The hot-exit backup's file (ED-014 section 4.3): written and read back, a file cut short refused, another format refused.
import { describe, expect, test } from 'vitest';
import { BackupFormatError, decodeBackup, encodeBackup, type BackupFile } from '../../src/backup.js';

const sample = (over: Partial<BackupFile> = {}): BackupFile => ({
  uri: 'file:///work/a.docx',
  name: 'a.docx',
  meta: { name: 'a.docx', firstEdit: { kind: 'typing', at: 3 } },
  base: new Uint8Array([80, 75, 3, 4, 1, 2, 3]),
  update: new Uint8Array([9, 8, 7]),
  ...over,
});

describe('the backup file', () => {
  test('written and read back, the source and the meta with it', () => {
    const backup = sample({ source: 'file:///work/notes.md', meta: { name: 'notes.docx', source: 'notes.md' } });
    const bytes = encodeBackup(backup);
    expect(new TextDecoder().decode(bytes.subarray(0, 8))).toBe('D4JBKUP1');
    const read = decodeBackup(bytes);
    expect(read).toEqual(backup);
    // The bytes are copies of their own, never views over the file's buffer.
    expect(read.base.buffer).not.toBe(bytes.buffer);
  });

  test('the file as last written and the Markdown source, where given, read back; a backup without them has none', () => {
    const backup = sample({ disk: new Uint8Array([80, 75, 9]), sourceBytes: new TextEncoder().encode('# Notes\n') });
    expect(decodeBackup(encodeBackup(backup))).toEqual(backup);
    const plain = decodeBackup(encodeBackup(sample()));
    expect(plain.disk).toBeUndefined();
    expect(plain.sourceBytes).toBeUndefined();
    // Cut inside the source's bytes: refused.
    const bytes = encodeBackup(backup);
    expect(() => decodeBackup(bytes.subarray(0, bytes.byteLength - 2))).toThrow(BackupFormatError);
  });

  test('an empty update, a backup of the base alone, reads back empty', () => {
    const read = decodeBackup(encodeBackup(sample({ update: new Uint8Array() })));
    expect(read.update.byteLength).toBe(0);
    expect([...read.base]).toEqual([80, 75, 3, 4, 1, 2, 3]);
  });

  test('a file cut short, or longer than its header says, is refused', () => {
    const bytes = encodeBackup(sample());
    for (const cut of [4, 10, bytes.byteLength - 1]) expect(() => decodeBackup(bytes.subarray(0, cut)), String(cut)).toThrow(BackupFormatError);
    const longer = new Uint8Array(bytes.byteLength + 1);
    longer.set(bytes);
    expect(() => decodeBackup(longer)).toThrow(BackupFormatError);
  });

  test('another format is refused: another magic, a header that does not read, one that is not a backup\'s', () => {
    expect(() => decodeBackup(new TextEncoder().encode('PK\u0003\u0004 a zip, not a backup'))).toThrow(BackupFormatError);
    const header = (json: string): Uint8Array => {
      const h = new TextEncoder().encode(json);
      const out = new Uint8Array(12 + h.length);
      out.set(new TextEncoder().encode('D4JBKUP1'));
      new DataView(out.buffer).setUint32(8, h.length);
      out.set(h, 12);
      return out;
    };
    expect(() => decodeBackup(header('{not json'))).toThrow(BackupFormatError);
    expect(() => decodeBackup(header('{"uri":1,"name":"a","meta":{"name":"a"},"baseLength":0,"updateLength":0}'))).toThrow(BackupFormatError);
    expect(() => decodeBackup(header('{"uri":"u","name":"a","meta":{},"baseLength":0,"updateLength":0}'))).toThrow(BackupFormatError);
    expect(() => decodeBackup(header('{"uri":"u","name":"a","meta":{"name":"a"},"baseLength":-1,"updateLength":0}'))).toThrow(BackupFormatError);
    expect(decodeBackup(header('{"uri":"u","name":"a","meta":{"name":"a"},"baseLength":0,"updateLength":0}')).name).toBe('a');
  });
});
