// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-013 section 9: a file changed on disk is judged by its bytes as well as its time, so that a time that moved over the
// bytes the editor last read or wrote (Android's MediaStore, after the editor's own write) is no change. In Node, whose
// `File` reads its bytes (jsdom's has no `arrayBuffer`).
import { describe, test, expect } from 'vitest';
import { changedSince } from '../src/app/files.js';
describe('changedSince', () => {
  const bytes = (...b: number[]): Uint8Array => new Uint8Array(b);
  const fileOf = (b: Uint8Array, lastModified: number): File => new File([b as BlobPart], 'a.docx', { lastModified });

  test('the same time is no change, whatever the bytes', async () => {
    expect(await changedSince(fileOf(bytes(1, 2, 3), 1000), 1000, bytes(9))).toBe(false);
  });

  test('a time that moved over the bytes known is no change (ED-013 section 9: Android, after the editor\'s own write)', async () => {
    expect(await changedSince(fileOf(bytes(1, 2, 3), 2000), 1000, bytes(1, 2, 3))).toBe(false);
  });

  test('a moved time with other bytes, another size, or none known, is a change', async () => {
    expect(await changedSince(fileOf(bytes(1, 2, 4), 2000), 1000, bytes(1, 2, 3))).toBe(true);
    expect(await changedSince(fileOf(bytes(1, 2), 2000), 1000, bytes(1, 2, 3))).toBe(true);
    expect(await changedSince(fileOf(bytes(1, 2, 3), 2000), 1000)).toBe(true);
  });
});
