// SPDX-License-Identifier: GPL-3.0-or-later

/** The source: the 255-page document. */
export const LARGE: string;

/** The sizes the figures are taken at: 200 and 1,000. */
export const SIZES: readonly number[];

export const AUTHORS: readonly string[];

/** `anon-250.docx` (or `source`) with `count` tracked edits made by the engine: its bytes (ED-005 section 9.4 item 1). */
export function buildRevisionScale(count: number, source?: Uint8Array): Promise<Uint8Array>;
