// SPDX-License-Identifier: GPL-3.0-or-later

/** The ids the first round tries (ED-005 section 9.2 item 3). */
export const IDS: readonly number[];

/** One file of the bisection: a footnote or a comment whose `w:id` is `id`. */
export interface IdFile { name: string; bytes: Uint8Array; kind: 'footnote' | 'comment'; id: number }

/** The files for each kind and each id. */
export function buildIdsForWord(ids?: readonly number[]): Promise<IdFile[]>;
