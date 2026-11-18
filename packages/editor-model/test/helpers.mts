// SPDX-License-Identifier: GPL-3.0-or-later
import { readFile, readdir } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { ZipPartStore, WordprocessingMLPackage } from '@docx4j/core-ts';
import type { Node as PMNode } from 'prosemirror-model';

export const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), 'fixtures');

export async function fixture(name: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(join(fixturesDir, name)));
}

/**
 * Every .docx in the fixtures directory, in `opendope/` (E2.b's templates), in `revisions/` (E4.c's
 * Word-made revisions and their base), in `comments/` (E4.e's Word saves) and in `review/` (E4.g's
 * Word session: the editor's file, Word's saves of it untouched, accepted and rejected, and the
 * measurements), sorted.
 */
export async function allFixtures(): Promise<string[]> {
  const root = (await readdir(fixturesDir)).filter((n) => n.endsWith('.docx'));
  const sub = async (dir: string): Promise<string[]> => (await readdir(join(fixturesDir, dir))).filter((n) => n.endsWith('.docx')).map((n) => `${dir}/${n}`);
  return [...root, ...(await sub('opendope')), ...(await sub('revisions')), ...(await sub('comments')), ...(await sub('review'))].sort();
}

export function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** The parts that regenerate by design (CR-001 section 12): content types and relationships. */
export function regenerated(name: string): boolean {
  return name === '[Content_Types].xml' || name.endsWith('.rels');
}

export interface PartDifference {
  name: string;
  reason: 'missing' | 'added' | 'bytes';
}

/**
 * ED-002 section 3.7's oracle: same entries, every part byte-identical except the regenerated
 * ones and those listed in `expectChanged`. Returns the differences (empty when the oracle holds).
 */
export function comparePackages(before: Uint8Array, after: Uint8Array, expectChanged: string[] = []): PartDifference[] {
  const a = new ZipPartStore(before);
  const b = new ZipPartStore(after);
  const an = new Set(a.partNames());
  const bn = new Set(b.partNames());
  const out: PartDifference[] = [];
  for (const n of an) if (!bn.has(n)) out.push({ name: n, reason: 'missing' });
  for (const n of bn) if (!an.has(n)) out.push({ name: n, reason: 'added' });
  for (const n of an) {
    if (!bn.has(n) || regenerated(n) || expectChanged.includes(n)) continue;
    if (!bytesEqual(a.loadSync(n), b.loadSync(n))) out.push({ name: n, reason: 'bytes' });
  }
  return out;
}

/** The main document part of saved bytes as text. */
export async function mainXml(bytes: Uint8Array): Promise<string> {
  const pkg = await WordprocessingMLPackage.load(bytes);
  return pkg.getMainDocumentPart().getXml();
}

/** A structural copy without PARENT pointers, for deep equality (the engine's `plain`). */
export function plain<T>(value: T): T {
  if (value === undefined) return value;
  return JSON.parse(JSON.stringify(value, (k, v) => (k === 'PARENT' ? undefined : v))) as T;
}

/** The text of a block node, hard breaks and tabs as their characters. */
export function textOf(node: PMNode): string {
  let out = '';
  node.descendants((n) => {
    if (n.isText) out += n.text;
    else if (n.type.name === 'tab') out += '\t';
    else if (n.type.name === 'hardBreak') out += '\n';
    return true;
  });
  return out;
}

/**
 * A date as Word writes `w:date`: the local wall-clock time, with a `Z` (core-ts 0.3.0's
 * `wordCalendarOf`; CR-002 section 29, ED-005 section 12.5 item 12). What it gives depends on the
 * process's time zone, which is the point: `npm run test:zones` runs the date tests in four.
 */
export function wordDate(date: Date | string): string {
  const d = typeof date === 'string' ? new Date(date) : date;
  const two = (n: number): string => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${two(d.getMonth() + 1)}-${two(d.getDate())}T${two(d.getHours())}:${two(d.getMinutes())}:${two(d.getSeconds())}Z`;
}

/**
 * A revision's author and date as the editor writes them since core-ts 0.3.0, for a pattern:
 * `w16du:dateUtc` (UTC) before the author, then `w:date` on the local wall clock.
 */
export function revisionAttrs(author: string, iso: string): string {
  return `w16du:dateUtc="${iso}" w:author="${author}" w:date="${wordDate(iso)}"`;
}

/** A pattern with `@C` standing for Author C's attributes at 2026-09-28T10:00:00Z, the tests' tracking date. */
export function dated(pattern: RegExp): RegExp {
  return new RegExp(pattern.source.replaceAll('@C', revisionAttrs('Author C', '2026-09-28T10:00:00Z')), pattern.flags);
}
