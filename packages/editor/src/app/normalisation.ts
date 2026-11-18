// SPDX-License-Identifier: GPL-3.0-or-later
//
// The normalisation report (ED-008 section 8.8): what writing the document's values in their normal form would
// change, and what the engine noted reading its parts. A save writes values as read (`normaliseOnSave` off, Jason,
// 2026-10-26), so this is a report and nothing more: it reads a package of its own, made from the bytes the next save
// would write, so the session's edits are in it in every part, the notes and headers too. That save writes nothing
// for an unedited document, which still saves byte for byte; for an edited one it does what Re-paginate's does.
import { WordprocessingMLPackage, savePackage, type PackageWarning, type PartSink, type ValueNormalised } from '@docx4j/core-ts';
import type { EditorDocument } from '@docx4j/editor-model';

/** One value and what normal form makes of it, null where it is left out, with how often the kind has it. */
export interface ValueChange { before: string; after: string | null; count: number }
/** The changes to one element or attribute of a part: `w:color/@w:val`, `w:highlight`. */
export interface NormalisedKind { kind: string; count: number; changes: ValueChange[] }
export interface NormalisedPart { partName: string; count: number; kinds: NormalisedKind[] }
/** What the engine noted reading the parts, by code; each message names its part. */
export interface PartNotes { code: string; messages: string[] }

export interface NormalisationReport {
  values: number;
  parts: NormalisedPart[];
  notes: PartNotes[];
}

/** The kind of a value's path: its element and attribute, positions left out. */
export function kindOf(path: string): string {
  const steps = path.split('/').filter((s) => s !== '').map((s) => s.replace(/\[\d+\]$/, ''));
  const last = steps[steps.length - 1] ?? path;
  return last.startsWith('@') ? `${steps[steps.length - 2] ?? ''}/${last}` : last;
}

/** The most frequent first, then by name, so a report reads the same each time. */
const byCount = <T extends { count: number }>(name: (item: T) => string) => (a: T, b: T): number => b.count - a.count || name(a).localeCompare(name(b));

/** A document can hold thousands of these: by part, by kind, by value, each with its count. */
export function groupNormalised(entries: readonly ValueNormalised[]): NormalisedPart[] {
  const parts = new Map<string, Map<string, Map<string, ValueChange>>>();
  for (const e of entries) {
    let kinds = parts.get(e.partName);
    if (!kinds) parts.set(e.partName, kinds = new Map());
    const kind = kindOf(e.path);
    let changes = kinds.get(kind);
    if (!changes) kinds.set(kind, changes = new Map());
    const before = e.before ?? '';
    const key = JSON.stringify([before, e.after]);
    const change = changes.get(key);
    if (change) change.count++;
    else changes.set(key, { before, after: e.after, count: 1 });
  }
  return [...parts].map(([partName, kinds]) => {
    const list = [...kinds].map(([kind, changes]) => ({
      kind,
      count: [...changes.values()].reduce((n, c) => n + c.count, 0),
      changes: [...changes.values()].sort(byCount((c) => c.before)),
    })).sort(byCount((k) => k.kind));
    return { partName, count: list.reduce((n, k) => n + k.count, 0), kinds: list };
  }).sort(byCount((p) => p.partName));
}

export function groupNotes(warnings: readonly PackageWarning[]): PartNotes[] {
  const codes = new Map<string, string[]>();
  for (const w of warnings) {
    const messages = codes.get(w.code);
    if (messages) messages.push(w.message);
    else codes.set(w.code, [w.message]);
  }
  return [...codes].map(([code, messages]) => ({ code, messages })).sort((a, b) => b.messages.length - a.messages.length || a.code.localeCompare(b.code));
}

/** The report's save wants what the engine learns from writing, never the bytes. */
const DISCARD: PartSink<void> = { put: () => undefined, finish: () => Promise.resolve() };

export async function normalisationReport(session: EditorDocument): Promise<NormalisationReport> {
  const { bytes } = await session.save({ selfCheck: false });
  const pkg = await WordprocessingMLPackage.load(bytes);
  pkg.normaliseOnSave = true;
  // Every part, so each is marshalled: a part only read is copied from its source bytes, and normalises nothing.
  await pkg.unmarshalAll();
  await savePackage(pkg, DISCARD);
  const normalised = pkg.lastSave?.normalised ?? [];
  return { values: normalised.length, parts: groupNormalised(normalised), notes: groupNotes(pkg.warnings) };
}
