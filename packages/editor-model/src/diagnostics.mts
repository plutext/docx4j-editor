// SPDX-License-Identifier: GPL-3.0-or-later
//
// What the projection could not make editable, reported rather than hidden (ED-001 section 11
// item 8): every opaque node with its kind and address, and the tracked changes: those the edit
// layer holds as content since E4.c (ED-005 section 5.2), and those it keeps as they are, shown in
// Word only (section 5.2's preserved list).

export interface OpaqueEntry {
  /** The element's local name (`tbl`, `fldChar`, `bookmarkStart`, `oMath`, ...). */
  kind: string;
  /** `body/3`, `body/3/sdt/0`, `body/5/r/2` ...: the path of block indexes from the body. */
  address: string;
  inline: boolean;
  sourceId: string;
}

export interface Diagnostics {
  opaque: OpaqueEntry[];
  tracked: {
    /**
     * Tracked changes the edit layer holds (ED-005 section 5.2): text inserted, deleted or moved
     * (a `w:ins`, `w:del`, `w:moveTo` or `w:moveFrom` element each), a run's or a paragraph's
     * properties changed, a paragraph mark or a row inserted or deleted.
     */
    shown: number;
    /** Tracked changes kept as they are and shown in Word only, by element (`tblPrChange`, `sectPrChange`, `ins` in a note, ...). */
    preserved: Record<string, number>;
  };
  /** `w:lastRenderedPageBreak` markers seen; they are decoration sources, not nodes. */
  pageBreakMarkers: number;
  /** Distinct comment ids whose ranges were projected as marks. */
  comments: number;
}

export function emptyDiagnostics(): Diagnostics {
  return { opaque: [], tracked: { shown: 0, preserved: {} }, pageBreakMarkers: 0, comments: 0 };
}

/** How many tracked changes are kept and shown in Word only. */
export function preservedCount(d: Diagnostics): number {
  return Object.values(d.tracked.preserved).reduce((a, b) => a + b, 0);
}

/** A one-line summary for a banner: "12 elements shown read-only; 3 tracked changes; 1 tracked change shown in Word only". */
export function summarize(d: Diagnostics): string {
  const parts: string[] = [];
  if (d.opaque.length > 0) parts.push(`${d.opaque.length} element${d.opaque.length === 1 ? '' : 's'} shown read-only`);
  if (d.tracked.shown > 0) parts.push(`${d.tracked.shown} tracked change${d.tracked.shown === 1 ? '' : 's'}`);
  const kept = preservedCount(d);
  if (kept > 0) parts.push(`${kept} tracked change${kept === 1 ? '' : 's'} shown in Word only`);
  return parts.join('; ');
}
