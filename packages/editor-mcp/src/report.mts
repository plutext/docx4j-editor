// SPDX-License-Identifier: GPL-3.0-or-later
//
// The change report (docs/developer/agent.md section 2.7; vocabulary rule 6): what a mutating call did, as
// structured content and a one-line summary. Each touched and created paragraph with its new text
// as accepted (the agent's next `if_text`), the blocks removed, the revisions the call wrote, the
// parts it changed untracked (rule 1), and warnings. Ordinals are as of the moment the call's change
// was made.
//
// The revisions are listed as Office JS lists them (the engine's `Body.getTrackedChanges`): touching
// insertions by one author across paragraphs are one change, at the paragraph it starts in. So they
// are listed over the top-level blocks the call involved, each contiguous run of them as a body of
// its own, and what the agent had pending there before the call is taken out.
import { Body, Paragraph, Table, type Element, type TrackedChange } from '@docx4j/core-ts';
import { describeParagraph } from './address.mjs';

export interface ReportParagraph {
  address: string;
  paraId?: string;
  text: string;
}

export interface ReportRevision {
  kind: 'insertion' | 'deletion' | 'formatting' | 'properties';
  author: string;
  address: string;
  text: string;
}

export interface ChangeReport {
  summary: string;
  touched: ReportParagraph[];
  created: ReportParagraph[];
  removed: { address: string }[];
  revisions: ReportRevision[];
  parts: string[];
  warnings: string[];
  /** Set when the call was a dry run: nothing was written. */
  dry_run?: true;
  [key: string]: unknown;
}

function kindOf(change: TrackedChange): ReportRevision['kind'] {
  if (change.type === 'Added') return 'insertion';
  if (change.type === 'Deleted') return 'deletion';
  return change.target.kind === 'paragraphProperties' ? 'properties' : 'formatting';
}

const keyOf = (c: TrackedChange): string => `${kindOf(c)}|${c.text}`;

/** The report of one call, built as it runs. */
export class ReportBuilder {
  private readonly touched: Paragraph[] = [];
  private readonly created: (Paragraph | Table)[] = [];
  /** A removed block and its address as it was: an agent's own insertion deleted goes for real, and has none after. */
  private readonly removed: { block: Paragraph | Table; address: string }[] = [];
  private readonly parts = new Set<string>();
  readonly warnings: string[] = [];
  /** The top-level blocks the call involves, by the value object, and the agent's revisions there before it. */
  private readonly blocks = new Set<object>();
  private readonly before: string[] = [];
  /** A call that tracks nothing (a comment, rule 1): its report lists no revisions. */
  private tracksNothing = false;

  constructor(readonly body: Body, readonly author: string) {}

  /** The top-level block an element is in (itself, a table holding it, a control holding it). */
  private topOf(value: object): object | undefined {
    let v = value as { PARENT?: object };
    for (let depth = 0; v && depth < 64; depth++) {
      if (v.PARENT === this.body.container) return v;
      v = v.PARENT as { PARENT?: object };
    }
    return undefined;
  }

  /** The agent's revisions over top-level blocks, each contiguous run of them listed as Office JS lists a body. */
  private revisionsIn(blocks: ReadonlySet<object>): TrackedChange[] {
    const content = this.body.content;
    const indices = content.map((el, i) => (blocks.has(el.value as object) ? i : -1)).filter((i) => i >= 0);
    const out: TrackedChange[] = [];
    for (let k = 0; k < indices.length;) {
      let end = k;
      while (end + 1 < indices.length && indices[end + 1] === indices[end]! + 1) end++;
      const run = new Body(this.body.part, { content: content.slice(indices[k], indices[end]! + 1) }, this.body.prefix, this.body.package_);
      out.push(...run.getTrackedChanges().filter((c) => c.author === this.author));
      k = end + 1;
    }
    return out;
  }

  /** A block's top-level block noted, and the agent's revisions there taken as they are before the call. */
  private involve(value: object): void {
    const top = this.topOf(value);
    if (!top || this.blocks.has(top)) return;
    this.blocks.add(top);
    for (const c of this.revisionsIn(new Set([top]))) this.before.push(keyOf(c));
  }

  /** A paragraph the call is about to change; call it before the change. */
  touch(paragraph: Paragraph): void {
    if (this.touched.some((p) => p.p === paragraph.p)) return;
    this.involve(paragraph.p);
    this.touched.push(paragraph);
  }

  /** A block the call made (call it after the insertion). */
  create(block: Paragraph | Table): void {
    this.created.push(block);
    const top = this.topOf(block instanceof Paragraph ? block.p : block.tbl);
    if (top) this.blocks.add(top);
  }

  /** A block the call is about to delete (as a tracked change it stays in the document, deleted); call it before the deletion. */
  remove(block: Paragraph | Table): void {
    this.involve(block instanceof Paragraph ? block.p : block.tbl);
    const address = block instanceof Paragraph ? this.body.addressOf(block) : this.body.addressOf({ element: block.element, container: block.container });
    this.removed.push({ block, address: address ?? '?' });
  }

  /** A package part the call changed beside the content (rule 1: untracked). */
  part(name: string): void {
    this.parts.add(name);
  }

  /**
   * The call makes no revision (a comment, rule 1). Its markers may still part a revision the agent
   * made before into two listed changes, which the before-and-after comparison would take for new
   * ones (ED-005 section 12.29).
   */
  untracked(): void {
    this.tracksNothing = true;
  }

  warn(message: string): void {
    if (!this.warnings.includes(message)) this.warnings.push(message);
  }

  /** Where a change is: its paragraph's address, or its row's or table's. */
  private addressOfChange(c: TrackedChange): string {
    if (c.paragraph) return this.body.addressOf(c.paragraph) ?? '?';
    const t = c.target as { tr?: Element; tbl?: Element };
    const element = t.tr ?? t.tbl;
    return element ? this.body.addressOf({ element, container: [] }) ?? '?' : '?';
  }

  /** The report as it stands now; `summary` is the call's one line. */
  finish(summary: string, extra: Record<string, unknown> = {}): ChangeReport {
    const earlier = [...this.before];
    const revisions: ReportRevision[] = [];
    for (const c of this.tracksNothing ? [] : this.revisionsIn(this.blocks)) {
      const at = earlier.indexOf(keyOf(c));
      if (at >= 0) { earlier.splice(at, 1); continue; }
      revisions.push({ kind: kindOf(c), author: c.author, address: this.addressOfChange(c), text: c.text });
    }
    const describe = (block: Paragraph | Table): ReportParagraph => (block instanceof Paragraph
      ? describeParagraph(this.body, block)
      : { address: this.body.addressOf({ element: block.element, container: block.container }) ?? '?', text: block.text });
    return {
      summary,
      touched: this.touched.map((p) => describeParagraph(this.body, p)),
      created: this.created.map(describe),
      removed: this.removed.map((r) => ({ address: r.address })),
      revisions,
      parts: [...this.parts],
      warnings: this.warnings,
      ...extra,
    };
  }
}
