// SPDX-License-Identifier: GPL-3.0-or-later
//
// Hand-overs across a Save As (ED-014 section 4.3, P25), with no `vscode` import. VS Code replaces the editor with a new
// one for the target, which opens a new page from the bytes just written; the old page's place, and its store if an edit
// came after the export, wait here for the target's `openCustomDocument`. Each is taken once, and goes after ten seconds:
// if the person's associations open the target in another editor, nothing takes it.

export const HANDOVER_MS = 10_000;

export interface Timers {
  setTimeout(f: () => void, ms: number): unknown;
  clearTimeout(handle: unknown): void;
}

const realTimers: Timers = {
  setTimeout: (f, ms) => setTimeout(f, ms),
  clearTimeout: (h) => clearTimeout(h as ReturnType<typeof setTimeout>),
};

export class Handovers<T> {
  private readonly kept = new Map<string, { value: T; timer: unknown }>();

  constructor(private readonly ms = HANDOVER_MS, private readonly timers: Timers = realTimers) {}

  /** Kept for the target, in place of any hand-over already waiting for it. */
  put(target: string, value: T): void {
    this.drop(target);
    const timer = this.timers.setTimeout(() => this.kept.delete(target), this.ms);
    this.kept.set(target, { value, timer });
  }

  /** Taken once. */
  take(target: string): T | undefined {
    const kept = this.kept.get(target);
    if (!kept) return undefined;
    this.drop(target);
    return kept.value;
  }

  has(target: string): boolean {
    return this.kept.has(target);
  }

  private drop(target: string): void {
    const kept = this.kept.get(target);
    if (!kept) return;
    this.timers.clearTimeout(kept.timer);
    this.kept.delete(target);
  }
}
