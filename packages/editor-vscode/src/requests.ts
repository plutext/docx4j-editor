// SPDX-License-Identifier: GPL-3.0-or-later
//
// Requests to the page (ED-014 section 4.3), with no `vscode` import: each an id and a promise, answered by the page's
// message with that id. When the page goes (the panel disposed, or the page loaded again), the requests still waiting fail.

export class PageGone extends Error {
  override name = 'PageGone';
}

export class Requests<T> {
  private next = 1;
  private readonly waiting = new Map<number, { resolve: (value: T) => void; reject: (error: Error) => void }>();

  /**
   * A new request: `post` sends it with its id; the promise is its answer. `reuse` asks again under an id already
   * answered, for a request that continues another (a Save As's hand-over, under its save's id).
   */
  ask(post: (id: number) => void, reuse?: number): { id: number; answer: Promise<T> } {
    if (reuse !== undefined && this.waiting.has(reuse)) return { id: reuse, answer: Promise.reject(new Error(`request ${reuse} is still waiting`)) };
    const id = reuse ?? this.next++;
    const answer = new Promise<T>((resolve, reject) => this.waiting.set(id, { resolve, reject }));
    try {
      post(id);
    } catch (e) {
      this.waiting.delete(id);
      return { id, answer: Promise.reject(e instanceof Error ? e : new Error(String(e))) };
    }
    return { id, answer };
  }

  /** The page's answer; false for an id no request waits on (a stale page's, or one already answered). */
  answer(id: number, value: T): boolean {
    const waiting = this.waiting.get(id);
    if (!waiting) return false;
    this.waiting.delete(id);
    waiting.resolve(value);
    return true;
  }

  /** Every request still waiting fails. */
  failAll(error: Error = new PageGone('the editor\'s page has gone')): void {
    const all = [...this.waiting.values()];
    this.waiting.clear();
    for (const w of all) w.reject(error);
  }

  get size(): number {
    return this.waiting.size;
  }
}
