// SPDX-License-Identifier: GPL-3.0-or-later
//
// One page's channel (ED-014 section 4.3), with no `vscode` import: nothing is posted before the page says `ready`, and
// each request is an id and a promise its answer settles. When the page goes, the requests still waiting fail; a page
// that says `ready` again is a new page, so the old one's requests fail then too.
import { Requests } from './requests.js';
import type { FromPage, ToPage } from './shared.js';

/** The page's answers to requests: a save's bytes or refusal, a backup, a hand-over. */
export type Answer = Extract<FromPage, { type: 'saved' } | { type: 'backup' } | { type: 'handoff' }>;

export class Channel {
  private readied = 0;
  private outbox: ToPage[] = [];
  private readonly requests = new Requests<Answer>();
  private closed = false;

  constructor(private readonly send: (message: ToPage) => void) {}

  /** Whether the page has said `ready` at least once and has not gone. */
  get isReady(): boolean {
    return this.readied > 0 && !this.closed;
  }

  /** How many times a page has said `ready` here: more than once when it loaded again. */
  get loads(): number {
    return this.readied;
  }

  /** The page said `ready`: `first` goes before anything that waited. A page loaded again fails the old one's requests. */
  ready(first: ToPage): void {
    if (this.closed) return;
    if (this.readied > 0) this.requests.failAll();
    this.readied++;
    this.send(first);
    for (const m of this.outbox.splice(0)) this.send(m);
  }

  /**
   * A poster for the answers to a request of the page as it is now. An answer that comes after the page loaded again is
   * dropped, since the new page's ids start again and it would settle another of its requests (section 9.23 finding 2).
   * Take it before the first `await`.
   */
  replier(): (message: ToPage) => void {
    const load = this.readied;
    return (message) => { if (this.readied === load) this.post(message); };
  }

  post(message: ToPage): void {
    if (this.closed) return;
    if (this.readied > 0) this.send(message); else this.outbox.push(message);
  }

  /** A request: posted with a fresh id (or `reuse`, an answered one's), answered by the page's message with that id. */
  ask(make: (id: number) => ToPage, reuse?: number): { id: number; answer: Promise<Answer> } {
    if (this.closed) return { id: 0, answer: Promise.reject(new Error('the editor\'s page has gone')) };
    return this.requests.ask((id) => this.post(make(id)), reuse);
  }

  /** A message from the page: an answer settles its request; true if it was one. */
  answer(message: FromPage): boolean {
    if (message.type !== 'saved' && message.type !== 'backup' && message.type !== 'handoff') return false;
    this.requests.answer(message.id, message);
    return true;
  }

  /** The panel was disposed. */
  close(): void {
    this.closed = true;
    this.outbox = [];
    this.requests.failAll();
  }

  get waiting(): number {
    return this.requests.size;
  }
}
