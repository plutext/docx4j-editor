// SPDX-License-Identifier: GPL-3.0-or-later
//
// What a tab reads when it asks whether a session is open elsewhere (ED-011 section 3.6), on the paths a browser
// suite never takes: a page with no Web Locks, and a refusal that is not a wait given up. The locks themselves are
// the browser's, in e2e/recovery-tabs.spec.ts.
import { afterEach, describe, expect, it, vi } from 'vitest';
import { holdSession, sessionInUse } from '../src/app/persistence.js';

/** A stand-in for `navigator.locks` whose every request ends as `request` says. */
function withLocks(request: (...args: unknown[]) => Promise<unknown>): void {
  Object.defineProperty(navigator, 'locks', { configurable: true, value: { request } });
}

afterEach(() => { delete (navigator as unknown as { locks?: unknown }).locks; vi.restoreAllMocks(); });

describe('a session open elsewhere', () => {
  it('reads as free where the browser has no Web Locks, and holding one does nothing', async () => {
    expect((navigator as unknown as { locks?: unknown }).locks).toBeUndefined();
    expect(await sessionInUse('s')).toBe(false);
    expect(() => holdSession('s')()).not.toThrow();
  });

  it('reads as in use only when the wait for the lock ran out', async () => {
    withLocks(() => Promise.reject(new DOMException('waited', 'TimeoutError')));
    expect(await sessionInUse('s')).toBe(true);
  });

  it('reads as free when the browser refuses to lock at all: an offer must not be lost to that', async () => {
    withLocks(() => Promise.reject(new DOMException('no', 'SecurityError')));
    expect(await sessionInUse('s')).toBe(false);
  });

  it('asks for the lock whole, and holds it shared', async () => {
    const asked: Array<Record<string, unknown>> = [];
    withLocks((_name, options, callback) => { asked.push(options as Record<string, unknown>); return (callback as () => Promise<unknown>)(); });
    expect(await sessionInUse('s')).toBe(false);
    expect(asked[0]!['mode']).toBeUndefined();
    const release = holdSession('s');
    expect(asked[1]!['mode']).toBe('shared');
    release();
  });

  it('a hold refused does not reject where nothing is waiting for it', async () => {
    withLocks(() => Promise.reject(new DOMException('no', 'SecurityError')));
    const release = holdSession('s');
    await Promise.resolve();
    expect(() => release()).not.toThrow();
  });
});
