// SPDX-License-Identifier: GPL-3.0-or-later
//
// A watcher on an open file's handle (ED-011 section 3.2): the file's modification time read through
// `getFile()` every few seconds while the tab is visible, and at once when the tab becomes visible or
// the window takes focus, which is when a person comes back from Word. A change is reported once;
// a later write is reported again. `getFile()` reads metadata alone and never touches the document.

export interface FileWatcher {
  /** Read the file's time now and report a change. */
  check(): Promise<void>;
  stop(): void;
}

export interface WatchOptions {
  /** The polling interval while visible; five seconds by default (ED-011 P2). */
  intervalMs?: number;
}

export function watchFile(handle: FileSystemFileHandle, openedAt: number, onChange: (lastModified: number) => void, options: WatchOptions = {}): FileWatcher {
  const intervalMs = options.intervalMs ?? 5_000;
  let reported = openedAt;
  let stopped = false;
  let checking = false;
  let timer: number | undefined;
  const check = async (): Promise<void> => {
    if (stopped || checking) return;
    checking = true;
    try {
      const file = await handle.getFile();
      if (!stopped && file.lastModified !== reported) {
        reported = file.lastModified;
        onChange(file.lastModified);
      }
    } catch {
      // A handle whose permission lapsed, or whose file is gone: nothing to report; the save says so.
    } finally {
      checking = false;
    }
  };
  const start = (): void => {
    if (timer === undefined && document.visibilityState === 'visible') timer = window.setInterval(() => void check(), intervalMs);
  };
  const pause = (): void => {
    if (timer !== undefined) { clearInterval(timer); timer = undefined; }
  };
  const onVisibility = (): void => {
    if (document.visibilityState === 'visible') { start(); void check(); } else pause();
  };
  const onFocus = (): void => void check();
  document.addEventListener('visibilitychange', onVisibility);
  window.addEventListener('focus', onFocus);
  start();
  return {
    check,
    stop() {
      stopped = true;
      pause();
      document.removeEventListener('visibilitychange', onVisibility);
      window.removeEventListener('focus', onFocus);
    },
  };
}
