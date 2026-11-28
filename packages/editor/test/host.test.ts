// SPDX-License-Identifier: GPL-3.0-or-later
//
// The host seam's VS Code half (ED-014 section 4.1) over a fake `acquireVsCodeApi`: each message both ways, the queue and
// `ready` once, the options filled at boot and written back, a refused save, `changed` held to its interval and posted
// again when an edit came between the export and the write, and a backup answered while a save is in flight. The page is
// a fake here; the application's own half is held by `host-app.test.tsx` and `e2e/host-vscode.spec.ts`.
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { initHost, resetHost, vscodeHost, browserHost, AskFailure, type HostFeature, type HostPage, type VsCodeHost } from '../src/app/host.js';
import { HOST_META, CHANGED_INTERVAL, PAGE_COMMANDS, type FromPage, type ToPage } from '../src/app/host-protocol.js';
import { describeFailure, setKeyCommand } from '../src/app/ask.js';
import { t } from '../src/app/i18n.js';

/** The features the browser has and VS Code lacks (`documentMacros` VS Code has in a trusted workspace alone, below). */
const FEATURES: HostFeature[] = ['open', 'new', 'close', 'saveInPlace', 'saveAs', 'download', 'print', 'sendCopy', 'join', 'recovery', 'diskInTab', 'layoutChoice', 'startPage', 'appKeys', 'unloadGuard', 'documentMacros', 'spelling', 'companionOptions', 'askOptions'];

/** A fake of VS Code's API: what the page posts, and an acquire that throws the second time, as VS Code's does. */
function fakeApi() {
  const posted: FromPage[] = [];
  let acquired = 0;
  const api = { postMessage: (m: unknown) => { posted.push(m as FromPage); } };
  return { posted, api, acquire: () => { acquired++; if (acquired > 1) throw new Error('An instance of the VS Code API has already been acquired'); return api; } };
}

type Call = [keyof HostPage, unknown[]];

/** A page that records its calls; `over` replaces a method. */
function fakePage(over: Partial<HostPage> = {}): HostPage & { calls: Call[] } {
  const calls: Call[] = [];
  const record = <K extends keyof HostPage>(name: K, fallback: (...args: never[]) => unknown) => ((...args: unknown[]) => {
    calls.push([name, args]);
    const f = (over[name] ?? fallback) as (...a: unknown[]) => unknown;
    return f(...args);
  }) as HostPage[K];
  return {
    calls,
    open: record('open', async () => undefined),
    save: record('save', async () => ({ bytes: new Uint8Array([1, 2, 3]) })),
    written: record('written', () => undefined),
    writeFailed: record('writeFailed', () => undefined),
    backup: record('backup', () => ({ update: new Uint8Array([9]), meta: { name: 'a.docx' } })),
    handoff: record('handoff', () => ({ place: { at: 'here' } })),
    revert: record('revert', async () => undefined),
    diskChanged: record('diskChanged', () => undefined),
    showPalette: record('showPalette', () => undefined),
    pasteText: record('pasteText', () => undefined),
    setCompanion: record('setCompanion', () => undefined),
    reloadFromDisk: record('reloadFromDisk', async () => undefined),
  };
}

const send = (m: ToPage | Record<string, unknown>): void => { window.dispatchEvent(new MessageEvent('message', { data: m })); };
const tick = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));
const ofType = <T extends FromPage['type']>(posted: FromPage[], type: T): Extract<FromPage, { type: T }>[] => posted.filter((m): m is Extract<FromPage, { type: T }> => m.type === type);

let hosts: VsCodeHost[] = [];
const make = (api: { postMessage: (m: unknown) => void }): VsCodeHost => { const h = vscodeHost(api); hosts.push(h); return h; };

beforeEach(() => { localStorage.clear(); sessionStorage.clear(); });
afterEach(() => {
  for (const h of hosts) h.stop();
  hosts = [];
  resetHost();
  delete (globalThis as { acquireVsCodeApi?: unknown }).acquireVsCodeApi;
  document.head.querySelectorAll(`meta[name="${HOST_META}"]`).forEach((m) => m.remove());
  localStorage.clear();
  vi.useRealTimers();
});

describe('choosing the host', () => {
  test('VS Code\'s where acquireVsCodeApi exists, acquired once though initHost is called twice', () => {
    const fake = fakeApi();
    (globalThis as { acquireVsCodeApi?: unknown }).acquireVsCodeApi = fake.acquire;
    expect(initHost().kind).toBe('vscode');
    expect(() => initHost()).not.toThrow();
    expect(initHost().kind).toBe('vscode');
  });

  test('the browser\'s otherwise, with every feature but the page\'s own plain paste; VS Code\'s has that one alone', () => {
    expect(initHost().kind).toBe('browser');
    const browser = browserHost();
    const vscode = make(fakeApi().api);
    for (const f of FEATURES) {
      expect(browser.has(f), f).toBe(true);
      expect(vscode.has(f), f).toBe(false);
    }
    // The browser's own Ctrl+Shift+V works; VS Code's host page cancels it, so the page takes the key (ED-014 P34).
    expect(browser.has('clipboardText')).toBe(false);
    expect(vscode.has('clipboardText')).toBe(true);
  });
});

describe('the queue and ready', () => {
  test('ready is posted once across a double attach, and an open sent before the attach goes to the page attached last', async () => {
    const fake = fakeApi();
    const h = make(fake.api);
    send({ type: 'open', name: 'a.docx', bytes: new Uint8Array([80, 75]) });
    const first = fakePage();
    const second = fakePage();
    // React's development StrictMode: mount, unmount, mount, in one commit.
    const detach = h.attach(first);
    detach();
    h.attach(second);
    await tick();
    expect(ofType(fake.posted, 'ready')).toEqual([{ type: 'ready', origin: location.origin }]);
    expect(first.calls).toEqual([]);
    expect(second.calls.map(([name]) => name)).toEqual(['open']);
    const [file] = second.calls[0]![1] as [{ name: string; bytes: Uint8Array }];
    expect(file.name).toBe('a.docx');
    expect(file.bytes).toBeInstanceOf(Uint8Array);
    expect([...file.bytes]).toEqual([80, 75]);
  });

  test('a message that arrives between the attach and the queue\'s drain keeps its place behind the queue', async () => {
    const h = make(fakeApi().api);
    send({ type: 'open', name: 'a.docx', bytes: new Uint8Array([80, 75]) });
    const page = fakePage();
    h.attach(page);
    send({ type: 'palette' });
    await tick();
    expect(page.calls.map(([name]) => name)).toEqual(['open', 'showPalette']);
  });
});

describe('the messages from the extension', () => {
  test('each reaches the page', async () => {
    const fake = fakeApi();
    const h = make(fake.api);
    const page = fakePage();
    h.attach(page);
    await tick();
    const link = { room: 'r'.repeat(16), token: 't', relay: 'ws://127.0.0.1:4175' };
    send({ type: 'open', name: 'shared.docx', join: link });
    send({ type: 'open', name: 'b.docx', bytes: new Uint8Array([1]), resume: { update: new Uint8Array([7]), meta: { name: 'b.docx', joined: true } } });
    // A Node Buffer as VS Code's serializer turns it into a plain object (S0): read as the bytes it was.
    send({ type: 'revert', bytes: { type: 'Buffer', data: [4, 5] } } as Record<string, unknown>);
    send({ type: 'disk', bytes: new Uint8Array([6]), lastModified: 42 });
    send({ type: 'companion', url: 'http://127.0.0.1:4174', relay: 'ws://127.0.0.1:4175' });
    send({ type: 'palette' });
    send({ type: 'pasteText', text: 'plain' });
    send({ type: 'writeFailed', id: 3, error: 'disk full' });
    send({ type: 'backup', id: 4 });
    send({ type: 'nonsense' });
    send({});
    await tick();
    const names = page.calls.map(([name]) => name);
    // In order, each after an open or a revert has finished (ED-014 section 4.3); the backup at once, never behind them.
    expect(names).toEqual(['open', 'backup', 'open', 'revert', 'diskChanged', 'setCompanion', 'showPalette', 'pasteText', 'writeFailed']);
    expect(page.calls.find(([name]) => name === 'pasteText')![1]).toEqual(['plain']);
    expect(page.calls[0]![1]).toEqual([{ name: 'shared.docx', join: link }]);
    const [resumed, how] = page.calls[2]![1] as [{ bytes: Uint8Array }, { resume: { update: Uint8Array; meta: unknown } }];
    expect([...resumed.bytes]).toEqual([1]);
    expect([...how.resume.update]).toEqual([7]);
    expect(how.resume.meta).toEqual({ name: 'b.docx', joined: true });
    expect([...(page.calls[3]![1][0] as Uint8Array)]).toEqual([4, 5]);
    expect(page.calls[4]![1][1]).toBe(42);
    expect(page.calls[5]![1]).toEqual(['http://127.0.0.1:4174', 'ws://127.0.0.1:4175']);
    expect(page.calls[8]![1]).toEqual(['disk full']);
    expect(ofType(fake.posted, 'backup')).toEqual([{ type: 'backup', id: 4, update: new Uint8Array([9]), meta: { name: 'a.docx' } }]);
  });

  test('a backup with no document open is empty', async () => {
    const fake = fakeApi();
    const h = make(fake.api);
    h.attach(fakePage({ backup: () => undefined }));
    await tick();
    send({ type: 'backup', id: 1 });
    expect(ofType(fake.posted, 'backup')).toEqual([{ type: 'backup', id: 1, empty: true }]);
  });
});

describe('the routes the page takes', () => {
  test('each is posted', async () => {
    const fake = fakeApi();
    const h = make(fake.api);
    const link = { room: 'r'.repeat(16), token: 't', relay: 'ws://127.0.0.1:4175' };
    await h.openDocument({ name: 'Document1.docx', bytes: new Uint8Array([1]), lastModified: 0 });
    await h.openDocument({ name: '', join: link });
    h.reload(true);
    h.openExternal('https://example.org/');
    h.download(new Uint8Array([2]), 'a.diagnosis.docx');
    h.base(new Uint8Array([3]));
    h.status('Could not open a.docx');
    expect(fake.posted).toEqual([
      { type: 'untitled', name: 'Document1.docx', bytes: new Uint8Array([1]) },
      { type: 'untitled', name: '', join: link },
      { type: 'revertRequest' },
      { type: 'openExternal', url: 'https://example.org/' },
      { type: 'download', bytes: new Uint8Array([2]), name: 'a.diagnosis.docx' },
      { type: 'base', bytes: new Uint8Array([3]) },
      { type: 'status', text: 'Could not open a.docx' },
    ]);
  });

  test('a room link is built on the companion\'s address once the host has posted it, never the webview\'s', async () => {
    const h = make(fakeApi().api);
    h.attach(fakePage());
    expect(h.siteAddress()).toBe(location.href.replace(/#.*$/, ''));
    send({ type: 'companion', url: 'http://127.0.0.1:4174', relay: 'ws://127.0.0.1:4175' });
    expect(h.siteAddress()).toBe('http://127.0.0.1:4174/');
  });
});

describe('saving in two phases', () => {
  test('a refused save, and one that throws, are posted as refused', async () => {
    const fake = fakeApi();
    const h = make(fake.api);
    let throws = false;
    h.attach(fakePage({ save: async () => { if (throws) throw new Error('boom'); return { refused: 'Not saved: the round-trip check failed' }; } }));
    await tick();
    send({ type: 'save', id: 1, name: 'a.docx' });
    await tick();
    throws = true;
    send({ type: 'save', id: 2, name: 'a.docx' });
    await tick();
    expect(ofType(fake.posted, 'saved')).toEqual([
      { type: 'saved', id: 1, refused: 'Not saved: the round-trip check failed' },
      { type: 'saved', id: 2, refused: 'boom' },
    ]);
  });

  test('written calls the page saved only when no edit came after the export; otherwise changed is posted again', async () => {
    const fake = fakeApi();
    const h = make(fake.api);
    const page = fakePage();
    h.attach(page);
    await tick();
    h.changed();
    send({ type: 'save', id: 1, name: 'a.docx' });
    await tick();
    send({ type: 'written', id: 1, name: 'a.docx' });
    // An edit between `saved` and `written`.
    send({ type: 'save', id: 2, name: 'a.docx' });
    await tick();
    h.changed();
    const before = ofType(fake.posted, 'changed').length;
    send({ type: 'written', id: 2, name: 'a.docx' });
    // Posted at once, the interval notwithstanding: after VS Code has cleared its dirty state.
    expect(ofType(fake.posted, 'changed').length).toBe(before + 1);
    // A written the page never exported for is not clean either.
    send({ type: 'written', id: 9, name: 'a.docx' });
    const written = page.calls.filter(([name]) => name === 'written').map(([, args]) => args);
    expect(written).toEqual([['a.docx', true], ['a.docx', false], ['a.docx', false]]);
    expect(ofType(fake.posted, 'saved').map((m) => m.id)).toEqual([1, 2]);
  });

  test('after a written naming .xml, the next save asks the page for that name, which decides Flat OPC', async () => {
    const h = make(fakeApi().api);
    const page = fakePage();
    h.attach(page);
    await tick();
    send({ type: 'save', id: 1, name: 'a.xml' });
    await tick();
    send({ type: 'written', id: 1, name: 'a.xml' });
    send({ type: 'save', id: 2, name: 'a.xml' });
    await tick();
    expect(page.calls.filter(([name]) => name === 'save').map(([, args]) => args)).toEqual([['a.xml'], ['a.xml']]);
    expect(page.calls.find(([name]) => name === 'written')?.[1]).toEqual(['a.xml', true]);
  });

  test('a backup asked during a save in flight is answered at once, never queued behind the save', async () => {
    const fake = fakeApi();
    const h = make(fake.api);
    let finish = (): void => undefined;
    h.attach(fakePage({ save: () => new Promise((resolve) => { finish = () => resolve({ bytes: new Uint8Array([1]) }); }) }));
    await tick();
    send({ type: 'save', id: 1, name: 'a.docx' });
    send({ type: 'backup', id: 2 });
    expect(fake.posted.filter((m) => m.type === 'saved' || m.type === 'backup').map((m) => m.type)).toEqual(['backup']);
    finish();
    await tick();
    expect(fake.posted.filter((m) => m.type === 'saved' || m.type === 'backup').map((m) => m.type)).toEqual(['backup', 'saved']);
  });
});

describe('changed', () => {
  test('is held to its interval: the first at once, the rest as one at the interval\'s end', () => {
    vi.useFakeTimers();
    const fake = fakeApi();
    const h = make(fake.api);
    const count = (): number => ofType(fake.posted, 'changed').length;
    h.changed();
    expect(count()).toBe(1);
    for (let i = 0; i < 5; i++) h.changed();
    expect(count()).toBe(1);
    vi.advanceTimersByTime(CHANGED_INTERVAL - 1);
    expect(count()).toBe(1);
    vi.advanceTimersByTime(1);
    expect(count()).toBe(2);
    vi.advanceTimersByTime(CHANGED_INTERVAL * 4);
    expect(count()).toBe(2);
    h.changed();
    expect(count()).toBe(3);
  });

  test('a clean written spends a changed the interval was holding, which would otherwise mark VS Code\'s document dirty again', async () => {
    vi.useFakeTimers();
    const fake = fakeApi();
    const h = make(fake.api);
    const page = fakePage();
    h.attach(page);
    await vi.advanceTimersByTimeAsync(0);
    const count = (): number => ofType(fake.posted, 'changed').length;
    // Two keys within the interval: the first posted, the second held.
    h.changed();
    h.changed();
    expect(count()).toBe(1);
    send({ type: 'save', id: 1, name: 'a.docx' });
    await vi.advanceTimersByTimeAsync(0);
    send({ type: 'written', id: 1, name: 'a.docx' });
    expect(page.calls.find(([name]) => name === 'written')?.[1]).toEqual(['a.docx', true]);
    await vi.advanceTimersByTimeAsync(CHANGED_INTERVAL * 4);
    expect(count()).toBe(1);
    // The next edit is posted at once.
    h.changed();
    expect(count()).toBe(2);
  });
});

describe('the options (P19)', () => {
  const boot = (options: Record<string, string>): void => {
    const meta = document.createElement('meta');
    meta.name = HOST_META;
    meta.content = JSON.stringify({ options });
    document.head.appendChild(meta);
  };

  test('are filled from the boot data, the host\'s copy the truth for every shared key, and the fill posts nothing', async () => {
    localStorage.setItem('docx4j-editor.relay', 'wss://old.example.org');
    localStorage.setItem('docx4j-editor.session', 'kept-per-page');
    localStorage.setItem('docx4j-editor.ask.key', 'sk-page');
    boot({ 'docx4j-editor.name': 'Ann', 'docx4j-editor.zoom': '1.25', 'docx4j-editor.ask.key': 'sk-host', 'elsewhere': 'x' });
    const fake = fakeApi();
    make(fake.api);
    await tick();
    expect(localStorage.getItem('docx4j-editor.name')).toBe('Ann');
    expect(localStorage.getItem('docx4j-editor.zoom')).toBe('1.25');
    expect(localStorage.getItem('docx4j-editor.relay')).toBeNull();
    expect(localStorage.getItem('docx4j-editor.session')).toBe('kept-per-page');
    // Never shared, so never filled from the boot; and the page never holds the Ask box's key now (ED-014 P45): one typed
    // into S1 to S5's Options is removed at the host's start, and nothing is posted for it.
    expect(localStorage.getItem('docx4j-editor.ask.key')).toBeNull();
    expect(localStorage.getItem('elsewhere')).toBeNull();
    expect(ofType(fake.posted, 'options')).toEqual([]);
  });

  test('the page\'s writes of a shared key are posted, coalesced; the excluded keys, other keys and unchanged values are not', async () => {
    boot({ 'docx4j-editor.zoom': '1' });
    const fake = fakeApi();
    make(fake.api);
    localStorage.setItem('docx4j-editor.name', 'Bo');
    localStorage.setItem('docx4j-editor.name', 'Bea');
    localStorage.setItem('docx4j-editor.zoom', '1');
    localStorage.removeItem('docx4j-editor.zoom');
    localStorage.removeItem('docx4j-editor.never-set');
    localStorage.setItem('docx4j-editor.session', 's');
    localStorage.setItem('docx4j-editor.editCauses', '[]');
    localStorage.setItem('docx4j-editor.ask.key', 'sk');
    localStorage.setItem('another.app', 'x');
    sessionStorage.setItem('docx4j-editor.name', 'Session');
    await tick();
    expect(ofType(fake.posted, 'options')).toEqual([{ type: 'options', set: { 'docx4j-editor.name': 'Bea' }, remove: ['docx4j-editor.zoom'] }]);
  });

  test('another editor\'s change reaches the storage, and is not posted back; stop puts the storage back', async () => {
    const fake = fakeApi();
    const h = make(fake.api);
    localStorage.setItem('docx4j-editor.name', 'Ann');
    await tick();
    fake.posted.length = 0;
    send({ type: 'options', set: { 'docx4j-editor.layout': 'full', 'docx4j-editor.ask.key': 'nope' }, remove: ['docx4j-editor.name'] });
    await tick();
    expect(localStorage.getItem('docx4j-editor.layout')).toBe('full');
    expect(localStorage.getItem('docx4j-editor.name')).toBeNull();
    expect(localStorage.getItem('docx4j-editor.ask.key')).toBeNull();
    expect(fake.posted).toEqual([]);
    h.stop();
    localStorage.setItem('docx4j-editor.name', 'After');
    await tick();
    expect(fake.posted).toEqual([]);
  });
});

describe('S3\'s additions (ED-014 section 4.3)', () => {
  const bootWith = (data: Record<string, unknown>): void => {
    const meta = document.createElement('meta');
    meta.name = HOST_META;
    meta.content = JSON.stringify({ options: {}, ...data });
    document.head.appendChild(meta);
  };

  test('the boot names VS Code\'s language and the trust: documentMacros only in a trusted workspace (P27)', () => {
    bootWith({ language: 'fr', trusted: true });
    const trusted = make(fakeApi().api);
    expect(trusted.language()).toBe('fr');
    expect(trusted.has('documentMacros')).toBe(true);
    trusted.stop();
    document.head.querySelectorAll(`meta[name="${HOST_META}"]`).forEach((m) => m.remove());
    bootWith({ language: 'de', trusted: false });
    const untrusted = make(fakeApi().api);
    expect(untrusted.language()).toBe('de');
    expect(untrusted.has('documentMacros')).toBe(false);
    untrusted.stop();
    document.head.querySelectorAll(`meta[name="${HOST_META}"]`).forEach((m) => m.remove());
    // No boot data: no language, no macros.
    const none = make(fakeApi().api);
    expect(none.language()).toBeUndefined();
    expect(none.has('documentMacros')).toBe(false);
    // The browser offers them, and names no language of its own.
    expect(browserHost().has('documentMacros')).toBe(true);
    expect(browserHost().language()).toBeUndefined();
  });

  test('a hand-over with no edit after the export: the page stops and gives its place, and no store goes with it (P25)', async () => {
    const fake = fakeApi();
    const h = make(fake.api);
    const page = fakePage();
    h.attach(page);
    await tick();
    send({ type: 'save', id: 1, name: 'b.docx' });
    await tick();
    send({ type: 'handoff', id: 1 });
    expect(page.calls.filter(([name]) => name === 'handoff')).toHaveLength(1);
    expect(page.calls.filter(([name]) => name === 'backup')).toHaveLength(0);
    expect(ofType(fake.posted, 'handoff')).toEqual([{ type: 'handoff', id: 1, place: { at: 'here' }, late: false }]);
  });

  test('a hand-over with an edit after the export carries the store as a backup\'s update and meta (section 9.8 finding 1)', async () => {
    const fake = fakeApi();
    const h = make(fake.api);
    const page = fakePage();
    h.attach(page);
    await tick();
    send({ type: 'save', id: 1, name: 'b.docx' });
    await tick();
    h.changed();
    send({ type: 'handoff', id: 1 });
    expect(ofType(fake.posted, 'handoff')).toEqual([{ type: 'handoff', id: 1, place: { at: 'here' }, resume: { update: new Uint8Array([9]), meta: { name: 'a.docx' } }, late: true }]);
    // The changed the edit armed is spent: the editor is closing.
    const changed = ofType(fake.posted, 'changed').length;
    await new Promise((r) => setTimeout(r, CHANGED_INTERVAL * 2));
    expect(ofType(fake.posted, 'changed').length).toBe(changed);
  });

  test('a cancelled save forgets its export: a written under that id afterwards is not clean', async () => {
    const h = make(fakeApi().api);
    const page = fakePage();
    h.attach(page);
    await tick();
    send({ type: 'save', id: 1, name: 'a.docx' });
    await tick();
    send({ type: 'saveCancelled', id: 1 });
    // The page says nothing of a cancelled save.
    expect(page.calls.map(([name]) => name)).toEqual(['save']);
    send({ type: 'written', id: 1, name: 'a.docx' });
    expect(page.calls.find(([name]) => name === 'written')?.[1]).toEqual(['a.docx', false]);
  });

  test('a save, a disk change and a hand-over that come while the page opens wait for the open', async () => {
    let finish = (): void => undefined;
    const h = make(fakeApi().api);
    const page = fakePage({ open: () => new Promise<void>((resolve) => { finish = resolve; }) });
    h.attach(page);
    await tick();
    send({ type: 'open', name: 'a.docx', bytes: new Uint8Array([1]), place: { caret: 2 } });
    send({ type: 'save', id: 1, name: 'a.docx' });
    send({ type: 'disk', bytes: new Uint8Array([2]), lastModified: 1 });
    send({ type: 'backup', id: 2 });
    await tick();
    expect(page.calls.map(([name]) => name)).toEqual(['open', 'backup']);
    expect(page.calls[0]![1][1]).toEqual({ place: { caret: 2 } });
    finish();
    await tick();
    expect(page.calls.map(([name]) => name)).toEqual(['open', 'backup', 'save', 'diskChanged']);
  });
});

describe('S5\'s additions (ED-014 section 4.5)', () => {
  const bootWith = (data: Record<string, unknown>): void => {
    const meta = document.createElement('meta');
    meta.name = HOST_META;
    meta.content = JSON.stringify({ options: {}, ...data });
    document.head.appendChild(meta);
  };
  const link = { room: 'r'.repeat(16), token: 'tok', relay: 'ws://127.0.0.1:22642' };

  test('the companion where the boot names the agent, none otherwise; the browser finds its own, and keeps the options (P41)', () => {
    expect(make(fakeApi().api).companion()).toBeUndefined();
    bootWith({ agent: 'Agent' });
    expect(make(fakeApi().api).companion()?.agent).toBe('Agent');
    expect(browserHost().companion()).toBeUndefined();
    expect(browserHost().has('companionOptions')).toBe(true);
    expect(make(fakeApi().api).has('companionOptions')).toBe(false);
  });

  test('start asks the extension, and its answer settles it at once, ahead of an open the queue holds; the page hears the address too', async () => {
    bootWith({ agent: 'Agent' });
    const fake = fakeApi();
    const h = make(fake.api);
    let finishOpen!: () => void;
    const page = fakePage({ open: () => new Promise<void>((r) => { finishOpen = r; }) });
    h.attach(page);
    await tick();
    send({ type: 'open', name: 'a.docx', bytes: new Uint8Array([1]) });
    await tick();
    const started = h.companion()!.start();
    const asked = ofType(fake.posted, 'companionStart');
    expect(asked).toEqual([{ type: 'companionStart', id: 1 }]);
    send({ type: 'companion', id: 1, url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
    await expect(started).resolves.toEqual({ url: 'http://127.0.0.1:24402', relay: 'ws://127.0.0.1:22642' });
    expect(h.siteAddress()).toBe('http://127.0.0.1:24402/');
    // The page hears it once the open is done, in order.
    expect(page.calls.filter(([n]) => n === 'setCompanion')).toHaveLength(0);
    finishOpen();
    await tick();
    await tick();
    expect(page.calls.find(([n]) => n === 'setCompanion')![1]).toEqual(['http://127.0.0.1:24402', 'ws://127.0.0.1:22642']);
  });

  test('a start refused rejects with the extension\'s words, and the page hears nothing', async () => {
    bootWith({ agent: 'Agent' });
    const fake = fakeApi();
    const h = make(fake.api);
    const page = fakePage();
    h.attach(page);
    await tick();
    const started = h.companion()!.start();
    send({ type: 'companion', id: 1, error: 'The agent\'s companion is not started in Restricted Mode.' });
    await expect(started).rejects.toThrow('Restricted Mode');
    await tick();
    expect(page.calls.filter(([n]) => n === 'setCompanion')).toHaveLength(0);
  });

  test('invite posts the room and its token; invited settles it, or rejects with why', async () => {
    bootWith({ agent: 'Agent' });
    const fake = fakeApi();
    const h = make(fake.api);
    h.attach(fakePage());
    await tick();
    const ok = h.companion()!.invite(link);
    expect(ofType(fake.posted, 'invite')).toEqual([{ type: 'invite', id: 1, room: link.room, token: 'tok' }]);
    send({ type: 'invited', id: 1 });
    await expect(ok).resolves.toBeUndefined();
    const refused = h.companion()!.invite(link);
    send({ type: 'invited', id: 2, error: 'wrong token for this room' });
    await expect(refused).rejects.toThrow('wrong token for this room');
    // A stray answer is no message for the page.
    send({ type: 'invited', id: 99 });
  });

  test('room: posted when the session enters a room and when it leaves, once per change', () => {
    const fake = fakeApi();
    const h = make(fake.api);
    h.room(undefined);
    h.room(link.relay);
    h.room(link.relay);
    h.room(undefined);
    expect(ofType(fake.posted, 'room')).toEqual([{ type: 'room', relay: link.relay }, { type: 'room' }]);
  });

  test('a hand-over in a room carries the store with no edit after the export, late false (P42); open with clean opens the resume as saved', async () => {
    const fake = fakeApi();
    const h = make(fake.api);
    const page = fakePage({ backup: () => ({ update: new Uint8Array([9]), meta: { name: 'a.docx', room: link } }) });
    h.attach(page);
    await tick();
    h.room(link.relay);
    send({ type: 'save', id: 1, name: 'b.docx' });
    await tick();
    send({ type: 'handoff', id: 1 });
    expect(ofType(fake.posted, 'handoff')).toEqual([{ type: 'handoff', id: 1, place: { at: 'here' }, resume: { update: new Uint8Array([9]), meta: { name: 'a.docx', room: link } }, late: false }]);
    // The target's page: a resume that opens saved, in its room.
    send({ type: 'open', name: 'b.docx', bytes: new Uint8Array([1]), resume: { update: new Uint8Array([9]), meta: { name: 'b.docx', room: link } }, clean: true });
    await tick();
    const [, how] = page.calls.find(([n]) => n === 'open')![1] as [unknown, { clean?: boolean; resume?: unknown }];
    expect(how.clean).toBe(true);
    expect(how.resume).toBeDefined();
    // clean without a resume means nothing.
    send({ type: 'open', name: 'c.docx', bytes: new Uint8Array([1]), clean: true });
    await tick();
    expect(page.calls.filter(([n]) => n === 'open').at(-1)![1][1]).toBeUndefined();
  });
});

describe('S6\'s additions, the Ask box (ED-014 section 4.6)', () => {
  const bootWith = (data: Record<string, unknown>): void => {
    const meta = document.createElement('meta');
    meta.name = HOST_META;
    meta.content = JSON.stringify({ options: {}, ...data });
    document.head.appendChild(meta);
  };
  const model = { vendor: 'docx4j-test', id: 'claude-test', version: '1', name: 'Test Claude', family: 'claude-test', maxInputTokens: 100_000 };
  const request = { model: 'docx4j-test/claude-test', system: 'S', prompt: 'P', maxTokens: 4096 };

  test('the browser sends its own, and keeps the key and address in Options; VS Code\'s host sends, the options hidden (P45)', () => {
    expect(browserHost().ask()).toBeUndefined();
    expect(browserHost().has('askOptions')).toBe(true);
    const h = make(fakeApi().api);
    expect(h.ask()).toBeDefined();
    expect(h.ask()).toBe(h.ask());
    expect(h.has('askOptions')).toBe(false);
  });

  test('whether a key is stored: the boot\'s askKey, then askKeyState, which waits behind an open', async () => {
    bootWith({ askKey: true });
    const fake = fakeApi();
    const h = make(fake.api);
    const ask = h.ask()!;
    expect(ask.keyStored()).toBe(true);
    const heard: boolean[] = [];
    const off = ask.onKeyStored((stored) => heard.push(stored));
    let finishOpen!: () => void;
    h.attach(fakePage({ open: () => new Promise<void>((r) => { finishOpen = r; }) }));
    await tick();
    send({ type: 'open', name: 'a.docx', bytes: new Uint8Array([1]) });
    await tick();
    send({ type: 'askKeyState', stored: false });
    await tick();
    // A push waits behind the open, as every message does.
    expect(heard).toEqual([]);
    expect(ask.keyStored()).toBe(true);
    finishOpen();
    await tick();
    await tick();
    expect(heard).toEqual([false]);
    expect(ask.keyStored()).toBe(false);
    off();
    send({ type: 'askKeyState', stored: true });
    await tick();
    expect(heard).toEqual([false]);
    expect(ask.keyStored()).toBe(true);
    // No boot key: not stored.
    document.head.querySelectorAll(`meta[name="${HOST_META}"]`).forEach((m) => m.remove());
    expect(make(fakeApi().api).ask()!.keyStored()).toBe(false);
  });

  test('the models: asked once, answered ahead of an open the queue holds; the list kept; asked again when fresh; pushes heard', async () => {
    const fake = fakeApi();
    const h = make(fake.api);
    let finishOpen!: () => void;
    h.attach(fakePage({ open: () => new Promise<void>((r) => { finishOpen = r; }) }));
    await tick();
    send({ type: 'open', name: 'a.docx', bytes: new Uint8Array([1]) });
    await tick();
    const ask = h.ask()!;
    const heard: unknown[] = [];
    ask.onModels((list) => heard.push(list));
    const first = ask.models();
    const again = ask.models();
    expect(ofType(fake.posted, 'askModels')).toEqual([{ type: 'askModels', id: 1 }]);
    send({ type: 'askModels', id: 1, models: [model] });
    await expect(first).resolves.toEqual([model]);
    await expect(again).resolves.toEqual([model]);
    // Kept: no second question.
    await expect(ask.models()).resolves.toEqual([model]);
    expect(ofType(fake.posted, 'askModels')).toHaveLength(1);
    // Fresh asks again.
    const fresh = ask.models(true);
    expect(ofType(fake.posted, 'askModels')).toEqual([{ type: 'askModels', id: 1 }, { type: 'askModels', id: 2 }]);
    send({ type: 'askModels', id: 2, models: [] });
    await expect(fresh).resolves.toEqual([]);
    // A push (no id) waits behind the open.
    send({ type: 'askModels', models: [model, { ...model, vendor: 'copilot' }] });
    await tick();
    // Each answer heard once (the two first callers shared one question), and the push not yet.
    expect(heard).toEqual([[model], []]);
    finishOpen();
    await tick();
    await tick();
    expect(heard).toEqual([[model], [], [model, { ...model, vendor: 'copilot' }]]);
    await expect(ask.models()).resolves.toEqual([model, { ...model, vendor: 'copilot' }]);
  });

  test('send posts the route and the request; asked settles it ahead of an open the queue holds: an answer, or a failure by kind', async () => {
    const fake = fakeApi();
    const h = make(fake.api);
    h.attach(fakePage({ open: () => new Promise<void>(() => undefined) }));
    await tick();
    send({ type: 'open', name: 'a.docx', bytes: new Uint8Array([1]) });
    await tick();
    const ask = h.ask()!;
    const answered = ask.send('vscode', request, new AbortController().signal);
    expect(ofType(fake.posted, 'ask')).toEqual([{ type: 'ask', id: 1, route: 'vscode', request }]);
    send({ type: 'asked', id: 1, answer: { text: 'T', model: 'Test Claude', stopReason: null } });
    await expect(answered).resolves.toEqual({ text: 'T', model: 'Test Claude', stopReason: null });
    const failed = ask.send('key', { ...request, model: 'claude-sonnet-5' }, new AbortController().signal);
    send({ type: 'asked', id: 2, failure: { kind: 'tooLong', message: '9 > 1', tokens: 9, limit: 1 } });
    const e = await failed.then(() => undefined, (x: unknown) => x);
    expect(e).toBeInstanceOf(AskFailure);
    expect((e as AskFailure).failure).toEqual({ kind: 'tooLong', message: '9 > 1', tokens: 9, limit: 1 });
    const cancelled = ask.send('vscode', request, new AbortController().signal);
    send({ type: 'asked', id: 3, cancelled: true });
    await expect(cancelled).rejects.toMatchObject({ name: 'AbortError' });
  });

  test('a Cancel deletes the request\'s entry, then posts askCancel; a late asked is dropped', async () => {
    const fake = fakeApi();
    const h = make(fake.api);
    const page = fakePage();
    h.attach(page);
    await tick();
    const controller = new AbortController();
    const pending = h.ask()!.send('key', request, controller.signal);
    controller.abort();
    await expect(pending).rejects.toMatchObject({ name: 'AbortError' });
    expect(ofType(fake.posted, 'askCancel')).toEqual([{ type: 'askCancel', id: 1 }]);
    send({ type: 'asked', id: 1, answer: { text: 'late', model: 'm', stopReason: null } });
    await tick();
    // Nothing reached the page, and nothing was thrown.
    expect(page.calls).toEqual([]);
    // A signal aborted before the send posts nothing.
    const gone = new AbortController();
    gone.abort();
    await expect(h.ask()!.send('key', request, gone.signal)).rejects.toMatchObject({ name: 'AbortError' });
    expect(ofType(fake.posted, 'ask')).toHaveLength(1);
  });

  test('pushes that come before the application attaches are kept for it: the key\'s state and the models (section 9.23)', async () => {
    bootWith({ askKey: true });
    const h = make(fakeApi().api);
    const ask = h.ask()!;
    const keys: boolean[] = [];
    const lists: unknown[] = [];
    ask.onKeyStored((stored) => keys.push(stored));
    ask.onModels((list) => lists.push(list));
    send({ type: 'askKeyState', stored: false });
    send({ type: 'askModels', models: [model] });
    await tick();
    expect(keys).toEqual([]);
    expect(lists).toEqual([]);
    h.attach(fakePage());
    await tick();
    await tick();
    expect(keys).toEqual([false]);
    expect(lists).toEqual([[model]]);
    await expect(ask.models()).resolves.toEqual([model]);
  });

  test('a refused key is worded by where the key is set: Options in a browser, the command in VS Code (section 9.23 finding 1)', () => {
    expect(describeFailure({ kind: 'key', message: 'refused', status: 401 })).toMatch(/Options/);
    (globalThis as { acquireVsCodeApi?: unknown }).acquireVsCodeApi = fakeApi().acquire;
    expect(initHost().kind).toBe('vscode');
    const words = describeFailure({ kind: 'key', message: 'refused', status: 401 });
    expect(words).toContain('Set the Anthropic API Key');
    expect(words).not.toMatch(/Options/);
  });

  test('the Options dialog\'s buttons post askKey with their action', () => {
    const fake = fakeApi();
    const h = make(fake.api);
    for (const action of ['set', 'remove', 'settings'] as const) h.ask()!.key(action);
    expect(ofType(fake.posted, 'askKey')).toEqual([{ type: 'askKey', action: 'set' }, { type: 'askKey', action: 'remove' }, { type: 'askKey', action: 'settings' }]);
  });

  test('a key a page held before S6 is removed at the host\'s start, and nothing is posted for it', async () => {
    localStorage.setItem('docx4j-editor.ask.key', 'sk-old');
    localStorage.setItem('docx4j-editor.ask.model', 'claude-opus-5');
    const fake = fakeApi();
    make(fake.api);
    await tick();
    expect(localStorage.getItem('docx4j-editor.ask.key')).toBeNull();
    expect(localStorage.getItem('docx4j-editor.ask.model')).toBe('claude-opus-5');
    expect(fake.posted).toEqual([]);
  });
});

describe('S7\'s additions, the palette\'s names (ED-014 section 4.7, P49)', () => {
  const bootWith = (data: Record<string, unknown>): void => {
    const meta = document.createElement('meta');
    meta.name = HOST_META;
    meta.content = JSON.stringify({ options: {}, ...data });
    document.head.appendChild(meta);
  };
  const FRENCH = 'docx4j editor: Définir la clé API Anthropic…';

  test('the boot\'s name for the key\'s command reaches the three messages that name it, in VS Code\'s language, not the page\'s', () => {
    // VS Code in French, the page in English (a language chosen in Options): the palette's entry is the French one.
    bootWith({ commands: { 'docx4j.ask.setKey': FRENCH, 'docx4j.join': 7 } });
    (globalThis as { acquireVsCodeApi?: unknown }).acquireVsCodeApi = fakeApi().acquire;
    const h = initHost();
    expect(h.commandName('docx4j.ask.setKey')).toBe(FRENCH);
    // A name that is not a string is not one.
    expect(h.commandName('docx4j.join')).toBeUndefined();
    expect(setKeyCommand()).toBe(FRENCH);
    expect(describeFailure({ kind: 'key', message: 'refused', status: 401 })).toBe(`The provider refused the key (401): set it again with "${FRENCH}"`);
    expect(describeFailure({ kind: 'noKey', message: 'none' })).toBe(`No Anthropic API key is set: set one with "${FRENCH}"`);
    expect(t('ask.noKey.vscode', { command: setKeyCommand() })).toContain(`("${FRENCH}")`);
  });

  test('a boot that names none, and a browser: the English name', () => {
    bootWith({});
    (globalThis as { acquireVsCodeApi?: unknown }).acquireVsCodeApi = fakeApi().acquire;
    expect(initHost().commandName('docx4j.ask.setKey')).toBeUndefined();
    expect(setKeyCommand()).toBe(PAGE_COMMANDS['docx4j.ask.setKey']);
    expect(describeFailure({ kind: 'noKey', message: 'none' })).toContain('"docx4j editor: Set the Anthropic API Key…"');
    resetHost();
    expect(browserHost().commandName('docx4j.ask.setKey')).toBeUndefined();
  });
});
