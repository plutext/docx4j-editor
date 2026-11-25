// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-013 section 3.8 (P8): Send a copy's hand-over to the Web Share API, over a stubbed `navigator` (jsdom has none).
import { afterEach, describe, test, expect } from 'vitest';
import { canShareFiles, isChromium, sendFile, DOCX_TYPE } from '../src/app/files.js';

type Stub = {
  canShare?: (data: { files: File[] }) => boolean;
  share?: (data: { files: File[] }) => Promise<void>;
  userActivation?: { isActive: boolean };
  userAgentData?: { brands: { brand: string }[] };
  userAgent?: string;
};

const SAFARI = 'Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Mobile/15E148 Safari/604.1';
const CHROME_ANDROID = 'Mozilla/5.0 (Linux; Android 10; K) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/157.0.0.0 Mobile Safari/537.36';
const CHROME_IOS = 'Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) CriOS/157.0.0.0 Mobile/15E148 Safari/604.1';

function stub(fields: Stub): void {
  for (const [key, value] of Object.entries(fields)) Object.defineProperty(navigator, key, { configurable: true, value });
}

afterEach(() => {
  for (const key of ['canShare', 'share', 'userActivation', 'userAgentData', 'userAgent']) delete (navigator as unknown as Record<string, unknown>)[key];
});

const docx = (): File => new File([new Uint8Array([80, 75])], 'a.docx', { type: DOCX_TYPE });
const xml = (): File => new File([new Uint8Array([60])], 'a.xml', { type: 'application/xml' });
const refusing = (name: string) => async (): Promise<void> => { throw new DOMException('refused', name); };

describe('isChromium', () => {
  test("by the brands where the browser states them, else by its user agent; Chrome on iOS is WebKit's", () => {
    expect(isChromium({ userAgentData: { brands: [{ brand: 'Not A;Brand' }, { brand: 'Chromium' }, { brand: 'Google Chrome' }] }, userAgent: SAFARI } as never)).toBe(true);
    expect(isChromium({ userAgentData: { brands: [{ brand: 'Other' }] }, userAgent: CHROME_ANDROID } as never)).toBe(false);
    expect(isChromium({ userAgent: CHROME_ANDROID } as never)).toBe(true);
    expect(isChromium({ userAgent: SAFARI } as never)).toBe(false);
    expect(isChromium({ userAgent: CHROME_IOS } as never)).toBe(false);
  });
});

describe('canShareFiles', () => {
  test('false without the API, or with share alone', () => {
    expect(canShareFiles()).toBe(false);
    stub({ share: async () => {} });
    expect(canShareFiles()).toBe(false);
  });

  test("the browser's answer for a .docx, false when canShare throws", () => {
    const asked: File[] = [];
    stub({ share: async () => {}, canShare: (d) => { asked.push(...d.files); return true; } });
    expect(canShareFiles()).toBe(true);
    expect(asked.map((f) => [f.name, f.type])).toEqual([['document.docx', DOCX_TYPE]]);
    stub({ canShare: () => { throw new TypeError('no'); } });
    expect(canShareFiles()).toBe(false);
  });

  test("never in Chromium, whose canShare says yes to a .docx its share refuses (UPSTREAM-28)", () => {
    stub({ share: async () => {}, canShare: () => true, userAgentData: { brands: [{ brand: 'Chromium' }] } });
    expect(canShareFiles()).toBe(false);
    stub({ userAgentData: undefined as never, userAgent: CHROME_ANDROID });
    expect(canShareFiles()).toBe(false);
    stub({ userAgent: SAFARI });
    expect(canShareFiles()).toBe(true);
  });
});

describe('sendFile', () => {
  test('sent: the file handed to share as it is', async () => {
    const given: File[] = [];
    stub({ canShare: () => true, share: async (d) => { given.push(...d.files); } });
    const file = docx();
    expect(await sendFile(file)).toEqual({ kind: 'sent' });
    expect(given).toEqual([file]);
  });

  test('unsupported: no share, or a file canShare refuses (Flat OPC)', async () => {
    expect(await sendFile(docx())).toEqual({ kind: 'unsupported' });
    stub({ canShare: (d) => d.files.every((f) => f.name.endsWith('.docx')), share: async () => {} });
    expect(await sendFile(xml())).toEqual({ kind: 'unsupported' });
  });

  test('cancelled: the share sheet closed without a choice', async () => {
    stub({ canShare: () => true, share: refusing('AbortError') });
    expect(await sendFile(docx())).toEqual({ kind: 'cancelled' });
  });

  test('ready: no activation left, before any share; the second tap shares', async () => {
    let calls = 0;
    stub({ canShare: () => true, share: async () => { calls++; }, userActivation: { isActive: false } });
    const file = docx();
    expect(await sendFile(file)).toEqual({ kind: 'ready', file });
    expect(calls).toBe(0);
    expect(await sendFile(file, true)).toEqual({ kind: 'sent' });
    expect(calls).toBe(1);
  });

  test("refused: NotAllowedError with the tap's activation there, as Chromium refuses a .docx, or on the second tap", async () => {
    const file = docx();
    stub({ canShare: () => true, share: refusing('NotAllowedError'), userActivation: { isActive: true } });
    expect(await sendFile(file)).toEqual({ kind: 'refused' });
    // A browser that cannot say whether the activation was there: the first refusal is taken as the activation spent.
    delete (navigator as unknown as Record<string, unknown>)['userActivation'];
    expect(await sendFile(file)).toEqual({ kind: 'ready', file });
    expect(await sendFile(file, true)).toEqual({ kind: 'refused' });
  });

  test('another refusal is thrown', async () => {
    stub({ canShare: () => true, share: refusing('DataError') });
    await expect(sendFile(docx())).rejects.toMatchObject({ name: 'DataError' });
  });
});
