// SPDX-License-Identifier: GPL-3.0-or-later
//
// The companion probe's address: the option when set, else the companion that served the page, which
// marks the page it serves (`<meta name="docx4j-companion">`). A page without the mark, a static site
// such as the demo on GitHub Pages, probes nothing: there a probe would be a 404 on every open.
import { describe, test, expect } from 'vitest';
import { effectiveCompanion, servedCompanion, COMPANION_META } from '../src/app/companion.js';

describe('effectiveCompanion', () => {
  test('the option wins, trimmed, marked page or not', () => {
    expect(effectiveCompanion(' http://127.0.0.1:4174/companion ', 'https://example.org/editor/', '')).toBe('http://127.0.0.1:4174/companion');
    expect(effectiveCompanion('http://elsewhere:4174/companion', 'http://127.0.0.1:4174/', '/companion')).toBe('http://elsewhere:4174/companion');
  });

  test('a page the companion served probes where its mark says, whatever the page\'s query and hash', () => {
    expect(effectiveCompanion('', 'http://127.0.0.1:4174/?x=1#room=abc', '/companion')).toBe('http://127.0.0.1:4174/companion');
  });

  test('an unmarked page probes nothing: a static site at the root or under a subpath', () => {
    expect(effectiveCompanion('', 'http://localhost:4173/', '')).toBe('');
    expect(effectiveCompanion('', 'https://plutext.github.io/docx4j-editor/', '')).toBe('');
  });

  test('a page not served over HTTP probes nothing', () => {
    expect(effectiveCompanion('', 'file:///tmp/index.html', '/companion')).toBe('');
    expect(effectiveCompanion('', 'not a url', '/companion')).toBe('');
  });
});

describe('servedCompanion', () => {
  test('reads the mark from the page, and nothing from a page without one', () => {
    const doc = document.implementation.createHTMLDocument('t');
    expect(servedCompanion(doc)).toBe('');
    const meta = doc.createElement('meta');
    meta.name = COMPANION_META;
    meta.content = ' /companion ';
    doc.head.append(meta);
    expect(servedCompanion(doc)).toBe('/companion');
  });

  test('this test page has no mark, so the default probes nothing', () => {
    expect(effectiveCompanion('', 'http://localhost:4173/')).toBe('');
  });
});
