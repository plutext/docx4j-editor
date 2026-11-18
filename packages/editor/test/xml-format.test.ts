// SPDX-License-Identifier: GPL-3.0-or-later
//
// The parts panel's Format (formatXml): indentation where an element holds only elements, text left exactly as it is.
import { describe, expect, test } from 'vitest';
import { formatXml, elementPathAt } from '../src/app/xml.js';

describe('formatXml', () => {
  test('indents element-only content, keeps the declaration, and leaves text, mixed content and xml:space alone', () => {
    const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="urn:w"><w:style w:styleId="A"><w:name w:val="a"/><w:rPr><w:b/></w:rPr></w:style>'
      + '<w:p><w:r><w:t>Hello</w:t></w:r><w:r><w:t xml:space="preserve"> two  spaces </w:t></w:r></w:p>'
      + '<data><name>Ann <b>bold</b> tail</name><empty></empty><blank>  </blank></data></w:styles>';
    const out = formatXml(xml);
    expect(out.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles')).toBe(true);
    expect(out).toContain('\n  <w:style w:styleId="A">\n    <w:name w:val="a"/>\n    <w:rPr>\n      <w:b/>\n    </w:rPr>\n  </w:style>');
    expect(out).toContain('<w:t>Hello</w:t>');
    expect(out).toContain('<w:t xml:space="preserve"> two  spaces </w:t>');
    expect(out).toContain('<name>Ann <b>bold</b> tail</name>');
    expect(out).toContain('<blank>  </blank>');
    // The same document, whitespace aside.
    const strip = (s: string): string => s.replace(/>\s+</g, '><');
    expect(strip(out)).toBe(strip(xml.replace('<empty></empty>', '<empty/>')));
    // Formatting twice changes nothing.
    expect(formatXml(out)).toBe(out);
  });

  test('a text that does not parse throws', () => {
    expect(() => formatXml('<a><b></a>')).toThrow();
  });
});

describe('elementPathAt', () => {
  const xml = '<?xml version="1.0"?>\n<a xmlns="urn:x">\n  <!-- note -->\n  <b id="1" name=\'two > three\'>Bob</b>\n  <c/>\n  <b>Bea<d>x</d></b>\n</a>';
  const at = (needle: string, plus = 0): number => xml.indexOf(needle) + plus;
  test('the element an offset is in or on, by position among element children, and the attribute it is on', () => {
    expect(elementPathAt(xml, at('<a'))).toEqual({ path: [1] });
    expect(elementPathAt(xml, at('Bob', 1))).toEqual({ path: [1, 1] });
    expect(elementPathAt(xml, at('id="1"', 4))).toEqual({ path: [1, 1], attribute: 'id' });
    expect(elementPathAt(xml, at('three'))).toEqual({ path: [1, 1], attribute: 'name' });
    expect(elementPathAt(xml, at('</b>', 2))).toEqual({ path: [1, 1] });
    expect(elementPathAt(xml, at('<c/>', 1))).toEqual({ path: [1, 2] });
    expect(elementPathAt(xml, at('Bea'))).toEqual({ path: [1, 3] });
    expect(elementPathAt(xml, at('>x<', 1))).toEqual({ path: [1, 3, 1] });
    expect(elementPathAt(xml, at('note'))).toEqual({ path: [1] });
    expect(elementPathAt(xml, at('version'))).toBeUndefined();
    expect(elementPathAt(xml, xml.length)).toBeUndefined();
  });
});
