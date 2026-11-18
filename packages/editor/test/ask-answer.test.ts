// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-003 section 7.5: the three fenced forms and the explanation.
import { describe, test, expect } from 'vitest';
import { parseAnswer } from '../src/app/ask-answer.js';

describe('the ask answer', () => {
  test('a wml block is content, with the prose around it as the note', () => {
    const form = parseAnswer('Here is the clause.\n```wml\n<w:p><w:r><w:t>Hello</w:t></w:r></w:p>\n```\nInsert it after the heading.');
    expect(form).toEqual({ kind: 'wml', xml: '<w:p><w:r><w:t>Hello</w:t></w:r></w:p>', note: 'Here is the clause.\n\nInsert it after the heading.' });
  });

  test('an xml tag counts as wml; text and ts are the other two forms; a trailing newline is dropped', () => {
    expect(parseAnswer('```xml\n<w:p/>\n```').kind).toBe('wml');
    expect(parseAnswer('```text\nOne.\nTwo.\n```')).toEqual({ kind: 'text', text: 'One.\nTwo.', note: '' });
    expect(parseAnswer('```ts\nreturn body.paragraphs.length;\n```')).toEqual({ kind: 'script', code: 'return body.paragraphs.length;', note: '' });
    expect(parseAnswer('```typescript\nx\n```').kind).toBe('script');
    expect(parseAnswer('```javascript\nx\n```').kind).toBe('script');
  });

  test('prose with no fence, or a fence in another language, is an explanation', () => {
    expect(parseAnswer('  The w:jc element sets the alignment.  ')).toEqual({ kind: 'explanation', text: 'The w:jc element sets the alignment.' });
    expect(parseAnswer('```python\nprint(1)\n```').kind).toBe('explanation');
  });

  test('the first block decides when there are several', () => {
    expect(parseAnswer('```text\na\n```\n```ts\nb\n```').kind).toBe('text');
  });
});
