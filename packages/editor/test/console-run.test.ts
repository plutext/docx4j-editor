// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-002 section 6.6 step D3: the wrapper keeps line numbers, TypeScript is stripped, and an
// error's line maps back to the user's code.
import { describe, test, expect } from 'vitest';
import { wrap, compile, lineOf, BINDINGS } from '../src/app/console-run.js';

describe('console runner', () => {
  test('the wrapper takes the bindings as parameters and keeps the user code line for line', () => {
    const code = 'const a: number = 1;\nreturn a + 1;';
    const wrapped = wrap(code);
    expect(wrapped.split('\n')[0]).toBe(`export default async ({ ${BINDINGS.join(', ')} }) => {`);
    expect(wrapped.split('\n')[1]).toBe('const a: number = 1;');
    const js = compile(code);
    expect(js).not.toContain(': number');
    expect(js.split('\n')[1]).toBe('const a = 1;');
    expect(js.split('\n')[2]).toBe('return a + 1;');
  });

  test('a syntax error is thrown by the compiler with its line', () => {
    expect(() => compile('const x = ;')).toThrow();
  });

  test("an error's line is the user's, the wrapper's line subtracted", () => {
    const url = 'blob:http://localhost/abc-123';
    const error = new Error('boom');
    error.stack = `Error: boom\n    at default (${url}:3:9)\n    at run (app.js:1:1)`;
    expect(lineOf(error, url)).toBe(2);
    expect(lineOf(new Error('no stack'), url)).toBeUndefined();
  });
});
