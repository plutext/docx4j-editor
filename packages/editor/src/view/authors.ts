// SPDX-License-Identifier: GPL-3.0-or-later
//
// An author's colour (ED-005 section 6.3): one hue per name, the same in the text's tracked
// changes, the comments margin and the People list, so an author, or the agent of E4.f, keeps one
// colour everywhere; and a pattern per name beside it, so that authors differ by more than colour
// (section 9.5). The body view owns a stylesheet into which one rule per author is written the
// first time a mark view or a node view meets the name: `[data-author="…"] { --author-hue: N }`,
// from which `editor.css` draws every revision. Nothing is a decoration, and a keystroke costs a
// set lookup when a mark view is made.

/** A hue from the author's name, so one author keeps one colour across documents. */
export function authorHue(author: string): number {
  let h = 0;
  for (let i = 0; i < author.length; i++) h = (h * 31 + author.charCodeAt(i)) >>> 0;
  return h % 360;
}

export function authorColor(author: string): string {
  return `hsl(${authorHue(author)} 65% 42%)`;
}

/**
 * An author's pattern (ED-005 section 9.5, accessibility): a hatch at an angle of its own, so that
 * authors differ by more than colour. None is a line style the kinds use (an underline, a
 * strikethrough, a doubled or dotted line), so no horizontal hatch. Chosen by a hash apart from the
 * hue's, so that two authors of near hues are unlikely to share a pattern too.
 */
const PATTERNS: readonly ((line: string) => string)[] = [
  (c) => `repeating-linear-gradient(45deg, ${c} 0 1px, transparent 1px 5px)`,
  (c) => `repeating-linear-gradient(-45deg, ${c} 0 1px, transparent 1px 5px)`,
  (c) => `repeating-linear-gradient(90deg, ${c} 0 1px, transparent 1px 4px)`,
  (c) => `repeating-linear-gradient(45deg, ${c} 0 1px, transparent 1px 6px), repeating-linear-gradient(-45deg, ${c} 0 1px, transparent 1px 6px)`,
];

export function authorPatternIndex(author: string): number {
  let h = 2166136261;
  for (let i = 0; i < author.length; i++) h = Math.imul(h ^ author.charCodeAt(i), 16777619) >>> 0;
  return h % PATTERNS.length;
}

/** The author's pattern as a CSS image, its lines in `line` (the author's hue, faint, behind text; white over a swatch of the colour). */
export function authorPattern(author: string, line = `hsl(${authorHue(author)} 65% 42% / 0.22)`): string {
  return PATTERNS[authorPatternIndex(author)]!(line);
}

/** A swatch beside an author's name: the colour with the pattern over it in white, to match the text's by. */
export function authorSwatch(author: string): { backgroundColor: string; backgroundImage: string } {
  return { backgroundColor: authorColor(author), backgroundImage: authorPattern(author, 'rgb(255 255 255 / 0.75)') };
}

/**
 * A string as a CSS string literal: the backslash, the quote and line breaks escaped, which is
 * all a quoted attribute value needs. Not `CSS.escape`, which jsdom does not have (ED-005
 * section 6.3).
 */
export function cssString(s: string): string {
  return `"${s.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/[\r\n]/g, ' ')}"`;
}

/** The authors met so far and their rules, in a `<style>` element in the document's head. */
export class AuthorStyles {
  private readonly seen = new Set<string>();
  readonly element: HTMLStyleElement;

  constructor(owner: Document) {
    this.element = owner.createElement('style');
    this.element.dataset['docx4jEditor'] = 'authors';
    owner.head.appendChild(this.element);
  }

  /** The rule for an author, written once. */
  note(author: string | null | undefined): void {
    const name = author ?? '';
    if (this.seen.has(name)) return;
    this.seen.add(name);
    this.element.append(`.docx-editor [data-author=${cssString(name)}] { --author-hue: ${authorHue(name)}; --author-pattern: ${authorPattern(name)}; }\n`);
  }

  /** How many authors have a rule (tests). */
  get size(): number {
    return this.seen.size;
  }

  destroy(): void {
    this.element.remove();
  }
}
