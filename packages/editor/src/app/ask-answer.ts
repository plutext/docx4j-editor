// SPDX-License-Identifier: GPL-3.0-or-later
//
// The ask mode's answer (ED-001 section 3.5; ED-003 section 7.5): the model answers in one of
// three fenced forms and the console knows what to do with each. The first fenced block decides:
// a `wml` (or `xml`) block is content to preview and insert; a `text` block is plain paragraphs;
// a `ts` (or `js`) block is a script for the console editor, never run; anything else is an
// explanation. The prose around a block is kept as its note.

export type AskForm =
  | { kind: 'wml'; xml: string; note: string }
  | { kind: 'text'; text: string; note: string }
  | { kind: 'script'; code: string; note: string }
  | { kind: 'explanation'; text: string };

const FENCE = /```([\w-]*)[ \t]*\n([\s\S]*?)```/;

export function parseAnswer(answer: string): AskForm {
  const m = FENCE.exec(answer);
  if (!m) return { kind: 'explanation', text: answer.trim() };
  const lang = (m[1] ?? '').toLowerCase();
  const body = m[2] ?? '';
  const note = `${answer.slice(0, m.index)}${answer.slice(m.index + m[0].length)}`.trim();
  if (lang === 'wml' || lang === 'xml') return { kind: 'wml', xml: body.trim(), note };
  if (lang === 'text' || lang === 'txt' || lang === 'plain' || lang === 'plaintext') return { kind: 'text', text: body.replace(/\n$/, ''), note };
  if (lang === 'ts' || lang === 'typescript' || lang === 'js' || lang === 'javascript') return { kind: 'script', code: body.replace(/\n$/, ''), note };
  return { kind: 'explanation', text: answer.trim() };
}

/** The contract the system prompt states, in the words the model is given. */
export const FENCED_FORMS = `Answer in exactly one of these forms:
1. Content to put in the document, as a WordprocessingML fragment: a fenced block tagged wml holding one or more block-level elements (w:p, w:tbl) with the w: prefix, styles by the document's style ids (w:pStyle w:val="Heading1"), no w:document or w:body wrapper. Use this for anything that goes into the document.
2. Plain text to put in the document: a fenced block tagged text, one paragraph per line, when no formatting is wanted.
3. A script for the console: a fenced block tagged ts, TypeScript over the bindings pkg, body, main, selection, el, wml and Word (the content API in Office JS shapes, as declared below). The editor shows it; the user reads and runs it. Never both a script and content.
4. An explanation: prose with no fenced block, for a question that changes nothing.
A short sentence before or after the block is fine; nothing else outside it.`;
