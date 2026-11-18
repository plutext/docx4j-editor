// SPDX-License-Identifier: GPL-3.0-or-later
//
// core-ts check 34 (run by Jason in current Word, 2026-10-03; `fixtures/revisions/check34/`, copied from
// docx4j-core-ts's `test/fixtures/revisions/check34/`): Word's XML Mapping Pane, Insert Content Control > Plain Text
// on the node `name` of `<data><name>Ann</name><empty/></data>`, one case per paragraph of `34-input.docx`; Word's
// saves are `34a-mapped.docx` (tracking off), `34b-tracked.docx` (1j, tracking on) and `34b-tracked-reject-all.docx`
// (Reject All over the last). The Data tab's bind does each
// case through `contentControl.insertBound`, and each paragraph is held to Word's. Two cases are refused at Jason's
// word, where Word does something else (1f, 1g; docs/developer/word-divergences.md). ED-005 sections 12.56 and 12.57.
import { describe, expect, test } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, runCommandAsync, plainTextControlRefusal, controlToBind, stateFor, readControl, sdtBinding, PLACEHOLDER_STYLE } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

const PART = '{443D3445-BAA1-420E-A5E7-67CD0C9BC805}';
/** The binding as Word's pane writes it: the canonical path in the part. */
const NAME = { xpath: '/data[1]/name[1]', storeItemID: PART };
const EMPTY = { xpath: '/data[1]/empty[1]', storeItemID: PART };

/**
 * A block as text: a run control's content in brackets, a block control's as `BLOCK [..]`, bold as `*..*`, deleted
 * text as `-{..}` and inserted as `+{..}`, a placeholder as `{placeholder}` (Word's wording and the editor's differ: the
 * engine's), a picture as `<picture>`; markers left out. Without `formatting`, bold is not shown.
 */
function shown(node: PMNode, formatting = true): string {
  let out = '';
  node.forEach((c) => {
    if (c.isText) {
      const has = (name: string): boolean => c.marks.some((m) => m.type.name === name);
      if (c.marks.some((m) => m.type.name === 'runStyle' && m.attrs['id'] === PLACEHOLDER_STYLE)) { out += '{placeholder}'; return; }
      let text = c.text!;
      if (formatting && has('bold')) text = `*${text}*`;
      out += has('deletion') ? `-{${text}}` : has('insertion') ? `+{${text}}` : text;
    } else if (c.type.name === 'sdtInline') {
      out += `[${shown(c, formatting)}]`;
    } else if (c.type.name === 'sdtBlock') {
      out += `BLOCK [${shown(c, formatting)}]`;
    } else if (c.type.name === 'paragraph') {
      out += shown(c, formatting);
    } else if (c.type.name === 'image') {
      out += '<picture>';
    }
  });
  // Neighbouring runs of one kind read as one.
  for (let prev = ''; prev !== out;) { prev = out; out = out.replace(/([-+])\{([^{}]*)\}\1\{/g, '$1{$2'); }
  return out.replace(/\*\*/g, '');
}

/** A top-level block as text: a block control as `BLOCK [..]`. */
function block(node: PMNode, formatting = true): string {
  return node.type.name === 'sdtBlock' ? `BLOCK [${shown(node, formatting)}]` : shown(node, formatting);
}

function blockStart(doc: PMNode, i: number): number {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += doc.child(k).nodeSize;
  return pos;
}

/** The position just after `text`'s first occurrence in a block's text, inside run controls too. */
function after(doc: PMNode, block: number, text: string): number {
  const node = doc.child(block);
  const target = node.textContent.indexOf(text) + text.length;
  expect(target).toBeGreaterThanOrEqual(text.length);
  let seen = 0;
  let found = -1;
  node.descendants((n, p) => {
    if (found >= 0) return false;
    if (n.isText) {
      if (seen + n.text!.length >= target) found = blockStart(doc, block) + 1 + p + (target - seen);
      seen += n.text!.length;
    }
    return true;
  });
  return found;
}

/** The position at the end of a block's content. */
const end = (doc: PMNode, block: number): number => blockStart(doc, block) + doc.child(block).nodeSize - 1;

async function input(): Promise<EditorDocument> {
  const session = await EditorDocument.open(await fixture('revisions/check34/34-input.docx'));
  await session.template.load();
  return session;
}

async function word(name: '34a-mapped.docx' | '34b-tracked.docx'): Promise<PMNode> {
  return (await EditorDocument.open(await fixture(`revisions/check34/${name}`))).doc;
}

describe("core-ts check 34: the Data tab's bind does what Word's XML Mapping Pane does", () => {
  test('1a to 1e, 1h and 1i, tracking off: each paragraph as Word left it; the binding Word wrote', async () => {
    const session = await input();
    const bind = (selection: { from: number; to?: number }, binding = NAME): void => {
      expect(runCommand(session, 'contentControl.insertBound', binding, selection).ok).toBe(true);
    };
    bind({ from: after(session.doc, 1, '1a. The qu') });                                   // 1a: inside "quick"
    bind({ from: after(session.doc, 2, '1b. The quick') });                                // 1b: between "quick" and "brown"
    bind({ from: end(session.doc, 3) });                                                   // 1c: at the paragraph's end
    bind({ from: blockStart(session.doc, 5) + 1 });                                        // 1d: the empty paragraph
    bind({ from: after(session.doc, 6, '1e. The '), to: after(session.doc, 6, 'quick brown') });   // 1e: "quick brown"
    // 1h: inside the existing rich text control, a new control nested in it (Jason's word, ED-005 section 12.57).
    const inRich = { from: after(session.doc, 10, 'existing rich') };
    expect(controlToBind(stateFor(session, inRich))).toBeUndefined();
    bind(inRich);
    bind({ from: end(session.doc, 11) }, EMPTY);                                           // 1i: an empty node
    const mapped = await word('34a-mapped.docx');
    for (const [label, at] of [['1a', 1], ['1b', 2], ['1c', 3], ['1e', 6], ['1h', 10], ['1i', 11]] as const) {
      expect(`${label}: ${block(session.doc.child(at))}`).toBe(`${label}: ${block(mapped.child(at))}`);
    }
    // 1d, the empty paragraph: Word makes a block-level control around it; the editor a run-level one inside it, by
    // Jason's decision of 2026-10-08 (docs/developer/word-divergences.md): a block control holding a document's last paragraph
    // cannot be left by the caret. Word's result stays pinned.
    expect(block(mapped.child(5))).toBe('BLOCK [Ann]');
    expect(block(session.doc.child(5))).toBe('[Ann]');
    expect(shown(session.doc.child(10))).toBe('1h. An [existing rich[Ann] text control]: ');
    // Every control made is plain text, bound as Word binds it; 1i shows the placeholder, as Word's does.
    const made: PMNode[] = [];
    session.doc.descendants((n) => { if ((n.type.name === 'sdtInline' || n.type.name === 'sdtBlock') && sdtBinding(n.attrs['sdtPr'])) made.push(n); return true; });
    expect(made).toHaveLength(7);
    for (const c of made) {
      expect(readControl(c.attrs['sdtPr']).kind).toBe('PlainText');
      expect(sdtBinding(c.attrs['sdtPr'])).toMatchObject({ storeItemID: PART });
    }
    expect(readControl(made[6]!.attrs['sdtPr']).showingPlaceholder).toBe(true);
    // Saved and reopened (the save's self-check runs), the paragraphs read the same.
    const reopened = await EditorDocument.open((await session.save()).bytes);
    for (const at of [1, 2, 3, 5, 6, 10, 11]) expect(block(reopened.doc.child(at))).toBe(block(session.doc.child(at)));
  });

  test('1j, tracking on: the selected text deleted and the value inserted inside the control, as Word has them', async () => {
    const session = await input();
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Author A';
    const r = runCommand(session, 'contentControl.insertBound', NAME, { from: after(session.doc, 12, '1j. The '), to: after(session.doc, 12, 'quick brown') });
    expect(r.ok).toBe(true);
    const tracked = await word('34b-tracked.docx');
    expect(shown(session.doc.child(12), false)).toBe(shown(tracked.child(12), false));
    expect(shown(session.doc.child(12), false)).toBe('1j. The [-{quick brown}+{Ann}] fox.');
    // Word's plain text control gives its content one formatting, the deleted "brown" losing its bold; the editor keeps
    // it, which a reject puts back (docs/developer/word-divergences.md). Word marks the control itself inserted
    // (w:customXmlInsRangeStart), which the editor's model cannot hold: the control is not marked.
    expect(shown(tracked.child(12))).toBe('1j. The [-{quick brown}+{Ann}] fox.');
    expect(shown(session.doc.child(12))).toBe('1j. The [-{quick *brown*}+{Ann}] fox.');
    // Word's Reject All (`34b-tracked-reject-all.docx`, Jason, 2026-10-03) rejects the control's insertion too: the
    // control goes, the text is back without its bold, and the part still holds "Ann". The editor's reject puts the
    // text back in the control, which stays bound, and the text is then written to the node, as an edit in a bound
    // control is (docs/developer/word-divergences.md, open; ED-005 section 12.57).
    const rejected = await EditorDocument.open(await fixture('revisions/check34/34b-tracked-reject-all.docx'));
    await rejected.template.load();
    expect(shown(rejected.doc.child(12))).toBe('1j. The quick brown fox.');
    expect(rejected.template.partNamed(PART)!.selectSingleNode('/data/name')!.text).toBe('Ann');
    expect((await runCommandAsync(session, 'trackedChanges.rejectAll', {})).ok).toBe(true);
    expect(shown(session.doc.child(12))).toBe('1j. The [quick *brown*] fox.');
    expect(session.template.partNamed(PART)!.selectSingleNode('/data/name')!.text).toBe('quick brown');
  });

  test('1f and 1g, refused where Word does something else (Jason: keep refusing); Word\'s result pinned', async () => {
    const session = await input();
    // 1f: from "quick" in the first of two paragraphs to "brown" in the second. Word replaced the first paragraph's part.
    const across = { from: after(session.doc, 7, 'two: the '), to: after(session.doc, 8, 'brown') };
    expect(plainTextControlRefusal(stateFor(session, across))).toBe('paragraphs');
    expect(runCommand(session, 'contentControl.insertBound', NAME, across).ok).toBe(false);
    // 1g: "here: " through the picture. Word replaced the text and moved the picture to the paragraph's end.
    const picture = { from: after(session.doc, 9, 'A picture '), to: end(session.doc, 9) };
    expect(plainTextControlRefusal(stateFor(session, picture))).toBe('content');
    expect(runCommand(session, 'contentControl.insertBound', NAME, picture).ok).toBe(false);
    const mapped = await word('34a-mapped.docx');
    expect(shown(mapped.child(7))).toBe('1f. First paragraph of two: the [Ann]');
    expect(shown(mapped.child(8))).toBe('brown fox, second paragraph of two.');
    expect(shown(mapped.child(9))).toBe('1g. A picture [Ann] after it.<picture>');
  });
});
