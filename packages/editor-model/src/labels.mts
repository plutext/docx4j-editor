// SPDX-License-Identifier: GPL-3.0-or-later
//
// List labels for display (ED-002 section 5.3): what the Emulator (CR-001 Phase B) says each
// source paragraph's label is, keyed by the paragraph's `sourceId`. The engine's own
// `Body.listLabels()` walks unmarshalled parts only, so this walks the private tree the
// projection came from (the engine's story walk, over the body's content), which costs the
// numbering part nothing: the emulator reads its definitions privately. Labels come from the
// source elements, so a paragraph typed new in a list has none and a deletion's followers are
// stale until the next open; the list commands and live renumbering are E2.
import { isElement, selectedBranchOf, NumberingState, type Element, type ListLabel } from '@docx4j/core-ts';
import type * as mce from '@docx4j/generated-objects-ts/modules/org_docx4j_mce';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorDocument } from './session.mjs';
import type { IndentAttrs, ParagraphAttrs } from './schema.mjs';
import { numberingParts, numberingPPrOf, levelIndentOf, type Numbering } from './commands/lists.mjs';

export type { ListLabel };

/** The engine's story walk: every paragraph in order, a text box starting its own counters, a kept `mc:AlternateContent` by its selected branch. */
function walkStory(root: unknown, state: NumberingState, emit: (p: wml.P, state: NumberingState) => void): void {
  const walk = (value: unknown, current: NumberingState): void => {
    if (Array.isArray(value)) { for (const item of value) walk(item, current); return; }
    if (typeof value !== 'object' || value === null) return;
    if (isElement(value)) { walk((value as Element).value, current); return; }
    const node = value as Record<string, unknown> & { TYPE_NAME?: string };
    let inner = current;
    if (node['TYPE_NAME'] === 'org_docx4j_wml.P') emit(node as unknown as wml.P, current);
    else if (node['TYPE_NAME'] === 'org_docx4j_wml.CTTxbxContent') inner = new NumberingState();
    // A kept `mc:AlternateContent` (core-ts CR-010 phase 4): the branch the engine reads, not every branch's paragraphs.
    else if (node['TYPE_NAME'] === 'org_docx4j_mce.AlternateContent') { walk(selectedBranchOf(node as mce.AlternateContent) ?? [], inner); return; }
    for (const key of Object.keys(node)) {
      if (key === 'PARENT' || key === 'TYPE_NAME') continue;
      walk(node[key], inner);
    }
  };
  walk(root, state);
}

/**
 * Word's bullet library lives in the private-use range of the Symbol and Wingdings faces (docx4j's
 * default bullet is Symbol F0B7), which no web font has: the label is shown as the Unicode
 * character the glyph is, and any other private-use character as a bullet, until E2.f draws the
 * label in the level's own font.
 */
const SYMBOL_GLYPHS: Record<string, string> = {
  '\uF0B7': '•', '\uF0A7': '▪', '\uF0A8': '◦', '\uF0D8': '➢', '\uF0FC': '✓', '\uF076': '❖', '\uF0A4': '◊', '\uF0AD': '✱', '\uF06E': '■', '\uF0B2': '◆', '\uF0BE': '➤',
};

export function displayLabel(numString: string): string {
  return numString.replace(/[\uE000-\uF8FF]/g, (ch) => SYMBOL_GLYPHS[ch] ?? '•');
}

/** The labels of the body's list paragraphs by `sourceId`; empty when the document has no numbering. */
export function listLabels(session: EditorDocument): Map<string, ListLabel> {
  const out = new Map<string, ListLabel>();
  let emulator;
  try {
    emulator = session.main.numberingDefinitionsPart?.getEmulator();
  } catch {
    return out;
  }
  if (!emulator) return out;
  const byElement = new Map<object, ListLabel>();
  walkStory(session.body.content ?? [], new NumberingState(), (p, counters) => {
    const result = emulator.getNumber(p.pPr, counters);
    if (result === undefined || result.numString === undefined) return;
    byElement.set(p, {
      listString: displayLabel(result.numString),
      level: Number(result.ilvl ?? '0'),
      numId: result.numId ?? '',
      siblingIndex: (result.count ?? 1) - (result.level?.start ?? 1),
      isBullet: result.isBullet,
    });
  });
  if (byElement.size === 0) return out;
  session.doc.descendants((node) => {
    if (node.type.name !== 'paragraph') return true;
    const id = node.attrs['sourceId'] as string | null;
    const el = id ? session.ctx.ids.get(id) : undefined;
    const label = el ? byElement.get(el.value as object) : undefined;
    if (label && id) out.set(id, label);
    return false;
  });
  return out;
}

// ---------------------------------------------------------------------------------------------
// Live labels over the editor's own document (ED-003 section 5.4, step C3)
// ---------------------------------------------------------------------------------------------

/** The level's `w:suff`: what separates the label from the text. */
export type LabelSuffix = 'tab' | 'space' | 'nothing';

/** A label with the geometry its widget and its paragraph lay out with. */
export interface LiveLabel extends ListLabel {
  /** The indent the paragraph lays out with: the level's with the paragraph's own over it (`levelIndentOf`), or null when the level states none. */
  indent: IndentAttrs | null;
  suff: LabelSuffix;
}

/**
 * The labels of the document's list paragraphs, keyed by node, computed over the nodes rather
 * than the tree: a walk of the paragraphs in story order (inside tables and controls too; an
 * opaque block is not entered, so a text box starts nothing here) builds for each the only
 * `w:pPr` the emulator's `numRefFor` reads, `{ pStyle, numPr }` from the attributes, and asks
 * `getNumber` with one `NumberingState` for the story, exactly as `listLabels` does over the
 * tree. A typed paragraph is numbered like any other, and a deletion's followers renumber.
 * Empty when the document has no numbering part. `numbering` is the session's unless given (the
 * preview's clone has its own, with the lists an XHTML conversion made).
 */
export function liveLabels(doc: PMNode, session: EditorDocument, numbering: Numbering | undefined = numberingParts(session)): Map<PMNode, LiveLabel> {
  const out = new Map<PMNode, LiveLabel>();
  if (!numbering) return out;
  const state = new NumberingState();
  doc.descendants((node) => {
    if (node.type.name === 'opaqueBlock') return false;
    if (node.type.name !== 'paragraph') return true;
    const attrs = node.attrs as ParagraphAttrs;
    const result = numbering.emulator.getNumber(numberingPPrOf(attrs), state);
    if (result === undefined || result.numString === undefined) return false;
    const suff = result.level?.suff;
    out.set(node, {
      listString: displayLabel(result.numString),
      level: Number(result.ilvl ?? '0'),
      numId: result.numId ?? '',
      siblingIndex: (result.count ?? 1) - (result.level?.start ?? 1),
      isBullet: result.isBullet,
      indent: levelIndentOf(attrs, numbering),
      suff: suff === 'space' || suff === 'nothing' ? suff : 'tab',
    });
    return false;
  });
  return out;
}
