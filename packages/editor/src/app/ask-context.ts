// SPDX-License-Identifier: GPL-3.0-or-later
//
// The ask mode's context builders (ED-001 section 3.5; ED-003 section 7.5), reused by E4: what
// the console already has, handed to the model with the question. The selection's XML and text
// (`exportBlock` over the selected blocks, marshalled), the outline (the headings of the document
// on screen, by style, from the projected document rather than the tree), the styles in use with
// their names, and the API declarations (the content API's classes and the shim's index, raw and
// capped). The request built from them is shown before it is sent (ED-001 section 11 item 5).
import { documentStyles, paragraphText, acceptedTextBetween, type EditorDocument, type StyleEntry, type SourceIds } from '@docx4j/editor-model';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorState } from 'prosemirror-state';
import type { EditorSurface } from '../view/surface.js';
import { stylesInUse } from '../view/css.js';
import { blocksIn, xmlOf } from './xml.js';
import { FENCED_FORMS } from './ask-answer.js';
import type { AskRequestData } from './host-protocol.js';
import { chatPrompt } from './ask-send.js';

/** The request as one message for a chat, or for a VS Code model, which has no system role (ED-014 section 4.6). */
export { chatPrompt };

export interface AskContext {
  selectionXml: string;
  selectionText: string;
  /** How many top-level blocks the selection touches. */
  blocks: number;
  outline: string;
  styles: string;
  declarations: string;
  /** True when the declarations were cut at the cap. */
  declarationsCapped: boolean;
}

export type AskRequest = AskRequestData;

/** The declarations the model reads: the content API's classes an answer's script or fragment would use, and the shim. */
const declarationFiles = import.meta.glob([
  '../../../../node_modules/@docx4j/core-ts/dist/model/content/Body.d.mts',
  '../../../../node_modules/@docx4j/core-ts/dist/model/content/Paragraph.d.mts',
  '../../../../node_modules/@docx4j/core-ts/dist/model/content/Range.d.mts',
  '../../../../node_modules/@docx4j/core-ts/dist/model/content/Font.d.mts',
  '../../../../node_modules/@docx4j/core-ts/dist/model/content/Table.d.mts',
  '../../../../node_modules/@docx4j/core-ts/dist/model/content/ContentControl.d.mts',
  '../../../../node_modules/@docx4j/core-ts/dist/office-js/index.d.mts',
], { query: '?raw', import: 'default' }) as Record<string, () => Promise<string>>;

/** The declarations' cap in characters (about 15,000 tokens): enough for the classes above, and said to be a cap when it cuts. */
export const DECLARATIONS_CAP = 60_000;

let declarationsCache: Promise<{ text: string; capped: boolean }> | undefined;

export function apiDeclarations(): Promise<{ text: string; capped: boolean }> {
  declarationsCache ??= (async () => {
    const parts: string[] = [];
    for (const [path, load] of Object.entries(declarationFiles).sort(([a], [b]) => a.localeCompare(b))) {
      const name = path.slice(path.lastIndexOf('/dist/') + 6);
      // Comments are the bulk of a declaration file; the signatures are what the model needs.
      const text = (await load()).replace(/\/\*\*[\s\S]*?\*\//g, '').replace(/^\s*\n/gm, '');
      parts.push(`// ${name}\n${text}`);
    }
    const whole = parts.join('\n');
    return whole.length > DECLARATIONS_CAP ? { text: whole.slice(0, DECLARATIONS_CAP), capped: true } : { text: whole, capped: false };
  })();
  return declarationsCache;
}

/**
 * The headings of the document on screen, by their paragraph style (`Heading1` to `Heading9`, `Title`, `Subtitle`), indented by
 * level. Each heading's text is the accepted text, as the find bar reads it (ED-008 proposal P4): a deletion pending in the
 * document is not text the model is told the document holds.
 */
export function outlineOf(doc: PMNode, limit = 200, ids?: SourceIds): string {
  const lines: string[] = [];
  doc.descendants((node) => {
    if (lines.length >= limit) return false;
    if (node.type.name !== 'paragraph') return true;
    const style = (node.attrs['pStyle'] as string | null) ?? '';
    const m = /^Heading(\d)$/.exec(style);
    const level = m ? Number(m[1]) : style === 'Title' ? 0 : style === 'Subtitle' ? 1 : -1;
    if (level < 0) return false;
    const text = paragraphText(node, 0, ids).text.trim();
    if (text !== '') lines.push(`${'  '.repeat(Math.max(0, level - 1))}${style}: ${text}`);
    return false;
  });
  return lines.join('\n');
}

/** The styles the model may name: the ones in use first, then the quick styles, as `id (name)` lines. */
export function stylesLine(entries: readonly StyleEntry[], inUse: { paragraph: Set<string>; character: Set<string> }, limit = 80): string {
  const used = entries.filter((e) => inUse.paragraph.has(e.id) || inUse.character.has(e.id));
  const quick = entries.filter((e) => e.quick && !used.includes(e));
  const rest = entries.filter((e) => !used.includes(e) && !quick.includes(e));
  const listed = [...used, ...quick, ...rest].slice(0, limit);
  return listed.map((e) => `${e.id} (${e.name}, ${e.type}${used.includes(e) ? ', in use' : ''})`).join('\n');
}

/** Everything the request carries besides the question. */
export async function buildContext(session: EditorDocument, surface: EditorSurface): Promise<AskContext> {
  const state: EditorState = surface.state;
  const blocks = blocksIn(state);
  const xmls = await Promise.all(blocks.nodes.map(async (n) => xmlOf(surface.exportBlock(n)).catch((e: Error) => `<!-- ${e.message} -->`)));
  // The accepted text (ED-008 proposal P4): what a reader of the document would read, deletions out.
  const selectionText = acceptedTextBetween(state.doc, state.selection.from, state.selection.to, session.ctx.ids);
  const [entries, declarations] = await Promise.all([documentStyles(session), apiDeclarations()]);
  return {
    selectionXml: xmls.join('\n'),
    selectionText,
    blocks: blocks.nodes.length,
    outline: outlineOf(state.doc, 200, session.ctx.ids),
    styles: stylesLine(entries, stylesInUse(session, state.doc)),
    declarations: declarations.text,
    declarationsCapped: declarations.capped,
  };
}

/** The system prompt: the contract, then the document's context, then the declarations (the stable part last so a provider can cache it). */
export function systemPrompt(ctx: AskContext): string {
  return [
    'You help edit a Word document in the docx4j editor. The user sees the document; you see the selection, the outline and the styles below.',
    FENCED_FORMS,
    `Document outline (headings by style):\n${ctx.outline === '' ? '(no headings)' : ctx.outline}`,
    `Styles (id (name, type)); use the ids in w:pStyle and w:rStyle:\n${ctx.styles}`,
    `API declarations for scripts (the subset of Office JS Word.* that @docx4j/core-ts implements)${ctx.declarationsCapped ? ', cut at the cap' : ''}:\n${ctx.declarations}`,
  ].join('\n\n');
}

/** The user message: the question, with the selection as text and XML. */
export function userMessage(prompt: string, ctx: AskContext): string {
  const selection = ctx.blocks === 0 ? '(nothing selected)' : `Selected text:\n${ctx.selectionText === '' ? '(a caret, no text)' : ctx.selectionText}\n\nSelected XML (${ctx.blocks} block${ctx.blocks === 1 ? '' : 's'}):\n${ctx.selectionXml}`;
  return `${prompt.trim()}\n\n${selection}`;
}

export function buildRequest(model: string, prompt: string, ctx: AskContext, maxTokens = 4096): AskRequest {
  return { model, system: systemPrompt(ctx), prompt: userMessage(prompt, ctx), maxTokens };
}
