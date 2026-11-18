// SPDX-License-Identifier: GPL-3.0-or-later
//
// Markdown in (docs/developer/agent.md section 2.6; ED-005 section 8.4): the model's reading
// (`@docx4j/editor-model`'s `readMarkdown`, which paste and Open share since ED-005 section 12.66),
// held to the agent's subset. A construct outside it is `markdown.unsupported`, naming the
// construct: `insert_ooxml` is the escape hatch.
import { readMarkdown, MarkdownUnsupported, type MdBlock } from '@docx4j/editor-model';
import { ToolError } from './errors.mjs';

export { buildBlocks, plainBlocks, inlineText, type BuildContext, type Inline, type ListPlace, type MdBlock } from '@docx4j/editor-model';

/** Markdown as the subset's blocks; `markdown.unsupported` for anything outside it. */
export function parseMarkdown(markdown: string): MdBlock[] {
  try {
    return readMarkdown(markdown).blocks;
  } catch (e) {
    if (!(e instanceof MarkdownUnsupported)) throw e;
    throw new ToolError('markdown.unsupported', e.message, 'use paragraphs, # headings, **bold**, *italic*, ~~strike~~, lists, GFM tables and [links](https://...); insert_ooxml takes anything else', { construct: e.construct });
  }
}
