// SPDX-License-Identifier: GPL-3.0-or-later
//
// The editing tools (docs/developer/agent.md section 3.3). Each is one tree run, tracked as the agent: blocks
// replaced, added and removed whole (`edit`, `insert`, `delete`, `insert_ooxml`), text inside a
// paragraph through `replace_text`, `set_format` and `delete` with a span, a paragraph's style and
// properties through `set_style`; and a comment, never tracked (rule 1). Everything a call needs is
// checked before anything changes, and each returns the change report.
import { randomInt } from 'node:crypto';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { Paragraph, Table, p as paragraphOf, r as runOf, typeNameOf, type Comment, type Element } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { ToolError } from './errors.mjs';
import { addressSchema, checkIfText, describeAddress, hasSpan, isParagraph, isTable, kindOf, resolveBlock, resolveParagraph, resolveSpan, spanShape, tableOf, type AddressInput } from './address.mjs';
import { buildBlocks, parseMarkdown, plainBlocks, type MdBlock } from './markdown.mjs';
import { answer, buildContext, counted, documentArg, dryRunArg, ifTextArg, inRoom, insertAt, LOCATIONS, minute, mutate, nextStyleOf, placeFor, removeReplaced, reportResult, searchOptions, searchShape, styleIdFor, viewsOf, type Edit, type Location, type Place, type ToolHost } from './kit.mjs';

/** The style a plain paragraph takes at a place (section 2.6): the replaced paragraph's; next to one, the style Word's Enter would give (`w:next`); Normal elsewhere. */
async function plainStyleAt(edit: Edit, place: Place, location: Location): Promise<string | undefined> {
  if (place.kind !== 'block' || !place.paragraph) return undefined;
  const style = place.paragraph.p.pPr?.pStyle?.val;
  return location === 'Replace' ? style : nextStyleOf(edit, style);
}

/** "after body/3", "at the end of the body", for a summary. */
function where(place: Place, location: Location): string {
  if (place.kind === 'container') return `at the ${location === 'Start' ? 'start' : 'end'} of ${place.label}`;
  return `${location === 'Before' ? 'before' : 'after'} ${place.address}`;
}

/** The content of a mutating call that adds blocks: Markdown or plain text, parsed before anything runs. */
function blocksOf(markdown: string | undefined, text: string | undefined): MdBlock[] {
  if ((markdown === undefined) === (text === undefined)) throw new ToolError('arguments.invalid', 'give markdown or text, one of them', 'markdown is the subset of section 2.6; text is plain, a paragraph per line', { argument: 'markdown' });
  const blocks = markdown !== undefined ? parseMarkdown(markdown) : plainBlocks(text!);
  if (blocks.length === 0) throw new ToolError('arguments.invalid', 'there is nothing to insert', 'give the paragraphs as markdown or text', { argument: 'markdown' });
  return blocks;
}

/** A comment's content: plain paragraphs with bold, italic and strikethrough; nothing a comment part cannot hold on its own. */
function commentContent(markdown: string): Element<wml.P>[] {
  const out: Element<wml.P>[] = [];
  for (const block of parseMarkdown(markdown)) {
    if (block.kind !== 'paragraph' || block.heading !== undefined || block.list !== undefined) {
      throw new ToolError('markdown.unsupported', 'a comment holds paragraphs of text: no headings, lists or tables', 'write the comment as plain paragraphs', { construct: block.kind === 'table' ? 'a table in a comment' : block.heading !== undefined ? 'a heading in a comment' : 'a list in a comment' });
    }
    if (block.inlines.some((i) => i.link !== undefined)) throw new ToolError('markdown.unsupported', 'a link in a comment is not supported', 'write the address as text', { construct: 'a link in a comment' });
    out.push(paragraphOf(block.inlines.filter((i) => i.text !== '').map((i) => runOf(i.text, { ...(i.bold ? { bold: true } : {}), ...(i.italic ? { italic: true } : {}), ...(i.strike ? { strikeThrough: true } : {}) }))));
  }
  if (out.length === 0) throw new ToolError('arguments.invalid', 'the comment is empty', 'give its text', { argument: 'text' });
  return out;
}


/** The comment parts the package has now, for the report (rule 1: untracked). */
function commentParts(edit: Edit): string[] {
  const main = edit.main as unknown as Record<string, { partName?: { toString(): string } } | undefined>;
  return ['commentsPart', 'commentsExtendedPart', 'commentsIdsPart', 'commentsExtensiblePart', 'peoplePart']
    .map((k) => main[k]?.partName?.toString())
    .filter((n): n is string => n !== undefined);
}

/** Every comment of a list and their replies, depth first. */
function allComments(list: readonly Comment[]): Comment[] {
  return list.flatMap((c) => [c, ...allComments(c.replies)]);
}

const editShape = {
  document: documentArg,
  address: addressSchema.optional(),
  markdown: z.string().optional().describe('paragraphs, # headings, **bold**, *italic*, ~~strike~~, - and 1. lists, GFM tables, [links](https://...)'),
  text: z.string().optional().describe('plain text, a paragraph per line, in place of markdown'),
  if_text: ifTextArg,
  dry_run: dryRunArg,
};

export function registerEditTools(server: McpServer, host: ToolHost): void {
  const addBlocks = async (tool: 'edit' | 'insert', args: { document: string; address?: AddressInput | undefined; markdown?: string | undefined; text?: string | undefined; location?: Location | undefined; if_text?: string | undefined; dry_run?: boolean | undefined }) => {
    const blocks = blocksOf(args.markdown, args.text);
    const location = args.location ?? (tool === 'edit' ? 'Replace' : 'After');
    const allowed: readonly Location[] = tool === 'edit' ? LOCATIONS : ['Before', 'After', 'Start', 'End'];
    return reportResult(await inRoom(host, args.document, (room) => mutate(host, room, args.dry_run ?? false, 'TrackAll', async (edit) => {
      const place = await placeFor(edit, args.address, location, allowed, args.if_text);
      const elements = await buildBlocks(blocks, buildContext(edit, await plainStyleAt(edit, place, location)));
      const made = insertAt(edit, place, location, elements);
      for (const v of made) edit.report.create(v);
      if (location === 'Replace') removeReplaced(edit, place, elements);
      const summary = location === 'Replace' && place.kind === 'block' ? `Replaced ${place.address} with ${counted(made)}` : `Inserted ${counted(made)} ${where(place, location)}`;
      return edit.report.finish(summary);
    })));
  };

  server.registerTool('edit', {
    title: 'Edit blocks',
    description: 'Replaces a paragraph or table with Markdown (or plain text), or adds it Before or After a block, or at the Start or End of a container (the body without an address, a table cell, a content control). Whole paragraphs: to change words inside one, use replace_text with the paragraph as scope. Tracked as the agent\'s insertions and deletions: a replaced block stays at its address, deleted, until the person resolves it, and the new blocks follow it (the report\'s created addresses).',
    inputSchema: { ...editShape, location: z.enum(LOCATIONS).optional().describe('default Replace') },
  }, (args) => answer(() => addBlocks('edit', args)));

  server.registerTool('insert', {
    title: 'Insert blocks',
    description: 'Adds Markdown (or plain text) After a block (the default) or Before it, or at the Start or End of a container. A GFM table inserts a table. Tracked as the agent\'s insertions.',
    // Every location in the schema, so that one insert does not take is the contract's location.invalid, not a schema error.
    inputSchema: { ...editShape, location: z.enum(LOCATIONS).optional().describe('Before, After (the default), Start or End; not Replace') },
  }, (args) => answer(() => addBlocks('insert', args)));

  server.registerTool('delete', {
    title: 'Delete',
    description: 'Deletes a paragraph or table, or, with a span (anchor, or offset and length), the span\'s text in a paragraph. Tracked as the agent\'s deletion; the agent\'s own pending insertion goes for real.',
    inputSchema: { document: documentArg, address: addressSchema, ...spanShape, if_text: ifTextArg, dry_run: dryRunArg },
  }, (args) => answer(async () => reportResult(await inRoom(host, args.document, (room) => mutate(host, room, args.dry_run ?? false, 'TrackAll', async (edit) => {
    const { block, warnings } = resolveBlock(edit.body, args.address);
    for (const w of warnings) edit.report.warn(w);
    const at = edit.body.addressOf(block) ?? describeAddress(args.address);
    if (isParagraph(block.element)) {
      const paragraph = new Paragraph(block.element as Element<wml.P>, block.container, edit.body);
      await checkIfText(edit.body, paragraph, args.if_text);
      if (hasSpan(args)) {
        const range = resolveSpan(paragraph, args, args.address);
        const text = range.text;
        if (text === '') throw new ToolError('range.invalid', 'the span is empty', 'give the words, or an offset and a length above 0', { address: args.address });
        edit.report.touch(paragraph);
        range.delete();
        return edit.report.finish(`Deleted ${JSON.stringify(text)} in ${at}`);
      }
      edit.report.remove(paragraph);
      paragraph.delete();
      return edit.report.finish(`Deleted the paragraph ${at}`);
    }
    if (hasSpan(args)) throw new ToolError('arguments.invalid', `a span is part of a paragraph, and ${at} is ${kindOf(block.element)}`, 'address one of its paragraphs', { argument: 'anchor' });
    if (args.if_text !== undefined) throw new ToolError('arguments.invalid', `if_text names a paragraph, and ${at} is ${kindOf(block.element)}`, 'drop if_text', { argument: 'if_text' });
    if (isTable(block.element)) {
      const table = tableOf(edit.body, block);
      edit.report.remove(table);
      table.delete();
      return edit.report.finish(`Deleted the table ${at}`);
    }
    throw new ToolError('target.invalid', `delete takes a paragraph or a table, and ${at} is ${kindOf(block.element)}`, 'remove a content control with remove_content_control; delete a row\'s or a cell\'s paragraphs', { address: args.address });
  })))));

  server.registerTool('replace_text', {
    title: 'Replace text',
    description: 'Replaces every match of find with replace in scope (an address: a paragraph, a table, a cell or a content control; the whole document by default), as Word\'s Replace All: all matches or none, one change. $1 in replace names a group with match_regexp. Paragraph ids survive. if_text applies when the scope is one paragraph.',
    inputSchema: {
      document: documentArg,
      find: z.string().min(1),
      replace: z.string(),
      scope: addressSchema.optional(),
      ...searchShape,
      if_text: ifTextArg,
      dry_run: dryRunArg,
    },
  }, (args) => answer(async () => {
    const options = searchOptions(args);
    return reportResult(await inRoom(host, args.document, (room) => mutate(host, room, args.dry_run ?? false, 'TrackAll', async (edit) => {
      let paragraphs: Paragraph[];
      if (args.scope === undefined) {
        if (args.if_text !== undefined) throw new ToolError('arguments.invalid', 'if_text names a paragraph, and the scope is the whole document', 'give the paragraph as scope, or drop if_text', { argument: 'if_text' });
        paragraphs = edit.body.paragraphs;
      } else {
        const { block, warnings } = resolveBlock(edit.body, args.scope);
        for (const w of warnings) edit.report.warn(w);
        if (isParagraph(block.element)) {
          const paragraph = new Paragraph(block.element as Element<wml.P>, block.container, edit.body);
          await checkIfText(edit.body, paragraph, args.if_text);
          paragraphs = [paragraph];
        } else {
          if (args.if_text !== undefined) throw new ToolError('arguments.invalid', `if_text names a paragraph, and the scope is ${kindOf(block.element)}`, 'drop if_text', { argument: 'if_text' });
          const at = edit.body.addressOf(block) ?? describeAddress(args.scope);
          paragraphs = edit.body.sub({ content: [block.element] }, at).paragraphs.map((q) => new Paragraph(q.element, q.container, edit.body));
        }
      }
      // Every match first, so a pattern that does not parse refuses before anything changes.
      const matched: { paragraph: Paragraph; n: number }[] = [];
      for (const paragraph of paragraphs) {
        let n: number;
        try { n = paragraph.searchMatches(args.find, options).length; } catch (e) { throw new ToolError('arguments.invalid', `the search does not parse: ${(e as Error).message}`, 'check the pattern', { argument: 'find' }); }
        if (n > 0) matched.push({ paragraph, n });
      }
      for (const { paragraph } of matched) {
        edit.report.touch(paragraph);
        paragraph.replaceText(args.find, args.replace, options);
      }
      const total = matched.reduce((s, m) => s + m.n, 0);
      const summary = total === 0
        ? `No match for ${JSON.stringify(args.find)}`
        : `Replaced ${total === 1 ? 'one match' : `${total} matches`} of ${JSON.stringify(args.find)} with ${JSON.stringify(args.replace)} in ${matched.length === 1 ? 'one paragraph' : `${matched.length} paragraphs`}`;
      return edit.report.finish(summary, { matches: total });
    })));
  }));

  server.registerTool('set_style', {
    title: 'Set a paragraph\'s style and properties',
    description: 'Sets a paragraph\'s style (looked up by id first, then by name; added from docx4j\'s defaults when the document lacks it) and any of its alignment, indents and spacing, in points. Tracked as a paragraph-properties change.',
    inputSchema: {
      document: documentArg,
      address: addressSchema,
      style: z.string().min(1).optional(),
      alignment: z.enum(['Left', 'Centered', 'Right', 'Justified']).optional(),
      left_indent: z.number().optional(),
      first_line_indent: z.number().optional().describe('negative for a hanging indent'),
      space_before: z.number().min(0).optional(),
      space_after: z.number().min(0).optional(),
      line_spacing: z.number().min(0).optional(),
      if_text: ifTextArg,
      dry_run: dryRunArg,
    },
  }, (args) => answer(async () => {
    const given = ['style', 'alignment', 'left_indent', 'first_line_indent', 'space_before', 'space_after', 'line_spacing'].filter((k) => (args as Record<string, unknown>)[k] !== undefined);
    if (given.length === 0) throw new ToolError('arguments.invalid', 'nothing to set', 'give a style or a property', { argument: 'style' });
    return reportResult(await inRoom(host, args.document, (room) => mutate(host, room, args.dry_run ?? false, 'TrackAll', async (edit) => {
      const { paragraph, warnings } = resolveParagraph(edit.body, args.address);
      for (const w of warnings) edit.report.warn(w);
      await checkIfText(edit.body, paragraph, args.if_text);
      const styleId = args.style !== undefined ? await styleIdFor(edit, args.style) : undefined;
      edit.report.touch(paragraph);
      if (styleId !== undefined) paragraph.styleId = styleId;
      if (args.alignment !== undefined) paragraph.alignment = args.alignment;
      if (args.left_indent !== undefined) paragraph.leftIndent = args.left_indent;
      if (args.first_line_indent !== undefined) paragraph.firstLineIndent = args.first_line_indent;
      if (args.space_before !== undefined) paragraph.spaceBefore = args.space_before;
      if (args.space_after !== undefined) paragraph.spaceAfter = args.space_after;
      if (args.line_spacing !== undefined) paragraph.lineSpacing = args.line_spacing;
      return edit.report.finish(`Set ${given.join(', ')} on ${edit.body.addressOf(paragraph) ?? describeAddress(args.address)}`);
    })));
  }));

  server.registerTool('set_format', {
    title: 'Format text',
    description: 'Character formatting over a span of a paragraph (anchor, or offset and length; the whole paragraph without one): bold, italic, underline, strike, font, size (points), color (#RRGGBB), highlight (a Word highlight colour, or null to clear). Tracked as a formatting change.',
    inputSchema: {
      document: documentArg,
      address: addressSchema,
      ...spanShape,
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      strike: z.boolean().optional(),
      font: z.string().min(1).optional(),
      size: z.number().positive().optional(),
      color: z.string().regex(/^#?[0-9a-fA-F]{6}$|^auto$/).optional(),
      highlight: z.string().min(1).nullable().optional(),
      if_text: ifTextArg,
      dry_run: dryRunArg,
    },
  }, (args) => answer(async () => {
    const given = ['bold', 'italic', 'underline', 'strike', 'font', 'size', 'color', 'highlight'].filter((k) => (args as Record<string, unknown>)[k] !== undefined);
    if (given.length === 0) throw new ToolError('arguments.invalid', 'nothing to set', 'give a property: bold, italic, underline, strike, font, size, color or highlight', { argument: 'bold' });
    return reportResult(await inRoom(host, args.document, (room) => mutate(host, room, args.dry_run ?? false, 'TrackAll', async (edit) => {
      const { paragraph, warnings } = resolveParagraph(edit.body, args.address);
      for (const w of warnings) edit.report.warn(w);
      await checkIfText(edit.body, paragraph, args.if_text);
      const range = resolveSpan(paragraph, args, args.address);
      if (range.start === range.end) throw new ToolError('range.invalid', 'the span is empty', 'give the words, or an offset and a length above 0', { address: args.address });
      edit.report.touch(paragraph);
      const font = range.font;
      if (args.bold !== undefined) font.bold = args.bold;
      if (args.italic !== undefined) font.italic = args.italic;
      if (args.underline !== undefined) font.underline = args.underline ? 'Single' : 'None';
      if (args.strike !== undefined) font.strikeThrough = args.strike;
      if (args.font !== undefined) font.name = args.font;
      if (args.size !== undefined) font.size = args.size;
      if (args.color !== undefined) font.color = args.color === 'auto' ? 'auto' : `#${args.color.replace(/^#/, '').toUpperCase()}`;
      if (args.highlight !== undefined) font.highlightColor = args.highlight;
      return edit.report.finish(`Set ${given.join(', ')} on ${JSON.stringify(range.text)} in ${edit.body.addressOf(paragraph) ?? describeAddress(args.address)}`);
    })));
  }));

  server.registerTool('insert_ooxml', {
    title: 'Insert OOXML',
    description: 'Inserts WordprocessingML (w:p, w:tbl, or a flat OPC pkg:package, whose referenced parts are copied in) at a block or container, as edit places Markdown: the escape hatch for what Markdown cannot say. Tracked as the agent\'s insertions. It inserts content: no part is edited as XML.',
    inputSchema: {
      document: documentArg,
      address: addressSchema.optional(),
      ooxml: z.string().min(1),
      location: z.enum(LOCATIONS).describe('where: Replace, Before or After the block; Start or End of a container'),
      if_text: ifTextArg,
      dry_run: dryRunArg,
    },
  }, (args) => answer(async () => reportResult(await inRoom(host, args.document, (room) => mutate(host, room, args.dry_run ?? false, 'TrackAll', async (edit) => {
    const place = await placeFor(edit, args.address, args.location, LOCATIONS, args.if_text);
    const isPackage = /<pkg:package[\s>]/.test(args.ooxml);
    if (!isPackage) {
      // A bare fragment brings no relationships: every one it names must be the document's.
      const rels = edit.main.getRelationshipsPart();
      for (const m of args.ooxml.matchAll(/\br:(?:id|embed|link|pict|dm|lo|qs|cs)="([^"]+)"/g)) {
        if (!rels?.getRelationshipById(m[1]!)) throw new ToolError('ooxml.invalid', `the OOXML names relationship ${m[1]}, which the document does not have`, 'send a flat OPC pkg:package with the part it refers to, or leave the reference out', {});
      }
    }
    if (edit.dry && isPackage) {
      edit.report.warn('a dry run does not insert a package, whose parts it would copy in');
      return edit.report.finish(`Would insert the package ${where(place, args.location === 'Replace' ? 'After' : args.location)}`);
    }
    let inserted: (Paragraph | { element: Element; container: Element[] })[];
    try {
      if (place.kind === 'container') {
        // The contract's block placement (docs/developer/agent.md section 2.5): Before the container's first block or After its
        // last. The engine's Start and End place a package as Word's paste does since core-ts 0.8.0, joining its first
        // or last paragraph into the container's, which would edit a paragraph the agent did not address.
        const content = place.body.content;
        const edge = args.location === 'Start' ? content[0] : content[content.length - 1];
        inserted = edge
          ? await place.body.insertOoxml(args.ooxml, args.location === 'Start' ? 'Before' : 'After', { element: edge, container: content })
          : await place.body.insertOoxml(args.ooxml, args.location === 'Start' ? 'Start' : 'End');
      } else {
        inserted = await edit.body.insertOoxml(args.ooxml, args.location === 'Before' ? 'Before' : 'After', place.block);
      }
    } catch (e) {
      throw new ToolError('ooxml.invalid', `the OOXML was not inserted: ${(e as Error).message}`, 'give well-formed w:p or w:tbl elements (namespaces declared or the standard prefixes), or a flat OPC pkg:package', {});
    }
    const views: (Paragraph | Table)[] = [];
    for (const v of inserted) {
      if (v instanceof Paragraph) views.push(new Paragraph(v.element, v.container, edit.body));
      else if (typeNameOf(v.element) === 'org_docx4j_wml.Tbl') views.push(...viewsOf(edit.body, [v.element], v.container));
    }
    for (const v of views) edit.report.create(v);
    const others = inserted.length - views.length;
    if (others > 0) edit.report.warn(`${others === 1 ? 'one inserted block is' : `${others} inserted blocks are`} neither a paragraph nor a table (a content control, say): read shows ${others === 1 ? 'it' : 'them'}`);
    if (args.location === 'Replace') removeReplaced(edit, place, inserted.map((v) => v.element));
    const summary = args.location === 'Replace' && place.kind === 'block' ? `Replaced ${place.address} with ${counted(views)}` : `Inserted ${counted(views)} ${where(place, args.location)}`;
    return edit.report.finish(summary);
  })))));

  server.registerTool('comment', {
    title: 'Comment',
    description: 'A comment on a span of a paragraph (anchor, or offset and length; the whole paragraph without one), or, with reply_to, a reply in that comment\'s thread; resolve: true resolves the thread. The text is Markdown paragraphs. Comments are never tracked (Word does not track them).',
    inputSchema: {
      document: documentArg,
      address: addressSchema.optional(),
      ...spanShape,
      reply_to: z.string().min(1).optional().describe('a comment id, as get_comments gives it'),
      text: z.string().min(1),
      resolve: z.boolean().optional(),
      if_text: ifTextArg,
      dry_run: dryRunArg,
    },
  }, (args) => answer(async () => {
    if ((args.address === undefined) === (args.reply_to === undefined)) throw new ToolError('arguments.invalid', 'give address (a new comment) or reply_to (a reply), one of them', 'get_comments lists the threads and their ids', { argument: 'address' });
    const content = commentContent(args.text);
    return reportResult(await inRoom(host, args.document, async (room) => {
      await room.session.stories.comments();
      return mutate(host, room, args.dry_run ?? false, 'Off', async (edit) => {
        edit.report.untracked();
        const options = { author: { name: host.name }, date: minute() };
        if (args.reply_to !== undefined) {
          if (hasSpan(args)) throw new ToolError('arguments.invalid', 'a reply is anchored where its thread is', 'drop the span', { argument: 'anchor' });
          const known = new Set<string>();
          const collect = (ts: ReturnType<typeof room.session.stories.commentThreads>): void => { for (const t of ts) { known.add(t.id); collect(t.replies); } };
          collect(room.session.stories.commentThreads(edit.doc));
          if (!known.has(args.reply_to)) throw new ToolError('address.not_found', `no comment ${args.reply_to}`, 'get_comments lists the threads and their ids', { comment: args.reply_to });
          if (edit.dry) return edit.report.finish(`Would reply to comment ${args.reply_to}${args.resolve ? ' and resolve its thread' : ''}`);
          const target = allComments(await edit.body.getComments()).find((c) => String(c.id) === args.reply_to);
          if (!target) throw new ToolError('address.not_found', `comment ${args.reply_to} is not anchored in the body`, 'get_comments lists the threads', { comment: args.reply_to });
          const reply = await target.reply(content, options);
          let root = target;
          while (root.parent) root = root.parent;
          if (args.resolve) root.resolved = true;
          for (const name of commentParts(edit)) edit.report.part(name);
          return edit.report.finish(`Replied to comment ${args.reply_to}${args.resolve ? ' and resolved its thread' : ''}`, { comment: { id: String(reply.id), thread: String(root.id) } });
        }
        const { paragraph, warnings } = resolveParagraph(edit.body, args.address!);
        for (const w of warnings) edit.report.warn(w);
        await checkIfText(edit.body, paragraph, args.if_text);
        const range = resolveSpan(paragraph, args, args.address!);
        edit.report.touch(paragraph);
        if (edit.dry) return edit.report.finish(`Would comment on ${JSON.stringify(range.text)} in ${edit.body.addressOf(paragraph)}`);
        // In a room, a random id, as the editor gives the comments made there (ED-005 section 7.2).
        const comment = await range.insertComment(content, { ...options, ...(room.session.inRoom ? { id: randomInt(2 ** 20, 2 ** 31 - 1) } : {}) });
        if (args.resolve) comment.resolved = true;
        for (const name of commentParts(edit)) edit.report.part(name);
        return edit.report.finish(`Commented on ${JSON.stringify(range.text)} in ${edit.body.addressOf(paragraph)}`, { comment: { id: String(comment.id), thread: String(comment.id) } });
      });
    }));
  }));
}
