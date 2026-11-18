// SPDX-License-Identifier: GPL-3.0-or-later
//
// The content-control tools (docs/developer/agent.md section 3.4), the editor's own. A control wrapped or
// unwrapped, its properties, and a binding to the document's custom XML are never tracked (rule 1:
// the model holds no revision for a control, ED-005 proposal 5); the content inside one is tracked
// like any other, so removing a control with its content deletes the content as a tracked change.
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { Paragraph, deepCopy, linkParents, nextSdtId, r as runOf, sdt, type ContentControl, type Element, type SdtKind } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { parseTag, withTagParam, formatTag, readTag } from '@docx4j/editor-model';
import { ToolError } from './errors.mjs';
import { addressSchema, checkIfText, describeAddress, hasSpan, isParagraph, isTable, kindOf, resolveBlock, resolveSpan, spanShape, type AddressInput } from './address.mjs';
import { answer, documentArg, dryRunArg, ifTextArg, inRoom, mutate, reportResult, type Edit, type ToolHost } from './kit.mjs';

/** The contract's kinds (as Office JS names them in `insertContentControl`) and the engine's. */
const KINDS = { RichText: 'RichText', PlainText: 'PlainText', Checkbox: 'CheckBox', DropDownList: 'DropDownList', Date: 'DatePicker', Picture: 'Picture', RepeatingSection: 'RepeatingSection' } as const satisfies Record<string, SdtKind>;

/** A control named by its `w:id`, its tag (the first with it), or the address of the block it is. */
const controlArg = z.union([
  z.object({ id: z.number().int() }).strict(),
  z.object({ tag: z.string().min(1) }).strict(),
  z.object({ address: addressSchema }).strict(),
]).describe('{ id } (its w:id), { tag } (the first control with it), or { address } (a block content control\'s address)');
type ControlArg = z.infer<typeof controlArg>;

function findControl(edit: Edit, spec: ControlArg): ContentControl {
  const all = edit.body.contentControls;
  if ('id' in spec) {
    const found = all.find((c) => c.id === spec.id);
    if (!found) throw new ToolError('address.not_found', `no content control has id ${spec.id}`, 'read shows the controls; outline gives their addresses', { control: spec });
    return found;
  }
  if ('tag' in spec) {
    const found = all.filter((c) => c.tag === spec.tag);
    if (found.length === 0) throw new ToolError('address.not_found', `no content control has the tag ${JSON.stringify(spec.tag)}`, 'name it by its id or address', { control: spec });
    if (found.length > 1) edit.report.warn(`${found.length} content controls have the tag ${JSON.stringify(spec.tag)}; the first was used`);
    return found[0]!;
  }
  const { block } = resolveBlock(edit.body, spec.address);
  const found = all.find((c) => c.element === block.element);
  if (!found) throw new ToolError('target.invalid', `${describeAddress(spec.address)} is ${kindOf(block.element)}, not a content control`, 'address the control itself, or name it by id or tag', { control: spec });
  return found;
}

/** Where a control is: its own address for a block one, its paragraph's for one inside a paragraph. */
function controlAddress(edit: Edit, control: ContentControl): string | undefined {
  let v = control.element.value as { TYPE_NAME?: string; PARENT?: object } | undefined;
  for (let depth = 0; v && depth < 16; depth++) {
    if (v.TYPE_NAME === 'org_docx4j_wml.P') return edit.body.addressOf({ element: { value: v } as unknown as Element, container: [] });
    v = v.PARENT as typeof v;
  }
  return edit.body.addressOf({ element: control.element, container: [] });
}

function describeControl(edit: Edit, control: ContentControl): Record<string, unknown> {
  return { id: control.id, tag: control.tag, title: control.title, kind: control.type, address: controlAddress(edit, control) ?? null };
}

/** Blocks of one container wrapped in a new block control (the engine wraps one paragraph; several, or a table, are wrapped here). */
function wrapBlocks(edit: Edit, items: Element[], container: Element[], kind: SdtKind, tag: string | undefined, title: string | undefined): ContentControl {
  const at = container.indexOf(items[0]!);
  const owner = (items[0]!.value as { PARENT?: object }).PARENT;
  const control = sdt(items, { kind, id: nextSdtId(edit.document), form: 'block', ...(tag !== undefined ? { tag } : {}), ...(title !== undefined ? { title } : {}) });
  container.splice(at, items.length, control as Element);
  linkParents(control, owner);
  const view = edit.body.contentControls.find((c) => c.element === control);
  if (!view) throw new Error('the control made was not found in the body');
  return view;
}

export function registerControlTools(server: McpServer, host: ToolHost): void {
  server.registerTool('insert_content_control', {
    title: 'Insert a content control',
    description: 'Wraps a paragraph (or, with to, the blocks from address to to in one container; or a table; or, with a span, words in a paragraph) in a new content control of a kind, with a tag and title. Not tracked (rule 1).',
    inputSchema: {
      document: documentArg,
      address: addressSchema,
      ...spanShape,
      to: addressSchema.optional().describe('the last block to wrap, in the same container as address'),
      kind: z.enum(['RichText', 'PlainText', 'Checkbox', 'DropDownList', 'Date', 'Picture', 'RepeatingSection']).optional().describe('default RichText'),
      tag: z.string().optional(),
      title: z.string().optional(),
      if_text: ifTextArg,
      dry_run: dryRunArg,
    },
  }, (args) => answer(async () => reportResult(await inRoom(host, args.document, (room) => mutate(host, room, args.dry_run ?? false, 'Off', async (edit) => {
    const kind = KINDS[args.kind ?? 'RichText'];
    const { block, warnings } = resolveBlock(edit.body, args.address);
    for (const w of warnings) edit.report.warn(w);
    const at = edit.body.addressOf(block) ?? describeAddress(args.address);
    let control: ContentControl;
    if (args.to !== undefined) {
      if (hasSpan(args)) throw new ToolError('arguments.invalid', 'to wraps whole blocks; a span is part of one paragraph', 'give to or a span, not both', { argument: 'to' });
      const end = resolveBlock(edit.body, args.to as AddressInput).block;
      if (end.container !== block.container) throw new ToolError('target.invalid', 'address and to are in different containers', 'wrap blocks of one container (the body, one cell, one control)', { address: args.address });
      const from = block.container.indexOf(block.element);
      const to = block.container.indexOf(end.element);
      if (to < from) throw new ToolError('target.invalid', 'to comes before address', 'give the first block as address and the last as to', { address: args.address });
      if (args.if_text !== undefined) throw new ToolError('arguments.invalid', 'if_text names one paragraph, and to wraps several blocks', 'drop if_text', { argument: 'if_text' });
      const items = block.container.slice(from, to + 1);
      for (const item of items) if (isParagraph(item)) edit.report.touch(new Paragraph(item as Element<wml.P>, block.container, edit.body));
      control = wrapBlocks(edit, items, block.container, kind, args.tag, args.title);
    } else if (isParagraph(block.element)) {
      const paragraph = new Paragraph(block.element as Element<wml.P>, block.container, edit.body);
      await checkIfText(edit.body, paragraph, args.if_text);
      edit.report.touch(paragraph);
      control = hasSpan(args) ? resolveSpan(paragraph, args, args.address).insertContentControl(kind) : paragraph.insertContentControl(kind);
    } else if (isTable(block.element)) {
      if (hasSpan(args) || args.if_text !== undefined) throw new ToolError('arguments.invalid', `a span and if_text name a paragraph, and ${at} is a table`, 'drop them', { argument: 'anchor' });
      control = wrapBlocks(edit, [block.element], block.container, kind, args.tag, args.title);
    } else {
      throw new ToolError('target.invalid', `${at} is ${kindOf(block.element)}`, 'wrap a paragraph, words in one, a table, or blocks from address to to', { address: args.address });
    }
    if (args.tag !== undefined) control.tag = args.tag;
    if (args.title !== undefined) control.title = args.title;
    return edit.report.finish(`Wrapped ${args.to !== undefined ? `${at} to ${edit.body.addressOf(resolveBlock(edit.body, args.to as AddressInput).block) ?? '?'}` : at} in a ${args.kind ?? 'RichText'} content control`, { control: describeControl(edit, control) });
  })))));

  server.registerTool('remove_content_control', {
    title: 'Remove a content control',
    description: 'Unwraps a content control, keeping its content (not tracked, rule 1), or, with keep_content: false, removes the content with it, the content\'s deletion tracked.',
    inputSchema: { document: documentArg, control: controlArg, keep_content: z.boolean().optional().describe('default true'), dry_run: dryRunArg },
  }, (args) => answer(async () => {
    const keep = args.keep_content ?? true;
    return reportResult(await inRoom(host, args.document, (room) => mutate(host, room, args.dry_run ?? false, keep ? 'Off' : 'TrackAll', async (edit) => {
      const control = findControl(edit, args.control);
      const described = describeControl(edit, control);
      if (!keep) {
        // The content deleted as the agent's tracked change, then the control unwrapped around it.
        if (control.form === 'Run') {
          control.getRange('Content').delete();
        } else {
          for (const paragraph of control.paragraphs) { edit.report.remove(paragraph); paragraph.delete(); }
          for (const table of control.tables) { edit.report.remove(table); table.delete(); }
        }
      } else {
        for (const paragraph of control.paragraphs) edit.report.touch(paragraph);
      }
      control.delete(true);
      return edit.report.finish(`Removed the content control ${String(described['id'])}${keep ? ', keeping its content' : ' and its content'}`, { control: described });
    })));
  }));

  server.registerTool('set_content_control', {
    title: 'Set a content control\'s properties',
    description: 'Sets a content control\'s tag, title, placeholder text, lock (none, content, control, both), and its kind\'s own properties: checked (a checkbox), items (a drop-down list, as { text, value? }), date_format (a date). Not tracked (rule 1).',
    inputSchema: {
      document: documentArg,
      control: controlArg,
      tag: z.string().optional(),
      title: z.string().optional(),
      placeholder: z.string().optional(),
      lock: z.enum(['none', 'content', 'control', 'both']).optional(),
      checked: z.boolean().optional(),
      items: z.array(z.object({ text: z.string().min(1), value: z.string().optional() }).strict()).optional(),
      date_format: z.string().min(1).optional(),
      dry_run: dryRunArg,
    },
  }, (args) => answer(async () => {
    const given = ['tag', 'title', 'placeholder', 'lock', 'checked', 'items', 'date_format'].filter((k) => (args as Record<string, unknown>)[k] !== undefined);
    if (given.length === 0) throw new ToolError('arguments.invalid', 'nothing to set', 'give a property', { argument: 'tag' });
    return reportResult(await inRoom(host, args.document, (room) => mutate(host, room, args.dry_run ?? false, 'Off', async (edit) => {
      const control = findControl(edit, args.control);
      const refuse = (what: string, kind: string): never => { throw new ToolError('target.invalid', `${what} is a ${kind}'s property, and this control is ${control.type}`, 'set only the properties its kind has', { control: args.control }); };
      if (args.checked !== undefined) { const box = control.checkboxContentControl; if (!box) refuse('checked', 'checkbox'); box!.isChecked = args.checked; }
      if (args.items !== undefined) {
        const list = control.dropDownListContentControl ?? control.comboBoxContentControl;
        if (!list) refuse('items', 'drop-down list');
        list!.deleteAllListItems();
        for (const item of args.items) list!.addListItem(item.text, item.value);
      }
      if (args.date_format !== undefined) { const date = control.datePickerContentControl; if (!date) refuse('date_format', 'date'); date!.dateFormat = args.date_format; }
      if (args.tag !== undefined) control.tag = args.tag;
      if (args.title !== undefined) control.title = args.title;
      if (args.placeholder !== undefined) control.placeholderText = args.placeholder;
      if (args.lock !== undefined) {
        control.cannotEdit = args.lock === 'content' || args.lock === 'both';
        control.cannotDelete = args.lock === 'control' || args.lock === 'both';
      }
      return edit.report.finish(`Set ${given.join(', ')} on the content control ${control.id}`, { control: describeControl(edit, control) });
    })));
  }));

  server.registerTool('bind', {
    title: 'Bind a content control',
    description: 'Binds a content control to a node of the document\'s custom XML (the data part by the OpenDoPE rules, or store_item), as the Template perspective does: w:dataBinding, and, where the template uses OpenDoPE XPaths, its entry and the od:xpath tag. The control shows the node\'s value. Not tracked (rule 1).',
    inputSchema: {
      document: documentArg,
      control: controlArg,
      xpath: z.string().min(1),
      store_item: z.string().min(1).optional().describe('the custom XML part\'s itemID; the data part when omitted'),
      prefix_mappings: z.string().optional().describe("xmlns:ns0='...' for the prefixes the XPath uses"),
      dry_run: dryRunArg,
    },
  }, (args) => answer(async () => reportResult(await inRoom(host, args.document, async (room) => {
    const template = room.session.template;
    await template.load();
    return mutate(host, room, args.dry_run ?? false, 'Off', async (edit) => {
      const control = findControl(edit, args.control);
      const bare = (id: string): string => id.replace(/[{}]/g, '').toLowerCase();
      const part = args.store_item !== undefined ? template.parts.find((p) => bare(p.id) === bare(args.store_item!)) : template.dataPart.part;
      if (!part) {
        throw args.store_item !== undefined
          ? new ToolError('arguments.invalid', `no custom XML part has itemID ${args.store_item}`, 'give the itemID of one of the document\'s custom XML parts', { argument: 'store_item' })
          : new ToolError('target.invalid', 'the document has no custom XML data part to bind to', 'a person can add one in the Template perspective\'s Data tab', {});
      }
      let node;
      try { node = part.selectSingleNode(args.xpath, args.prefix_mappings); } catch (e) { throw new ToolError('arguments.invalid', `the XPath does not parse: ${(e as Error).message}`, 'XPath 1.0 over the data part; declare its prefixes in prefix_mappings', { argument: 'xpath' }); }
      const repeating = control.type === 'RepeatingSection';
      if (!node) {
        throw repeating
          ? new ToolError('binding.no_repeat_node', `${JSON.stringify(args.xpath)} selects no node in the data, so there is nothing to repeat`, 'give the data an item to repeat first', { xpath: args.xpath })
          : new ToolError('binding.no_node', `${JSON.stringify(args.xpath)} selects no node in the data part`, 'Word binds to a node that exists: check the XPath against the part', { xpath: args.xpath });
      }
      if (node.node === part.documentElement.node && !repeating) {
        throw new ToolError('binding.document_element', 'the XPath selects the data\'s document element, and a control shows one node\'s value', 'bind to an element or attribute under it', { xpath: args.xpath });
      }
      const first = control.paragraphs[0];
      if (first) edit.report.touch(new Paragraph(first.element, first.container, edit.body));
      if (!edit.dry) {
        if (!control.xmlMapping.setMapping(args.xpath, args.prefix_mappings, part)) {
          throw new ToolError('binding.no_node', `the control could not be bound to ${JSON.stringify(args.xpath)}`, 'check the XPath against the part', { xpath: args.xpath });
        }
        // The control shows the node's value at once, as Word's does (docx4j's BindingHandler), in the control's run properties.
        if (!repeating) {
          const run = runOf(node.text);
          if (control.runProperties) run.value.rPr = deepCopy(control.runProperties);
          control.setBoundContent([run], false);
        }
        // OpenDoPE (REC-002): an XPaths entry and the od:xpath tag, where the template uses them.
        const existing = readTag(control.tag).xpath;
        if (template.openDopeParts.xpaths !== undefined || existing !== undefined) {
          const entry = template.upsertEntry({ ...(existing !== undefined ? { id: existing } : {}), xpath: args.xpath, storeItemID: part.id, ...(args.prefix_mappings !== undefined ? { prefixMappings: args.prefix_mappings } : {}) });
          control.tag = formatTag(withTagParam(parseTag(control.tag), 'od:xpath', entry.id));
          const xpathsPart = template.openDopeParts.xpaths;
          if (xpathsPart) edit.report.part(xpathsPart.part.partName.toString());
        }
      }
      return edit.report.finish(`Bound the content control ${control.id} to ${args.xpath}`, { control: describeControl(edit, control) });
    });
  }))));
}
