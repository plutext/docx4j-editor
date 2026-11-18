// SPDX-License-Identifier: GPL-3.0-or-later
//
// Addresses and spans (docs/developer/agent.md sections 2.2 to 2.4): a block named in one of the vocabulary's
// three forms, as the engine resolves them (`Body.elementAt`), and part of a paragraph named by its
// words or by offsets in its text as accepted (the engine's `Paragraph.text`). What does not
// resolve is `address.not_found` with the nearest address that exists, or `range.invalid`.
import { z } from 'zod';
import { Paragraph, Range, Table, Texts, typeNameOf, childrenOf, type Address, type Body, type BlockElement, type Element } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { ToolError } from './errors.mjs';

/** An address as a tool takes it: an ordinal (`"body/3"`), a paragraph id (`"w14:5A2B1C3D"`, the hex alone, or `{ paraId }`), or text (`{ contains }`). */
export const addressSchema = z.union([
  z.string().min(1).describe('an ordinal ("body/3", "body/7/0/1/0"), or a paragraph id ("5A2B1C3D" or "w14:5A2B1C3D")'),
  z.object({ paraId: z.string().min(1) }).strict().describe('a paragraph id, as outline and find give it'),
  z.object({ contains: z.string().min(1) }).strict().describe('the first paragraph whose text contains this'),
]);
export type AddressInput = z.infer<typeof addressSchema>;

/** A span of a paragraph (section 2.3): the words (their first occurrence, or the `occurrence`th), or offsets in its text. */
export const spanShape = {
  anchor: z.string().min(1).optional().describe('the words the span covers, in the paragraph\'s text as accepted'),
  occurrence: z.number().int().min(1).optional().describe('which occurrence of anchor, 1-based (default 1)'),
  offset: z.number().int().min(0).optional().describe('where the span starts in the paragraph\'s text (UTF-16 code units)'),
  length: z.number().int().min(0).optional().describe('how long the span is'),
};
export interface SpanInput {
  anchor?: string | undefined;
  occurrence?: number | undefined;
  offset?: number | undefined;
  length?: number | undefined;
}

const PARA_ID = /^[0-9a-f]{8}$/i;

/** An input address in the engine's form: a paragraph id upper case and without its prefix. */
export function normalizeAddress(input: AddressInput): Address {
  if (typeof input === 'string') {
    const s = input.trim();
    if (/^w14:/i.test(s)) return { paraId: s.slice(4).toUpperCase() };
    if (PARA_ID.test(s)) return { paraId: s.toUpperCase() };
    return s;
  }
  if ('paraId' in input) return { paraId: input.paraId.trim().replace(/^w14:/i, '').toUpperCase() };
  return { contains: input.contains };
}

/** An address in words, for messages. */
export function describeAddress(input: AddressInput): string {
  if (typeof input === 'string') return input;
  return 'paraId' in input ? `paraId ${input.paraId}` : `the paragraph containing ${JSON.stringify(input.contains)}`;
}

/** The ordinal nearest a missing one that exists: its longest prefix that resolves, the last index clamped. */
function nearestOrdinal(body: Body, ordinal: string): string | undefined {
  if (!ordinal.startsWith(`${body.prefix}/`)) return body.content.length > 0 ? `${body.prefix}/0` : undefined;
  const segments = ordinal.slice(body.prefix.length + 1).split('/');
  let current: object = body.container;
  const path: number[] = [];
  for (const seg of segments) {
    const children = childrenOf(current);
    const i = Number(seg);
    if (!children || children.length === 0 || !Number.isInteger(i) || i < 0) break;
    const at = Math.min(i, children.length - 1);
    path.push(at);
    if (at !== i) break;
    current = children[at]!.value as object;
  }
  return path.length > 0 ? [body.prefix, ...path].join('/') : undefined;
}

/** What `address.not_found` carries. */
function notFound(body: Body, input: AddressInput): ToolError {
  const address = normalizeAddress(input);
  const nearest = typeof address === 'string' ? nearestOrdinal(body, address) : undefined;
  return new ToolError('address.not_found', `nothing is at ${describeAddress(input)}`, 'call outline for the addresses as they are now; prefer a paragraph id, which the person\'s edits do not shift', { address: input, nearest: nearest ?? null });
}

/** How many paragraphs a text address matches, for the warning that later ones were left. */
function textMatches(body: Body, contains: string): number {
  let n = 0;
  for (const p of body.paragraphs) if (p.text.includes(contains)) n++;
  return n;
}

/** A block the address names, and any warning its resolution gives. */
export interface ResolvedBlock {
  block: BlockElement;
  warnings: string[];
}

/** The block an address names (a paragraph, a table, a row, a cell, a content control). */
export function resolveBlock(body: Body, input: AddressInput): ResolvedBlock {
  const address = normalizeAddress(input);
  const block = body.elementAt(address);
  if (!block) throw notFound(body, input);
  const warnings: string[] = [];
  if (typeof address === 'object' && 'contains' in address && textMatches(body, address.contains) > 1) {
    warnings.push(`the text address matched ${body.addressOf(block) ?? '?'}; later matches were not changed`);
  }
  return { block, warnings };
}

export const isParagraph = (el: Element): boolean => typeNameOf(el) === 'org_docx4j_wml.P';
export const isTable = (el: Element): boolean => typeNameOf(el) === 'org_docx4j_wml.Tbl';

/** The paragraph an address names; a table or other block there is `target.invalid`. */
export function resolveParagraph(body: Body, input: AddressInput): { paragraph: Paragraph; warnings: string[] } {
  const { block, warnings } = resolveBlock(body, input);
  if (!isParagraph(block.element)) {
    throw new ToolError('target.invalid', `${describeAddress(input)} is ${kindOf(block.element)}, not a paragraph`, 'address one of its paragraphs (outline lists them)', { address: input });
  }
  return { paragraph: new Paragraph(block.element as Element<wml.P>, block.container, body), warnings };
}

/** What kind of block an element is, in words. */
export function kindOf(el: Element): string {
  switch (typeNameOf(el)) {
    case 'org_docx4j_wml.P': return 'a paragraph';
    case 'org_docx4j_wml.Tbl': return 'a table';
    case 'org_docx4j_wml.Tr': return 'a table row';
    case 'org_docx4j_wml.Tc': return 'a table cell';
    case 'org_docx4j_wml.SdtBlock': return 'a content control';
    default: return 'something other than a paragraph';
  }
}

/** A paragraph's id as outputs spell it: the 8 hex digits alone, upper case; undefined when it has none. */
export function paraIdOf(p: wml.P): string | undefined {
  return p.paraId ? p.paraId.toUpperCase() : undefined;
}

/** A paragraph as reports and reads name it (section 2.7): its address, its id, its text as accepted. */
export function describeParagraph(body: Body, paragraph: Paragraph): { address: string; paraId?: string; text: string } {
  const paraId = paraIdOf(paragraph.p);
  return { address: body.addressOf(paragraph) ?? '?', ...(paraId ? { paraId } : {}), text: paragraph.text };
}

/** The table view of a block that is a table. */
export function tableOf(body: Body, block: BlockElement): Table {
  return new Table(block.element as Element<wml.Tbl>, block.container, body);
}

/** Whether any span field is given. */
export function hasSpan(span: SpanInput): boolean {
  return span.anchor !== undefined || span.offset !== undefined || span.length !== undefined || span.occurrence !== undefined;
}

/** The span of a paragraph a tool acts on (section 2.3): the whole paragraph without one. */
export function resolveSpan(paragraph: Paragraph, span: SpanInput, address: AddressInput): Range {
  const text = paragraph.text;
  if (span.anchor !== undefined && (span.offset !== undefined || span.length !== undefined)) {
    throw new ToolError('arguments.invalid', 'a span is anchor (with occurrence) or offset and length, not both', 'give one form', { argument: 'anchor' });
  }
  if (span.occurrence !== undefined && span.anchor === undefined) {
    throw new ToolError('arguments.invalid', 'occurrence needs anchor', 'give the words as anchor', { argument: 'occurrence' });
  }
  if (span.anchor !== undefined) {
    const occurrence = span.occurrence ?? 1;
    let at = -1;
    for (let k = 0; k < occurrence; k++) {
      at = text.indexOf(span.anchor, at + 1);
      if (at < 0) break;
    }
    if (at < 0) {
      throw new ToolError('range.invalid', `${JSON.stringify(span.anchor)} occurs fewer than ${occurrence} times in ${describeAddress(address)}`, 'read the paragraph again: its text is in the error\'s current field', { address, anchor: span.anchor, current: { text } });
    }
    return new Range(paragraph, at, at + span.anchor.length);
  }
  if (span.offset !== undefined || span.length !== undefined) {
    const offset = span.offset ?? 0;
    const length = span.length ?? text.length - offset;
    if (offset > text.length || offset + length > text.length || length < 0) {
      throw new ToolError('range.invalid', `offset ${offset} and length ${length} run past the end of ${describeAddress(address)}, whose text is ${text.length} long`, 'offsets count the paragraph\'s text as accepted, in UTF-16 code units', { address, offset, current: { text } });
    }
    return new Range(paragraph, offset, offset + length);
  }
  return new Range(paragraph, 0, text.length);
}

/**
 * A paragraph as the person left it, by its text, for a call refused because their typing reached
 * it while the call ran (section 2.10): described as `if_text`'s refusal describes one, or its text
 * alone if it cannot be found again.
 */
export function paragraphNow(body: Body, text: string): { address?: string; paraId?: string; text: string } {
  try {
    return describeParagraph(body, resolveParagraph(body, { contains: text }).paragraph);
  } catch {
    return { text };
  }
}

/**
 * `if_text` (section 2.4): the paragraph's text now against the agent's reading; `target.stale` with the current text
 * when they differ. The reading is the paragraph's text as accepted in either of its two spellings: Office JS's
 * (`Paragraph.text`, what the change reports and `target.stale` carry) or the text model's (what outline, find and read
 * show, ED-008: a note's reference no character, a line break a newline). A paragraph that changed differs in both.
 */
export async function checkIfText(body: Body, paragraph: Paragraph, ifText: string | undefined): Promise<void> {
  if (ifText === undefined) return;
  if (paragraph.text === ifText) return;
  // The paragraph alone, no package: its accepted runs, as the text model reads one paragraph (core-ts CR-012).
  const read = await Texts.of(paragraph.p as object, { view: 'accepted' });
  if (read.blocks.length === 1 && read.blocks[0]!.text === ifText) return;
  throw new ToolError('target.stale', 'the paragraph changed since it was read', 'decide again from its current text; pass that as if_text', { current: describeParagraph(body, paragraph) });
}
