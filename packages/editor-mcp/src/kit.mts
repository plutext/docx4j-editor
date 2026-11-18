// SPDX-License-Identifier: GPL-3.0-or-later
//
// What every tool shares (docs/developer/agent.md section 2): the room a `document` handle names, one call at
// a time per room, the agent always tracked, a mutating call run as one tree run (or, dry, over a
// copy), the parts a build touches written only when not dry, and where a location puts a block.
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { z } from 'zod';
import { Paragraph, Table, typeNameOf, type Body, type BlockElement, type Element } from '@docx4j/core-ts';
import type { EditorDocument, TreeBindings, TreeRunTracking } from '@docx4j/editor-model';
import { ToolError, errorResult, type ErrorResult } from './errors.mjs';
import { PathError } from './paths.mjs';
import { ReportBuilder, type ChangeReport } from './report.mjs';
import { checkIfText, describeAddress, kindOf, paragraphNow, resolveBlock, tableOf, type AddressInput } from './address.mjs';
import type { BuildContext, ListPlace } from './markdown.mjs';

/** A room the agent is in, as the tools see it (`server.mts` makes them). */
export interface ToolRoom {
  readonly id: string;
  /** The file it came from; none for a document a person invited the agent into from their browser. */
  readonly path: string | undefined;
  readonly link: string;
  readonly session: EditorDocument;
  /** Set when the person pressed Stop for the agent in this room (section 2.11). */
  stopped: boolean;
  /** The files the agent saved this room to, which a Stop covers with the room's own (section 2.11). */
  readonly saved: Set<string>;
  /** One call at a time, in order (section 2.10). */
  queue<T>(work: () => Promise<T>): Promise<T>;
  /** Who is in the room now, from its awareness. */
  people(): { name: string; agent: boolean }[];
  /**
   * After a refused call (ED-005 proposal 42): let the room's session go, since its tree holds the
   * refused call, and join the room afresh; `session` is the new one after it.
   */
  renew(): Promise<void>;
  close(): Promise<void>;
}

/** The companion as the tools see it. */
export interface ToolHost {
  /** The agent's name: its tracked changes' and comments' author. */
  readonly name: string;
  readonly allowResolve: boolean;
  readonly roots: readonly string[];
  /** Section 2.9's cap on a result's characters. */
  readonly maxInlineChars: number;
  readonly rooms: ReadonlyMap<string, ToolRoom>;
  /** Opens a file under a root into a room, or answers the room it is already open in. */
  open(path: string): Promise<ToolRoom>;
  /** Whether the relay is running, after any restart in progress (absent: always). An edit that would reach no one is refused. */
  relayReady?(): Promise<boolean>;
}

/** A tool's answer: text for the model and the same as structured content. */
export interface ToolResult {
  content: { type: 'text'; text: string }[];
  structuredContent: Record<string, unknown>;
  [key: string]: unknown;
}

export function ok(structured: Record<string, unknown>, text = JSON.stringify(structured)): ToolResult {
  return { content: [{ type: 'text', text }], structuredContent: structured };
}

/** A path refusal as the contract's error. */
/** The tools' argument for each of the engine's measurement setters a tool hands a value to (`set_style`, `set_format`). */
const SETTER_ARGUMENTS: Readonly<Record<string, string>> = {
  'Paragraph.leftIndent': 'left_indent', 'Paragraph.firstLineIndent': 'first_line_indent', 'Paragraph.spaceBefore': 'space_before',
  'Paragraph.spaceAfter': 'space_after', 'Paragraph.lineSpacing': 'line_spacing', 'Font.size': 'size',
};

function mapError(e: unknown): unknown {
  // A measurement past what Word stores or takes (core-ts 0.7.0, CR-002 section 65): the engine refuses it before anything
  // changes, its message naming the setter, the range and what Word would read. Matched by `code`, as an add-in branches
  // on Word's errors; the class is not on the engine's index.
  if (e instanceof Error && (e as { code?: unknown }).code === 'InvalidArgument') {
    const argument = SETTER_ARGUMENTS[/^([A-Za-z]+\.[A-Za-z]+)/.exec(e.message)?.[1] ?? ''];
    return new ToolError('arguments.invalid', e.message, 'give a value inside the range the message names', argument ? { argument } : {});
  }
  // The engine reads a block in a kept mc:AlternateContent's branch and does not edit it (core-ts 0.6.0,
  // CR-013; ED-009 section 9.5 item 2): its refusal is the target's, not the companion's.
  if (e instanceof Error && e.name === 'NotSupportedError') {
    return new ToolError('target.invalid', e.message, 'the block is in a branch of an mc:AlternateContent, which the engine reads and does not edit; a person can edit it in the editor');
  }
  if (e instanceof PathError) {
    return e.code === 'not_found'
      ? new ToolError('path.not_found', e.message, 'a path is under a --root the companion was started with; a relative one is taken from the first root')
      : new ToolError('path.not_allowed', e.message, 'only files under the companion\'s --root folders can be opened or written', {});
  }
  return e;
}

/** A tool's work, its failure as the contract's error result. */
export async function answer(work: () => Promise<ToolResult>): Promise<ToolResult | ErrorResult> {
  try {
    return await work();
  } catch (e) {
    return errorResult(mapError(e));
  }
}

/** A mutating call's report as its answer: the one-line summary first, then the report. */
export function reportResult(report: ChangeReport): ToolResult {
  const line = report.dry_run ? `${report.summary} (dry run: nothing was written)` : report.summary;
  return ok(report, `${line}\n${JSON.stringify(report)}`);
}

// ---------------------------------------------------------------- schemas

export const documentArg = z.string().min(1).describe('the room id open returned (list_documents lists them)');
export const dryRunArg = z.boolean().optional().describe('return the change report without writing anything');
export const ifTextArg = z.string().optional().describe('the paragraph\'s text as you last read it; the call is refused with target.stale, writing nothing, when it has changed');
export const searchShape = {
  match_case: z.boolean().optional(),
  match_whole_word: z.boolean().optional(),
  match_wildcards: z.boolean().optional().describe('Word\'s wildcards: ? * <word> [a-z] @ {n,m}'),
  match_regexp: z.boolean().optional().describe('a JavaScript regular expression; $1 in a replacement names a group'),
};
export const LOCATIONS = ['Replace', 'Before', 'After', 'Start', 'End'] as const;
export type Location = (typeof LOCATIONS)[number];

export function searchOptions(args: { match_case?: boolean | undefined; match_whole_word?: boolean | undefined; match_wildcards?: boolean | undefined; match_regexp?: boolean | undefined }): { matchCase?: boolean; matchWholeWord?: boolean; matchWildcards?: boolean; matchRegExp?: boolean } {
  if (args.match_wildcards && args.match_regexp) throw new ToolError('arguments.invalid', 'match_wildcards and match_regexp cannot be used together', 'choose one', { argument: 'match_regexp' });
  return {
    ...(args.match_case ? { matchCase: true } : {}),
    ...(args.match_whole_word ? { matchWholeWord: true } : {}),
    ...(args.match_wildcards ? { matchWildcards: true } : {}),
    ...(args.match_regexp ? { matchRegExp: true } : {}),
  };
}

// ---------------------------------------------------------------- rooms and calls

/**
 * What the agent is told when the person has stopped it (docs/developer/agent.md section 2.11). The Stop
 * covers the document: the hint must never offer a way round it, as the first wording did ("open the
 * file again"), which Claude Code took (ED-005 section 12.29).
 */
export const STOPPED_HINT = 'stop working on this document, and do not open its file or a copy of it again; tell the person you have stopped. They can bring you back with Invite in the editor';

function stopped(document: string): ToolError {
  return new ToolError('agent.stopped', 'the person pressed Stop for you in this document', STOPPED_HINT, { document });
}

export function roomOf(host: ToolHost, document: string): ToolRoom {
  const room = host.rooms.get(document);
  if (!room) throw new ToolError('document.not_open', `no document is open as ${document}`, 'call list_documents for the rooms you are in, or open the file', { document });
  if (room.stopped) throw stopped(document);
  return room;
}

/** A call on a room: after the calls before it, refused once the person stopped the agent. */
export function inRoom<T>(host: ToolHost, document: string, work: (room: ToolRoom) => Promise<T>): Promise<T> {
  const room = roomOf(host, document);
  return room.queue(async () => {
    if (room.stopped) throw stopped(document);
    return work(room);
  });
}

/** A change's date: the call's, to the minute (ED-005 section 8.5). */
export function minute(): Date {
  return new Date(Math.floor(Date.now() / 60_000) * 60_000);
}

/** What a mutating call runs with: the tree, the room, whether it is a dry run, and its report. */
export interface Edit extends TreeBindings {
  room: ToolRoom;
  host: ToolHost;
  dry: boolean;
  report: ReportBuilder;
}

/**
 * A mutating call (rule 2): one tree run tracked as the agent (`mode`: `TrackAll`, or `Off` for
 * what is never tracked, rule 1), whatever the document's own setting, its change reaching the
 * room as one transaction; or, dry, the same over a copy of the tree, nothing written. The work
 * checks everything before it changes anything, so a refusal writes nothing.
 */
export async function mutate(host: ToolHost, room: ToolRoom, dry: boolean, mode: 'TrackAll' | 'Off', work: (edit: Edit) => Promise<ChangeReport>): Promise<ChangeReport> {
  const tracking: TreeRunTracking = { author: host.name, mode, date: minute() };
  if (!dry && host.relayReady && !(await host.relayReady())) {
    throw new ToolError('relay.down', 'the companion\'s relay stopped and could not be restarted, so nothing reaches the rooms', 'tell the person the companion needs restarting; nothing you do now would reach the document');
  }
  const edit = (b: TreeBindings, isDry: boolean): Edit => ({ ...b, room, host, dry: isDry, report: new ReportBuilder(b.body, host.name) });
  if (dry) {
    const report = await room.session.readTree((b) => work(edit(b, true)), { copy: true, tracking });
    return { ...report, dry_run: true };
  }
  // Guarded (ED-005 proposal 42): the person's typing that arrived while the call ran is let in
  // before the merge, and a block the call replaced that it reached refuses the call, writing nothing.
  const run = await room.session.runOnTree((b) => work(edit(b, false)), undefined, tracking, { guard: true });
  if (run.refused) {
    await room.renew();
    // The paragraph as the person left it, described as `if_text`'s refusal describes one (section 2.10).
    const text = run.refused.blocks.find((b) => b.text !== undefined)?.text;
    const current = text === undefined ? null : await room.session.readTree(({ body }) => paragraphNow(body, text));
    throw new ToolError('target.stale', 'the person changed a paragraph your call was changing, while it ran; nothing was written', 'read the paragraph again and redo your edit with the person\'s words in it', { current });
  }
  if (run.error) throw run.error;
  return run.value!;
}

// ---------------------------------------------------------------- parts a build touches

const RELS_PART = '/word/_rels/document.xml.rels';

export function stylesPartName(b: TreeBindings): string {
  return b.main.styleDefinitionsPart?.partName.toString() ?? '/word/styles.xml';
}

function numberingPartName(b: TreeBindings): string {
  return b.main.numberingDefinitionsPart?.partName.toString() ?? '/word/numbering.xml';
}

/** The document's styles, as the styles part holds them now. */
export async function stylesOf(b: TreeBindings): Promise<wml.Style[]> {
  const part = b.main.styleDefinitionsPart;
  if (!part) return [];
  return ((await part.getContents()) as wml.Styles).style ?? [];
}

/**
 * A style's id, looked up by id first, then by name (so an id wins where both match), then added
 * from docx4j's defaults (untracked, rule 1: the styles part is named in the report). A dry run
 * names the part without adding, and cannot tell whether the defaults have it.
 */
export async function styleIdFor(edit: Edit, style: string): Promise<string> {
  const styles = await stylesOf(edit);
  if (styles.some((s) => s.styleId === style)) return style;
  const byName = styles.find((s) => (s.name?.val ?? '').toLowerCase() === style.toLowerCase());
  if (byName?.styleId) return byName.styleId;
  const candidate = style.split(/\s+/).filter((w) => w !== '').map((w) => w.charAt(0).toUpperCase() + w.slice(1)).join('');
  if (styles.some((s) => s.styleId === candidate)) return candidate;
  if (edit.dry) {
    edit.report.part(stylesPartName(edit));
    edit.report.warn(`${candidate} would be added from docx4j's default styles; a dry run does not check that the defaults have it`);
    return candidate;
  }
  await edit.room.session.ensureStyles([candidate]);
  if ((await stylesOf(edit)).some((s) => s.styleId === candidate)) {
    edit.report.part(stylesPartName(edit));
    return candidate;
  }
  throw new ToolError('style.not_found', `no style ${JSON.stringify(style)}, by id or by name, in the document or in docx4j's defaults`, 'use a style id or name the document has: read shows each paragraph\'s style', { style });
}

/** The style Word gives the paragraph after one in this style (`w:next`, as Enter does), else the same. */
export async function nextStyleOf(b: TreeBindings, styleId: string | undefined): Promise<string | undefined> {
  if (styleId === undefined) return undefined;
  const style = (await stylesOf(b)).find((s) => s.styleId === styleId);
  return style?.next?.val ?? styleId;
}

/** What building Markdown needs, each part written (rule 1) only when the call is not dry. */
export function buildContext(edit: Edit, paragraphStyle: string | undefined): BuildContext {
  const lists = new Map<string, string>();
  let firstNumbered: string | undefined;
  let linkStyle: string | undefined;
  return {
    paragraphStyle,
    headingStyle: (level) => styleIdFor(edit, `heading ${level}`),
    numId: async (place: ListPlace) => {
      const made = lists.get(place.key);
      if (made !== undefined) return made;
      edit.report.part(numberingPartName(edit));
      let numId: string;
      if (edit.dry) numId = '0';
      // A second numbered list in one call starts again at 1, over the first's definition.
      else if (place.ordered && firstNumbered !== undefined) numId = await edit.room.session.restartList(firstNumbered);
      else numId = await edit.room.session.newList(!place.ordered, edit.doc);
      if (place.ordered && firstNumbered === undefined) firstNumbered = numId;
      lists.set(place.key, numId);
      return numId;
    },
    linkId: async (address) => {
      edit.report.part(RELS_PART);
      return edit.dry ? 'rId0' : edit.room.session.addHyperlink(address);
    },
    linkStyle: async () => (linkStyle ??= await styleIdFor(edit, 'Hyperlink')),
    tableStyle: async () => {
      try { return await styleIdFor(edit, 'TableGrid'); } catch { return undefined; }
    },
  };
}

// ---------------------------------------------------------------- where a location puts blocks

/** Where blocks go: next to (or in place of) a block, or at the start or end of a container. */
export type Place =
  | { kind: 'block'; block: BlockElement; address: string; paragraph?: Paragraph; table?: Table }
  | { kind: 'container'; body: Body; label: string };

function locationInvalid(location: string, allowed: readonly string[], why: string): ToolError {
  return new ToolError('location.invalid', why, `this call takes ${allowed.join(', ')}`, { location, allowed });
}

/**
 * The place a location names (section 2.5): `Start` and `End` a container (the body without an
 * address, a table cell, a content control), the others a block (a paragraph or a table; a
 * content control for `Before` and `After`). `if_text` checks a paragraph target.
 */
export async function placeFor(edit: Edit, address: AddressInput | undefined, location: Location, allowed: readonly Location[], ifText: string | undefined): Promise<Place> {
  if (!allowed.includes(location)) throw locationInvalid(location, allowed, `this call does not take ${location}`);
  if (location === 'Start' || location === 'End') {
    if (ifText !== undefined) throw new ToolError('arguments.invalid', 'if_text names a paragraph; Start and End name a container', 'drop if_text, or use Before or After a paragraph', { argument: 'if_text' });
    if (address === undefined) return { kind: 'container', body: edit.body, label: 'the body' };
    const { block, warnings } = resolveBlock(edit.body, address);
    for (const w of warnings) edit.report.warn(w);
    const at = edit.body.addressOf(block) ?? describeAddress(address);
    const type = typeNameOf(block.element);
    if (type === 'org_docx4j_wml.Tc') return { kind: 'container', body: edit.body.sub(block.element.value as { content?: Element[] }, at), label: `the cell ${at}` };
    if (type === 'org_docx4j_wml.SdtBlock') {
      const content = (block.element.value as wml.SdtBlock).sdtContent;
      if (!content) throw new ToolError('target.invalid', `the content control at ${at} has no content to hold blocks`, 'insert before or after it instead', { address });
      return { kind: 'container', body: edit.body.sub(content as { content?: Element[] }, at), label: `the content control ${at}` };
    }
    throw locationInvalid(location, allowed, `${location} names a container (the body, a table cell or a content control), and ${describeAddress(address)} is ${kindOf(block.element)}`);
  }
  if (address === undefined) throw new ToolError('arguments.invalid', `${location} needs an address`, 'give the block\'s address; Start and End without one mean the body', { argument: 'address' });
  const { block, warnings } = resolveBlock(edit.body, address);
  for (const w of warnings) edit.report.warn(w);
  const at = edit.body.addressOf(block) ?? describeAddress(address);
  const type = typeNameOf(block.element);
  if (type === 'org_docx4j_wml.P') {
    const paragraph = new Paragraph(block.element as Element<wml.P>, block.container, edit.body);
    await checkIfText(edit.body, paragraph, ifText);
    return { kind: 'block', block, address: at, paragraph };
  }
  if (ifText !== undefined) throw new ToolError('arguments.invalid', `if_text names a paragraph, and ${at} is ${kindOf(block.element)}`, 'drop if_text', { argument: 'if_text' });
  if (type === 'org_docx4j_wml.Tbl') return { kind: 'block', block, address: at, table: tableOf(edit.body, block) };
  if (type === 'org_docx4j_wml.SdtBlock' && location !== 'Replace') return { kind: 'block', block, address: at };
  throw new ToolError('target.invalid', `${location} takes a paragraph or a table, and ${at} is ${kindOf(block.element)}`, location === 'Replace' ? 'remove a content control with remove_content_control' : 'address a paragraph or a table; Start and End take a cell or a content control', { address });
}

/** Blocks put at a place (the engine tracks their insertion, everything in them since core-ts 0.3.1); the views of the paragraphs and tables among them. */
export function insertAt(edit: Edit, place: Place, location: Location, elements: Element[]): (Paragraph | Table)[] {
  let container: Element[];
  if (place.kind === 'container') {
    place.body.insertElement(elements, location === 'Start' ? 'Start' : 'End');
    container = place.body.content;
  } else {
    edit.body.insertElement(elements, location === 'Before' ? 'Before' : 'After', place.block);
    container = place.block.container;
  }
  return viewsOf(edit.body, elements, container);
}

/** The paragraph and table views of elements in a container. */
export function viewsOf(body: Body, elements: readonly Element[], container: Element[]): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = [];
  for (const e of elements) {
    if (typeNameOf(e) === 'org_docx4j_wml.P') out.push(new Paragraph(e as Element<wml.P>, container, body));
    else if (typeNameOf(e) === 'org_docx4j_wml.Tbl') out.push(new Table(e as Element<wml.Tbl>, container, body));
  }
  return out;
}

/**
 * The block `Replace` takes out, as a tracked deletion. `made` are the blocks the call put after it, which
 * keep the properties the call gave them: since core-ts 0.3.2 a whole paragraph deleted hands the next one
 * nothing, as Word does (core-ts check 33; UPSTREAM-18, closed).
 */
export function removeReplaced(edit: Edit, place: Place, _made: readonly Element[] = []): void {
  if (place.kind !== 'block') return;
  if (place.paragraph) { edit.report.remove(place.paragraph); place.paragraph.delete(); }
  else if (place.table) { edit.report.remove(place.table); place.table.delete(); }
}

/** "2 paragraphs and a table", for a summary. */
export function counted(views: readonly (Paragraph | Table)[]): string {
  const paragraphs = views.filter((v) => v instanceof Paragraph).length;
  const tables = views.length - paragraphs;
  const parts: string[] = [];
  if (paragraphs > 0) parts.push(paragraphs === 1 ? 'a paragraph' : `${paragraphs} paragraphs`);
  if (tables > 0) parts.push(tables === 1 ? 'a table' : `${tables} tables`);
  return parts.length > 0 ? parts.join(' and ') : 'nothing';
}
