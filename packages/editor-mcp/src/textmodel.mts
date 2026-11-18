// SPDX-License-Identifier: GPL-3.0-or-later
//
// The text model's options on the reading tools (docs/developer/agent.md section 2.13; ED-008 sections 3.1 and 3.2):
// `outline`, `find`, `read` and `extract_text` read the room's document through the engine's text model
// (core-ts's `Texts`, CR-012), the one docx4j's and the Python servers' readers are held to, taking the text
// extraction specification's options by its names. What a tool does not take, and a value the engine does not
// implement, is `arguments.invalid` naming the specification's statement, never accepted silently (P5).
import { z } from 'zod';
import { Texts, TextOptions, TextOptionError, TEXT_OPTION_NAMES, TextResult, type TextBlock } from '@docx4j/core-ts';
import type { TreeBindings } from '@docx4j/editor-model';
import { ToolError } from './errors.mjs';
import type { ToolRoom } from './kit.mjs';

const flag = (what: string) => z.boolean().optional().describe(what);

/** The view and the scope (WD section 10): what text is read. Every reading tool takes them. */
export const SCOPE_SHAPE = {
  view: z.enum(['accepted', 'original', 'markup']).optional().describe('accepted (the default): the document as Accept All leaves it; original: as Reject All leaves it; markup: both, insertions and deletions marked (CriticMarkup in read and extract_text)'),
  include: z.array(z.enum(['headers', 'footers', 'glossary'])).optional().describe('stories beyond the main one, the notes and the text boxes: headers, footers, the glossary'),
  text_boxes: flag('text boxes after their anchoring paragraph (default true)'),
  notes: flag('footnotes and endnotes after the body (default true)'),
  note_references: flag('a note\'s reference in the text, [n] (default true)'),
  comments: flag('comments at their references (default false)'),
  hidden: flag('hidden text (default false)'),
  field_codes: flag('fields\' instructions beside their results (default false)'),
  fields: z.enum(['cached', 'update']).optional().describe('a field\'s cached result (the default); update is not implemented'),
  toc: flag('a table of contents (default false)'),
  drop_caps: z.enum(['join', 'asis']).optional().describe('a drop cap joined to its word (the default), or apart as Word\'s text has it'),
  placeholders: flag('a content control\'s placeholder text (default true)'),
  bindings: z.enum(['bound', 'surface']).optional().describe('a bound control\'s value from its data (the default), or the text the document shows'),
  list_labels: flag('list labels from the numbering (default false)'),
  charts: flag('charts (default true)'),
  smartart: flag('SmartArt (default true)'),
  embedded: flag('embedded packages (default false; true is not implemented)'),
  altchunks: flag('altChunk content (default true)'),
  warnings: flag('warnings where the text is not the whole document (default true)'),
};

/** How the text is written (WD section 10): `read` and `extract_text` alone, which return a serialisation. */
export const SERIALISATION_SHAPE = {
  alt_text: flag('a picture\'s description, [image: ...] (default false)'),
  metadata: flag('the document properties, as YAML front matter in read (default false)'),
  symbols: z.enum(['asis', 'unicode']).optional().describe('a symbol character as stored (the default); unicode is not implemented'),
  small_caps: z.enum(['asis', 'upper']).optional().describe('small capitals as stored (the default), or in capitals'),
  default_language: z.string().min(1).optional().describe('a BCP 47 tag: the language capitals are mapped by where the document names none (default en-US)'),
  math: z.enum(['latex', 'unicodemath']).optional().describe('equations as LaTeX (the default); unicodemath is not implemented'),
  tables: z.enum(['auto', 'tsv', 'grid']).optional().describe('tables in plain text: simple or listed (auto, the default), tab-separated, or a grid'),
  markup_detail: z.enum(['type', 'author', 'author_date']).optional().describe('what a markup mark says beyond its type'),
  page_markers: flag('[page n] where Word\'s pages began when it last laid the document out: current only when the file was paginated after its last edit, never after a keystroke in the editor (Re-paginate) (default false)'),
  layout: z.enum(['none', 'indent']).optional().describe('a paragraph\'s indent as leading spaces in plain text'),
  profile: z.enum(['word-text']).optional().describe('a preset: word-text is Word\'s Save As plain text as the options reach it; an option given beside it wins'),
};

/** The specification's statement for each option a tool may refuse by name (WD section 10). */
const STATEMENTS: Readonly<Record<string, string>> = {
  alt_text: 'REQ-514', metadata: 'REQ-515', symbols: 'REQ-601', small_caps: 'REQ-607', default_language: 'REQ-607', math: 'REQ-604',
  tables: 'REQ-605', markup_detail: 'REQ-702', page_markers: 'REQ-708', layout: 'REQ-709', profile: 'REQ-1002', repair: 'REQ-207',
  duplicate_entries: 'REQ-206',
};

export type ReadingTool = 'outline' | 'find' | 'read' | 'extract_text';

/** Whether a tool takes the serialisation options. */
const SERIALISES: Readonly<Record<ReadingTool, boolean>> = { outline: false, find: false, read: true, extract_text: true };

function refused(message: string, argument: string, statement: string, hint: string): ToolError {
  return new ToolError('arguments.invalid', message, hint, { argument, statement });
}

/**
 * The text model's options from a reading tool's arguments (section 2.13): every option the tool takes, set
 * by its specification name, and refused as the engine refuses it; an argument the tool does not take at all,
 * an option of the specification among them, refused by name. `own` are the tool's other arguments.
 */
export function textOptionsOf(tool: ReadingTool, args: Readonly<Record<string, unknown>>, own: readonly string[]): TextOptions {
  const options = TextOptions.of({});
  const names = new Set<string>(TEXT_OPTION_NAMES);
  const serialisation = new Set(Object.keys(SERIALISATION_SHAPE));
  for (const [name, value] of Object.entries(args)) {
    if (value === undefined || own.includes(name)) continue;
    if (name === 'repair' || name === 'duplicate_entries') {
      throw refused(`${tool} does not take ${name}: it is how a file is loaded, and the room's file was read when it was opened, as Word repairs it (open's repaired)`, name, STATEMENTS[name]!, 'drop it; open reads a file Word repairs as Word does');
    }
    if (serialisation.has(name) && !SERIALISES[tool]) {
      throw refused(`${tool} does not take ${name}: it returns no serialisation (${STATEMENTS[name]})`, name, STATEMENTS[name]!, 'read and extract_text take the serialisation options; outline and find take the view and the scope');
    }
    if (!names.has(name)) throw refused(`${tool} takes no argument ${name}`, name, 'REQ-1001', 'docs/developer/agent.md section 2.13 lists the options by the text extraction specification\'s names');
    try {
      options.set(name, value);
    } catch (e) {
      if (e instanceof TextOptionError) throw refused(e.message, e.option, e.statement, 'the engine refuses this value rather than ignore it; leave the option at its default');
      throw e;
    }
  }
  try {
    options.checkImplemented();
  } catch (e) {
    if (e instanceof TextOptionError) throw refused(e.message, e.option, e.statement, 'the engine refuses this value rather than ignore it; leave the option at its default');
    throw e;
  }
  return options;
}

/**
 * The room's document as it is now through the text model: the person's typing and the session's notes,
 * headers and comments included (`readTree`'s `stories`, its `textPackage`; ED-008 section 3.1).
 */
export async function readText<T>(room: ToolRoom, options: TextOptions, fn: (result: TextResult, b: TreeBindings) => T | Promise<T>): Promise<T> {
  // A comment a peer made is adopted when the comments are read (ED-005 section 12.30).
  if (options.comments) await room.session.stories.comments();
  return room.session.readTree(async (b) => {
    let result: TextResult;
    try {
      result = await Texts.of(b.textPackage, options);
    } catch (e) {
      if (e instanceof TextOptionError) throw refused(e.message, e.option, e.statement, 'the engine refuses this value rather than ignore it; leave the option at its default');
      throw e;
    }
    return fn(result, b);
  }, { stories: true });
}

/** Every block of a result in reading order, a table followed by its cells' blocks at every depth. */
export function* blocksIn(blocks: readonly TextBlock[]): Generator<TextBlock> {
  for (const block of blocks) {
    yield block;
    for (const row of block.rows) for (const cell of row.cells) yield* blocksIn(cell.blocks);
  }
}

/** Whether a block is at an address, or under it, or joined to the block there. */
export function blockAt(block: TextBlock, address: string): boolean {
  const at = (a: string): boolean => a === address || a.startsWith(`${address}/`);
  return at(block.address) || block.joined.some(at);
}

/**
 * The blocks an address names: the result's own where one is at or under it (a table whole, a paragraph with
 * those joined to it), else the blocks nested in a table at or under it (a cell's paragraph).
 */
export function blocksAt(result: TextResult, address: string): TextBlock[] {
  const top = result.blocks.filter((b) => blockAt(b, address));
  if (top.length > 0) return top;
  const out: TextBlock[] = [];
  const visit = (blocks: readonly TextBlock[]): void => {
    for (const block of blocks) {
      if (blockAt(block, address)) { out.push(block); continue; }
      for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks);
    }
  };
  visit(result.blocks);
  return out;
}

/** A result of some of another's blocks, its comments, warnings, metadata and options kept, for one serialisation. */
export function resultOf(result: TextResult, blocks: readonly TextBlock[]): TextResult {
  return new TextResult(blocks, result.comments, result.warnings, result.metadata, result.options);
}

/** A warning as a reading tool returns it: the model's (REQ-801) by its JSON, the editor's own with a message. */
export type ReadWarning = { kind: string; address?: string; selector?: string; detail?: number | string } | { kind: string; message: string };

export function warningsOf(result: TextResult): ReadWarning[] {
  return result.warnings.map((w) => {
    const json = w.toJSON() as { kind: string; address?: string; selector?: string; detail?: number | string };
    return Object.fromEntries(Object.entries(json).filter(([, v]) => v !== undefined && v !== null)) as ReadWarning;
  });
}

/** A warning in words, for the text a tool answers with. */
export function warningText(w: ReadWarning): string {
  if ('message' in w) return w.message;
  return `${w.kind}${w.address ? ` at ${w.address}` : ''}${w.selector ? ` (${w.selector})` : ''}`;
}
