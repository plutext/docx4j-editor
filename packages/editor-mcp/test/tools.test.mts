// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 8.9, step F3: each tool held to its contract (docs/developer/agent.md) through an MCP client
// over the SDK's in-memory transport, against a companion with real rooms: the tools listed, the
// room's tools, reading, editing (Markdown in, change reports, if_text, dry runs, errors), comments,
// tracked changes and their resolution, content controls, and the paths.
import { afterEach, beforeEach, describe, expect, test } from 'vitest';
import { mkdtemp, readFile, rm, writeFile, stat } from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { WordprocessingMLPackage, Texts, PlainText, MarkdownExporter } from '@docx4j/core-ts';
import { createCompanion, type AgentRoom, type Companion } from '../src/index.mjs';
import { COMPANION_VERSION } from '../src/tools.mjs';
import { PHASE0 as phase0, suitePresent, loadRunner } from '../../editor-model/test/text-extraction-suite.mjs';

const fixtures = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'editor-model', 'test', 'fixtures');
/** The text extraction suite's probe (ED-008 step 2): its fixtures and goldens, docx4j's. */

interface Env { companion: Companion; client: Client; root: string }
const cleanups: (() => Promise<void> | void)[] = [];
afterEach(async () => { for (const c of cleanups.splice(0).reverse()) await c(); });

/** A companion over a fresh root holding the fixtures named, and a client connected in memory. */
async function start(files: Record<string, string>, options: { allowResolve?: boolean; maxInlineChars?: number } = {}, from = fixtures): Promise<Env> {
  const root = await mkdtemp(join(tmpdir(), 'tools-root-'));
  const app = await mkdtemp(join(tmpdir(), 'tools-app-'));
  const data = await mkdtemp(join(tmpdir(), 'tools-rooms-'));
  await writeFile(join(app, 'index.html'), '<!doctype html><title>docx4j editor</title>');
  for (const [name, fixture] of Object.entries(files)) await writeFile(join(root, name), await readFile(join(from, fixture)));
  const companion = createCompanion({ roots: [root], port: 0, app, data, log: () => undefined, ...options });
  await companion.listen();
  const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
  const server = companion.mcp();
  await server.connect(serverSide);
  const client = new Client({ name: 'contract-test', version: '1' });
  await client.connect(clientSide);
  cleanups.push(() => rm(root, { recursive: true, force: true }), () => rm(app, { recursive: true, force: true }), () => rm(data, { recursive: true, force: true }));
  cleanups.push(async () => { await client.close(); await companion.close(); });
  return { companion, client, root };
}

type Structured = Record<string, unknown>;

async function call(env: Env, name: string, args: Record<string, unknown> = {}): Promise<{ isError: boolean; data: Structured; text: string }> {
  const result = await env.client.callTool({ name, arguments: args }) as { isError?: boolean; structuredContent?: Structured; content?: { text?: string }[] };
  return { isError: result.isError === true, data: result.structuredContent ?? {}, text: result.content?.[0]?.text ?? '' };
}

/** A call that must succeed: its structured content. */
async function ok(env: Env, name: string, args: Record<string, unknown> = {}): Promise<Structured> {
  const r = await call(env, name, args);
  expect(r.isError, `${name}: ${r.text}`).toBe(false);
  return r.data;
}

/** A call that must fail: its error's structured content. */
async function fails(env: Env, name: string, args: Record<string, unknown>, code: string): Promise<Structured> {
  const r = await call(env, name, args);
  expect(r.isError, `${name} should fail with ${code}`).toBe(true);
  expect(r.data['code'], r.text).toBe(code);
  expect(typeof r.data['hint']).toBe('string');
  return r.data;
}

async function openLease(env: Env): Promise<string> {
  return (await ok(env, 'open', { path: 'lease.docx' }))['document'] as string;
}

interface Block { address: string; paraId?: string; style: string; text: string; kind: string }
async function outline(env: Env, document: string): Promise<Block[]> {
  return (await ok(env, 'outline', { document, budget: 1000 }))['blocks'] as Block[];
}
const blockWith = (blocks: Block[], text: string): Block => {
  const b = blocks.find((x) => x.kind === 'paragraph' && x.text.includes(text));
  if (!b) throw new Error(`no paragraph with ${text}`);
  return b;
};

/** The main part of the room's document as a save would write it. */
async function mainXml(env: Env, document: string): Promise<string> {
  const room = env.companion.rooms.get(document) as AgentRoom;
  const saved = await room.session.save();
  const main = (await WordprocessingMLPackage.load(saved.bytes)).getMainDocumentPart();
  await main.getContents();
  return main.getXml();
}

test('COMPANION_VERSION is the package\'s version (scripts/set-version.mjs sets both)', () => {
  expect(COMPANION_VERSION).toBe(JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8')).version);
});

describe('the tools (docs/developer/agent.md, ED-005 section 8.9 step F3)', () => {
  let env: Env;
  beforeEach(async () => { env = await start({ 'lease.docx': 'tracked-changes.docx', 'tables.docx': 'tables.docx', 'bound.docx': 'opendope/binding-simple.docx', 'boxes.docx': 'loadAndSave.docx', 'mcblock.docx': 'mc/mc-block.docx' }); });

  test('every tool of the contract is listed, with an input schema', async () => {
    const { tools } = await env.client.listTools();
    const names = tools.map((t) => t.name).sort();
    expect(names).toEqual([
      'accept_tracked_changes', 'bind', 'close', 'comment', 'delete', 'edit', 'extract_text', 'find', 'get_comments', 'get_tracked_changes', 'insert', 'insert_content_control',
      'insert_ooxml', 'list_documents', 'open', 'outline', 'read', 'reject_tracked_changes', 'remove_content_control', 'replace_text', 'save', 'set_content_control',
      'set_format', 'set_style', 'track_changes',
    ]);
    for (const t of tools) expect(t.inputSchema.type).toBe('object');
  });

  test('open: a room and its link; the same file answers the same room; a path outside the roots or missing is refused', async () => {
    const first = await ok(env, 'open', { path: 'lease.docx' });
    expect(first['link']).toMatch(/#room=.*&token=.*&relay=/);
    expect((await ok(env, 'open', { path: join(env.root, 'lease.docx') }))['document']).toBe(first['document']);
    await fails(env, 'open', { path: '/etc/passwd' }, 'path.not_allowed');
    await fails(env, 'open', { path: 'missing.docx' }, 'path.not_found');
    const listed = (await ok(env, 'list_documents'))['documents'] as { document: string; people: { name: string; agent: boolean }[] }[];
    expect(listed.map((d) => d.document)).toEqual([first['document']]);
    expect(listed[0]!.people).toContainEqual({ name: 'Claude', agent: true });
    await fails(env, 'outline', { document: 'nope' }, 'document.not_open');
    await ok(env, 'close', { document: first['document'] });
    await fails(env, 'outline', { document: first['document'] }, 'document.not_open');
  });

  test('track_changes: always the agent\'s, the document\'s own setting reported; Off refused', async () => {
    const document = await openLease(env);
    expect(await ok(env, 'track_changes', { document })).toEqual({ mode: 'TrackAll', author: 'Claude', document_mode: 'Off' });
    await fails(env, 'track_changes', { document, mode: 'Off' }, 'tracking.not_allowed');
  });

  test('a block in a kept mc:AlternateContent\'s selected branch is listed at its address and refused for editing (ED-009 section 9.5; core-ts 0.6.0, CR-013)', async () => {
    const document = (await ok(env, 'open', { path: 'mcblock.docx' }))['document'] as string;
    const outline = await ok(env, 'outline', { document });
    const blocks = outline['blocks'] as { address: string; text: string }[];
    expect(blocks.map((b) => b.address)).toEqual(['body/0', 'body/1', 'body/2/0']);
    expect(blocks[2]!.text).toBe('w16sdtfl');
    const error = await fails(env, 'edit', { document, address: 'body/2/0', text: 'Edited' }, 'target.invalid');
    expect(String(error['hint'])).toContain('mc:AlternateContent');
  });

  test('outline, read and find: addresses, paragraph ids, the text as accepted', async () => {
    const document = await openLease(env);
    const blocks = await outline(env, document);
    const title = blockWith(blocks, 'Docx sample document');
    expect(title).toMatchObject({ address: 'body/0', style: 'Title', kind: 'paragraph' });
    expect(blocks.some((b) => b.kind === 'table')).toBe(true);
    // Deleted text is out, inserted text in (section 2.2).
    const tracking = blockWith(blocks, 'Here is some change tracking');
    expect(tracking.text).toBe('Here is some change tracking. An insertion Followed by.');
    const budget = await ok(env, 'outline', { document, budget: 3 });
    expect((budget['blocks'] as unknown[]).length).toBe(3);
    expect(budget['truncated']).toBe(true);

    const read = await ok(env, 'read', { document });
    expect(read['markdown']).toContain('<!-- body/0 -->\nDocx sample document');
    expect(read['markdown']).toContain('# Tables');
    expect(read['markdown']).toContain('|Cell text||Shaded grey|');
    const one = await ok(env, 'read', { document, address: tracking.address });
    expect(one['markdown']).toBe(`<!-- ${tracking.address} -->\nHere is some change tracking. An insertion Followed by.`);

    const found = await ok(env, 'find', { document, text: 'insertion' });
    expect(found['hits']).toContainEqual(expect.objectContaining({ address: tracking.address, offset: 33, length: 9 }));
    const byXPath = await ok(env, 'find', { document, xpath: '//w:tbl' });
    expect((byXPath['hits'] as { kind: string }[]).every((h) => h.kind === 'table')).toBe(true);
    expect((byXPath['hits'] as unknown[]).length).toBeGreaterThan(0);
    // A hit inside a text box (both branches of its kept mc:AlternateContent, ED-005 section 12.106) is the
    // paragraph holding the box, once; before, the box's own paragraph had no address and the hit was dropped.
    const boxes = (await ok(env, 'open', { path: 'boxes.docx' }))['document'] as string;
    const inBox = await ok(env, 'find', { document: boxes, xpath: '//w:txbxContent//w:p' });
    expect(inBox['total']).toBe(1);
    // one hit, at the text box's own address (core-ts 0.6.0, CR-012 D4): the Fallback's twin folded into it (ED-008 section 8.8 item 15)
    expect(inBox['hits']).toEqual([expect.objectContaining({ kind: 'paragraph', address: expect.stringMatching(/^body\/\d+\/textbox:\d+\/\d+$/), snippet: 'Hello world' })]);
    await fails(env, 'find', { document, text: 'x', xpath: '//w:p' }, 'arguments.invalid');
    await fails(env, 'read', { document, address: 'body/9999' }, 'address.not_found');
  });

  test('edit: Markdown replaces a paragraph as the agent\'s tracked change, with headings, formatting, lists, a link and a table; the report names them', async () => {
    const document = await openLease(env);
    const target = blockWith(await outline(env, document), 'Some text.');
    const markdown = '## New heading\n\nA **bold** and *italic* ~~old~~ line with a [link](https://example.com/).\n\n- one\n- two\n  - two a\n\n1. first\n2. second\n\n| A | B |\n| --- | --- |\n| 1 | 2 |';
    const report = await ok(env, 'edit', { document, address: target.address, markdown, if_text: 'Some text.' });
    expect(report['summary']).toBe(`Replaced ${target.address} with 7 paragraphs and a table`);
    expect(report['removed']).toEqual([{ address: target.address }]);
    const created = report['created'] as { text: string }[];
    expect(created.map((c) => c.text)).toEqual(['New heading', 'A bold and italic old line with a link.', 'one', 'two', 'two a', 'first', 'second', 'A\nB\n1\n2']);
    expect((report['revisions'] as { kind: string; author: string }[]).every((r) => r.author === 'Claude')).toBe(true);
    // As Office JS lists them (the engine since core-ts 0.3.1): a paragraph mark is "\r", touching insertions across
    // paragraphs are one change, a link's holder parts them (check 30), and a table groups with what touches it (check 31).
    expect(report['revisions']).toEqual([
      expect.objectContaining({ kind: 'deletion', text: 'Some text.\r' }),
      expect.objectContaining({ kind: 'insertion', address: 'body/5', text: 'New heading\rA bold and italic old line with a ' }),
      expect.objectContaining({ kind: 'insertion', address: 'body/6', text: 'link' }),
      expect.objectContaining({ kind: 'insertion', address: 'body/6', text: '.\rone\rtwo\rtwo a\rfirst\rsecond\rA\tB\r\n1\t2\r\n' }),
    ]);
    expect(report['parts']).toEqual(expect.arrayContaining(['/word/numbering.xml', '/word/_rels/document.xml.rels']));
    const xml = await mainXml(env, document);
    // The heading keeps its style: the replaced paragraph's deletion hands it nothing (core-ts 0.3.2; UPSTREAM-18, closed).
    expect(xml).toMatch(/<w:p><w:pPr><w:pStyle w:val="Heading2"\/><w:rPr><w:ins [^>]*\/><\/w:rPr><\/w:pPr><w:ins [^>]*><w:r><w:t>New heading</);
    expect(xml).not.toMatch(/w:pPrChange/);
    // The link's text is in the insertion too (the engine since core-ts 0.3.1): <w:hyperlink><w:ins><w:r>.
    expect(xml).toMatch(/<w:hyperlink [^>]*r:id="[^"]+"[^>]*><w:ins w:id="\d+"[^>]* w:author="Claude"[^>]*><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/><\/w:rPr><w:t>link<\/w:t>/);
    expect(xml).toMatch(/<w:numPr><w:ilvl w:val="1"\/><w:numId w:val="\d+"\/><\/w:numPr>/);
    expect(xml).toMatch(/<w:del w:id="\d+"[^>]* w:author="Claude"[^>]*><w:r><w:delText>Some text\.<\/w:delText>/);
    expect(xml).toMatch(/<w:ins w:id="\d+"[^>]* w:author="Claude"[^>]*>(?:(?!<\/w:ins>).)*<w:r><w:rPr><w:b\/><w:bCs\/><\/w:rPr><w:t>bold<\/w:t>/);
    // The document's own tracking setting was left alone (ED-005 section 8.5).
    expect(xml).not.toContain('w:trackRevisions');
    const room = env.companion.rooms.get(document) as AgentRoom;
    expect(room.session.trackingMode).toBe('Off');
  });

  test('insert: After by default, Before, Start and End of the body; Replace is not insert\'s; plain text a paragraph per line', async () => {
    const document = await openLease(env);
    const title = blockWith(await outline(env, document), 'Docx sample document');
    const after = await ok(env, 'insert', { document, address: title.address, text: 'Line one\nLine two' });
    expect(after['summary']).toBe(`Inserted 2 paragraphs after ${title.address}`);
    expect((after['created'] as { address: string }[]).map((c) => c.address)).toEqual(['body/1', 'body/2']);
    const start = await ok(env, 'insert', { document, text: 'At the very start', location: 'Start' });
    expect(start['created']).toEqual([expect.objectContaining({ address: 'body/0', text: 'At the very start' })]);
    const end = await ok(env, 'insert', { document, markdown: 'At the end', location: 'End' });
    expect(end['summary']).toBe('Inserted a paragraph at the end of the body');
    await fails(env, 'insert', { document, address: title.address, text: 'x', location: 'Replace' }, 'location.invalid');
    await fails(env, 'insert', { document, address: title.address, text: 'x', location: 'Start' }, 'location.invalid');
    await fails(env, 'insert', { document, markdown: '```\ncode\n```' }, 'markdown.unsupported');
    await fails(env, 'insert', { document, markdown: '![a](b.png)' }, 'markdown.unsupported');
    await fails(env, 'insert', { document, markdown: 'x', text: 'y' }, 'arguments.invalid');
  });

  test('if_text: a stale paragraph is refused with its current text, and nothing is written', async () => {
    const document = await openLease(env);
    const target = blockWith(await outline(env, document), 'Some text.');
    const room = env.companion.rooms.get(document) as AgentRoom;
    const before = room.session.doc;
    const error = await fails(env, 'edit', { document, address: target.address, text: 'New', if_text: 'Some text (as read earlier).' }, 'target.stale');
    expect(error['current']).toEqual({ address: target.address, text: 'Some text.' });
    expect(room.session.doc.eq(before)).toBe(true);
    await fails(env, 'delete', { document, address: blockWith(await outline(env, document), 'Tables').address.replace(/\d+$/, (n) => String(Number(n) + 2)), if_text: 'x' }, 'arguments.invalid');
  });

  test('a dry run reports without writing', async () => {
    const document = await openLease(env);
    const target = blockWith(await outline(env, document), 'Some text.');
    const room = env.companion.rooms.get(document) as AgentRoom;
    const before = room.session.doc;
    const report = await ok(env, 'edit', { document, address: target.address, markdown: '# Heading\n\n- item', dry_run: true });
    expect(report['dry_run']).toBe(true);
    expect(report['created']).toEqual([expect.objectContaining({ text: 'Heading' }), expect.objectContaining({ text: 'item' })]);
    expect(room.session.doc.eq(before)).toBe(true);
    expect((await room.session.save()).changed).toBe(false);
  });

  test('delete: a span\'s text, a paragraph, a table; the agent\'s own insertion goes for real', async () => {
    const document = await openLease(env);
    let blocks = await outline(env, document);
    const target = blockWith(blocks, 'Some text.');
    const span = await ok(env, 'delete', { document, address: target.address, anchor: 'text' });
    expect(span['summary']).toBe(`Deleted "text" in ${target.address}`);
    expect(span['touched']).toEqual([expect.objectContaining({ address: target.address, text: 'Some .' })]);
    expect(span['revisions']).toEqual([{ kind: 'deletion', author: 'Claude', address: target.address, text: 'text' }]);
    const para = await ok(env, 'delete', { document, address: target.address });
    expect(para['removed']).toEqual([{ address: target.address }]);
    blocks = await outline(env, document);
    const table = blocks.find((b) => b.kind === 'table')!;
    expect((await ok(env, 'delete', { document, address: table.address }))['summary']).toBe(`Deleted the table ${table.address}`);
    // The agent's own paragraph, inserted then deleted: gone, not marked.
    const made = await ok(env, 'insert', { document, text: 'Mine alone', location: 'End' });
    const mine = (made['created'] as { address: string }[])[0]!.address;
    await ok(env, 'delete', { document, address: mine });
    expect((await outline(env, document)).some((b) => b.text === 'Mine alone')).toBe(false);
    await fails(env, 'delete', { document, address: target.address, anchor: 'nowhere' }, 'range.invalid');
  });

  test('replace_text: every match, one change, the touched paragraphs with their new text; $1 with match_regexp; if_text with a paragraph scope', async () => {
    const document = await openLease(env);
    const all = await ok(env, 'replace_text', { document, find: 'indented', replace: 'INDENTED', match_whole_word: true });
    expect(all['matches']).toBe(20);
    expect(all['summary']).toBe('Replaced 20 matches of "indented" with "INDENTED" in one paragraph');
    const touched = all['touched'] as { text: string; address: string }[];
    expect(touched[0]!.text).toContain('INDENTED INDENTED');
    const regexp = await ok(env, 'replace_text', { document, find: '(Left) (indent)', replace: '$2 $1', match_regexp: true });
    expect((regexp['touched'] as { text: string }[])[0]!.text).toContain('indent Left');
    const target = blockWith(await outline(env, document), 'Some text.');
    await fails(env, 'replace_text', { document, find: 'Some', replace: 'Any', scope: target.address, if_text: 'stale' }, 'target.stale');
    await fails(env, 'replace_text', { document, find: 'Some', replace: 'Any', if_text: 'Some text.' }, 'arguments.invalid');
    const none = await ok(env, 'replace_text', { document, find: 'no such words', replace: 'x' });
    expect(none['summary']).toBe('No match for "no such words"');
  });

  test('set_style and set_format: a style by name added from the defaults, properties in points, formatting over a span; tracked', async () => {
    const document = await openLease(env);
    const target = blockWith(await outline(env, document), 'Some text.');
    const styled = await ok(env, 'set_style', { document, address: target.address, style: 'Heading 3', alignment: 'Centered', space_after: 12 });
    expect(styled['revisions']).toContainEqual(expect.objectContaining({ kind: 'properties', author: 'Claude' }));
    const formatted = await ok(env, 'set_format', { document, address: target.address, anchor: 'text', bold: true, color: '#ff0000' });
    expect(formatted['revisions']).toContainEqual(expect.objectContaining({ kind: 'formatting', text: 'text' }));
    const xml = await mainXml(env, document);
    expect(xml).toMatch(/<w:pStyle w:val="Heading3"\/>/);
    expect(xml).toMatch(/<w:jc w:val="center"\/>/);
    expect(xml).toMatch(/<w:b\/>[^]*<w:color w:val="FF0000"\/>[^]*<w:rPrChange /);
    await fails(env, 'set_style', { document, address: target.address, style: 'No Such Style At All' }, 'style.not_found');
    await fails(env, 'set_format', { document, address: target.address }, 'arguments.invalid');
    // A measurement past what Word stores: the engine's refusal (core-ts 0.7.0, CR-002 section 65), naming the argument.
    const tooFar = await fails(env, 'set_style', { document, address: target.address, left_indent: 400000 }, 'arguments.invalid');
    expect(tooFar['argument']).toBe('left_indent');
    expect(String(tooFar['message'])).toMatch(/Word 365 would read it as/);
    const table = (await outline(env, document)).find((b) => b.kind === 'table')!;
    await fails(env, 'set_style', { document, address: table.address, alignment: 'Left' }, 'target.invalid');
  });

  test('insert_ooxml: a fragment inserted and tracked; one that does not parse, or names a relationship it does not bring, refused', async () => {
    const document = await openLease(env);
    const title = blockWith(await outline(env, document), 'Docx sample document');
    const xml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>From OOXML</w:t></w:r></w:p>';
    const report = await ok(env, 'insert_ooxml', { document, address: title.address, ooxml: xml, location: 'After' });
    expect(report['created']).toEqual([expect.objectContaining({ address: 'body/1', text: 'From OOXML' })]);
    expect(report['revisions']).toEqual([expect.objectContaining({ kind: 'insertion', text: 'From OOXML\r' })]);
    // A block content control is the agent's insertion with its paragraph, in Word's form (core-ts 0.3.1, check 30):
    // insertion range markers around the control, its paragraph's mark and runs inserted.
    const control = '<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sdtPr><w:tag w:val="inserted"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Inside a control</w:t></w:r></w:p></w:sdtContent></w:sdt>';
    const wrapped = await ok(env, 'insert_ooxml', { document, address: title.address, ooxml: control, location: 'After' });
    expect(wrapped['warnings']).toEqual(['one inserted block is neither a paragraph nor a table (a content control, say): read shows it']);
    expect(await mainXml(env, document)).toMatch(/<w:customXmlInsRangeStart [^>]*w:author="Claude"[^>]*\/><w:sdt><w:sdtPr><w:tag w:val="inserted"\/><\/w:sdtPr><w:sdtContent><w:customXmlInsRangeEnd [^>]*\/><w:p><w:pPr><w:rPr><w:ins [^>]*w:author="Claude"[^>]*\/><\/w:rPr><\/w:pPr><w:ins [^>]*w:author="Claude"[^>]*><w:r><w:t>Inside a control<\/w:t>/);
    await fails(env, 'insert_ooxml', { document, address: title.address, ooxml: '<w:p><unclosed>', location: 'After' }, 'ooxml.invalid');
    await fails(env, 'insert_ooxml', { document, address: title.address, ooxml: '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:hyperlink r:id="rId999"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p>', location: 'After' }, 'ooxml.invalid');
  });

  test('insert_ooxml: a whole package at the body\'s Start and End goes in as blocks of its own, joining no paragraph (docs/developer/agent.md section 2.5)', async () => {
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const part = (name: string, type: string, xml: string): string => `<pkg:part pkg:name="${name}" pkg:contentType="${type}"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`;
    const pkg = '<?xml version="1.0" standalone="yes"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
      + part('/_rels/.rels', 'application/vnd.openxmlformats-package.relationships+xml', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
      + part('/word/document.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Package first</w:t></w:r></w:p><w:p><w:r><w:t>Package last</w:t></w:r></w:p></w:body></w:document>`)
      + '</pkg:package>';
    const document = await openLease(env);
    const before = (await outline(env, document)).map((b) => b.text);
    await ok(env, 'insert_ooxml', { document, ooxml: pkg, location: 'End' });
    await ok(env, 'insert_ooxml', { document, ooxml: pkg, location: 'Start' });
    const after = (await outline(env, document)).map((b) => b.text);
    expect(after).toEqual(['Package first', 'Package last', ...before, 'Package first', 'Package last']);
  });

  test('comment: a comment on words, a reply, the thread resolved; never tracked; get_comments lists it', async () => {
    const document = await openLease(env);
    const target = blockWith(await outline(env, document), 'Some text.');
    const made = await ok(env, 'comment', { document, address: target.address, anchor: 'text', text: 'Is this **right**?' });
    const id = (made['comment'] as { id: string }).id;
    expect(made['revisions']).toEqual([]);
    expect(made['parts']).toContain('/word/comments.xml');
    await ok(env, 'comment', { document, reply_to: id, text: 'Yes.', resolve: true });
    const threads = (await ok(env, 'get_comments', { document }))['threads'] as { id: string; author: string; resolved: boolean; anchor: string; address: string; text: string; replies: { text: string }[] }[];
    const thread = threads.find((t) => t.id === id)!;
    expect(thread).toMatchObject({ author: 'Claude', resolved: true, anchor: 'text', address: target.address, text: 'Is this right?' });
    expect(thread.replies.map((r) => r.text)).toEqual(['Yes.']);
    await fails(env, 'comment', { document, reply_to: '424242', text: 'x' }, 'address.not_found');
    await fails(env, 'comment', { document, address: target.address, text: '# Heading' }, 'markdown.unsupported');
    const xml = await mainXml(env, document);
    expect(xml).toMatch(/<w:commentRangeStart w:id="\d+"\/>/);
    expect(made['warnings'] ?? []).toEqual([]);
  });

  test('comment: on words in a pending insertion, the anchor stays on those words, in Word\'s form (core-ts 0.3.1; ED-005 sections 12.29 and 12.50)', async () => {
    const document = await openLease(env);
    const target = blockWith(await outline(env, document), 'Left indent');
    const edited = await ok(env, 'edit', { document, address: target.address, markdown: 'The tenant pays **monthly**.' });
    const paragraph = (edited['created'] as { address: string }[])[0]!;
    const made = await ok(env, 'comment', { document, address: paragraph.address, anchor: 'monthly', text: 'Was quarterly.' });
    // No warning: the engine puts the start marker inside the w:ins and splits it after the words, as Word does.
    expect(made['warnings']).toEqual([]);
    expect(made['revisions']).toEqual([]);
    // The plain reference run parts nothing: one insertion with its mark, as Office JS lists it (check 32).
    const listed = (await ok(env, 'get_tracked_changes', { document, author: 'Claude' }))['changes'] as { kind: string; text: string }[];
    expect(listed.map((c) => [c.kind, c.text])).toEqual([['deletion', 'Left indent\r'], ['insertion', 'The tenant pays monthly.\r']]);
    const threads = (await ok(env, 'get_comments', { document }))['threads'] as { id: string; anchor: string }[];
    expect(threads.find((t) => t.id === (made['comment'] as { id: string }).id)!.anchor).toBe('monthly');
    expect(await mainXml(env, document)).toMatch(/<w:ins [^>]*w:author="Claude"[^>]*><w:r>(?:<w:rPr>(?:(?!<\/w:rPr>).)*<\/w:rPr>)?<w:t xml:space="preserve">The tenant pays <\/w:t><\/w:r><w:commentRangeStart w:id="\d+"\/>/);
  });

  test('get_tracked_changes: as the review panel lists them, with keys; the agent may not resolve without --allow-resolve', async () => {
    const document = await openLease(env);
    const before = (await ok(env, 'get_tracked_changes', { document }))['changes'] as { kind: string; author: string; text: string; key?: string; address: string }[];
    expect(before.map((c) => [c.kind, c.author, c.text])).toEqual([['insertion', 'Jason Harrop', 'An insertion'], ['deletion', 'Jason Harrop', ' A deletion']]);
    const target = blockWith(await outline(env, document), 'Some text.');
    await ok(env, 'insert', { document, address: target.address, text: 'By the agent' });
    const after = (await ok(env, 'get_tracked_changes', { document, author: 'Claude' }))['changes'] as { kind: string; text: string; address: string }[];
    expect(after).toEqual([expect.objectContaining({ kind: 'insertion', text: 'By the agent\r' })]);
    await fails(env, 'accept_tracked_changes', { document }, 'resolve.not_allowed');
    await fails(env, 'reject_tracked_changes', { document, key: 'x' }, 'resolve.not_allowed');
  });

  test('content controls: wrapped, set, bound and removed, none of it tracked', async () => {
    const document = (await ok(env, 'open', { path: 'bound.docx' }))['document'] as string;
    const blocks = await outline(env, document);
    const made = await ok(env, 'insert_content_control', { document, address: blocks[1]!.address, kind: 'RichText', tag: 'second', title: 'The second' });
    const control = made['control'] as { id: number; tag: string; kind: string };
    expect(control).toMatchObject({ tag: 'second', kind: 'RichText' });
    expect(made['revisions']).toEqual([]);
    const set = await ok(env, 'set_content_control', { document, control: { tag: 'second' }, title: 'Renamed', lock: 'control' });
    expect(set['control']).toMatchObject({ title: 'Renamed' });
    await fails(env, 'set_content_control', { document, control: { tag: 'second' }, checked: true }, 'target.invalid');
    // The fixture's first inline control, bound to element 1, bound again to element 3.
    const bound = await ok(env, 'bind', { document, control: { id: -1309316605 }, xpath: '/myxml[1]/element3[1]' });
    expect(bound['summary']).toBe('Bound the content control -1309316605 to /myxml[1]/element3[1]');
    const read = await ok(env, 'read', { document, address: blocks[0]!.address });
    expect(read['markdown']).toContain('element 3 contents');
    await fails(env, 'bind', { document, control: { id: -1309316605 }, xpath: '/myxml[1]/element9[1]' }, 'binding.no_node');
    await fails(env, 'bind', { document, control: { id: -1309316605 }, xpath: '/myxml' }, 'binding.document_element');
    const removed = await ok(env, 'remove_content_control', { document, control: { id: control.id } });
    expect(removed['summary']).toBe(`Removed the content control ${control.id}, keeping its content`);
    await fails(env, 'remove_content_control', { document, control: { id: control.id } }, 'address.not_found');
  });

  test('save: to its file, or to output_path under a root; an existing other file needs overwrite; outside the roots refused', async () => {
    const document = await openLease(env);
    await ok(env, 'insert', { document, text: 'Saved by the agent', location: 'End' });
    const saved = await ok(env, 'save', { document });
    expect(saved['path']).toBe(join(env.root, 'lease.docx'));
    expect((await stat(join(env.root, 'lease.docx'))).size).toBe(saved['bytes']);
    await ok(env, 'save', { document, output_path: 'copy.docx' });
    await fails(env, 'save', { document, output_path: 'tables.docx' }, 'path.exists');
    await ok(env, 'save', { document, output_path: 'tables.docx', overwrite: true });
    await fails(env, 'save', { document, output_path: '/tmp/outside.docx' }, 'path.not_allowed');
    const reopened = await WordprocessingMLPackage.load(new Uint8Array(await readFile(join(env.root, 'copy.docx'))));
    expect(await reopened.getMainDocumentPart().getXml()).toContain('Saved by the agent');
  });

  test('Stop: every tool on a stopped room answers agent.stopped, reading tools included', async () => {
    const document = await openLease(env);
    (env.companion.rooms.get(document) as AgentRoom).stopped = true;
    await fails(env, 'outline', { document }, 'agent.stopped');
    await fails(env, 'edit', { document, address: 'body/0', text: 'x' }, 'agent.stopped');
  });
});

describe('the tools with --allow-resolve and the inline cap', () => {
  test('accept by key, reject all of an author\'s; through the four commands', async () => {
    const env = await start({ 'lease.docx': 'tracked-changes.docx' }, { allowResolve: true });
    const document = await openLease(env);
    const changes = (await ok(env, 'get_tracked_changes', { document }))['changes'] as { key: string; text: string }[];
    const accepted = await ok(env, 'accept_tracked_changes', { document, key: changes[0]!.key });
    expect(accepted['resolved']).toBe(1);
    await fails(env, 'accept_tracked_changes', { document, key: changes[0]!.key }, 'target.stale');
    const rejected = await ok(env, 'reject_tracked_changes', { document, author: 'Jason Harrop' });
    expect(rejected['resolved']).toBe(1);
    expect((await ok(env, 'get_tracked_changes', { document }))['changes']).toEqual([]);
    const text = blockWith(await outline(env, document), 'Here is some change tracking').text;
    expect(text).toBe('Here is some change tracking. An insertion Followed by A deletion.');
  });

  test('a read past the inline cap is truncated, or written to output_path', async () => {
    const env = await start({ 'lease.docx': 'tracked-changes.docx' }, { maxInlineChars: 500 });
    const document = await openLease(env);
    const cut = await ok(env, 'read', { document });
    expect(cut['truncated']).toBe(true);
    expect((cut['markdown'] as string).length).toBe(500);
    expect(cut['chars']).toBeGreaterThan(500);
    const written = await ok(env, 'read', { document, output_path: 'lease.md' });
    expect(written['path']).toBe(join(env.root, 'lease.md'));
    expect((await readFile(join(env.root, 'lease.md'), 'utf8')).length).toBe(written['chars']);
  });
});

describe('the reading tools over the text model (ED-008 step 3, docs/developer/agent.md section 2.13)', () => {
  /** Markdown without the address comments the tool adds before each block (section 2.6). */
  const withoutAddresses = (md: string): string => md.split('\n').filter((line) => !/^<!-- \S+( paraId=[0-9A-F]+)? -->$/.test(line)).join('\n');
  const trimmed = (text: string): string => text.replace(/\n+$/, '');

  // The suite's files: the public repository leaves them out, and there this test skips (text-extraction-suite.mts).
  test.skipIf(!suitePresent)('read and extract_text in each view: the engine\'s own Markdown and plain text of the file, exactly, and the suite\'s goldens after REQ-1201\'s normalisation', async () => {
    const { normalize, imageDestinations, table } = await loadRunner();
    const probes = ['extraction-probe.docx', 'extraction-probe-external.docm', 'extraction-probe-structure.docx'];
    const env = await start(Object.fromEntries(probes.map((f) => [f, f])), { maxInlineChars: 1_000_000 }, phase0);
    const documents = new Map<string, string>();
    for (const f of probes) documents.set(f, (await ok(env, 'open', { path: f }))['document'] as string);
    let golden = 0;
    for (const row of table(join(phase0, 'goldens.tsv'))) {
      const fixture = row['fixture']!;
      const view = row['view']!;
      const pkg = await WordprocessingMLPackage.load(await readFile(join(phase0, fixture)));
      const document = documents.get(fixture)!;
      let ours: string;
      if (row['format'] === 'md') {
        const read = await ok(env, 'read', { document, view });
        ours = withoutAddresses(read['markdown'] as string);
        // What the tool adds is the address comments, and nothing else (ED-008 section 5).
        expect(ours, `${fixture} ${view} md`).toBe(trimmed(await MarkdownExporter.export(await Texts.of(pkg, { view: view as 'accepted' }), pkg)));
      } else {
        const extracted = await ok(env, 'extract_text', { document, view });
        ours = extracted['text'] as string;
        expect(ours, `${fixture} ${view} txt`).toBe(PlainText.of(await Texts.of(pkg, { view: view as 'accepted' })));
      }
      // The suite's goldens the engine meets exactly (the model's test, test/text-extraction.test.mts): the external and
      // structure probes'; extraction-probe.docx's hold the lines of the engine's phase 5 (LaTeX, charts, chunks).
      if (fixture === 'extraction-probe.docx') continue;
      const expected = readFileSync(join(phase0, row['golden']!), 'utf8');
      const md = row['format'] === 'md';
      expect(normalize(md ? imageDestinations(ours) : ours), row['golden']).toBe(normalize(md ? imageDestinations(expected) : expected));
      golden++;
    }
    expect(golden).toBe(8);
  }, 120_000);

  test('find: under accepted a span in the paragraph\'s text, none in a note or across a reference; under original and markup the deleted words, with no span', async () => {
    const env = await start({ 'lease.docx': 'tracked-changes.docx', 'notes.docx': 'footnote.docx' });
    const notes = (await ok(env, 'open', { path: 'notes.docx' }))['document'] as string;
    // "Here is a document with 2 footnotes\u0002:": the reference is no character in the text model's text
    const inText = await ok(env, 'find', { document: notes, text: '2 footnotes' });
    expect(inText['hits']).toEqual([expect.objectContaining({ address: 'body/0', offset: 24, length: 11, snippet: 'Here is a document with 2 footnotes:' })]);
    const across = await ok(env, 'find', { document: notes, text: 'footnotes:' });
    expect(across['hits']).toEqual([{ address: 'body/0', snippet: 'Here is a document with 2 footnotes:' }]);
    const inNote = await ok(env, 'find', { document: notes, text: 'Footnote 1' });
    expect(inNote['hits']).toEqual([{ address: 'footnote:3/0', story: 'footnote', snippet: ' Footnote 1' }]);
    expect((await ok(env, 'find', { document: notes, text: 'Footnote 1', notes: false }))['hits']).toEqual([]);

    const lease = await openLease(env);
    expect((await ok(env, 'find', { document: lease, text: 'A deletion' }))['total']).toBe(0);
    for (const view of ['original', 'markup']) {
      const found = await ok(env, 'find', { document: lease, text: 'A deletion', view });
      expect(found['hits'], view).toEqual([expect.objectContaining({ address: 'body/60' })]);
      expect(found['hits'] as Record<string, unknown>[]).toEqual([expect.not.objectContaining({ offset: expect.anything() })]);
      expect(found['spans'], view).toMatch(/accepted/);
    }
  });

  test('outline: the block model in the view; a note after the body; markup marked', async () => {
    const env = await start({ 'lease.docx': 'tracked-changes.docx', 'notes.docx': 'footnote.docx' });
    const notes = (await ok(env, 'open', { path: 'notes.docx' }))['document'] as string;
    const blocks = (await ok(env, 'outline', { document: notes }))['blocks'] as Record<string, unknown>[];
    expect(blocks[0]).toMatchObject({ address: 'body/0', kind: 'paragraph', text: 'Here is a document with 2 footnotes:' });
    expect(blocks).toContainEqual(expect.objectContaining({ address: 'footnote:3/0', story: 'footnote', text: ' Footnote 1' }));
    const lease = await openLease(env);
    const markup = (await ok(env, 'outline', { document: lease, view: 'markup' }))['blocks'] as Record<string, unknown>[];
    const tracked = markup.find((b) => b['address'] === 'body/60')!;
    expect(tracked['text']).toBe('Here is some change tracking. {++An insertion++} Followed by{-- A deletion--}.');
    expect(tracked['revisions']).toBeGreaterThan(0);
    const original = (await ok(env, 'outline', { document: lease, view: 'original' }))['blocks'] as Record<string, unknown>[];
    expect(original.find((b) => b['address'] === 'body/60')!['text']).toBe('Here is some change tracking.  Followed by A deletion.');
  });

  test('extract_text and read at an address, a note\'s among them; headers where included', async () => {
    const env = await start({ 'notes.docx': 'footnote.docx', 'headers.docx': 'headers.docx' });
    const notes = (await ok(env, 'open', { path: 'notes.docx' }))['document'] as string;
    expect((await ok(env, 'extract_text', { document: notes, address: 'body/0' }))['text']).toBe('Here is a document with 2 footnotes[1]:\n');
    expect((await ok(env, 'extract_text', { document: notes, address: 'footnote:3/0' }))['text']).toBe('[2] Footnote 1\n'); // the second note referenced
    expect((await ok(env, 'read', { document: notes, address: 'body/0' }))['markdown']).toBe('<!-- body/0 -->\nHere is a document with 2 footnotes[^1]:');
    const headers = (await ok(env, 'open', { path: 'headers.docx' }))['document'] as string;
    const without = (await ok(env, 'extract_text', { document: headers }))['text'] as string;
    const withHeaders = (await ok(env, 'extract_text', { document: headers, include: ['headers', 'footers'] }))['text'] as string;
    expect(withHeaders.length).toBeGreaterThan(without.length);
    const pkg = await WordprocessingMLPackage.load(await readFile(join(fixtures, 'headers.docx')));
    expect(withHeaders).toBe(PlainText.of(await Texts.of(pkg, { include: ['headers', 'footers'] })));
  });

  test('an option a tool does not take, or a value the engine does not implement, is arguments.invalid naming the statement (P5)', async () => {
    const env = await start({ 'lease.docx': 'tracked-changes.docx' });
    const document = await openLease(env);
    expect(await fails(env, 'outline', { document, tables: 'tsv' }, 'arguments.invalid')).toMatchObject({ argument: 'tables', statement: 'REQ-605' });
    expect(await fails(env, 'find', { document, text: 'x', metadata: true }, 'arguments.invalid')).toMatchObject({ argument: 'metadata', statement: 'REQ-515' });
    expect(await fails(env, 'read', { document, fields: 'update' }, 'arguments.invalid')).toMatchObject({ argument: 'fields', statement: 'REQ-506' });
    expect(await fails(env, 'read', { document, embedded: true }, 'arguments.invalid')).toMatchObject({ argument: 'embedded', statement: 'REQ-512' });
    expect(await fails(env, 'extract_text', { document, math: 'unicodemath' }, 'arguments.invalid')).toMatchObject({ argument: 'math', statement: 'REQ-604' });
    expect(await fails(env, 'extract_text', { document, repair: 'word' }, 'arguments.invalid')).toMatchObject({ argument: 'repair', statement: 'REQ-207' });
    expect(await fails(env, 'outline', { document, hiden: true }, 'arguments.invalid')).toMatchObject({ argument: 'hiden', statement: 'REQ-1001' });
    await fails(env, 'find', { document, xpath: '//w:p', view: 'original' }, 'arguments.invalid');
    // and what the engine implements is taken
    expect((await ok(env, 'extract_text', { document, tables: 'tsv', markup_detail: 'author', view: 'markup' }))['text']).toContain('{++Jason Harrop: An insertion++}');
  });

  test('if_text takes the paragraph\'s text as outline shows it (the text model\'s), or as Office JS spells it', async () => {
    const env = await start({ 'notes.docx': 'footnote.docx' });
    const document = (await ok(env, 'open', { path: 'notes.docx' }))['document'] as string;
    const shown = ((await ok(env, 'outline', { document }))['blocks'] as Block[])[0]!.text;
    expect(shown).toBe('Here is a document with 2 footnotes:');
    await ok(env, 'set_format', { document, address: 'body/0', anchor: 'document', bold: true, if_text: shown });
    await ok(env, 'set_format', { document, address: 'body/0', anchor: 'Here', italic: true, if_text: 'Here is a document with 2 footnotes\u0002:' });
    await fails(env, 'set_format', { document, address: 'body/0', anchor: 'Here', bold: true, if_text: 'Here is a document with 2 footnotes[1]:' }, 'target.stale');
  });
});
