// SPDX-License-Identifier: GPL-3.0-or-later
//
// OpenDoPE for an authoring tool (ED-003 section 4.2, 4.6 and 4.7; the OpenDoPE Specification v3
// WD 2026-09-18, cited by requirement number): the tag syntax (its section 3), the XPaths,
// Conditions, Components and Questions parts read into plain records and written back keeping
// every attribute they carry (REQ-012), the data part chosen by the specification's four rules
// (its section 2.2), and the evaluation a processor would do: an XPaths entry as a boolean in the
// declared conversion mode (table 7), a condition tree (table 6), a repeat's base and count
// (REQ-039). Pure over DOMs, strings and the engine's `XPathEngine`; nothing here touches the
// session, the store or the tree. The session-bound view is `template.mts`.
import type { CustomXmlPart, XPathEngine } from '@docx4j/core-ts';
import { parsePrefixMappings, booleanValue } from '@docx4j/core-ts';

// ---------------------------------------------------------------- namespaces (table 1)

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const XPATHS_NS = 'http://opendope.org/xpaths';
export const CONDITIONS_NS = 'http://opendope.org/conditions';
export const COMPONENTS_NS = 'http://opendope.org/components';
export const QUESTIONS_NS = 'http://opendope.org/questions';
export const ANSWERS_NS = 'http://opendope.org/answers';
export const OPENDOPE_PREFIX = 'http://opendope.org/';

/** The specification version this editor authors against, and writes to a new XPaths part (REQ-009). */
export const OPENDOPE_VERSION = '3.0';

/** The namespaces whose parts the data-part rules skip (section 2.2, docx4j's `CustomXmlDataStoragePartSelector`). */
export const COVER_PAGE_NS = 'http://schemas.microsoft.com/office/2006/coverPageProps';
export const DOCUMENT_MANAGEMENT_NS = 'http://schemas.microsoft.com/office/2006/metadata/properties';
export const CORE_PROPERTIES_NS = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
export const BIBLIOGRAPHY_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography';
const EXCLUDED_BINDING_NAMESPACES = [COVER_PAGE_NS, DOCUMENT_MANAGEMENT_NS, CORE_PROPERTIES_NS];
const WELL_KNOWN_PART_NAMESPACES = new Set([COVER_PAGE_NS, DOCUMENT_MANAGEMENT_NS, CORE_PROPERTIES_NS, BIBLIOGRAPHY_NS]);

/** True for a part whose root is in an `http://opendope.org/` namespace (the definition of an OpenDoPE part). */
export function isOpenDopePart(part: CustomXmlPart): boolean {
  return part.namespaceUri.startsWith(OPENDOPE_PREFIX);
}

/** The error the evaluator and the writers throw, naming the requirement where one applies. */
export class OpenDopeError extends Error {
  constructor(message: string, readonly requirement?: string) {
    super(requirement ? `${message} (${requirement})` : message);
    this.name = 'OpenDopeError';
  }
}

// ---------------------------------------------------------------- the tag syntax (section 3)

/** One parameter of a tag: the key percent-decoded, the value literal. */
export interface TagParam {
  key: string;
  value: string;
}

/**
 * `tag = param *( "&" param )`, `param = key [ "=" value ]`: keys percent-decoded with `+` a
 * space, values literal (they may contain `=`), a parameter without `=` a key whose value is the
 * key itself. Empty parameters (a trailing `&`) are dropped.
 */
export function parseTag(tag: string | null | undefined): TagParam[] {
  if (!tag) return [];
  const out: TagParam[] = [];
  for (const param of tag.split('&')) {
    if (param === '') continue;
    const eq = param.indexOf('=');
    const key = decodeKey(eq < 0 ? param : param.slice(0, eq));
    out.push({ key, value: eq < 0 ? key : param.slice(eq + 1) });
  }
  return out;
}

function decodeKey(raw: string): string {
  const plus = raw.replace(/\+/g, ' ');
  try {
    return decodeURIComponent(plus);
  } catch {
    return plus;
  }
}

/** The inverse of `parseTag`: keys percent-encoded where the grammar needs it, values as they are. */
export function formatTag(params: readonly TagParam[]): string {
  return params.map(({ key, value }) => {
    if (value.includes('&')) throw new OpenDopeError(`a tag value cannot contain "&" (parameter ${key})`);
    if (key === '') throw new OpenDopeError('a tag parameter needs a key');
    return `${encodeKey(key)}=${value}`;
  }).join('&');
}

function encodeKey(key: string): string {
  return key.replace(/[%&=+ ]/g, (c) => (c === ' ' ? '+' : `%${c.charCodeAt(0).toString(16).toUpperCase().padStart(2, '0')}`));
}

/** The value of a parameter, or undefined. */
export function tagParam(params: readonly TagParam[], key: string): string | undefined {
  return params.find((p) => p.key === key)?.value;
}

/** The parameters with `key` set to `value` (in place when present, appended otherwise) or removed when `value` is undefined. */
export function withTagParam(params: readonly TagParam[], key: string, value: string | undefined): TagParam[] {
  const at = params.findIndex((p) => p.key === key);
  if (value === undefined) return at < 0 ? [...params] : [...params.slice(0, at), ...params.slice(at + 1)];
  if (at < 0) return [...params, { key, value }];
  return params.map((p, i) => (i === at ? { key, value } : p));
}

/** The three parameters of which a control may carry at most one (REQ-007). */
export const ROLE_KEYS = ['od:condition', 'od:repeat', 'od:component'] as const;
/** The parameters only a processor writes (table 10, REQ-077). */
export const PROCESSOR_KEYS = ['od:rptd', 'od:RptOcc', 'od:resultConditionFalse', 'od:resultRepeatZero', 'w15:resultRepeatZero'] as const;

export type OpenDopeRole = 'none' | 'condition' | 'repeat' | 'component';

/** What the tag says, in the specification's vocabulary (table 10). */
export interface TagReading {
  params: TagParam[];
  /** The `od:xpath` entry id. */
  xpath?: string;
  role: OpenDopeRole;
  /** The condition, XPaths entry or component id the role refers to. */
  roleId?: string;
  /** `od:RptPosCon`, a position condition inside a repeat (section 8.3). */
  positionCondition?: string;
  contentType?: string;
  progid?: string;
  handler?: string;
  finish?: string;
  /** Parameters a processor set (an instance document opened for editing): shown read-only, never written by the editor. */
  processorSet: TagParam[];
  /** What the specification says is wrong with the tag: more than one role (REQ-007), processor-set parameters (REQ-077). */
  problems: string[];
}

export function readTag(tag: string | null | undefined): TagReading {
  const params = parseTag(tag);
  const roles = ROLE_KEYS.filter((k) => tagParam(params, k) !== undefined);
  const problems: string[] = [];
  if (roles.length > 1) problems.push(`a content control may carry only one of ${ROLE_KEYS.join(', ')}; this one has ${roles.join(' and ')} (REQ-007)`);
  const processorSet = params.filter((p) => (PROCESSOR_KEYS as readonly string[]).includes(p.key));
  if (processorSet.length > 0) problems.push(`${processorSet.map((p) => p.key).join(', ')}: set by a processor, not by a template (REQ-077)`);
  const role: OpenDopeRole = roles[0] === 'od:condition' ? 'condition' : roles[0] === 'od:repeat' ? 'repeat' : roles[0] === 'od:component' ? 'component' : 'none';
  const reading: TagReading = { params, role, processorSet, problems };
  if (roles[0]) reading.roleId = tagParam(params, roles[0]);
  const optional = (key: string, into: 'xpath' | 'positionCondition' | 'contentType' | 'progid' | 'handler' | 'finish'): void => {
    const v = tagParam(params, key);
    if (v !== undefined) reading[into] = v;
  };
  optional('od:xpath', 'xpath');
  optional('od:RptPosCon', 'positionCondition');
  optional('od:ContentType', 'contentType');
  optional('od:progid', 'progid');
  optional('od:Handler', 'handler');
  optional('od:finish', 'finish');
  return reading;
}

// ---------------------------------------------------------------- the parts as plain records

export type BooleanMode = 'java' | 'xpath1' | 'xpath2';

/** One `xpath` entry (table 3): the binding, and every other attribute of the entry under `attrs`. */
export interface XPathEntry {
  id: string;
  xpath: string;
  storeItemID: string;
  prefixMappings: string;
  /** `@name`, `@description`, `@questionID`, `@type`, ... and anything else the entry carries; kept as read (REQ-012). */
  attrs: Record<string, string>;
}

/** The XPaths part (table 2 on the root, table 3 per entry). */
export interface XPathsPart {
  version?: string;
  booleanConversion?: BooleanMode;
  authoringTool?: string;
  authoringToolVersion?: string;
  /** Every other root attribute, kept as read. */
  attrs: Record<string, string>;
  entries: XPathEntry[];
}

export type Operand =
  | { op: 'xpathref'; id: string }
  | { op: 'conditionref'; id: string }
  | { op: 'and' | 'or'; operands: Operand[] }
  | { op: 'not'; operand: Operand };

/** One `condition` (table 6): exactly one operand, and the documentation attributes under `attrs`. */
export interface Condition {
  id: string;
  attrs: Record<string, string>;
  operand: Operand;
}

export interface ConditionsPart {
  attrs: Record<string, string>;
  conditions: Condition[];
}

export interface Component {
  id: string;
  iri: string;
  attrs: Record<string, string>;
}

export interface ComponentsPart {
  attrs: Record<string, string>;
  components: Component[];
}

export interface Question {
  id: string;
  text: string;
  help?: string;
  hint?: string;
  appearance?: string;
  response: { kind: 'free' } | { kind: 'fixed'; canSelectMany: boolean; items: { label: string; value: string }[] };
}

export interface QuestionsPart {
  title?: string;
  subtitle?: string;
  questions: Question[];
}

/** The parts of a template, found by root namespace and local name (REQ-001), never by part name. */
export interface OpenDopeParts {
  xpaths?: CustomXmlPart;
  conditions?: CustomXmlPart;
  components?: CustomXmlPart;
  questions?: CustomXmlPart;
  answers?: CustomXmlPart;
}

export function findOpenDopeParts(parts: readonly CustomXmlPart[]): OpenDopeParts {
  const out: OpenDopeParts = {};
  for (const part of parts) {
    const root = part.document.documentElement;
    if (!root) continue;
    const local = root.localName ?? root.nodeName;
    const ns = part.namespaceUri;
    if (ns === XPATHS_NS && local === 'xpaths') out.xpaths ??= part;
    else if (ns === CONDITIONS_NS && local === 'conditions') out.conditions ??= part;
    else if (ns === COMPONENTS_NS && local === 'components') out.components ??= part;
    else if (ns === QUESTIONS_NS && local === 'questionnaire') out.questions ??= part;
    else if (ns === ANSWERS_NS && local === 'answers') out.answers ??= part;
  }
  return out;
}

function attributesOf(element: Element, except: readonly string[] = []): Record<string, string> {
  const out: Record<string, string> = {};
  const attrs = element.attributes;
  for (let i = 0; attrs && i < attrs.length; i++) {
    const a = attrs.item(i);
    if (!a || a.name === 'xmlns' || a.name.startsWith('xmlns:')) continue;
    const name = a.localName ?? a.name;
    if (except.includes(name)) continue;
    out[name] = a.value;
  }
  return out;
}

function childElementsOf(node: Node): Element[] {
  const out: Element[] = [];
  for (let c = node.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) out.push(c as Element);
  return out;
}

function localNameOf(element: Element): string {
  return element.localName ?? element.nodeName.replace(/^.*:/, '');
}

function textOfChild(element: Element, name: string): string | undefined {
  const child = childElementsOf(element).find((c) => localNameOf(c) === name);
  return child ? (child.textContent ?? '') : undefined;
}

export function readXPaths(part: CustomXmlPart): XPathsPart {
  const root = part.document.documentElement;
  if (!root) throw new OpenDopeError('the XPaths part has no document element');
  const attrs = attributesOf(root, ['version', 'booleanConversion', 'authoringTool', 'authoringToolVersion']);
  const out: XPathsPart = { attrs, entries: [] };
  const version = root.getAttribute('version');
  if (version) out.version = version;
  const mode = root.getAttribute('booleanConversion');
  if (mode === 'java' || mode === 'xpath1' || mode === 'xpath2') out.booleanConversion = mode;
  else if (mode) out.attrs['booleanConversion'] = mode;   // an unknown mode: kept, reported by `xpathsProblems`
  const tool = root.getAttribute('authoringTool');
  if (tool) out.authoringTool = tool;
  const toolVersion = root.getAttribute('authoringToolVersion');
  if (toolVersion) out.authoringToolVersion = toolVersion;
  for (const entry of childElementsOf(root)) {
    if (localNameOf(entry) !== 'xpath') continue;
    const binding = childElementsOf(entry).find((c) => localNameOf(c) === 'dataBinding');
    // The binding's attributes are unqualified (xpaths.xsd: attributeFormDefault unqualified; docx4j writes and
    // reads them so). A file whose writer qualified them with the WordprocessingML namespace (core-ts's check 37
    // and 38 templates, ED-005 section 12.87) is read all the same; the editor writes the schema's form.
    const attr = (name: string): string => binding?.getAttribute(name) ?? binding?.getAttributeNS(W_NS, name) ?? '';
    out.entries.push({
      id: entry.getAttribute('id') ?? '',
      xpath: attr('xpath'),
      storeItemID: attr('storeItemID'),
      prefixMappings: attr('prefixMappings'),
      attrs: attributesOf(entry, ['id']),
    });
  }
  return out;
}

/** The qualified name a child of `root` is written with: the root's prefix, so the part keeps its spelling. */
function qualified(root: Element, local: string): string {
  return root.prefix ? `${root.prefix}:${local}` : local;
}

/** Remove every element child named `local` and the whitespace around them, keeping anything else. */
function clearChildren(root: Element, local: string): void {
  for (let c = root.firstChild; c;) {
    const next = c.nextSibling;
    if ((c.nodeType === 1 && localNameOf(c as Element) === local) || (c.nodeType === 3 && (c.nodeValue ?? '').trim() === '')) root.removeChild(c);
    c = next;
  }
}

/**
 * Writes the record back into the part's DOM: the root's attributes set from the record (those it
 * carries; none removed, REQ-012), the entries rebuilt in order. Marks the part for re-marshalling.
 */
export function writeXPaths(part: CustomXmlPart, xpaths: XPathsPart): void {
  const document = part.document;
  const root = document.documentElement;
  if (!root) throw new OpenDopeError('the XPaths part has no document element');
  const rootAttrs: Record<string, string | undefined> = { ...xpaths.attrs, version: xpaths.version, booleanConversion: xpaths.booleanConversion, authoringTool: xpaths.authoringTool, authoringToolVersion: xpaths.authoringToolVersion };
  for (const [k, v] of Object.entries(rootAttrs)) if (v !== undefined) root.setAttribute(k, v);
  clearChildren(root, 'xpath');
  for (const entry of xpaths.entries) {
    const el = document.createElementNS(XPATHS_NS, qualified(root, 'xpath'));
    el.setAttribute('id', entry.id);
    for (const [k, v] of Object.entries(entry.attrs)) el.setAttribute(k, v);
    const binding = document.createElementNS(XPATHS_NS, qualified(root, 'dataBinding'));
    binding.setAttribute('xpath', entry.xpath);
    binding.setAttribute('storeItemID', entry.storeItemID);
    if (entry.prefixMappings !== '') binding.setAttribute('prefixMappings', entry.prefixMappings);
    el.appendChild(binding);
    root.appendChild(el);
  }
  part.touch();
}

/** The XML of a new XPaths part: version 3.0 in the `xpath2` mode (REQ-009, REQ-031, REC-004). */
export function newXPathsXml(mode: BooleanMode = 'xpath2'): string {
  return `<xpaths xmlns="${XPATHS_NS}" version="${OPENDOPE_VERSION}" booleanConversion="${mode}"/>`;
}

export function newConditionsXml(): string {
  return `<conditions xmlns="${CONDITIONS_NS}"/>`;
}

function readOperand(element: Element): Operand {
  const name = localNameOf(element);
  const children = childElementsOf(element);
  switch (name) {
    case 'xpathref': return { op: 'xpathref', id: element.getAttribute('id') ?? '' };
    case 'conditionref': return { op: 'conditionref', id: element.getAttribute('id') ?? '' };
    case 'and':
    case 'or': return { op: name, operands: children.map(readOperand) };
    case 'not': {
      if (children.length !== 1) throw new OpenDopeError(`"not" takes exactly one operand; this one has ${children.length}`);
      return { op: 'not', operand: readOperand(children[0]!) };
    }
    default: throw new OpenDopeError(`unknown condition operator "${name}" (table 6 allows xpathref, conditionref, and, or, not)`);
  }
}

export function readConditions(part: CustomXmlPart): ConditionsPart {
  const root = part.document.documentElement;
  if (!root) throw new OpenDopeError('the Conditions part has no document element');
  const out: ConditionsPart = { attrs: attributesOf(root), conditions: [] };
  for (const el of childElementsOf(root)) {
    if (localNameOf(el) !== 'condition') continue;
    const children = childElementsOf(el);
    const id = el.getAttribute('id') ?? '';
    if (children.length !== 1) throw new OpenDopeError(`condition ${id} must hold exactly one of xpathref, and, or, not; it holds ${children.length}`, 'table 6');
    out.conditions.push({ id, attrs: attributesOf(el, ['id']), operand: readOperand(children[0]!) });
  }
  return out;
}

function writeOperand(document: Document, root: Element, operand: Operand): Element {
  const el = document.createElementNS(CONDITIONS_NS, qualified(root, operand.op));
  switch (operand.op) {
    case 'xpathref':
    case 'conditionref':
      el.setAttribute('id', operand.id);
      break;
    case 'and':
    case 'or':
      for (const o of operand.operands) el.appendChild(writeOperand(document, root, o));
      break;
    case 'not':
      el.appendChild(writeOperand(document, root, operand.operand));
      break;
  }
  return el;
}

export function writeConditions(part: CustomXmlPart, conditions: ConditionsPart): void {
  const document = part.document;
  const root = document.documentElement;
  if (!root) throw new OpenDopeError('the Conditions part has no document element');
  for (const [k, v] of Object.entries(conditions.attrs)) root.setAttribute(k, v);
  clearChildren(root, 'condition');
  for (const c of conditions.conditions) {
    if (c.operand.op === 'conditionref') throw new OpenDopeError(`condition ${c.id}: a condition cannot hold a conditionref directly; wrap it in and, or or not`, 'table 6');
    const el = document.createElementNS(CONDITIONS_NS, qualified(root, 'condition'));
    el.setAttribute('id', c.id);
    for (const [k, v] of Object.entries(c.attrs)) el.setAttribute(k, v);
    el.appendChild(writeOperand(document, root, c.operand));
    root.appendChild(el);
  }
  part.touch();
}

export function readComponents(part: CustomXmlPart): ComponentsPart {
  const root = part.document.documentElement;
  if (!root) throw new OpenDopeError('the Components part has no document element');
  const out: ComponentsPart = { attrs: attributesOf(root), components: [] };
  for (const el of childElementsOf(root)) {
    if (localNameOf(el) !== 'component') continue;
    out.components.push({ id: el.getAttribute('id') ?? '', iri: el.getAttribute('iri') ?? '', attrs: attributesOf(el, ['id', 'iri']) });
  }
  return out;
}

/** The Questions part, read for display (docx4j's `org.opendope.questions`; the specification describes interactive assembly in outline only). */
export function readQuestions(part: CustomXmlPart): QuestionsPart {
  const root = part.document.documentElement;
  if (!root) throw new OpenDopeError('the Questions part has no document element');
  const out: QuestionsPart = { questions: [] };
  const title = textOfChild(root, 'title');
  if (title !== undefined) out.title = title;
  const subtitle = textOfChild(root, 'subtitle');
  if (subtitle !== undefined) out.subtitle = subtitle;
  const questions = childElementsOf(root).find((c) => localNameOf(c) === 'questions');
  for (const q of questions ? childElementsOf(questions) : []) {
    if (localNameOf(q) !== 'question') continue;
    const question: Question = { id: q.getAttribute('id') ?? '', text: textOfChild(q, 'text') ?? '', response: { kind: 'free' } };
    const help = textOfChild(q, 'help');
    if (help !== undefined) question.help = help;
    const hint = textOfChild(q, 'hint');
    if (hint !== undefined) question.hint = hint;
    const appearance = q.getAttribute('appearance');
    if (appearance) question.appearance = appearance;
    const response = childElementsOf(q).find((c) => localNameOf(c) === 'response');
    const fixed = response ? childElementsOf(response).find((c) => localNameOf(c) === 'fixed') : undefined;
    if (fixed) {
      question.response = {
        kind: 'fixed',
        canSelectMany: fixed.getAttribute('canSelectMany') === 'true',
        items: childElementsOf(fixed).filter((c) => localNameOf(c) === 'item').map((item) => ({ label: textOfChild(item, 'label') ?? '', value: textOfChild(item, 'value') ?? '' })),
      };
    }
    out.questions.push(question);
  }
  return out;
}

// ---------------------------------------------------------------- what is wrong with the parts

export interface PartProblem {
  part: 'xpaths' | 'conditions';
  id?: string;
  message: string;
  requirement: string;
}

/** Duplicate ids (REQ-013), the declared version and mode (REQ-009, REQ-010, REQ-031), empty expressions. */
export function xpathsProblems(xpaths: XPathsPart): PartProblem[] {
  const out: PartProblem[] = [];
  const seen = new Set<string>();
  for (const e of xpaths.entries) {
    if (e.id === '') out.push({ part: 'xpaths', message: 'an entry has no id', requirement: 'REQ-013' });
    else if (seen.has(e.id)) out.push({ part: 'xpaths', id: e.id, message: `entry id ${e.id} is used twice`, requirement: 'REQ-013' });
    seen.add(e.id);
    if (e.xpath === '') out.push({ part: 'xpaths', id: e.id, message: `entry ${e.id} has no expression`, requirement: 'REQ-014' });
    if (e.storeItemID === '') out.push({ part: 'xpaths', id: e.id, message: `entry ${e.id} names no store item id`, requirement: 'table 3' });
  }
  if (xpaths.version === undefined) out.push({ part: 'xpaths', message: 'no version declared: authored against 2.3 or earlier; the mode is the processor\'s default (java)', requirement: 'REQ-009' });
  else if (versionNewerThan(xpaths.version, OPENDOPE_VERSION)) out.push({ part: 'xpaths', message: `declares version ${xpaths.version}, newer than ${OPENDOPE_VERSION}: the OpenDoPE parts are read-only here`, requirement: 'REQ-010' });
  if (xpaths.version !== undefined && xpaths.booleanConversion === undefined) {
    const declared = xpaths.attrs['booleanConversion'];
    out.push({ part: 'xpaths', message: declared ? `booleanConversion "${declared}" is not java, xpath1 or xpath2` : 'no booleanConversion declared', requirement: 'REQ-031' });
  }
  return out;
}

/** `3.1` is newer than `3.0`; a version that does not parse counts as newer (refused, REQ-010's spirit). */
export function versionNewerThan(version: string, than: string): boolean {
  const parse = (v: string): number[] | undefined => {
    const m = /^\s*(\d+)(?:\.(\d+))?/.exec(v);
    return m ? [Number(m[1]), Number(m[2] ?? '0')] : undefined;
  };
  const a = parse(version);
  const b = parse(than)!;
  if (!a) return true;
  return a[0]! > b[0]! || (a[0] === b[0] && a[1]! > b[1]!);
}

/** Duplicate ids (REQ-028), references to missing conditions or entries, cycles (REQ-029). */
export function conditionsProblems(conditions: ConditionsPart, xpaths: XPathsPart | undefined): PartProblem[] {
  const out: PartProblem[] = [];
  const byId = new Map<string, Condition>();
  for (const c of conditions.conditions) {
    if (c.id === '') out.push({ part: 'conditions', message: 'a condition has no id', requirement: 'REQ-028' });
    else if (byId.has(c.id)) out.push({ part: 'conditions', id: c.id, message: `condition id ${c.id} is used twice`, requirement: 'REQ-028' });
    byId.set(c.id, c);
  }
  const entries = new Set(xpaths?.entries.map((e) => e.id) ?? []);
  const refs = (o: Operand, into: { conditions: string[]; xpaths: string[] }): void => {
    if (o.op === 'conditionref') into.conditions.push(o.id);
    else if (o.op === 'xpathref') into.xpaths.push(o.id);
    else if (o.op === 'not') refs(o.operand, into);
    else for (const x of o.operands) refs(x, into);
  };
  for (const c of conditions.conditions) {
    const r = { conditions: [] as string[], xpaths: [] as string[] };
    refs(c.operand, r);
    for (const id of r.conditions) if (!byId.has(id)) out.push({ part: 'conditions', id: c.id, message: `condition ${c.id} refers to condition ${id}, which does not exist`, requirement: 'REQ-029' });
    for (const id of r.xpaths) if (!entries.has(id)) out.push({ part: 'conditions', id: c.id, message: `condition ${c.id} refers to XPaths entry ${id}, which does not exist`, requirement: 'REQ-029' });
    if ((c.operand.op === 'and' || c.operand.op === 'or') && c.operand.operands.length === 0) out.push({ part: 'conditions', id: c.id, message: `condition ${c.id}: ${c.operand.op} needs at least one operand`, requirement: 'table 6' });
  }
  // Cycles: a depth-first walk over conditionref edges.
  const state = new Map<string, 'visiting' | 'done'>();
  const visit = (id: string, path: string[]): void => {
    const s = state.get(id);
    if (s === 'done') return;
    if (s === 'visiting') { out.push({ part: 'conditions', id, message: `conditions refer to each other in a cycle: ${[...path, id].join(' -> ')}`, requirement: 'REQ-029' }); return; }
    const c = byId.get(id);
    if (!c) return;
    state.set(id, 'visiting');
    const r = { conditions: [] as string[], xpaths: [] as string[] };
    refs(c.operand, r);
    for (const next of r.conditions) visit(next, [...path, id]);
    state.set(id, 'done');
  };
  for (const c of conditions.conditions) visit(c.id, []);
  return out;
}

// ---------------------------------------------------------------- the data part (section 2.2)

export interface DataPartChoice {
  part?: CustomXmlPart;
  /** Which of the four rules decided. */
  rule: 1 | 2 | 3 | 4;
  reason: string;
}

/** A `w:dataBinding` seen in the main document body, in document order (rule 2's input). */
export interface BindingSeen {
  storeItemID: string;
  prefixMappings: string;
}

/** Store item ids compare case-insensitively, with or without braces (REQ-005). */
export function sameStoreItemId(a: string, b: string): boolean {
  return normaliseId(a) === normaliseId(b);
}

function normaliseId(id: string): string {
  return id.replace(/[{}]/g, '').toLowerCase();
}

export function partByStoreItemId(parts: readonly CustomXmlPart[], id: string): CustomXmlPart | undefined {
  if (id === '') return undefined;
  return parts.find((p) => sameStoreItemId(p.id, id));
}

function excludedMappings(prefixMappings: string): boolean {
  return EXCLUDED_BINDING_NAMESPACES.some((ns) => prefixMappings.includes(ns));
}

/**
 * The specification's four rules, in order: the first XPaths entry whose store item id resolves
 * (skipping entries bound to the property stores); else, with no XPaths part, the first
 * `w:dataBinding` of the body likewise; else the one part that is neither OpenDoPE's nor a
 * well-known Microsoft part (REQ-003: never a choice between several); else undetermined (REQ-004).
 */
export function selectDataPart(parts: readonly CustomXmlPart[], xpaths: XPathsPart | undefined, bindings: readonly BindingSeen[]): DataPartChoice {
  if (xpaths && xpaths.entries.length > 0) {
    for (const entry of xpaths.entries) {
      if (excludedMappings(entry.prefixMappings)) continue;
      const part = partByStoreItemId(parts, entry.storeItemID);
      if (part && !isOpenDopePart(part)) return { part, rule: 1, reason: `named by XPaths entry ${entry.id}` };
    }
    return { rule: 4, reason: 'the XPaths part has entries, but none names a custom XML part of this package' };
  }
  if (!xpaths) {
    for (const b of bindings) {
      if (excludedMappings(b.prefixMappings)) continue;
      const part = partByStoreItemId(parts, b.storeItemID);
      if (part && !isOpenDopePart(part)) return { part, rule: 2, reason: 'named by the first w:dataBinding in the body' };
    }
  }
  const candidates = parts.filter((p) => !isOpenDopePart(p) && !WELL_KNOWN_PART_NAMESPACES.has(p.namespaceUri));
  if (candidates.length === 1) return { part: candidates[0]!, rule: 3, reason: 'the only custom XML part that is neither OpenDoPE\'s nor a well-known Microsoft part' };
  if (candidates.length === 0) return { rule: 4, reason: 'no custom XML part could hold data' };
  return { rule: 4, reason: `${candidates.length} custom XML parts could hold data and nothing names one (REQ-003)` };
}

// ---------------------------------------------------------------- evaluation (sections 7 and 8)

/** What the evaluator needs: the parts (for store item ids), the engine, the mode and the two definition parts. */
export interface EvaluationContext {
  parts: readonly CustomXmlPart[];
  engine: XPathEngine;
  mode: BooleanMode;
  xpaths: XPathsPart | undefined;
  conditions: ConditionsPart | undefined;
}

/** The XPaths entry an id names, or an error naming it. */
export function entryNamed(ctx: EvaluationContext, id: string): XPathEntry {
  const entry = ctx.xpaths?.entries.find((e) => e.id === id);
  if (!entry) throw new OpenDopeError(`XPaths entry ${id} does not exist`, ctx.xpaths ? 'REQ-017' : 'table 1');
  return entry;
}

function contextOf(ctx: EvaluationContext, entry: XPathEntry): { node: Node; namespaces: Record<string, string> } {
  const part = partByStoreItemId(ctx.parts, entry.storeItemID);
  if (!part) throw new OpenDopeError(`XPaths entry ${entry.id} names store item ${entry.storeItemID || '(none)'}, which is not a custom XML part of this package`);
  return { node: part.document, namespaces: parsePrefixMappings(entry.prefixMappings) };
}

/**
 * Table 7 through the engine's `booleanValue` (core-ts CR-005 phase A, 0.1.3): `java` and
 * `xpath1` over any engine; `xpath2` as `xs:boolean(expression)` through the FontoXPath engine,
 * docx4j's own strategy, and otherwise the error REQ-032 asks for (evaluate in the declared mode
 * or refuse), which core-ts raises naming the engine to install. One implementation of the
 * table across the two packages, so the preview and the bind goldens cannot drift.
 */
export function booleanIn(engine: XPathEngine, expression: string, context: Node, namespaces: Record<string, string>, mode: BooleanMode): boolean {
  try {
    return booleanValue(engine, expression, context, namespaces, mode);
  } catch (e) {
    if (mode === 'xpath2' && e instanceof Error && /XPath 2\.0 engine|fontoxpath/i.test(e.message)) throw new OpenDopeError(e.message, 'REQ-032');
    throw e;
  }
}

/** An XPaths entry as a boolean in the context's mode. */
export function evaluateEntryBoolean(ctx: EvaluationContext, id: string): boolean {
  const entry = entryNamed(ctx, id);
  const { node, namespaces } = contextOf(ctx, entry);
  return booleanIn(ctx.engine, entry.xpath, node, namespaces, ctx.mode);
}

/** The XPath string value of an entry, trimmed (REQ-018, REQ-019); undefined when it selects nothing. */
export function evaluateEntryValue(ctx: EvaluationContext, id: string): string | undefined {
  const entry = entryNamed(ctx, id);
  const { node, namespaces } = contextOf(ctx, entry);
  const value = ctx.engine.selectValue(entry.xpath, node, namespaces);
  if (value === undefined) return undefined;
  let text = String(value).trim();
  if (/^\s*local-name/.test(entry.xpath)) text = text.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, hex: string) => String.fromCharCode(parseInt(hex, 16)));   // REQ-020
  return text;
}

/** The value a `w:dataBinding` selects now, trimmed, or undefined when it selects nothing (the bound-value verdict). */
export function evaluateBinding(ctx: EvaluationContext, binding: { xpath: string; storeItemID: string; prefixMappings: string }): string | undefined {
  const part = partByStoreItemId(ctx.parts, binding.storeItemID);
  if (!part) throw new OpenDopeError(`the binding names store item ${binding.storeItemID || '(none)'}, which is not a custom XML part of this package`);
  const nodes = ctx.engine.select(binding.xpath, part.document, parsePrefixMappings(binding.prefixMappings));
  const first = nodes[0];
  return first ? (first.textContent ?? '').trim() : undefined;
}

/** A condition by id, with the short-circuit rules of table 6 and the cycle and missing-reference errors of REQ-029 and REQ-030. */
export function evaluateCondition(ctx: EvaluationContext, id: string, visiting: readonly string[] = []): boolean {
  if (visiting.includes(id)) throw new OpenDopeError(`conditions refer to each other in a cycle: ${[...visiting, id].join(' -> ')}`, 'REQ-029');
  const condition = ctx.conditions?.conditions.find((c) => c.id === id);
  if (!condition) throw new OpenDopeError(`condition ${id} does not exist`, visiting.length > 0 ? 'REQ-029' : 'REQ-030');
  return evaluateOperand(ctx, condition.operand, [...visiting, id]);
}

export function evaluateOperand(ctx: EvaluationContext, operand: Operand, visiting: readonly string[]): boolean {
  switch (operand.op) {
    case 'xpathref': return evaluateEntryBoolean(ctx, operand.id);
    case 'conditionref': return evaluateCondition(ctx, operand.id, visiting);
    case 'not': return !evaluateOperand(ctx, operand.operand, visiting);
    case 'and': {
      for (const o of operand.operands) if (!evaluateOperand(ctx, o, visiting)) return false;
      return true;
    }
    case 'or': {
      for (const o of operand.operands) if (evaluateOperand(ctx, o, visiting)) return true;
      return false;
    }
  }
}

/** The repeat base (section 8.1): the expression less one trailing `/` or one trailing `[1]`. */
export function repeatBase(xpath: string): string {
  const trimmed = xpath.trim();
  if (trimmed.endsWith('/')) return trimmed.slice(0, -1);
  if (trimmed.endsWith('[1]')) return trimmed.slice(0, -3);
  return trimmed;
}

/** The nodes the repeat base selects: REQ-039's instance count is their number. */
export function repeatNodes(ctx: EvaluationContext, id: string): Node[] {
  const entry = entryNamed(ctx, id);
  const { node, namespaces } = contextOf(ctx, entry);
  return ctx.engine.select(repeatBase(entry.xpath), node, namespaces);
}

export function repeatCount(ctx: EvaluationContext, id: string): number {
  return repeatNodes(ctx, id).length;
}

/**
 * REQ-038's warning: a repeat must select the elements to repeat, not their container. When the
 * base selects exactly one element whose element children all share one name and there are at
 * least two of them, the author most likely pointed at the container.
 */
export function repeatLooksLikeContainer(ctx: EvaluationContext, id: string): boolean {
  const nodes = repeatNodes(ctx, id);
  if (nodes.length !== 1 || nodes[0]!.nodeType !== 1) return false;
  const children = childElementsOf(nodes[0]!);
  if (children.length < 2) return false;
  const name = localNameOf(children[0]!);
  return children.every((c) => localNameOf(c) === name);
}

/** A fresh id with `prefix` not among `used` (`x1`, `x2`, ... or `c1`, ...). */
export function freshId(prefix: string, used: Iterable<string>): string {
  const taken = new Set(used);
  for (let n = 1; ; n++) {
    const id = `${prefix}${n}`;
    if (!taken.has(id)) return id;
  }
}

/** An XPath's location steps, split on the `/` outside predicates and quotes: `/invoice[1]/items/item[1]` is `['invoice[1]', 'items', 'item[1]']`. */
export function xpathSteps(xpath: string): string[] {
  const steps: string[] = [];
  let depth = 0;
  let quote: string | null = null;
  let current = '';
  for (const ch of xpath) {
    if (quote) { current += ch; if (ch === quote) quote = null; continue; }
    if (ch === '"' || ch === "'") { quote = ch; current += ch; continue; }
    if (ch === '[') depth++;
    if (ch === ']') depth--;
    if (ch === '/' && depth === 0) { if (current !== '') steps.push(current); current = ''; continue; }
    current += ch;
  }
  if (current !== '') steps.push(current);
  return steps;
}

const stepName = (step: string): string => step.replace(/\[.*$/s, '');
const stepPredicates = (step: string): string => step.slice(stepName(step).length);

/**
 * A binding's XPath written relative to the repeat it sits in (the specification's contextualising,
 * REQ-045 to REQ-049): docx4j's `XPathEnhancerParser` rewrites a binding for each row by its textual
 * prefix, so a binding inside a repeat must begin with the repeat's XPath verbatim, then the instance
 * (`[1]` in the template) and the rest of its steps: for the repeat `/invoice[1]/items/item`, the
 * Data tab's `/invoice[1]/items[1]/item[1]/name[1]` becomes `/invoice[1]/items/item[1]/name[1]`
 * (docx4j's own invoice sample writes `/invoice[1]/items/item[1]/name`). Undefined when the XPath
 * does not lie under the repeat's, step name by step name, or the repeat's last step has a predicate.
 */
export function relativeToRepeat(xpath: string, repeatXPath: string): string | undefined {
  const base = xpathSteps(repeatXPath);
  const steps = xpathSteps(xpath);
  if (base.length === 0 || steps.length <= base.length) return undefined;
  if (!xpath.startsWith('/') || !repeatXPath.startsWith('/')) return undefined;
  for (let i = 0; i < base.length; i++) if (stepName(steps[i]!) !== stepName(base[i]!)) return undefined;
  if (stepPredicates(base[base.length - 1]!) !== '') return undefined;
  const instance = stepPredicates(steps[base.length - 1]!) || '[1]';
  return `${repeatXPath}${instance}/${steps.slice(base.length).join('/')}`;
}
