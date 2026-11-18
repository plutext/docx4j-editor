// SPDX-License-Identifier: GPL-3.0-or-later
//
// The template view of a session (ED-003 section 4): the custom XML parts through the engine's
// `pkg.customXmlParts`, the OpenDoPE parts as records (opendope.mts), the data part by the
// specification's rules, the parts edited in place under E1's `markPartChanged` protocol
// (section 4.2: no undo for a part edit, said so in the panel), and the verdicts of section 4.7
// for one control's properties. Reading never marks a part; `customXmlParts.load()` parses the
// DOMs and leaves the bytes alone, so the untouched round trip holds after every panel was read.
import type { CustomXmlPart } from '@docx4j/core-ts';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorDocument } from './session.mjs';
import { APPLICATION_NAME } from './appxml.mjs';
import { readControlEntries } from './yjs.mjs';
import { sdtBinding, sdtKind, sdtTag } from './sdtpr.mjs';
import {
  findOpenDopeParts, readXPaths, writeXPaths, readConditions, writeConditions, readComponents, readQuestions,
  newXPathsXml, newConditionsXml, selectDataPart, isOpenDopePart, xpathsProblems, conditionsProblems, versionNewerThan, freshId, readTag,
  evaluateCondition, evaluateEntryValue, evaluateBinding, repeatCount, repeatLooksLikeContainer, repeatBase, entryNamed,
  OPENDOPE_VERSION, OpenDopeError,
  type OpenDopeParts, type XPathsPart, type XPathEntry, type ConditionsPart, type Condition, type Operand, type ComponentsPart, type QuestionsPart,
  type BooleanMode, type DataPartChoice, type BindingSeen, type PartProblem, type EvaluationContext, type OpenDopeRole,
} from './opendope.mjs';

/** What the preview says for one control (ED-003 section 4.7). */
export interface Verdict {
  role: OpenDopeRole | 'positionCondition' | 'repeatingSection';
  /** The condition, entry or component id the role names. */
  id?: string;
  /** A condition: kept (true) or dropped (REQ-033, REQ-034). */
  kept?: boolean;
  /** A repeat or a Word repeating section: the instance count (REQ-039, REQ-052); zero means removed (REQ-040). */
  count?: number;
  /** The bound value, trimmed (REQ-017 to REQ-019); '' when empty (a placeholder, REQ-024). */
  value?: string;
  /** The expression a position condition would apply (section 8.3); shown, not evaluated. */
  positionPredicate?: string;
  /** The binding or entry selects nothing in the data part. */
  selectsNothing?: boolean;
  error?: string;
  warnings: string[];
}

/** One use of an XPaths entry: what refers to it, and as what. */
export interface EntryUse {
  kind: 'bound' | 'repeat' | 'position' | 'condition';
  label: string;
}

export class TemplateModel {
  private ready = false;
  private cache: { parts?: OpenDopeParts; xpaths?: XPathsPart | null; conditions?: ConditionsPart | null; components?: ComponentsPart | null; questions?: QuestionsPart | null } = {};
  /** True once an OpenDoPE part was edited: the save stamps the authoring tool (REQ-011). */
  touched = false;

  constructor(private readonly session: EditorDocument) {}

  /**
   * Parse the custom XML parts and warm the XPath engine; once, before anything else here. The
   * engine becomes core-ts's FontoXPath one (XPath 3.1, the `xpath2` mode as `xs:boolean(...)`,
   * section 9 item 13), loaded here rather than with the package so that a session that never
   * opens the Template perspective never pays for it; when `fontoxpath` cannot be loaded the
   * default engine stays and the `xpath2` mode is refused with REQ-032's message. With
   * `engine: false` only the parts are parsed (the application asks `isTemplate` of every
   * document it opens, section 4.9's suggestion, and must not pay for the engine to do so); a
   * later `load()` installs the engine.
   *
   * `fontoxpath` is the loaded module for a browser: core-ts's engine loads the package by a
   * bare `import()` a bundler leaves alone, so the application imports it into its Template
   * chunk and hands it over (`new FontoXPathEngine(fontoxpath)`, core-ts 0.1.4, CR-005 section
   * 7); without it, Node's `import('fontoxpath')` is what happens.
   */
  async load(options: { engine?: boolean; fontoxpath?: unknown } = {}): Promise<void> {
    const pkg = this.session.pkg;
    const before = pkg.xpathEngine;
    if (!this.engineInstalled && options.engine !== false) {
      try {
        const { FontoXPathEngine } = await import('@docx4j/core-ts/xpath-fonto');
        pkg.xpathEngine = options.fontoxpath !== undefined ? new FontoXPathEngine(options.fontoxpath) : new FontoXPathEngine();
        await pkg.xpathEngine.ready();
        this.engineInstalled = true;
      } catch {
        pkg.xpathEngine = before;
      }
    }
    await pkg.customXmlParts.load();
    this.ready = true;
    this.invalidate();
  }
  private engineInstalled = false;

  get loaded(): boolean {
    return this.ready;
  }

  /** Whether the FontoXPath engine is installed (the `xpath2` mode evaluates). */
  get engineLoaded(): boolean {
    return this.engineInstalled;
  }

  /** The custom XML part a store item id names, or undefined (REQ-005's comparison). */
  partNamed(storeItemID: string): CustomXmlPart | undefined {
    return this.parts.find((p) => sameId(p.id, storeItemID));
  }

  /** Drop the cached reads (after a tree run, or an edit made outside these methods), and tell the listeners. */
  invalidate(): void {
    this.cache = {};
    this.notify();
  }

  private readonly listeners = new Set<() => void>();

  /** Hear of every change to what the template view answers: the load, a part write, an invalidation (the view refreshes its verdict badges). */
  subscribe(fn: () => void): () => void {
    this.listeners.add(fn);
    return () => { this.listeners.delete(fn); };
  }

  private notify(): void {
    for (const fn of this.listeners) fn();
  }

  /** The custom XML parts, in the package's order. */
  get parts(): CustomXmlPart[] {
    if (!this.ready) throw new OpenDopeError('the template is not loaded: await session.template.load() once');
    return this.session.pkg.customXmlParts.items;
  }

  /** The OpenDoPE parts, found by root namespace and local name (REQ-001). */
  get openDopeParts(): OpenDopeParts {
    return (this.cache.parts ??= findOpenDopeParts(this.parts));
  }

  get xpaths(): XPathsPart | undefined {
    if (this.cache.xpaths === undefined) this.cache.xpaths = this.openDopeParts.xpaths ? readXPaths(this.openDopeParts.xpaths) : null;
    return this.cache.xpaths ?? undefined;
  }

  get conditions(): ConditionsPart | undefined {
    if (this.cache.conditions === undefined) this.cache.conditions = this.openDopeParts.conditions ? readConditions(this.openDopeParts.conditions) : null;
    return this.cache.conditions ?? undefined;
  }

  get components(): ComponentsPart | undefined {
    if (this.cache.components === undefined) this.cache.components = this.openDopeParts.components ? readComponents(this.openDopeParts.components) : null;
    return this.cache.components ?? undefined;
  }

  get questions(): QuestionsPart | undefined {
    if (this.cache.questions === undefined) this.cache.questions = this.openDopeParts.questions ? readQuestions(this.openDopeParts.questions) : null;
    return this.cache.questions ?? undefined;
  }

  /** The boolean conversion mode: the declared one, else `java`, which is what docx4j does for a template that declares none. */
  get mode(): BooleanMode {
    return this.xpaths?.booleanConversion ?? 'java';
  }

  /** Whether the OpenDoPE parts may be edited here: not when the template declares a newer version (REQ-010's spirit). */
  get editable(): { ok: boolean; reason?: string } {
    const version = this.xpaths?.version;
    if (version !== undefined && versionNewerThan(version, OPENDOPE_VERSION)) return { ok: false, reason: `the template declares OpenDoPE version ${version}, newer than the ${OPENDOPE_VERSION} this editor implements` };
    return { ok: true };
  }

  /** What is wrong with the parts (REQ-009, REQ-010, REQ-013, REQ-028 to REQ-031). */
  get problems(): PartProblem[] {
    const out: PartProblem[] = [];
    if (this.xpaths) out.push(...xpathsProblems(this.xpaths));
    if (this.conditions) out.push(...conditionsProblems(this.conditions, this.xpaths));
    return out;
  }

  /** Whether the document suggests the Template perspective (section 4.9): OpenDoPE parts or any binding. */
  get isTemplate(): boolean {
    const p = this.openDopeParts;
    return p.xpaths !== undefined || p.conditions !== undefined || p.components !== undefined || this.bindingsInBody().length > 0;
  }

  /** Every `w:dataBinding` of the body in document order (rule 2's input), from the nodes and the group entries. */
  bindingsInBody(doc: PMNode = this.session.doc): BindingSeen[] {
    const out: BindingSeen[] = [];
    const push = (sdtPr: unknown): void => {
      const b = sdtBinding(sdtPr);
      if (b) out.push({ storeItemID: b.storeItemID, prefixMappings: b.prefixMappings });
    };
    const groups = readControlEntries(this.session.store);
    doc.descendants((node) => {
      if (node.type.name === 'sdtBlock' || node.type.name === 'sdtInline') push(node.attrs['sdtPr']);
      if (node.type.name === 'table_row' || node.type.name === 'table_cell') {
        for (const g of (node.attrs['sdt'] as { groupId: string; first: boolean }[] | null) ?? []) if (g.first) push(groups.get(g.groupId)?.sdtPr ?? null);
      }
      return true;
    });
    return out;
  }

  /** The data part by the specification's four rules (section 2.2). */
  get dataPart(): DataPartChoice {
    return selectDataPart(this.parts, this.xpaths, this.bindingsInBody());
  }

  // ---------------------------------------------------------------- the parts edited in place

  private requireEditable(): void {
    const e = this.editable;
    if (!e.ok) throw new OpenDopeError(e.reason!, 'REQ-010');
  }

  private markTouched(part: CustomXmlPart): void {
    this.touched = true;
    this.session.markPartChanged(part.part.partName.toString());
  }

  /** The XPaths part, created on first use as a version 3.0 part in the `xpath2` mode (REC-004) and related from the main part (REQ-002). */
  ensureXPathsPart(): CustomXmlPart {
    const existing = this.openDopeParts.xpaths;
    if (existing) return existing;
    this.requireEditable();
    const part = this.session.addCustomXmlPart(newXPathsXml());
    this.cache = {};
    this.markTouched(part);
    return part;
  }

  ensureConditionsPart(): CustomXmlPart {
    const existing = this.openDopeParts.conditions;
    if (existing) return existing;
    this.requireEditable();
    const part = this.session.addCustomXmlPart(newConditionsXml());
    this.cache = {};
    this.markTouched(part);
    return part;
  }

  /**
   * A custom XML part to hold data, added to the package from its XML (section 11.2 item 31):
   * a document with no data part gets one, which rule 3 then names when it is the only
   * candidate. The XML must parse to a document with one root element; it is a part edit
   * (marked for the save, no undo). The OpenDoPE parts are not touched: nothing names the new
   * part until an entry or a binding does.
   */
  addDataPart(xml: string): CustomXmlPart {
    if (!this.ready) throw new OpenDopeError('the template is not loaded: await session.template.load() once');
    let part: CustomXmlPart;
    try {
      part = this.session.addCustomXmlPart(xml);   // the engine parses, and throws on malformed XML
    } catch (e) {
      throw new OpenDopeError(`not a custom XML part: ${(e as Error).message}`);
    }
    if (isOpenDopePart(part)) {
      this.session.pkg.customXmlParts.removePart(part);
      throw new OpenDopeError('an OpenDoPE part is not a data part; the OpenDoPE tab makes those');
    }
    this.cache = {};
    this.session.markPartChanged(part.part.partName.toString());
    this.notify();
    return part;
  }

  /**
   * Edit the XPaths part through its record; the part is written back, marked, and named as
   * touched. `fn` works on a copy: a refused write (REQ-013) leaves the cached record as it was.
   */
  updateXPaths(fn: (xpaths: XPathsPart) => void): XPathsPart {
    this.requireEditable();
    const part = this.ensureXPathsPart();
    const record = structuredClone(this.xpaths ?? readXPaths(part));
    fn(record);
    const problems = xpathsProblems(record).filter((p) => p.requirement === 'REQ-013');
    if (problems.length > 0) throw new OpenDopeError(problems[0]!.message, 'REQ-013');
    writeXPaths(part, record);
    this.cache.xpaths = undefined;
    this.markTouched(part);
    this.notify();
    return this.xpaths!;
  }

  updateConditions(fn: (conditions: ConditionsPart) => void): ConditionsPart {
    this.requireEditable();
    const part = this.ensureConditionsPart();
    const record = structuredClone(this.conditions ?? readConditions(part));
    fn(record);
    const problems = conditionsProblems(record, this.xpaths).filter((p) => p.requirement === 'REQ-028' || p.requirement === 'REQ-029');
    if (problems.length > 0) throw new OpenDopeError(problems[0]!.message, problems[0]!.requirement);
    writeConditions(part, record);
    this.cache.conditions = undefined;
    this.markTouched(part);
    this.notify();
    return this.conditions!;
  }

  freshEntryId(): string {
    return freshId('x', this.xpaths?.entries.map((e) => e.id) ?? []);
  }

  freshConditionId(): string {
    return freshId('c', this.conditions?.conditions.map((c) => c.id) ?? []);
  }

  /**
   * Add or replace an XPaths entry: a fresh id when none is given, the data part's store item id
   * when none is given (REQ-004 when there is no data part), the other attributes kept when the
   * entry existed.
   */
  upsertEntry(entry: { id?: string; xpath: string; storeItemID?: string; prefixMappings?: string; attrs?: Record<string, string> }): XPathEntry {
    const id = entry.id ?? this.freshEntryId();
    let storeItemID = entry.storeItemID;
    if (storeItemID === undefined) {
      const data = this.dataPart;
      if (!data.part) throw new OpenDopeError(`no data part to bind to: ${data.reason}`, 'REQ-004');
      storeItemID = data.part.id;
    }
    let out: XPathEntry | undefined;
    this.updateXPaths((x) => {
      const at = x.entries.findIndex((e) => e.id === id);
      const previous = at >= 0 ? x.entries[at]! : undefined;
      out = { id, xpath: entry.xpath, storeItemID: storeItemID!, prefixMappings: entry.prefixMappings ?? previous?.prefixMappings ?? '', attrs: { ...(previous?.attrs ?? {}), ...(entry.attrs ?? {}) } };
      if (at >= 0) x.entries[at] = out; else x.entries.push(out);
    });
    return out!;
  }

  /** Add or replace a condition; a fresh id when none is given. */
  upsertCondition(condition: { id?: string; operand: Operand; attrs?: Record<string, string> }): Condition {
    const id = condition.id ?? this.freshConditionId();
    let out: Condition | undefined;
    this.updateConditions((c) => {
      const at = c.conditions.findIndex((x) => x.id === id);
      const previous = at >= 0 ? c.conditions[at]! : undefined;
      out = { id, attrs: { ...(previous?.attrs ?? {}), ...(condition.attrs ?? {}) }, operand: condition.operand };
      if (at >= 0) c.conditions[at] = out; else c.conditions.push(out);
    });
    return out!;
  }

  /** Remove an entry, unless a tag or a condition uses it (the panel names the users). */
  deleteEntry(id: string, doc: PMNode = this.session.doc): void {
    const users = this.usersOfEntry(id, doc);
    if (users.length > 0) throw new OpenDopeError(`XPaths entry ${id} is used by ${users.join(', ')}`);
    this.updateXPaths((x) => { x.entries = x.entries.filter((e) => e.id !== id); });
  }

  deleteCondition(id: string, doc: PMNode = this.session.doc): void {
    const users = this.usersOfCondition(id, doc);
    if (users.length > 0) throw new OpenDopeError(`condition ${id} is used by ${users.join(', ')}`);
    this.updateConditions((c) => { c.conditions = c.conditions.filter((x) => x.id !== id); });
  }

  /** Who refers to an entry: tags (`od:xpath`, `od:repeat`, `od:RptPosCon`) and conditions (`xpathref`). */
  usersOfEntry(id: string, doc: PMNode = this.session.doc): string[] {
    return this.usesOfEntry(id, doc).map((u) => u.label);
  }

  /** The uses of an entry, each with how it is used: a control's binding, a repeat, a position condition, or a condition's operand. */
  usesOfEntry(id: string, doc: PMNode = this.session.doc): EntryUse[] {
    const out: EntryUse[] = [];
    for (const { tag, label } of this.tagsIn(doc)) {
      const t = readTag(tag);
      if (t.xpath === id) out.push({ kind: 'bound', label });
      if (t.role === 'repeat' && t.roleId === id) out.push({ kind: 'repeat', label });
      if (t.positionCondition === id) out.push({ kind: 'position', label });
    }
    for (const c of this.conditions?.conditions ?? []) if (refersToEntry(c.operand, id)) out.push({ kind: 'condition', label: `condition ${c.id}` });
    return out;
  }

  usersOfCondition(id: string, doc: PMNode = this.session.doc): string[] {
    const out: string[] = [];
    for (const { tag, label } of this.tagsIn(doc)) {
      const t = readTag(tag);
      if (t.role === 'condition' && t.roleId === id) out.push(label);
    }
    for (const c of this.conditions?.conditions ?? []) if (refersToCondition(c.operand, id)) out.push(`condition ${c.id}`);
    return out;
  }

  /** The controls tagged `od:component=id`. */
  usersOfComponent(id: string, doc: PMNode = this.session.doc): string[] {
    const out: string[] = [];
    for (const { tag, label } of this.tagsIn(doc)) {
      const t = readTag(tag);
      if (t.role === 'component' && t.roleId === id) out.push(label);
    }
    return out;
  }

  private tagsIn(doc: PMNode): { tag: string; label: string }[] {
    const out: { tag: string; label: string }[] = [];
    const groups = readControlEntries(this.session.store);
    doc.descendants((node, pos) => {
      if (node.type.name === 'sdtBlock' || node.type.name === 'sdtInline') out.push({ tag: sdtTag(node.attrs['sdtPr']), label: `the ${node.type.name === 'sdtBlock' ? 'block' : 'run'} control at ${pos}` });
      if (node.type.name === 'table_row' || node.type.name === 'table_cell') {
        for (const g of (node.attrs['sdt'] as { groupId: string; first: boolean }[] | null) ?? []) if (g.first) out.push({ tag: sdtTag(groups.get(g.groupId)?.sdtPr ?? null), label: `the ${node.type.name === 'table_row' ? 'row' : 'cell'} control ${g.groupId}` });
      }
      return true;
    });
    return out;
  }

  /** Every save that touched a part names this editor and its `XX.YYYY` version (REQ-011). */
  stampAuthoringTool(version: string): void {
    const part = this.openDopeParts.xpaths;
    if (!part) return;
    const record = readXPaths(part);
    record.authoringTool = APPLICATION_NAME;
    record.authoringToolVersion = version;
    writeXPaths(part, record);
    this.cache.xpaths = undefined;
    this.session.markPartChanged(part.part.partName.toString());
  }

  // ---------------------------------------------------------------- evaluation (section 4.7)

  get evaluation(): EvaluationContext {
    return { parts: this.parts, engine: this.session.pkg.xpathEngine, mode: this.mode, xpaths: this.xpaths, conditions: this.conditions };
  }

  /** What a processor would make of one control's properties. */
  verdictFor(sdtPr: unknown): Verdict {
    const tag = readTag(sdtTag(sdtPr));
    const binding = sdtBinding(sdtPr);
    const kind = sdtKind(sdtPr);
    const ctx = this.evaluation;
    const verdict: Verdict = { role: tag.role, warnings: [...tag.problems] };
    const attempt = (fn: () => void): void => {
      try { fn(); } catch (e) { verdict.error = e instanceof Error ? e.message : String(e); }
    };
    if (tag.role === 'condition' && tag.roleId !== undefined) {
      verdict.id = tag.roleId;
      if (binding && kind !== 'RepeatingSection') verdict.warnings.push('a processor does not evaluate od:condition on a control with a w:dataBinding (REQ-008)');
      attempt(() => { verdict.kept = evaluateCondition(ctx, tag.roleId!); });
    } else if (tag.role === 'repeat' && tag.roleId !== undefined) {
      verdict.id = tag.roleId;
      if (binding) verdict.warnings.push('a processor does not evaluate od:repeat on a control with a w:dataBinding (REQ-008)');
      attempt(() => {
        verdict.count = repeatCount(ctx, tag.roleId!);
        if (repeatLooksLikeContainer(ctx, tag.roleId!)) verdict.warnings.push(`the repeat selects one element whose children look repeated: a repeat expression must select the elements to repeat, not their container (REQ-038); the base is ${repeatBase(entryNamed(ctx, tag.roleId!).xpath)}`);
      });
    } else if (tag.role === 'component' && tag.roleId !== undefined) {
      verdict.id = tag.roleId;
      if (!this.components?.components.some((c) => c.id === tag.roleId)) verdict.warnings.push(`component ${tag.roleId} is not in the Components part`);
    } else if (kind === 'RepeatingSection') {
      verdict.role = 'repeatingSection';
      if (binding) attempt(() => { verdict.count = ctx.engine.select(repeatBase(binding.xpath), this.partFor(binding.storeItemID).document, prefixMap(binding.prefixMappings)).length; });
    }
    if (tag.positionCondition !== undefined) {
      verdict.role = tag.role === 'none' ? 'positionCondition' : verdict.role;
      attempt(() => { verdict.positionPredicate = entryNamed(ctx, tag.positionCondition!).xpath; });
    }
    // The bound value: the od:xpath entry where it exists, else the w:dataBinding (REQ-017).
    const entry = tag.xpath !== undefined ? ctx.xpaths?.entries.find((e) => e.id === tag.xpath) : undefined;
    if (entry) {
      attempt(() => {
        const value = evaluateEntryValue(ctx, entry.id);
        if (value === undefined) verdict.selectsNothing = true; else verdict.value = value;
      });
      if (binding && (binding.xpath !== entry.xpath || !sameId(binding.storeItemID, entry.storeItemID))) verdict.warnings.push(`the w:dataBinding (${binding.xpath}) and XPaths entry ${entry.id} (${entry.xpath}) differ (REC-002)`);
    } else if (tag.xpath !== undefined) {
      verdict.warnings.push(`od:xpath names entry ${tag.xpath}, which does not exist (REQ-017: the w:dataBinding applies)`);
      if (binding) attempt(() => { const v = evaluateBinding(ctx, binding); if (v === undefined) verdict.selectsNothing = true; else verdict.value = v; });
    } else if (binding && kind !== 'RepeatingSection' && kind !== 'Group') {
      attempt(() => { const v = evaluateBinding(ctx, binding); if (v === undefined) verdict.selectsNothing = true; else verdict.value = v; });
    }
    return verdict;
  }

  private partFor(storeItemID: string): CustomXmlPart {
    const part = this.partNamed(storeItemID);
    if (!part) throw new OpenDopeError(`store item ${storeItemID || '(none)'} is not a custom XML part of this package`);
    return part;
  }
}

function sameId(a: string, b: string): boolean {
  return a.replace(/[{}]/g, '').toLowerCase() === b.replace(/[{}]/g, '').toLowerCase();
}

function prefixMap(mappings: string): Record<string, string> {
  const out: Record<string, string> = {};
  const re = /xmlns(?::([\w.\-]+))?\s*=\s*(['"])(.*?)\2/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(mappings)) !== null) out[m[1] ?? ''] = m[3]!;
  return out;
}

function refersToEntry(o: Operand, id: string): boolean {
  switch (o.op) {
    case 'xpathref': return o.id === id;
    case 'conditionref': return false;
    case 'not': return refersToEntry(o.operand, id);
    default: return o.operands.some((x) => refersToEntry(x, id));
  }
}

function refersToCondition(o: Operand, id: string): boolean {
  switch (o.op) {
    case 'conditionref': return o.id === id;
    case 'xpathref': return false;
    case 'not': return refersToCondition(o.operand, id);
    default: return o.operands.some((x) => refersToCondition(x, id));
  }
}
