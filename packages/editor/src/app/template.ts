// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Template perspective's shared pieces (ED-003 section 4.9): the template view loaded once
// per session with its XPath engine, the labels the tabs share, and the running of a command
// with the reason a refusal shows in the status line.
import { useEffect, useState } from 'react';
// The XPath 3.1 engine's module, imported here so that it is in the Template chunk (this file is
// reached only through the lazily loaded tabs) and handed to core-ts's `FontoXPathEngine`, which
// cannot resolve a bare `import('fontoxpath')` in a browser (ED-003 section 11.2 item 21).
import * as fontoxpath from 'fontoxpath';
import { readControl, readTag, fragmentFactsOf, isFlatOpc, imageBase64Of, type EditorDocument, type ControlTarget, type Verdict, type Operand, type XPathsPart, type ConditionsPart, type FragmentFacts } from '@docx4j/editor-model';
import type { CustomXmlPart } from '@docx4j/core-ts';
import { kindLabelOf } from '../view/nodeviews/index.js';
import { t } from './i18n.js';

export interface TemplateStatus {
  loaded: boolean;
  error: string | null;
}

const loads = new WeakMap<EditorDocument, Promise<void>>();

/** `session.template.load()` once per session: the custom XML parts parsed and the FontoXPath engine installed as `pkg.xpathEngine`. */
export function loadTemplate(session: EditorDocument): Promise<void> {
  let p = loads.get(session);
  if (!p) { p = session.template.load({ fontoxpath }); loads.set(session, p); }
  return p;
}

/** `loadTemplate` for a tab: it renders once the load is done. */
export function useTemplate(session: EditorDocument): TemplateStatus {
  const [status, setStatus] = useState<TemplateStatus>(() => ({ loaded: session.template.loaded && session.template.engineLoaded, error: null }));
  useEffect(() => {
    let cancelled = false;
    const p = loadTemplate(session);
    p.then(() => { if (!cancelled) setStatus({ loaded: true, error: null }); }, (e: unknown) => { if (!cancelled) setStatus({ loaded: false, error: (e as Error).message }); });
    return () => { cancelled = true; };
  }, [session]);
  return status;
}

/** A control's form as the tabs name it: block, run, row, cell (`control.form.<form>`). */
export function formLabel(form: ControlTarget['form']): string {
  return t(`control.form.${form}`);
}

export function kindLabel(kind: string): string {
  return kindLabelOf(kind);
}

/** The short name a tree row or a badge shows for a control's properties. */
export function controlLabel(sdtPr: unknown): string {
  const r = readControl(sdtPr);
  return r.title || r.tag || kindLabel(r.kind);
}

/** A part as the panels name it: the root element's local name and the part name. */
export function partLabel(part: CustomXmlPart): string {
  const root = part.document.documentElement;
  const local = root ? root.localName ?? root.nodeName : t('template.part.empty');
  return `${local} · ${part.part.partName.toString()}`;
}

/** One line for a verdict (the tree's badge; the panel says more). */
export function verdictSummary(v: Verdict | undefined): string {
  if (!v) return '';
  if (v.error) return t('verdict.error', { error: v.error });
  const parts: string[] = [];
  if (v.kept !== undefined) parts.push(v.kept ? t('verdict.kept') : t('verdict.dropped'));
  if (v.count !== undefined) parts.push(v.count === 0 ? t('verdict.removedFalse') : t('verdict.instances', { count: v.count }));
  if (v.value !== undefined) parts.push(v.value === '' ? t('verdict.emptyValue') : t('verdict.value', { value: v.value.length > 40 ? `${v.value.slice(0, 40)}…` : v.value }));
  if (v.selectsNothing) parts.push(t('verdict.selectsNothing'));
  if (v.positionPredicate !== undefined) parts.push(t('verdict.position', { position: v.positionPredicate }));
  return parts.join(', ');
}

/** A condition's operand tree as one line: `and(x5, not(c6))`. */
export function operandText(o: Operand, xpaths: XPathsPart | undefined): string {
  switch (o.op) {
    case 'xpathref': { const e = xpaths?.entries.find((x) => x.id === o.id); return e ? `${o.id} [${e.xpath}]` : t('operand.missing', { id: o.id }); }
    case 'conditionref': return t('operand.condition', { id: o.id });
    case 'not': return `not(${operandText(o.operand, xpaths)})`;
    default: return `${o.op}(${o.operands.map((x) => operandText(x, xpaths)).join(', ')})`;
  }
}

export function conditionText(id: string, conditions: ConditionsPart | undefined, xpaths: XPathsPart | undefined): string {
  const c = conditions?.conditions.find((x) => x.id === id);
  return c ? operandText(c.operand, xpaths) : t('condition.missing');
}

/** The role badge text a tree row shows. */
export function roleBadge(sdtPr: unknown): string | undefined {
  const tag = readTag(readControl(sdtPr).tag);
  if (tag.role !== 'none') return tag.roleId ? t('role.withId', { role: t(`role.${tag.role}`), id: tag.roleId }) : t(`role.${tag.role}`);
  if (tag.positionCondition !== undefined) return t('role.position', { id: tag.positionCondition });
  if (readControl(sdtPr).binding || tag.xpath !== undefined) return t('role.bound');
  return undefined;
}

/** The data URL of a node's image data, for a thumbnail (ED-005 section 12.117). */
export function imageDataUrl(contentType: string, value: string): string {
  return `data:${contentType};base64,${imageBase64Of(value).replace(/\s+/g, '')}`;
}

/**
 * A node value's facts as a Word fragment (ED-005 proposal 46), read through the engine in the background: `null` while
 * they are read, undefined for a value that is no Flat OPC or does not load.
 */
export function useFragmentFacts(value: string | undefined): FragmentFacts | undefined | null {
  const fragment = value !== undefined && isFlatOpc(value);
  const [facts, setFacts] = useState<FragmentFacts | undefined | null>(fragment ? null : undefined);
  useEffect(() => {
    if (!fragment) { setFacts(undefined); return; }
    let current = true;
    setFacts(null);
    void fragmentFactsOf(value!).then((f) => { if (current) setFacts(f); });
    return () => { current = false; };
  }, [value, fragment]);
  return facts;
}

/** A fragment's facts in the words the Data and Control tabs share. */
export function fragmentSummary(facts: FragmentFacts | undefined | null): string {
  if (facts === null) return t('fragment.reading');
  if (facts === undefined) return t('fragment.unreadable');
  return t('fragment.summary', { size: Math.max(1, Math.round(facts.bytes / 1024)), paragraphs: facts.paragraphs, tables: facts.tables });
}

