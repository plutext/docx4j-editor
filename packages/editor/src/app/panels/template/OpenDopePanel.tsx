// SPDX-License-Identifier: GPL-3.0-or-later
//
// The OpenDoPE tab (ED-003 section 4.6): the XPaths part's entries listed with what each selects
// and how it converts in the declared mode, added, edited and deleted (an entry in use names its
// users and stays; ids unique, REQ-013; an expression checked by evaluating it, REQ-014); the
// Conditions part's conditions as operator trees with an editor that adds, removes and reorders
// operands and picks entries and conditions from lists (a condition holds one operand; cycles
// and missing references refused, REQ-029 and REQ-030; ids unique, REQ-028), each shown with
// its boolean in the declared mode; the root's version and boolean conversion; the Questions
// and Components parts read-only. Every write goes through the template view (`updateXPaths`,
// `updateConditions`), which creates a missing part on first use and marks the part for the
// save; no undo for a part edit (section 9 item 7), said at the top.
import { useMemo, useState } from 'react';
import {
  evaluateEntryBoolean, evaluateEntryValue, evaluateCondition, OpenDopeError, OPENDOPE_VERSION,
  type EditorDocument, type XPathEntry, type Condition, type Operand, type BooleanMode, type EvaluationContext, type EntryUse,
} from '@docx4j/editor-model';
import type { CustomXmlPart } from '@docx4j/core-ts';
import { useEditorState, useStoreRevision } from '../../state.js';
import { useTemplate, partLabel, operandText } from '../../template.js';
import type { TabContext } from '../../perspectives.js';
import { QuestionsList, ComponentsList, EntryUses } from './opendope-lists.js';
import { Fold } from '../Fold.js';
import { t, useT } from '../../i18n.js';

type Outcome<T> = { value: T; error?: undefined } | { value?: undefined; error: string };

function attempt<T>(fn: () => T): Outcome<T> {
  try { return { value: fn() }; } catch (e) { return { error: e instanceof Error ? e.message : String(e) }; }
}

function withoutEmpty(attrs: Record<string, string>): Record<string, string> {
  return Object.fromEntries(Object.entries(attrs).filter(([, v]) => v !== ''));
}

export default function OpenDopePanel({ editor, session, onStatus }: TabContext) {
  useT();
  const state = useEditorState(editor);
  const revision = useStoreRevision(session);
  const tpl = useTemplate(session);
  const template = session.template;
  const [error, setError] = useState<string | null>(null);

  const edit = (what: string, fn: () => void): boolean => {
    try {
      fn();
      setError(null);
      onStatus(t('od.edit.done', { what }));
      return true;
    } catch (e) {
      const message = e instanceof OpenDopeError ? e.message : (e as Error).message;
      setError(message);
      onStatus(message);
      return false;
    }
  };

  const users = useMemo(() => {
    void revision;
    if (!tpl.loaded || !state) return { entry: (): EntryUse[] => [], condition: (): string[] => [], component: (): string[] => [] };
    return { entry: (id: string) => template.usesOfEntry(id, state.doc), condition: (id: string) => template.usersOfCondition(id, state.doc), component: (id: string) => template.usersOfComponent(id, state.doc) };
  }, [tpl.loaded, state, template, revision]);

  if (!tpl.loaded) {
    return <section className="panel opendope-panel" aria-label={t('tab.opendope')} data-testid="panel-opendope"><h2>{t('tab.opendope')}</h2>{tpl.error ? <p className="error" role="alert">{tpl.error}</p> : <p className="meta">{t('data.loading')}</p>}</section>;
  }

  const od = template.openDopeParts;
  const xpaths = template.xpaths;
  const conditions = template.conditions;
  const parts = template.parts;
  const data = template.dataPart;
  const editable = template.editable;
  const ctx = template.evaluation;
  const problems = template.problems;

  return (
    <section className="panel opendope-panel" aria-label={t('tab.opendope')} data-testid="panel-opendope">
      <h2>{t('tab.opendope')} <span className="meta">{t('od.meta')}</span></h2>
      <dl className="od-root" data-testid="od-root">
        <dt>{t('od.xpathsPart')}</dt><dd>{od.xpaths ? partLabel(od.xpaths) : <span className="meta">{t('od.xpathsPart.none', { version: OPENDOPE_VERSION })}</span>}</dd>
        {xpaths && (
          <>
            <dt>{t('od.version')}</dt><dd data-testid="od-version">{xpaths.version ?? <span className="warning">{t('od.version.none')}</span>}</dd>
            <dt>{t('od.mode')}</dt>
            <dd>
              <select aria-label={t('od.mode')} data-testid="od-mode" value={xpaths.booleanConversion ?? ''} disabled={!editable.ok}
                onChange={(e) => edit(t('od.mode.set', { mode: e.target.value || t('od.mode.none') }), () => template.updateXPaths((x) => { if (e.target.value) x.booleanConversion = e.target.value as BooleanMode; else delete x.booleanConversion; }))}>
                <option value="">{t('od.mode.default')}</option>
                <option value="java">java</option><option value="xpath1">xpath1</option><option value="xpath2">xpath2</option>
              </select>
              {template.mode === 'xpath2' && !template.engineLoaded && <span className="warning">{t('od.mode.needsEngine')}</span>}
            </dd>
            {xpaths.authoringTool && <><dt>{t('od.authoringTool')}</dt><dd>{xpaths.authoringTool} {xpaths.authoringToolVersion}</dd></>}
          </>
        )}
        <dt>{t('od.conditionsPart')}</dt><dd>{od.conditions ? partLabel(od.conditions) : <span className="meta">{t('od.conditionsPart.none')}</span>}</dd>
      </dl>
      {!editable.ok && <p className="error" data-testid="od-readonly">{t('od.readOnly', { reason: editable.reason ?? '' })}</p>}
      {problems.length > 0 && <ul className="warnings" data-testid="od-problems">{problems.map((p) => <li key={`${p.part}-${p.id ?? ''}-${p.requirement}-${p.message}`}>{p.message} ({p.requirement})</li>)}</ul>}
      {error && <p className="error" role="alert" data-testid="od-error">{error}</p>}
      <XPathsSection entries={xpaths?.entries ?? []} parts={parts} dataPart={data.part} ctx={ctx} template={template} editable={editable.ok} users={users.entry} edit={edit} />
      <ConditionsSection conditions={conditions?.conditions ?? []} entries={xpaths?.entries ?? []} ctx={ctx} template={template} editable={editable.ok} users={users.condition} edit={edit} />
      {od.questions && template.questions && <QuestionsList questions={template.questions} xpaths={xpaths} meta={partLabel(od.questions)} />}
      {od.components && template.components && <ComponentsList components={template.components} meta={partLabel(od.components)} users={users.component} />}
    </section>
  );
}

// ---------------------------------------------------------------- the XPaths part

interface EntryDraft { id: string; xpath: string; storeItemID: string; prefixMappings: string; name: string; description: string }

function XPathsSection({ entries, parts, dataPart, ctx, template, editable, users, edit }: {
  entries: XPathEntry[]; parts: CustomXmlPart[]; dataPart: CustomXmlPart | undefined; ctx: EvaluationContext; template: EditorDocument['template']; editable: boolean;
  users: (id: string) => EntryUse[]; edit: (what: string, fn: () => void) => boolean;
}) {
  const [draft, setDraft] = useState<EntryDraft | null>(null);
  const [editingId, setEditingId] = useState<string | null>(null);   // null: a new entry
  const [check, setCheck] = useState<string | null>(null);

  const open = (entry?: XPathEntry): void => {
    setCheck(null);
    if (entry) {
      setEditingId(entry.id);
      setDraft({ id: entry.id, xpath: entry.xpath, storeItemID: entry.storeItemID, prefixMappings: entry.prefixMappings, name: entry.attrs['name'] ?? '', description: entry.attrs['description'] ?? '' });
    } else {
      setEditingId(null);
      setDraft({ id: template.freshEntryId(), xpath: '', storeItemID: dataPart?.id ?? parts[0]?.id ?? '', prefixMappings: '', name: '', description: '' });
    }
  };

  /** REQ-014: the expression is well-formed when the engine evaluates it against the part. */
  const validate = (d: EntryDraft): string | null => {
    if (!d.id.trim()) return t('od.entry.noId');
    if (!d.xpath.trim()) return t('od.entry.noExpression');
    const part = template.partNamed(d.storeItemID);
    if (!part) return t('od.entry.noPart');
    try { part.selectNodes(d.xpath.trim(), d.prefixMappings || undefined); } catch (e) { return t('od.entry.notEvaluated', { error: (e as Error).message }); }
    return null;
  };

  const apply = (): void => {
    if (!draft) return;
    const problem = validate(draft);
    if (problem) { setCheck(problem); return; }
    const record: XPathEntry = { id: draft.id.trim(), xpath: draft.xpath.trim(), storeItemID: draft.storeItemID, prefixMappings: draft.prefixMappings.trim(), attrs: {} };
    const ok = edit(t(editingId === null ? 'od.entry.added' : 'od.entry.updated', { id: record.id }), () => template.updateXPaths((x) => {
      const at = editingId === null ? -1 : x.entries.findIndex((e) => e.id === editingId);
      const previous = at >= 0 ? x.entries[at]! : undefined;
      const kept = Object.fromEntries(Object.entries(previous?.attrs ?? {}).filter(([k]) => k !== 'name' && k !== 'description'));
      record.attrs = { ...kept, ...withoutEmpty({ name: draft.name.trim(), description: draft.description.trim() }) };
      if (at >= 0) x.entries[at] = record; else x.entries.push(record);
    }));
    if (ok) { setDraft(null); setEditingId(null); setCheck(null); }
  };

  const remove = (id: string): void => {
    const u = users(id);
    if (u.length > 0) { edit('', () => { throw new OpenDopeError(t('od.entry.inUse', { id, users: u.map((x) => x.label).join(', ') })); }); return; }
    edit(t('od.entry.deleted', { id }), () => template.deleteEntry(id));
  };

  // The boolean column is for an entry that is, or might be, a condition's: an entry used as a repeat or a binding
  // has no boolean worth showing (Jason, ED-005 section 12.84), so its cell stays empty and nothing is evaluated.
  const evaluated = (e: XPathEntry, kinds: EntryUse['kind'][]): { value: string; valueClass: string; boolean: string } => {
    const v = attempt(() => evaluateEntryValue(ctx, e.id));
    const asBoolean = !kinds.includes('repeat') && !kinds.includes('bound');
    const b = asBoolean ? attempt(() => evaluateEntryBoolean(ctx, e.id)) : undefined;
    return {
      value: v.error ? t('od.value.error', { error: v.error }) : v.value === undefined ? t('od.value.selectsNothing') : v.value === '' ? t('od.value.empty') : v.value.length > 40 ? `${v.value.slice(0, 40)}…` : v.value,
      // An entry that selects nothing or does not evaluate stands out: it is what a template author most needs to see.
      valueClass: v.error ? 'error' : v.value === undefined ? 'selects-nothing' : 'meta',
      boolean: b === undefined ? '' : b.error ? t('od.value.error', { error: b.error }) : String(b.value),
    };
  };

  return (
    <Fold id="opendope.xpaths" level={3} className="od-section" testId="od-xpaths" title={t('data.opendope.xpaths')} meta={t('od.xpaths.meta', { count: entries.length, mode: ctx.mode })}>
      {entries.length === 0 ? <p className="meta">{t('od.xpaths.none')}</p> : (
        <table className="od-table" aria-label={t('od.xpaths.table')}>
          <thead><tr><th>{t('od.col.id')}</th><th>{t('od.col.expression')}</th><th>{t('od.col.part')}</th><th>{t('od.col.name')}</th><th>{t('od.col.value')}</th><th>{t('od.col.boolean')}</th><th>{t('od.col.usedAs')}</th><th></th></tr></thead>
          <tbody>
            {entries.map((e) => {
              const u = users(e.id);
              const part = template.partNamed(e.storeItemID);
              const kinds = [...new Set(u.map((x) => x.kind))];
              const ev = evaluated(e, kinds);
              return (
                <tr key={e.id} data-testid={`od-entry-${e.id}`} className={kinds.map((k) => `use-${k}`).join(' ')}>
                  <td><b>{e.id}</b></td>
                  <td><code>{e.xpath}</code>{e.prefixMappings && <div className="meta"><code>{e.prefixMappings}</code></div>}</td>
                  <td className="meta">{part ? partLabel(part).split(' · ')[0] : <span className="error">{t('od.part.notInPackage', { id: e.storeItemID || t('od.part.none') })}</span>}</td>
                  <td>{e.attrs['name'] ?? ''}{e.attrs['description'] ? <div className="meta">{e.attrs['description']}</div> : null}{Object.entries(e.attrs).filter(([k]) => k !== 'name' && k !== 'description').map(([k, v]) => <div key={k} className="meta">{k}={v}</div>)}</td>
                  <td className={ev.valueClass} data-testid={`od-entry-${e.id}-value`}>{ev.value}</td>
                  <td className="meta" data-testid={`od-entry-${e.id}-boolean`}>{ev.boolean}</td>
                  <td data-testid={`od-entry-${e.id}-uses`}><EntryUses uses={u} /></td>
                  <td className="actions">
                    <button type="button" className="link" disabled={!editable} onClick={() => open(e)}>{t('od.edit')}</button>
                    <button type="button" className="link" disabled={!editable} aria-label={t('od.entry.delete.label', { id: e.id })} onClick={() => remove(e.id)}>{t('od.delete')}</button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}
      {draft ? (
        <div className="od-form" data-testid="od-entry-form">
          <h4>{editingId === null ? t('od.entry.new') : t('od.entry.heading', { id: editingId })}</h4>
          <label>{t('od.form.id')} <input type="text" aria-label={t('od.entry.id.label')} data-testid="od-entry-id" value={draft.id} disabled={editingId !== null} onChange={(e) => setDraft({ ...draft, id: e.target.value })} /></label>
          <label>{t('od.form.expression')} <input type="text" aria-label={t('od.entry.expression.label')} data-testid="od-entry-xpath" value={draft.xpath} placeholder="/invoice/customer/name" onChange={(e) => setDraft({ ...draft, xpath: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); apply(); } }} /></label>
          <label>{t('od.form.part')}
            <select aria-label={t('od.entry.part.label')} data-testid="od-entry-part" value={draft.storeItemID} onChange={(e) => setDraft({ ...draft, storeItemID: e.target.value })}>
              {parts.length === 0 && <option value="">{t('od.entry.part.none')}</option>}
              {parts.map((p) => <option key={p.id} value={p.id}>{dataPart === p ? t('data.dataPart', { part: partLabel(p) }) : partLabel(p)}</option>)}
            </select>
          </label>
          <label>{t('od.form.prefixes')} <input type="text" aria-label={t('od.entry.prefixes.label')} value={draft.prefixMappings} placeholder="xmlns:ns0='…'" onChange={(e) => setDraft({ ...draft, prefixMappings: e.target.value })} /></label>
          <label>{t('od.form.name')} <input type="text" aria-label={t('od.entry.name.label')} data-testid="od-entry-name" value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} /></label>
          <label>{t('od.form.description')} <input type="text" aria-label={t('od.entry.description.label')} value={draft.description} onChange={(e) => setDraft({ ...draft, description: e.target.value })} /></label>
          {check && <p className="error" role="alert" data-testid="od-entry-check">{check}</p>}
          <div className="pane-actions">
            <button type="button" data-testid="od-entry-apply" onClick={apply}>{t('od.apply')}</button>
            <button type="button" onClick={() => { setDraft(null); setCheck(null); }}>{t('od.cancel')}</button>
            <span className="meta">{t('od.entry.note')}</span>
          </div>
        </div>
      ) : <div className="pane-actions"><button type="button" data-testid="od-entry-new" disabled={!editable || parts.length === 0} onClick={() => open()}>{t('od.entry.new')}</button>{parts.length === 0 && <span className="meta">{t('od.entry.noParts')}</span>}</div>}
    </Fold>
  );
}

// ---------------------------------------------------------------- the Conditions part

function ConditionsSection({ conditions, entries, ctx, template, editable, users, edit }: {
  conditions: Condition[]; entries: XPathEntry[]; ctx: EvaluationContext; template: EditorDocument['template']; editable: boolean;
  users: (id: string) => string[]; edit: (what: string, fn: () => void) => boolean;
}) {
  const [draft, setDraft] = useState<{ id: string; name: string; operand: Operand } | null>(null);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [check, setCheck] = useState<string | null>(null);

  const open = (c?: Condition): void => {
    setCheck(null);
    if (c) { setEditingId(c.id); setDraft({ id: c.id, name: c.attrs['name'] ?? '', operand: c.operand }); }
    else { setEditingId(null); setDraft({ id: template.freshConditionId(), name: '', operand: { op: 'xpathref', id: entries[0]?.id ?? '' } }); }
  };

  const incomplete = (o: Operand): string | null => {
    switch (o.op) {
      case 'xpathref': return o.id === '' ? t('od.condition.pickEntry') : null;
      case 'conditionref': return o.id === '' ? t('od.condition.pickCondition') : null;
      case 'not': return incomplete(o.operand);
      default: return o.operands.length === 0 ? t('od.condition.needsOperand', { op: o.op }) : o.operands.map(incomplete).find((x) => x !== null) ?? null;
    }
  };

  const apply = (): void => {
    if (!draft) return;
    if (!draft.id.trim()) { setCheck(t('od.condition.noId')); return; }
    const problem = incomplete(draft.operand);
    if (problem) { setCheck(problem); return; }
    const id = draft.id.trim();
    const ok = edit(t(editingId === null ? 'od.condition.added' : 'od.condition.updated', { id }), () => template.updateConditions((c) => {
      const at = editingId === null ? -1 : c.conditions.findIndex((x) => x.id === editingId);
      const previous = at >= 0 ? c.conditions[at]! : undefined;
      const kept = Object.fromEntries(Object.entries(previous?.attrs ?? {}).filter(([k]) => k !== 'name'));
      const record: Condition = { id, attrs: { ...kept, ...withoutEmpty({ name: draft.name.trim() }) }, operand: draft.operand };
      if (at >= 0) c.conditions[at] = record; else c.conditions.push(record);
    }));
    if (ok) { setDraft(null); setEditingId(null); setCheck(null); }
  };

  const remove = (id: string): void => {
    const u = users(id);
    if (u.length > 0) { edit('', () => { throw new OpenDopeError(t('od.condition.inUse', { id, users: u.join(', ') })); }); return; }
    edit(t('od.condition.deleted', { id }), () => template.deleteCondition(id));
  };

  const verdict = (id: string): string => {
    const b = attempt(() => evaluateCondition(ctx, id));
    return b.error ? t('od.value.error', { error: b.error }) : b.value ? t('od.condition.true') : t('od.condition.false');
  };

  return (
    <Fold id="opendope.conditions" level={3} className="od-section" testId="od-conditions" title={t('data.opendope.conditions')} meta={t('od.conditions.meta', { count: conditions.length, mode: ctx.mode })}>
      {conditions.length === 0 ? <p className="meta">{t('od.conditions.none')}</p> : (
        <table className="od-table" aria-label={t('od.conditions.table')}>
          <thead><tr><th>{t('od.col.id')}</th><th>{t('od.col.condition')}</th><th>{t('od.col.name')}</th><th>{t('od.col.verdict')}</th><th>{t('od.col.usedBy')}</th><th></th></tr></thead>
          <tbody>
            {conditions.map((c) => {
              const u = users(c.id);
              return (
                <tr key={c.id} data-testid={`od-condition-${c.id}`}>
                  <td><b>{c.id}</b></td>
                  <td><code>{operandText(c.operand, template.xpaths)}</code></td>
                  <td>{c.attrs['name'] ?? ''}{Object.entries(c.attrs).filter(([k]) => k !== 'name').map(([k, v]) => <div key={k} className="meta">{k}={v}</div>)}</td>
                  <td className="meta" data-testid={`od-condition-${c.id}-verdict`}>{verdict(c.id)}</td>
                  <td className="meta">{u.length === 0 ? t('od.unused') : u.join(', ')}</td>
                  <td className="actions">
                    <button type="button" className="link" disabled={!editable} onClick={() => open(c)}>{t('od.edit')}</button>
                    <button type="button" className="link" disabled={!editable} aria-label={t('od.condition.delete.label', { id: c.id })} onClick={() => remove(c.id)}>{t('od.delete')}</button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}
      {draft ? (
        <div className="od-form" data-testid="od-condition-form">
          <h4>{editingId === null ? t('od.condition.new') : t('od.condition.heading', { id: editingId })}</h4>
          <label>{t('od.form.id')} <input type="text" aria-label={t('od.condition.id.label')} data-testid="od-condition-id" value={draft.id} disabled={editingId !== null} onChange={(e) => setDraft({ ...draft, id: e.target.value })} /></label>
          <label>{t('od.form.name')} <input type="text" aria-label={t('od.condition.name.label')} data-testid="od-condition-name" value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} /></label>
          <div className="operand-tree" data-testid="od-operand-tree">
            <OperandEditor value={draft.operand} onChange={(operand) => setDraft({ ...draft, operand })} entries={entries} conditions={conditions.filter((c) => c.id !== draft.id)} ctx={ctx} path="root" />
          </div>
          <p className="meta">{t('od.condition.reads')} <code>{operandText(draft.operand, template.xpaths)}</code></p>
          {check && <p className="error" role="alert" data-testid="od-condition-check">{check}</p>}
          <div className="pane-actions">
            <button type="button" data-testid="od-condition-apply" onClick={apply}>{t('od.apply')}</button>
            <button type="button" onClick={() => { setDraft(null); setCheck(null); }}>{t('od.cancel')}</button>
            <span className="meta">{t('od.condition.note')}</span>
          </div>
        </div>
      ) : <div className="pane-actions"><button type="button" data-testid="od-condition-new" disabled={!editable || entries.length === 0} onClick={() => open()}>{t('od.condition.new')}</button>{entries.length === 0 && <span className="meta">{t('od.condition.needsEntries')}</span>}</div>}
    </Fold>
  );
}

// ---------------------------------------------------------------- the operator tree (table 6)

/** The operators, each labelled by `od.op.<op>`. */
const OPS: readonly Operand['op'][] = ['xpathref', 'conditionref', 'not', 'and', 'or'];

/** Another operator in place of this one, keeping what it can of the operands. */
function convert(value: Operand, op: Operand['op'], entries: XPathEntry[], conditions: Condition[]): Operand {
  if (value.op === op) return value;
  switch (op) {
    case 'xpathref': return { op, id: entries[0]?.id ?? '' };
    case 'conditionref': return { op, id: conditions[0]?.id ?? '' };
    case 'not': return { op, operand: value.op === 'and' || value.op === 'or' ? value.operands[0] ?? { op: 'xpathref', id: entries[0]?.id ?? '' } : value };
    default: return { op, operands: value.op === 'and' || value.op === 'or' ? value.operands : value.op === 'not' ? [value.operand] : [value] };
  }
}

export function OperandEditor({ value, onChange, entries, conditions, ctx, path }: { value: Operand; onChange: (o: Operand) => void; entries: XPathEntry[]; conditions: Condition[]; ctx: EvaluationContext; path: string }) {
  const leafBoolean = (): string => {
    if (value.op === 'xpathref' && value.id) { const b = attempt(() => evaluateEntryBoolean(ctx, value.id)); return b.error ? t('od.value.error', { error: b.error }) : `= ${String(b.value)}`; }
    if (value.op === 'conditionref' && value.id) { const b = attempt(() => evaluateCondition(ctx, value.id)); return b.error ? t('od.value.error', { error: b.error }) : `= ${String(b.value)}`; }
    return '';
  };
  const children = value.op === 'and' || value.op === 'or' ? value.operands : value.op === 'not' ? [value.operand] : [];
  const setChild = (i: number, child: Operand): void => {
    if (value.op === 'not') onChange({ op: 'not', operand: child });
    else if (value.op === 'and' || value.op === 'or') onChange({ op: value.op, operands: value.operands.map((o, k) => (k === i ? child : o)) });
  };
  const list = value.op === 'and' || value.op === 'or' ? value : undefined;
  const move = (i: number, by: number): void => {
    if (!list) return;
    const ops = [...list.operands];
    const j = i + by;
    if (j < 0 || j >= ops.length) return;
    [ops[i], ops[j]] = [ops[j]!, ops[i]!];
    onChange({ op: list.op, operands: ops });
  };
  return (
    <div className="operand" data-testid={`operand-${path}`}>
      <div className="operand-row">
        <select aria-label={t('od.operand.operator', { path })} data-testid={`operand-${path}-op`} value={value.op} onChange={(e) => onChange(convert(value, e.target.value as Operand['op'], entries, conditions))}>
          {OPS.map((op) => <option key={op} value={op}>{t(`od.op.${op}`)}</option>)}
        </select>
        {value.op === 'xpathref' && (
          <select aria-label={t('od.operand.entry', { path })} data-testid={`operand-${path}-id`} value={value.id} onChange={(e) => onChange({ op: 'xpathref', id: e.target.value })}>
            <option value="">…</option>
            {entries.map((e) => <option key={e.id} value={e.id}>{e.id} · {e.xpath}</option>)}
          </select>
        )}
        {value.op === 'conditionref' && (
          <select aria-label={t('od.operand.condition', { path })} data-testid={`operand-${path}-id`} value={value.id} onChange={(e) => onChange({ op: 'conditionref', id: e.target.value })}>
            <option value="">…</option>
            {conditions.map((c) => <option key={c.id} value={c.id}>{c.id}{c.attrs['name'] ? ` · ${c.attrs['name']}` : ''}</option>)}
          </select>
        )}
        <span className="meta">{leafBoolean()}</span>
        {list && <button type="button" className="link" data-testid={`operand-${path}-add`} onClick={() => onChange({ op: list.op, operands: [...list.operands, { op: 'xpathref', id: entries[0]?.id ?? '' }] })}>{t('od.operand.add')}</button>}
      </div>
      {children.length > 0 && (
        <ul className="operands">
          {children.map((child, i) => (
            <li key={i}>
              {list && (
                <span className="operand-tools">
                  <button type="button" className="link" aria-label={t('od.operand.up', { n: i + 1 })} disabled={i === 0} onClick={() => move(i, -1)}>↑</button>
                  <button type="button" className="link" aria-label={t('od.operand.down', { n: i + 1 })} disabled={i === list.operands.length - 1} onClick={() => move(i, 1)}>↓</button>
                  <button type="button" className="link" aria-label={t('od.operand.remove', { n: i + 1 })} data-testid={`operand-${path}-${i}-remove`} disabled={list.operands.length === 1} onClick={() => onChange({ op: list.op, operands: list.operands.filter((_o, k) => k !== i) })}>×</button>
                </span>
              )}
              <OperandEditor value={child} onChange={(o) => setChild(i, o)} entries={entries} conditions={conditions} ctx={ctx} path={`${path}-${i}`} />
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
