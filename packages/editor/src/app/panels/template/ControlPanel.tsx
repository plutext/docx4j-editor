// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Control tab (ED-003 section 4.4): the control at the caret, or the selected group. Word's
// properties editable through `contentControl.properties`; the typed kind's own fields shown;
// the binding (`w:dataBinding` resolved to its part and the node it selects, the `od:xpath`
// entry beside it) edited through `contentControl.bind`; the OpenDoPE role (none, condition,
// repeat; component read-only) with its definition inline, set through `opendope.condition`,
// `opendope.repeat` and `opendope.positionCondition`; and the verdict of section 4.7. Nothing
// here writes the tree directly. A part write (an XPaths entry, a condition) happens in the same
// command and is not undone (section 9 item 7), which the panel says beside its Apply.
import { XHTML_CONTENT_TYPE } from '@docx4j/editor-model';
import { refreshXhtmlControls } from '../../../view/xhtmlbind.js';
import { refreshPictureControls } from '../../../view/picturebind.js';
import { refreshFragmentControls } from '../../../view/fragmentbind.js';
import { refreshBoundAtCaret } from '../../../view/boundrefresh.js';
import { useEffect, useMemo, useState, type ReactNode } from 'react';
import {
  controlAt, sdtPrOf, readControl, readTag, sdtItem, repeatBase, OpenDopeError, pictureBindingOf, pictureFactsOf, withPictureWidth, boundNodeOf,
  fragmentBindingOf, isFlatOpc, WORD_DOCUMENT_PROGID,
  type ControlProperties, type ControlReading, type TagReading, type Verdict, type EditorDocument, type SdtPrItem, type ContentControlAppearance,
} from '@docx4j/editor-model';
import type { DocxEditorView } from '../../../view/DocxEditorView.js';
import { useEditorState, useStoreRevision } from '../../state.js';
import { useTemplate, formLabel, kindLabel, partLabel, conditionText, verdictSummary, imageDataUrl, useFragmentFacts, fragmentSummary } from '../../template.js';
import { t, useT } from '../../i18n.js';
import type { TabContext } from '../../perspectives.js';
import { Fold } from '../Fold.js';

/** The kind element of `w:sdtPr` a typed control carries, by kind. */
const KIND_ELEMENT: Record<string, string> = {
  PlainText: 'text', CheckBox: 'checkbox', DatePicker: 'date', DropDownList: 'dropDownList', ComboBox: 'comboBox', Picture: 'picture',
  RepeatingSection: 'repeatingSection', RepeatingSectionItem: 'repeatingSectionItem', Group: 'group', BuildingBlockGallery: 'docPartObj',
  Citation: 'citation', Bibliography: 'bibliography', Equation: 'equation',
};

type Run = (name: string, args: unknown, refused?: string) => boolean;

/** A text box that commits on Enter or blur, and follows the model when it changes elsewhere. */
export function TextField({ value, label, testId, placeholder, onCommit, disabled }: { value: string; label: string; testId?: string; placeholder?: string; onCommit: (v: string) => void; disabled?: boolean }) {
  const [draft, setDraft] = useState(value);
  useEffect(() => { setDraft(value); }, [value]);
  const commit = (): void => { if (draft !== value) onCommit(draft); };
  return (
    <input type="text" aria-label={label} data-testid={testId} value={draft} placeholder={placeholder} disabled={disabled}
      onChange={(e) => setDraft(e.target.value)} onBlur={commit} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); commit(); } }} />
  );
}

function attempt<T>(fn: () => T): { value?: T; error?: string } {
  try { return { value: fn() }; } catch (e) { return { error: e instanceof Error ? e.message : String(e) }; }
}

export default function ControlPanel({ editor, session, onStatus }: TabContext) {
  useT();
  const state = useEditorState(editor);
  const revision = useStoreRevision(session);
  const tpl = useTemplate(session);
  const target = state ? controlAt(state) : undefined;
  const sdtPr = target ? sdtPrOf(target, session.controlEntry) : null;
  const reading = useMemo(() => readControl(sdtPr), [sdtPr]);
  const tag = useMemo(() => readTag(reading.tag), [reading.tag]);
  const verdict = useMemo(() => (tpl.loaded && target ? attempt(() => session.template.verdictFor(sdtPr)) : undefined), [tpl.loaded, target, sdtPr, session, revision]);   // eslint-disable-line react-hooks/exhaustive-deps

  const run: Run = (name, args, refused) => {
    try {
      const ok = editor.run(name, args);
      if (!ok) onStatus(refused ?? t('control.notApplied', { command: name }));
      return ok;
    } catch (e) {
      onStatus(e instanceof OpenDopeError ? e.message : t('control.error', { command: name, error: (e as Error).message }));
      return false;
    }
  };

  if (!target) {
    return (
      <section className="panel control-panel" aria-label={t('tab.control')} data-testid="panel-control">
        <h2>{t('tab.control')}</h2>
        <p className="meta">{t('control.none')}</p>
        {tpl.error && <p className="error" role="alert">{tpl.error}</p>}
      </section>
    );
  }

  return (
    <section className="panel control-panel" aria-label={t('tab.control')} data-testid="panel-control">
      <h2>{t('control.heading')} <span className="meta" data-testid="control-summary">{t('control.summary', { form: formLabel(target.form), kind: kindLabel(reading.kind), id: (reading.id ?? null) !== null ? String(reading.id) : t('control.id.none') })}</span></h2>
      {tpl.error && <p className="error" role="alert">{tpl.error}</p>}
      <Properties reading={reading} tag={tag} sdtPr={sdtPr} run={run} />
      {tpl.loaded && <Binding session={session} editor={editor} reading={reading} tag={tag} run={run} onStatus={onStatus} />}
      {tpl.loaded && (reading.kind === 'Picture' || tag.handler === 'picture') && <PictureFacts session={session} editor={editor} sdtPr={sdtPr} reading={reading} tag={tag} run={run} onStatus={onStatus} />}
      {tpl.loaded && <FragmentFacts session={session} sdtPr={sdtPr} />}
      {tpl.loaded && <Role session={session} reading={reading} tag={tag} verdict={verdict?.value} run={run} onStatus={onStatus} />}
      {tpl.loaded && <VerdictView session={session} verdict={verdict} />}
      {!tpl.loaded && !tpl.error && <p className="meta">{t('control.loading')}</p>}
    </section>
  );
}

// ---------------------------------------------------------------- Word's properties

function Properties({ reading, tag, sdtPr, run }: { reading: ControlReading; tag: TagReading; sdtPr: unknown; run: Run }) {
  const [tagError, setTagError] = useState<string | null>(null);
  const set = (patch: ControlProperties): void => { run('contentControl.properties', patch); };
  const setTag = (value: string): void => {
    const problems = readTag(value).problems.filter((p) => p.includes('REQ-007'));
    if (problems.length > 0) { setTagError(problems[0]!); return; }
    setTagError(null);
    set({ tag: value });
  };
  const kindItem = sdtItem(sdtPr, KIND_ELEMENT[reading.kind] ?? '', '*');
  return (
    <Fold id="control.properties" level={3} className="control-group" testId="control-properties" title={t('control.properties')} meta="w:sdtPr">
      <dl>
        <dt><label htmlFor="cc-title">{t('control.title')}</label></dt><dd><TextField value={reading.title} label={t('control.title')} testId="cc-title" onCommit={(v) => set({ title: v })} /></dd>
        <dt><label htmlFor="cc-tag">{t('control.tag')}</label></dt>
        <dd>
          <TextField value={reading.tag} label={t('control.tag')} testId="cc-tag" onCommit={setTag} />
          {tagError && <p className="error" role="alert" data-testid="cc-tag-error">{tagError}</p>}
          {tag.params.length > 0 && (
            <ul className="tag-params" data-testid="cc-tag-params">
              {tag.params.map((p, i) => {
                const processor = tag.processorSet.includes(p);
                return <li key={`${p.key}-${i}`} className={processor ? 'processor-set' : ''}><span className="attr-name">{p.key}</span>=<span className="attr-value">{p.value}</span>{processor ? <span className="meta">{t('control.tag.processorSet')}</span> : null}</li>;
              })}
            </ul>
          )}
          {tag.problems.filter((p) => !p.includes('REQ-077')).map((p) => <p key={p} className="error">{p}</p>)}
        </dd>
        <dt>{t('control.kind')}</dt><dd>{kindLabel(reading.kind)}</dd>
        <dt>{t('control.appearance')}</dt>
        <dd>
          <select aria-label={t('control.appearance')} data-testid="cc-appearance" value={reading.appearance} onChange={(e) => set({ appearance: e.target.value as ContentControlAppearance })}>
            {(['BoundingBox', 'Tags', 'Hidden'] as const).map((a) => <option key={a} value={a}>{t(`control.appearance.${a}`)}</option>)}
          </select>
        </dd>
        <dt>{t('control.colour')}</dt>
        <dd>
          <input type="color" aria-label={t('control.colour')} data-testid="cc-color" value={reading.color || '#000000'} onChange={(e) => set({ color: e.target.value.toUpperCase() })} />
          {reading.color ? <button type="button" className="link" onClick={() => set({ color: '' })}>{t('control.colour.none')}</button> : <span className="meta"> {t('control.colour.none')}</span>}
        </dd>
        <dt>{t('control.lock')}</dt>
        <dd>
          <label><input type="checkbox" data-testid="cc-cannot-delete" checked={reading.cannotDelete} onChange={(e) => set({ cannotDelete: e.target.checked })} /> {t('control.lock.cannotDelete')}</label>
          <label><input type="checkbox" data-testid="cc-cannot-edit" checked={reading.cannotEdit} onChange={(e) => set({ cannotEdit: e.target.checked })} /> {t('control.lock.cannotEdit')}</label>
        </dd>
        <dt>{t('control.flags')}</dt>
        <dd>
          <label><input type="checkbox" checked={reading.removeWhenEdited} onChange={(e) => set({ removeWhenEdited: e.target.checked })} /> {t('control.flags.removeWhenEdited')}</label>
          <label><input type="checkbox" checked={reading.showingPlaceholder} onChange={(e) => set({ showingPlaceholder: e.target.checked })} /> {t('control.flags.showingPlaceholder')}</label>
          {reading.kind === 'PlainText' && <label><input type="checkbox" checked={reading.multiLine} onChange={(e) => set({ multiLine: e.target.checked })} /> {t('control.flags.multiLine')}</label>}
        </dd>
        {kindItem && <><dt>{kindLabel(reading.kind)}</dt><dd><KindFields item={kindItem} /></dd></>}
      </dl>
    </Fold>
  );
}

/** The typed kind's own fields, read from its element as it is (editing them is a later step). */
function KindFields({ item }: { item: SdtPrItem }) {
  const entries = Object.entries(item.value).filter(([k]) => k !== 'TYPE_NAME');
  if (entries.length === 0) return <span className="meta">{t('control.kindFields.none')}</span>;
  return (
    <ul className="kind-fields" data-testid="cc-kind-fields">
      {entries.map(([k, v]) => <li key={k}><span className="attr-name">{k}</span>: {show(v)}</li>)}
    </ul>
  );
}

function show(v: unknown): ReactNode {
  if (v === null || v === undefined) return '';
  if (Array.isArray(v)) return <ul>{v.map((x, i) => <li key={i}>{show(x)}</li>)}</ul>;
  if (typeof v === 'object') {
    const o = v as Record<string, unknown>;
    if ('val' in o && Object.keys(o).every((k) => k === 'val' || k === 'TYPE_NAME')) return String(o['val']);
    return Object.entries(o).filter(([k]) => k !== 'TYPE_NAME').map(([k, x]) => `${k}=${typeof x === 'object' ? JSON.stringify(x) : String(x)}`).join(' ');
  }
  return String(v);
}

// ---------------------------------------------------------------- the binding

function Binding({ session, editor, reading, tag, run, onStatus }: { session: EditorDocument; editor: DocxEditorView; reading: ControlReading; tag: TagReading; run: Run; onStatus: (s: string) => void }) {
  const template = session.template;
  const binding = reading.binding;
  const part = binding ? template.partNamed(binding.storeItemID) : undefined;
  const selected = binding
    ? !part ? { error: t('control.binding.notInPart', { id: binding.storeItemID }) }
      : attempt(() => { const n = part.selectSingleNode(binding.xpath, binding.prefixMappings); return n ? { xpath: n.xpath, text: n.text } : undefined; })
    : undefined;
  const entry = tag.xpath !== undefined ? template.xpaths?.entries.find((e) => e.id === tag.xpath) : undefined;
  const data = template.dataPart;
  const parts = template.parts;
  const [xpath, setXpath] = useState(binding?.xpath ?? '');
  const [prefixes, setPrefixes] = useState(binding?.prefixMappings ?? '');
  const [storeItemID, setStoreItemID] = useState(binding?.storeItemID ?? data.part?.id ?? '');
  useEffect(() => {
    setXpath(binding?.xpath ?? '');
    setPrefixes(binding?.prefixMappings ?? '');
    setStoreItemID(binding?.storeItemID ?? data.part?.id ?? '');
  }, [binding?.xpath, binding?.prefixMappings, binding?.storeItemID, data.part?.id]);
  const opendope = template.openDopeParts.xpaths !== undefined || tag.xpath !== undefined;
  const bind = (): void => {
    if (!xpath.trim()) { onStatus(t('control.binding.enterXPath')); return; }
    const bound = run('contentControl.bind', { xpath: xpath.trim(), ...(prefixes ? { prefixMappings: prefixes } : {}), ...(storeItemID ? { storeItemID } : {}) }, t('control.binding.notBound'));
    // A node holding markup binds as an XHTML control (ED-005 section 12.79), a Word document as Word's mapped rich text
    // (section 12.124): the control filled here as the switch and the Data tab's bind fill it (section 12.86), that control alone.
    if (bound) void refreshBoundAtCaret(editor, session).then((notes) => { for (const note of notes) onStatus(note); });
  };
  const entryAttr = (name: string, value: string): void => {
    if (!entry) return;
    try {
      template.upsertEntry({ id: entry.id, xpath: entry.xpath, storeItemID: entry.storeItemID, prefixMappings: entry.prefixMappings, attrs: { [name]: value } });
    } catch (e) { onStatus((e as Error).message); }
  };
  return (
    <Fold id="control.binding" level={3} className="control-group" testId="control-binding" title={t('control.binding')} meta={binding ? (binding.w15 ? 'w15:dataBinding' : 'w:dataBinding') : t('control.binding.none')}>
      {binding && (
        <dl>
          <dt>{t('control.binding.part')}</dt><dd>{part ? partLabel(part) : <span className="error">{t('control.binding.notInPackage', { id: binding.storeItemID || t('control.binding.noStoreItem') })}</span>}</dd>
          <dt>{t('control.binding.xpath')}</dt><dd><code>{binding.xpath}</code></dd>
          {binding.prefixMappings && <><dt>{t('control.binding.prefixes')}</dt><dd><code>{binding.prefixMappings}</code></dd></>}
          <dt>{t('control.binding.selects')}</dt>
          <dd data-testid="cc-binding-selects">
            {selected?.error ? <span className="error">{selected.error}</span>
              : selected?.value ? <><code>{selected.value.xpath}</code> <span className="meta">{selected.value.text.length > 80 ? `${selected.value.text.slice(0, 80)}…` : selected.value.text}</span></>
                : <span className="error">{t('control.binding.selectsNothing')}</span>}
          </dd>
        </dl>
      )}
      {(binding || tag.xpath !== undefined) && (tag.contentType === XHTML_CONTENT_TYPE || !(selected?.value && isFlatOpc(selected.value.text))) && (
        <label className="xhtml-switch">
          <input type="checkbox" data-testid="cc-xhtml" checked={tag.contentType === XHTML_CONTENT_TYPE}
            onChange={(e) => { if (run('contentControl.xhtml', { on: e.target.checked })) void refreshXhtmlControls(editor.view, session).then((r) => { for (const note of r.notes) onStatus(note); }); }} />
          {' '}{t('control.binding.xhtml')} <span className="meta">{t('control.binding.xhtml.note')}</span>
        </label>
      )}
      {(binding || tag.xpath !== undefined) && (selected?.value && isFlatOpc(selected.value.text) || tag.progid === WORD_DOCUMENT_PROGID) && (
        <label className="xhtml-switch">
          {/* ED-005 proposal 46 (decision (e)): OpenDoPE's form, which a processor resolves, or Word's own mapping, which Word fills on open. */}
          <input type="checkbox" data-testid="cc-fragment" checked={tag.progid === WORD_DOCUMENT_PROGID}
            onChange={(e) => {
              if (!run('contentControl.wordFragment', { on: e.target.checked })) return;
              // The control filled with its Word document, either form (the engine's pass in a tree run, part 3).
              void refreshFragmentControls(editor, session, reading.id !== undefined ? { ids: new Set([reading.id]) } : {});
            }} />
          {' '}{t('control.binding.fragment')} <span className="meta">{t('control.binding.fragment.note')}</span>
        </label>
      )}
      {tag.xpath !== undefined && (
        <dl data-testid="cc-entry">
          <dt>od:xpath</dt>
          <dd>{entry ? <>{entry.id} <code>{entry.xpath}</code></> : <span className="error">{t('control.binding.noEntry', { id: tag.xpath })}</span>}</dd>
          {entry && (
            <>
              <dt>{t('control.binding.entryName')}</dt><dd><TextField value={entry.attrs['name'] ?? ''} label={t('control.binding.entryName.label')} onCommit={(v) => entryAttr('name', v)} /></dd>
              <dt>{t('control.binding.entryDescription')}</dt><dd><TextField value={entry.attrs['description'] ?? ''} label={t('control.binding.entryDescription.label')} onCommit={(v) => entryAttr('description', v)} /></dd>
              {Object.entries(entry.attrs).filter(([k]) => k !== 'name' && k !== 'description').map(([k, v]) => <div key={k} className="contents"><dt>{k}</dt><dd>{v}</dd></div>)}
            </>
          )}
        </dl>
      )}
      <div className="bind-form">
        <label>{t('control.binding.form.xpath')} <input type="text" aria-label={t('control.binding.form.xpath.label')} data-testid="bind-xpath" value={xpath} onChange={(e) => setXpath(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); bind(); } }} placeholder="/root/element" /></label>
        <label>{t('control.binding.form.prefixes')} <input type="text" aria-label={t('control.binding.form.prefixes.label')} data-testid="bind-prefixes" value={prefixes} onChange={(e) => setPrefixes(e.target.value)} placeholder="xmlns:ns0='…'" /></label>
        <label>{t('control.binding.form.part')}
          <select aria-label={t('control.binding.form.part.label')} data-testid="bind-part" value={storeItemID} onChange={(e) => setStoreItemID(e.target.value)}>
            {parts.map((p) => <option key={p.id} value={p.id}>{data.part === p ? t('control.binding.form.dataPart', { part: partLabel(p) }) : partLabel(p)}</option>)}
          </select>
        </label>
        <div className="pane-actions">
          <button type="button" data-testid="bind-apply" onClick={bind}>{binding ? t('control.binding.rebind') : t('control.binding.bind')}</button>
          {binding && <button type="button" data-testid="bind-remove" onClick={() => run('contentControl.bind', null)}>{t('control.binding.unbind')}</button>}
          <span className="meta">{opendope ? t('control.binding.note.opendope') : t('control.binding.note.plain')}</span>
        </div>
      </div>
    </Fold>
  );
}

// ---------------------------------------------------------------- the picture

/** How a handler control's picture is sized: its `width` absent (the drawing kept), `auto`, or a number of points. */
type WidthMode = 'keep' | 'auto' | 'points';

/**
 * The picture binding's facts (ED-005 section 12.117, proposal 45 part 3): Word's picture control or OpenDoPE's handler, the
 * node's image data (its size, the image's format and pixels, a thumbnail), and for the handler its `width`, set here: none
 * (REQ-061, the drawing kept and its picture replaced, by the processor), `auto` or points (REQ-062, a new inline picture at
 * its natural size, at most that wide). The picture is placed again at once at a new width.
 */
function PictureFacts({ session, editor, sdtPr, reading, tag, run, onStatus }: { session: EditorDocument; editor: DocxEditorView; sdtPr: unknown; reading: ControlReading; tag: TagReading; run: Run; onStatus: (s: string) => void }) {
  const picture = pictureBindingOf(sdtPr, session.template);
  const handler = tag.handler === 'picture';
  const current = picture?.width;
  const [mode, setMode] = useState<WidthMode>(current === undefined ? 'keep' : current === 'auto' ? 'auto' : 'points');
  const [points, setPoints] = useState(typeof current === 'number' ? String(current) : '');
  useEffect(() => {
    setMode(current === undefined ? 'keep' : current === 'auto' ? 'auto' : 'points');
    setPoints(typeof current === 'number' ? String(current) : '');
  }, [current]);
  const data = picture ? boundNodeOf(session, picture.binding) : undefined;
  const value = data?.nodeValue.trim() ?? '';
  const facts = value !== '' ? pictureFactsOf(value) : undefined;
  const applyWidth = (): void => {
    const width = mode === 'keep' ? undefined : mode === 'auto' ? 'auto' as const : Number(points);
    if (typeof width === 'number' && !(Number.isFinite(width) && width > 0)) { onStatus(t('control.picture.width.invalid')); return; }
    if (run('contentControl.properties', { tag: withPictureWidth(reading.tag, width) })) {
      const ids = reading.id !== undefined ? new Set([reading.id]) : undefined;
      void refreshPictureControls(editor.view, session, undefined, ids).then((r) => { for (const note of r.notes) onStatus(note); });
    }
  };
  return (
    <Fold id="control.picture" level={3} className="control-group" testId="control-picture" title={t('control.picture')} meta={handler ? 'od:Handler=picture' : 'w:picture'}>
      <dl>
        <dt>{t('control.picture.kind')}</dt>
        <dd data-testid="cc-picture-kind">{t(handler ? 'control.picture.kind.handler' : 'control.picture.kind.picture')}</dd>
        <dt>{t('control.picture.data')}</dt>
        <dd data-testid="cc-picture-data">
          {!picture ? <span className="meta">{t('control.binding.none')}</span>
            : !data ? <span className="error">{t('control.binding.selectsNothing')}</span>
              : value === '' ? <span className="meta">{t('control.picture.data.empty')}</span>
                : facts ? t('control.picture.data.size', { size: Math.max(1, Math.round(facts.bytes / 1024)) })
                  : <span className="error">{t('control.picture.data.notImage')}</span>}
        </dd>
        {facts && (
          <>
            <dt>{t('control.picture.image')}</dt>
            <dd data-testid="cc-picture-image">
              <img className="picture-thumbnail" src={imageDataUrl(facts.contentType, value)} alt={t('picture.alt', { xpath: picture!.binding.xpath })} />
              {' '}{facts.format === 'SVG' ? t('picture.sizeSvg') : t('picture.size', { format: facts.format, width: facts.widthPx, height: facts.heightPx })}
            </dd>
          </>
        )}
        {handler && (
          <>
            <dt>{t('control.picture.width')}</dt>
            <dd>
              <select aria-label={t('control.picture.width.label')} data-testid="cc-picture-width-mode" value={mode} onChange={(e) => setMode(e.target.value as WidthMode)}>
                <option value="keep">{t('control.picture.width.keep')}</option>
                <option value="auto">{t('control.picture.width.auto')}</option>
                <option value="points">{t('control.picture.width.points')}</option>
              </select>
              {mode === 'points' && (
                <input type="number" min="1" step="any" aria-label={t('control.picture.width.points.label')} data-testid="cc-picture-width" value={points}
                  onChange={(e) => setPoints(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); applyWidth(); } }} />
              )}
              <button type="button" data-testid="cc-picture-width-apply" onClick={applyWidth}>{t('control.picture.width.apply')}</button>
              <p className="meta">{t('control.picture.width.note')}</p>
            </dd>
          </>
        )}
      </dl>
    </Fold>
  );
}

// ---------------------------------------------------------------- the Word fragment

/**
 * A Word fragment's facts (ED-005 proposal 46): OpenDoPE's `od:progid` control or Word's `w15`-mapped rich text whose node
 * holds a Word document as Flat OPC. Its kind, what the document holds and keeps, and what binding it leaves behind, as
 * Word 365 measured (part 6(a)). Nothing shows for a `w15` mapping whose node holds plain text, Word's ordinary mapping.
 */
function FragmentFacts({ session, sdtPr }: { session: EditorDocument; sdtPr: unknown }) {
  const fragment = fragmentBindingOf(sdtPr, session.template);
  const data = fragment ? boundNodeOf(session, fragment.binding) : undefined;
  const value = data?.nodeValue ?? '';
  const holds = value !== '' && isFlatOpc(value);
  const facts = useFragmentFacts(holds ? value : undefined);
  if (!fragment || (fragment.shape === 'w15' && !holds)) return null;
  return (
    <Fold id="control.fragment" level={3} className="control-group" testId="control-fragment" title={t('control.fragment')} meta={fragment.shape === 'progid' ? 'od:progid=Word.Document' : 'w15:dataBinding'}>
      <dl>
        <dt>{t('control.kind')}</dt>
        <dd data-testid="cc-fragment-kind">{t(fragment.shape === 'progid' ? 'control.fragment.kind.progid' : 'control.fragment.kind.w15')}</dd>
        <dt>{t('control.fragment.data')}</dt>
        <dd data-testid="cc-fragment-data">
          {!data ? <span className="error">{t('control.binding.selectsNothing')}</span>
            : value === '' ? <span className="meta">{t('control.picture.data.empty')}</span>
              : !holds ? <span className="error">{t('control.fragment.data.notFragment')}</span>
                : fragmentSummary(facts)}
        </dd>
        {facts && (
          <>
            <dt>{t('control.fragment.kept')}</dt>
            <dd data-testid="cc-fragment-kept">{t('control.fragment.kept.value', { pictures: facts.pictures, notes: facts.notes, controls: facts.controls })}</dd>
            {(facts.comments > 0 || facts.sectionBreaks > 0) && (
              <>
                <dt>{t('control.fragment.dropped')}</dt>
                <dd data-testid="cc-fragment-dropped">{t('control.fragment.dropped.value', { comments: facts.comments, breaks: facts.sectionBreaks })}</dd>
              </>
            )}
          </>
        )}
      </dl>
    </Fold>
  );
}

// ---------------------------------------------------------------- the OpenDoPE role

function Role({ session, reading, tag, verdict, run, onStatus }: { session: EditorDocument; reading: ControlReading; tag: TagReading; verdict: Verdict | undefined; run: Run; onStatus: (s: string) => void }) {
  const template = session.template;
  const bound = reading.binding !== undefined && reading.kind !== 'RepeatingSection';
  const editable = template.editable;
  const [role, setRole] = useState<'condition' | 'repeat'>(tag.role === 'repeat' ? 'repeat' : 'condition');
  const [choice, setChoice] = useState<'existing' | 'new'>('existing');
  const [existing, setExisting] = useState('');
  const [expression, setExpression] = useState('');
  const [name, setName] = useState('');
  const [predicate, setPredicate] = useState('');
  useEffect(() => { setPredicate(verdict?.positionPredicate ?? ''); }, [verdict?.positionPredicate]);
  const conditions = template.conditions?.conditions ?? [];
  const entries = template.xpaths?.entries ?? [];
  const options = role === 'condition' ? conditions.map((c) => ({ id: c.id, text: `${c.id}${c.attrs['name'] ? ` · ${c.attrs['name']}` : ''}` })) : entries.map((e) => ({ id: e.id, text: `${e.id} · ${e.xpath}` }));
  const cannot = bound ? t('control.role.cannot.bound') : !editable.ok ? editable.reason! : tag.role === 'component' ? t('control.role.cannot.component') : undefined;

  const apply = (): void => {
    if (cannot) { onStatus(cannot); return; }
    const command = role === 'condition' ? 'opendope.condition' : 'opendope.repeat';
    // Another role first cleared (REQ-007): two commands, two undo steps.
    if (tag.role !== 'none' && tag.role !== role && !run(tag.role === 'condition' ? 'opendope.condition' : 'opendope.repeat', { id: null })) return;
    if (choice === 'existing') {
      if (!existing) { onStatus(t(role === 'condition' ? 'control.role.pickCondition' : 'control.role.pickEntry')); return; }
      run(command, { id: existing }, t(role === 'condition' ? 'control.role.missingCondition' : 'control.role.missingEntry', { id: existing }));
      return;
    }
    if (!expression.trim()) { onStatus(t('control.role.enterExpression')); return; }
    try {
      if (role === 'condition') {
        const entry = template.upsertEntry({ xpath: expression.trim(), ...(name ? { attrs: { name } } : {}) });
        run(command, { id: template.freshConditionId(), define: { op: 'xpathref', id: entry.id }, ...(name ? { name } : {}) });
      } else {
        run(command, { id: template.freshEntryId(), define: { xpath: expression.trim() }, ...(name ? { name } : {}) });
      }
      setExpression('');
      setName('');
    } catch (e) { onStatus((e as Error).message); }
  };
  const clear = (): void => { if (tag.role === 'condition' || tag.role === 'repeat') run(tag.role === 'condition' ? 'opendope.condition' : 'opendope.repeat', { id: null }); };
  const setPosition = (): void => {
    if (!predicate.trim()) { run('opendope.positionCondition', { id: null }); return; }
    run('opendope.positionCondition', { id: tag.positionCondition ?? template.freshEntryId(), define: { xpath: predicate.trim() } });
  };

  return (
    <Fold id="control.role" level={3} className="control-group" testId="control-role" title={t('control.role')} meta={<span data-testid="cc-role">{tag.role === 'none' ? t('control.role.none') : t(`role.${tag.role}`)}{tag.roleId ? ` ${tag.roleId}` : ''}</span>}>
      {tag.role === 'condition' && tag.roleId !== undefined && (
        <dl>
          <dt>{t('control.role.condition')}</dt><dd>{tag.roleId}: <code data-testid="cc-condition-text">{conditionText(tag.roleId, template.conditions, template.xpaths)}</code></dd>
          <dt>{t('control.role.verdict')}</dt><dd data-testid="cc-role-verdict">{verdict ? verdictSummary(verdict) : '…'}</dd>
        </dl>
      )}
      {tag.role === 'repeat' && tag.roleId !== undefined && (() => {
        const entry = entries.find((e) => e.id === tag.roleId);
        return (
          <dl>
            <dt>{t('control.role.repeat')}</dt><dd>{tag.roleId}: {entry ? <code>{entry.xpath}</code> : <span className="error">{t('control.role.noEntry')}</span>}</dd>
            {entry && <><dt>{t('control.role.base')}</dt><dd><code>{repeatBase(entry.xpath)}</code></dd></>}
            <dt>{t('control.role.count')}</dt><dd data-testid="cc-repeat-count">{verdict?.count !== undefined ? t('control.role.count.value', { count: verdict.count }) : verdict?.error ?? '…'}</dd>
          </dl>
        );
      })()}
      {tag.role === 'component' && (
        <dl><dt>{t('control.role.component')}</dt><dd>{tag.roleId} <span className="meta">{t('control.role.component.note')}</span></dd></dl>
      )}
      {(tag.role === 'condition' || tag.role === 'repeat') && <div className="pane-actions"><button type="button" data-testid="cc-role-clear" onClick={clear}>{t('control.role.clear')}</button><span className="meta">{t('control.role.clear.note')}</span></div>}
      {cannot ? <p className="meta" data-testid="cc-role-cannot">{cannot}</p> : (
        <div className="role-form">
          <label>{t('control.role.set')} <select aria-label={t('control.role.kind.label')} data-testid="cc-role-kind" value={role} onChange={(e) => { setRole(e.target.value as 'condition' | 'repeat'); setExisting(''); }}><option value="condition">{t('control.role.kind.condition')}</option><option value="repeat">{t('control.role.kind.repeat')}</option></select></label>
          <label><input type="radio" name="cc-role-choice" checked={choice === 'existing'} onChange={() => setChoice('existing')} /> {t('control.role.existing')}
            <select aria-label={role === 'condition' ? t('control.role.existing.condition') : t('control.role.existing.entry')} data-testid="cc-role-existing" value={existing} onChange={(e) => { setExisting(e.target.value); setChoice('existing'); }}>
              <option value="">…</option>
              {options.map((o) => <option key={o.id} value={o.id}>{o.text}</option>)}
            </select>
          </label>
          <label><input type="radio" name="cc-role-choice" checked={choice === 'new'} onChange={() => setChoice('new')} /> {t('control.role.new')}
            <input type="text" aria-label={t('control.role.new.expression')} data-testid="cc-role-expression" value={expression} placeholder={role === 'condition' ? '/invoice/paid = "true"' : '/invoice/items/item'} onChange={(e) => { setExpression(e.target.value); setChoice('new'); }} />
            <input type="text" aria-label={t('control.role.new.name')} data-testid="cc-role-name" value={name} placeholder={t('control.role.new.name.placeholder')} onChange={(e) => setName(e.target.value)} />
          </label>
          <div className="pane-actions">
            <button type="button" data-testid="cc-role-apply" onClick={apply}>{t('control.role.apply')}</button>
            <span className="meta">{t(role === 'condition' ? 'control.role.apply.note.condition' : 'control.role.apply.note.repeat')}</span>
          </div>
        </div>
      )}
      <dl>
        <dt>{t('control.position')}</dt>
        <dd>
          <input type="text" aria-label={t('control.position.label')} data-testid="cc-position" value={predicate} placeholder="position()=last()" onChange={(e) => setPredicate(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); setPosition(); } }} />
          <button type="button" onClick={setPosition}>{tag.positionCondition !== undefined ? (predicate.trim() ? t('control.position.update') : t('control.position.clear')) : t('control.position.set')}</button>
          <span className="meta"> {t('control.position.note', { id: tag.positionCondition !== undefined ? ` ${tag.positionCondition}` : '' })}</span>
        </dd>
      </dl>
    </Fold>
  );
}

// ---------------------------------------------------------------- the verdict

function VerdictView({ session, verdict }: { session: EditorDocument; verdict: { value?: Verdict; error?: string } | undefined }) {
  const template = session.template;
  const v = verdict?.value;
  return (
    <Fold id="control.verdict" level={3} className="control-group" testId="control-verdict" title={t('control.verdict')} meta={t(template.mode === 'xpath2' && !template.engineLoaded ? 'control.verdict.meta.needsEngine' : 'control.verdict.meta', { mode: template.mode })}>
      {verdict?.error && <p className="error">{verdict.error}</p>}
      {v && (
        <>
          <p data-testid="cc-verdict">{verdictSummary(v) || (v.role === 'none' ? t('control.verdict.unchanged') : v.role)}</p>
          {v.error && <p className="error" data-testid="cc-verdict-error">{v.error}</p>}
          {v.warnings.length > 0 && <ul className="warnings" data-testid="cc-verdict-warnings">{v.warnings.map((w) => <li key={w}>{w}</li>)}</ul>}
        </>
      )}
    </Fold>
  );
}
