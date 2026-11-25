// SPDX-License-Identifier: GPL-3.0-or-later
//
// The paragraph and section dialogs (ED-002 section 5.4): forms over `paragraph.*` and
// `section.*`; a field left as it was runs nothing.
import { isInTable, selectedRect } from 'prosemirror-tables';
import { tableAutoFitOf, type TableAutoFit } from '@docx4j/editor-model';
import { useEffect, useRef, useState, type KeyboardEvent, type ReactNode } from 'react';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { sectionAt, linkAt, documentBookmarks, paragraphsIn, paragraphTabs, twips, TAB_ALIGNMENTS, TAB_LEADERS, type SelectionFormatting, type ParagraphAttrs, type TabStopArg } from '@docx4j/editor-model';
import type { EditorSurface } from '../view/surface.js';
import { PAGE_SIZES, MARGIN_PRESETS as MARGINS_TWIPS, type PageSizePaper } from './newdoc.js';
import { useT } from './i18n.js';

/** The controls Tab moves between inside a dialog. */
export const FOCUSABLE = 'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])';

/** Whether no element has the focus (the one that had it removed): the time to give it back to the opener. */
export function focusIsNowhere(): boolean {
  const el = document.activeElement;
  return el === null || el === document.body || el === document.documentElement;
}

/**
 * A modal dialog (ED-005 section 9.5, accessibility): labelled, Escape closing it, Tab and
 * Shift+Tab kept inside it, and the focus given back where it was when it closes.
 */
export function Dialog({ title, onClose, onOk, children, okLabel }: { title: string; onClose: () => void; onOk: () => void; children: ReactNode; okLabel?: string }) {
  const t = useT();
  const form = useRef<HTMLFormElement>(null);
  // Taken as the dialog renders, before its own control takes the focus.
  const [opener] = useState(() => (typeof document !== 'undefined' ? document.activeElement : null) as HTMLElement | null);
  useEffect(() => {
    const el = form.current;
    // A field the dialog asks for takes the focus (its `autoFocus`), else OK. OK's own `autoFocus` came
    // later in the document and took it from every field (ED-005 section 12.41).
    if (el && !el.contains(document.activeElement)) el.querySelector<HTMLElement>('button[type="submit"]')?.focus();
    return () => {
      // React's StrictMode unmounts once at mount without removing the element (development only): nothing to give back then.
      if (el?.isConnected) return;
      // Only when the focus went with the element: what took it in the same commit (a dialog's field, ED-013
      // section 9) keeps it.
      if (opener?.isConnected && focusIsNowhere()) opener.focus();
    };
  }, [opener]);
  const trap = (e: KeyboardEvent<HTMLFormElement>): void => {
    if (e.key === 'Escape') { onClose(); return; }
    if (e.key !== 'Tab' || !form.current) return;
    const items = [...form.current.querySelectorAll<HTMLElement>(FOCUSABLE)];
    if (items.length === 0) return;
    const first = items[0]!;
    const last = items[items.length - 1]!;
    if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
    else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
  };
  return (
    <div className="dialog-backdrop" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <form ref={form} className="dialog" role="dialog" aria-modal="true" aria-label={title} onSubmit={(e) => { e.preventDefault(); onOk(); }}
        onKeyDown={trap}>
        <h2>{title}</h2>
        {children}
        <div className="dialog-buttons">
          <button type="button" onClick={onClose}>{t('dialog.cancel')}</button>
          <button type="submit">{okLabel ?? t('dialog.ok')}</button>
        </div>
      </form>
    </div>
  );
}

function Field({ label, value, onChange, step = 0.5, min, unit = 'pt', testId }: { label: string; value: string; onChange: (v: string) => void; step?: number; min?: number; unit?: 'pt' | 'none'; testId?: string }) {
  const t = useT();
  return (
    <label className="field">
      <span>{label}</span>
      <input type="number" step={step} min={min} value={value} placeholder={t('dialog.field.mixed')} data-testid={testId} onChange={(e) => onChange(e.target.value)} />
      {unit === 'pt' ? <span className="unit">{t('dialog.unit.pt')}</span> : <span />}
    </label>
  );
}

const show = (v: number | null | 'Mixed'): string => (v === 'Mixed' || v === null ? '' : String(v));

/**
 * A row of the Tabs section (ED-003 section 8.5, step F4): the first selected paragraph's own
 * stops (every `w:tab` it states, `clear` and `bar` included) and, after them, the stops its
 * style chain contributes, which the resolver has already merged (a stop the paragraph cleared
 * is not among them). Only the paragraph's own rows are written; removing a style's row adds a
 * `clear` at its position, as Word's dialog does.
 */
interface TabRow {
  key: number;
  position: string;
  alignment: string;
  leader: string;
  origin: 'own' | 'style';
}

function tabRowsOf(editor: EditorSurface): TabRow[] {
  const state = editor.state;
  const first = paragraphsIn(state)[0];
  if (!first) return [];
  const own = (first.node.attrs as ParagraphAttrs).tabs ?? [];
  const rows: TabRow[] = own.map((t, i) => ({ key: i, position: String(t.pos / 20), alignment: t.val, leader: t.leader ?? 'none', origin: 'own' }));
  const ownPositions = new Set(own.map((t) => t.pos));
  for (const s of paragraphTabs(first.node, state).stops) {
    if (ownPositions.has(s.pos)) continue;
    rows.push({ key: rows.length, position: String(s.pos / 20), alignment: s.val, leader: s.leader ?? 'none', origin: 'style' });
  }
  return rows;
}

const stopsOf = (rows: TabRow[]): TabStopArg[] => rows.filter((r) => r.origin === 'own' && r.position !== '').map((r) => ({ position: Number(r.position), alignment: r.alignment, leader: r.leader }));

export function ParagraphDialog({ editor, formatting, onClose }: { editor: EditorSurface; formatting: SelectionFormatting; onClose: () => void }) {
  const t = useT();
  const p = formatting.paragraph;
  const initial = { leftIndent: show(p.leftIndent), rightIndent: show(p.rightIndent), firstLineIndent: show(p.firstLineIndent), spaceBefore: show(p.spaceBefore), spaceAfter: show(p.spaceAfter), lineSpacing: show(p.lineSpacing) };
  const [v, setV] = useState(initial);
  const [initialRows] = useState(() => tabRowsOf(editor));
  const [rows, setRows] = useState(initialRows);
  const ok = (): void => {
    for (const key of Object.keys(initial) as (keyof typeof initial)[]) {
      if (v[key] !== initial[key] && v[key] !== '') editor.run(`paragraph.${key}`, { value: Number(v[key]) });
    }
    const stops = stopsOf(rows);
    if (JSON.stringify(stops) !== JSON.stringify(stopsOf(initialRows))) editor.run('paragraph.tabStops', { stops });
    onClose();
  };
  const set = (key: keyof typeof initial) => (value: string) => setV({ ...v, [key]: value });
  const setRow = (key: number, patch: Partial<TabRow>): void => setRows(rows.map((r) => (r.key === key ? { ...r, ...patch } : r)));
  const remove = (row: TabRow): void => {
    // A style's stop is cancelled by an own `clear` at its position; an own stop simply goes.
    if (row.origin === 'style') setRows(rows.map((r) => (r.key === row.key ? { ...r, origin: 'own', alignment: 'clear', leader: 'none' } : r)));
    else setRows(rows.filter((r) => r.key !== row.key));
  };
  const add = (): void => setRows([...rows, { key: rows.reduce((m, r) => Math.max(m, r.key), -1) + 1, position: '', alignment: 'left', leader: 'none', origin: 'own' }]);
  const tabsEnabled = paragraphsIn(editor.state).length > 0;
  return (
    <Dialog title={t('paragraphDialog.title')} onClose={onClose} onOk={ok}>
      <Field label={t('paragraphDialog.leftIndent')} value={v.leftIndent} onChange={set('leftIndent')} />
      <Field label={t('paragraphDialog.rightIndent')} value={v.rightIndent} onChange={set('rightIndent')} />
      <Field label={t('paragraphDialog.firstLine')} value={v.firstLineIndent} onChange={set('firstLineIndent')} />
      <Field label={t('paragraphDialog.spaceBefore')} value={v.spaceBefore} onChange={set('spaceBefore')} />
      <Field label={t('paragraphDialog.spaceAfter')} value={v.spaceAfter} onChange={set('spaceAfter')} />
      <Field label={t('paragraphDialog.lineSpacing')} value={v.lineSpacing} onChange={set('lineSpacing')} />
      <fieldset className="tab-stops" data-testid="tab-stops">
        <legend>{t('paragraphDialog.tabs')}</legend>
        {rows.length === 0 && <p className="meta">{t('paragraphDialog.tabs.none')}</p>}
        {rows.map((r) => (
          <div className="tab-stop" key={r.key} data-testid="tab-stop" data-origin={r.origin}>
            <input type="number" step={0.5} min={0} aria-label={t('paragraphDialog.tab.position')} value={r.position} placeholder={t('dialog.unit.pt')} disabled={r.origin === 'style'} onChange={(e) => setRow(r.key, { position: e.target.value })} />
            <span className="unit">{t('dialog.unit.pt')}</span>
            <select aria-label={t('paragraphDialog.tab.alignment')} value={r.alignment} disabled={r.origin === 'style'} onChange={(e) => setRow(r.key, { alignment: e.target.value })}>
              {TAB_ALIGNMENTS.map((a) => <option key={a} value={a}>{t(`tabStop.alignment.${a}`)}</option>)}
            </select>
            <select aria-label={t('paragraphDialog.tab.leader')} value={r.leader} disabled={r.origin === 'style'} onChange={(e) => setRow(r.key, { leader: e.target.value })}>
              {TAB_LEADERS.map((l) => <option key={l} value={l}>{t(`tabStop.leader.${l}`)}</option>)}
            </select>
            <span className="meta">{r.origin === 'style' ? t('paragraphDialog.tab.fromStyle') : t('paragraphDialog.tab.own')}</span>
            <button type="button" aria-label={r.origin === 'style' ? t('paragraphDialog.tab.clear') : t('paragraphDialog.tab.remove')} onClick={() => remove(r)}>×</button>
          </div>
        ))}
        <button type="button" data-testid="tab-add" disabled={!tabsEnabled} onClick={add}>{t('paragraphDialog.tab.add')}</button>
      </fieldset>
    </Dialog>
  );
}

const PT = 20;

/** Word's paper sizes in points (A4 11906 x 16838 twips, Letter 12240 x 15840), in portrait. */
const PAPER = { a4: { w: 11906 / PT, h: 16838 / PT }, letter: { w: 12240 / PT, h: 15840 / PT } } as const;
/** Word's margin presets in points (`newdoc.ts` has them in twips, for new documents); the header and footer distances are left alone. */
const MARGIN_PRESETS = { normal: MARGINS_TWIPS.normal / PT, narrow: MARGINS_TWIPS.narrow / PT } as const;
const MARGIN_SIDES = ['top', 'bottom', 'left', 'right'] as const;

/**
 * The Section dialog, in Word's groups: Page (orientation and size, with A4 and Letter), Margins (with Word's Normal
 * and Narrow presets), Layout (where the section starts, its column count) and the section break. A field left as it
 * was runs nothing; a preset fills the fields and OK runs the commands once.
 */
export function SectionDialog({ editor, onClose }: { editor: EditorSurface; onClose: () => void }) {
  const t = useT();
  const target = sectionAt(editor.state);
  const s = (target.sectPr ?? {}) as wml.SectPr;
  const g = { w: (s.pgSz?.w ?? 11906) / PT, h: (s.pgSz?.h ?? 16838) / PT, top: (s.pgMar?.top ?? 1440) / PT, right: (s.pgMar?.right ?? 1440) / PT, bottom: (s.pgMar?.bottom ?? 1440) / PT, left: (s.pgMar?.left ?? 1440) / PT, header: (s.pgMar?.header ?? 708) / PT, footer: (s.pgMar?.footer ?? 708) / PT, gutter: (s.pgMar?.gutter ?? 0) / PT };
  const initial = {
    width: String(g.w), height: String(g.h), orientation: s.pgSz?.orient === 'landscape' ? 'landscape' : 'portrait',
    top: String(g.top), right: String(g.right), bottom: String(g.bottom), left: String(g.left), header: String(g.header), footer: String(g.footer), gutter: String(g.gutter),
    type: s.type?.val ?? 'nextPage', columns: String(s.cols?.num ?? 1),
  };
  const [v, setV] = useState(initial);
  const set = (key: keyof typeof initial) => (value: string) => setV({ ...v, [key]: value });
  // A paper size fills width and height for the orientation chosen; a margin preset the four sides.
  const paper = (name: keyof typeof PAPER): void => {
    const { w, h } = PAPER[name];
    setV(v.orientation === 'landscape' ? { ...v, width: String(h), height: String(w) } : { ...v, width: String(w), height: String(h) });
  };
  const isPaper = (name: keyof typeof PAPER): boolean => {
    const { w, h } = PAPER[name];
    const [pw, ph] = v.orientation === 'landscape' ? [h, w] : [w, h];
    return Math.abs(Number(v.width) - pw) < 0.5 && Math.abs(Number(v.height) - ph) < 0.5;
  };
  const marginPreset = (name: keyof typeof MARGIN_PRESETS): void => {
    const next = { ...v };
    for (const side of MARGIN_SIDES) next[side] = String(MARGIN_PRESETS[name]);
    setV(next);
  };
  const isMarginPreset = (name: keyof typeof MARGIN_PRESETS): boolean => MARGIN_SIDES.every((side) => Number(v[side]) === MARGIN_PRESETS[name]);
  const ok = (): void => {
    if (v.orientation !== initial.orientation) editor.run('section.orientation', { value: v.orientation });
    if (v.width !== initial.width || v.height !== initial.height) editor.run('section.pageSize', { width: Number(v.width), height: Number(v.height) });
    const margins: Record<string, number> = {};
    for (const k of ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter'] as const) if (v[k] !== initial[k]) margins[k] = Number(v[k]);
    if (Object.keys(margins).length > 0) editor.run('section.margins', margins);
    if (v.type !== initial.type) editor.run('section.type', { value: v.type });
    if (v.columns !== initial.columns) editor.run('section.columns', { count: Math.max(1, Math.round(Number(v.columns) || 1)) });
    onClose();
  };
  return (
    <Dialog title={target.kind === 'final' ? t('sectionDialog.title.final') : t('sectionDialog.title')} onClose={onClose} onOk={ok}>
      <fieldset data-testid="section-page">
        <legend>{t('sectionDialog.page')}</legend>
        <label className="field"><span>{t('sectionDialog.orientation')}</span>
          <select value={v.orientation} onChange={(e) => set('orientation')(e.target.value)}><option value="portrait">{t('sectionDialog.portrait')}</option><option value="landscape">{t('sectionDialog.landscape')}</option></select>
        </label>
        <Field label={t('sectionDialog.pageWidth')} value={v.width} onChange={set('width')} testId="section-width" />
        <Field label={t('sectionDialog.pageHeight')} value={v.height} onChange={set('height')} testId="section-height" />
        <div className="field presets">
          <span>{t('sectionDialog.paper')}</span>
          <button type="button" aria-pressed={isPaper('a4')} onClick={() => paper('a4')}>{t('sectionDialog.paper.a4')}</button>
          <button type="button" aria-pressed={isPaper('letter')} onClick={() => paper('letter')}>{t('sectionDialog.paper.letter')}</button>
        </div>
      </fieldset>
      <fieldset data-testid="section-margins">
        <legend>{t('sectionDialog.margins')}</legend>
        <Field label={t('sectionDialog.marginTop')} value={v.top} onChange={set('top')} testId="section-margin-top" />
        <Field label={t('sectionDialog.marginBottom')} value={v.bottom} onChange={set('bottom')} />
        <Field label={t('sectionDialog.marginLeft')} value={v.left} onChange={set('left')} testId="section-margin-left" />
        <Field label={t('sectionDialog.marginRight')} value={v.right} onChange={set('right')} />
        <Field label={t('sectionDialog.header')} value={v.header} onChange={set('header')} />
        <Field label={t('sectionDialog.footer')} value={v.footer} onChange={set('footer')} />
        <Field label={t('sectionDialog.gutter')} value={v.gutter} onChange={set('gutter')} />
        <div className="field presets">
          <span>{t('sectionDialog.margins.preset')}</span>
          <button type="button" aria-pressed={isMarginPreset('normal')} onClick={() => marginPreset('normal')}>{t('sectionDialog.margins.normal')}</button>
          <button type="button" aria-pressed={isMarginPreset('narrow')} onClick={() => marginPreset('narrow')}>{t('sectionDialog.margins.narrow')}</button>
        </div>
      </fieldset>
      <fieldset data-testid="section-layout">
        <legend>{t('sectionDialog.layout')}</legend>
        <label className="field"><span>{t('sectionDialog.start')}</span>
          <select value={v.type} onChange={(e) => set('type')(e.target.value)}>
            {(['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn'] as const).map((type) => <option key={type} value={type}>{t(`sectionDialog.start.${type}`)}</option>)}
          </select>
        </label>
        <Field label={t('sectionDialog.columns')} value={v.columns} onChange={set('columns')} step={1} min={1} unit="none" testId="section-columns" />
      </fieldset>
      <fieldset data-testid="section-break">
        <legend>{t('sectionDialog.break')}</legend>
        <div className="field presets">
          <span />
          <button type="button" onClick={() => { editor.run('section.insertBreak', {}); onClose(); }} disabled={target.kind === 'paragraph' && target.node === editor.state.selection.$from.parent}>{t('sectionDialog.break.insert')}</button>
          <button type="button" onClick={() => { editor.run('section.removeBreak'); onClose(); }}>{t('sectionDialog.break.remove')}</button>
        </div>
      </fieldset>
    </Dialog>
  );
}

/** Table properties (ED-005 section 12.78): Word's AutoFit behaviour, fixed column widths, AutoFit to window or AutoFit to contents. */
export function TableDialog({ editor, onClose, onFit }: { editor: EditorSurface; onClose: () => void; onFit?: () => void }) {
  const t = useT();
  const rect = isInTable(editor.state) ? selectedRect(editor.state) : undefined;
  const initial: TableAutoFit = rect ? tableAutoFitOf(rect.table.attrs['tblPr']) : 'fixed';
  const [autoFit, setAutoFit] = useState<TableAutoFit>(initial);
  const ok = (): void => {
    if (autoFit !== initial) editor.run('table.properties', { autoFit });
    onClose();
  };
  return (
    <Dialog title={t('tableDialog.title')} onClose={onClose} onOk={ok}>
      <label className="field"><span>{t('tableDialog.autoFit')}</span>
        <select data-testid="table-autofit" value={autoFit} onChange={(e) => setAutoFit(e.target.value as TableAutoFit)}>
          {(['fixed', 'window', 'contents'] as const).map((mode) => <option key={mode} value={mode}>{t(`tableDialog.autoFit.${mode}`)}</option>)}
        </select>
      </label>
      <p className="meta">{t(`tableDialog.autoFit.${autoFit}.note`)}</p>
      {onFit && (
        <>
          <div className="field presets"><span>{t('tableDialog.fit')}</span>
            <button type="button" data-testid="table-fit-dialog" onClick={() => { onClose(); onFit(); }}>{t('tableDialog.fit')}</button><span />
          </div>
          <p className="meta">{t('tableDialog.fit.note')}</p>
        </>
      )}
    </Dialog>
  );
}

/**
 * Word's Insert Table dialog (ED-005 section 12.83): the size, and the AutoFit behaviour with a fixed column width
 * (Auto: the text width shared) over `table.insert`.
 */
export function InsertTableDialog({ editor, onClose }: { editor: EditorSurface; onClose: () => void }) {
  const t = useT();
  const [columns, setColumns] = useState('5');
  const [rows, setRows] = useState('2');
  const [autoFit, setAutoFit] = useState<TableAutoFit>('fixed');
  const [width, setWidth] = useState('');
  const ok = (): void => {
    const c = Math.max(1, Math.min(63, Math.round(Number(columns) || 0)));
    const r = Math.max(1, Math.round(Number(rows) || 0));
    const w = Number(width);
    editor.run('table.insert', { rows: r, columns: c, autoFit, ...(autoFit === 'fixed' && width.trim() !== '' && w > 0 ? { columnWidth: w } : {}) });
    onClose();
  };
  return (
    <Dialog title={t('insertTableDialog.title')} onClose={onClose} onOk={ok}>
      <fieldset>
        <legend>{t('insertTableDialog.size')}</legend>
        <label className="field"><span>{t('insertTableDialog.columns')}</span>
          <input type="number" min={1} max={63} step={1} data-testid="insert-table-columns" value={columns} autoFocus onFocus={(e) => e.target.select()} onChange={(e) => setColumns(e.target.value)} />
        </label>
        <label className="field"><span>{t('insertTableDialog.rows')}</span>
          <input type="number" min={1} step={1} data-testid="insert-table-rows" value={rows} onChange={(e) => setRows(e.target.value)} />
        </label>
      </fieldset>
      <fieldset>
        <legend>{t('insertTableDialog.autoFit')}</legend>
        <label className="check"><input type="radio" name="insert-table-autofit" data-testid="insert-table-fixed" checked={autoFit === 'fixed'} onChange={() => setAutoFit('fixed')} /> {t('tableDialog.autoFit.fixed')}
          <input type="number" min={1} step={0.5} className="inline" aria-label={t('insertTableDialog.width')} data-testid="insert-table-width" placeholder={t('insertTableDialog.width.auto')} value={width} disabled={autoFit !== 'fixed'} onChange={(e) => setWidth(e.target.value)} />
          <span className="unit">{t('dialog.unit.pt')}</span>
        </label>
        <label className="check"><input type="radio" name="insert-table-autofit" data-testid="insert-table-contents" checked={autoFit === 'contents'} onChange={() => setAutoFit('contents')} /> {t('tableDialog.autoFit.contents')}</label>
        <label className="check"><input type="radio" name="insert-table-autofit" data-testid="insert-table-window" checked={autoFit === 'window'} onChange={() => setAutoFit('window')} /> {t('tableDialog.autoFit.window')}</label>
      </fieldset>
    </Dialog>
  );
}

/** The page size for new documents, asked the first time New is used; the choice is kept as an option. */
export function PageSizeDialog({ onChoose, onClose }: { onChoose: (size: PageSizePaper) => void; onClose: () => void }) {
  const t = useT();
  const [size, setSize] = useState<PageSizePaper>('A4');
  return (
    <Dialog title={t('pageSizeDialog.title')} onClose={onClose} onOk={() => onChoose(size)}>
      <p className="meta">{t('pageSizeDialog.note')}</p>
      <label className="field page-size">
        <span>{t('pageSizeDialog.size')}</span>
        <select aria-label={t('pageSizeDialog.size')} data-testid="page-size" value={size} autoFocus onChange={(e) => setSize(e.target.value as PageSizePaper)}>
          {PAGE_SIZES.map((s) => <option key={s.id} value={s.id}>{t(`pageSize.${s.id}`)}</option>)}
        </select>
      </label>
    </Dialog>
  );
}


/**
 * The hyperlink dialog (ED-003 section 5.5): Word's fields over `range.hyperlink`. Over the
 * selection or the link at the caret (its address and text shown); at a bare caret the text is
 * asked for and inserted. An address is external; a bookmark is a place in this document; both
 * at once are allowed (`address#bookmark`). Remove keeps the text.
 */
export function LinkDialog({ editor, onClose }: { editor: EditorSurface; onClose: () => void }) {
  const t = useT();
  const state = editor.state;
  const existing = linkAt(state);
  const selected = state.selection.empty ? '' : state.doc.textBetween(state.selection.from, state.selection.to, '\t');
  const bookmarks = documentBookmarks(state.doc, editor.session);
  const [address, setAddress] = useState(existing?.address ?? '');
  const [bookmark, setBookmark] = useState(existing?.anchor ?? '');
  const [text, setText] = useState(existing?.text ?? selected);
  const bare = !existing && selected === '';
  const apply = (): void => {
    if (address.trim() === '' && bookmark === '') return;
    void editor.runAsync('range.hyperlink', { address: address.trim() || undefined, subAddress: bookmark || undefined, text: text !== (existing?.text ?? selected) || bare ? text : undefined });
    onClose();
  };
  const remove = (): void => { void editor.runAsync('range.hyperlink', {}); onClose(); };
  return (
    <Dialog title={t('linkDialog.title')} onClose={onClose} onOk={apply}>
      <label className="field"><span>{t('linkDialog.address')}</span><input type="text" data-testid="link-address" value={address} placeholder="https://…" autoFocus onChange={(e) => setAddress(e.target.value)} /></label>
      <label className="field"><span>{t('linkDialog.bookmark')}</span>
        <select data-testid="link-bookmark" value={bookmark} onChange={(e) => setBookmark(e.target.value)}>
          <option value="">{t('linkDialog.bookmark.none')}</option>
          {bookmarks.map((b) => <option key={b.name} value={b.name}>{b.name}</option>)}
        </select>
      </label>
      <label className="field"><span>{t('linkDialog.text')}</span><input type="text" data-testid="link-text" value={text} onChange={(e) => setText(e.target.value)} placeholder={bare ? t('linkDialog.text.placeholder') : undefined} /></label>
      {existing && <button type="button" data-testid="link-remove" onClick={remove}>{t('linkDialog.remove')}</button>}
    </Dialog>
  );
}

/**
 * Turning tracking on with no name set (ED-005 sections 5.4 and 6.6): tracked changes carry their
 * author's name, so it is asked for first, the guest name offered. OK keeps it as the "Your name"
 * option and turns tracking on; Cancel leaves it off.
 */
/**
 * Reload from disk with unsaved changes (ED-011 section 9.9; Jason, 2026-11-09): what is lost said before it goes,
 * since the reload deletes the session, so neither Undo nor recovery brings the changes back.
 */
export function DiscardReloadDialog({ name, onConfirm, onClose }: { name: string; onConfirm: () => void; onClose: () => void }) {
  const t = useT();
  return (
    <Dialog title={t('discardReloadDialog.title')} okLabel={t('discardReloadDialog.ok')} onClose={onClose} onOk={onConfirm}>
      <p data-testid="discard-reload-body">{t('discardReloadDialog.body', { name })}</p>
    </Dialog>
  );
}

/** The question before Close discards unsaved changes, when the option asks for it. */
export function DiscardCloseDialog({ name, onConfirm, onClose }: { name: string; onConfirm: () => void; onClose: () => void }) {
  const t = useT();
  return (
    <Dialog title={t('closeDialog.title')} okLabel={t('closeDialog.ok')} onClose={onClose} onOk={onConfirm}>
      <p data-testid="close-ask-body">{t('closeDialog.body', { name })}</p>
    </Dialog>
  );
}

/** Send a copy's second tap (ED-013 section 3.8): the copy is made, but the first tap's activation was spent on making it. */
export function SendReadyDialog({ name, onSend, onClose }: { name: string; onSend: () => void; onClose: () => void }) {
  const t = useT();
  return (
    <Dialog title={t('send.ready.title')} okLabel={t('send.ready.ok')} onClose={onClose} onOk={onSend}>
      <p data-testid="send-ready-body">{t('send.ready.body', { name })}</p>
    </Dialog>
  );
}

export function TrackingNameDialog({ initial, onChoose, onClose, purpose = 'tracking' }: { initial: string; onChoose: (name: string) => void; onClose: () => void; purpose?: 'tracking' | 'comments' }) {
  const t = useT();
  const [name, setName] = useState(initial);
  return (
    <Dialog title={purpose === 'comments' ? t('comments.name.title') : t('tracking.name.title')} onClose={onClose} onOk={() => { if (name.trim() !== '') onChoose(name.trim()); }}>
      <p className="meta">{purpose === 'comments' ? t('comments.name.prompt') : t('tracking.name.prompt')}</p>
      <label className="field">
        <span>{t('options.name')}</span>
        <input type="text" data-testid="tracking-name" value={name} autoFocus onChange={(e) => setName(e.target.value)} />
      </label>
    </Dialog>
  );
}
