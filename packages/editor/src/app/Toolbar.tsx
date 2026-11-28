// SPDX-License-Identifier: GPL-3.0-or-later
//
// The formatting toolbar (ED-001 section 3.4; ED-002 section 5.4): every control is a command
// from the registry with its arguments; pressed and enabled states come from the effective
// formatting reading and the specs, nothing else. Laid out as Word's and ONLYOFFICE's Home ribbon:
// groups of two rows over a caption, Undo, Font (the run's properties), Paragraph and Styles, then
// the insert items, the content controls and the table's tools in groups of their own.
import { useRef, useEffect, useMemo, useState, type CSSProperties, type ReactNode } from 'react';
import { FONT_SIZES, type StyleEntry, type SelectionFormatting, type SdtKind } from '@docx4j/editor-model';
import type { EditorSurface } from '../view/surface.js';
import { useToolbarReading } from './state.js';
import { resolvedFontName } from './fonts.js';
import { stylePreview, highlightCss } from '../view/css.js';
import { ParagraphDialog, SectionDialog, LinkDialog, TableDialog, InsertTableDialog } from './Dialogs.js';
import type { ToolbarPreset } from './perspectives.js';
import { useT } from './i18n.js';
import { placeMenuPane } from './menus.js';

/** The kinds the insert menu offers (ED-003 section 4.3), in Word's order; each labelled by `control.kind.<kind>`. */
const INSERT_KINDS: readonly SdtKind[] = ['RichText', 'PlainText', 'CheckBox', 'DatePicker', 'DropDownList', 'ComboBox', 'Picture', 'RepeatingSection', 'Group'];

interface ToolButtonProps {
  editor: EditorSurface;
  command: string;
  args?: unknown;
  /** A command tried first; `command` runs when it does not apply (the indent buttons on a list item change its level). */
  first?: string;
  label: string;
  pressed?: boolean;
  disabled?: boolean;
  children: ReactNode;
  after?: () => void;
}

/** Insert picture (ED-005 section 12.76): a file chosen, read as base64, inserted at the caret through `range.insertInlinePicture`. */
function PictureButton({ editor, label, disabled }: { editor: EditorSurface; label: string; disabled: boolean }) {
  const input = useRef<HTMLInputElement | null>(null);
  const onFile = async (file: File | undefined): Promise<void> => {
    if (!file) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    let binary = '';
    for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    await editor.runAsync('range.insertInlinePicture', { base64: btoa(binary), name: file.name.replace(/\.[^.]+$/, '') });
    editor.focus();
  };
  return (
    <>
      <button type="button" className="tool" aria-label={label} title={label} data-testid="insert-picture" disabled={disabled} onMouseDown={(e) => e.preventDefault()} onClick={() => input.current?.click()}>🖼</button>
      <input ref={input} type="file" accept="image/png,image/jpeg,image/gif,image/bmp,image/tiff,image/svg+xml" hidden data-testid="insert-picture-file" onChange={(e) => { const f = e.target.files?.[0]; e.target.value = ''; void onFile(f); }} />
    </>
  );
}

export function ToolButton({ editor, command, args, first, label, pressed, disabled, children, after }: ToolButtonProps) {
  return (
    <button
      type="button"
      className={`tool${pressed ? ' pressed' : ''}`}
      aria-label={label}
      title={label}
      aria-pressed={pressed === undefined ? undefined : pressed}
      disabled={disabled}
      data-command={command}
      onMouseDown={(e) => e.preventDefault()}
      onClick={() => { if (!(first && editor.run(first))) editor.run(command, args); after?.(); editor.focus(); }}
    >
      {children}
    </button>
  );
}

/** A ribbon group: its controls in one or two rows over its caption; `launcher`, Word's dialog launcher, at the caption's end. */
function Group({ label, caption = true, launcher, testId, children }: { label: string; caption?: boolean; launcher?: ReactNode; testId?: string; children: ReactNode }) {
  return (
    <div className="ribbon-group" role="group" aria-label={label} data-testid={testId}>
      <div className="ribbon-body">{children}</div>
      <div className="ribbon-caption">{caption && <span aria-hidden="true">{label}</span>}{launcher}</div>
    </div>
  );
}

const Row = ({ children }: { children: ReactNode }) => <div className="ribbon-row">{children}</div>;

/** Word's alignment icons: four lines, long and short, against the edge they align to; justified, all long. */
function AlignIcon({ align }: { align: 'left' | 'centre' | 'right' | 'justify' }) {
  return (
    <svg className="icon" viewBox="0 0 16 16" width="16" height="16" fill="currentColor" aria-hidden="true" focusable="false">
      {[14, 9, 14, 9].map((long, i) => {
        const width = align === 'justify' ? 14 : long;
        const x = align === 'right' ? 15 - width : align === 'centre' ? (16 - width) / 2 : 1;
        return <rect key={i} x={x} y={2 + i * 3.5} width={width} height={1.6} rx={0.4} />;
      })}
    </svg>
  );
}

/** A glyph over a colour bar (the highlight's, the font colour's), as Word draws those two buttons. */
function SwatchIcon({ glyph, colour }: { glyph: ReactNode; colour: string | null }) {
  return (
    <span className="swatch-icon" aria-hidden="true">
      <span className="glyph">{glyph}</span>
      <span className={`bar${colour === null ? ' none' : ''}`} style={colour === null ? undefined : { background: colour }} />
    </span>
  );
}

/** A colour a swatch menu offers: its id (a name's key, or a highlight name), its CSS colour and its name for people. */
interface Swatch { id: string; colour: string; name: string }

// Word's ten standard colours, by Word's names and values: the second row of Row shading and of the font colour.
const STANDARD_COLOURS: readonly (readonly [string, string])[] = [
  ['darkRed', '#C00000'], ['red', '#FF0000'], ['orange', '#FFC000'], ['yellow', '#FFFF00'], ['lightGreen', '#92D050'],
  ['green', '#00B050'], ['lightBlue', '#00B0F0'], ['blue', '#0070C0'], ['darkBlue', '#002060'], ['purple', '#7030A0'],
];
// The Row shading menu's colours (ED-005 section 12.128): white and three greys, the shades rows are most often given,
// then the standard colours.
const ROW_SHADES: readonly (readonly [string, string])[] = [
  ['white', '#FFFFFF'], ['grey5', '#F2F2F2'], ['grey15', '#D9D9D9'], ['grey25', '#BFBFBF'], ...STANDARD_COLOURS,
];

/**
 * A menu of colour swatches (ED-005 sections 12.128 and 12.131): Row shading, the highlighter's arrow and the font colour.
 * A `<details>` whose summary is a tool. Its pane holds the swatches, an entry for no colour (first, as Word puts the font
 * colour's Automatic, or after them), and, when `more` is given, the browser's picker for any other colour. Picking closes
 * it and gives the focus back to the text; so do Escape and a press outside. The selection stays where it was while the
 * menu is open: every press in it keeps the focus off the menu.
 */
function SwatchMenu({ editor, testId, label, command, summary, summaryClass, swatches, columns, isCurrent, onPick, none, more }: {
  editor: EditorSurface;
  testId: string;
  label: string;
  command: string;
  summary: ReactNode;
  summaryClass?: string;
  swatches: readonly Swatch[];
  columns: number;
  isCurrent: (swatch: Swatch) => boolean;
  /** A swatch, or null for the no-colour entry. */
  onPick: (swatch: Swatch | null) => void;
  none: { label: string; testId: string; first?: boolean };
  /** The browser's picker: the colour it starts at when the menu opens, and what a chosen colour does. */
  more?: { label: string; testId: string; initial: () => string; onPick: (hex: string) => void };
}) {
  const menu = useRef<HTMLDetailsElement>(null);
  const picker = useRef<HTMLInputElement>(null);
  const close = (): void => { if (menu.current) menu.current.open = false; };
  const done = (): void => { close(); editor.focus(); };
  const moreRef = useRef(more);
  moreRef.current = more;
  const doneRef = useRef(done);
  doneRef.current = done;
  // The picker's colour once it is chosen (the native change event), not each colour the picker passes through on the way.
  useEffect(() => {
    const input = picker.current;
    if (!input) return;
    const onChange = (): void => { moreRef.current?.onPick(input.value.toUpperCase()); doneRef.current(); };
    input.addEventListener('change', onChange);
    return () => input.removeEventListener('change', onChange);
  }, []);
  // Always listening, the menu's own state asked at the event: `toggle` is queued, so a key straight after the click that
  // opens the menu can come before any state the toggle sets.
  const editorRef = useRef(editor);
  editorRef.current = editor;
  useEffect(() => {
    const onDown = (e: PointerEvent): void => { if (menu.current?.open && !menu.current.contains(e.target as Node)) close(); };
    const onKey = (e: KeyboardEvent): void => {
      if (e.key !== 'Escape' || !menu.current?.open) return;
      e.preventDefault();
      close();
      editorRef.current.focus();
    };
    document.addEventListener('pointerdown', onDown, true);
    document.addEventListener('keydown', onKey, true);
    return () => { document.removeEventListener('pointerdown', onDown, true); document.removeEventListener('keydown', onKey, true); };
  }, []);
  const keep = (e: { preventDefault(): void }): void => e.preventDefault();
  const noneButton = (
    <button type="button" className="none" data-testid={none.testId} onMouseDown={keep} onClick={() => { onPick(null); done(); }}>{none.label}</button>
  );
  return (
    <details ref={menu} className="tool-menu swatch-menu" data-testid={testId}
      onToggle={(e) => {
        // The picker starts at the current colour: uncontrolled, since React would put a controlled one back mid-pick.
        if (e.currentTarget.open && picker.current && moreRef.current) picker.current.value = moreRef.current.initial().toLowerCase();
        placeMenuPane(e.currentTarget);
      }}>
      <summary className={`tool${summaryClass ? ` ${summaryClass}` : ''}`} aria-label={label} title={label} data-command={command} onMouseDown={keep}>
        {summary}
      </summary>
      <div className="menu-pane" role="group" aria-label={label}>
        {none.first && noneButton}
        <div className="swatches" style={{ '--swatch-columns': columns } as CSSProperties}>
          {swatches.map((s) => (
            <button key={s.id} type="button" className={`swatch${isCurrent(s) ? ' pressed' : ''}`} style={{ background: s.colour }} data-colour={s.colour}
              aria-label={s.name} title={s.name} aria-pressed={isCurrent(s)}
              onMouseDown={keep} onClick={() => { onPick(s); done(); }} />
          ))}
        </div>
        {!none.first && noneButton}
        {more && (
          <label className="more">{more.label}
            <input ref={picker} type="color" data-testid={more.testId} defaultValue="#ffffff" />
          </label>
        )}
      </div>
    </details>
  );
}

/** Row shading (ED-005 section 12.128): the selected rows' cells' colour, No colour, and the picker for any other. */
function RowShadingMenu({ editor, current }: { editor: EditorSurface; current: string | null | 'Mixed' | undefined }) {
  const t = useT();
  const bar = current && current !== 'Mixed' ? current : null;
  const swatches = useMemo(() => ROW_SHADES.map(([id, colour]) => ({ id, colour, name: t(`toolbar.table.shading.${id}`) })), [t]);
  return (
    <SwatchMenu editor={editor} testId="row-shading" label={t('toolbar.table.shading')} command="row.shading"
      summary={<><SwatchIcon glyph="▤" colour={bar} /><span className="menu-arrow" aria-hidden="true">▾</span></>}
      swatches={swatches} columns={7} isCurrent={(s) => current === s.colour}
      onPick={(s) => { editor.run('row.shading', { color: s ? s.colour : null }); }}
      none={{ label: t('toolbar.table.shading.none'), testId: 'row-shading-none' }}
      more={{ label: t('toolbar.table.shading.more'), testId: 'row-shading-more', initial: () => bar ?? '#FFFFFF', onPick: (color) => { editor.run('row.shading', { color }); } }} />
  );
}

// Word's highlight palette, in the order of Word's menu, five to a row.
const HIGHLIGHTS = ['yellow', 'green', 'cyan', 'magenta', 'blue', 'red', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray', 'black'];

// The font colour menu's colours (ED-005 section 12.131): black, two greys and white, then the standard colours, each with
// the key of its name.
const FONT_COLOURS: readonly (readonly [string, string, string])[] = [
  ['black', '#000000', 'highlight.black'], ['darkGray', '#808080', 'highlight.darkGray'], ['lightGray', '#C0C0C0', 'highlight.lightGray'],
  ['white', '#FFFFFF', 'toolbar.table.shading.white'],
  ...STANDARD_COLOURS.map(([id, colour]) => [id, colour, `toolbar.table.shading.${id}`] as const),
];

// The highlighter's pen (ED-005 section 12.52 item 8): the colour the button applies, yellow until a colour is
// picked, '' once "No highlight" is, kept per browser as Word keeps its last colour.
const PEN_KEY = 'docx4j-editor.highlightPen';
function loadPen(): string {
  try {
    const v = localStorage.getItem(PEN_KEY);
    if (v === '' || (v && HIGHLIGHTS.includes(v))) return v;
  } catch { /* storage may be unavailable */ }
  return 'yellow';
}
function savePen(pen: string): void {
  try { localStorage.setItem(PEN_KEY, pen); } catch { /* storage may be unavailable */ }
}

export interface ToolbarProps {
  editor: EditorSurface;
  styles: StyleEntry[];
  fonts: string[];
  /** Everything (the default), or the paragraph and character basics (the Template perspective). */
  preset?: Exclude<ToolbarPreset, 'none'>;
  /** The content control group (ED-003 section 4.9): insert kinds, wrap rows and cells, unwrap, delete, the Control tab. */
  controls?: boolean;
  /** Bring the Control tab to the front (properties and binding live there). */
  onControl?: () => void;
  /** Fit the table's columns to its content (ED-007): the application measures and reports. */
  onFitTable?: () => void;
  /**
   * Hands the application the opening of the link dialog while the toolbar is shown, and null when it goes: Ctrl+K is the
   * surfaces' shortcut (`Mod-k`), so that in VS Code the key gate keeps it from VS Code's chords (ED-014 section 4.4).
   */
  bindLink?: (open: (() => void) | null) => void;
}

export function Toolbar({ editor, styles, fonts, preset = 'full', controls = false, onControl, onFitTable, bindLink }: ToolbarProps) {
  const t = useT();
  // One memoised reading per transaction (ED-003 section 8.2): the toolbar re-renders only when it changes.
  const reading = useToolbarReading(editor);
  const fmt: SelectionFormatting = reading.fmt;
  const [dialog, setDialog] = useState<'paragraph' | 'section' | 'link' | 'table' | 'insertTable' | null>(null);
  // Ctrl+K opens the hyperlink dialog from any editing surface: the surfaces' shortcut calls what is bound here.
  useEffect(() => {
    if (!bindLink) return;
    bindLink(() => setDialog('link'));
    return () => bindLink(null);
  }, [bindLink]);
  const font = fmt.font;
  const para = fmt.paragraph;
  const value = <T,>(v: T | 'Mixed', fallback: T): T => (v === 'Mixed' ? fallback : v);
  const fontName = font.name === 'Mixed' ? '' : font.name === '' ? resolvedFontName(editor) : font.name;
  const fontSize = value(font.size, 0);
  const styleId = value(para.style, null) ?? '';
  const enabled = (name: string): boolean => reading.enabled[name] ?? true;
  const paragraphStyles = styles.filter((s) => s.type === 'paragraph');
  const fontList = fontName && !fonts.includes(fontName) ? [fontName, ...fonts] : fonts;
  const full = preset === 'full';
  const inControl = reading.inControl;
  // Bullets or Numbering pressed from the caret paragraph's kind (ED-003 section 5.7).
  const listKind = reading.listKind;

  // The caret's style as the list names it: a paragraph with no style has the default one, which the list offers as ''.
  const current = paragraphStyles.find((s) => s.id === styleId) ?? (styleId === '' ? paragraphStyles.find((s) => s.isDefault) : undefined);
  const pickerValue = current ? (current.isDefault ? '' : current.id) : styleId;
  // Each paragraph style's entry in the list shown in the style's look (font, size, bold, colour), as Word's pane shows them.
  const previews = useMemo(() => new Map(paragraphStyles.map((s) => [s.id, stylePreview(editor.session, s.id)])), [editor.session, styles]);
  // The highlighter: the pen's colour on the button, pressed when the selection is highlighted in it (a mark may hold
  // the name or its hex), a click applying the pen, or clearing the selection's highlight when it already has the pen's colour, as Word's does.
  const [pen, setPen] = useState(loadPen);
  const penCss = pen === '' ? null : highlightCss(pen);
  const highlighted = font.highlightColor !== 'Mixed' && font.highlightColor !== null && pen !== '' && highlightCss(font.highlightColor) === penCss;
  const penName = pen === '' ? t('toolbar.highlight.none') : t(`highlight.${pen}`);
  const applyHighlight = (value: string): void => { editor.run('font.highlightColor', { value: value || null }); editor.focus(); };
  const pickPen = (value: string): void => { setPen(value); savePen(value); editor.run('font.highlightColor', { value: value || null }); };
  const highlightSwatches = useMemo(() => HIGHLIGHTS.map((h) => ({ id: h, colour: highlightCss(h), name: t(`highlight.${h}`) })), [t]);
  const colour = font.color !== 'Mixed' && font.color.startsWith('#') ? font.color : '#000000';
  const fontColourSwatches = useMemo(() => FONT_COLOURS.map(([id, c, key]) => ({ id, colour: c, name: t(key) })), [t]);

  return (
    <div className="toolbar" role="toolbar" aria-label={t('toolbar.label')} data-testid="toolbar">
      <Group label={t('toolbar.group.undo')} testId="group-undo">
        <Row><ToolButton editor={editor} command="history.undo" label={t('toolbar.undo')} disabled={!enabled('history.undo')}>↶</ToolButton></Row>
        <Row><ToolButton editor={editor} command="history.redo" label={t('toolbar.redo')} disabled={!enabled('history.redo')}>↷</ToolButton></Row>
      </Group>
      <Group label={t('toolbar.group.font')} testId="group-font">
        {full && (
          <Row>
            <select aria-label={t('toolbar.font')} data-testid="font-name" className="font-name" value={fontName} onChange={(e) => { editor.run('font.name', { value: e.target.value }); editor.focus(); }}>
              {fontName === '' && <option value="">{t('toolbar.font.mixed')}</option>}
              {fontList.map((f) => <option key={f} value={f}>{f}</option>)}
            </select>
            <select aria-label={t('toolbar.fontSize')} data-testid="font-size" className="font-size" value={fontSize || ''} onChange={(e) => { editor.run('font.size', { value: Number(e.target.value) }); editor.focus(); }}>
              {fontSize === 0 && <option value="">–</option>}
              {(FONT_SIZES.includes(fontSize) ? FONT_SIZES : [fontSize, ...FONT_SIZES].sort((a, b) => a - b)).map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
            {/* A tool button lays its children out as flex items, where a bare <sub> or <sup> is not lowered or raised: the span keeps the text one inline box. */}
            <ToolButton editor={editor} command="font.grow" label={t('toolbar.grow')}><span>A<sup>+</sup></span></ToolButton>
            <ToolButton editor={editor} command="font.shrink" label={t('toolbar.shrink')}><span>A<sup>−</sup></span></ToolButton>
            <ToolButton editor={editor} command="font.clear" label={t('toolbar.clear')}><span>A<sub>✕</sub></span></ToolButton>
          </Row>
        )}
        <Row>
          <ToolButton editor={editor} command="font.bold" args={{ value: 'toggle' }} label={t('toolbar.bold')} pressed={font.bold === true}><b>B</b></ToolButton>
          <ToolButton editor={editor} command="font.italic" args={{ value: 'toggle' }} label={t('toolbar.italic')} pressed={font.italic === true}><i>I</i></ToolButton>
          <ToolButton editor={editor} command="font.underline" args={{ value: 'toggle' }} label={t('toolbar.underline')} pressed={font.underline !== 'Mixed' && font.underline !== 'None'}><u>U</u></ToolButton>
          {full && (
            <>
              <ToolButton editor={editor} command="font.strikeThrough" args={{ value: 'toggle' }} label={t('toolbar.strikeThrough')} pressed={font.strikeThrough === true}><s>ab</s></ToolButton>
              <ToolButton editor={editor} command="font.subscript" args={{ value: 'toggle' }} label={t('toolbar.subscript')} pressed={font.subscript === true}>x₂</ToolButton>
              <ToolButton editor={editor} command="font.superscript" args={{ value: 'toggle' }} label={t('toolbar.superscript')} pressed={font.superscript === true}>x²</ToolButton>
            </>
          )}
          {/* The highlighter, Word's split button, in every preset (a reviewer marks text with it too): the pen's colour under
              the glyph, a click applies it, the arrow's menu of swatches picks it and applies it too (ED-005 section 12.131). */}
          <span className="split-tool" data-testid="highlight">
            <button type="button" className={`tool${highlighted ? ' pressed' : ''}`} aria-label={t('toolbar.highlight')} title={t('toolbar.highlight.title', { colour: penName })}
              aria-pressed={highlighted} data-command="font.highlightColor" onMouseDown={(e) => e.preventDefault()}
              onClick={() => applyHighlight(highlighted ? '' : pen)}>
              <SwatchIcon glyph="ab" colour={penCss} />
            </button>
            <SwatchMenu editor={editor} testId="highlight-menu" label={t('toolbar.highlight.pick')} command="font.highlightColor"
              summary={<span aria-hidden="true">▾</span>} summaryClass="split-arrow"
              swatches={highlightSwatches} columns={5} isCurrent={(s) => s.id === pen}
              onPick={(s) => pickPen(s ? s.id : '')}
              none={{ label: t('toolbar.highlight.none'), testId: 'highlight-none' }} />
          </span>
          {/* The font colour (ED-005 section 12.131): Automatic first, as in Word, then the swatches and the picker. */}
          {full && (
            <SwatchMenu editor={editor} testId="font-colour" label={t('toolbar.color')} command="font.color"
              summary={<><SwatchIcon glyph="A" colour={colour} /><span className="menu-arrow" aria-hidden="true">▾</span></>}
              swatches={fontColourSwatches} columns={7} isCurrent={(s) => font.color !== 'Mixed' && font.color.toUpperCase() === s.colour}
              onPick={(s) => { editor.run('font.color', { value: s ? s.colour : 'auto' }); }}
              none={{ label: t('review.value.auto'), testId: 'font-colour-auto', first: true }}
              more={{ label: t('toolbar.table.shading.more'), testId: 'font-colour-more', initial: () => colour, onPick: (value) => { editor.run('font.color', { value }); } }} />
          )}
        </Row>
      </Group>
      <Group label={t('toolbar.group.paragraph')} testId="group-paragraph"
        launcher={<button type="button" className="ribbon-launcher" aria-label={t('toolbar.paragraph')} title={t('toolbar.paragraph')} onMouseDown={(e) => e.preventDefault()} onClick={() => setDialog('paragraph')}>↘</button>}>
        <Row>
          <ToolButton editor={editor} command="list.toggle" args={{ bullet: true }} label={t('toolbar.bullets')} pressed={listKind === 'bullet'}>•≡</ToolButton>
          <ToolButton editor={editor} command="list.toggle" args={{ bullet: false }} label={t('toolbar.numbering')} pressed={listKind === 'number'}>1≡</ToolButton>
          <ToolButton editor={editor} command="paragraph.decreaseIndent" first="list.outdent" label={t('toolbar.decreaseIndent')} disabled={!enabled('paragraph.decreaseIndent') && !enabled('list.outdent')}>⇤</ToolButton>
          <ToolButton editor={editor} command="paragraph.increaseIndent" first="list.indent" label={t('toolbar.increaseIndent')}>⇥</ToolButton>
        </Row>
        <Row>
          <ToolButton editor={editor} command="paragraph.alignment" args={{ value: 'Left' }} label={t('toolbar.alignStart')} pressed={para.alignment === 'Left'}><AlignIcon align="left" /></ToolButton>
          <ToolButton editor={editor} command="paragraph.alignment" args={{ value: 'Centered' }} label={t('toolbar.centre')} pressed={para.alignment === 'Centered'}><AlignIcon align="centre" /></ToolButton>
          <ToolButton editor={editor} command="paragraph.alignment" args={{ value: 'Right' }} label={t('toolbar.alignEnd')} pressed={para.alignment === 'Right'}><AlignIcon align="right" /></ToolButton>
          <ToolButton editor={editor} command="paragraph.alignment" args={{ value: 'Justified' }} label={t('toolbar.justify')} pressed={para.alignment === 'Justified'}><AlignIcon align="justify" /></ToolButton>
        </Row>
      </Group>
      <Group label={t('toolbar.group.styles')} testId="group-styles">
        <Row>
          <select aria-label={t('toolbar.style')} data-testid="style-picker" className="style-picker" value={pickerValue} onMouseDown={(e) => e.stopPropagation()}
            onChange={(e) => { editor.run('paragraph.style', { value: e.target.value || null }); editor.focus(); }}>
            {!current && <option value={styleId}>{styleId || t('toolbar.style.default')}</option>}
            {paragraphStyles.map((s) => <option key={s.id} value={s.isDefault ? '' : s.id} className={`s-${s.id}`} style={previews.get(s.id)}>{s.name}{s.quick ? '' : ' ·'}</option>)}
          </select>
        </Row>
      </Group>
      {/* The rest, each in a group of its own (no caption for now), laid out in the ribbon's two rows. */}
      <Group label={t('toolbar.group.insert')} caption={false} testId="group-insert">
        <div className="ribbon-grid">
          <button type="button" className={`tool${reading.inLink ? ' pressed' : ''}`} aria-label={t('toolbar.hyperlink')} title={t('toolbar.hyperlink.title')} data-testid="link-button" onMouseDown={(e) => e.preventDefault()} onClick={() => setDialog('link')} disabled={!enabled('range.hyperlink')}>🔗</button>
          {full && (
            <>
              <ToolButton editor={editor} command="insertBreak" args={{ type: 'Page' }} label={t('toolbar.pageBreak')} disabled={!enabled('insertBreak')}>⤓</ToolButton>
              <ToolButton editor={editor} command="range.insertFootnote" label={t('toolbar.footnote')} disabled={!enabled('range.insertFootnote')}>a¹</ToolButton>
              <PictureButton editor={editor} label={t('toolbar.picture')} disabled={!enabled('range.insertInlinePicture')} />
              <ToolButton editor={editor} command="range.insertEndnote" label={t('toolbar.endnote')} disabled={!enabled('range.insertEndnote')}>aⁱ</ToolButton>
              <button type="button" className="tool" aria-label={t('toolbar.section')} title={editor.story ? t('toolbar.section.inStory') : t('toolbar.section')} onMouseDown={(e) => e.preventDefault()} onClick={() => setDialog('section')} disabled={editor.story !== undefined}>§…</button>
              <button type="button" className="tool" aria-label={t('toolbar.pageNumbers')} title={t('toolbar.pageNumbers')} data-testid="page-numbers" onMouseDown={(e) => e.preventDefault()} onClick={() => { void editor.runAsync('document.pageNumbers').then(() => editor.focus()); }} disabled={!enabled('document.pageNumbers')}>№</button>
            </>
          )}
        </div>
      </Group>
      {controls && (
        <Group label={t('toolbar.controls')} caption={false} testId="control-tools">
          <div className="ribbon-grid">
            <select aria-label={t('toolbar.control.insert')} data-testid="insert-control" value="" onMouseDown={(e) => e.stopPropagation()}
              onChange={(e) => { if (e.target.value) { editor.run('contentControl.insert', { kind: e.target.value as SdtKind }); editor.focus(); } }} disabled={!enabled('contentControl.insert')}>
              <option value="">{t('toolbar.control.insert.prompt')}</option>
              {INSERT_KINDS.map((k) => <option key={k} value={k}>{t(`control.kind.${k}`)}</option>)}
            </select>
            <ToolButton editor={editor} command="contentControl.wrapRows" label={t('toolbar.control.wrapRows')} disabled={!enabled('contentControl.wrapRows')}>⊟⇉</ToolButton>
            <ToolButton editor={editor} command="contentControl.wrapCells" label={t('toolbar.control.wrapCells')} disabled={!enabled('contentControl.wrapCells')}>⊞⇉</ToolButton>
            <ToolButton editor={editor} command="contentControl.delete" args={{ keepContent: true }} label={t('toolbar.control.unwrap')} disabled={!inControl}>⇊</ToolButton>
            <ToolButton editor={editor} command="contentControl.delete" args={{ keepContent: false }} label={t('toolbar.control.delete')} disabled={!inControl}>⇊✕</ToolButton>
            <button type="button" className="tool" aria-label={t('toolbar.control.properties')} title={t('toolbar.control.properties.title')} data-testid="control-tab-button" onMouseDown={(e) => e.preventDefault()} onClick={() => onControl?.()} disabled={!onControl}>⚙</button>
          </div>
        </Group>
      )}
      <Group label={t('toolbar.table')} caption={false} testId={reading.inTable ? 'table-tools' : undefined}>
        <div className="ribbon-grid">
          {reading.inTable
            ? (
              <>
                <ToolButton editor={editor} command="table.addRows" args={{ location: 'Before' }} label={t('toolbar.table.addRowAbove')}>⬆⁺</ToolButton>
                <ToolButton editor={editor} command="table.addRows" args={{ location: 'After' }} label={t('toolbar.table.addRowBelow')}>⬇⁺</ToolButton>
                <ToolButton editor={editor} command="table.deleteRows" label={t('toolbar.table.deleteRow')}>⬍✕</ToolButton>
                <ToolButton editor={editor} command="table.addColumns" args={{ location: 'Before' }} label={t('toolbar.table.addColumnBefore')}>⬅⁺</ToolButton>
                <ToolButton editor={editor} command="table.addColumns" args={{ location: 'After' }} label={t('toolbar.table.addColumnAfter')}>➡⁺</ToolButton>
                <ToolButton editor={editor} command="table.deleteColumns" label={t('toolbar.table.deleteColumn')}>⬌✕</ToolButton>
                <ToolButton editor={editor} command="table.mergeCells" label={t('toolbar.table.merge')} disabled={!enabled('table.mergeCells')}>⊞→⊟</ToolButton>
                <ToolButton editor={editor} command="table.splitCell" label={t('toolbar.table.split')} disabled={!enabled('table.splitCell')}>⊟→⊞</ToolButton>
                <ToolButton editor={editor} command="row.header" label={t('toolbar.table.headerRow')} pressed={reading.headerRow}>⊤</ToolButton>
                <RowShadingMenu editor={editor} current={reading.rowShading} />
                <ToolButton editor={editor} command="table.delete" label={t('toolbar.table.delete')}>▦✕</ToolButton>
                <button type="button" className="tool" aria-label={t('toolbar.table.properties')} title={t('toolbar.table.properties')} data-testid="table-properties" onMouseDown={(e) => e.preventDefault()} onClick={() => setDialog('table')}>▦…</button>
                <button type="button" className="tool" aria-label={t('toolbar.table.fit')} title={t('toolbar.table.fit')} data-testid="table-fit" disabled={!onFitTable} onMouseDown={(e) => e.preventDefault()} onClick={() => onFitTable?.()}>▦⇲</button>
                <button type="button" className="tool" aria-label={t('toolbar.table.insertInCell')} title={t('toolbar.table.insertInCell')} data-command="table.insert" disabled={!enabled('table.insert')} onMouseDown={(e) => e.preventDefault()} onClick={() => setDialog('insertTable')}>▦⁺</button>
              </>
            )
            : <button type="button" className="tool" aria-label={t('toolbar.table.insert')} title={t('toolbar.table.insert')} data-command="table.insert" disabled={!enabled('table.insert')} onMouseDown={(e) => e.preventDefault()} onClick={() => setDialog('insertTable')}>▦</button>}
        </div>
      </Group>
      {dialog === 'paragraph' && <ParagraphDialog editor={editor} formatting={fmt} onClose={() => { setDialog(null); editor.focus(); }} />}
      {dialog === 'section' && <SectionDialog editor={editor} onClose={() => { setDialog(null); editor.refreshCss(); editor.focus(); }} />}
      {dialog === 'table' && <TableDialog editor={editor} onClose={() => { setDialog(null); editor.focus(); }} onFit={onFitTable} />}
      {dialog === 'insertTable' && <InsertTableDialog editor={editor} onClose={() => { setDialog(null); editor.focus(); }} />}
      {dialog === 'link' && <LinkDialog editor={editor} onClose={() => { setDialog(null); editor.focus(); }} />}
    </div>
  );
}
