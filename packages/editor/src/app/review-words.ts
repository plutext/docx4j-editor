// SPDX-License-Identifier: GPL-3.0-or-later
//
// What a formatting or properties change changed, in words, with its values, as Word's Reviewing
// Pane words them (ED-005 proposal 43, section 9.7): "Font: Bold", "Font: Not Bold", "Font: 16 pt",
// "Centred", "Left indent: 36 pt", "Heading 2", "Numbered + Level: 1 + 1.". The model gives each
// changed property with its value now (`changedValues`); the words, the style names (the styles part
// through the resolver, as the style box shows them), the list's kind and label (the numbering
// definitions and the labels the view draws) and the units (points, as the Paragraph dialog has
// them) are the application's. And a note reference inserted or deleted, whose text is empty: the
// number the view draws for it.
import type { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { changedValues, listKindOf, numberingOf, highlightNameForColor, type ChangedValue, type EditorDocument, type ParagraphAttrs, type TrackedChangeInfo } from '@docx4j/editor-model';
import { labelLookup } from '../view/plugins/labels.js';
import { decorationsKey } from '../view/plugins/decorations.js';
import { t } from './i18n.js';

/** A property's own word: the command's label. */
const WORDS: Record<string, string> = {
  bold: 'command.font.bold', italic: 'command.font.italic', underline: 'command.font.underline', strike: 'command.font.strikeThrough', rtl: 'command.font.rtl',
  'font.name': 'command.font.name', 'font.size': 'command.font.size', 'font.color': 'command.font.color', 'font.highlight': 'command.font.highlightColor',
  runStyle: 'command.font.style', pStyle: 'command.paragraph.style', jc: 'command.paragraph.alignment', tabs: 'command.paragraph.tabStops',
  'ind.start': 'command.paragraph.leftIndent', 'ind.end': 'command.paragraph.rightIndent', 'ind.firstLine': 'command.paragraph.firstLineIndent',
  'spacing.before': 'command.paragraph.spaceBefore', 'spacing.after': 'command.paragraph.spaceAfter', 'spacing.line': 'command.paragraph.lineSpacing',
};

/** Twips in points, to a tenth. */
const points = (twips: number): string => t('review.value.points', { size: Math.round(twips / 2) / 10 });

/** A style's name as the styles part gives it, else its id. */
function styleName(session: EditorDocument, id: string): string {
  return session.pkg.propertyResolverOrUndefined?.getStyle(id)?.name?.val ?? id;
}

/** The properties Word words under "Font:". */
const FONT = new Set(['bold', 'italic', 'rtl', 'strike', 'underline', 'vertAlign', 'font.name', 'font.size']);

/** A font property's value as it follows "Font:" ("Bold", "Not Bold", "16 pt", "Arial"); undefined for a size taken away, which has no word there. */
function fontValueWords(v: ChangedValue): string | undefined {
  const word = t(WORDS[v.property] ?? v.property);
  const not = (property: string): string => t('review.value.not', { property });
  const value = v.value;
  switch (v.property) {
    case 'strike': return value === null ? not(word) : value === 'double' ? t('command.font.doubleStrikeThrough') : word;
    case 'underline': return value === null ? not(word) : value === 'Single' || value === true ? word : t('review.value.underlineStyle', { underline: word, style: String(value) });
    case 'vertAlign': return value === 'subscript' ? t('command.font.subscript') : value === 'superscript' ? t('command.font.superscript') : t('review.value.baseline');
    case 'font.name': return typeof value === 'string' ? value : t('review.value.default');
    case 'font.size': return typeof value === 'number' ? t('review.value.points', { size: value }) : undefined;
    default: return value === null ? not(word) : word;
  }
}

/** One changed property with its value, in words. */
export function valueWords(v: ChangedValue, state: EditorState, session: EditorDocument, node: PMNode | null): string {
  const word = (): string => t(WORDS[v.property] ?? v.property);
  const named = (value: string, property = word()): string => t('review.value.named', { property, value });
  const deflt = t('review.value.default');
  const value = v.value;
  if (FONT.has(v.property)) {
    const words = fontValueWords(v);
    return words === undefined ? named(deflt) : named(words, t('toolbar.font'));
  }
  switch (v.property) {
    case 'font.color':
      return named(value === null ? deflt : value === 'auto' ? t('review.value.auto') : String(value));
    case 'font.highlight': {
      if (value === null) return t('toolbar.highlight.none');
      // The name by the object model's table, the one the engine reads `w:highlight` by (a copy here had Word's palette ahead of the objects release).
      const name = highlightNameForColor(String(value));
      return named(name ? t(`highlight.${name}`) : String(value));
    }
    case 'runStyle':
      return value === null ? named(deflt) : styleName(session, String(value));
    case 'pStyle': {
      const id = typeof value === 'string' ? value : session.pkg.propertyResolverOrUndefined?.getDefaultParagraphStyleId();
      return id ? styleName(session, id) : named(deflt);
    }
    case 'jc':
      return value === null ? named(deflt) : t('review.value.jc', { jc: String(value) });
    case 'bidi':
      return t('review.value.direction', { rtl: value === true ? 'true' : 'false' });
    case 'numPr': {
      const attrs = node?.attrs as ParagraphAttrs | undefined;
      const kind = attrs ? listKindOf(attrs, numberingOf(state)) : undefined;
      if (!kind || !node) return t('review.value.noList');
      const label = labelLookup(state)(node);
      return t('review.value.list', { kind, level: (label?.level ?? (value as { ilvl?: number } | null)?.ilvl ?? 0) + 1, label: label?.listString.trim() || 'none' });
    }
    case 'tabs': {
      const stops = (value as { pos: number; val: string }[] | null) ?? [];
      if (stops.length === 0) return named(deflt);
      return named(stops.map((s) => (s.val === 'clear' ? t('review.value.notAt', { position: points(s.pos) }) : points(s.pos))).join(', '));
    }
    case 'ind.firstLine':
      if (typeof value === 'number' && value < 0) return named(points(-value), t('review.value.hanging'));
      return named(typeof value === 'number' ? points(value) : deflt);
    case 'spacing.line': {
      const line = value as { line: number; rule: string | null } | null;
      if (!line) return named(deflt);
      if (line.rule === 'exact') return named(t('review.value.exactly', { size: points(line.line) }));
      if (line.rule === 'atLeast') return named(t('review.value.atLeast', { size: points(line.line) }));
      return named(t('review.value.lines', { lines: Math.round(line.line / 24) / 10 }));
    }
    default:
      // The indents and the spacing before and after, in twips.
      return named(typeof value === 'number' ? points(value) : deflt);
  }
}

/** What a formatting or properties change changed, with the values, as one line; empty for any other change. */
export function changeWords(state: EditorState, session: EditorDocument, change: TrackedChangeInfo): string {
  const values = changedValues(state.doc, change);
  if (values.length === 0) return '';
  const piece = change.pieces.find((p) => p.what === 'properties');
  const node = piece ? state.doc.nodeAt(piece.pos) : null;
  // A list made states the level's indents: the paragraph's own taken away are the list's, which its item describes (Word's pane does the same).
  const listed = values.some((v) => v.property === 'numPr' && v.value !== null);
  const shown = values.filter((v) => !(listed && v.property.startsWith('ind.') && v.value === null));
  // The font's properties under one "Font:", as Word writes them ("Font: Italic, Underline").
  const font = shown.map((v) => (FONT.has(v.property) ? fontValueWords(v) : undefined)).filter((w): w is string => w !== undefined);
  const rest = shown.filter((v) => !FONT.has(v.property) || fontValueWords(v) === undefined).map((v) => valueWords(v, state, session, node));
  return [...(font.length > 0 ? [t('review.value.named', { property: t('toolbar.font'), value: font.join(', ') })] : []), ...rest].join(', ');
}

/** A change that is a note reference inserted or deleted, whose text is empty: the note and the number the view draws for it. */
export function noteReferenceWords(state: EditorState, change: TrackedChangeInfo): string | undefined {
  if (change.text !== '') return undefined;
  for (const p of change.pieces) {
    if (p.what !== 'atom') continue;
    const node = state.doc.nodeAt(p.pos);
    if (node?.type.name !== 'footnoteRef') continue;
    const kind = node.attrs['kind'] === 'endnote' ? 'endnote' : 'footnote';
    const label = decorationsKey.getState(state)?.noteLabels.get(`${kind}:${String(node.attrs['footnoteId'])}`);
    return t('review.noteRef', { kind, label: label ?? String(node.attrs['footnoteId']) });
  }
  return undefined;
}
