// SPDX-License-Identifier: GPL-3.0-or-later
//
// Proofing (Jason, 2026-11-11, from the browser's red squiggles under trade names): the editor checks no spelling itself,
// but tells the browser's spell checker what it needs.
// - "Check spelling as you type", an option kept per browser: `spellcheck` on every editing surface.
// - The document's own language on each surface (`lang`): the default paragraph style's `w:lang`, through the document
//   defaults, where the browser otherwise takes the application's language.
// - Per run, as inline decorations from the decorations plugin (`proofingOf`): a language that differs from the
//   document's, and Word's "Do not check spelling or grammar" (`w:noProof`) as `spellcheck="false"`. Each from the run's
//   own properties, its character style, or its paragraph's style, as Word resolves them, with `w:lang`'s East Asian or
//   complex-script value for text in those scripts.
// Which dictionary a browser uses for a `lang` is the browser's: Firefox picks its dictionary from it, Chrome checks
// against the languages enabled in its own settings.
import { Plugin, PluginKey, type EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import type { Element, PropertyResolver } from '@docx4j/core-ts';
import type { EditorDocument, SourceIds } from '@docx4j/editor-model';
import { surfaces } from './surface.js';
import { characterStyleRPr } from './css.js';
import { PARTS_CHANGED } from './plugins/labels.js';

let spellcheck = true;

const proofingKey = new PluginKey<string | undefined>('docx4j-proofing');

/** Whether the editing surfaces ask the browser to check spelling (the option, kept per browser by the application). */
export function spellcheckEnabled(): boolean {
  return spellcheck;
}

/** The option set: every live surface's `spellcheck` follows at once (a transaction recomputes a view's attributes). */
export function setSpellcheck(on: boolean): void {
  if (on === spellcheck) return;
  spellcheck = on;
  for (const s of surfaces.values()) s.view.dispatch(s.state.tr.setMeta(proofingKey, true));
}

/** A `w:lang`'s value for some text: the East Asian one for CJK text, the complex-script one for RTL and Indic text, else `w:val`. */
const CJK = /[⺀-鿿가-힯豈-﫿＀-￯]/;
const COMPLEX = /[֐-ࣿऀ-෿฀-໿ༀ-࿿ក-៿יִ-﷿ﹰ-﻿]/;
function languageOf(lang: wml.CTLanguage | undefined, text: string): string | undefined {
  if (!lang) return undefined;
  const v = CJK.test(text) ? lang.eastAsia : COMPLEX.test(text) ? lang.bidi : lang.val;
  return v && v !== 'x-none' ? v : undefined;
}

/** A `w:noProof` (or any on/off property): undefined when absent, else on unless `w:val` says off. */
function onOff(value: { val?: boolean | null } | undefined | null): boolean | undefined {
  return value === undefined || value === null ? undefined : value.val !== false;
}

let languageCache = new WeakMap<PropertyResolver, { lang: string | undefined }>();

/** The document's language: its default paragraph style's `w:lang/@w:val`, through the document defaults; cached per resolver until a part changes. */
export function documentLanguage(session: EditorDocument): string | undefined {
  const resolver = session.pkg.propertyResolverOrUndefined;
  if (!resolver) return undefined;
  const hit = languageCache.get(resolver);
  if (hit) return hit.lang;
  const lang = languageOf(defaultRPr(resolver)?.lang ?? undefined, '');
  languageCache.set(resolver, { lang });
  return lang;
}

/** A part the look depends on changed: the language and the styles' proofing are read again. */
function forget(): void {
  languageCache = new WeakMap();
  styleCache = new WeakMap();
}

/** The document defaults through the default paragraph style, as the stylesheet reads them (`css.ts`). */
function defaultRPr(resolver: PropertyResolver): wml.RPr | undefined {
  const id = resolver.defaultParagraphStyleId;
  return id !== undefined ? resolver.getEffectiveRPr(id) : resolver.getEffectiveRPr(undefined, undefined);
}

interface StyleProofing { lang?: wml.CTLanguage; noProof?: boolean }

let styleCache = new WeakMap<PropertyResolver, Map<string, StyleProofing>>();

/** A paragraph style's (`p:`) or a character style's (`r:`) language and proofing, cached per resolver. */
function styleProofing(resolver: PropertyResolver, key: string): StyleProofing {
  let cache = styleCache.get(resolver);
  if (!cache) { cache = new Map(); styleCache.set(resolver, cache); }
  const hit = cache.get(key);
  if (hit) return hit;
  const id = key.slice(2);
  const rPr = key.startsWith('p:') ? (id === '' ? defaultRPr(resolver) : resolver.getEffectiveRPr(id)) : characterStyleRPr(resolver, id);
  const out: StyleProofing = {};
  if (rPr?.lang) out.lang = rPr.lang;
  const noProof = onOff(rPr?.noProof);
  if (noProof !== undefined) out.noProof = noProof;
  cache.set(key, out);
  return out;
}

/** A stretch of a paragraph whose proofing differs from the surface's: its own language, or no spell checking. */
export interface RelProofing {
  from: number;
  to: number;
  attrs: { lang?: string; spellcheck?: 'false' };
}

/**
 * The runs of a paragraph whose language is not the document's, or which Word marks "Do not check spelling or grammar",
 * as offsets from the paragraph's start (the theme colours' convention), touching runs with the same attributes merged.
 * A run's properties come from its source run (`runSource`), its character style (the `runStyle` mark, else the source's),
 * then its paragraph's style; text with no source run takes its paragraph's.
 */
export function proofingOf(paragraph: PMNode, ids: SourceIds, resolver: PropertyResolver, documentLang: string | undefined): RelProofing[] {
  const out: RelProofing[] = [];
  const para = styleProofing(resolver, `p:${(paragraph.attrs['pStyle'] as string | null) ?? ''}`);
  paragraph.descendants((child, rel) => {
    if (!child.isText) return true;
    const runSource = child.marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'] as string | undefined;
    const rPr = runSource ? (ids.get(runSource) as Element<wml.R> | undefined)?.value.rPr ?? undefined : undefined;
    const rStyle = (child.marks.find((m) => m.type.name === 'runStyle')?.attrs['id'] as string | undefined) ?? rPr?.rStyle?.val ?? undefined;
    const char = rStyle ? styleProofing(resolver, `r:${rStyle}`) : {};
    const lang = languageOf(rPr?.lang ?? char.lang ?? para.lang, child.text ?? '');
    const noProof = onOff(rPr?.noProof) ?? char.noProof ?? para.noProof ?? false;
    const attrs: RelProofing['attrs'] = {};
    if (lang && lang !== documentLang) attrs.lang = lang;
    if (noProof) attrs.spellcheck = 'false';
    if (attrs.lang === undefined && attrs.spellcheck === undefined) return false;
    const from = 1 + rel;
    const to = from + child.nodeSize;
    const last = out[out.length - 1];
    if (last && last.to === from && last.attrs.lang === attrs.lang && last.attrs.spellcheck === attrs.spellcheck) last.to = to;
    else out.push({ from, to, attrs });
    return false;
  });
  return out;
}

/**
 * Every editing surface's `spellcheck` and `lang` (the option, the document's language), as a plugin's attributes, which a
 * view merges with its own. The language is read again when a part the look depends on changes (`PARTS_CHANGED`).
 */
export function proofingPlugin(session: EditorDocument): Plugin<string | undefined> {
  return new Plugin<string | undefined>({
    key: proofingKey,
    state: {
      init: () => documentLanguage(session),
      apply: (tr, value) => { if (!tr.getMeta(PARTS_CHANGED)) return value; forget(); return documentLanguage(session); },
    },
    props: {
      attributes: (state: EditorState): Record<string, string> => {
        const lang = proofingKey.getState(state);
        return { spellcheck: spellcheck ? 'true' : 'false', ...(lang ? { lang } : {}) };
      },
    },
  });
}
