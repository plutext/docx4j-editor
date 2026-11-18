// SPDX-License-Identifier: GPL-3.0-or-later
//
// The font box's name when the effective reading is '' (a theme font, ED-002 section 10.2 item
// 6): resolved through the run-font selector, as the CSS is.
import type { Element } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { paragraphPropertiesFor, runPropertiesFor } from '@docx4j/editor-model';
import type { EditorSurface } from '../view/surface.js';
import { latinFontOf } from '../view/css.js';

export function resolvedFontName(editor: EditorSurface): string {
  const { state, session } = editor;
  const resolver = session.pkg.propertyResolverOrUndefined;
  const selector = session.main.runFontSelectorOrUndefined;
  if (!resolver || !selector) return '';
  const $from = state.selection.$from;
  const paragraph = $from.node($from.depth);
  if (paragraph.type.name !== 'paragraph') return '';
  const marks = state.storedMarks ?? $from.marks();
  const runSource = marks.find((m) => m.type.name === 'runSource')?.attrs['sourceId'] as string | undefined;
  const source = runSource ? (session.ctx.ids.get(runSource) as Element<wml.R> | undefined) : undefined;
  const pPr = paragraphPropertiesFor(paragraph, state);
  const rPr = resolver.getEffectiveRPr(runPropertiesFor(source?.value.rPr, marks), pPr);
  return latinFontOf(selector, pPr, rPr);
}

// ---------------------------------------------------------------- embedded fonts (ED-003 section 8.4)

import { embeddedFaces, looksLikeFont, type EditorDocument, type EmbeddedFace } from '@docx4j/editor-model';
import { t } from './i18n.js';

export interface EmbeddedFontStatus {
  /** Faces loaded, as `family weight/style`. */
  loaded: string[];
  /** Faces that would not load, with the reason. */
  failed: string[];
  /** Faces whose bytes the model repaired on the copy before the browser would take them (`EmbeddedFace.repairs`, UPSTREAM-5). */
  repaired: string[];
  faces: EmbeddedFace[];
}

const loadedFaces = new WeakMap<EditorDocument, { faces: FontFace[]; status: EmbeddedFontStatus }>();

/** One line for the status bar. */
export function describeEmbeddedFonts(status: EmbeddedFontStatus): string {
  const families = new Map<string, { faces: number; subset: number }>();
  for (const face of status.faces) {
    const entry = families.get(face.family) ?? { faces: 0, subset: 0 };
    entry.faces++;
    if (face.subsetted) entry.subset++;
    families.set(face.family, entry);
  }
  // A family Word subsetted (`w:subsetted`) holds the characters the document used when it was saved; a character outside them falls back to the next font in the stack.
  const parts = [...families].map(([f, { faces, subset }]) => t('fonts.face', { family: f, count: faces, subset: subset > 0 ? subset === faces ? t('fonts.subset.all') : t('fonts.subset.some', { count: subset }) : '' }));
  const head = status.loaded.length > 0 ? t('fonts.loaded', { families: parts.join(', ') }) : '';
  return status.failed.length > 0 ? `${head}${t(head ? 'fonts.failedAfter' : 'fonts.failed', { faces: status.failed.join(', ') })}` : head;
}

/**
 * The document's embedded faces as `FontFace`s for the session (section 9 item 34): added to
 * `document.fonts`, so the stacks that name the family first (`css.ts`) render with them, and
 * removed with `unloadEmbeddedFonts` when the document closes. Nothing is written. A face that
 * fails to parse is reported and the stack falls through as before.
 */
export async function loadEmbeddedFonts(session: EditorDocument): Promise<EmbeddedFontStatus> {
  const status: EmbeddedFontStatus = { loaded: [], failed: [], repaired: [], faces: [] };
  if (typeof FontFace === 'undefined' || typeof document === 'undefined' || !document.fonts) return status;
  let faces: EmbeddedFace[];
  try { faces = await embeddedFaces(session); } catch (e) { status.failed.push(t('fonts.tableUnreadable', { error: (e as Error).message })); return status; }
  const made: FontFace[] = [];
  for (const face of faces) {
    const label = `${face.family} ${face.weight}/${face.style}`;
    if (!looksLikeFont(face.bytes)) { status.failed.push(t('fonts.notAFont', { face: label })); continue; }
    try {
      const ff = new FontFace(face.family, face.bytes as BufferSource, { weight: face.weight, style: face.style });
      await ff.load();
      document.fonts.add(ff);
      made.push(ff);
      status.loaded.push(label);
      status.faces.push(face);
      if (face.repairs.length > 0) status.repaired.push(label);
    } catch (e) {
      status.failed.push(t('fonts.faceError', { face: label, error: (e as Error).message }));
    }
  }
  loadedFaces.set(session, { faces: made, status });
  return status;
}

export function unloadEmbeddedFonts(session: EditorDocument): void {
  const entry = loadedFaces.get(session);
  if (!entry) return;
  for (const ff of entry.faces) { try { document.fonts.delete(ff); } catch { /* gone already */ } }
  loadedFaces.delete(session);
}

/** What was loaded for a session, for the diagnostics. */
export function embeddedFontStatusOf(session: EditorDocument): EmbeddedFontStatus | undefined {
  return loadedFaces.get(session)?.status;
}
