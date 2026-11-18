// SPDX-License-Identifier: GPL-3.0-or-later
//
// The application's words for things that carry an English name of their own (ED-003 section
// 12.2): a command's label (the registry's `label` is the command's canonical name, the API's
// and the palette's match), a perspective's and a tab's (a plugin's own label when the shipped
// files have no key for it), and the sentences built from the model's data.
import { preservedCount, commandNamed, type CommandSpec, type Diagnostics, type LockRefusal } from '@docx4j/editor-model';
import { t, tOr, plainOr } from './i18n.js';
import type { Perspective, TabSpec } from './perspectives.js';

/** A command's label in the current language: `command.<name>`, else the registry's English. */
export function commandLabel(spec: Pick<CommandSpec<unknown>, 'name' | 'label'>): string {
  return tOr(`command.${spec.name}`, spec.label);
}

/** A command argument's description in the current language, else the registry's: plain text, since it quotes code in braces. */
export function argumentDescription(spec: Pick<CommandSpec<unknown>, 'name'>, arg: string, fallback: string): string {
  return plainOr(`command.${spec.name}.arg.${arg}`, fallback);
}

export function perspectiveLabel(p: Pick<Perspective, 'id' | 'label'>): string {
  return tOr(`perspective.${p.id}`, p.label);
}

export function tabLabel(tab: Pick<TabSpec, 'id' | 'label'>): string {
  return tOr(`tab.${tab.id}`, tab.label);
}

/** The status line for an edit that ran untracked with tracking on (ED-005 section 5.4): in a story, or an untracked command. */
export function untrackedStatus(what: string): string {
  if (what === 'story') return t('status.untracked.story');
  const spec = commandNamed(what);
  return t('status.untracked.command', { command: spec ? commandLabel(spec) : what });
}

/** The banner's sentence over the diagnostics (the model's `summarize`, in the current language), with the tracked changes as the review panel lists them. */
export function diagnosticSummary(d: Diagnostics, changes: number): string {
  const parts: string[] = [];
  if (d.opaque.length > 0) parts.push(t('banner.summary.opaque', { count: d.opaque.length }));
  if (changes > 0) parts.push(t('banner.summary.tracked', { count: changes }));
  const kept = preservedCount(d);
  if (kept > 0) parts.push(t('banner.summary.preserved', { count: kept }));
  // Comments too, so that a document with comments and nothing else is offered Review (ED-005 section 6.7).
  if (d.comments > 0) parts.push(t('banner.summary.comments', { count: d.comments }));
  return parts.length === 0 ? '' : t('banner.summary', { parts: parts.join('; ') });
}

/** The status line for an edit a content control's lock refused (ED-005 section 12.60). */
export function lockRefusedStatus(reason: LockRefusal): string {
  return t(reason === 'cannotDelete' ? 'status.lock.cannotDelete' : 'status.lock.cannotEdit');
}
