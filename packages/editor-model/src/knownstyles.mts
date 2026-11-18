// SPDX-License-Identifier: GPL-3.0-or-later
//
// The styles a document lacks, spliced in through the engine's `styles.ensure`, which since core-ts 0.4.0
// (CR-002 section 41) falls back to Word's built-ins from docx4j's KnownStyles.xml, closing over what each is
// based on and linked to. Until then this module carried eleven of Word's definitions of its own (UPSTREAM-20,
// ED-005 sections 12.66 and 1.11); they are gone. An id the engine has no definition for is skipped, as its
// refusal was before.
import { Docx4JException, type WordprocessingMLPackage } from '@docx4j/core-ts';

/**
 * The styles the document lacks, spliced in through the engine. The ids added, in order (what a style is
 * based on or linked to before it, as the engine closes); the caller marks the styles part changed when there
 * are any. An id with no definition anywhere is skipped.
 */
export async function ensureStylesIn(pkg: WordprocessingMLPackage, ids: readonly string[]): Promise<string[]> {
  const added: string[] = [];
  for (const id of ids) {
    try {
      added.push(...await pkg.styles.ensure(id));
    } catch (e) {
      if (!(e instanceof Docx4JException)) throw e;
    }
  }
  return added;
}
